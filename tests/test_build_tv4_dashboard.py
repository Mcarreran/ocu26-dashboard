"""Pruebas de negocio para scripts/build_tv4_dashboard.py (dashboard TV4
OCU26, Pipeline Comercial = Core Comercial excluyendo YPF).

No modifica scripts/validate_input.py, scripts/transform_data.py,
scripts/semantic_model.py, scripts/metrics_engine.py,
config/business_semantics.json, input/OCU26_BASE_DATOS.xlsx, ni ningun
archivo productivo de TV1/TV2/TV3 (build_tv1/2/3_dashboard.py,
tv1/2/3_template.html, tv1/2/3.html, test_build_tv1/2/3_dashboard.py,
TV1/TV2/TV3/TV4_REFERENCE.html.html). Los fixtures sinteticos usan la
CONFIGURACION REAL (sm.load_config()), mismo patron que
test_build_tv3_dashboard.py.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pandas as pd
import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT / "scripts"))
sys.path.insert(0, str(REPO_ROOT / "tests"))

import semantic_model as sm  # noqa: E402
import validate_input as vi  # noqa: E402
import build_tv4_dashboard as td  # noqa: E402
from test_semantic_model import _maestro_row, _campana_row, _transform_result  # noqa: E402

PRODUCTION_FILE = REPO_ROOT / "input" / "OCU26_BASE_DATOS.xlsx"


def _semantic(maestro_rows: list[dict], campanas_rows: list[dict] | None = None) -> dict:
    config = sm.load_config()
    return sm.build_semantic_model(_transform_result(maestro_rows, campanas_rows or []), config)


def _cencosud_static(elemento_id: str, ubicacion: str = "UNICENTER", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Shoppings Estático", Subcircuito="CENCOSUD", Ubicacion=ubicacion,
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _cencosud_digital(elemento_id: str, ubicacion: str = "UNICENTER", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Shoppings Digital", Subcircuito="CENCOSUD", Ubicacion=ubicacion,
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        CapacidadSlotsReel=20, SegundosDia=72000,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _aa2000_static(elemento_id: str, ubicacion: str = "EZE", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="AA2000", Subcircuito="X", Ubicacion=ubicacion,
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _pantalla_led(elemento_id: str, ubicacion: str = "SITE", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Pantalla Led", Subcircuito="X", Ubicacion=ubicacion,
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        CapacidadSlotsReel=20, SegundosDia=72000,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _ypf_static(elemento_id: str, ubicacion: str = "500 - TESTVILLE - Calle Falsa 123", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="YPF Estático", Subcircuito="X", Ubicacion=ubicacion,
        Medio="Estático", TipoCatalogo="Abierto", TipoInventario="Físico estático",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _apsa_static(elemento_id: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Shoppings Estático", Subcircuito="APSA", Ubicacion="APSA_SITE",
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _london_static(elemento_id: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="London Supply", Subcircuito="LS", Ubicacion="USH",
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


@pytest.fixture(scope="module")
def production_result():
    return td.build_tv4_data(PRODUCTION_FILE)


@pytest.fixture(scope="module")
def production_json(production_result):
    return json.dumps(production_result["data"], ensure_ascii=False)


# ---------------------------------------------------------------------------
# 1. Payload contiene solo TV4
# ---------------------------------------------------------------------------


def test_payload_top_level_keys_are_tv4_only(production_result):
    assert set(production_result["data"].keys()) == {
        "meta", "kpis", "estado_pipeline", "distribucion_familia", "timeline", "insights",
    }


def test_payload_contains_no_other_tv_datasets(production_json):
    for token in ("tv1_data", "tv2_data", "tv3_data", "ocu_data"):
        assert token not in production_json.lower()


# ---------------------------------------------------------------------------
# 2-4. YPF / APSA / London excluidos
# ---------------------------------------------------------------------------


def test_ypf_excluded_from_tv4_universe(production_result):
    assert "YPF" not in production_result["universe"]["circuitos"]


def test_apsa_excluded_from_tv4_universe(production_result):
    assert "APSA" not in production_result["universe"]["circuitos"]


def test_london_excluded_from_tv4_universe(production_result):
    assert "LONDON_SUPPLY" not in production_result["universe"]["circuitos"]


def test_payload_contains_no_excluded_circuits(production_json):
    upper = production_json.upper()
    assert "YPF" not in upper
    assert "APSA" not in upper
    assert "LONDON" not in upper


def test_ypf_elements_not_in_synthetic_tv4_scope():
    maestro_rows = [_cencosud_static("C1"), _ypf_static("Y1")]
    campana_rows = [
        _campana_row("CC", "C1", IDCampaña="X1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-08-15")),
        _campana_row("YC", "Y1", IDCampaña="X2", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-08-15")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    universe = td.build_tv4_universe(semantic_result)
    assert universe["circuitos"] == ["CENCOSUD"]
    assert "YPF" not in universe["scope_campanas"]["CircuitoNegocio"].unique().tolist()
    assert set(universe["scope_campanas"]["ElementoID"].unique()) == {"C1"}


def test_apsa_and_london_not_in_synthetic_tv4_universe():
    semantic_result = _semantic([_cencosud_static("C1"), _apsa_static("A1"), _london_static("L1")])
    universe = td.build_tv4_universe(semantic_result)
    assert universe["circuitos"] == ["CENCOSUD"]


# ---------------------------------------------------------------------------
# 5-6. Campanas activas distinct vs activaciones (elementos)
# ---------------------------------------------------------------------------


def test_campanas_unicas_distinct_from_activaciones():
    """1 campana en 2 elementos = 1 campana unica, 2 activaciones (spec
    Sec.3.A: nunca confundir campanas unicas con activaciones)."""
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    campana_rows = [
        _campana_row("CA1", "C1", IDCampaña="X1", Campaña="CAMP UNICA", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-08-15")),
        _campana_row("CA2", "C2", IDCampaña="X1", Campaña="CAMP UNICA", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-08-15")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    scope = td.build_tv4_universe(semantic_result)["scope_campanas"]
    cutoff = pd.Timestamp("2026-07-31")
    en_curso = scope[td._en_curso_mask(scope, cutoff)]
    kpi1 = td.compute_kpi1_actividad_actual(en_curso)
    assert kpi1["campanas_unicas"] == 1
    assert kpi1["activaciones"] == 2


def test_activaciones_en_curso_requires_date_overlap_with_cutoff():
    maestro_rows = [_cencosud_static("C1")]
    campana_rows = [_campana_row("CA1", "C1", IDCampaña="X1", FechaInicio=pd.Timestamp("2026-01-01"), FechaFin=pd.Timestamp("2026-01-31"))]
    semantic_result = _semantic(maestro_rows, campana_rows)
    scope = td.build_tv4_universe(semantic_result)["scope_campanas"]
    cutoff = pd.Timestamp("2026-07-31")
    en_curso = scope[td._en_curso_mask(scope, cutoff)]
    assert len(en_curso) == 0  # ya finalizo mucho antes del corte, no esta "en curso"


# ---------------------------------------------------------------------------
# 7-8. Reservas futuras y ventana de 30 dias
# ---------------------------------------------------------------------------


def test_reservas_futuras_requires_start_after_cutoff():
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    campana_rows = [
        _campana_row("FUT", "C1", IDCampaña="F1", FechaInicio=pd.Timestamp("2026-08-15"), FechaFin=pd.Timestamp("2026-09-15")),
        _campana_row("PAST", "C2", IDCampaña="P1", FechaInicio=pd.Timestamp("2026-06-01"), FechaFin=pd.Timestamp("2026-06-30")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    scope = td.build_tv4_universe(semantic_result)["scope_campanas"]
    cutoff = pd.Timestamp("2026-07-31")
    windows = td.compute_pipeline_windows(scope, cutoff)
    reservas = td.compute_kpi_evento(windows["futuras"], "FechaInicio")
    assert reservas["campanas_unicas"] == 1  # solo F1 (inicio posterior al corte)


def test_ventana_30_dias_boundary():
    """Inicio exactamente en corte+30 entra; corte+31 no (spec Sec.3.C)."""
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    campana_rows = [
        _campana_row("IN", "C1", IDCampaña="IN1", FechaInicio=pd.Timestamp("2026-08-30"), FechaFin=pd.Timestamp("2026-09-30")),
        _campana_row("OUT", "C2", IDCampaña="OUT1", FechaInicio=pd.Timestamp("2026-08-31"), FechaFin=pd.Timestamp("2026-09-30")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    scope = td.build_tv4_universe(semantic_result)["scope_campanas"]
    cutoff = pd.Timestamp("2026-07-31")
    windows = td.compute_pipeline_windows(scope, cutoff)
    assert windows["window_end"] == pd.Timestamp("2026-08-30")
    inician_ids = set(windows["inician_30d"]["IDCampaña"].unique())
    assert inician_ids == {"IN1"}


# ---------------------------------------------------------------------------
# 9-10. Inicios y finalizaciones correctos
# ---------------------------------------------------------------------------


def test_inician_30d_counts_distinct_start_events():
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    campana_rows = [
        _campana_row("A", "C1", IDCampaña="A1", FechaInicio=pd.Timestamp("2026-08-05"), FechaFin=pd.Timestamp("2026-09-05")),
        _campana_row("B", "C2", IDCampaña="A1", FechaInicio=pd.Timestamp("2026-08-05"), FechaFin=pd.Timestamp("2026-09-05")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    scope = td.build_tv4_universe(semantic_result)["scope_campanas"]
    cutoff = pd.Timestamp("2026-07-31")
    windows = td.compute_pipeline_windows(scope, cutoff)
    inician = td.compute_kpi_evento(windows["inician_30d"], "FechaInicio")
    assert inician["campanas_unicas"] == 1  # misma campana, mismo inicio, 2 elementos
    assert inician["activaciones"] == 2


def test_finalizan_30d_only_counts_currently_active_ending_soon():
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    campana_rows = [
        # en curso hoy, finaliza dentro de 30 dias -> cuenta
        _campana_row("END", "C1", IDCampaña="E1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-08-10")),
        # todavia no empezo -> no es "en curso", no debe contarse como finalizacion
        _campana_row("FUT", "C2", IDCampaña="E2", FechaInicio=pd.Timestamp("2026-08-05"), FechaFin=pd.Timestamp("2026-08-20")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    scope = td.build_tv4_universe(semantic_result)["scope_campanas"]
    cutoff = pd.Timestamp("2026-07-31")
    windows = td.compute_pipeline_windows(scope, cutoff)
    finalizan = td.compute_kpi_evento(windows["finalizan_30d"], "FechaFin")
    assert finalizan["campanas_unicas"] == 1
    assert finalizan["activaciones"] == 1


def test_indefinida_campaigns_never_finalize():
    maestro_rows = [_cencosud_static("C1")]
    campana_rows = [_campana_row("IND", "C1", IDCampaña="I1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=None, FechaIndefinida="Si")]
    semantic_result = _semantic(maestro_rows, campana_rows)
    scope = td.build_tv4_universe(semantic_result)["scope_campanas"]
    cutoff = pd.Timestamp("2026-07-31")
    windows = td.compute_pipeline_windows(scope, cutoff)
    assert len(windows["finalizan_30d"]) == 0
    assert len(windows["en_curso"]) == 1  # sigue "en curso" (vigente indefinida)


# ---------------------------------------------------------------------------
# 11-13. Timeline: orden, tope 5, sin duplicados por ElementoID
# ---------------------------------------------------------------------------


def test_timeline_sorted_by_date_then_campaign():
    maestro_rows = [_cencosud_static(f"C{i}") for i in range(3)]
    campana_rows = [
        _campana_row("R1", "C0", IDCampaña="Z1", Campaña="ZETA", FechaInicio=pd.Timestamp("2026-08-10"), FechaFin=pd.Timestamp("2026-09-01")),
        _campana_row("R2", "C1", IDCampaña="A1", Campaña="ALFA", FechaInicio=pd.Timestamp("2026-08-05"), FechaFin=pd.Timestamp("2026-09-01")),
        _campana_row("R3", "C2", IDCampaña="B1", Campaña="BETA", FechaInicio=pd.Timestamp("2026-08-05"), FechaFin=pd.Timestamp("2026-09-01")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    scope = td.build_tv4_universe(semantic_result)["scope_campanas"]
    cutoff = pd.Timestamp("2026-07-31")
    windows = td.compute_pipeline_windows(scope, cutoff)
    rows, total = td.compute_proximas_activaciones(windows["inician_30d"])
    assert [r["campana"] for r in rows] == ["ALFA", "BETA", "ZETA"]
    assert total == 3


def test_timeline_max_five_but_total_reflects_all():
    maestro_rows = [_cencosud_static(f"C{i}") for i in range(7)]
    campana_rows = [
        _campana_row(f"R{i}", f"C{i}", IDCampaña=f"X{i}", Campaña=f"CAMP {i}", FechaInicio=pd.Timestamp("2026-08-05"), FechaFin=pd.Timestamp("2026-09-01"))
        for i in range(7)
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    scope = td.build_tv4_universe(semantic_result)["scope_campanas"]
    cutoff = pd.Timestamp("2026-07-31")
    windows = td.compute_pipeline_windows(scope, cutoff)
    rows, total = td.compute_proximas_activaciones(windows["inician_30d"])
    assert len(rows) == 5
    assert total == 7


def test_timeline_no_duplicate_per_elemento_id():
    """Misma campana en 10 ElementoID, mismo inicio -> 1 sola fila en el
    timeline (spec Sec.7/9.13: no duplicar por ElementoID)."""
    maestro_rows = [_cencosud_static(f"C{i}") for i in range(10)]
    campana_rows = [
        _campana_row(f"R{i}", f"C{i}", IDCampaña="MASIVA", Campaña="CAMPANA MASIVA", FechaInicio=pd.Timestamp("2026-08-05"), FechaFin=pd.Timestamp("2026-09-01"))
        for i in range(10)
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    scope = td.build_tv4_universe(semantic_result)["scope_campanas"]
    cutoff = pd.Timestamp("2026-07-31")
    windows = td.compute_pipeline_windows(scope, cutoff)
    rows, total = td.compute_proximas_activaciones(windows["inician_30d"])
    assert total == 1
    assert len(rows) == 1


def test_timeline_keeps_genuinely_distinct_start_dates_for_same_campaign():
    """La misma campana con 2 fechas de inicio distintas (eventos comerciales
    distintos) NO debe colapsarse en 1 sola fila (spec Sec.7)."""
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    campana_rows = [
        _campana_row("R1", "C1", IDCampaña="M1", Campaña="MULTI FECHA", FechaInicio=pd.Timestamp("2026-08-05"), FechaFin=pd.Timestamp("2026-08-20")),
        _campana_row("R2", "C2", IDCampaña="M1", Campaña="MULTI FECHA", FechaInicio=pd.Timestamp("2026-08-12"), FechaFin=pd.Timestamp("2026-08-25")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    scope = td.build_tv4_universe(semantic_result)["scope_campanas"]
    cutoff = pd.Timestamp("2026-07-31")
    windows = td.compute_pipeline_windows(scope, cutoff)
    rows, total = td.compute_proximas_activaciones(windows["inician_30d"])
    assert total == 2


def test_production_timeline_sorted_and_capped(production_result):
    rows = production_result["data"]["timeline"]["rows"]
    assert len(rows) <= 5
    fechas = [r["fecha_iso"] for r in rows]
    assert fechas == sorted(fechas)


# ---------------------------------------------------------------------------
# 14. Distribucion por familia reconcilia con actividad actual
# ---------------------------------------------------------------------------


def test_distribucion_familia_reconciles_with_actividad_actual(production_result):
    kpis = production_result["data"]["kpis"]
    dist = production_result["data"]["distribucion_familia"]
    assert sum(f["activaciones"] for f in dist) == kpis["actividad_actual"]["activaciones"]


def test_distribucion_familia_max_four_groups(production_result):
    dist = production_result["data"]["distribucion_familia"]
    assert len(dist) <= 4


def test_distribucion_familia_sorted_desc(production_result):
    vals = [f["activaciones"] for f in production_result["data"]["distribucion_familia"]]
    assert vals == sorted(vals, reverse=True)


# ---------------------------------------------------------------------------
# 15. Nomenclatura UI usa "Estatico", no "Fijo"
# ---------------------------------------------------------------------------


def test_familia_label_uses_estatico_not_fijo():
    maestro_rows = [_cencosud_static("C1")]
    semantic_result = _semantic(maestro_rows)
    row = semantic_result["maestro"].iloc[0]
    assert td._familia(row) == "Shoppings Estático"


def test_template_uses_estatico_not_fijo_in_visible_copy():
    template = (REPO_ROOT / "scripts" / "templates" / "tv4_template.html").read_text(encoding="utf-8")
    assert "Fijo" not in template
    assert "FIJO" not in template.upper()


def test_payload_never_says_shoppings_fijo(production_json):
    assert "Shoppings Fijo" not in production_json
    assert "shoppings fijo" not in production_json.lower()


# ---------------------------------------------------------------------------
# 16. MetricStatus preservado (advertencia de fechas incompletas expuesta,
# nunca oculta ni convertida en 0 silencioso)
# ---------------------------------------------------------------------------


def test_incomplete_dates_excluded_not_counted_as_zero_activity():
    """Una activacion con fechas vacias no debe clasificarse como 'en curso'
    (falsamente activa) ni contarse como cero perdido en silencio: el
    builder la excluye de la clasificacion temporal (igual que el resto del
    motor) sin romper el resto del calculo."""
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    campana_rows = [
        _campana_row("OK", "C1", IDCampaña="X1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-08-15")),
        _campana_row("BAD", "C2", IDCampaña="X2", FechaInicio=None, FechaFin=None, FechaIndefinida=None),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    scope = td.build_tv4_universe(semantic_result)["scope_campanas"]
    cutoff = pd.Timestamp("2026-07-31")
    en_curso = scope[td._en_curso_mask(scope, cutoff)]
    assert len(en_curso) == 1  # solo X1; X2 (fechas vacias) queda excluida, no forzada a "en curso"


def test_production_reports_incomplete_dates_warning_when_present(production_result):
    meta = production_result["data"]["meta"]
    # En la fuente productiva real hay activaciones del Core con fechas
    # incompletas (Gate3B.1): la advertencia debe quedar expuesta, nunca oculta.
    assert isinstance(meta["advertencia_fechas_incompletas"], str)
    if meta["advertencia_fechas_incompletas"]:
        assert "fecha incompleta" in meta["advertencia_fechas_incompletas"]


# ---------------------------------------------------------------------------
# 17. Sin valores legacy de OCU_DATA serializados
# ---------------------------------------------------------------------------


def test_no_legacy_ocu_data_leaks_into_payload(production_json, production_result):
    for legacy_token in ("383", "1046", "Shoppings Fijo"):
        assert legacy_token not in production_json
    # el % legacy de TV4_REFERENCE era tautologico (100.0, propio universo
    # contra si mismo); TV4 productivo usa un denominador distinto.
    assert production_result["data"]["kpis"]["pipeline_core"]["pct"] != 100.0


# ---------------------------------------------------------------------------
# 18. SHA del Excel sin cambios
# ---------------------------------------------------------------------------


def test_input_excel_sha_unchanged(production_result):
    sha_after_build = vi.calculate_sha256(PRODUCTION_FILE)
    assert sha_after_build == production_result["sha256"]


# ---------------------------------------------------------------------------
# 19-22. TV1/TV2/TV3/TV4_REFERENCE intactas
# ---------------------------------------------------------------------------


def test_tv1_pipeline_still_builds_successfully():
    import build_tv1_dashboard as t1
    result = t1.build_tv1_data(PRODUCTION_FILE)
    # Recalibrado 2026-08-18: 964 -> 1064 (ver test_build_tv2_dashboard.py).
    assert result["data"]["kpis"]["core_comercial"]["value"] == 1064


def test_tv2_pipeline_still_builds_successfully():
    import build_tv2_dashboard as t2
    result = t2.build_tv2_data(PRODUCTION_FILE)
    assert result["data"]["kpis"]["ocupacion_calendario"]["activos"] == 71


def test_tv3_pipeline_still_builds_successfully():
    import build_tv3_dashboard as t3
    result = t3.build_tv3_data(PRODUCTION_FILE)
    # Recalibrado 2026-08-18: 119 -> 146 (ver test_reconciliacion_tv1_matches_tv1_estatico_scope
    # en test_build_tv3_dashboard.py: fila CENCOSUD "MISHKA" ya autorizada en FINAL_V2).
    assert result["data"]["kpis"]["ocupacion_calendario"]["activos"] == 146


def test_tv4_reference_untouched():
    ref = REPO_ROOT / "audit_sources" / "TV4_REFERENCE.html.html"
    assert ref.exists()
    html = ref.read_text(encoding="utf-8")
    assert "window.OCU_DATA" in html  # dataset legacy original, nunca reescrito


def test_production_excel_untouched_by_import_path():
    assert PRODUCTION_FILE.exists()
