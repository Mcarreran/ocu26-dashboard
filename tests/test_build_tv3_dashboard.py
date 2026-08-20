"""Pruebas de negocio para scripts/build_tv3_dashboard.py (dashboard TV3
OCU26, Core Comercial Estatico = Shoppings Estatico + AA2000 Estatico +
Pilar Frontlight).

No modifica scripts/validate_input.py, scripts/transform_data.py,
scripts/semantic_model.py, scripts/metrics_engine.py,
config/business_semantics.json, input/OCU26_BASE_DATOS.xlsx, ni ningun
archivo productivo de TV1/TV2 (build_tv1_dashboard.py, build_tv2_dashboard.py,
tv1_template.html, tv2_template.html, tv1.html, tv2.html,
test_build_tv1_dashboard.py, test_build_tv2_dashboard.py,
TV1_REFERENCE.html, TV2_REFERENCE.html). Los fixtures sinteticos usan la
CONFIGURACION REAL (sm.load_config()), mismo patron que
test_build_tv2_dashboard.py.
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
import build_tv3_dashboard as td  # noqa: E402
from metrics_engine import MetricsEngine  # noqa: E402
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


def _remeros_static(elemento_id: str, ubicacion: str = "REMEROS", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Shoppings Estático", Subcircuito="REMEROS", Ubicacion=ubicacion,
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático",
        CapacidadSlotsReel=0, SegundosDia=0,
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


def _pilar_frontlight(elemento_id: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Pilar", Subcircuito="X", Ubicacion="PILAR",
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


def _cencomedia(elemento_id: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Supermercados", Subcircuito="X", Ubicacion="MARTINEZ",
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Flexible gráfico",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


@pytest.fixture(scope="module")
def production_result():
    return td.build_tv3_data(PRODUCTION_FILE)


@pytest.fixture(scope="module")
def production_json(production_result):
    return json.dumps(production_result["data"], ensure_ascii=False)


# ---------------------------------------------------------------------------
# 1. Payload contiene solo TV3
# ---------------------------------------------------------------------------


def test_payload_top_level_keys_are_tv3_only(production_result):
    assert set(production_result["data"].keys()) == {
        "meta", "kpis", "familias", "ranking_shoppings", "soportes_top", "evolution", "insights",
    }


def test_payload_contains_no_other_tv_datasets(production_json):
    for token in ("tv1_data", "tv2_data", "tv4", "tv5", "tv6", "ocu_data"):
        assert token not in production_json.lower()


def test_no_legacy_ocu_data_leaks_into_payload(production_json):
    for legacy_token in ("533", "588", "1444", "2770", "P.PILAR"):
        assert legacy_token not in production_json


# ---------------------------------------------------------------------------
# 2-5. APSA / London / YPF / Cencomedia excluidos
# ---------------------------------------------------------------------------


def test_apsa_excluded_from_tv3_universe(production_result):
    assert "APSA" not in production_result["universe"]["circuitos"]


def test_london_excluded_from_tv3_universe(production_result):
    assert "LONDON_SUPPLY" not in production_result["universe"]["circuitos"]


def test_ypf_excluded_from_tv3_universe(production_result):
    assert "YPF" not in production_result["universe"]["circuitos"]


def test_cencomedia_excluded_from_tv3_universe(production_result):
    assert "CENCOMEDIA" not in production_result["universe"]["circuitos"]


def test_payload_contains_no_excluded_circuits(production_json):
    upper = production_json.upper()
    assert "APSA" not in upper
    assert "LONDON" not in upper


def test_apsa_elements_not_in_synthetic_tv3_universe():
    semantic_result = _semantic([_cencosud_static("C1"), _apsa_static("A1")])
    universe = td.build_tv3_universe(semantic_result)
    assert "A1" not in universe["element_ids"]
    assert "C1" in universe["element_ids"]


def test_london_elements_not_in_synthetic_tv3_universe():
    semantic_result = _semantic([_cencosud_static("C1"), _london_static("L1")])
    universe = td.build_tv3_universe(semantic_result)
    assert "L1" not in universe["element_ids"]


def test_ypf_elements_not_in_synthetic_tv3_universe():
    semantic_result = _semantic([_cencosud_static("C1"), _ypf_static("Y1")])
    universe = td.build_tv3_universe(semantic_result)
    assert "Y1" not in universe["element_ids"]


def test_cencomedia_elements_not_in_synthetic_tv3_universe():
    semantic_result = _semantic([_cencosud_static("C1"), _cencomedia("M1")])
    universe = td.build_tv3_universe(semantic_result)
    assert "M1" not in universe["element_ids"]


def test_digital_elements_not_in_synthetic_tv3_universe():
    """Medio=Digital nunca entra al Core Estatico, ni siquiera de circuitos
    que si pertenecen al scope (ej. Pantalla Led no es Core Estatico)."""
    semantic_result = _semantic([_cencosud_static("C1"), _pantalla_led("P1")])
    universe = td.build_tv3_universe(semantic_result)
    assert "P1" not in universe["element_ids"]
    assert "C1" in universe["element_ids"]


# ---------------------------------------------------------------------------
# 6. Ocupacion calendario usa DISTINCT ElementoID
# ---------------------------------------------------------------------------


def test_ocupacion_calendario_uses_distinct_elementoid():
    maestro_rows = [_cencosud_static("C1")]
    campana_rows = [
        _campana_row("CA1", "C1", IDCampaña="CAMP-A", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-05")),
        _campana_row("CA2", "C1", IDCampaña="CAMP-B", FechaInicio=pd.Timestamp("2026-07-10"), FechaFin=pd.Timestamp("2026-07-15")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    kpi1 = td.compute_kpi1_ocupacion_calendario(engine, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert kpi1["activos"] == 1  # 2 campanas en el mismo elemento cuentan 1 activo, no 2
    assert kpi1["elegibles"] == 1


def test_no_duplicate_counting_across_families():
    maestro_rows = [_cencosud_static("C1"), _remeros_static("R1"), _aa2000_static("A1"), _pilar_frontlight("PF1")]
    semantic_result = _semantic(maestro_rows)
    engine = MetricsEngine(semantic_result)
    period, previous = ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30")
    kpi1 = td.compute_kpi1_ocupacion_calendario(engine, period, previous)
    assert kpi1["elegibles"] == 4  # cada elemento cuenta una sola vez


# ---------------------------------------------------------------------------
# 7-8. Reconciliacion familias = Core
# ---------------------------------------------------------------------------


def test_family_elegibles_reconcile_with_core(production_result):
    kpis = production_result["data"]["kpis"]
    fam = production_result["data"]["familias"]
    total_elegibles = fam["shoppings"]["elegibles"] + fam["aa2000"]["elegibles"] + fam["pilar_frontlight"]["elegibles"]
    assert total_elegibles == kpis["ocupacion_calendario"]["elegibles"]


def test_family_activos_reconcile_with_core(production_result):
    kpis = production_result["data"]["kpis"]
    fam = production_result["data"]["familias"]
    total_activos = fam["shoppings"]["activos"] + fam["aa2000"]["activos"] + fam["pilar_frontlight"]["activos"]
    assert total_activos == kpis["ocupacion_calendario"]["activos"]


def test_kpi_shape_matches_tv2_system(production_result):
    """Mismo sistema que TV2 (spec TV3 Sec.3): activos/elegibles/pct/delta,
    no el ratio dia-based de TV1."""
    kpis = production_result["data"]["kpis"]
    for kpi in (kpis["ocupacion_calendario"], kpis["shoppings_estatico"], kpis["aa2000_estatico"]):
        assert set(kpi.keys()) == {"activos", "elegibles", "anterior_activos", "pct_actual", "pct_anterior", "delta_pp"}
        assert kpi["pct_actual"] == round(kpi["activos"] / kpi["elegibles"] * 100.0, 1)


# ---------------------------------------------------------------------------
# 9. AA2000 - metrica valida, cero real no inventado
# ---------------------------------------------------------------------------


def test_aa2000_zero_is_real_not_invented(production_result):
    """AA2000 Estatico tiene CompletitudMaestro=PARCIAL (faltan Mendoza/
    Cordoba): elementos_con_actividad/registrados no requieren cobertura
    completa, asi que el 0 real de la produccion viene de una query valida,
    nunca de forzar NO_APLICA a cero."""
    aa2000 = production_result["data"]["kpis"]["aa2000_estatico"]
    assert aa2000["elegibles"] > 0
    assert aa2000["activos"] >= 0
    assert aa2000["pct_actual"] is not None


def test_aa2000_synthetic_activity_is_detected():
    maestro_rows = [_aa2000_static("A1")]
    campana_rows = [_campana_row("CA1", "A1", IDCampaña="CAMP-A", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-05"))]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    kpi3 = td.compute_kpi3_aa2000_estatico(engine, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert kpi3["activos"] == 1
    assert kpi3["elegibles"] == 1


# ---------------------------------------------------------------------------
# 10-11. Disponibles correctos
# ---------------------------------------------------------------------------


def test_disponibles_equals_elegibles_menos_activos(production_result):
    kpis = production_result["data"]["kpis"]
    oc, disp = kpis["ocupacion_calendario"], kpis["disponibles"]
    assert disp["disponibles"] == oc["elegibles"] - oc["activos"]


def test_disponibles_pct_is_correct(production_result):
    disp = production_result["data"]["kpis"]["disponibles"]
    esperado = round(disp["disponibles"] / disp["elegibles"] * 100.0, 1)
    assert disp["pct_actual"] == esperado


# ---------------------------------------------------------------------------
# 12-13. Ranking ordenado y con tope
# ---------------------------------------------------------------------------


def test_ranking_shoppings_sorted_by_ocupacion_desc(production_result):
    rows = production_result["data"]["ranking_shoppings"]
    pcts = [r["ocup_pct"] for r in rows]
    assert pcts == sorted(pcts, reverse=True)


def test_ranking_shoppings_max_five(production_result):
    assert len(production_result["data"]["ranking_shoppings"]) <= 5


def test_ranking_shoppings_sort_tiebreak_activos_desc_then_name_asc():
    maestro_rows = [
        _cencosud_static("C1", ubicacion="UNICENTER"),
        _cencosud_static("C2", ubicacion="P.OESTE"),
        _remeros_static("R1", ubicacion="REMEROS"),
    ]
    campana_rows = [
        _campana_row("CA1", "C1", IDCampaña="X1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
        _campana_row("CA2", "C2", IDCampaña="X2", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
        _campana_row("CA3", "R1", IDCampaña="X3", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    ranking = td.compute_ranking_shoppings(engine, ("2026-07-01", "2026-07-31"))
    # los 3 tienen 100% ocupacion (1/1): orden final por nombre asc.
    assert [r["sitio"] for r in ranking] == sorted(r["sitio"] for r in ranking)


# ---------------------------------------------------------------------------
# 14. Soportes Top 3 maximo
# ---------------------------------------------------------------------------


def test_soportes_top_max_three(production_result):
    assert len(production_result["data"]["soportes_top"]) <= 3


def test_soportes_top_sorted_desc(production_result):
    rows = production_result["data"]["soportes_top"]
    vals = [r["activos"] for r in rows]
    assert vals == sorted(vals, reverse=True)


def test_soportes_top_only_positive_activity():
    maestro_rows = [_cencosud_static("C1", Descripcion="Frontlight Ingreso"), _remeros_static("R1", Descripcion="Chupete Backlight")]
    campana_rows = [_campana_row("CA1", "C1", IDCampaña="X1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-05"))]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    soportes = td.compute_soportes_top(engine, ("2026-07-01", "2026-07-31"))
    assert sum(r["activos"] for r in soportes) == 1
    assert all(r["activos"] > 0 for r in soportes)


# ---------------------------------------------------------------------------
# 15-16. Evolucion Ene-Jul unicamente, sin futuro
# ---------------------------------------------------------------------------


def test_evolution_excludes_future_months(production_result):
    ev = production_result["data"]["evolution"]
    assert len(ev["meses"]) == td.REPORT_MONTH
    assert ev["meses"][-1] == "Jul"
    assert "Ago" not in ev["meses"]


def test_evolution_has_single_core_series(production_result):
    ev = production_result["data"]["evolution"]
    assert set(ev.keys()) == {"meses", "core"}
    assert len(ev["core"]) == len(ev["meses"])


def test_previous_month_is_june(production_result):
    assert td._previous_month(2026, 7) == (2026, 6)
    assert production_result["data"]["meta"]["previous_month_label"] == "Junio"


# ---------------------------------------------------------------------------
# 17. Deltas en puntos porcentuales, no % relativo
# ---------------------------------------------------------------------------


def test_delta_is_percentage_points_not_relative_percent():
    maestro_rows = [_cencosud_static(f"C{i}") for i in range(10)]
    campana_rows = []
    for i in range(5):
        campana_rows.append(_campana_row(f"J{i}", f"C{i}", IDCampaña=f"CJ{i}", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10")))
    for i in range(4):
        campana_rows.append(_campana_row(f"N{i}", f"C{i}", IDCampaña=f"CN{i}", FechaInicio=pd.Timestamp("2026-06-05"), FechaFin=pd.Timestamp("2026-06-10")))
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    kpi1 = td.compute_kpi1_ocupacion_calendario(engine, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert kpi1["pct_actual"] == 50.0
    assert kpi1["pct_anterior"] == 40.0
    assert kpi1["delta_pp"] == 10.0  # 50 - 40, NUNCA (50-40)/40*100=25


# ---------------------------------------------------------------------------
# 18. MetricStatus preservado (reconciliacion TV1)
# ---------------------------------------------------------------------------


def test_reconciliacion_tv1_preserves_metric_status(production_result):
    rec = production_result["reconciliacion_tv1"]
    assert rec["status"] in ("OK", "PARTIAL", "NO_APLICA", "REQUIERE_CONFIRMACION")
    assert rec["ocupacion_dia_based_pct"] is not None


def test_reconciliacion_tv1_matches_tv1_estatico_scope(production_result):
    """La ocupacion dia-based del scope reconciliado (Shoppings+Pilar,
    unico subconjunto con CoberturaCatalogo/CompletitudMaestro completos)
    debe coincidir con el valor ya validado en TV1 para Julio 2026.
    Recalibrado 2026-08-18: 29.4% -> 35.7%. circuitos_elegibles y el
    denominador_dias (11.842) no cambiaron -- el numerador_dias subio de
    3.482 a 4.227 porque la fila CENCOSUD "MISHKA" (PPIL-BN-C-4B), ya
    autorizada en FINAL_V2 y nunca antes promovida, quedo activa en el
    periodo al promover la base junto con YPF."""
    rec = production_result["reconciliacion_tv1"]
    assert rec["ocupacion_dia_based_pct"] == 35.7
    assert set(rec["circuitos_elegibles"]) == {"CENCOSUD", "REMEROS", "PILAR_FRONTLIGHT"}


# ---------------------------------------------------------------------------
# 19. SHA del Excel sin cambios
# ---------------------------------------------------------------------------


def test_input_excel_sha_unchanged(production_result):
    sha_after_build = vi.calculate_sha256(PRODUCTION_FILE)
    assert sha_after_build == production_result["sha256"]


# ---------------------------------------------------------------------------
# 20-22. TV1/TV2/referencias intactas
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


def test_tv3_reference_untouched():
    ref = REPO_ROOT / "audit_sources" / "TV3_REFERENCE.html.html"
    assert ref.exists()
    html = ref.read_text(encoding="utf-8")
    assert "window.OCU_DATA" in html  # dataset legacy original, nunca reescrito


def test_tv1_and_tv2_references_untouched():
    for name in ("TV1_REFERENCE.html.html", "TV2_REFERENCE.html.html"):
        ref = REPO_ROOT / "audit_sources" / name
        assert ref.exists()
        assert "window.OCU_DATA" in ref.read_text(encoding="utf-8")


# ---------------------------------------------------------------------------
# 23. Nomenclatura visible "Estático", no "Fijo"
# ---------------------------------------------------------------------------


def test_template_uses_estatico_not_fijo_in_visible_copy():
    template = (REPO_ROOT / "scripts" / "templates" / "tv3_template.html").read_text(encoding="utf-8")
    assert "Core Comercial Est" in template  # "Core Comercial Estático"
    # "Fijo" (nomenclatura legacy visible) no debe aparecer en el template productivo.
    assert "Fijo" not in template
    assert "FIJO" not in template.upper()
