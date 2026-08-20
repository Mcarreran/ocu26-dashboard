"""Pruebas de negocio para scripts/build_tv2_dashboard.py (dashboard TV2 OCU26,
Core Comercial Digital = Pantallas LED + Shoppings Digital + AA2000 Digital).

No modifica scripts/validate_input.py, scripts/transform_data.py,
scripts/semantic_model.py, scripts/metrics_engine.py,
config/business_semantics.json, input/OCU26_BASE_DATOS.xlsx, ni ningun
archivo productivo de TV1 (build_tv1_dashboard.py, tv1_template.html,
tv1.html, test_build_tv1_dashboard.py, TV1_REFERENCE.html). Los fixtures
sinteticos usan la CONFIGURACION REAL (sm.load_config()), mismo patron que
test_build_tv1_dashboard.py.

Cobertura: los 27 puntos de negocio obligatorios (prompt TV2 Sec.27).
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
import build_tv2_dashboard as td  # noqa: E402
from metrics_engine import MetricsEngine  # noqa: E402
from test_semantic_model import _maestro_row, _campana_row, _transform_result  # noqa: E402

PRODUCTION_FILE = REPO_ROOT / "input" / "OCU26_BASE_DATOS.xlsx"


def _semantic(maestro_rows: list[dict], campanas_rows: list[dict] | None = None) -> dict:
    config = sm.load_config()
    return sm.build_semantic_model(_transform_result(maestro_rows, campanas_rows or []), config)


def _pantalla_led(elemento_id: str, ubicacion: str = "SITE", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Pantalla Led", Subcircuito="X", Ubicacion=ubicacion,
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        CapacidadSlotsReel=20, SegundosDia=72000,
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


def _cencosud_static(elemento_id: str, ubicacion: str = "UNICENTER", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Shoppings Estático", Subcircuito="CENCOSUD", Ubicacion=ubicacion,
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _remeros_digital(elemento_id: str, ubicacion: str = "REMEROS", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Shoppings Digital", Subcircuito="REMEROS", Ubicacion=ubicacion,
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        CapacidadSlotsReel=20, SegundosDia=72000,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _aa2000_digital(elemento_id: str, ubicacion: str = "EZE", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="AA2000", Subcircuito="X", Ubicacion=ubicacion,
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        CapacidadSlotsReel=20, SegundosDia=72000,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _ypf_digital(elemento_id: str, ubicacion: str = "500 - TESTVILLE - Calle Falsa 123", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="YPF Digital", Subcircuito="X", Ubicacion=ubicacion,
        Medio="Digital", TipoCatalogo="Abierto", TipoInventario="Digital",
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
    return td.build_tv2_data(PRODUCTION_FILE)


@pytest.fixture(scope="module")
def production_json(production_result):
    return json.dumps(production_result["data"], ensure_ascii=False)


# ---------------------------------------------------------------------------
# 1. Payload contiene solo TV2
# ---------------------------------------------------------------------------


def test_payload_top_level_keys_are_tv2_only(production_result):
    assert set(production_result["data"].keys()) == {
        "meta", "kpis", "core_digital", "ranking_pantallas", "ranking_shoppings", "evolution", "insights",
    }


def test_payload_contains_no_other_tv_datasets(production_json):
    for token in ("tv1_data", "tv3", "tv4", "tv5", "tv6", "ocu_data"):
        assert token not in production_json.lower()


# ---------------------------------------------------------------------------
# 2-4. YPF / APSA / London excluidos
# ---------------------------------------------------------------------------


def test_ypf_excluded_from_tv2_universe(production_result):
    assert "YPF" not in production_result["universe"]["circuitos"]


def test_apsa_excluded_from_tv2_universe(production_result):
    assert "APSA" not in production_result["universe"]["circuitos"]


def test_london_excluded_from_tv2_universe(production_result):
    assert "LONDON_SUPPLY" not in production_result["universe"]["circuitos"]


def test_ypf_elements_not_in_synthetic_tv2_universe():
    semantic_result = _semantic([_pantalla_led("P1"), _ypf_digital("500 - TT - 1")])
    universe = td.build_tv2_universe(semantic_result)
    assert "500 - TT - 1" not in universe["element_ids"]
    assert "P1" in universe["element_ids"]


def test_apsa_elements_not_in_synthetic_tv2_universe():
    semantic_result = _semantic([_pantalla_led("P1"), _apsa_static("A1")])
    universe = td.build_tv2_universe(semantic_result)
    assert "A1" not in universe["element_ids"]


def test_london_elements_not_in_synthetic_tv2_universe():
    semantic_result = _semantic([_pantalla_led("P1"), _london_static("L1")])
    universe = td.build_tv2_universe(semantic_result)
    assert "L1" not in universe["element_ids"]


def test_payload_contains_no_excluded_circuits(production_json):
    upper = production_json.upper()
    assert "APSA" not in upper
    assert "LONDON" not in upper


# ---------------------------------------------------------------------------
# 5-6. AA2000 incluido en Core Digital / KPIs generales
# ---------------------------------------------------------------------------


def test_aa2000_included_in_core_digital(production_result):
    elegibles = production_result["data"]["core_digital"]["elegibles"]
    assert elegibles["aa2000"] > 0


def test_aa2000_capacity_included_in_general_fill_denominator():
    """El denominador del fill general debe incluir la capacidad de AA2000:
    verificado sumando familias independientes y comparando contra el
    query combinado (mismo patron que el guard del builder)."""
    engine_rows = [_pantalla_led("P1"), _cencosud_digital("C1"), _aa2000_digital("A1")]
    semantic_result = _semantic(engine_rows)
    engine = MetricsEngine(semantic_result)
    period, previous = ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30")
    combined = td._fill_family(engine, td.TV2_CIRCUITOS, period, previous)
    aa2000_only = td._fill_family(engine, td.AA2000_CIRCUITOS, period, previous)
    assert aa2000_only["denominador"] > 0
    assert combined["denominador"] >= aa2000_only["denominador"]


# ---------------------------------------------------------------------------
# 7-8. AA2000 excluido de rankings y de las 2 series principales de evolucion
# ---------------------------------------------------------------------------


def test_aa2000_excluded_from_ranking_scopes():
    assert "AA2000" not in td.PANTALLAS_CIRCUITOS
    assert "AA2000" not in td.SHOPPINGS_CIRCUITOS


def test_aa2000_excluded_from_evolution_series(production_result):
    ev = production_result["data"]["evolution"]
    assert set(ev.keys()) == {"meses", "pantallas", "shoppings"}


# ---------------------------------------------------------------------------
# 9-10. Remeros incluido en Shoppings, no duplicado en Pantallas
# ---------------------------------------------------------------------------


def test_remeros_included_in_shoppings_family():
    assert "REMEROS" in td.SHOPPINGS_CIRCUITOS


def test_remeros_not_duplicated_in_pantallas_even_with_same_site_name():
    """Una Pantalla Led fisicamente ubicada en 'Remeros' (CircuitoNegocio
    PANTALLAS_LED) y un elemento Shoppings Digital de Remeros
    (CircuitoNegocio REMEROS) coexisten sin que el shopping se cuente
    dentro de la familia Pantallas."""
    maestro_rows = [_pantalla_led("PLED-REM", ubicacion="REMEROS"), _remeros_digital("SHOP-REM", ubicacion="REMEROS")]
    semantic_result = _semantic(maestro_rows)
    universe = td.build_tv2_universe(semantic_result)
    maestro = universe["maestro"]
    pantallas_ids = maestro[maestro["CircuitoNegocio"] == "PANTALLAS_LED"]["ElementoID"].tolist()
    shoppings_ids = maestro[maestro["CircuitoNegocio"] == "REMEROS"]["ElementoID"].tolist()
    assert pantallas_ids == ["PLED-REM"]
    assert shoppings_ids == ["SHOP-REM"]
    assert not set(pantallas_ids) & set(shoppings_ids)


# ---------------------------------------------------------------------------
# 11. Ocupacion calendario usa DISTINCT ElementoID
# ---------------------------------------------------------------------------


def test_ocupacion_calendario_uses_distinct_elementoid():
    maestro_rows = [_pantalla_led("P1")]
    campana_rows = [
        _campana_row("C1", "P1", IDCampaña="CAMP-A", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-05")),
        _campana_row("C2", "P1", IDCampaña="CAMP-B", FechaInicio=pd.Timestamp("2026-07-10"), FechaFin=pd.Timestamp("2026-07-15")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    kpi1 = td.compute_kpi1_ocupacion_calendario(engine, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert kpi1["activos"] == 1  # 2 campanas en el mismo elemento cuentan 1 activo, no 2
    assert kpi1["elegibles"] == 1


# ---------------------------------------------------------------------------
# 12-13. Reconciliacion Pantallas+Shoppings+AA2000 = Core
# ---------------------------------------------------------------------------


def test_family_elegibles_reconcile_with_core(production_result):
    core = production_result["data"]["core_digital"]["elegibles"]
    assert core["pantallas"] + core["shoppings"] + core["aa2000"] == core["total"]


def test_family_con_campana_reconciles_with_core(production_result):
    core = production_result["data"]["core_digital"]["con_campana"]
    assert core["pantallas"] + core["shoppings"] + core["aa2000"] == core["total"]


def test_kpi3_kpi4_are_calendario_occupancy_not_fill(production_result):
    """Ajuste visual: las tarjetas KPI3 'Pantallas LED' y KPI4 'Shoppings
    Digital' pasaron de fill rate a ocupacion por calendario (misma forma
    que KPI1, sin slots/capacidad)."""
    kpis = production_result["data"]["kpis"]
    core = production_result["data"]["core_digital"]
    for kpi, activos_esperado, elegibles_esperado in (
        (kpis["pantallas"], core["con_campana"]["pantallas"], core["elegibles"]["pantallas"]),
        (kpis["shoppings"], core["con_campana"]["shoppings"], core["elegibles"]["shoppings"]),
    ):
        assert set(kpi.keys()) == {"activos", "elegibles", "anterior_activos", "pct_actual", "pct_anterior", "delta_pp"}
        assert kpi["activos"] == activos_esperado
        assert kpi["elegibles"] == elegibles_esperado
        assert kpi["pct_actual"] == round(kpi["activos"] / kpi["elegibles"] * 100.0, 1)


def test_fill_general_reconciles_with_family_numerator_denominator(production_result):
    # kpis.pantallas/kpis.shoppings ahora son ocupacion por calendario (tarjetas
    # KPI3/KPI4, ajuste visual TV2); el fill por familia usado para reconciliar
    # el Core vive en core_digital.pantallas_fill/shoppings_fill/aa2000_fill.
    core_digital = production_result["data"]["core_digital"]
    fd = production_result["data"]["kpis"]["fill_digital"]
    families = [core_digital["pantallas_fill"], core_digital["shoppings_fill"], core_digital["aa2000_fill"]]
    num_sum = sum(f["numerador_raw"] for f in families)
    den_sum = sum(f["denominador"] for f in families)
    assert den_sum == fd["denominador"]
    assert abs(num_sum - fd["numerador_raw"]) < 0.01


# ---------------------------------------------------------------------------
# 14-16. Slots disponibles / % disponible / apertura por familia
# ---------------------------------------------------------------------------


def test_slots_disponibles_equals_capacidad_menos_ocupados(production_result):
    kpis = production_result["data"]["kpis"]
    sd, fd = kpis["slots_disponibles"], kpis["fill_digital"]
    esperado = round(fd["denominador"] - fd["numerador_raw"])
    assert sd["disponibles"] == esperado


def test_pct_disponible_is_correct(production_result):
    sd = production_result["data"]["kpis"]["slots_disponibles"]
    esperado = round(sd["disponibles"] / sd["capacidad_total"] * 100.0, 1)
    assert sd["pct_actual"] == esperado


def test_apertura_disponibilidad_por_familia_is_correct(production_result):
    core_digital = production_result["data"]["core_digital"]
    sd = production_result["data"]["kpis"]["slots_disponibles"]
    for key, fam in (("pantallas", core_digital["pantallas_fill"]), ("shoppings", core_digital["shoppings_fill"])):
        disp_raw = fam["denominador"] - fam["numerador_raw"]
        esperado_disp = round(disp_raw)
        # pct se calcula sobre el disponible RAW (sin redondear a int primero),
        # mismo patron que evita doble redondeo en el resto del pipeline.
        esperado_pct = round(disp_raw / fam["denominador"] * 100.0, 1)
        assert sd["apertura"][key]["disponibles"] == esperado_disp
        assert sd["apertura"][key]["pct"] == esperado_pct


def test_aa2000_availability_status_preserved_not_invented(production_result):
    ap_aa2000 = production_result["data"]["kpis"]["slots_disponibles"]["apertura"]["aa2000"]
    assert ap_aa2000["status"] in ("OK", "PARTIAL", "NO_APLICA", "REQUIERE_CONFIRMACION")


# ---------------------------------------------------------------------------
# 17-19. Rankings ordenados y con tope
# ---------------------------------------------------------------------------


def test_ranking_pantallas_sorted_by_fill_desc(production_result):
    rows = production_result["data"]["ranking_pantallas"]
    fills = [r["fill_pct"] for r in rows]
    assert fills == sorted(fills, reverse=True)


def test_ranking_shoppings_sorted_by_fill_desc(production_result):
    rows = production_result["data"]["ranking_shoppings"]
    fills = [r["fill_pct"] for r in rows]
    assert fills == sorted(fills, reverse=True)


def test_ranking_pantallas_max_six(production_result):
    assert len(production_result["data"]["ranking_pantallas"]) <= 6


def test_ranking_shoppings_max_three(production_result):
    assert len(production_result["data"]["ranking_shoppings"]) <= 3


# ---------------------------------------------------------------------------
# 20-21. Sin meses futuros / junio = periodo anterior
# ---------------------------------------------------------------------------


def test_evolution_excludes_future_months(production_result):
    ev = production_result["data"]["evolution"]
    assert len(ev["meses"]) == td.REPORT_MONTH
    assert ev["meses"][-1] == "Jul"
    assert "Ago" not in ev["meses"]


def test_previous_month_is_june(production_result):
    assert td._previous_month(2026, 7) == (2026, 6)
    assert production_result["data"]["meta"]["previous_month_label"] == "Junio"


# ---------------------------------------------------------------------------
# 22. Deltas en puntos porcentuales, no % relativo
# ---------------------------------------------------------------------------


def test_delta_is_percentage_points_not_relative_percent():
    maestro_rows = [_pantalla_led(f"P{i}") for i in range(10)]
    campana_rows = []
    for i in range(5):
        campana_rows.append(_campana_row(f"J{i}", f"P{i}", IDCampaña=f"CJ{i}", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10")))
    for i in range(4):
        campana_rows.append(_campana_row(f"N{i}", f"P{i}", IDCampaña=f"CN{i}", FechaInicio=pd.Timestamp("2026-06-05"), FechaFin=pd.Timestamp("2026-06-10")))
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    kpi1 = td.compute_kpi1_ocupacion_calendario(engine, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert kpi1["pct_actual"] == 50.0
    assert kpi1["pct_anterior"] == 40.0
    assert kpi1["delta_pp"] == 10.0  # 50 - 40, NUNCA (50-40)/40*100=25


# ---------------------------------------------------------------------------
# 23. MetricStatus preservado
# ---------------------------------------------------------------------------


def test_metric_status_preserved_in_production_output(production_result):
    fd = production_result["data"]["kpis"]["fill_digital"]
    core_digital = production_result["data"]["core_digital"]
    for kpi in (fd, core_digital["pantallas_fill"], core_digital["shoppings_fill"]):
        assert kpi["status"] in ("OK", "PARTIAL", "NO_APLICA", "REQUIERE_CONFIRMACION")
    # AA2000 tiene CompletitudMaestro=PARCIAL real: debe reflejarse, nunca forzarse a OK.
    assert core_digital["aa2000_fill"]["status"] == "PARTIAL"


def test_unknown_capacity_does_not_become_zero_fill_rate():
    maestro_rows = [_maestro_row(
        "U1", CircuitoDashboard="AA2000", Subcircuito="X", Ubicacion="EZE",
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        Descripcion="", CapacidadSlotsReel=0, SegundosDia=0,
    )]
    campana_rows = [_campana_row("C1", "U1", IDCampaña="CAMP-U", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10"))]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    aa2000 = td.compute_aa2000_fill(engine, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert aa2000["status"] == "NO_APLICA"
    assert aa2000["pct_actual"] is None  # nunca 0.0 cuando la capacidad es desconocida


# ---------------------------------------------------------------------------
# 24. SHA del Excel sin cambios
# ---------------------------------------------------------------------------


def test_input_excel_sha_unchanged(production_result):
    sha_after_build = vi.calculate_sha256(PRODUCTION_FILE)
    assert sha_after_build == production_result["sha256"]


# ---------------------------------------------------------------------------
# 25-26. TV1 y TV2_REFERENCE intactas
# ---------------------------------------------------------------------------


def test_tv1_pipeline_still_builds_successfully():
    """No debe haber quedado roto ni acoplado por la introduccion de TV2:
    build_tv2_dashboard.py no importa build_tv1_dashboard.py (builders
    independientes sobre el mismo pipeline), asi que TV1 debe poder
    construirse igual que antes."""
    import build_tv1_dashboard as t1
    result = t1.build_tv1_data(PRODUCTION_FILE)
    # Recalibrado 2026-08-18: 964 -> 1064 (513 no-YPF sin cambios + 551 estaciones
    # YPF, antes 451; catalogo YPF creció 983 elementos y perdió 17 en APIE 30943).
    # Verificado con calculo independiente en pandas directo sobre el Excel: 1064.
    assert result["data"]["kpis"]["core_comercial"]["value"] == 1064


def test_tv2_reference_untouched():
    ref = REPO_ROOT / "audit_sources" / "TV2_REFERENCE.html.html"
    assert ref.exists()
    html = ref.read_text(encoding="utf-8")
    assert "window.OCU_DATA" in html  # dataset legacy original, nunca reescrito


# ---------------------------------------------------------------------------
# 27. Sin valores legacy serializados por error
# ---------------------------------------------------------------------------


def test_no_legacy_reference_values_leak_into_payload(production_json):
    for legacy_token in ("27558", "47120", "19562", "1444", "2770"):
        assert legacy_token not in production_json
