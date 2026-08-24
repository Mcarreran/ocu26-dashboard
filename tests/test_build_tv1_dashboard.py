"""Pruebas de negocio para scripts/build_tv1_dashboard.py (dashboard TV1 OCU26).

No modifica scripts/validate_input.py, scripts/transform_data.py,
scripts/semantic_model.py, scripts/metrics_engine.py,
config/business_semantics.json ni input/OCU26_BASE_DATOS.xlsx. Los
fixtures sinteticos usan la CONFIGURACION REAL (sm.load_config()), mismo
patron que test_export_data.py, para ejercitar reglas de negocio reales
con datos controlados. Los tests de reconciliacion/payload corren contra
el archivo productivo real (module-scoped, se ejecuta una sola vez).

Cobertura: los 24 puntos de negocio obligatorios (prompt TV1 Sec.54).
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
import build_tv1_dashboard as td  # noqa: E402
from metrics_engine import MetricsEngine  # noqa: E402
from test_semantic_model import _maestro_row, _campana_row, _transform_result  # noqa: E402

PRODUCTION_FILE = REPO_ROOT / "input" / "OCU26_BASE_DATOS.xlsx"


def _semantic(maestro_rows: list[dict], campanas_rows: list[dict] | None = None) -> dict:
    config = sm.load_config()
    return sm.build_semantic_model(_transform_result(maestro_rows, campanas_rows or []), config)


def _digital_aa2000_unknown_capacity(elemento_id: str, **overrides) -> dict:
    """Elemento digital sin perfil de FormatoNegocio confirmado y sin
    capacidad legacy (CapacidadSlotsReel=0): fuerza SlotsComerciales=
    REQUIERE_CONFIRMACION (nunca 0)."""
    row = dict(
        CircuitoDashboard="AA2000", Subcircuito="X", Ubicacion="EZE",
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        Descripcion="", CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _static_cencosud(elemento_id: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Shoppings Estático", Subcircuito="CENCOSUD", Ubicacion="UNICENTER",
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _digital_cencosud(elemento_id: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Shoppings Digital", Subcircuito="CENCOSUD", Ubicacion="UNICENTER",
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        CapacidadSlotsReel=20, SegundosDia=72000,
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


def _ypf_digital(elemento_id: str, ubicacion: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="YPF Digital", Subcircuito="X", Ubicacion=ubicacion,
        Medio="Digital", TipoCatalogo="Abierto", TipoInventario="Digital",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _ypf_static(elemento_id: str, ubicacion: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="YPF Estático", Subcircuito="X", Ubicacion=ubicacion,
        Medio="Estático", TipoCatalogo="Abierto", TipoInventario="Físico estático",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


@pytest.fixture(scope="module")
def production_result():
    return td.build_tv1_data(PRODUCTION_FILE)


@pytest.fixture(scope="module")
def production_json(production_result):
    return json.dumps(production_result["data"], ensure_ascii=False)


# ---------------------------------------------------------------------------
# 1-2. APSA / London Supply excluidos
# ---------------------------------------------------------------------------


def test_apsa_excluded_from_tv1_universe(production_result):
    assert "APSA" not in production_result["universe"]["circuitos"]


def test_london_excluded_from_tv1_universe(production_result):
    assert "LONDON_SUPPLY" not in production_result["universe"]["circuitos"]


def test_apsa_elements_not_in_synthetic_universe():
    semantic_result = _semantic([_static_cencosud("C1"), _apsa_static("A1")])
    universe = td.build_tv1_universe(semantic_result)
    assert "A1" not in universe["element_ids"]
    assert "C1" in universe["element_ids"]


def test_london_elements_not_in_synthetic_universe():
    semantic_result = _semantic([_static_cencosud("C1"), _london_static("L1")])
    universe = td.build_tv1_universe(semantic_result)
    assert "L1" not in universe["element_ids"]
    assert "C1" in universe["element_ids"]


# ---------------------------------------------------------------------------
# 3-4. IDCampaña DISTINCT (una campaña en 50 elementos = 1 campaña)
# ---------------------------------------------------------------------------


def test_one_campaign_across_fifty_elements_counts_once():
    maestro_rows = [_static_cencosud(f"E{i}") for i in range(50)]
    campana_rows = [
        _campana_row(f"CARGA{i}", f"E{i}", IDCampaña="CAMP-X", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10"))
        for i in range(50)
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    n = td._distinct_campanas(engine, universe["element_ids"], "2026-07-01", "2026-07-31")
    assert n == 1


def test_campanas_ytd_uses_distinct_idcampana(production_result):
    n_distinct = production_result["data"]["kpis"]["campanas_unicas"]["ytd"]
    assert n_distinct > 0
    # sanity: el conteo distinct debe ser materialmente menor a las filas
    # crudas superpuestas en el mismo periodo (evidencia de que SI dedupe).
    assert n_distinct < 9503  # total de filas de CAMPANAS en el Excel fuente


# ---------------------------------------------------------------------------
# 5. Campañas YTD no se obtienen sumando meses
# ---------------------------------------------------------------------------


def test_ytd_campanas_not_sum_of_monthly_counts():
    maestro_rows = [_static_cencosud("E1")]
    campana_rows = [_campana_row("C1", "E1", IDCampaña="CAMP-LONG", FechaInicio=pd.Timestamp("2026-01-15"), FechaFin=pd.Timestamp("2026-07-15"))]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    element_ids = ["E1"]

    ytd = td._distinct_campanas(engine, element_ids, "2026-01-01", "2026-07-31")
    assert ytd == 1

    suma_mensual = 0
    for month in range(1, 8):
        start, end = td._period_bounds(2026, month)
        suma_mensual += td._distinct_campanas(engine, element_ids, start, end)
    assert suma_mensual == 7  # la misma campaña activa en los 7 meses
    assert ytd != suma_mensual  # YTD NO es la suma de los meses


# ---------------------------------------------------------------------------
# 6. ElementoID activo DISTINCT
# ---------------------------------------------------------------------------


def test_element_with_two_campaigns_counts_once_as_active():
    maestro_rows = [_static_cencosud("E1")]
    campana_rows = [
        _campana_row("C1", "E1", IDCampaña="CAMP-A", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-05")),
        _campana_row("C2", "E1", IDCampaña="CAMP-B", FechaInicio=pd.Timestamp("2026-07-10"), FechaFin=pd.Timestamp("2026-07-15")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    kpi3 = td.compute_kpi3_actividad(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"), core_value=1)
    assert kpi3["value"] == 1


# ---------------------------------------------------------------------------
# 7-9. YPF: participa en actividad TV1, NO en digital fill rate, serie independiente
# ---------------------------------------------------------------------------


def test_ypf_participates_in_tv1_activity(production_result):
    assert production_result["data"]["kpis"]["unidades_actividad"]["value"] > 0
    ypf_family = next(f for f in production_result["data"]["composition"]["familias"] if f["nombre"] == "YPF")
    assert ypf_family["total"] > 0


def test_ypf_excluded_from_digital_fill_universe(production_result):
    assert "YPF" not in production_result["universe"]["digital_circuitos"]


def test_ypf_is_independent_evolution_series(production_result):
    ev = production_result["data"]["evolution"]
    assert "ypf" in ev and "digital" in ev and "estatico" in ev
    assert len(ev["ypf"]) == len(ev["digital"]) == len(ev["estatico"])
    # Julio 2026: YPF tiene actividad propia que no esta mezclada dentro de "digital"
    assert ev["ypf"][-1] > 0
    assert ev["ypf"][-1] != ev["digital"][-1]


# ---------------------------------------------------------------------------
# 10-11. Cencomedia: actividad estática real, sin denominador falso
# ---------------------------------------------------------------------------


def test_cencomedia_excluded_from_tv1_universe(production_result):
    assert "CENCOMEDIA" not in production_result["universe"]["circuitos"]


def test_cencomedia_absent_from_composition_payload(production_result):
    """Decision 21/08/2026: Cencomedia no aparece en el payload de TV1, ni
    siquiera como familia en cero (distinto del diseño anterior)."""
    nombres = [f["nombre"] for f in production_result["data"]["composition"]["familias"]]
    assert "Cencomedia" not in nombres
    barras = [b["nombre"] for b in production_result["data"]["apertura_circuito"]["barras"]]
    assert "Cencomedia" not in barras


def test_cencomedia_absent_from_tv1_html():
    html = td.DEFAULT_OUTPUT_HTML.read_text(encoding="utf-8") if td.DEFAULT_OUTPUT_HTML.exists() else ""
    if html:
        assert "CENCOMEDIA" not in html.upper()


def test_cencomedia_excluded_from_static_occupancy_denominator(production_result):
    # Cencomedia tiene CompletitudMaestro=NO_APLICA: no puede aportar un
    # denominador fabricado a la ocupacion estatica "elegible" de KPI4.
    assert "CENCOMEDIA" not in production_result["data"]["kpis"]["estatico"]["universo_elegible"]


def test_cencomedia_rows_in_fixture_do_not_change_any_kpi():
    """Cambiar/agregar filas Cencomedia en un fixture no debe modificar
    ningun KPI de TV1: se excluyen (via TV1_EXCLUDED_CIRCUITOS) antes de
    construir el universo, igual que APSA/London."""
    base_rows = [_static_cencosud("C1")]
    base_campanas = [_campana_row(
        "C1CAMP", "C1", IDCampaña="CX", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10"),
    )]
    universe_a = td.build_tv1_universe(_semantic(base_rows, base_campanas))
    kpi1_a = td.compute_kpi1_core(universe_a)

    cencomedia_row = _maestro_row(
        "CM1", CircuitoDashboard="Alguna Cadena Jumbo", Subcircuito="X", Ubicacion="X",
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Flexible gráfico",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    cencomedia_campana = _campana_row(
        "CM1CAMP", "CM1", IDCampaña="CY", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31"),
    )
    universe_b = td.build_tv1_universe(_semantic(base_rows + [cencomedia_row], base_campanas + [cencomedia_campana]))
    kpi1_b = td.compute_kpi1_core(universe_b)

    assert kpi1_a == kpi1_b
    assert "CM1" not in universe_b["element_ids"]


def test_cencosud_still_participates_after_cencomedia_exclusion():
    universe = td.build_tv1_universe(_semantic([_static_cencosud("C1")]))
    assert "CENCOSUD" in universe["circuitos"]
    assert "C1" in universe["element_ids"]


# ---------------------------------------------------------------------------
# 12-13. Pantallas = Digital, Cencomedia = Estático en composición
# ---------------------------------------------------------------------------


def test_pantallas_led_maestro_is_all_digital():
    from export_data import load_pipeline
    _tr, semantic_result, _engine = load_pipeline(PRODUCTION_FILE)
    maestro = semantic_result["maestro"]
    pantallas = maestro[maestro["CircuitoNegocio"] == "PANTALLAS_LED"]
    assert len(pantallas) > 0
    assert (pantallas["Medio"] == "Digital").all()


def test_cencomedia_maestro_is_all_static():
    from export_data import load_pipeline
    _tr, semantic_result, _engine = load_pipeline(PRODUCTION_FILE)
    maestro = semantic_result["maestro"]
    cencomedia = maestro[maestro["CircuitoNegocio"] == "CENCOMEDIA"]
    assert len(cencomedia) > 0
    assert (cencomedia["Medio"] == "Estático").all()


# ---------------------------------------------------------------------------
# 14. Composición reconcilia con activos
# ---------------------------------------------------------------------------


def test_composition_reconciles_with_catalogo_comercial(production_result):
    data = production_result["data"]
    total_familias = sum(f["total"] for f in data["composition"]["familias"])
    assert total_familias == data["catalogo_comercial"]["value"]


# ---------------------------------------------------------------------------
# 15. Mes anterior es consecutivo
# ---------------------------------------------------------------------------


def test_previous_month_is_consecutive():
    assert td._previous_month(2026, 7) == (2026, 6)
    assert td._previous_month(2026, 1) == (2025, 12)  # rollover de anio


# ---------------------------------------------------------------------------
# 16. Porcentajes comparan en puntos porcentuales (pp), no % relativo
# ---------------------------------------------------------------------------


def test_delta_is_percentage_points_not_relative_percent():
    maestro_rows = [_digital_cencosud(f"D{i}") for i in range(10)]
    # Julio: 5 de 10 activos (50%). Junio: 4 de 10 activos (40%).
    campana_rows = []
    for i in range(5):
        campana_rows.append(_campana_row(f"J{i}", f"D{i}", IDCampaña=f"CJ{i}", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10")))
    for i in range(4):
        campana_rows.append(_campana_row(f"N{i}", f"D{i}", IDCampaña=f"CN{i}", FechaInicio=pd.Timestamp("2026-06-05"), FechaFin=pd.Timestamp("2026-06-10")))
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    kpi5 = td.compute_kpi5_digital_calendario(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert kpi5["pct_actual"] == 50.0
    assert kpi5["pct_anterior"] == 40.0
    assert kpi5["delta_pp"] == 10.0  # 50 - 40, NUNCA (50-40)/40*100=25


# ---------------------------------------------------------------------------
# 17. YTD estático no promedia porcentajes (agrega numerador/denominador)
# ---------------------------------------------------------------------------


def test_static_ytd_aggregates_numerator_denominator_not_average_of_percents(production_result):
    kpi4 = production_result["data"]["kpis"]["estatico"]
    # El YTD real (Ene-Jul) no coincide con el promedio simple de
    # ocupacion_actual/ocupacion_anterior: confirma que se agrega
    # numerador/denominador en vez de promediar porcentajes mensuales.
    naive_average = round((kpi4["ocupacion_actual"] + kpi4["ocupacion_anterior"]) / 2, 1)
    assert kpi4["ytd"] != naive_average


# ---------------------------------------------------------------------------
# 18. Unknown no pasa a 0
# ---------------------------------------------------------------------------


def test_unknown_capacity_does_not_become_zero_fill_rate():
    maestro_rows = [_digital_aa2000_unknown_capacity("U1")]
    campana_rows = [_campana_row("C1", "U1", IDCampaña="CAMP-U", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10"))]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    kpi6 = td.compute_kpi6_digital_fill(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert kpi6["status"] == "NO_APLICA"
    assert kpi6["pct_actual"] is None  # nunca 0.0 cuando la capacidad es desconocida


# ---------------------------------------------------------------------------
# 19. MetricStatus se conserva
# ---------------------------------------------------------------------------


def test_metric_status_preserved_in_production_output(production_result):
    kpi4 = production_result["data"]["kpis"]["estatico"]
    kpi6 = production_result["data"]["kpis"]["digital_fill"]
    # Julio 2026 tiene fechas incompletas / maestro parcial reales: el
    # dashboard debe exponer el status real, nunca forzarlo silenciosamente a OK.
    assert kpi4["status"] in ("OK", "PARTIAL", "NO_APLICA", "REQUIERE_CONFIRMACION")
    assert kpi6["status"] in ("OK", "PARTIAL", "NO_APLICA", "REQUIERE_CONFIRMACION")
    # Recalibrado 2026-08-18 (promocion base OCU26+YPF): kpi4 paso de PARTIAL a OK
    # porque el conjunto de cambios promovidos (retiro de 17 elementos legacy de
    # APIE 30943 + reemplazo integro de campanas YPF + 4 filas CENCOSUD ya
    # autorizadas en FINAL_V2 y nunca antes promovidas) elimino las fechas
    # incompletas que generaban el PARTIAL. kpi6 no cambio.
    assert kpi4["status"] == "OK"
    assert kpi6["status"] == "PARTIAL"


# ---------------------------------------------------------------------------
# 20. Input Excel mantiene SHA
# ---------------------------------------------------------------------------


def test_input_excel_sha_unchanged(production_result):
    sha_after_build = vi.calculate_sha256(PRODUCTION_FILE)
    assert sha_after_build == production_result["sha256"]


# ---------------------------------------------------------------------------
# 21-22. Payload TV1 aislado (sin tv2-tv6, sin APSA/London)
# ---------------------------------------------------------------------------


def test_payload_top_level_keys_are_tv1_only(production_result):
    assert set(production_result["data"].keys()) == {
        "meta", "kpis", "evolution", "catalogo_comercial", "composition",
        "espacios", "soportes_fisicos", "campanas_catalogo", "evolution_espacios", "apertura_circuito", "insights",
    }


def test_payload_contains_no_other_tv_datasets(production_json):
    for token in ("tv2", "tv3", "tv4", "tv5", "tv6"):
        assert token not in production_json.lower()


def test_payload_contains_no_excluded_circuits(production_json):
    upper = production_json.upper()
    assert "APSA" not in upper
    assert "LONDON" not in upper
    assert "CENCOMEDIA" not in upper


# ---------------------------------------------------------------------------
# 23. Campañas acumuladas usan DISTINCT IDCampaña
# ---------------------------------------------------------------------------


def test_campanas_acumuladas_are_distinct_not_row_count(production_result):
    data = production_result["data"]
    campanas = data["campanas_catalogo"]["campanas_unicas"]
    assert str(campanas) in data["insights"]["lectura"]


# ---------------------------------------------------------------------------
# 24. Gráfico de evolución no incluye meses posteriores al report_month
# ---------------------------------------------------------------------------


def test_evolution_chart_excludes_future_months(production_result):
    ev = production_result["data"]["evolution"]
    assert len(ev["meses"]) == td.REPORT_MONTH
    assert ev["meses"][-1] == "Jul"
    assert "Ago" not in ev["meses"]
    assert "Sep" not in ev["meses"]


# ---------------------------------------------------------------------------
# Correccion APIE — grano comercial YPF por estacion (prompt Sec.16, items 1-14)
# ---------------------------------------------------------------------------


def _one_station_three_elements():
    """Una estacion YPF (mismo prefijo + misma localidad de Ubicacion) con
    3 ElementoID de formatos distintos (TT, TT, MB), todos con campana
    activa en julio 2026."""
    ubic = "500 - TESTVILLE - Calle Falsa 123"
    maestro_rows = [
        _ypf_digital("500 - TT - 1", ubic),
        _ypf_digital("500 - TT - 2", ubic),
        _ypf_digital("500 - MB - 1", ubic),
    ]
    campana_rows = [
        _campana_row("YC1", "500 - TT - 1", IDCampaña="CY1", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10")),
        _campana_row("YC2", "500 - TT - 2", IDCampaña="CY2", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10")),
        _campana_row("YC3", "500 - MB - 1", IDCampaña="CY3", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10")),
    ]
    return maestro_rows, campana_rows


# 1. YPF core usa COUNT DISTINCT APIE (estacion), no ElementoID.
def test_ypf_core_uses_distinct_station_not_elementoid():
    maestro_rows, _ = _one_station_three_elements()
    maestro_rows = [_static_cencosud("C1")] + maestro_rows  # 1 elemento no-YPF de control
    semantic_result = _semantic(maestro_rows)
    universe = td.build_tv1_universe(semantic_result)
    kpi1 = td.compute_kpi1_core(universe)
    assert kpi1["non_ypf"] == 1
    assert kpi1["ypf_estaciones"] == 1  # 3 ElementoID YPF colapsan a 1 estacion
    assert kpi1["value"] == 2


# 2. Multiples ElementoID de una misma APIE = 1 unidad (catalogo).
def test_multiple_elementid_same_station_counts_as_one_unit():
    maestro_rows, _ = _one_station_three_elements()
    semantic_result = _semantic(maestro_rows)
    universe = td.build_tv1_universe(semantic_result)
    assert universe["ypf_station_catalog_count"] == 1
    assert len(universe["ypf_element_ids"]) == 3


# 3-4. YPF actividad mensual usa APIE distinct; una estacion con varios
# formatos activos sigue contando 1 estacion activa.
def test_ypf_monthly_activity_uses_distinct_station():
    maestro_rows, campana_rows = _one_station_three_elements()
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    active = td._ypf_active_stations(engine, universe["ypf_element_ids"], universe["ypf_station_map"], "2026-07-01", "2026-07-31")
    assert len(active) == 1  # 3 ElementoID activos, 1 sola estacion


# 5. Evolucion YPF usa APIE (no ElementoID/formatos).
def test_evolution_ypf_series_uses_station_grain():
    maestro_rows, campana_rows = _one_station_three_elements()
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    ev = td.compute_evolution(engine, universe, 2026, 7)
    assert ev["ypf"][-1] == 1  # julio: 1 estacion activa, no 3 ElementoID


# 6-7. Composicion YPF usa APIE (catalogo) y NO se divide Digital/Estatico.
def test_composition_ypf_uses_station_grain_without_digital_estatico_split():
    maestro_rows, _campana_rows = _one_station_three_elements()
    semantic_result = _semantic(maestro_rows)
    universe = td.build_tv1_universe(semantic_result)
    catalogo = td.compute_catalogo_total(universe)
    comp = td.compute_composition(universe, catalogo["value"])
    ypf_family = next(f for f in comp["familias"] if f["nombre"] == "YPF")
    assert ypf_family["split"] is False
    assert ypf_family["digital"] == 0
    assert ypf_family["estatico"] == 0
    assert ypf_family["total"] == 1  # 1 estacion (catalogo), no 3 ElementoID ni 2 Dig/1 MB


# 8. Total Core = No YPF ElementoID distinct + YPF APIE distinct.
def test_core_total_equals_non_ypf_plus_ypf_stations():
    maestro_rows, _ = _one_station_three_elements()
    maestro_rows = [_static_cencosud("C1"), _static_cencosud("C2")] + maestro_rows
    semantic_result = _semantic(maestro_rows)
    universe = td.build_tv1_universe(semantic_result)
    kpi1 = td.compute_kpi1_core(universe)
    assert kpi1["value"] == kpi1["non_ypf"] + kpi1["ypf_estaciones"]


# 9. Total actividad = No YPF ElementoID activo + YPF APIE activo.
def test_unidades_actividad_total_equals_non_ypf_active_plus_ypf_active_stations():
    maestro_rows, campana_rows = _one_station_three_elements()
    static_rows = [_static_cencosud("C1")]
    static_campanas = [_campana_row("SC1", "C1", IDCampaña="CS1", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10"))]
    semantic_result = _semantic(static_rows + maestro_rows, static_campanas + campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    kpi3 = td.compute_kpi3_actividad(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"), core_value=2)
    assert kpi3["value"] == kpi3["non_ypf_activo"] + kpi3["ypf_estaciones_activas"]
    assert kpi3["non_ypf_activo"] == 1
    assert kpi3["ypf_estaciones_activas"] == 1
    assert kpi3["value"] == 2


# 10. % del core usa el mismo grano mixto en numerador y denominador.
def test_pct_core_uses_same_mixed_grain_numerator_and_denominator():
    maestro_rows, campana_rows = _one_station_three_elements()
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    kpi1 = td.compute_kpi1_core(universe)
    kpi3 = td.compute_kpi3_actividad(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"), kpi1["value"])
    assert kpi3["pct_core"] == round(kpi3["value"] / kpi1["value"] * 100.0, 1)


# 11. Campañas siguen usando IDCampaña distinct.
# Recalibrado 2026-08-18: 425 -> 443. TV1 incluye YPF en este KPI (solo excluye
# APSA/LONDON_SUPPLY); el reemplazo integro del bloque legacy YPF (10 campanas)
# por YPF Etapa 2 (27 campanas mas granulares) mas las 4 filas CENCOSUD ya
# autorizadas en FINAL_V2 explican el delta. Verificado con calculo
# independiente en pandas directo sobre el Excel (no via el pipeline): 443.
def test_campanas_unchanged_by_apie_correction(production_result):
    assert production_result["data"]["kpis"]["campanas_unicas"]["ytd"] == 443


# 12. Digital por calendario sigue excluyendo YPF.
def test_digital_calendario_excludes_ypf_elements():
    maestro_rows, campana_rows = _one_station_three_elements()
    maestro_rows = maestro_rows + [_digital_cencosud("D1")]
    campana_rows = campana_rows + [_campana_row("DC1", "D1", IDCampaña="CD1", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10"))]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    kpi5 = td.compute_kpi5_digital_calendario(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert kpi5["elegibles"] == 1  # solo D1 (Cencosud digital); los 3 ElementoID YPF quedan fuera


# 13. Digital fill sigue excluyendo YPF (ya cubierto tambien en produccion).
def test_digital_fill_excludes_ypf_universe(production_result):
    assert "YPF" not in production_result["universe"]["digital_circuitos"]


# 14. APIE nulo no genera fallback silencioso a ElementoID.
def test_ypf_element_without_derivable_station_blocks_build():
    bad_row = _ypf_digital("999 - TT - 1", ubicacion="SIN_FORMATO_DE_LOCALIDAD")  # sin " - " -> sin token de localidad
    semantic_result = _semantic([bad_row])
    with pytest.raises(td.BuildError, match="sin estacion"):
        td.build_tv1_universe(semantic_result)


# ---------------------------------------------------------------------------
# Ajuste final — tarjeta YPF, "Unidades con campaña" y composición de catálogo
# ---------------------------------------------------------------------------


def test_kpi_ypf_pct_uses_unidades_con_campana_as_denominator(production_result):
    k = production_result["data"]["kpis"]
    esperado = round(k["ypf"]["estaciones_activas"] / k["unidades_actividad"]["value"] * 100.0, 1)
    assert k["ypf"]["pct_sobre_unidades_campana"] == esperado


def test_unidades_con_campana_breakdown_sums_to_total(production_result):
    k = production_result["data"]["kpis"]
    act = k["unidades_actividad"]
    assert act["digital_activo"] + act["estatico_activo"] + k["ypf"]["estaciones_activas"] == act["value"]


def test_catalogo_comercial_includes_portfolio_complementario(production_result):
    data = production_result["data"]
    # A diferencia de Core Comercial (solo PortfolioTier=CORE), el catalogo
    # de composicion incluye tambien COMPLEMENTARIO (MAB; Cencomedia
    # excluida desde el 21/08/2026) para que esa familia tenga denominador
    # real (spec original Sec.20).
    assert data["catalogo_comercial"]["value"] > data["kpis"]["core_comercial"]["value"]


def test_composition_otros_content_is_mab_and_pilar_frontlight(production_result):
    otros = production_result["data"]["composition"]["otros_circuitos"]
    assert set(otros) == {"MAB", "PILAR_FRONTLIGHT"}


def test_composition_shoppings_merges_cencosud_and_remeros():
    cencosud_row = _static_cencosud("C1")
    remeros_row = _maestro_row(
        "R1", CircuitoDashboard="Shoppings Estático", Subcircuito="REMEROS", Ubicacion="REMEROS",
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático", CapacidadSlotsReel=0, SegundosDia=0,
    )
    semantic_result = _semantic([cencosud_row, remeros_row])
    universe = td.build_tv1_universe(semantic_result)
    catalogo = td.compute_catalogo_total(universe)
    comp = td.compute_composition(universe, catalogo["value"])
    shoppings = next(f for f in comp["familias"] if f["nombre"] == "Shoppings")
    assert shoppings["total"] == 2  # Cencosud + Remeros fusionados en una sola familia


# ---------------------------------------------------------------------------
# Espacios publicitarios v2 (prompt "CORRECCION DEFINITIVA DE TV1" 2026-08-21):
# catalogo completo (ya no "nucleo exacto"), Cencomedia excluida, YPF con
# ocupacion propia Digital+Estatico.
# ---------------------------------------------------------------------------


def _pantalla_led(elemento_id: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Pantalla Led", Subcircuito="TESTSITE", Ubicacion="TESTSITE",
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        Descripcion="", CapacidadSlotsReel=20, SegundosDia=72000,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _totem_shopping(elemento_id: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Shoppings Digital", Subcircuito="CENCOSUD", Ubicacion="UNICENTER",
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        Descripcion="Totem digital de shopping", CapacidadSlotsReel=20, SegundosDia=72000,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _puente_led_cencosud(elemento_id: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Shoppings Digital", Subcircuito="CENCOSUD", Ubicacion="UNICENTER",
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        Descripcion="Puente LED shopping", CapacidadSlotsReel=13, SegundosDia=72000,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _triedro_cencosud(elemento_id: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Shoppings Digital", Subcircuito="CENCOSUD", Ubicacion="UNICENTER",
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        Descripcion="Triedro Digital - Test", CapacidadSlotsReel=10, SegundosDia=50400,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _tripstore_aa2000(elemento_id: str, **overrides) -> dict:
    """AA2000 "Totem Simple" (confirmado con el usuario 21/08/2026 = Tripstore):
    capacidad REGISTRADA (CapacidadSlotsReel), nunca la tasa de 10 de Shopping."""
    row = dict(
        CircuitoDashboard="AA2000", Subcircuito="EZE", Ubicacion="EZE",
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        Descripcion="Totem Simple - Ezeiza", CapacidadSlotsReel=20, SegundosDia=72000,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _static_aa2000(elemento_id: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="AA2000", Subcircuito="EZE", Ubicacion="EZE",
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _otro_digital_sin_regla(elemento_id: str, **overrides) -> dict:
    """OTRO fuera de REMEROS (p.ej. Cencosud): sigue sin regla de conversion.
    Nunca usar Subcircuito=REMEROS aqui -- desde la correccion 22/08/2026 ese
    circuito especifico SI tiene una regla confirmada (ver
    _remeros_totem_otro / test_espacios_remeros_otro_totem_capacidad_10)."""
    row = dict(
        CircuitoDashboard="Shoppings Digital", Subcircuito="CENCOSUD", Ubicacion="UNICENTER",
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        Descripcion="Patio de Comidas", CapacidadSlotsReel=10, SegundosDia=50400,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _remeros_totem_otro(elemento_id: str, **overrides) -> dict:
    """Correccion 22/08/2026 (regla de negocio confirmada por el usuario):
    los elementos Digital de Remeros con FormatoNegocio=OTRO ('TV Led', sin
    keyword rule que matchee Totem) son comercialmente totems de 10 slots."""
    row = dict(
        CircuitoDashboard="Shoppings Digital", Subcircuito="REMEROS", Ubicacion="REMEROS",
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        Descripcion="TV Led", CapacidadSlotsReel=10, SegundosDia=50400,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


# --- Tarjeta 1/2: capacidad != ocupacion, tasas nuevas (no Gate3) ---------


def test_espacios_pantalla_led_capacidad_20_no_ocupacion():
    """1 Pantalla LED = 20 espacios de capacidad. Con 1 sola campaña activa
    todo julio (nunca 2 concurrentes), la ocupación real es 1: NUNCA 20 solo
    porque el soporte "tiene campaña" (capacidad != ocupación)."""
    maestro_rows = [_pantalla_led("P1")]
    campana_rows = [_campana_row(
        "C1", "P1", IDCampaña="CAMP-P1",
        FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31"),
    )]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    espacios = td.compute_espacios(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert espacios["digital"]["por_categoria"]["PANTALLA_LED"]["totales"] == 20
    assert espacios["digital"]["por_categoria"]["PANTALLA_LED"]["ocupados"] == 1.0
    assert espacios["totales"] == 20
    assert espacios["ocupados"] == 1
    assert espacios["disponibles"] == 19


def test_espacios_totem_shopping_uses_10_not_legacy_20():
    maestro_rows = [_totem_shopping("T1")]
    semantic_result = _semantic(maestro_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    espacios = td.compute_espacios(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    detalle = espacios["digital"]["por_categoria"]["TOTEM_SHOPPING"]
    assert detalle["capacidad_por_unidad"] == 10
    assert detalle["totales"] == 10


def test_espacios_puente_led_uses_10_not_legacy_13():
    maestro_rows = [_puente_led_cencosud("PU1")]
    semantic_result = _semantic(maestro_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    espacios = td.compute_espacios(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    detalle = espacios["digital"]["por_categoria"]["PUENTE_LED"]
    assert detalle["capacidad_por_unidad"] == 10
    assert detalle["totales"] == 10


def test_espacios_tripstore_aa2000_uses_registered_capacity_not_shopping_rate():
    """Confirmado con el usuario 21/08/2026: Tripstore (AA2000, Totem Simple)
    usa la capacidad REGISTRADA (CapacidadSlotsReel=20 en este fixture),
    nunca la tasa de 10 de Tótem de Shopping."""
    maestro_rows = [_tripstore_aa2000("TS1", CapacidadSlotsReel=20)]
    semantic_result = _semantic(maestro_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    espacios = td.compute_espacios(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert espacios["aa2000_tripstore"]["elementos"] == 1
    assert espacios["aa2000_tripstore"]["totales"] == 20  # NUNCA 10


def test_espacios_triedro_uses_registered_capacity():
    maestro_rows = [_triedro_cencosud("TR1", CapacidadSlotsReel=10)]
    semantic_result = _semantic(maestro_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    espacios = td.compute_espacios(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    detalle = espacios["digital"]["por_categoria"]["TRIEDRO"]
    assert detalle["totales"] == 10
    assert detalle["elementos"] == 1


def test_espacios_estatico_es_1_a_1():
    maestro_rows = [_static_cencosud("E1"), _static_cencosud("E2")]
    campana_rows = [_campana_row(
        "C1", "E1", IDCampaña="CAMP-E1",
        FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10"),
    )]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    espacios = td.compute_espacios(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert espacios["estatico"]["totales"] == 2
    assert espacios["estatico"]["ocupados"] == 1


def test_espacios_sin_regla_confirmada_no_se_incluye_en_totales():
    """Un formato digital sin regla ('Patio de Comidas', FormatoNegocio=OTRO,
    circuito Cencosud) no debe sumar espacios ni aparecer en ninguna
    categoria con regla. No usa Remeros: ese circuito tiene su propia
    excepcion confirmada desde el 22/08/2026 (ver test siguiente)."""
    maestro_rows = [_otro_digital_sin_regla("O1")]
    semantic_result = _semantic(maestro_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    espacios = td.compute_espacios(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert espacios["totales"] == 0
    assert espacios["sin_regla_confirmada"]["elementos"] == 1


def test_espacios_remeros_otro_totem_capacidad_10():
    """Correccion 22/08/2026: un elemento Digital OTRO ('TV Led') en
    CircuitoNegocio=REMEROS SI tiene regla confirmada (Totem, 10 slots),
    a diferencia de un OTRO identico en cualquier otro circuito."""
    maestro_rows = [_remeros_totem_otro("REM-DB-1")]
    semantic_result = _semantic(maestro_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    espacios = td.compute_espacios(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert espacios["totales"] == 10
    assert espacios["digital"]["por_categoria"]["TOTEM_SHOPPING"]["totales"] == 10
    assert espacios["sin_regla_confirmada"]["elementos"] == 0


# --- Tarjeta 1: Soportes físicos (corrección 21/08/2026) ------------------


def test_espacios_totales_valor_de_referencia_actual(production_result):
    """Valor de referencia de control (correccion 22/08/2026: 3.898 -> 3.958,
    +60 espacios por los 6 totems Remeros Shoppings Digital confirmados como
    capacidad comercial), calculado dinámicamente (compute_espacios no
    depende de compute_soportes_fisicos: son capas independientes)."""
    assert production_result["data"]["espacios"]["totales"] == 3958


def test_soportes_fisicos_no_altera_espacios_totales():
    """1 Pantalla LED = 20 espacios comerciales pero 1 solo soporte físico
    instalado: contar soportes físicos nunca cambia Espacios Totales."""
    maestro_rows = [_pantalla_led("P1")]
    semantic_result = _semantic(maestro_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    espacios = td.compute_espacios(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    soportes = td.compute_soportes_fisicos(universe)
    assert espacios["totales"] == 20  # espacios comerciales
    assert soportes["total"] == 1  # soportes físicos
    assert soportes["digital_no_ypf"] == 1


def test_soportes_fisicos_ypf_incluye_los_cuatro_formatos():
    """YPF agrupa Punteras/Menú Board/Torres/Fotobox como soportes físicos,
    vía FormatoNegocio (tokens MB/TT/PPUNTER/FB de ElementoID)."""
    ubic = "500 - TESTVILLE - Calle Falsa 123"
    maestro_rows = [
        _ypf_digital("500 - MB - 1", ubic),
        _ypf_digital("500 - TT - 1", ubic),
        _ypf_digital("500 - TT - 2", ubic),
        _ypf_digital("500 - PPUNTER - 1", ubic),
        _ypf_digital("500 - PPUNTER - 2", ubic),
        _ypf_static("500 - FB - 1", ubic),
    ]
    semantic_result = _semantic(maestro_rows)
    universe = td.build_tv1_universe(semantic_result)
    soportes = td.compute_soportes_fisicos(universe)
    assert soportes["ypf_detalle"] == {"menu_board": 1, "torres": 2, "punteras": 2, "fotobox": 1}
    assert soportes["ypf"] == 6  # 5 digitales + 1 estático (fotobox), contados individualmente


def test_soportes_ypf_no_se_duplican_en_estatico_o_digital_no_ypf():
    """Los soportes YPF (digitales o estáticos) nunca se suman dentro de
    soportes_estaticos_no_ypf ni soportes_digitales_no_ypf."""
    ubic = "500 - TESTVILLE - Calle Falsa 123"
    maestro_rows = [
        _ypf_digital("500 - MB - 1", ubic),
        _ypf_static("500 - FB - 1", ubic),
        _static_cencosud("C1"),
    ]
    semantic_result = _semantic(maestro_rows)
    universe = td.build_tv1_universe(semantic_result)
    soportes = td.compute_soportes_fisicos(universe)
    assert soportes["estatico_no_ypf"] == 1  # solo C1
    assert soportes["digital_no_ypf"] == 0  # el MB de YPF no cuenta aquí
    assert soportes["ypf"] == 2


def test_soportes_fisicos_porcentajes_suman_100(production_result):
    sop = production_result["data"]["soportes_fisicos"]
    suma = round(sop["pct_estatico"] + sop["pct_digital"] + sop["pct_ypf"], 1)
    assert abs(suma - 100.0) <= 0.1


def test_una_estacion_con_varios_soportes_digitales_sigue_teniendo_5_espacios():
    """El conteo de soportes físicos (Tarjeta 1) es independiente de la
    capacidad comercial YPF: una estación con 2 Torres + 1 Menú Board (3
    soportes físicos) sigue aportando exactamente 5 espacios comerciales
    digitales (ESPACIOS_YPF_POR_ESTACION), nunca 3x5=15 ni 3."""
    maestro_rows, _ = _one_station_three_elements()  # 1 estacion, 2 TT + 1 MB
    semantic_result = _semantic(maestro_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    espacios = td.compute_espacios(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    soportes = td.compute_soportes_fisicos(universe)
    assert soportes["ypf"] == 3  # 3 soportes físicos (2 TT + 1 MB)
    assert espacios["ypf"]["digital"]["totales"] == 5  # 1 estación x 5, nunca 3x5


# --- Tarjeta 3: Campañas Únicas incluye YPF (corrección 21/08/2026) -------


def test_campanas_catalogo_incluye_ypf():
    """Campañas Únicas (Tarjeta 3) ya no excluye YPF: el universo pasa a ser
    universe["element_ids"] completo (ya excluye Cencomedia/APSA/London)."""
    maestro_rows, campana_rows = _one_station_three_elements()  # 1 estacion YPF, CY1/CY2/CY3
    maestro_rows = maestro_rows + [_static_cencosud("C1")]
    campana_rows = campana_rows + [_campana_row(
        "SC1", "C1", IDCampaña="CS1", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10"),
    )]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    campanas = td.compute_campanas_catalogo(engine, universe, "2026-01-01", ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert campanas["campanas_unicas"] == 4  # CS1 (Cencosud) + CY1/CY2/CY3 (YPF)
    assert "label" not in campanas  # rotulo "Catálogo sin YPF" eliminado


def test_campana_presente_en_ypf_y_otro_circuito_cuenta_una_sola_vez():
    """Una misma IDCampaña activa en YPF y en otro circuito no debe
    duplicarse en el total global: nunca se suman subtotales por circuito."""
    ubic = "500 - TESTVILLE - Calle Falsa 123"
    maestro_rows = [
        _maestro_row(
            "500 - MB - 1", CircuitoDashboard="YPF Digital", Subcircuito="X", Ubicacion=ubic,
            Medio="Digital", TipoCatalogo="Abierto", TipoInventario="Digital", CapacidadSlotsReel=0, SegundosDia=0,
        ),
        _static_cencosud("C1"),
    ]
    campana_rows = [
        _campana_row("Y1", "500 - MB - 1", IDCampaña="CAMP-COMPARTIDA", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10")),
        _campana_row("C1carga", "C1", IDCampaña="CAMP-COMPARTIDA", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    campanas = td.compute_campanas_catalogo(engine, universe, "2026-01-01", ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert campanas["campanas_unicas"] == 1  # misma IDCampaña en YPF y Cencosud: 1 sola vez
    assert campanas["presencias_en_elementos"] == 2  # 2 combinaciones campaña-elemento distintas


def test_campanas_unicas_es_distinct_idcampana_no_activaciones():
    maestro_rows = [_static_cencosud(f"E{i}") for i in range(50)]
    campana_rows = [_campana_row(
        f"CARGA{i}", f"E{i}", IDCampaña="CAMP-X",
        FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10"),
    ) for i in range(50)]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    campanas = td.compute_campanas_catalogo(engine, universe, "2026-01-01", ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert campanas["campanas_unicas"] == 1  # 1 campaña en 50 elementos = 1, no 50


def test_presencias_en_elementos_es_distinct_campana_elemento():
    maestro_rows = [_static_cencosud(f"E{i}") for i in range(3)]
    campana_rows = [_campana_row(
        f"CARGA{i}", f"E{i}", IDCampaña="CAMP-X",
        FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10"),
    ) for i in range(3)]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    campanas = td.compute_campanas_catalogo(engine, universe, "2026-01-01", ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert campanas["campanas_unicas"] == 1
    assert campanas["presencias_en_elementos"] == 3  # 1 campaña x 3 elementos = 3 presencias


# --- Tarjeta 6: YPF Digital + Estatico, con ocupacion propia --------------


def test_ypf_catalogo_suma_digital_y_estatico():
    """YPF Digital catalogo = estaciones digitales x5; YPF Estatico catalogo
    = elementos estaticos x1 (prompt Sec.11.A)."""
    ubic = "500 - TESTVILLE - Calle Falsa 123"
    maestro_rows = [
        _maestro_row(
            "500 - TT - 1", CircuitoDashboard="YPF Digital", Subcircuito="X", Ubicacion=ubic,
            Medio="Digital", TipoCatalogo="Abierto", TipoInventario="Digital", CapacidadSlotsReel=0, SegundosDia=0,
        ),
        _maestro_row(
            "500 - FB - 1", CircuitoDashboard="YPF Estático", Subcircuito="X", Ubicacion=ubic,
            Medio="Estático", TipoCatalogo="Abierto", TipoInventario="Físico estático", CapacidadSlotsReel=0, SegundosDia=0,
        ),
    ]
    semantic_result = _semantic(maestro_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    espacios = td.compute_espacios(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert espacios["ypf"]["digital"]["estaciones_catalogo"] == 1
    assert espacios["ypf"]["digital"]["totales"] == 5  # 1 estacion x 5
    assert espacios["ypf"]["estatico"]["totales"] == 1  # 1 elemento x 1
    assert espacios["ypf"]["totales"] == 6


def test_ypf_ocupado_es_concurrencia_real_no_estacion_activa_x5():
    """Aclaración prioritaria 21/08/2026: la ocupación digital YPF NUNCA es
    'estación activa x5' (eso sobreestima). Es mín(5, campañas REALMENTE
    concurrentes). En este fixture 3 campañas distintas (CY1/CY2/CY3) están
    todas activas en la misma estación Y se solapan exactamente en fechas
    (05-10/jul): concurrentes=3, ocupados=min(5,3)=3, NUNCA 5."""
    maestro_rows, campana_rows = _one_station_three_elements()  # 1 estacion, 3 ElementoID (2 TT + 1 MB), todos Digital
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    espacios = td.compute_espacios(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert espacios["ypf"]["digital"]["estaciones_activas"] == 1
    assert espacios["ypf"]["digital"]["ocupados"] == 3  # 3 campañas concurrentes, NUNCA 5 (capacidad) ni 0
    assert espacios["ypf"]["ocupados"] == 3


def test_ypf_ocupado_distintas_campanas_no_concurrentes_no_se_suman():
    """Dos campañas en la misma estación que NUNCA coexisten (una termina
    antes de que empiece la otra) deben dar concurrencia=1, no 2: la
    ocupación mide simultaneidad real, no cuántas campañas 'pasaron' por la
    estación en el mes."""
    ubic = "700 - TESTVILLE - Calle Falsa 789"
    maestro_rows = [_maestro_row(
        "700 - TT - 1", CircuitoDashboard="YPF Digital", Subcircuito="X", Ubicacion=ubic,
        Medio="Digital", TipoCatalogo="Abierto", TipoInventario="Digital", CapacidadSlotsReel=0, SegundosDia=0,
    )]
    campana_rows = [
        _campana_row("C1", "700 - TT - 1", IDCampaña="CA", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-05")),
        _campana_row("C2", "700 - TT - 1", IDCampaña="CB", FechaInicio=pd.Timestamp("2026-07-10"), FechaFin=pd.Timestamp("2026-07-15")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    espacios = td.compute_espacios(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert espacios["ypf"]["digital"]["ocupados"] == 1  # nunca activas al mismo tiempo -> concurrencia maxima 1
    assert espacios["ypf"]["digital"]["estaciones_activas"] == 1  # si tuvo actividad en el periodo (no concurrencia)


def test_ypf_estatico_ocupado_es_elementos_con_campana_x1():
    ubic = "600 - TESTVILLE - Calle Falsa 456"
    maestro_rows = [_maestro_row(
        "600 - FB - 1", CircuitoDashboard="YPF Estático", Subcircuito="X", Ubicacion=ubic,
        Medio="Estático", TipoCatalogo="Abierto", TipoInventario="Físico estático", CapacidadSlotsReel=0, SegundosDia=0,
    )]
    campana_rows = [_campana_row(
        "YC1", "600 - FB - 1", IDCampaña="CY1",
        FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10"),
    )]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    espacios = td.compute_espacios(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert espacios["ypf"]["estatico"]["ocupados"] == 1
    assert espacios["ypf"]["ocupados"] == 1  # sin estaciones digitales activas en este fixture


def test_ypf_porcentaje_ocupacion_es_ocupado_sobre_catalogo(production_result):
    ypf = production_result["data"]["espacios"]["ypf"]
    esperado = round(ypf["ocupados"] / ypf["totales"] * 100.0, 1)
    assert ypf["pct_ocupacion"] == esperado


# --- Tarjetas 4/5: Estatico y Digital SIN YPF -----------------------------


def test_espacios_estatico_excluye_ypf():
    maestro_rows, campana_rows = _one_station_three_elements()  # YPF, sin estatico
    maestro_rows = maestro_rows + [_static_cencosud("C1")]
    semantic_result = _semantic(maestro_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    espacios = td.compute_espacios(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert espacios["estatico"]["totales"] == 1  # solo C1; los YPF (todos Digital) no suman aqui


def test_espacios_digital_excluye_ypf_pero_incluye_tripstore(production_result):
    """Auditoría del denominador 21/08/2026: la tarjeta 5 (Digital) excluye
    YPF, pero SÍ incluye Tripstore AA2000 (200 espacios) -- de lo contrario
    quedaba fuera del total general y de la apertura por circuito a la vez
    que fuera del denominador Digital, rompiendo la reconciliación. Su total
    debe ser estrictamente menor al Espacios Totales global (YPF/Estático
    siguen afuera) y debe reconciliar exactamente con sus 5 categorías,
    Tripstore incluido."""
    d = production_result["data"]
    assert d["espacios"]["digital"]["totales"] < d["espacios"]["totales"]
    cat = d["espacios"]["digital"]["por_categoria"]
    assert d["espacios"]["digital"]["totales"] == (
        cat["PANTALLA_LED"]["totales"] + cat["TOTEM_SHOPPING"]["totales"]
        + cat["PUENTE_LED"]["totales"] + cat["TRIEDRO"]["totales"] + cat["TRIPSTORE_AA2000"]["totales"]
    )
    assert cat["TRIPSTORE_AA2000"]["totales"] == d["espacios"]["aa2000_tripstore"]["totales"] == 200


def test_espacios_digital_totales_no_deja_afuera_los_200_de_tripstore():
    """No permitir que los 200 espacios de AA2000/Tripstore se incluyan en
    el total general y en la apertura por circuito, pero queden fuera del
    denominador de la tarjeta Digital (hallazgo de la auditoría 21/08/2026)."""
    maestro_rows = [_tripstore_aa2000("TS1", CapacidadSlotsReel=20)]
    semantic_result = _semantic(maestro_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    espacios = td.compute_espacios(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert espacios["digital"]["totales"] == 20  # Tripstore SI cuenta en el denominador Digital
    assert espacios["totales"] == 20


# --- Evolucion mensual de espacios ocupados (no elementos) ----------------


def test_evolution_espacios_eje_es_espacios_no_elementos():
    """Con 1 Pantalla LED (20 espacios) activa todo el mes, la serie digital
    debe reportar ocupacion en unidades de espacios concurrentes (~1), igual
    que la tarjeta, no un conteo de elementos con actividad."""
    maestro_rows = [_pantalla_led("P1")]
    campana_rows = [_campana_row(
        "C1", "P1", IDCampaña="CAMP-P1",
        FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31"),
    )]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    ev = td.compute_evolution_espacios(engine, universe, 2026, 7)
    assert ev["digital"][-1] == 1
    assert ev["meses"][-1] == "Jul"


def test_evolution_espacios_tres_series_sin_doble_conteo(production_result):
    ev = production_result["data"]["evolution_espacios"]
    assert set(ev.keys()) == {"meses", "estatico", "digital", "ypf"}
    assert len(ev["meses"]) == len(ev["estatico"]) == len(ev["digital"]) == len(ev["ypf"])


def test_evolution_espacios_ypf_usa_concurrencia_real_no_estacion_x5():
    """La serie mensual YPF debe usar la misma fórmula corregida que la
    tarjeta (mín(5, concurrentes)), nunca estaciones activas × 5."""
    ubic = "500 - TESTVILLE - Calle Falsa 123"
    maestro_rows = [
        _maestro_row(
            "500 - TT - 1", CircuitoDashboard="YPF Digital", Subcircuito="X", Ubicacion=ubic,
            Medio="Digital", TipoCatalogo="Abierto", TipoInventario="Digital", CapacidadSlotsReel=0, SegundosDia=0,
        ),
    ]
    campana_rows = [_campana_row(
        "YC1", "500 - TT - 1", IDCampaña="CY1",
        FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10"),
    )]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    ev = td.compute_evolution_espacios(engine, universe, 2026, 7)
    assert ev["ypf"][-1] == 1  # 1 campaña concurrente en julio, NUNCA 5


# --- Apertura por circuito -------------------------------------------------


def test_apertura_circuito_barras_reconcilian_con_totales(production_result):
    d = production_result["data"]
    total_barras = sum(b["total"] for b in d["apertura_circuito"]["barras"])
    assert total_barras == d["espacios"]["totales"]


def test_apertura_circuito_porcentajes_suman_100_con_tolerancia(production_result):
    barras = production_result["data"]["apertura_circuito"]["barras"]
    suma = round(sum(b["pct_total"] for b in barras), 1)
    assert abs(suma - 100.0) <= 0.5


def test_apertura_circuito_tripstore_dentro_de_aa2000_sin_duplicar(production_result):
    barras = {b["nombre"]: b for b in production_result["data"]["apertura_circuito"]["barras"]}
    tripstore = production_result["data"]["espacios"]["aa2000_tripstore"]
    assert barras["AA2000"]["digital"] == tripstore["totales"]
    # Tripstore no debe aparecer duplicado dentro de Shoppings.
    shoppings_totem = production_result["data"]["espacios"]["digital"]["por_categoria"]["TOTEM_SHOPPING"]["totales"]
    assert shoppings_totem <= barras["Shoppings"]["digital"]


def test_apertura_circuito_ypf_incluye_digital_y_estatico(production_result):
    barras = {b["nombre"]: b for b in production_result["data"]["apertura_circuito"]["barras"]}
    ypf = production_result["data"]["espacios"]["ypf"]
    assert barras["YPF"]["digital"] == ypf["digital"]["totales"]
    assert barras["YPF"]["estatico"] == ypf["estatico"]["totales"]
    assert barras["YPF"]["total"] == ypf["totales"]


def test_apertura_circuito_pilar_y_mab_en_otros_no_duplicados():
    pilar_row = _maestro_row(
        "PIL1", CircuitoDashboard="Pilar", Subcircuito="X", Ubicacion="PILAR",
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático", CapacidadSlotsReel=0, SegundosDia=0,
    )
    mab_row = _maestro_row(
        "MAB1", CircuitoDashboard="MAB", Subcircuito="X", Ubicacion="X",
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático", CapacidadSlotsReel=0, SegundosDia=0,
    )
    semantic_result = _semantic([pilar_row, mab_row])
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv1_universe(semantic_result)
    espacios = td.compute_espacios(engine, universe, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    apertura = td.compute_apertura_circuito(universe, espacios)
    otros = next(b for b in apertura["barras"] if b["nombre"] == "Otros")
    assert otros["total"] == 2  # Pilar + MAB, cada uno 1 vez


# ---------------------------------------------------------------------------
# Identidades obligatorias (correccion controlada 21/08/2026, Objetivo 4)
# ---------------------------------------------------------------------------


def test_identidad_reconciliacion_catalogo(production_result):
    """espacios_totales = estatico_catalogo + digital_catalogo + ypf_catalogo."""
    esp = production_result["data"]["espacios"]
    assert esp["totales"] == esp["estatico"]["totales"] + esp["digital"]["totales"] + esp["ypf"]["totales"]


def test_identidad_reconciliacion_ocupacion(production_result):
    """espacios_ocupados = estatico_ocupados + digital_ocupados + ypf_ocupados."""
    esp = production_result["data"]["espacios"]
    assert esp["ocupados"] == esp["estatico"]["ocupados"] + esp["digital"]["ocupados"] + esp["ypf"]["ocupados"]


def test_identidad_participacion_suma_100_con_tolerancia(production_result):
    """participacion_estatico + participacion_digital + participacion_ypf = 100%,
    tolerancia maxima 0,1 pp de redondeo."""
    esp = production_result["data"]["espacios"]
    suma = (
        esp["estatico"]["participacion_ocupado_pct"]
        + esp["digital"]["participacion_ocupado_pct"]
        + esp["ypf"]["participacion_ocupado_pct"]
    )
    assert abs(round(suma, 1) - 100.0) <= 0.1


def test_identidad_ocupacion_individual_por_familia(production_result):
    """ocupacion_familia = familia_ocupados / familia_catalogo, para las tres
    familias, recalculado independientemente del campo pct_ocupacion ya
    almacenado (no se compara un campo contra si mismo)."""
    esp = production_result["data"]["espacios"]
    for familia in ("estatico", "digital"):
        f = esp[familia]
        esperado = round(f["ocupados"] / f["totales"] * 100.0, 1)
        assert f["pct_ocupacion"] == esperado
    ypf = esp["ypf"]
    esperado_ypf = round(ypf["ocupados"] / ypf["totales"] * 100.0, 1)
    assert ypf["pct_ocupacion"] == esperado_ypf


def test_identidad_participacion_es_numerador_propio_no_resto_hasta_100(production_result):
    """La participacion de cada familia NUNCA debe calcularse como el resto
    hasta 100% de las otras dos: se verifica recalculando cada una de forma
    independiente contra ocupados_totales."""
    esp = production_result["data"]["espacios"]
    total_ocupados = esp["ocupados"]
    for familia in ("estatico", "digital", "ypf"):
        f = esp[familia]
        esperado = round(f["ocupados"] / total_ocupados * 100.0, 1)
        assert f["participacion_ocupado_pct"] == esperado


def test_identidad_apertura_circuito_mismo_universo_que_totales(production_result):
    """La apertura por circuitos debe sumar exactamente el mismo universo de
    espacios que ESPACIOS TOTALES."""
    d = production_result["data"]
    assert sum(b["total"] for b in d["apertura_circuito"]["barras"]) == d["espacios"]["totales"]


def test_identidad_cencomedia_apsa_london_excluidos_de_todo(production_result):
    """Cencomedia, APSA y London deben continuar excluidos de todos los
    numeradores, denominadores, tarjetas, series mensuales, barras y
    porcentajes (verificado sobre el JSON completo del payload)."""
    payload = json.dumps(production_result["data"], ensure_ascii=False).upper()
    for token in ("CENCOMEDIA", "APSA", "LONDON"):
        assert token not in payload


# ---------------------------------------------------------------------------
# Sobrecarga YPF (agregado de diseño futuro, 21/08/2026): capacidad_slots,
# campanas_concurrentes (solapamiento real), slots_ocupados, excedentes,
# indice_presion, estado -- funciones puras _max_overlap_count /
# _estacion_ocupacion_ypf, testeables sin pasar por el Excel.
# ---------------------------------------------------------------------------


def _iv(inicio: str, fin: str) -> tuple[pd.Timestamp, pd.Timestamp]:
    return pd.Timestamp(inicio), pd.Timestamp(fin)


def test_sobrecarga_ypf_0_campanas():
    r = td._estacion_ocupacion_ypf([])
    assert r == {
        "capacidad_slots": 5, "campanas_concurrentes": 0, "slots_ocupados": 0,
        "campanas_excedentes": 0, "porcentaje_ocupacion": 0.0, "indice_presion": 0.0, "estado": "DISPONIBLE",
    }


def test_sobrecarga_ypf_1_campana():
    r = td._estacion_ocupacion_ypf([_iv("2026-07-01", "2026-07-10")])
    assert r["campanas_concurrentes"] == 1
    assert r["slots_ocupados"] == 1
    assert r["campanas_excedentes"] == 0
    assert r["porcentaje_ocupacion"] == 20.0
    assert r["estado"] == "DISPONIBLE"


def test_sobrecarga_ypf_2_campanas_concurrentes():
    """Ejemplo obligatorio del prompt: 2 IDCampaña distintas activas y
    solapadas -> capacidad 5, ocupacion 2, disponibilidad 3, 40%."""
    intervals = [_iv("2026-07-01", "2026-07-15"), _iv("2026-07-05", "2026-07-20")]
    r = td._estacion_ocupacion_ypf(intervals)
    assert r["capacidad_slots"] == 5
    assert r["campanas_concurrentes"] == 2
    assert r["slots_ocupados"] == 2
    assert r["capacidad_slots"] - r["slots_ocupados"] == 3  # disponibilidad
    assert r["porcentaje_ocupacion"] == 40.0
    assert r["estado"] == "DISPONIBLE"


def test_sobrecarga_ypf_5_campanas_completa():
    intervals = [_iv("2026-07-01", "2026-07-10")] * 5
    # 5 intervalos identicos pero con IDCampaña distintas simuladas por 5 tuplas independientes
    r = td._estacion_ocupacion_ypf(intervals)
    assert r["campanas_concurrentes"] == 5
    assert r["slots_ocupados"] == 5
    assert r["campanas_excedentes"] == 0
    assert r["porcentaje_ocupacion"] == 100.0
    assert r["indice_presion"] == 100.0
    assert r["estado"] == "COMPLETA"


def test_sobrecarga_ypf_7_campanas_sobrecapacidad():
    """Ejemplo obligatorio del prompt: capacidad 5, 7 campañas concurrentes
    -> catalogo 5, ocupados 5, ocupacion 100%, concurrentes 7, excedentes 2,
    indice de presion 140%, estado SOBRECAPACIDAD. Las 7 campañas se
    conservan/contabilizan (campanas_concurrentes=7) pero NUNCA se
    convierten en 7 espacios de catalogo ni 7 ocupados."""
    intervals = [_iv("2026-07-01", "2026-07-10")] * 7
    r = td._estacion_ocupacion_ypf(intervals)
    assert r["capacidad_slots"] == 5
    assert r["slots_ocupados"] == 5
    assert r["porcentaje_ocupacion"] == 100.0
    assert r["campanas_concurrentes"] == 7
    assert r["campanas_excedentes"] == 2
    assert r["indice_presion"] == 140.0
    assert r["estado"] == "SOBRECAPACIDAD"


def test_sobrecarga_ypf_campanas_no_concurrentes_no_se_suman():
    """Dos campañas que NUNCA se solapan (fechas distintas sin superposicion)
    deben dar concurrencia maxima 1, no 2."""
    intervals = [_iv("2026-07-01", "2026-07-05"), _iv("2026-07-10", "2026-07-15")]
    r = td._estacion_ocupacion_ypf(intervals)
    assert r["campanas_concurrentes"] == 1
    assert r["estado"] == "DISPONIBLE"


def test_sobrecarga_ypf_campanas_adyacentes_incluidas_cuentan_concurrentes():
    """Fechas inclusive-inclusive: una campaña que termina el dia D y otra
    que empieza el mismo dia D SI se consideran concurrentes ese dia."""
    intervals = [_iv("2026-07-01", "2026-07-10"), _iv("2026-07-10", "2026-07-20")]
    r = td._estacion_ocupacion_ypf(intervals)
    assert r["campanas_concurrentes"] == 2


def test_sobrecarga_ypf_payload_expone_metricas_para_alerta_futura(production_result):
    """El payload debe exponer las metricas de sobrecarga aunque todavia no
    haya una tarjeta visible dedicada (prompt: 'dejar la metrica disponible
    en el payload y cubierta por tests')."""
    ypf_dig = production_result["data"]["espacios"]["ypf"]["digital"]
    for campo in ("estaciones_completa", "estaciones_sobrecapacidad", "campanas_excedentes_total", "estaciones_requiere_confirmacion"):
        assert campo in ypf_dig
        assert isinstance(ypf_dig[campo], int)


# --- Estructura visual (titulo, orden de tarjetas, tarjeta eliminada) ------


def test_titulo_principal_es_catalogo_de_espacios():
    html = td.TEMPLATE_PATH.read_text(encoding="utf-8")
    assert "<title>Brand Plus · TV1 · Catálogo de Espacios Publicitarios</title>" in html
    assert "Cat&aacute;logo de Espacios Publicitarios" in html  # hdr-title (HTML entities)
    assert "Visión general del negocio" not in html
    assert "Visi&oacute;n general del negocio" not in html


def test_orden_de_las_seis_tarjetas_en_template():
    html = td.TEMPLATE_PATH.read_text(encoding="utf-8")
    marcador = "renderTV1"
    idx = html.index(marcador)
    fragmento = html[idx: idx + 4000]
    orden_esperado = [
        "Espacios comerciales", "Espacios ocupados", "Campañas únicas",
        "Espacios Estáticos", "Espacios Digitales", "Espacios YPF",
    ]
    posiciones = [fragmento.index(etiqueta) for etiqueta in orden_esperado]
    assert posiciones == sorted(posiciones)


def test_tarjeta_fuera_del_nucleo_exacto_eliminada():
    html = td.TEMPLATE_PATH.read_text(encoding="utf-8")
    assert "Fuera del núcleo exacto" not in html
    assert "núcleo exacto" not in html.lower()


# ---------------------------------------------------------------------------
# Ajuste visual 22/08/2026: "Espacios comerciales", "de su catálogo",
# leyenda inferior sin recortar (solo presentación, ninguna métrica cambia)
# ---------------------------------------------------------------------------


def test_tarjeta1_titulo_es_espacios_comerciales():
    html = td.TEMPLATE_PATH.read_text(encoding="utf-8")
    idx = html.index("renderTV1")
    fragmento = html[idx: idx + 4000]
    assert "'Espacios comerciales'" in fragmento
    assert "'Espacios totales'" not in fragmento


def test_tarjeta1_leyendas_infraestructura_fisica_y_ypf():
    html = td.TEMPLATE_PATH.read_text(encoding="utf-8")
    idx = html.index("renderTV1")
    fragmento = html[idx: idx + 4000]
    assert "Infraestructura física" in fragmento
    assert "YPF: múltiples soportes comparten 5 slots por estación" in fragmento


def test_tarjeta1_no_muestra_porcentajes_de_soportes():
    """La distribución Estático/Digital/YPF de soportes se retira de la
    tarjeta visible (ajuste 22/08/2026): ya no se arma con pct_estatico/
    pct_digital/pct_ypf en la llamada de render de la Tarjeta 1."""
    html = td.TEMPLATE_PATH.read_text(encoding="utf-8")
    idx = html.index("renderTV1")
    fragmento = html[idx: idx + 4000]
    assert "sop.pct_estatico" not in fragmento
    assert "sop.pct_digital" not in fragmento
    assert "sop.pct_ypf" not in fragmento


def test_soportes_fisicos_porcentajes_se_conservan_en_payload(production_result):
    """Los porcentajes de soportes se retiran de la tarjeta visible pero
    siguen disponibles en el payload (para tests/informe)."""
    sop = production_result["data"]["soportes_fisicos"]
    assert set(("pct_estatico", "pct_digital", "pct_ypf")) <= sop.keys()
    assert sop["pct_estatico"] is not None
    assert sop["pct_digital"] is not None
    assert sop["pct_ypf"] is not None


def test_pct_ocupacion_de_su_catalogo_en_estatico_digital_ypf():
    html = td.TEMPLATE_PATH.read_text(encoding="utf-8")
    idx = html.index("renderTV1")
    fragmento = html[idx: idx + 4000]
    assert fragmento.count("% ocupación de su catálogo") == 3  # Estático, Digital, YPF
    assert "% ocupación del catálogo" not in fragmento


def test_tarjeta_espacios_ocupados_no_usa_de_su_catalogo():
    """La tarjeta general Espacios Ocupados usa el catálogo comercial total,
    no el catálogo particular de una familia: conserva '% ocupación' a
    secas, nunca 'de su catálogo'."""
    html = td.TEMPLATE_PATH.read_text(encoding="utf-8")
    idx = html.index("renderTV1")
    fragmento = html[idx: idx + 4000]
    assert "'% ocupación'," in fragmento


def test_apertura_footnote_soportes_sin_capacidad_confirmada():
    html = td.TEMPLATE_PATH.read_text(encoding="utf-8")
    assert "soportes digitales sin capacidad confirmada se excluyen de las barras." in html
    assert "sin regla de espacios confirmada, no incluidos en ninguna barra" not in html


def test_apertura_nota_categorias_texto_corto(production_result):
    """La nota de agrupacion por categoria (Shoppings/AA2000/YPF/Otros) se
    acorta para que la leyenda inferior del panel no quede recortada en
    1920x1080 (ajuste 22/08/2026)."""
    nota = production_result["data"]["apertura_circuito"]["nota"]
    assert nota == (
        "Shoppings incluye Tótems, Puentes y Triedros; AA2000 incluye Tripstore; YPF incluye "
        "Estático; Otros incluye Pilar y MAB."
    )
    assert len(nota) < 150  # sensiblemente mas corta que la version anterior (~230 caracteres)


def test_insights_no_mencionan_cencomedia_ni_nucleo(production_result):
    ins = production_result["data"]["insights"]
    texto = (ins["lectura"] + ins["punto_positivo"] + ins["a_atender"]).lower()
    assert "cencomedia" not in texto
    assert "núcleo" not in texto and "nucleo" not in texto
    assert "1.052" not in texto and "1052" not in texto


# --- Texto fantasma (backdrop-filter) ---------------------------------------


def test_kpi_cards_tienen_isolation_para_evitar_bleed_de_backdrop_filter():
    """Fix estructural del artefacto de texto fantasma detectado en la
    migracion anterior: tarjetas adyacentes con backdrop-filter:blur(10px) y
    gap:14px pueden "sangrar" contenido vecino en Chromium; isolation:isolate
    por tarjeta evita el artefacto sin cambiar el layout."""
    html = td.TEMPLATE_PATH.read_text(encoding="utf-8")
    kpi_rule_start = html.index(".kpi{")
    kpi_rule = html[kpi_rule_start: kpi_rule_start + 400]
    assert "backdrop-filter:blur(10px)" in kpi_rule
    assert "isolation:isolate" in kpi_rule


# ---------------------------------------------------------------------------
# Corrección horaria ART (prompt Sec.11)
# ---------------------------------------------------------------------------


def test_meta_generado_iso_has_art_offset(production_result):
    assert production_result["data"]["meta"]["generado_iso"].endswith("-03:00")


def test_meta_timezone_field_is_buenos_aires(production_result):
    assert production_result["data"]["meta"]["timezone"] == "America/Argentina/Buenos_Aires"


def test_art_timezone_offset_is_minus_three_hours():
    import datetime as _dt
    now_art = _dt.datetime.now(td.ART_TZ)
    assert now_art.utcoffset() == _dt.timedelta(hours=-3)


def test_meta_generado_matches_generado_iso_local_time(production_result):
    meta = production_result["data"]["meta"]
    import datetime as _dt
    iso_dt = _dt.datetime.fromisoformat(meta["generado_iso"])
    assert iso_dt.strftime("%d/%m/%Y %H:%M") == meta["generado"]


def test_template_clock_uses_explicit_buenos_aires_timezone():
    html = td.TEMPLATE_PATH.read_text(encoding="utf-8")
    assert "America/Argentina/Buenos_Aires" in html
    assert "Intl.DateTimeFormat" in html


def test_template_updated_reads_from_meta_not_client_open_time():
    """'Actualizado' debe leerse de D.meta.generado (momento real de
    generación del builder), nunca de la hora en que se abrió el HTML."""
    html = td.TEMPLATE_PATH.read_text(encoding="utf-8")
    assert "D.meta.generado" in html
    assert "now.getDate()" not in html  # patrón viejo: hora de apertura del HTML
