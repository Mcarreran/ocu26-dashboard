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


def test_cencomedia_present_as_static_family(production_result):
    familias = production_result["data"]["composition"]["familias"]
    cencomedia = next(f for f in familias if f["nombre"] == "Cencomedia")
    assert cencomedia["digital"] == 0  # Cencomedia nunca aporta digital


def test_cencomedia_excluded_from_static_occupancy_denominator(production_result):
    # Cencomedia tiene CompletitudMaestro=NO_APLICA: no puede aportar un
    # denominador fabricado a la ocupacion estatica "elegible" de KPI4.
    assert "CENCOMEDIA" not in production_result["data"]["kpis"]["estatico"]["universo_elegible"]


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
        "meta", "kpis", "evolution", "catalogo_comercial", "composition", "insights",
    }


def test_payload_contains_no_other_tv_datasets(production_json):
    for token in ("tv2", "tv3", "tv4", "tv5", "tv6"):
        assert token not in production_json.lower()


def test_payload_contains_no_excluded_circuits(production_json):
    upper = production_json.upper()
    assert "APSA" not in upper
    assert "LONDON" not in upper


# ---------------------------------------------------------------------------
# 23. Campañas acumuladas usan DISTINCT IDCampaña
# ---------------------------------------------------------------------------


def test_campanas_acumuladas_are_distinct_not_row_count(production_result):
    data = production_result["data"]
    ytd = data["kpis"]["campanas_unicas"]["ytd"]
    assert str(ytd) in data["insights"]["lectura"]


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
    # de composicion incluye tambien COMPLEMENTARIO (Cencomedia/MAB) para
    # que esas familias tengan denominador real (spec original Sec.20).
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
