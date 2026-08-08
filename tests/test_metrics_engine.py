"""Pruebas para scripts/metrics_engine.py (Gate 3B - motor de metricas OCU26).

No modifica scripts/validate_input.py, scripts/transform_data.py, sus tests,
ni input/OCU26_BASE_DATOS.xlsx. Los escenarios de fechas/spots/capacidad usan
fixtures sinteticos construidos con la CONFIGURACION REAL
(config/business_semantics.json) para que las pruebas ejerciten las reglas
de negocio reales (Puente LED=13, Pantalla LED=20, etc.), no una config de
juguete. Los tests de universos/cruces tambien se ejecutan contra el archivo
productivo real via transform_data().
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT / "scripts"))
sys.path.insert(0, str(REPO_ROOT / "tests"))

import semantic_model as sm  # noqa: E402
from transform_data import transform_data  # noqa: E402
from metrics_engine import (  # noqa: E402
    MetricsEngine,
    MetricsEngineError,
    UnknownDimensionError,
    UnknownMetricError,
    UnknownUniverseError,
    UnsupportedBusinessCaseError,
)
from test_semantic_model import _campana_row, _maestro_row, _transform_result  # noqa: E402

PRODUCTION_FILE = REPO_ROOT / "input" / "OCU26_BASE_DATOS.xlsx"


def _engine(maestro_rows: list[dict], campanas_rows: list[dict] | None = None) -> MetricsEngine:
    config = sm.load_config()
    result = sm.build_semantic_model(_transform_result(maestro_rows, campanas_rows or []), config)
    return MetricsEngine(result)


@pytest.fixture(scope="module")
def production_engine() -> MetricsEngine:
    result = sm.build_semantic_model(transform_data(PRODUCTION_FILE))
    return MetricsEngine(result)


# ---------------------------------------------------------------------------
# Ocupacion estatica: fechas inclusivas, solapamiento, Reservada/Cancelado
# ---------------------------------------------------------------------------


def _static_element(elemento_id="E1"):
    return _maestro_row(
        elemento_id, CircuitoDashboard="Shoppings Estático", Subcircuito="CENCOSUD", Ubicacion="UNICENTER",
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático", CapacidadSlotsReel=0, SegundosDia=0,
    )


def test_fecha_inicio_igual_fecha_fin_es_un_dia_ocupado():
    engine = _engine(
        [_static_element()],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-10"), FechaFin=pd.Timestamp("2026-03-10"), Estado="Activa")],
    )
    result = engine.query("elemento_dias_ocupados", start_date="2026-03-01", end_date="2026-03-31")
    assert result.iloc[0]["Value"] == 1


def test_campana_que_cruza_dos_meses_se_reparte_por_periodo_consultado():
    engine = _engine(
        [_static_element()],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-01-25"), FechaFin=pd.Timestamp("2026-02-10"), Estado="Activa")],
    )
    enero = engine.query("elemento_dias_ocupados", start_date="2026-01-01", end_date="2026-01-31")
    febrero = engine.query("elemento_dias_ocupados", start_date="2026-02-01", end_date="2026-02-28")
    assert enero.iloc[0]["Value"] == 7  # 25..31
    assert febrero.iloc[0]["Value"] == 10  # 1..10


def test_campanas_solapadas_no_duplican_dia_ocupado():
    engine = _engine(
        [_static_element()],
        [
            _campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-01"), FechaFin=pd.Timestamp("2026-03-10"), Estado="Activa"),
            _campana_row("C2", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-15"), Estado="Activa"),
        ],
    )
    result = engine.query("elemento_dias_ocupados", start_date="2026-03-01", end_date="2026-03-31")
    assert result.iloc[0]["Value"] == 15  # union 1..15, no 10+11


def test_reservada_bloquea_disponibilidad_futura():
    engine = _engine(
        [_static_element()],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-09-01"), FechaFin=pd.Timestamp("2026-09-10"), Estado="Reservada")],
    )
    result = engine.query("elemento_dias_ocupados", start_date="2026-09-01", end_date="2026-09-30")
    assert result.iloc[0]["Value"] == 10


def test_cancelada_no_bloquea_ocupacion():
    engine = _engine(
        [_static_element()],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-01"), FechaFin=pd.Timestamp("2026-03-10"), Estado="Cancelado")],
    )
    result = engine.query("elemento_dias_ocupados", start_date="2026-03-01", end_date="2026-03-31")
    assert result.iloc[0]["Value"] == 0


def test_fecha_indefinida_se_extiende_hasta_fin_del_periodo_consultado():
    engine = _engine(
        [_static_element()],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-20"), FechaFin=None, FechaIndefinida="Si", Estado="Activa")],
    )
    result = engine.query("elemento_dias_ocupados", start_date="2026-03-01", end_date="2026-03-31")
    assert result.iloc[0]["Value"] == 12  # 20..31 inclusive


def test_ocupacion_calendario_pct_calcula_correctamente_sobre_circuito_cerrado():
    engine = _engine(
        [_static_element()],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-01"), FechaFin=pd.Timestamp("2026-03-10"), Estado="Activa")],
    )
    result = engine.query("ocupacion_calendario_pct", start_date="2026-03-01", end_date="2026-03-31")
    assert result.iloc[0]["Numerator"] == 10
    assert result.iloc[0]["Denominator"] == 31
    assert result.iloc[0]["Value"] == pytest.approx(10 / 31 * 100)
    assert result.iloc[0]["MetricStatus"] == "OK"


def test_ocupacion_calendario_no_aplica_con_cobertura_desconocida():
    ypf_element = _maestro_row(
        "E1", CircuitoDashboard="YPF Estático", Subcircuito="1047", Ubicacion="1047 - RUTA X", Medio="Estático",
        TipoCatalogo="Abierto", TipoInventario="Físico estático",
    )
    engine = _engine([ypf_element])
    result = engine.query("ocupacion_calendario_pct", start_date="2026-01-01", end_date="2026-01-31")
    assert result.iloc[0]["MetricStatus"] == "NO_APLICA"
    assert pd.isna(result.iloc[0]["Value"])
    # actividad_sobre_registrados_pct SI se puede calcular (nombre distinto, sin gate de cobertura)
    result2 = engine.query("actividad_sobre_registrados_pct", start_date="2026-01-01", end_date="2026-01-31")
    assert result2.iloc[0]["MetricStatus"] == "OK"


# ---------------------------------------------------------------------------
# Spots / salidas / segundos vendidos
# ---------------------------------------------------------------------------


def _digital_pantalla_led(elemento_id="E1"):
    return _maestro_row(
        elemento_id, CircuitoDashboard="Pantalla Led", Subcircuito="PLED", Ubicacion="CABILDO",
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital", CapacidadSlotsReel=40, SegundosDia=75000,
    )


@pytest.mark.parametrize("salidas,esperado_segundos", [(1, 1800), (2, 3600), (3, 5400), (4, 7200)])
def test_segundos_vendidos_por_cantidad_de_salidas(salidas, esperado_segundos):
    engine = _engine(
        [_digital_pantalla_led()],
        [_campana_row(
            "C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"),
            Estado="Activa", DuracionSpotSeg=10.0, SalidasVendidas=float(salidas),
        )],
    )
    result = engine.query("segundos_vendidos", start_date="2026-03-05", end_date="2026-03-05")
    assert result.iloc[0]["Value"] == pytest.approx(esperado_segundos)


def test_spot_duracion_distinta_a_10_lanza_excepcion_controlada():
    engine = _engine(
        [_digital_pantalla_led()],
        [_campana_row(
            "C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"),
            Estado="Activa", DuracionSpotSeg=20.0, SalidasVendidas=2.0,
        )],
    )
    with pytest.raises(UnsupportedBusinessCaseError):
        engine.query("segundos_vendidos", start_date="2026-03-05", end_date="2026-03-05")


def test_exclusividad_lanza_excepcion_controlada_no_calcula_silenciosamente():
    engine = _engine(
        [_digital_pantalla_led()],
        [_campana_row(
            "C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"),
            Estado="Activa", TipoExclusividad="Día completo",
        )],
    )
    with pytest.raises(UnsupportedBusinessCaseError):
        engine.query("slots_ocupados", start_date="2026-03-05", end_date="2026-03-05")


def test_slots_ocupados_cuenta_campanas_concurrentes_no_salidas_vendidas():
    engine = _engine(
        [_digital_pantalla_led()],
        [
            _campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"), SalidasVendidas=4.0),
            _campana_row("C2", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"), SalidasVendidas=1.0),
        ],
    )
    result = engine.query("slots_ocupados", start_date="2026-03-05", end_date="2026-03-05")
    assert result.iloc[0]["Value"] == 2  # 2 campanas concurrentes, no 4+1=5 salidas


def test_fill_rate_slots_y_segundos():
    engine = _engine(
        [_digital_pantalla_led()],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"), SalidasVendidas=2.0)],
    )
    slots = engine.query("fill_rate_slots", start_date="2026-03-05", end_date="2026-03-05")
    segundos = engine.query("fill_rate_segundos", start_date="2026-03-05", end_date="2026-03-05")
    assert slots.iloc[0]["Value"] == pytest.approx(1 / 20 * 100)  # 1 slot ocupado / 20 comerciales (Pantalla Led)
    assert segundos.iloc[0]["Value"] == pytest.approx(3600 / 72000 * 100)


# ---------------------------------------------------------------------------
# Capacidad comercial vs legacy (via el motor, no solo semantic_model)
# ---------------------------------------------------------------------------


def test_slots_comerciales_puente_led_13_vs_legacy_10():
    puente = _maestro_row(
        "E1", CircuitoDashboard="Shoppings Digital", Subcircuito="CENCOSUD", Ubicacion="UNICENTER",
        Medio="Digital", Descripcion="Puente Led 2", CapacidadSlotsReel=10, SegundosDia=50400,
    )
    engine = _engine([puente])
    result = engine.query("slots_comerciales", start_date="2026-01-01", end_date="2026-01-31")
    assert result.iloc[0]["Value"] == 13


def test_slots_comerciales_totem_20():
    totem = _maestro_row(
        "E1", CircuitoDashboard="Shoppings Digital", Subcircuito="CENCOSUD", Ubicacion="UNICENTER",
        Medio="Digital", Descripcion="Totem 1", CapacidadSlotsReel=20, SegundosDia=100800,
    )
    engine = _engine([totem])
    result = engine.query("slots_comerciales", start_date="2026-01-01", end_date="2026-01-31")
    assert result.iloc[0]["Value"] == 20


def test_slots_comerciales_pantalla_led_20_aunque_legacy_sea_40():
    engine = _engine([_digital_pantalla_led()])
    result = engine.query("slots_comerciales", start_date="2026-01-01", end_date="2026-01-31")
    assert result.iloc[0]["Value"] == 20


def test_capacidad_fuente_cero_excluye_del_calculo_y_marca_partial():
    zero_cap = _maestro_row(
        "E1", CircuitoDashboard="London Supply", Subcircuito="USH", Ubicacion="USHUAIA", Ciudad="USHUAIA",
        Medio="Digital", TipoCatalogo="Abierto", Descripcion="Pantalla Digital", CapacidadSlotsReel=0, SegundosDia=0,
    )
    engine = _engine([zero_cap])
    result = engine.query("slots_comerciales", start_date="2026-01-01", end_date="2026-01-31")
    assert result.iloc[0]["Value"] == 0
    assert result.iloc[0]["MetricStatus"] == "PARTIAL"
    assert "REQUIERE_CONFIRMACION" in result.iloc[0]["Warnings"]


# ---------------------------------------------------------------------------
# Universos: APSA fuera de general/core, dentro de historico
# ---------------------------------------------------------------------------


def test_apsa_fuera_de_operativo_general_y_performance_core():
    apsa = _maestro_row("E1", CircuitoDashboard="Shoppings Estático", Subcircuito="APSA", Ubicacion="CABA", Medio="Estático")
    engine = _engine([apsa])
    assert engine.query("elementos_registrados", universe="OPERATIVO_GENERAL").iloc[0]["Value"] == 0
    assert engine.query("elementos_registrados", universe="PERFORMANCE_CORE").iloc[0]["Value"] == 0
    assert engine.query("elementos_registrados", universe="COMPLETO_HISTORICO").iloc[0]["Value"] == 1


# ---------------------------------------------------------------------------
# group_by / filters genericos y validacion de dimensiones
# ---------------------------------------------------------------------------


def test_group_by_generico_combina_dimensiones_de_maestro():
    rows = [
        _maestro_row("E1", CircuitoDashboard="Shoppings Estático", Subcircuito="CENCOSUD", Ubicacion="UNICENTER", Ciudad="CABA", Medio="Estático"),
        _maestro_row("E2", CircuitoDashboard="Shoppings Estático", Subcircuito="CENCOSUD", Ubicacion="P.OESTE", Ciudad="GBA", Medio="Estático"),
    ]
    engine = _engine(rows)
    result = engine.query("elementos_registrados", group_by=["CircuitoNegocio", "Ciudad"], universe="COMPLETO_HISTORICO")
    assert len(result) == 2
    assert set(result["Ciudad"]) == {"CABA", "GBA"}
    assert (result["CircuitoNegocio"] == "CENCOSUD").all()


def test_filters_genericos_equality_y_lista():
    rows = [
        _maestro_row("E1", CircuitoDashboard="Pantalla Led", Subcircuito="PLED", Ubicacion="CABILDO", Ciudad="CABA", Medio="Digital"),
        _maestro_row("E2", CircuitoDashboard="Pantalla Led", Subcircuito="PLED", Ubicacion="CORDOBA", Ciudad="CORDOBA", Medio="Digital"),
    ]
    engine = _engine(rows)
    assert engine.query("elementos_registrados", filters={"Ciudad": ["CABA", "CORDOBA"]}, universe="COMPLETO_HISTORICO").iloc[0]["Value"] == 2
    assert engine.query("elementos_registrados", filters={"Ciudad": "CABA"}, universe="COMPLETO_HISTORICO").iloc[0]["Value"] == 1


def test_filters_booleanos_sobre_flags_semanticos():
    rows = [
        _maestro_row("E1", CircuitoDashboard="Pantalla Led", Subcircuito="PLED", Ubicacion="CABILDO", Medio="Digital"),
        _maestro_row("E2", CircuitoDashboard="Shoppings Estático", Subcircuito="APSA", Ubicacion="CABA", Medio="Estático"),
    ]
    engine = _engine(rows)
    result = engine.query("elementos_registrados", filters={"IncluyePerformanceCore": True}, universe="COMPLETO_HISTORICO")
    assert result.iloc[0]["Value"] == 1


def test_dimension_inexistente_en_group_by_lanza_error_explicito():
    engine = _engine([_static_element()])
    with pytest.raises(UnknownDimensionError):
        engine.query("elementos_registrados", group_by=["DimensionQueNoExiste"])


def test_dimension_inexistente_en_filters_lanza_error_explicito():
    engine = _engine([_static_element()])
    with pytest.raises(UnknownDimensionError):
        engine.query("elementos_registrados", filters={"DimensionQueNoExiste": "x"})


def test_metrica_desconocida_lanza_error_explicito():
    engine = _engine([_static_element()])
    with pytest.raises(UnknownMetricError):
        engine.query("metrica_inventada")


def test_universo_desconocido_lanza_error_explicito():
    engine = _engine([_static_element()])
    with pytest.raises(UnknownUniverseError):
        engine.query("elementos_registrados", universe="NO_EXISTE")


def test_metrica_de_periodo_sin_fechas_lanza_error_explicito():
    engine = _engine([_static_element()])
    with pytest.raises(MetricsEngineError):
        engine.query("elemento_dias_ocupados")


# ---------------------------------------------------------------------------
# Contra produccion real: demostracion de cruces sin funciones dedicadas
# ---------------------------------------------------------------------------


def test_demo_operativo_general_excluye_apsa(production_engine):
    df = production_engine.query("elementos_registrados", group_by=["CircuitoNegocio"], universe="OPERATIVO_GENERAL")
    assert "APSA" not in set(df["CircuitoNegocio"])


def test_demo_performance_core_es_exactamente_los_seis_circuitos_core(production_engine):
    df = production_engine.query("elementos_registrados", group_by=["CircuitoNegocio"], universe="PERFORMANCE_CORE")
    assert set(df["CircuitoNegocio"]) == {"CENCOSUD", "REMEROS", "PANTALLAS_LED", "PILAR_FRONTLIGHT", "AA2000", "YPF"}


def test_demo_ypf_por_medio_y_formato_sin_funcion_dedicada(production_engine):
    df = production_engine.query(
        "elementos_registrados", group_by=["Medio", "FormatoNegocio"],
        filters={"CircuitoNegocio": "YPF"}, universe="COMPLETO_HISTORICO",
    )
    assert df["Value"].sum() > 0
    assert set(df["FormatoNegocio"]) == {"YPF_MENU_BOARD", "YPF_TORRE", "YPF_PUNTERA", "YPF_MUPI_FOTOBOX"}


def test_demo_fill_rate_slots_pantallas_led_denominador_consistente_con_elementos_registrados(production_engine):
    n_elementos = production_engine.query(
        "elementos_registrados", filters={"CircuitoNegocio": "PANTALLAS_LED"}, universe="COMPLETO_HISTORICO"
    ).iloc[0]["Value"]
    df = production_engine.query(
        "fill_rate_slots", filters={"CircuitoNegocio": "PANTALLAS_LED"}, universe="COMPLETO_HISTORICO",
        start_date="2026-07-01", end_date="2026-07-31",
    )
    assert df.iloc[0]["Denominator"] == n_elementos * 20


def test_demo_cliente_ciudad_medio_cruce_libre(production_engine):
    df = production_engine.query(
        "campanas", group_by=["Cliente", "Ciudad", "Medio"], universe="COMPLETO_HISTORICO",
        start_date="2026-01-01", end_date="2026-12-31",
    )
    assert not df.empty


# ---------------------------------------------------------------------------
# Gate3B.1 - correccion quirurgica de aplicabilidad de metricas
# ---------------------------------------------------------------------------


def _static_aa2000(elemento_id="E1"):
    return _maestro_row(
        elemento_id, CircuitoDashboard="AA2000", Subcircuito="EZEIZA", Ubicacion="EZEIZA",
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático", CapacidadSlotsReel=0, SegundosDia=0,
    )


def _digital_aa2000(elemento_id="E1"):
    return _maestro_row(
        elemento_id, CircuitoDashboard="AA2000", Subcircuito="TRIPSTORE", Ubicacion="EZEIZA",
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital", Descripcion="Totem Tripstore",
        CapacidadSlotsReel=20, SegundosDia=100800,
    )


def _digital_ypf(elemento_id="1047 - TT - 1"):
    return _maestro_row(
        elemento_id, CircuitoDashboard="YPF Digital", Subcircuito="1047", Ubicacion="1047 - RUTA X",
        Medio="Digital", TipoCatalogo="Abierto", TipoInventario="Digital", Descripcion="Mueble Torre",
        CapacidadSlotsReel=20, SegundosDia=100800,
    )


# -- A/B/C: AA2000 (CoberturaCatalogo=COMPLETO, CompletitudMaestro=PARCIAL) --


def test_A_aa2000_ocupacion_calendario_no_aplica_por_completitud_parcial():
    engine = _engine(
        [_static_aa2000()],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-01"), FechaFin=pd.Timestamp("2026-03-10"), Estado="Activa")],
    )
    result = engine.query("ocupacion_calendario_pct", start_date="2026-03-01", end_date="2026-03-31")
    assert result.iloc[0]["MetricStatus"] == "NO_APLICA"
    assert pd.isna(result.iloc[0]["Value"])


def test_B_aa2000_actividad_sobre_registrados_si_se_puede_calcular():
    engine = _engine(
        [_static_aa2000()],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-01"), FechaFin=pd.Timestamp("2026-03-10"), Estado="Activa")],
    )
    result = engine.query("actividad_sobre_registrados_pct", start_date="2026-03-01", end_date="2026-03-31")
    assert result.iloc[0]["MetricStatus"] == "OK"
    assert result.iloc[0]["Numerator"] == 10
    assert result.iloc[0]["Denominator"] == 31


def test_C_aa2000_digital_fill_rate_slots_partial_por_completitud_parcial():
    engine = _engine(
        [_digital_aa2000()],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"), SalidasVendidas=2.0)],
    )
    result = engine.query("fill_rate_slots", start_date="2026-03-05", end_date="2026-03-05")
    assert result.iloc[0]["MetricStatus"] == "PARTIAL"
    assert not pd.isna(result.iloc[0]["Value"])
    assert "incompleto" in result.iloc[0]["Warnings"]


# -- D/E/F/G: YPF (Brand Plus no administra el CMS, sin ocupacion de slots/segundos) --


def test_D_ypf_digital_fill_rate_slots_no_aplica():
    engine = _engine(
        [_digital_ypf()],
        [_campana_row("C1", "1047 - TT - 1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"), SalidasVendidas=2.0)],
    )
    result = engine.query("fill_rate_slots", start_date="2026-03-05", end_date="2026-03-05")
    assert result.iloc[0]["MetricStatus"] == "NO_APLICA"
    assert pd.isna(result.iloc[0]["Value"])


def test_E_ypf_digital_fill_rate_segundos_no_aplica():
    engine = _engine(
        [_digital_ypf()],
        [_campana_row("C1", "1047 - TT - 1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"), SalidasVendidas=2.0)],
    )
    result = engine.query("fill_rate_segundos", start_date="2026-03-05", end_date="2026-03-05")
    assert result.iloc[0]["MetricStatus"] == "NO_APLICA"
    assert pd.isna(result.iloc[0]["Value"])


def test_F_ypf_digital_actividad_sobre_registrados_funciona_con_dias_calendario():
    engine = _engine(
        [_digital_ypf()],
        [_campana_row("C1", "1047 - TT - 1", FechaInicio=pd.Timestamp("2026-03-01"), FechaFin=pd.Timestamp("2026-03-10"), Estado="Activa")],
    )
    result = engine.query("actividad_sobre_registrados_pct", start_date="2026-03-01", end_date="2026-03-31")
    assert result.iloc[0]["MetricStatus"] == "OK"
    assert result.iloc[0]["Numerator"] == 10
    assert result.iloc[0]["Denominator"] == 31


def test_G_ypf_no_entra_silenciosamente_en_grupo_mezclado_con_circuito_calculable():
    rows = [_digital_pantalla_led("E1"), _digital_ypf("1047 - TT - 1")]
    campanas = [
        _campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"), SalidasVendidas=2.0),
        _campana_row("C2", "1047 - TT - 1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"), SalidasVendidas=2.0),
    ]
    engine = _engine(rows, campanas)
    result = engine.query("fill_rate_slots", start_date="2026-03-05", end_date="2026-03-05")
    assert result.iloc[0]["Denominator"] == 20  # solo Pantalla Led; YPF excluido, no silenciosamente incluido
    assert result.iloc[0]["MetricStatus"] == "PARTIAL"
    assert "politica de circuito" in result.iloc[0]["Warnings"]


# -- H/I/J: fechas incompletas --


def test_H_fecha_inicio_vacia_no_inventa_fecha_y_queda_partial():
    engine = _engine(
        [_static_element()],
        [_campana_row("C1", "E1", FechaInicio=None, FechaFin=pd.Timestamp("2026-03-10"), FechaIndefinida="No", Estado="Activa")],
    )
    result = engine.query("elemento_dias_ocupados", start_date="2026-03-01", end_date="2026-03-31")
    assert result.iloc[0]["Value"] == 0  # la fila no se asigna a ningun periodo, no se inventa fecha
    assert result.iloc[0]["MetricStatus"] == "PARTIAL"
    assert "incompletas" in result.iloc[0]["Warnings"]


def test_I_fecha_fin_vacia_sin_indefinida_queda_partial():
    engine = _engine(
        [_static_element()],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-01"), FechaFin=None, FechaIndefinida="No", Estado="Activa")],
    )
    result = engine.query("elemento_dias_ocupados", start_date="2026-03-01", end_date="2026-03-31")
    assert result.iloc[0]["Value"] == 0
    assert result.iloc[0]["MetricStatus"] == "PARTIAL"
    assert "incompletas" in result.iloc[0]["Warnings"]


def test_J_fecha_fin_vacia_con_indefinida_si_sin_warning_de_fecha():
    engine = _engine(
        [_static_element()],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-20"), FechaFin=None, FechaIndefinida="Si", Estado="Activa")],
    )
    result = engine.query("elemento_dias_ocupados", start_date="2026-03-01", end_date="2026-03-31")
    assert result.iloc[0]["Value"] == 12  # 20..31 inclusive, comportamiento ya existente correcto
    assert result.iloc[0]["MetricStatus"] == "OK"
    assert result.iloc[0]["Warnings"] == ""


# -- K/L: SalidasVendidas vacia --


def test_K_salidas_vendidas_vacia_nunca_se_interpreta_como_cero_segundos():
    engine = _engine(
        [_digital_pantalla_led()],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"), SalidasVendidas=None)],
    )
    result = engine.query("segundos_vendidos", start_date="2026-03-05", end_date="2026-03-05")
    assert result.iloc[0]["MetricStatus"] == "REQUIERE_CONFIRMACION"
    assert pd.isna(result.iloc[0]["Value"])


def test_L_salidas_vendidas_vacia_slots_siguen_contando_la_campana():
    engine = _engine(
        [_digital_pantalla_led()],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"), SalidasVendidas=None)],
    )
    result = engine.query("slots_ocupados", start_date="2026-03-05", end_date="2026-03-05")
    assert result.iloc[0]["Value"] == 1
    assert result.iloc[0]["MetricStatus"] == "OK"


# -- produccion real: YPF/AA2000 con la politica y el gate ya activos --


def test_produccion_ypf_digital_slots_no_aplica(production_engine):
    df = production_engine.query(
        "fill_rate_slots", filters={"CircuitoNegocio": "YPF", "Medio": "Digital"}, universe="COMPLETO_HISTORICO",
        start_date="2026-07-01", end_date="2026-07-31",
    )
    assert df.iloc[0]["MetricStatus"] == "NO_APLICA"
    assert pd.isna(df.iloc[0]["Value"])


def test_produccion_ypf_actividad_sobre_registrados_por_medio(production_engine):
    df = production_engine.query(
        "actividad_sobre_registrados_pct", group_by=["Medio"], filters={"CircuitoNegocio": "YPF"},
        universe="COMPLETO_HISTORICO", start_date="2026-07-01", end_date="2026-07-31",
    )
    assert set(df["Medio"]) == {"Digital", "Estático"}
    assert (df["MetricStatus"] != "NO_APLICA").all()


def test_produccion_aa2000_ocupacion_calendario_no_se_presenta_como_total(production_engine):
    df = production_engine.query(
        "ocupacion_calendario_pct", filters={"CircuitoNegocio": "AA2000"}, universe="COMPLETO_HISTORICO",
        start_date="2026-07-01", end_date="2026-07-31",
    )
    assert df.iloc[0]["MetricStatus"] == "NO_APLICA"
    assert pd.isna(df.iloc[0]["Value"])


def test_produccion_aa2000_digital_slots_calculables_sobre_registrados_con_partial(production_engine):
    df = production_engine.query(
        "slots_comerciales", filters={"CircuitoNegocio": "AA2000", "Medio": "Digital"}, universe="COMPLETO_HISTORICO",
        start_date="2026-07-01", end_date="2026-07-31",
    )
    assert df.iloc[0]["Value"] > 0
    assert df.iloc[0]["MetricStatus"] == "PARTIAL"


# ---------------------------------------------------------------------------
# Gate3B.1.1 - orden de aplicacion: politicas ANTES de inspeccionar campanas
# ---------------------------------------------------------------------------


def test_A_ypf_exclusividad_no_lanza_excepcion_queda_no_aplica():
    """YPF esta bloqueado por politica para fill_rate_slots: su campana con
    exclusividad no debe ni siquiera inspeccionarse (_digital_period_activity
    nunca deberia recibir ese ElementoID), por lo tanto no debe lanzar
    UnsupportedBusinessCaseError."""
    engine = _engine(
        [_digital_ypf()],
        [_campana_row(
            "C1", "1047 - TT - 1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"),
            TipoExclusividad="Día completo",
        )],
    )
    result = engine.query("fill_rate_slots", start_date="2026-03-05", end_date="2026-03-05")
    assert result.iloc[0]["MetricStatus"] == "NO_APLICA"
    assert pd.isna(result.iloc[0]["Value"])


def test_B_grupo_mixto_ypf_exclusividad_no_rompe_el_resto_del_grupo():
    """Grupo mixto PANTALLAS_LED + YPF sin agrupar por circuito: YPF tiene
    una campana con exclusividad. El calculo debe hacerse solo sobre Pantalla
    Led (PARTIAL por exclusion de YPF), sin lanzar excepcion por la campana
    de YPF que nunca deberia inspeccionarse."""
    rows = [_digital_pantalla_led("E1"), _digital_ypf("1047 - TT - 1")]
    campanas = [
        _campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"), SalidasVendidas=2.0),
        _campana_row(
            "C2", "1047 - TT - 1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"),
            TipoExclusividad="Día completo",
        ),
    ]
    engine = _engine(rows, campanas)
    result = engine.query("fill_rate_slots", start_date="2026-03-05", end_date="2026-03-05")
    assert result.iloc[0]["Denominator"] == 20  # solo Pantalla Led; YPF excluido antes de mirar sus campanas
    assert result.iloc[0]["MetricStatus"] == "PARTIAL"
    assert "politica de circuito" in result.iloc[0]["Warnings"]


def test_C_slots_comerciales_no_depende_de_campanas_con_exclusividad():
    """Capacidad pura: no debe llamar a _digital_period_activity en absoluto,
    por lo que una campana con exclusividad en el periodo no debe afectarla
    ni lanzar excepcion."""
    engine = _engine(
        [_digital_pantalla_led()],
        [_campana_row(
            "C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"),
            TipoExclusividad="Día completo",
        )],
    )
    result = engine.query("slots_comerciales", start_date="2026-03-05", end_date="2026-03-05")
    assert result.iloc[0]["Value"] == 20
    assert result.iloc[0]["MetricStatus"] == "OK"


def test_D_segundos_comerciales_no_depende_de_campanas_con_exclusividad():
    engine = _engine(
        [_digital_pantalla_led()],
        [_campana_row(
            "C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"),
            TipoExclusividad="Día completo",
        )],
    )
    result = engine.query("segundos_comerciales", start_date="2026-03-05", end_date="2026-03-05")
    assert result.iloc[0]["Value"] == 72000
    assert result.iloc[0]["MetricStatus"] == "OK"


def test_capacidad_pura_no_depende_de_spot_duracion_no_soportada():
    """Mismo principio que C/D pero con DuracionSpotSeg!=10 en vez de
    exclusividad: la capacidad pura tampoco debe inspeccionar esa campana."""
    engine = _engine(
        [_digital_pantalla_led()],
        [_campana_row(
            "C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"),
            DuracionSpotSeg=20.0, SalidasVendidas=2.0,
        )],
    )
    slots = engine.query("slots_comerciales", start_date="2026-03-05", end_date="2026-03-05")
    segundos = engine.query("segundos_comerciales", start_date="2026-03-05", end_date="2026-03-05")
    assert slots.iloc[0]["Value"] == 20
    assert segundos.iloc[0]["Value"] == 72000


def test_ypf_spot_duracion_no_soportada_no_lanza_excepcion_en_fill_rate_slots():
    """Mismo principio que A pero con DuracionSpotSeg!=10 en vez de
    exclusividad."""
    engine = _engine(
        [_digital_ypf()],
        [_campana_row(
            "C1", "1047 - TT - 1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"),
            DuracionSpotSeg=20.0, SalidasVendidas=2.0,
        )],
    )
    result = engine.query("fill_rate_slots", start_date="2026-03-05", end_date="2026-03-05")
    assert result.iloc[0]["MetricStatus"] == "NO_APLICA"
    assert pd.isna(result.iloc[0]["Value"])


def test_no_policy_no_capacity_metric_sigue_lanzando_excepcion_esperada():
    """Control negativo: un circuito calculable (no bloqueado por politica)
    consultando una metrica de actividad (no de capacidad pura) SI debe
    seguir lanzando UnsupportedBusinessCaseError ante exclusividad, tal como
    Gate3B/Gate3B.1 ya exigian."""
    engine = _engine(
        [_digital_pantalla_led()],
        [_campana_row(
            "C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"),
            TipoExclusividad="Día completo",
        )],
    )
    with pytest.raises(UnsupportedBusinessCaseError):
        engine.query("slots_ocupados", start_date="2026-03-05", end_date="2026-03-05")
