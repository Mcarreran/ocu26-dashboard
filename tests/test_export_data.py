"""Pruebas para scripts/export_data.py (Gate 4A - capa de salida OCU26).

No modifica scripts/validate_input.py, scripts/transform_data.py,
scripts/semantic_model.py, scripts/metrics_engine.py,
config/business_semantics.json ni input/OCU26_BASE_DATOS.xlsx. Los
fixtures sinteticos usan la CONFIGURACION REAL (sm.load_config()), igual
patron que test_metrics_engine.py, para ejercitar reglas de negocio reales.
Los tests de round-trip/produccion corren contra el archivo productivo real.
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
import metrics_engine as me  # noqa: E402
import validate_input as vi  # noqa: E402
import export_data as ed  # noqa: E402
from metrics_engine import MetricsEngine  # noqa: E402
from test_semantic_model import _campana_row, _maestro_row, _transform_result  # noqa: E402

PRODUCTION_FILE = REPO_ROOT / "input" / "OCU26_BASE_DATOS.xlsx"


def _semantic(maestro_rows: list[dict], campanas_rows: list[dict] | None = None) -> dict:
    config = sm.load_config()
    return sm.build_semantic_model(_transform_result(maestro_rows, campanas_rows or []), config)


def _digital_pantalla_led(elemento_id="E1", **overrides):
    row = dict(
        CircuitoDashboard="Pantalla Led", Subcircuito="PLED", Ubicacion="CABILDO",
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        CapacidadSlotsReel=40, SegundosDia=75000,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _static_cencosud(elemento_id="E1", **overrides):
    row = dict(
        CircuitoDashboard="Shoppings Estático", Subcircuito="CENCOSUD", Ubicacion="UNICENTER",
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


@pytest.fixture(scope="module")
def production_pipeline():
    transform_result, semantic_result, engine = ed.load_pipeline(PRODUCTION_FILE)
    dim_elementos = ed.build_dim_elementos(semantic_result, engine)
    fact_campanas = ed.build_fact_campanas(semantic_result, dim_elementos)
    rango_inicio, rango_fin = ed.resolve_rango_temporal(fact_campanas)
    dim_calendario = ed.build_dim_calendario(rango_inicio, rango_fin)
    bridge = ed.build_bridge_campana_dia(engine, fact_campanas, rango_inicio, rango_fin)
    fact_metricas = ed.build_fact_metricas_diaria(engine, dim_elementos, rango_inicio, rango_fin)
    return dict(
        transform_result=transform_result, semantic_result=semantic_result, engine=engine,
        dim_elementos=dim_elementos, fact_campanas=fact_campanas, dim_calendario=dim_calendario,
        bridge=bridge, fact_metricas=fact_metricas, rango_inicio=rango_inicio, rango_fin=rango_fin,
    )


# ---------------------------------------------------------------------------
# Unicidad de PK / integridad referencial
# ---------------------------------------------------------------------------


def test_dim_elementos_falla_si_hay_elementoid_duplicado():
    semantic_result = _semantic([_digital_pantalla_led("E1"), _digital_pantalla_led("E1")])
    engine = MetricsEngine(semantic_result)
    with pytest.raises(ed.ExportError, match="ElementoID duplicado"):
        ed.build_dim_elementos(semantic_result, engine)


def test_dim_elementos_pk_unica_en_produccion(production_pipeline):
    dim_elementos = production_pipeline["dim_elementos"]
    assert dim_elementos["ElementoID"].is_unique
    assert dim_elementos["ElementoID"].notna().all()


def test_fact_campanas_falla_si_hay_cargaid_duplicado():
    semantic_result = _semantic(
        [_digital_pantalla_led("E1")],
        [_campana_row("C1", "E1"), _campana_row("C1", "E1")],
    )
    engine = MetricsEngine(semantic_result)
    dim_elementos = ed.build_dim_elementos(semantic_result, engine)
    with pytest.raises(ed.ExportError, match="CargaID duplicado"):
        ed.build_fact_campanas(semantic_result, dim_elementos)


def test_fact_campanas_pk_unica_en_produccion(production_pipeline):
    fact_campanas = production_pipeline["fact_campanas"]
    assert fact_campanas["CargaID"].is_unique
    assert fact_campanas["CargaID"].notna().all()


def test_fact_campanas_falla_si_elementoid_no_existe_en_dim_elementos():
    semantic_result = _semantic(
        [_digital_pantalla_led("E1")],
        [_campana_row("C1", "E1"), _campana_row("C2", "E999")],
    )
    engine = MetricsEngine(semantic_result)
    dim_elementos = ed.build_dim_elementos(semantic_result, engine)
    with pytest.raises(ed.ExportError, match="sin correspondencia en DIM_ELEMENTOS"):
        ed.build_fact_campanas(semantic_result, dim_elementos)


def test_fact_campanas_integridad_referencial_en_produccion(production_pipeline):
    ids_validos = set(production_pipeline["dim_elementos"]["ElementoID"])
    huerfanos = production_pipeline["fact_campanas"].loc[
        production_pipeline["fact_campanas"]["ElementoID"].notna()
        & ~production_pipeline["fact_campanas"]["ElementoID"].isin(ids_validos)
    ]
    assert huerfanos.empty


def test_bridge_campana_dia_referencia_solo_cargaid_existentes(production_pipeline):
    ids_validos = set(production_pipeline["fact_campanas"]["CargaID"])
    assert set(production_pipeline["bridge"]["CargaID"]).issubset(ids_validos)


def test_fact_metricas_diaria_referencia_solo_elementoid_existentes(production_pipeline):
    ids_validos = set(production_pipeline["dim_elementos"]["ElementoID"])
    assert set(production_pipeline["fact_metricas"]["ElementoID"]).issubset(ids_validos)


# ---------------------------------------------------------------------------
# IDCampaña vacio conservado
# ---------------------------------------------------------------------------


# Recalibrado 2026-08-18 (promocion base OCU26+YPF): 26 -> 25. Una de las
# filas legacy con IDCampaña vacio pertenecia al bloque YPF 10000-10009,
# retirado; las 13.616 filas nuevas de YPF Etapa 2 y las 4 filas CENCOSUD ya
# autorizadas en FINAL_V2 tienen IDCampaña siempre completo (0 vacios).
def test_25_idcampana_vacias_se_conservan_en_produccion(production_pipeline):
    fact_campanas = production_pipeline["fact_campanas"]
    vacias = fact_campanas[fact_campanas["IDCampaña"].isna()]
    assert len(vacias) == 25
    assert vacias["CargaID"].notna().all()
    assert vacias["CargaID"].is_unique


def test_idcampana_vacia_no_se_completa_ni_se_inventa():
    semantic_result = _semantic(
        [_digital_pantalla_led("E1")],
        [_campana_row("C1", "E1", IDCampaña=None)],
    )
    engine = MetricsEngine(semantic_result)
    dim_elementos = ed.build_dim_elementos(semantic_result, engine)
    fact_campanas = ed.build_fact_campanas(semantic_result, dim_elementos)
    assert len(fact_campanas) == 1
    assert pd.isna(fact_campanas.iloc[0]["IDCampaña"])
    assert fact_campanas.iloc[0]["CargaID"] == "C1"


# ---------------------------------------------------------------------------
# Crecimiento dinamico del maestro (Gate 4 no hardcodea cantidades/IDs)
# ---------------------------------------------------------------------------


def test_dim_elementos_absorbe_elemento_nuevo_sin_cambiar_codigo():
    semantic_result = _semantic([_digital_pantalla_led("E1"), _digital_pantalla_led("E2")])
    engine = MetricsEngine(semantic_result)
    dim_elementos = ed.build_dim_elementos(semantic_result, engine)
    assert len(dim_elementos) == 2
    assert set(dim_elementos["ElementoID"]) == {"E1", "E2"}


def test_elemento_de_circuito_no_reconocido_queda_no_clasificado_no_oculto():
    fila = _maestro_row("E1", CircuitoDashboard="CIRCUITO_INEXISTENTE_2027", Subcircuito="NUEVO", Medio="Digital")
    semantic_result = _semantic([fila])
    engine = MetricsEngine(semantic_result)
    dim_elementos = ed.build_dim_elementos(semantic_result, engine)
    fila_exportada = dim_elementos.iloc[0]
    assert fila_exportada["CircuitoNegocio"] == "NO_CLASIFICADO"
    assert fila_exportada["IncluyePerformanceCore"] == False  # noqa: E712
    assert fila_exportada["ElementoID"] == "E1"  # sigue visible, no se oculta


def test_produccion_row_counts_no_hardcodeados(production_pipeline):
    # Los conteos deben surgir de los datos, no de una constante en el codigo.
    esperado_maestro = len(production_pipeline["semantic_result"]["maestro"])
    assert len(production_pipeline["dim_elementos"]) == esperado_maestro
    esperado_campanas = len(production_pipeline["semantic_result"]["campanas"])
    assert len(production_pipeline["fact_campanas"]) == esperado_campanas


# ---------------------------------------------------------------------------
# Capacidades BI-safe: nullable numerico, nunca 0 para "desconocida"
# ---------------------------------------------------------------------------


def test_capacidad_desconocida_es_null_no_cero():
    fila = _maestro_row(
        "E1", CircuitoDashboard="London Supply", Subcircuito="USH", Ubicacion="USH", Medio="Digital",
        TipoCatalogo="Abierto", TipoInventario="Digital", Descripcion="Pantalla Digital",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    semantic_result = _semantic([fila])
    engine = MetricsEngine(semantic_result)
    dim_elementos = ed.build_dim_elementos(semantic_result, engine)
    fila_exportada = dim_elementos.iloc[0]
    assert fila_exportada["CapacidadSlotsDesconocida"] == True  # noqa: E712
    assert pd.isna(fila_exportada["SlotsComercialesValor"])  # NULL, nunca 0


def test_capacidad_conocida_es_numerica(production_pipeline):
    dim_elementos = production_pipeline["dim_elementos"]
    conocidos = dim_elementos[(~dim_elementos["CapacidadSlotsDesconocida"]) & (dim_elementos["Medio"] == "Digital")]
    valores = conocidos["SlotsComercialesValor"].dropna()
    assert len(valores) > 0
    assert pd.api.types.is_integer_dtype(valores)


def test_slotscomercialesvalor_nunca_mezcla_texto(production_pipeline):
    dim_elementos = production_pipeline["dim_elementos"]
    assert dim_elementos["SlotsComercialesValor"].apply(lambda v: isinstance(v, str)).sum() == 0
    assert str(dim_elementos["SlotsComercialesValor"].dtype) == "Int64"
    assert str(dim_elementos["SegundosComercialesValor"].dtype) == "Float64"


def test_8_elementos_con_capacidad_desconocida_en_produccion(production_pipeline):
    dim_elementos = production_pipeline["dim_elementos"]
    assert int(dim_elementos["CapacidadSlotsDesconocida"].sum()) == 8


# ---------------------------------------------------------------------------
# BValor/HValor/M2Valor: derivadas de b/h/m2, nullable, nunca 0. q sin QValor
# (semantica de negocio A VALIDAR). DimensionOptico/DimensionTotal/
# Observaciones permanecen solo texto, sin columna numerica derivada.
# ---------------------------------------------------------------------------


def _fila_dimensiones(elemento_id, **overrides):
    row = dict(CircuitoDashboard="Shoppings Estático", Subcircuito="CENCOSUD", Ubicacion="UNICENTER", Medio="Estático")
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


@pytest.fixture(scope="module")
def dim_elementos_dimensiones():
    filas = [
        _fila_dimensiones("E1", b=1.45, h=2.5, m2=10.0, q=3, DimensionOptico=0, DimensionTotal="58,5 x 97 cm", Observaciones="Texto libre"),
        _fila_dimensiones("E2", b="No aplica", h="No aplica", m2="No aplica", q="No aplica", DimensionOptico="Ver plano", DimensionTotal=0, Observaciones=0),
        _fila_dimensiones("E3", b="\xa0", h="\xa0", m2="\xa0", q="No"),
    ]
    semantic_result = _semantic(filas)
    engine = MetricsEngine(semantic_result)
    return ed.build_dim_elementos(semantic_result, engine)


def test_bvalor_hvalor_m2valor_dtype_numerico_nullable(dim_elementos_dimensiones):
    assert str(dim_elementos_dimensiones["BValor"].dtype) == "Float64"
    assert str(dim_elementos_dimensiones["HValor"].dtype) == "Float64"
    assert str(dim_elementos_dimensiones["M2Valor"].dtype) == "Float64"


def test_valor_numerico_valido_se_preserva_exacto(dim_elementos_dimensiones):
    fila_e1 = dim_elementos_dimensiones[dim_elementos_dimensiones["ElementoID"] == "E1"].iloc[0]
    assert fila_e1["BValor"] == 1.45
    assert fila_e1["HValor"] == 2.5
    assert fila_e1["M2Valor"] == 10.0


def test_no_aplica_se_convierte_a_na_nunca_a_cero(dim_elementos_dimensiones):
    fila_e2 = dim_elementos_dimensiones[dim_elementos_dimensiones["ElementoID"] == "E2"].iloc[0]
    # pd.isna(...) True ya prueba que no es 0: pd.NA no es igual a ningun numero,
    # incluido 0 (fila_e2["BValor"] != 0 devolveria pd.NA, no bool, si se comparara).
    assert pd.isna(fila_e2["BValor"])
    assert pd.isna(fila_e2["HValor"])
    assert pd.isna(fila_e2["M2Valor"])


def test_espacio_no_separable_se_convierte_a_na_nunca_a_cero(dim_elementos_dimensiones):
    fila_e3 = dim_elementos_dimensiones[dim_elementos_dimensiones["ElementoID"] == "E3"].iloc[0]
    assert pd.isna(fila_e3["BValor"])
    assert pd.isna(fila_e3["HValor"])
    assert pd.isna(fila_e3["M2Valor"])


def test_b_h_m2_originales_siguen_presentes_como_texto(dim_elementos_dimensiones):
    for col in ("b", "h", "m2"):
        assert col in dim_elementos_dimensiones.columns
        assert str(dim_elementos_dimensiones[col].dtype) == "string"


def test_q_sigue_presente_sin_qvalor(dim_elementos_dimensiones):
    assert "q" in dim_elementos_dimensiones.columns
    assert "QValor" not in dim_elementos_dimensiones.columns
    assert str(dim_elementos_dimensiones["q"].dtype) == "string"


def test_dimension_optico_total_observaciones_permanecen_texto(dim_elementos_dimensiones):
    for col in ("DimensionOptico", "DimensionTotal", "Observaciones"):
        assert str(dim_elementos_dimensiones[col].dtype) == "string"
        assert (col + "Valor") not in dim_elementos_dimensiones.columns


def test_produccion_bvalor_hvalor_m2valor_convertibilidad(production_pipeline):
    dim_elementos = production_pipeline["dim_elementos"]
    assert int(dim_elementos["BValor"].notna().sum()) == 617
    assert int(dim_elementos["HValor"].notna().sum()) == 617
    assert int(dim_elementos["M2Valor"].notna().sum()) == 631
    assert "QValor" not in dim_elementos.columns


# ---------------------------------------------------------------------------
# BRIDGE_CAMPANA_DIA: fidelidad temporal exacta con Gate 3
# ---------------------------------------------------------------------------


def test_bridge_fecha_inicio_igual_fin_es_un_dia():
    semantic_result = _semantic(
        [_static_cencosud("E1")],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-10"), FechaFin=pd.Timestamp("2026-03-10"), Estado="Activa")],
    )
    engine = MetricsEngine(semantic_result)
    dim_elementos = ed.build_dim_elementos(semantic_result, engine)
    fact_campanas = ed.build_fact_campanas(semantic_result, dim_elementos)
    bridge = ed.build_bridge_campana_dia(engine, fact_campanas, pd.Timestamp("2026-03-01"), pd.Timestamp("2026-03-31"))
    assert len(bridge) == 1
    assert bridge.iloc[0]["Fecha"] == pd.Timestamp("2026-03-10")


def test_bridge_cancelado_no_genera_filas():
    semantic_result = _semantic(
        [_static_cencosud("E1")],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-01"), FechaFin=pd.Timestamp("2026-03-10"), Estado="Cancelado")],
    )
    engine = MetricsEngine(semantic_result)
    dim_elementos = ed.build_dim_elementos(semantic_result, engine)
    fact_campanas = ed.build_fact_campanas(semantic_result, dim_elementos)
    bridge = ed.build_bridge_campana_dia(engine, fact_campanas, pd.Timestamp("2026-03-01"), pd.Timestamp("2026-03-31"))
    assert bridge.empty


def test_bridge_fecha_indefinida_se_extiende_hasta_rango_fin():
    semantic_result = _semantic(
        [_static_cencosud("E1")],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-20"), FechaFin=None, FechaIndefinida="Si", Estado="Activa")],
    )
    engine = MetricsEngine(semantic_result)
    dim_elementos = ed.build_dim_elementos(semantic_result, engine)
    fact_campanas = ed.build_fact_campanas(semantic_result, dim_elementos)
    bridge = ed.build_bridge_campana_dia(engine, fact_campanas, pd.Timestamp("2026-03-01"), pd.Timestamp("2026-03-31"))
    assert len(bridge) == 12  # 20..31 inclusive, igual que el motor
    assert bridge["Fecha"].max() == pd.Timestamp("2026-03-31")


def test_bridge_reservada_genera_pertenencia_temporal():
    semantic_result = _semantic(
        [_static_cencosud("E1")],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-09-01"), FechaFin=pd.Timestamp("2026-09-10"), Estado="Reservada")],
    )
    engine = MetricsEngine(semantic_result)
    dim_elementos = ed.build_dim_elementos(semantic_result, engine)
    fact_campanas = ed.build_fact_campanas(semantic_result, dim_elementos)
    bridge = ed.build_bridge_campana_dia(engine, fact_campanas, pd.Timestamp("2026-09-01"), pd.Timestamp("2026-09-30"))
    assert len(bridge) == 10


def test_bridge_fecha_incompleta_no_inventa_pertenencia():
    semantic_result = _semantic(
        [_static_cencosud("E1")],
        [_campana_row("C1", "E1", FechaInicio=None, FechaFin=pd.Timestamp("2026-03-10"), FechaIndefinida="No", Estado="Activa")],
    )
    engine = MetricsEngine(semantic_result)
    dim_elementos = ed.build_dim_elementos(semantic_result, engine)
    fact_campanas = ed.build_fact_campanas(semantic_result, dim_elementos)
    bridge = ed.build_bridge_campana_dia(engine, fact_campanas, pd.Timestamp("2026-03-01"), pd.Timestamp("2026-03-31"))
    assert bridge.empty  # fecha incompleta: no se inventa, no se cuenta


# Recalibrado 2026-08-18 (promocion base OCU26+YPF): 881.210 -> 1.185.499.
# Verificado con calculo independiente en pandas directo (suma de dias por
# CargaID sobre el rango resuelto, sin pasar por build_bridge_campana_dia):
# coincide exacto, 1.185.499.
def test_bridge_volumen_produccion(production_pipeline):
    assert len(production_pipeline["bridge"]) == 1_185_499
    assert list(production_pipeline["bridge"].columns) == ["CargaID", "Fecha"]


# ---------------------------------------------------------------------------
# FACT_METRICAS_DIARIA: sparse activity, componentes crudos
# ---------------------------------------------------------------------------


def test_dia_sin_fila_significa_actividad_cero_no_dia_inexistente():
    semantic_result = _semantic(
        [_static_cencosud("E1")],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"), Estado="Activa")],
    )
    engine = MetricsEngine(semantic_result)
    dim_elementos = ed.build_dim_elementos(semantic_result, engine)
    fact = ed.build_fact_metricas_diaria(engine, dim_elementos, pd.Timestamp("2026-03-01"), pd.Timestamp("2026-03-10"))
    assert len(fact) == 1  # sparse: solo el dia ocupado tiene fila
    assert fact.iloc[0]["Fecha"] == pd.Timestamp("2026-03-05")
    dias_calendario = pd.date_range("2026-03-01", "2026-03-10")
    dias_con_fila = set(fact["Fecha"])
    assert dias_con_fila.issubset(set(dias_calendario))
    assert len(dias_calendario) - len(dias_con_fila) == 9  # 9 dias sin fila = 9 dias en 0, no "inexistentes"


def test_elemento_estatico_no_lleva_componentes_digitales():
    semantic_result = _semantic(
        [_static_cencosud("E1")],
        [_campana_row("C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"), Estado="Activa")],
    )
    engine = MetricsEngine(semantic_result)
    dim_elementos = ed.build_dim_elementos(semantic_result, engine)
    fact = ed.build_fact_metricas_diaria(engine, dim_elementos, pd.Timestamp("2026-03-01"), pd.Timestamp("2026-03-10"))
    assert fact.iloc[0]["DiasOcupado"] == 1
    assert pd.isna(fact.iloc[0]["SlotsOcupadosDia"])
    assert pd.isna(fact.iloc[0]["SegundosVendidosDia"])
    assert pd.isna(fact.iloc[0]["HasSalidasIndeterminada"])


# Recalibrado 2026-08-18 (promocion base OCU26+YPF): 573.675 -> 737.475
# (digital 520.648 -> 677.894, no-digital 53.027 -> 59.581). Verificado con
# calculo independiente en pandas directo (pares unicos ElementoID+dia por
# explosion de rango, sin pasar por build_fact_metricas_diaria): coincide
# exacto en los 3 valores.
def test_fact_metricas_diaria_volumen_produccion(production_pipeline):
    fact = production_pipeline["fact_metricas"]
    assert len(fact) == 737_475
    digital = int(fact["ElementoID"].map(production_pipeline["dim_elementos"].set_index("ElementoID")["Medio"]).eq("Digital").sum())
    assert digital == 677_894
    assert len(fact) - digital == 59_581


def test_fact_metricas_diaria_falla_si_hay_exclusividad_en_el_rango():
    fila = _digital_pantalla_led("E1")
    campana = _campana_row(
        "C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"),
        Estado="Activa", TipoExclusividad="Día completo",
    )
    semantic_result = _semantic([fila], [campana])
    engine = MetricsEngine(semantic_result)
    dim_elementos = ed.build_dim_elementos(semantic_result, engine)
    with pytest.raises(ed.ExportError, match="exclusividad"):
        ed.build_fact_metricas_diaria(engine, dim_elementos, pd.Timestamp("2026-03-01"), pd.Timestamp("2026-03-10"))


def test_fact_metricas_diaria_falla_si_spot_duracion_no_soportada():
    fila = _digital_pantalla_led("E1")
    campana = _campana_row(
        "C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"),
        Estado="Activa", DuracionSpotSeg=20.0, SalidasVendidas=2.0,
    )
    semantic_result = _semantic([fila], [campana])
    engine = MetricsEngine(semantic_result)
    dim_elementos = ed.build_dim_elementos(semantic_result, engine)
    with pytest.raises(ed.ExportError, match="DuracionSpotSeg"):
        ed.build_fact_metricas_diaria(engine, dim_elementos, pd.Timestamp("2026-03-01"), pd.Timestamp("2026-03-10"))


# ---------------------------------------------------------------------------
# REQUIERE_CONFIRMACION sintetico (BLOQUEANTE) - SalidasVendidas vacia
# ---------------------------------------------------------------------------


def test_requiere_confirmacion_sintetico_salidas_vendidas_vacia():
    fila = _digital_pantalla_led("E1")
    campana = _campana_row(
        "C1", "E1", FechaInicio=pd.Timestamp("2026-03-05"), FechaFin=pd.Timestamp("2026-03-05"),
        Estado="Activa", DuracionSpotSeg=10.0, SalidasVendidas=None,
    )
    semantic_result = _semantic([fila], [campana])
    engine = MetricsEngine(semantic_result)
    dim_elementos = ed.build_dim_elementos(semantic_result, engine)
    fact = ed.build_fact_metricas_diaria(engine, dim_elementos, pd.Timestamp("2026-03-01"), pd.Timestamp("2026-03-10"))

    fila_dia = fact[(fact["ElementoID"] == "E1") & (fact["Fecha"] == pd.Timestamp("2026-03-05"))].iloc[0]
    assert fila_dia["HasSalidasIndeterminada"] == True  # noqa: E712
    # nunca se interpreta como 0: el componente numerico puede ser 0.0, pero el flag lo distingue
    assert fila_dia["SegundosVendidosDia"] == 0.0
    # slots SI son calculables (no dependen de SalidasVendidas)
    assert fila_dia["SlotsOcupadosDia"] == 1

    # ground truth: MetricsEngine real sobre los mismos datos sinteticos
    directo_segundos = engine.query("segundos_vendidos", start_date="2026-03-05", end_date="2026-03-05")
    assert directo_segundos.iloc[0]["MetricStatus"] == "REQUIERE_CONFIRMACION"
    directo_slots = engine.query("slots_ocupados", start_date="2026-03-05", end_date="2026-03-05")
    assert directo_slots.iloc[0]["MetricStatus"] == "OK"
    assert directo_slots.iloc[0]["Value"] == 1

    # reconstruccion generica (lo que haria la medida DAX) a partir del componente persistido
    status_reconstruido = "REQUIERE_CONFIRMACION" if fila_dia["HasSalidasIndeterminada"] else "OK"
    assert status_reconstruido == directo_segundos.iloc[0]["MetricStatus"]


# ---------------------------------------------------------------------------
# Round-trip: reconstruccion generica de Value/Numerator/Denominator/MetricStatus
# (mismos componentes validados en Gate 4A.0, ahora contra el output real de Gate 4)
# ---------------------------------------------------------------------------


def _peor_status(estados: list[str]) -> str:
    peor = "OK"
    for estado in estados:
        if me._STATUS_RANK[estado] > me._STATUS_RANK[peor]:
            peor = estado
    return peor


def _status_calendario(elemento_ids, dim_elementos) -> str:
    sub = dim_elementos[dim_elementos["ElementoID"].isin(elemento_ids)]
    if sub.empty:
        return "NO_APLICA"
    cobertura_ok = (sub["CoberturaCatalogo"] == "COMPLETO").all() and (sub["CompletitudMaestro"] == "COMPLETO").all()
    if not cobertura_ok:
        return "NO_APLICA"
    return "PARTIAL" if sub["FechaIncompletaCalendario"].any() else "OK"


def _status_actividad(elemento_ids, dim_elementos) -> str:
    sub = dim_elementos[dim_elementos["ElementoID"].isin(elemento_ids)]
    if sub.empty:
        return "NO_APLICA"
    return "PARTIAL" if sub["FechaIncompletaCalendario"].any() else "OK"


def _status_digital_slots(elemento_ids, dim_elementos, sensible_a_politica=True) -> str:
    sub = dim_elementos[dim_elementos["ElementoID"].isin(elemento_ids)]
    total = len(sub)
    if total == 0:
        return "NO_APLICA"
    bloqueados = sub["PolicyBloqueadaSlotSeconds"] if sensible_a_politica else pd.Series(False, index=sub.index)
    n_bloqueados = int(bloqueados.sum())
    if sensible_a_politica and n_bloqueados == total:
        return "NO_APLICA"
    usable = sub[~bloqueados] if sensible_a_politica else sub
    conocidos = usable[~usable["CapacidadSlotsDesconocida"]]
    if len(conocidos) == 0:
        return "NO_APLICA"
    hay_razon_partial = (
        n_bloqueados > 0
        or (usable["CompletitudMaestro"] != "COMPLETO").any()
        or usable["CapacidadSlotsDesconocida"].any()
        or usable["FechaIncompletaDigital"].any()
    )
    return "PARTIAL" if hay_razon_partial else "OK"


def _reconstruir_calendario(fact_metricas, dim_calendario, elemento_ids, start, end):
    dias = dim_calendario[(dim_calendario["Fecha"] >= pd.Timestamp(start)) & (dim_calendario["Fecha"] <= pd.Timestamp(end))]
    n_dias = len(dias)
    sub = fact_metricas[
        fact_metricas["ElementoID"].isin(elemento_ids)
        & (fact_metricas["Fecha"] >= pd.Timestamp(start))
        & (fact_metricas["Fecha"] <= pd.Timestamp(end))
    ]
    numerador = int(sub["DiasOcupado"].fillna(0).sum())
    denominador = n_dias * len(elemento_ids)
    return numerador, denominador


def _reconstruir_digital_slots(fact_metricas, dim_calendario, dim_elementos, elemento_ids, start, end):
    dias = dim_calendario[(dim_calendario["Fecha"] >= pd.Timestamp(start)) & (dim_calendario["Fecha"] <= pd.Timestamp(end))]
    n_dias = len(dias)
    sub_dim = dim_elementos[dim_elementos["ElementoID"].isin(elemento_ids)]
    usable_ids = set(sub_dim.loc[~sub_dim["PolicyBloqueadaSlotSeconds"], "ElementoID"])
    conocidos_ids = set(
        sub_dim.loc[sub_dim["ElementoID"].isin(usable_ids) & ~sub_dim["CapacidadSlotsDesconocida"], "ElementoID"]
    )
    sub_fact = fact_metricas[
        fact_metricas["ElementoID"].isin(conocidos_ids)
        & (fact_metricas["Fecha"] >= pd.Timestamp(start))
        & (fact_metricas["Fecha"] <= pd.Timestamp(end))
    ]
    suma_por_elemento = sub_fact.groupby("ElementoID")["SlotsOcupadosDia"].sum()
    numerador = 0.0
    for eid in conocidos_ids:
        numerador += float(suma_por_elemento.get(eid, 0) or 0) / n_dias
    denominador = float(
        dim_elementos.loc[dim_elementos["ElementoID"].isin(conocidos_ids), "SlotsComercialesValor"].sum()
    )
    return numerador, denominador


@pytest.mark.parametrize(
    "filtro,metric",
    [
        (dict(Subcircuito="CENCOSUD"), "ocupacion_calendario_pct"),
        (dict(CircuitoNegocio="AA2000"), "ocupacion_calendario_pct"),
        (dict(CircuitoNegocio="AA2000"), "actividad_sobre_registrados_pct"),
    ],
)
def test_round_trip_calendario_produccion(production_pipeline, filtro, metric):
    engine = production_pipeline["engine"]
    dim_elementos = production_pipeline["dim_elementos"]
    fact_metricas = production_pipeline["fact_metricas"]
    dim_calendario = production_pipeline["dim_calendario"]
    inicio, fin = "2026-01-01", "2026-01-10"

    directo = engine.query(metric, group_by=[], filters=filtro, universe="OPERATIVO_GENERAL", start_date=inicio, end_date=fin)

    campo, valor = next(iter(filtro.items()))
    elemento_ids = dim_elementos.loc[dim_elementos[campo] == valor, "ElementoID"]
    if metric in ("ocupacion_calendario_pct",):
        elemento_ids = dim_elementos.loc[(dim_elementos[campo] == valor) & (dim_elementos["Medio"] == "Estático"), "ElementoID"]
        status_reconstruido = _status_calendario(elemento_ids, dim_elementos)
    else:
        status_reconstruido = _status_actividad(elemento_ids, dim_elementos)

    numerador, denominador = _reconstruir_calendario(fact_metricas, dim_calendario, elemento_ids, inicio, fin)

    assert int(directo.iloc[0]["Numerator"]) == numerador if not pd.isna(directo.iloc[0]["Numerator"]) else True
    assert int(directo.iloc[0]["Denominator"]) == denominador
    assert directo.iloc[0]["MetricStatus"] == status_reconstruido


def test_round_trip_digital_pantallas_led_produccion(production_pipeline):
    engine = production_pipeline["engine"]
    dim_elementos = production_pipeline["dim_elementos"]
    fact_metricas = production_pipeline["fact_metricas"]
    dim_calendario = production_pipeline["dim_calendario"]
    inicio, fin = "2026-06-01", "2026-06-05"

    elemento_ids = dim_elementos.loc[dim_elementos["CircuitoNegocio"] == "PANTALLAS_LED", "ElementoID"]
    directo = engine.query("fill_rate_slots", filters=dict(CircuitoNegocio="PANTALLAS_LED"), universe="OPERATIVO_GENERAL", start_date=inicio, end_date=fin)
    numerador, denominador = _reconstruir_digital_slots(fact_metricas, dim_calendario, dim_elementos, elemento_ids, inicio, fin)
    status_reconstruido = _status_digital_slots(elemento_ids, dim_elementos)

    assert directo.iloc[0]["Value"] == pytest.approx(numerador / denominador * 100.0)
    assert directo.iloc[0]["MetricStatus"] == status_reconstruido == "OK"


def test_round_trip_digital_ypf_aislado_no_aplica(production_pipeline):
    dim_elementos = production_pipeline["dim_elementos"]
    engine = production_pipeline["engine"]
    elemento_ids = dim_elementos.loc[
        (dim_elementos["CircuitoNegocio"] == "YPF") & (dim_elementos["Medio"] == "Digital"), "ElementoID"
    ]
    directo = engine.query("fill_rate_slots", filters=dict(CircuitoNegocio="YPF"), universe="OPERATIVO_GENERAL", start_date="2026-06-01", end_date="2026-06-05")
    status_reconstruido = _status_digital_slots(elemento_ids, dim_elementos)
    assert directo.iloc[0]["MetricStatus"] == "NO_APLICA" == status_reconstruido


def test_round_trip_digital_grupo_mixto_pantalla_ypf_partial(production_pipeline):
    dim_elementos = production_pipeline["dim_elementos"]
    engine = production_pipeline["engine"]
    elemento_ids = dim_elementos.loc[
        dim_elementos["CircuitoNegocio"].isin(["PANTALLAS_LED", "YPF"]), "ElementoID"
    ]
    directo = engine.query(
        "fill_rate_slots", filters=dict(CircuitoNegocio=["PANTALLAS_LED", "YPF"]),
        universe="OPERATIVO_GENERAL", start_date="2026-06-01", end_date="2026-06-03",
    )
    status_reconstruido = _status_digital_slots(elemento_ids, dim_elementos)
    assert directo.iloc[0]["MetricStatus"] == "PARTIAL" == status_reconstruido


def test_round_trip_digital_london_supply_capacidad_desconocida_partial(production_pipeline):
    dim_elementos = production_pipeline["dim_elementos"]
    engine = production_pipeline["engine"]
    elemento_ids = dim_elementos.loc[
        (dim_elementos["CircuitoNegocio"] == "LONDON_SUPPLY") & (dim_elementos["Medio"] == "Digital"), "ElementoID"
    ]
    directo = engine.query(
        "fill_rate_slots", filters=dict(CircuitoNegocio="LONDON_SUPPLY"),
        universe="OPERATIVO_GENERAL", start_date="2026-06-01", end_date="2026-06-01",
    )
    status_reconstruido = _status_digital_slots(elemento_ids, dim_elementos)
    assert directo.iloc[0]["MetricStatus"] == "PARTIAL" == status_reconstruido


# ---------------------------------------------------------------------------
# Idempotencia
# ---------------------------------------------------------------------------


def test_idempotencia_dim_elementos(production_pipeline):
    semantic_result = production_pipeline["semantic_result"]
    engine = production_pipeline["engine"]
    primera = ed.build_dim_elementos(semantic_result, engine)
    segunda = ed.build_dim_elementos(semantic_result, engine)
    assert primera.equals(segunda)


def test_idempotencia_fact_metricas_diaria(production_pipeline):
    engine = production_pipeline["engine"]
    dim_elementos = production_pipeline["dim_elementos"]
    rango_inicio, rango_fin = production_pipeline["rango_inicio"], production_pipeline["rango_fin"]
    primera = ed.build_fact_metricas_diaria(engine, dim_elementos, rango_inicio, rango_fin)
    segunda = ed.build_fact_metricas_diaria(engine, dim_elementos, rango_inicio, rango_fin)
    assert primera.equals(segunda)


# ---------------------------------------------------------------------------
# Tipos Parquet (round-trip real, sin tocar el repo: tmp_path de pytest)
# ---------------------------------------------------------------------------


def test_parquet_roundtrip_preserva_tipos_nullable(tmp_path):
    df = pd.DataFrame(
        {
            "ElementoID": pd.array(["E1", "E2"], dtype="string"),
            "SlotsComercialesValor": pd.array([20, None], dtype="Int64"),
            "SegundosComercialesValor": pd.array([72000.0, None], dtype="Float64"),
            "CapacidadSlotsDesconocida": pd.array([False, True], dtype="boolean"),
            "Fecha": pd.to_datetime(["2026-01-01", "2026-01-02"]),
        }
    )
    ruta = tmp_path / "prueba.parquet"
    df.to_parquet(ruta, engine="pyarrow", index=False)
    leido = pd.read_parquet(ruta, engine="pyarrow")

    assert leido["SlotsComercialesValor"].tolist()[1] is pd.NA or pd.isna(leido["SlotsComercialesValor"].tolist()[1])
    assert leido["SlotsComercialesValor"].tolist()[0] == 20
    assert leido["CapacidadSlotsDesconocida"].tolist() == [False, True]
    assert pd.api.types.is_datetime64_any_dtype(leido["Fecha"])


def test_export_completo_escribe_5_parquet_y_manifest(tmp_path):
    resultado = ed.export_data(PRODUCTION_FILE, tmp_path)
    archivos = sorted(p.name for p in tmp_path.iterdir())
    assert archivos == [
        "_export_manifest.json",
        "bridge_campana_dia.parquet",
        "dim_calendario.parquet",
        "dim_elementos.parquet",
        "fact_campanas.parquet",
        "fact_metricas_diaria.parquet",
    ]
    assert resultado["manifest"]["source_sha256_match"] is True
    for nombre, df in resultado["tablas"].items():
        releido = pd.read_parquet(tmp_path / {
            "DIM_ELEMENTOS": "dim_elementos.parquet",
            "FACT_CAMPANAS": "fact_campanas.parquet",
            "DIM_CALENDARIO": "dim_calendario.parquet",
            "BRIDGE_CAMPANA_DIA": "bridge_campana_dia.parquet",
            "FACT_METRICAS_DIARIA": "fact_metricas_diaria.parquet",
        }[nombre], engine="pyarrow")
        assert len(releido) == len(df)


# ---------------------------------------------------------------------------
# SHA-256 del input intacto
# ---------------------------------------------------------------------------


def test_export_es_read_only_sobre_el_excel_fuente(tmp_path):
    sha_antes = vi.calculate_sha256(PRODUCTION_FILE)
    ed.export_data(PRODUCTION_FILE, tmp_path)
    sha_despues = vi.calculate_sha256(PRODUCTION_FILE)
    assert sha_antes == sha_despues
