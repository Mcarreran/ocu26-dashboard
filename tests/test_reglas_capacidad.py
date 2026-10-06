"""Reglas de negocio centrales OCU26 - Etapa 2A (2026-10-06).

Fuente unica de verdad: config/business_semantics.json (digital_capacity +
formato_negocio), resuelta por semantic_model (SlotsComerciales,
capacidad_espacio_digital, espacios_base_por_estacion) y por
metrics_engine.ocupacion_simultanea_por_estacion (YPF).

Todas las pruebas son hermeticas: filas sinteticas en memoria con la
CONFIGURACION REAL (sm.load_config()); ninguna lee ni escribe el Excel, los
HTML productivos ni output/.
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
import build_tv1_dashboard as tv1  # noqa: E402
import build_tv2_dashboard as tv2  # noqa: E402
import build_tv3_dashboard as tv3  # noqa: E402
import build_tv4_dashboard as tv4  # noqa: E402
import build_tv5_dashboard as tv5  # noqa: E402
import build_tv6_dashboard as tv6  # noqa: E402
from metrics_engine import (  # noqa: E402
    MetricsEngine,
    MetricsEngineError,
    max_campanas_simultaneas,
    ocupacion_simultanea_por_estacion,
    resumen_ocupacion_simultanea,
)
from test_semantic_model import _campana_row, _maestro_row, _transform_result  # noqa: E402


def _semantic(maestro_rows: list[dict], campanas_rows: list[dict] | None = None) -> dict:
    return sm.build_semantic_model(_transform_result(maestro_rows, campanas_rows or []), sm.load_config())


def _digital(elemento_id: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Shoppings Digital", Subcircuito="CENCOSUD", Ubicacion="UNICENTER",
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        Descripcion="", CapacidadSlotsReel=20, SegundosDia=100800,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _resuelto(maestro_row: dict) -> pd.Series:
    return _semantic([maestro_row])["maestro"].iloc[0]


def _capacidad(maestro_row: dict) -> tuple[float | None, str]:
    return sm.capacidad_espacio_digital(_resuelto(maestro_row))


# ---------------------------------------------------------------------------
# Fuente unica: config
# ---------------------------------------------------------------------------


def test_config_declara_las_reglas_vigentes():
    dc = sm.load_config()["digital_capacity"]
    assert dc["slots_profiles"] == {
        "PANTALLA_LED": 20, "TOTEM": 10, "TRIEDRO": 10, "PUENTE_LED": 10, "PATIO_COMIDAS": 10,
    }
    assert {k: v for k, v in dc["slots_por_circuito_digital"].items() if not k.startswith("_")} == {
        "AA2000": 10, "REMEROS": 10,
    }
    assert dc["capacidad_por_estacion"]["YPF"]["espacios_base_por_estacion"] == 5
    assert dc["default_segundos_comerciales"] == 72000
    assert sm.reglas_version().startswith("ETAPA2A")


def test_builders_no_definen_tasas_propias():
    """Ningun builder vuelve a definir una tabla de capacidad propia."""
    for n in range(1, 7):
        src = (REPO_ROOT / "scripts" / f"build_tv{n}_dashboard.py").read_text(encoding="utf-8")
        assert "ESPACIOS_POR_FORMATO_DIGITAL" not in src, n
        assert "CapacidadSlotsReel\"]" not in src.replace("row.get(\"CapacidadSlotsReel\")", ""), n


# ---------------------------------------------------------------------------
# 1-8. Capacidades por formato real
# ---------------------------------------------------------------------------


def test_01_pantalla_led_20():
    row = _digital("C99 - TEST", CircuitoDashboard="Pantalla Led", Subcircuito="PLED", CapacidadSlotsReel=40)
    assert _capacidad(row) == (20.0, "PANTALLA_LED")


def test_02_totem_10_aunque_el_excel_diga_20():
    assert _capacidad(_digital("T-1", Descripcion="Totem Digital - Test", CapacidadSlotsReel=20)) == (10.0, "TOTEM_SHOPPING")


def test_03_triedro_10():
    assert _capacidad(_digital("TR-1", Descripcion="Triedro Digital - Test", CapacidadSlotsReel=20)) == (10.0, "TRIEDRO")


def test_04_puente_led_10():
    assert _capacidad(_digital("PU-1", Descripcion="Puente Led 2 - Sector J", CapacidadSlotsReel=13)) == (10.0, "PUENTE_LED")


def test_05_aa2000_digital_10_tripstore_y_sin_descripcion():
    tripstore = _digital("AEP-TS-1", CircuitoDashboard="AA2000", Subcircuito="TRIPSTORE",
                         Descripcion="Totem Simple - Aeroparque", CapacidadSlotsReel=20)
    assert _capacidad(tripstore) == (10.0, "TRIPSTORE_AA2000")
    for eid in ("EZEPAW005", "EZEPAW011"):
        sin_desc = _digital(eid, CircuitoDashboard="AA2000", Subcircuito="EZEIZA", Descripcion=None, CapacidadSlotsReel=0)
        assert _capacidad(sin_desc) == (10.0, "OTRO_DIGITAL")


def test_06_patio_de_comidas_digital_10_y_estatico_no_cambia():
    digital = _digital("UNI-PACO-3L-1", Descripcion="Patio de Comidas - Nivel 3", CapacidadSlotsReel=10)
    assert _capacidad(digital) == (10.0, "PATIO_COMIDAS")
    estatico = _resuelto(_digital("UNI-EST-1", CircuitoDashboard="Shoppings Estático", Medio="Estático",
                                  TipoInventario="Físico estático", Descripcion="Patio de Comidas - Banner",
                                  CapacidadSlotsReel=0, SegundosDia=0))
    assert estatico["FormatoNegocio"] != "PATIO_COMIDAS"
    assert pd.isna(estatico["SlotsComerciales"])


def test_07_remeros_digital_10():
    rem = _digital("REM-DB-1", Subcircuito="REMEROS", Ubicacion="REMEROS", Descripcion="TV Led", CapacidadSlotsReel=20)
    assert _resuelto(rem)["FormatoNegocio"] == "TOTEM"
    assert _capacidad(rem) == (10.0, "TOTEM_SHOPPING")
    nuevo = _digital("REM-NUEVO-1", Subcircuito="REMEROS", Ubicacion="REMEROS", Descripcion="TV Led", CapacidadSlotsReel=20)
    assert _capacidad(nuevo)[0] == 10.0


def test_08_uni_puenteled_1_es_puente_led_10():
    row = _resuelto(_digital("UNI-PUENTELED-1", Descripcion=None, CapacidadSlotsReel=10))
    assert row["FormatoNegocio"] == "PUENTE_LED"
    assert sm.capacidad_espacio_digital(row) == (10.0, "PUENTE_LED")


def test_formato_desconocido_sin_regla_no_usa_capacidad_del_excel():
    """No se inventa capacidad ni se toma la capacidad historica del Excel."""
    row = _resuelto(_digital("UNI-X-1", Descripcion="Pantalla experimental", CapacidadSlotsReel=20))
    assert row["SlotsComerciales"] == "REQUIERE_CONFIRMACION"
    assert sm.capacidad_espacio_digital(row) == (None, "SIN_REGLA")


# ---------------------------------------------------------------------------
# 9. Segundos comerciales = 72.000
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("segundos_excel", [100800, 75000, 50400, 0])
def test_09_segundos_comerciales_72000_ignora_segundosdia_del_excel(segundos_excel):
    row = _resuelto(_digital("T-1", Descripcion="Totem Digital", SegundosDia=segundos_excel))
    assert row["SegundosComerciales"] == 72000
    assert row["SegundosDia"] == segundos_excel  # la columna original no se toca


# ---------------------------------------------------------------------------
# 10-12. YPF: catalogo 5 por estacion, ocupacion = maximo simultaneo SIN tope
# ---------------------------------------------------------------------------


def _asig(station: str, campaign: str, ini: str, fin: str) -> dict:
    return {"Estacion": station, "IDCampaña": campaign, "FechaInicio": pd.Timestamp(ini), "FechaFin": pd.Timestamp(fin)}


def test_10_ypf_catalogo_5_por_estacion_y_sin_capacidad_por_elemento():
    assert sm.espacios_base_por_estacion("YPF") == 5
    assert tv1.ESPACIOS_YPF_POR_ESTACION == 5
    assert tv4.DIGITAL_CAPACITY_PER_STATION == 5
    ypf = _resuelto(_digital("500 - TT - 1", CircuitoDashboard="YPF Digital", Subcircuito="500",
                             Ubicacion="500 - TESTVILLE - Calle 1", TipoCatalogo="Abierto", CapacidadSlotsReel=10))
    assert ypf["CircuitoNegocio"] == "YPF"
    assert pd.isna(ypf["SlotsComerciales"])  # capacidad por estacion, no por elemento


def test_11_ypf_siete_simultaneas_ocupan_siete():
    df = pd.DataFrame([_asig("E1", f"C{i}", "2026-07-05", "2026-07-10") for i in range(7)])
    r = ocupacion_simultanea_por_estacion(df, "2026-07-01", "2026-07-31", 5)
    fila = r.loc["E1"]
    assert fila["ocupacion"] == 7
    assert fila["exceso"] == 2
    assert fila["porcentaje_ocupacion"] == 140.0
    assert fila["estado"] == "SOBRECAPACIDAD"


def test_12_ypf_maximo_simultaneo_no_campanas_distintas_del_mes():
    """Ejemplo obligatorio: A y B 01-15/07, C y D 16-31/07 -> 2, NO 4."""
    df = pd.DataFrame([
        _asig("E1", "A", "2026-07-01", "2026-07-15"),
        _asig("E1", "B", "2026-07-01", "2026-07-15"),
        _asig("E1", "C", "2026-07-16", "2026-07-31"),
        _asig("E1", "D", "2026-07-16", "2026-07-31"),
    ])
    r = ocupacion_simultanea_por_estacion(df, "2026-07-01", "2026-07-31", 5)
    assert r.loc["E1", "max_simultaneas"] == 2
    assert r.loc["E1", "ocupacion"] == 2
    assert r.loc["E1", "campanas_periodo"] == 4  # informativo, nunca la ocupacion


def test_ypf_tres_y_tres_no_simultaneas_dan_tres():
    df = pd.DataFrame(
        [_asig("E1", f"Q1-{i}", "2026-07-01", "2026-07-15") for i in range(3)]
        + [_asig("E1", f"Q2-{i}", "2026-07-16", "2026-07-31") for i in range(3)]
    )
    assert ocupacion_simultanea_por_estacion(df, "2026-07-01", "2026-07-31", 5).loc["E1", "ocupacion"] == 3


def test_ypf_misma_campana_en_varios_elementos_cuenta_una_vez():
    df = pd.DataFrame([
        _asig("E1", "A", "2026-07-01", "2026-07-20"),
        _asig("E1", "A", "2026-07-05", "2026-07-25"),  # misma campaña, otro elemento
        _asig("E1", "B", "2026-07-10", "2026-07-12"),
    ])
    assert ocupacion_simultanea_por_estacion(df, "2026-07-01", "2026-07-31", 5).loc["E1", "ocupacion"] == 2


def test_ypf_misma_campana_con_hueco_no_rellena_el_hueco():
    """Una campaña con dos tramos separados no 'ocupa' el hueco intermedio."""
    df = pd.DataFrame([
        _asig("E1", "A", "2026-07-01", "2026-07-05"),
        _asig("E1", "A", "2026-07-20", "2026-07-25"),
        _asig("E1", "B", "2026-07-10", "2026-07-12"),
    ])
    assert ocupacion_simultanea_por_estacion(df, "2026-07-01", "2026-07-31", 5).loc["E1", "ocupacion"] == 1


def test_ypf_dias_inclusivos_y_recorte_al_periodo():
    assert max_campanas_simultaneas([
        (pd.Timestamp("2026-07-01"), pd.Timestamp("2026-07-10")),
        (pd.Timestamp("2026-07-10"), pd.Timestamp("2026-07-20")),
    ]) == 2  # comparten el 10/07
    df = pd.DataFrame([
        _asig("E1", "A", "2026-06-01", "2026-06-30"),  # fuera de julio
        _asig("E1", "B", "2026-06-20", "2026-07-02"),
        _asig("E1", "C", "2026-07-01", "2026-08-15"),
    ])
    r = ocupacion_simultanea_por_estacion(df, "2026-07-01", "2026-07-31", 5)
    assert r.loc["E1", "ocupacion"] == 2
    assert r.loc["E1", "campanas_periodo"] == 2


def test_ypf_entradas_invalidas():
    assert ocupacion_simultanea_por_estacion(pd.DataFrame(columns=["Estacion", "IDCampaña", "FechaInicio", "FechaFin"]),
                                            "2026-07-01", "2026-07-31", 5).empty
    with pytest.raises(MetricsEngineError):
        ocupacion_simultanea_por_estacion(pd.DataFrame(), "2026-07-01", "2026-07-31", 0)
    with pytest.raises(MetricsEngineError):
        ocupacion_simultanea_por_estacion(pd.DataFrame(), "2026-07-31", "2026-07-01", 5)
    resumen = resumen_ocupacion_simultanea(None)
    assert resumen["ocupados"] == 0 and resumen["max_pct_estacion"] is None


def test_ypf_tv1_tv4_tv5_misma_logica_para_la_misma_estacion():
    """Misma estacion, mismas campañas, mismo periodo -> misma ocupacion en
    TV1 (StationKey) y TV4 (APIE); TV5 usa la misma funcion en el dia de
    corte. Fixture: 7 campañas simultaneas 05-10/07 + 1 sola el 20/07."""
    ubic = "900 - TESTVILLE - Calle 9"
    maestro = [
        _maestro_row(f"900 - TT - {i}", CircuitoDashboard="YPF Digital", Subcircuito="900", Ubicacion=ubic,
                     Medio="Digital", TipoCatalogo="Abierto", TipoInventario="Digital",
                     CapacidadSlotsReel=10, SegundosDia=100800, RevisionMaestro="BASE LIMPIA TEST")
        for i in range(1, 9)
    ]
    campanas = [
        _campana_row(f"K{i}", f"900 - TT - {i}", IDCampaña=f"CAMP{i}",
                     FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10"))
        for i in range(1, 8)
    ] + [_campana_row("K8", "900 - TT - 8", IDCampaña="CAMP8",
                      FechaInicio=pd.Timestamp("2026-07-20"), FechaFin=pd.Timestamp("2026-07-20"))]
    semantic_result = _semantic(maestro, campanas)
    engine = MetricsEngine(semantic_result)

    u1 = tv1.build_tv1_universe(semantic_result)
    ypf_dig = u1["maestro"][u1["maestro"]["CircuitoNegocio"] == "YPF"].copy()
    ypf_dig["StationKey"] = ypf_dig["ElementoID"].map(u1["ypf_station_map"])
    r1 = tv1._ypf_digital_ocupados_por_estacion(engine, ypf_dig, "2026-07-01", "2026-07-31")

    u4 = tv4.build_tv4_universe(semantic_result)
    r4 = tv4._ocupacion_periodo(engine, u4, "2026-07-01", "2026-07-31")

    assert r1["ocupados"] == r4["ocupados_digitales"] == 7
    assert r1["estaciones_sobrecapacidad"] == r4["estaciones_sobre_capacidad"] == 1
    assert r1["campanas_excedentes_total"] == r4["exceso_sobre_capacidad"] == 2

    snap = tv5.build_tv5_ypf_snapshot(semantic_result, pd.Timestamp("2026-07-07"))
    assert snap["espacios"] == 7  # en el dia de corte, 7 simultaneas


# ---------------------------------------------------------------------------
# 13-14. Exclusiones: Cencomedia y London fuera de las TVs
# ---------------------------------------------------------------------------


def _universo_con_excluidos() -> dict:
    rows = [
        _digital("C1 - TEST", CircuitoDashboard="Pantalla Led", Subcircuito="PLED"),
        _maestro_row("UNI-EST-1", CircuitoDashboard="Shoppings Estático", Subcircuito="CENCOSUD", Ubicacion="UNICENTER",
                     Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático",
                     Descripcion="Banner", CapacidadSlotsReel=0, SegundosDia=0),
        _maestro_row("CENCO-1", CircuitoDashboard="Jumbo Palermo", Subcircuito="JUMBO", Ubicacion="PALERMO",
                     Medio="Estático", TipoCatalogo="Abierto", TipoInventario="Flexible gráfico",
                     CapacidadSlotsReel=0, SegundosDia=0),
        _maestro_row("LON-1", CircuitoDashboard="London Supply", Subcircuito="FTE", Ubicacion="USH",
                     Medio="Digital", TipoCatalogo="Abierto", TipoInventario="Digital",
                     Descripcion="Pantallas LED de 75 pulgadas", CapacidadSlotsReel=20, SegundosDia=100800),
        _maestro_row("LON-2", CircuitoDashboard="London Supply", Subcircuito="FTE", Ubicacion="USH",
                     Medio="Estático", TipoCatalogo="Abierto", TipoInventario="Físico estático",
                     CapacidadSlotsReel=0, SegundosDia=0),
    ]
    return _semantic(rows)


def test_13_cencomedia_excluido_de_las_tvs():
    sem = _universo_con_excluidos()
    m = sem["maestro"].set_index("ElementoID")
    assert m.loc["CENCO-1", "CircuitoNegocio"] == "CENCOMEDIA"
    for ids in (
        tv1.build_tv1_universe(sem)["element_ids"],
        tv2.build_tv2_universe(sem)["element_ids"],
        tv3.build_tv3_universe(sem)["element_ids"],
        tv6.build_tv6_universe(sem)["element_ids"],
    ):
        assert "CENCO-1" not in ids
    assert "CENCOMEDIA" in tv1.TV1_EXCLUDED_CIRCUITOS


def test_14_london_excluido_de_los_tableros():
    sem = _universo_con_excluidos()
    m = sem["maestro"].set_index("ElementoID")
    assert m.loc["LON-1", "CircuitoNegocio"] == "LONDON_SUPPLY"
    for ids in (
        tv1.build_tv1_universe(sem)["element_ids"],
        tv2.build_tv2_universe(sem)["element_ids"],
        tv3.build_tv3_universe(sem)["element_ids"],
        tv6.build_tv6_universe(sem)["element_ids"],
    ):
        assert "LON-1" not in ids and "LON-2" not in ids
    u5 = tv5.build_tv5_universe(sem)
    assert "LON-1" not in u5["digital_ids"] and "LON-2" not in u5["static_ids"]
