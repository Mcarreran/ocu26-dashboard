"""Pruebas de negocio para scripts/build_tv4_dashboard.py (dashboard TV4
OCU26, Pulso comercial YPF).

TV4 reemplaza el scope anterior de este slot (Pipeline Comercial): YPF pasa
a reportarse aqui, no en TV5 (decision explicita del usuario, ver docstring
de build_tv4_dashboard.py). Esta suite NO modifica scripts/validate_input.py,
scripts/transform_data.py, scripts/semantic_model.py, scripts/metrics_engine.py,
config/business_semantics.json, input/OCU26_BASE_DATOS.xlsx, ni ningun
archivo productivo de TV1/TV2/TV3/TV5/TV6 (build_tv1/2/3/5/6_dashboard.py,
tv1/2/3/5/6_template.html, tv1/2/3/5/6.html, test_build_tv1/2/3/5/6_dashboard.py,
audit_sources/*). Los fixtures sinteticos usan la CONFIGURACION REAL
(sm.load_config()), mismo patron que test_build_tv5_dashboard.py.
"""

from __future__ import annotations

import hashlib
import json
import re
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
from metrics_engine import MetricsEngine  # noqa: E402
from test_semantic_model import _maestro_row, _campana_row, _transform_result  # noqa: E402

PRODUCTION_FILE = REPO_ROOT / "input" / "OCU26_BASE_DATOS.xlsx"
BUILDER_SOURCE = (REPO_ROOT / "scripts" / "build_tv4_dashboard.py").read_text(encoding="utf-8")

PERIOD = ("2026-07-01", "2026-07-31")
PREVIOUS = ("2026-06-01", "2026-06-30")

# ---------------------------------------------------------------------------
# Proteccion: TV1/TV2/TV3/TV5/TV6 deben permanecer byte-identicas. Hashes
# tomados ANTES de tocar ningun archivo de este cambio (protocolo Sec.1 del
# pedido). Si esta prueba falla, algo fuera del scope de TV4 se modifico.
# ---------------------------------------------------------------------------

_PROTECTED_SHA256 = {
    # Etapa 2A (2026-10-06): se retiraron de este guard los builders TV1/TV2/
    # TV3/TV6 y los tests TV1/TV2, modificados A PROPOSITO por la
    # centralizacion de reglas (capacidad en config/semantic_model, YPF por
    # simultaneidad, input configurable). El guard sigue protegiendo lo que
    # esta etapa NO debe tocar: HTML productivos generados y templates.
    "tv1.html": "1727aead4de580b455a8b346ad6428b736c2e58ec8364fe663944a1a988c3d98",
    "tv3.html": "71578ae7ef4b739454084fc47776ccaa9e6ebde15425671e494a5d58a492fe17",
    "scripts/templates/tv1_template.html": "71380f34d9618384728ab8da1882c8b8c05e69b52ebb76563c0bd0bc175076ba",
    "scripts/templates/tv2_template.html": "e501bfbf44699934c31a7e2e39d69cb00f2124c0485cd01f6fb4e7587c415782",
    "scripts/templates/tv3_template.html": "776dea93829887042281ed27089188e6ac5c6ee6baccd854d9084ce092381cd9",
    "scripts/templates/tv6_template.html": "9ccad1be73344736f4df81f47c73ece6ad464e8bd9f6ece6b3b4904a994e5a25",
    "tests/test_build_tv3_dashboard.py": "9f68b4c356806a372a131b08e867ccd6c6c01fd4a0cb97bc5cd3d0f67c87866b",
    "tests/test_build_tv6_dashboard.py": "d58947dd3752a5a393fbf648b802d4731e170ccbcbc485af40e491120aeb288f",
    # Baseline TV6 actualizado al cierre aceptado TV1-TV6, commit c4fc8d8
    # (2026-08-23): reemplaza el hash "pre-TV4" de estos 3 archivos por el
    # vigente tras la migracion de TV6 a marcas/demanda por circuito.
    # scripts/build_tv6_dashboard.py y tests/test_build_tv6_dashboard.py
    # actualizados nuevamente el 24/08/2026: parche determinista (sorted()
    # en primer_mes_por_marca/circuitos_por_marca, ver compute_recurrencia y
    # compute_circuitos_por_marca) + fixture hermetico (build_and_write ya
    # no escribe tv6.html/output/tv6_data.json reales en los tests).
    # Baseline tests/test_build_tv2_dashboard.py actualizado el 24/08/2026:
    # refleja el parche autorizado de hermeticidad TV2 (fixture production_html
    # ahora escribe en tmp_path_factory en vez de sobreescribir tv2.html real).
    # tests/test_build_tv5_dashboard.py no forma parte de este guard (ver nota
    # de exclusion TV5 mas abajo); el parche de Selenium opt-in TV5 del
    # 24/08/2026 no requiere actualizar ningun hash aqui.
    # tv2.html / tv6.html: excluidos de este chequeo de hash exacto (2026-08-23).
    # Sus scripts/templates/tests fuente SI siguen protegidos arriba (byte-
    # identicos). Solo el .html generado cambia, y solo en el timestamp de
    # build ("generado"/"generado_iso"): correr la suite completa disparo un
    # rebuild propio de esas dos TVs como efecto colateral (confirmado con el
    # usuario, no revertido por no tener guardado el timestamp original exacto).
    # tv5.html / build_tv5_dashboard.py / tv5_template.html / test_build_tv5_
    # dashboard.py: excluidos de este chequeo (2026-08-23). TV5 fue reasignada
    # de "YPF" a "Pipeline Comercial" por una via externa a esta sesion
    # (confirmado explicitamente con el usuario) -- no es un cambio de TV4.
}


@pytest.mark.parametrize("rel_path", sorted(_PROTECTED_SHA256))
def test_protected_tv_files_remain_byte_identical(rel_path):
    path = REPO_ROOT / rel_path
    # El baseline historico de _PROTECTED_SHA256 mezcla finales de linea: los
    # fuentes .py/_template.html se hashearon en LF y los HTML de salida
    # (tv1.html, tv3.html) en CRLF nativo de Windows. Un checkout de Git en
    # otro sistema operativo (o con otro autocrlf) puede materializar
    # cualquiera de los dos sin que el contenido logico cambie, asi que el
    # guard acepta el hash crudo, su version LF o su version CRLF. Cualquier
    # cambio de CONTENIDO real (no solo de fin de linea) sigue sin coincidir
    # con ninguno de los tres y el test sigue fallando como corresponde.
    raw = path.read_bytes()
    lf = raw.replace(b"\r\n", b"\n")
    crlf = lf.replace(b"\n", b"\r\n")
    candidate_hashes = {
        hashlib.sha256(raw).hexdigest(),
        hashlib.sha256(lf).hexdigest(),
        hashlib.sha256(crlf).hexdigest(),
    }
    expected = _PROTECTED_SHA256[rel_path]
    assert expected in candidate_hashes, f"{rel_path} cambio respecto al baseline pre-TV4"


# ---------------------------------------------------------------------------
# Fixtures sinteticas
# ---------------------------------------------------------------------------


def _semantic(maestro_rows: list[dict], campanas_rows: list[dict] | None = None) -> dict:
    config = sm.load_config()
    return sm.build_semantic_model(_transform_result(maestro_rows, campanas_rows or []), config)


# ---------------------------------------------------------------------------
# Helpers de compatibilidad: tests/test_build_tv6_dashboard.py (protegido,
# no se modifica) importa estos 6 builders de fila desde este modulo
# (`from test_build_tv4_dashboard import _cencosud_static, ...`). Se
# conservan aqui, sin cambios, exclusivamente para no romper esa
# importacion; no se usan en las pruebas de TV4 de mas abajo (que usan
# _ypf_row/_digital/_static, especificos del universo YPF con APIE).
# ---------------------------------------------------------------------------


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


def _ypf_row(elemento_id: str, apie: str, medio: str = "Digital", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="YPF Digital" if medio == "Digital" else "YPF Estático",
        Subcircuito=str(apie),
        Ubicacion=f"{apie} - TESTVILLE - Calle Falsa 123",
        Medio=medio,
        TipoCatalogo="Abierto",
        TipoInventario=medio,
        CapacidadSlotsReel=0,
        SegundosDia=0,
        RevisionMaestro="BASE LIMPIA YPF ETAPA 1 CORREGIDA - TEST",
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _digital(apie: str, n: int, token: str = "TT", **overrides) -> dict:
    return _ypf_row(f"{apie} - {token} - {n}", apie, medio="Digital", **overrides)


def _static(apie: str, n: int, **overrides) -> dict:
    return _ypf_row(f"{apie} - FB - {n}", apie, medio="Estático", **overrides)


def _camp(carga_id: str, elemento_id: str, id_campana: str, start: str, end: str, **overrides) -> dict:
    return _campana_row(carga_id, elemento_id, IDCampaña=id_campana, FechaInicio=pd.Timestamp(start), FechaFin=pd.Timestamp(end), **overrides)


def _build(maestro_rows, campana_rows=None):
    semantic_result = _semantic(maestro_rows, campana_rows or [])
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv4_universe(semantic_result)
    return semantic_result, engine, universe


@pytest.fixture(scope="module")
def production_result():
    return td.build_tv4_data(PRODUCTION_FILE)


@pytest.fixture(scope="module")
def production_json(production_result):
    return json.dumps(production_result["data"], ensure_ascii=False)


@pytest.fixture(scope="module")
def production_html(production_result):
    return td.render_html(production_result["data"])


# ---------------------------------------------------------------------------
# 1. Estaciones contadas mediante APIE distintos
# ---------------------------------------------------------------------------


def test_apie_distinto_deduplica_estaciones():
    rows = [_digital("500", 1, "TT"), _digital("500", 2, "TT"), _digital("500", 3, "MB")]
    _, _, universe = _build(rows)
    assert universe["estaciones_catalogo"] == 1
    assert universe["elementos_catalogo"] == 3


def test_apie_distintos_no_se_colapsan():
    rows = [_digital("500", 1, "TT"), _digital("600", 1, "TT")]
    _, _, universe = _build(rows)
    assert universe["estaciones_catalogo"] == 2


# ---------------------------------------------------------------------------
# 2. Estacion digital con muchos elementos aporta 5 espacios (no mas)
# ---------------------------------------------------------------------------


def test_estacion_digital_muchos_elementos_da_cinco_espacios():
    rows = [_digital("500", i, "TT") for i in range(1, 9)]  # 8 elementos digitales
    _, _, universe = _build(rows)
    assert universe["espacios_digitales_catalogo"] == 5
    assert universe["espacios_totales_catalogo"] == 5
    assert universe["elementos_catalogo"] == 8  # los elementos fisicos no se pierden


# ---------------------------------------------------------------------------
# 3. Estacion digital con dos FB aporta 7 espacios (5 + 2)
# ---------------------------------------------------------------------------


def test_estacion_digital_con_dos_fb_da_siete_espacios():
    rows = [_digital("500", 1, "TT"), _digital("500", 2, "MB"), _static("500", 1), _static("500", 2)]
    _, _, universe = _build(rows)
    assert universe["espacios_digitales_catalogo"] == 5
    assert universe["espacios_estaticos_catalogo"] == 2
    assert universe["espacios_totales_catalogo"] == 7
    assert universe["estaciones_catalogo"] == 1


# ---------------------------------------------------------------------------
# 4. Estacion solo estatica con dos FB aporta 2 espacios (sin capacidad digital)
# ---------------------------------------------------------------------------


def test_estacion_estatica_con_dos_fb_da_dos_espacios():
    rows = [_static("500", 1), _static("500", 2)]
    _, _, universe = _build(rows)
    assert universe["espacios_digitales_catalogo"] == 0
    assert universe["espacios_estaticos_catalogo"] == 2
    assert universe["espacios_totales_catalogo"] == 2
    assert universe["estaciones_con_digital"] == 0


# ---------------------------------------------------------------------------
# 5. APIE 30943 (control real de produccion): 2 elementos FB, 2 espacios,
# sin capacidad digital.
# ---------------------------------------------------------------------------


def test_produccion_apie_30943_dos_espacios_sin_capacidad_digital(production_result):
    # Reconstruimos el universo real para inspeccionar el APIE puntual.
    _, semantic_result, _ = __import__("export_data").load_pipeline(PRODUCTION_FILE)
    universe = td.build_tv4_universe(semantic_result)
    maestro = universe["maestro"]
    sub = maestro[maestro["_APIE"] == "30943"]
    assert len(sub) == 2
    assert set(sub["_token"]) == {"FB"}
    assert "30943" not in universe["apie_con_digital"]
    assert "30943" in universe["apie_con_estatico"]


def test_apie_30943_shape_replicada_sinteticamente():
    rows = [_static("30943", 1), _static("30943", 2)]
    _, _, universe = _build(rows)
    assert universe["estaciones_catalogo"] == 1
    assert universe["elementos_catalogo"] == 2
    assert universe["espacios_totales_catalogo"] == 2
    assert universe["espacios_digitales_catalogo"] == 0


# ---------------------------------------------------------------------------
# 6. Misma campana en varios elementos digitales de una estacion = 1 espacio
# ---------------------------------------------------------------------------


def test_misma_campana_en_mb_tt_punter_ocupa_un_solo_espacio_digital():
    rows = [_digital("500", 1, "MB"), _digital("500", 1, "TT"), _digital("500", 1, "PPUNTER")]
    campanas = [
        _camp("C1", "500 - MB - 1", "CAMP1", "2026-07-05", "2026-07-10"),
        _camp("C2", "500 - TT - 1", "CAMP1", "2026-07-05", "2026-07-10"),
        _camp("C3", "500 - PPUNTER - 1", "CAMP1", "2026-07-05", "2026-07-10"),
    ]
    _, engine, universe = _build(rows, campanas)
    ocupacion = td._ocupacion_periodo(engine, universe, *PERIOD)
    assert ocupacion["ocupados_digitales"] == 1


# ---------------------------------------------------------------------------
# 7-8. Seis campanas digitales en una estacion -> 6/5 = 120%, exceso registrado
# ---------------------------------------------------------------------------


def test_seis_campanas_digitales_producen_120_pct_y_exceso():
    rows = [_digital("500", i, "TT") for i in range(1, 7)]  # 6 elementos, 1 estacion
    campanas = [
        _camp(f"C{i}", f"500 - TT - {i}", f"CAMP{i}", "2026-07-05", "2026-07-10") for i in range(1, 7)
    ]
    _, engine, universe = _build(rows, campanas)
    ocupacion = td._ocupacion_periodo(engine, universe, *PERIOD)
    assert ocupacion["ocupados_digitales"] == 6
    assert universe["espacios_digitales_catalogo"] == 5
    assert round(ocupacion["ocupados_digitales"] / universe["espacios_digitales_catalogo"] * 100.0, 1) == 120.0
    assert ocupacion["estaciones_sobre_capacidad"] == 1
    assert ocupacion["exceso_sobre_capacidad"] == 1
    assert ocupacion["max_pct_estacion"] == 120.0


def test_campanas_no_simultaneas_no_se_suman_en_la_estacion():
    """Etapa 2A (regla definitiva YPF): 3 campañas en la primera quincena y
    otras 3, NO simultaneas, en la segunda -> ocupacion 3, nunca 6 (no se
    usa 'campañas distintas del mes')."""
    rows = [_digital("500", i, "TT") for i in range(1, 7)]
    campanas = [
        _camp(f"C{i}", f"500 - TT - {i}", f"Q1-{i}", "2026-07-01", "2026-07-15") for i in range(1, 4)
    ] + [
        _camp(f"C{i}", f"500 - TT - {i}", f"Q2-{i}", "2026-07-16", "2026-07-31") for i in range(4, 7)
    ]
    _, engine, universe = _build(rows, campanas)
    ocupacion = td._ocupacion_periodo(engine, universe, *PERIOD)
    assert ocupacion["ocupados_digitales"] == 3
    assert ocupacion["estaciones_sobre_capacidad"] == 0


def test_siete_campanas_simultaneas_ocupan_siete_sin_tope():
    rows = [_digital("500", i, "TT") for i in range(1, 8)]
    campanas = [_camp(f"C{i}", f"500 - TT - {i}", f"CAMP{i}", "2026-07-05", "2026-07-10") for i in range(1, 8)]
    _, engine, universe = _build(rows, campanas)
    ocupacion = td._ocupacion_periodo(engine, universe, *PERIOD)
    assert ocupacion["ocupados_digitales"] == 7
    assert ocupacion["exceso_sobre_capacidad"] == 2
    assert ocupacion["max_pct_estacion"] == 140.0


def test_produccion_reporta_estaciones_sobre_capacidad_y_exceso(production_result):
    oc = production_result["data"]["kpis"]["ocupacion"]["actual"]
    assert oc["estaciones_sobre_capacidad"] >= 0
    assert oc["exceso_sobre_capacidad"] >= 0
    if oc["estaciones_sobre_capacidad"] > 0:
        assert oc["exceso_sobre_capacidad"] > 0
        assert oc["max_pct_estacion"] > 100.0


# ---------------------------------------------------------------------------
# 9. Estatico ocupado se cuenta por ElementoID distinto (no por fila/campana)
# ---------------------------------------------------------------------------


def test_fb_con_varias_campanas_sigue_siendo_un_solo_espacio_estatico_ocupado():
    rows = [_static("500", 1)]
    campanas = [
        _camp("C1", "500 - FB - 1", "CAMPA", "2026-07-01", "2026-07-10"),
        _camp("C2", "500 - FB - 1", "CAMPB", "2026-07-15", "2026-07-20"),
    ]
    _, engine, universe = _build(rows, campanas)
    ocupacion = td._ocupacion_periodo(engine, universe, *PERIOD)
    assert ocupacion["ocupados_estaticos"] == 1


# ---------------------------------------------------------------------------
# 10-11. Campanas por IDCampana distinto; activaciones por par distinto
# ---------------------------------------------------------------------------


def test_campanas_unicas_por_idcampana_distinto():
    rows = [_digital("1", 1, "TT"), _digital("2", 1, "TT"), _digital("3", 1, "TT")]
    campanas = [
        _camp("C1", "1 - TT - 1", "SAME", "2026-07-05", "2026-07-10"),
        _camp("C2", "2 - TT - 1", "SAME", "2026-07-05", "2026-07-10"),
        _camp("C3", "3 - TT - 1", "SAME", "2026-07-05", "2026-07-10"),
    ]
    _, engine, universe = _build(rows, campanas)
    resultado = td.compute_campanas(engine, universe, PERIOD, PREVIOUS)
    assert resultado["campanas_unicas_actual"] == 1
    assert resultado["activaciones_actual"] == 3


def test_activaciones_por_par_idcampana_elementoid():
    rows = [_digital("500", 1, "TT")]
    campanas = [
        _camp("C1", "500 - TT - 1", "CA", "2026-07-01", "2026-07-10"),
        _camp("C2", "500 - TT - 1", "CB", "2026-07-15", "2026-07-20"),
    ]
    _, engine, universe = _build(rows, campanas)
    resultado = td.compute_campanas(engine, universe, PERIOD, PREVIOUS)
    assert resultado["campanas_unicas_actual"] == 2
    assert resultado["activaciones_actual"] == 2


# ---------------------------------------------------------------------------
# 12. Filas duplicadas no inflan resultados
# ---------------------------------------------------------------------------


def test_filas_duplicadas_de_campana_no_inflan_activaciones():
    rows = [_digital("500", 1, "TT")]
    campanas = [
        _camp("C1", "500 - TT - 1", "CA", "2026-07-01", "2026-07-10"),
        _camp("C1-DUP", "500 - TT - 1", "CA", "2026-07-01", "2026-07-10"),
    ]
    _, engine, universe = _build(rows, campanas)
    resultado = td.compute_campanas(engine, universe, PERIOD, PREVIOUS)
    assert resultado["campanas_unicas_actual"] == 1
    assert resultado["activaciones_actual"] == 1
    ocupacion = td._ocupacion_periodo(engine, universe, *PERIOD)
    assert ocupacion["ocupados_digitales"] == 1


# ---------------------------------------------------------------------------
# 13-14. Ocupados totales = digital + estatico; espacios totales = digital + estatico
# ---------------------------------------------------------------------------


def test_ocupados_totales_es_suma_digital_mas_estatico():
    rows = [_digital("500", 1, "TT"), _static("500", 1)]
    campanas = [
        _camp("C1", "500 - TT - 1", "CD", "2026-07-01", "2026-07-10"),
        _camp("C2", "500 - FB - 1", "CE", "2026-07-01", "2026-07-10"),
    ]
    _, engine, universe = _build(rows, campanas)
    ocupacion = td._ocupacion_periodo(engine, universe, *PERIOD)
    assert ocupacion["ocupados_totales"] == ocupacion["ocupados_digitales"] + ocupacion["ocupados_estaticos"]
    assert ocupacion["ocupados_totales"] == 2


def test_espacios_totales_es_suma_digital_mas_estatico(production_result):
    cat = production_result["data"]["kpis"]["catalogo"]
    assert cat["espacios"] == cat["espacios_digitales"] + cat["espacios_estaticos"]


# ---------------------------------------------------------------------------
# 15. Comparacion julio contra junio (deltas correctos, periodos independientes)
# ---------------------------------------------------------------------------


def test_junio_se_calcula_independiente_de_julio():
    rows = [_digital("500", 1, "TT")]
    campanas = [_camp("C1", "500 - TT - 1", "CJUN", "2026-06-05", "2026-06-10")]
    _, engine, universe = _build(rows, campanas)
    ocupacion = td.compute_ocupacion(universe, engine, PERIOD, PREVIOUS)
    assert ocupacion["actual"]["ocupados_totales"] == 0
    assert ocupacion["anterior"]["ocupados_totales"] == 1
    assert ocupacion["delta_abs"] == -1


def test_produccion_delta_abs_es_actual_menos_anterior(production_result):
    oc = production_result["data"]["kpis"]["ocupacion"]
    assert oc["delta_abs"] == oc["actual"]["ocupados_totales"] - oc["anterior"]["ocupados_totales"]
    camp = production_result["data"]["kpis"]["campanas"]
    assert camp["delta_campanas"] == camp["campanas_unicas_actual"] - camp["campanas_unicas_anterior"]


# ---------------------------------------------------------------------------
# 16-17. Series mensuales con denominadores correctos, sin meses futuros
# ---------------------------------------------------------------------------


def test_historico_no_tiene_meses_futuros(production_result):
    hist = production_result["data"]["historico"]
    assert hist["meses"] == ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul"]
    for serie in hist["series"].values():
        assert len(serie) == 7


def test_historico_denominadores_correctos(production_result):
    hist = production_result["data"]["historico"]
    cat = production_result["data"]["kpis"]["catalogo"]
    for punto in hist["series"]["Digital"]:
        assert punto["capacidad"] == cat["espacios_digitales"]
    for punto in hist["series"]["Estático"]:
        assert punto["capacidad"] == cat["espacios_estaticos"]
    for punto in hist["series"]["Total"]:
        assert punto["capacidad"] == cat["espacios"]
    # pct = ocupados / capacidad * 100, coherente en cada punto
    for serie in hist["series"].values():
        for punto in serie:
            if punto["capacidad"]:
                esperado = round(punto["ocupados"] / punto["capacidad"] * 100.0, 1)
                assert punto["pct"] == esperado


def test_historico_no_capa_en_100_si_hay_sobreocupacion():
    rows = [_digital("500", i, "TT") for i in range(1, 7)]
    campanas = [_camp(f"C{i}", f"500 - TT - {i}", f"CAMP{i}", "2026-07-05", "2026-07-10") for i in range(1, 7)]
    _, engine, universe = _build(rows, campanas)
    hist = td.compute_historico(engine, universe, 2026, 7)
    julio = hist["series"]["Digital"][-1]
    assert julio["pct"] > 100.0


# ---------------------------------------------------------------------------
# 18. Composicion de elementos activos reconcilia 100%
# ---------------------------------------------------------------------------


def test_composicion_elementos_activos_reconcilia_100(production_result):
    comp = production_result["data"]["kpis"]["elementos_activos"]["composicion_actual"]
    total_pct = sum(v["pct"] for v in comp.values() if v["pct"] is not None)
    assert abs(total_pct - 100.0) < 0.5
    total_elementos = sum(v["elementos"] for v in comp.values())
    assert total_elementos == production_result["data"]["kpis"]["elementos_activos"]["actual"]


# ---------------------------------------------------------------------------
# 19. Mapa solo usa coordenadas validas
# ---------------------------------------------------------------------------


def test_mapa_solo_coordenadas_validas(production_result):
    mapa = production_result["data"]["mapa"]
    for punto in mapa["puntos"]:
        assert -90.0 <= punto["lat"] <= 90.0
        assert -180.0 <= punto["lon"] <= 180.0
        assert not (punto["lat"] == 0 and punto["lon"] == 0)


# ---------------------------------------------------------------------------
# 20. No aparecen numeros legacy (551, 4.048) como resultados hardcodeados
# ---------------------------------------------------------------------------


def test_no_hardcodea_numeros_legacy_en_el_builder():
    """551/4.048 solo pueden aparecer en el docstring del modulo (documentando
    que NO deben usarse); nunca en el codigo ejecutable (asignaciones,
    literales de diccionario, comparaciones)."""
    marker = '"""'
    first = BUILDER_SOURCE.find(marker)
    second = BUILDER_SOURCE.find(marker, first + 3)
    code_only = BUILDER_SOURCE[second + 3:] if second != -1 else BUILDER_SOURCE
    assert "551" not in code_only
    assert "4048" not in code_only
    assert "4.048" not in code_only


def test_catalogo_produccion_no_coincide_con_numeros_legacy(production_result):
    cat = production_result["data"]["kpis"]["catalogo"]
    assert cat["estaciones"] != 551
    assert cat["elementos"] != 4048


# ---------------------------------------------------------------------------
# 21. TV1/TV2/TV3/TV5/TV6 permanecen byte-identicas: ver
# test_protected_tv_files_remain_byte_identical arriba.
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# 22. HTML 1920x1080 sin scroll ni overflow
# ---------------------------------------------------------------------------


def test_template_declara_stage_1920x1080_sin_overflow():
    template_src = (REPO_ROOT / "scripts" / "templates" / "tv4_template.html").read_text(encoding="utf-8")
    assert "width:1920px" in template_src.replace(" ", "")
    assert "height:1080px" in template_src.replace(" ", "")
    assert "overflow:hidden" in template_src.replace(" ", "")


def test_rendered_html_contains_stage_and_no_scroll_markers(production_html):
    compact = production_html.replace(" ", "")
    assert "width:1920px" in compact
    assert "height:1080px" in compact
    assert "overflow:hidden" in compact


# ---------------------------------------------------------------------------
# Ajuste puntual (2026-08-23): grafico solo Digital/Estatico, tarjetas
# Ocupacion/Campanas rediseñadas, mapa con cobertura provincial.
# ---------------------------------------------------------------------------


_TEMPLATE_SRC = (REPO_ROOT / "scripts" / "templates" / "tv4_template.html").read_text(encoding="utf-8")


def _extract_js_function(src: str, name: str) -> str:
    marker = f"function {name}("
    start = src.index(marker)
    depth = 0
    i = src.index("{", start)
    body_start = i
    while True:
        if src[i] == "{":
            depth += 1
        elif src[i] == "}":
            depth -= 1
            if depth == 0:
                return src[body_start : i + 1]
        i += 1


def test_1_chart_order_contains_only_digital_y_estatico():
    fn = _extract_js_function(_TEMPLATE_SRC, "drawHistorico")
    m = re.search(r"var order = \[(.*?)\];", fn)
    assert m is not None
    assert m.group(1).replace(" ", "") == "'Digital','Estático'"


def test_2_no_aparece_total_en_la_leyenda_del_grafico():
    fn = _extract_js_function(_TEMPLATE_SRC, "drawHistorico")
    assert "Total" not in fn


def test_2b_payload_conserva_total_para_tarjeta_e_insights(production_result):
    # El calculo Total sigue viajando en el payload (Sec.1: "mantener el
    # calculo total para la tarjeta de ocupacion, el payload y los
    # insights"); solo se dejo de DIBUJAR como tercera linea.
    hist = production_result["data"]["historico"]
    assert "Total" in hist["series"]
    assert len(hist["series"]["Total"]) == 7


def test_3_numero_absoluto_de_ocupados_precede_al_porcentaje():
    render_fn = _extract_js_function(_TEMPLATE_SRC, "render")
    idx_abs = render_fn.index("Espacios ocupados")
    idx_pct = render_fn.index("% de ocupación")
    assert idx_abs < idx_pct, "el bloque 'Espacios ocupados' debe generarse (y por lo tanto quedar a la izquierda) antes que '% de ocupación'"
    idx_abs_val = render_fn.index("hero-value abs")
    idx_pct_val = render_fn.index("hero-value pct")
    assert idx_abs_val < idx_pct_val


def test_4_tarjeta_campanas_conserva_cifras_y_layout_compacto(production_result):
    camp = production_result["data"]["kpis"]["campanas"]
    render_fn = _extract_js_function(_TEMPLATE_SRC, "render")
    assert "kpi-cols" in render_fn
    assert "Activaciones" in render_fn
    assert "Únicas" in render_fn
    # Los valores en si no se tocan (Sec.3: "No modificar los valores de
    # campañas ni activaciones") -- verificado contra produccion real.
    assert camp["campanas_unicas_actual"] == 17
    assert camp["activaciones_actual"] == 10945


def test_4b_grid_kpis_es_proporcional_no_uniforme():
    m = re.search(r"\.kpis\{[^}]*grid-template-columns:([^;]+);", _TEMPLATE_SRC)
    assert m is not None
    cols = m.group(1).strip()
    assert cols != "repeat(5,1fr)"
    assert len(cols.split()) == 5


def test_chart_colors_digital_amarillo_estatico_celeste():
    fn = _extract_js_function(_TEMPLATE_SRC, "drawHistorico")
    m = re.search(r"var colors = \{ 'Digital': ([^,]+), 'Estático': ([^}]+) \};", fn)
    assert m is not None
    assert m.group(1).strip() == "C.warning"
    assert m.group(2).strip() == "C.blueSoft"


def test_body_grid_dos_columnas_mapa_ocupa_fila_completa():
    assert ".body-grid{display:grid" in _TEMPLATE_SRC.replace(" ", "")
    assert ".body-left{display:flex;flex-direction:column" in _TEMPLATE_SRC.replace(" ", "")
    # el panel del mapa es hermano directo de .body-left dentro de .body-grid
    # (mismo grid row -> misma altura que el stack grafico+insight)
    body_html = re.search(r'<div class="body-grid">(.*?)</div>\s*</section>', _TEMPLATE_SRC, re.S)
    assert body_html is not None
    assert body_html.group(1).count('class="panel map-panel"') == 1
    assert 'class="body-left"' in body_html.group(1)


def test_insight_compacto_es_una_sola_tarjeta_con_tres_bloques():
    assert 'class="panel insight-compact"' in _TEMPLATE_SRC
    assert _TEMPLATE_SRC.count("ic-lectura") >= 1
    assert _TEMPLATE_SRC.count("ic-positivo") >= 1
    assert _TEMPLATE_SRC.count("ic-atender") >= 1
    # ya no debe existir la franja de insight de ancho completo previa
    assert 'class="insight"' not in _TEMPLATE_SRC
    assert 'class="insight-col' not in _TEMPLATE_SRC


def test_insight_compacto_frases_generadas_desde_payload(production_result):
    """Las 3 oraciones cortas se arman en JS a partir de kpis.ocupacion
    (no hardcodeadas): verificamos que los numeros que DEBERIAN aparecer en
    pantalla coinciden con los del payload de produccion."""
    oc = production_result["data"]["kpis"]["ocupacion"]
    render_fn = _extract_js_function(_TEMPLATE_SRC, "drawCompactInsights")
    assert "ocA.ocupados_totales" in render_fn
    assert "oc.espacios_totales_catalogo" in render_fn
    assert "ocA.ocupacion_pct" in render_fn
    assert "oc.delta_pp" in render_fn
    assert "ocA.estaciones_sobre_capacidad" in render_fn
    # Etapa 2A: las cifras fijas (1.294 / 2.443 / 53,0% / 24,5 pp / 107)
    # correspondian a la regla anterior (campañas distintas del mes). Se
    # reemplazan por identidades exactas del payload.
    cat = production_result["data"]["kpis"]["catalogo"]
    assert oc["espacios_totales_catalogo"] == cat["espacios"] == cat["espacios_digitales"] + cat["espacios_estaticos"]
    assert oc["actual"]["ocupacion_pct"] == round(oc["actual"]["ocupados_totales"] / oc["espacios_totales_catalogo"] * 100.0, 1)
    assert oc["delta_abs"] == oc["actual"]["ocupados_totales"] - oc["anterior"]["ocupados_totales"]


def test_mapa_proyeccion_preserva_aspect_ratio_geografico():
    """La proyeccion (compartida por el mapa nacional y el recorte AMBA via
    buildProjection) debe usar una unica escala px/grado con correccion
    coseno de latitud, no escalas X/Y independientes: eso es lo que evita
    que el territorio se vea estirado o deformado cuando el aspect ratio
    del panel no coincide con el del territorio real."""
    fn = _extract_js_function(_TEMPLATE_SRC, "buildProjection")
    assert "cosLat" in fn
    assert "Math.min(availW / (lonSpan * cosLat), availH / latSpan)" in fn.replace("  ", " ")
    # ambos mapas (nacional y AMBA) reutilizan la misma funcion, no duplican
    # la logica de proyeccion con formulas propias.
    assert _TEMPLATE_SRC.count("buildProjection(") >= 3  # definicion + 2 usos (nacional, amba)


def test_mapa_nota_breve_no_inventa_coordenadas(production_result):
    mapa = production_result["data"]["mapa"]
    assert "nota_breve" in mapa
    assert "no se inventan" in mapa["nota_breve"].lower()


# ---------------------------------------------------------------------------
# Ajuste puntual (2026-08-23): mapa nacional + recorte cuadrado AMBA.
# ---------------------------------------------------------------------------


def test_amba_bounds_es_un_filtro_geografico_fijo_no_una_suposicion():
    assert td.AMBA_BOUNDS["lat_min"] < td.AMBA_BOUNDS["lat_max"]
    assert td.AMBA_BOUNDS["lon_min"] < td.AMBA_BOUNDS["lon_max"]
    # Auditado contra produccion: 55 de 78 puntos caen en AMBA.
    assert td._en_amba(-34.6037, -58.3816) is True  # CABA (Obelisco)
    assert td._en_amba(-32.9468, -60.6393) is False  # Rosario, fuera de AMBA


def test_puntos_del_payload_tienen_flag_amba_sin_duplicar_ni_alterar(production_result):
    mapa = production_result["data"]["mapa"]
    puntos = mapa["puntos"]
    assert all("amba" in p for p in puntos)
    n_amba = sum(1 for p in puntos if p["amba"])
    n_no_amba = sum(1 for p in puntos if not p["amba"])
    assert n_amba == mapa["n_puntos_amba"]
    assert n_amba + n_no_amba == len(puntos) == mapa["n_estaciones_con_geo"]
    # cada punto es o AMBA o nacional, nunca las dos cosas (no se duplica)
    apies = [p["apie"] for p in puntos]
    assert len(apies) == len(set(apies))


def test_amba_marker_es_promedio_real_no_centroide_provincial(production_result):
    """El marcador AMBA del mapa nacional debe ser un promedio ponderado de
    coordenadas REALES de estaciones AMBA (nunca el centroide de un
    poligono de provincia disfrazado de estacion)."""
    mapa = production_result["data"]["mapa"]
    puntos_amba = [p for p in mapa["puntos"] if p["amba"]]
    marker = mapa["amba_marker"]
    assert marker is not None
    assert marker["estaciones"] == len(puntos_amba)
    assert marker["activaciones"] == sum(p["activaciones"] for p in puntos_amba)
    lats = [p["lat"] for p in puntos_amba]
    lons = [p["lon"] for p in puntos_amba]
    assert min(lats) <= marker["lat"] <= max(lats)
    assert min(lons) <= marker["lon"] <= max(lons)


def test_mendoza_no_es_amba_queda_en_el_mapa_nacional(production_result):
    mapa = production_result["data"]["mapa"]
    mendoza_puntos = [p for p in mapa["puntos"] if p.get("ciudad") and "mendoza" in str(p["ciudad"]).lower()]
    for p in mendoza_puntos:
        assert p["amba"] is False
    mendoza_prov = next(p for p in mapa["provincias"] if p["nombre"] == "Mendoza")
    assert mendoza_prov["estaciones_activas"] == 26


def test_template_declara_composicion_dual_nacional_mas_amba():
    assert 'data-geo-svg-national' in _TEMPLATE_SRC
    assert 'data-geo-svg-amba' in _TEMPLATE_SRC
    assert 'map-amba-wrap' in _TEMPLATE_SRC
    assert 'aspect-ratio:1/1' in _TEMPLATE_SRC.replace(' ', '')
    assert 'Detalle AMBA' in _TEMPLATE_SRC
    assert 'data-geo-connector' in _TEMPLATE_SRC


def test_leyenda_unica_compartida_no_duplicada_en_amba():
    for fn_name in ("drawAmbaGeneral", "drawAmbaZoom"):
        fn = _extract_js_function(_TEMPLATE_SRC, fn_name)
        assert "map-legend" not in fn, fn_name


# ---------------------------------------------------------------------------
# Ajuste puntual (2026-08-23): "Correccion del mapa AMBA" -- geometria
# politica real (comunas CABA + partidos GBA), zoom CABA+Zona Norte.
# ---------------------------------------------------------------------------


def test_geometrias_politicas_locales_existen_y_tienen_los_conteos_oficiales():
    assert td.CABA_COMUNAS_PATH.exists()
    assert td.GBA_PARTIDOS_PATH.exists()
    comunas = json.loads(td.CABA_COMUNAS_PATH.read_text(encoding="utf-8"))
    partidos = json.loads(td.GBA_PARTIDOS_PATH.read_text(encoding="utf-8"))
    assert len(comunas["features"]) == 15  # 15 comunas de CABA (Ley 1777/2005)
    assert len(partidos["features"]) == 135  # 135 partidos de la Pcia. de Buenos Aires (IGN)


def test_point_in_geometry_clasifica_correctamente_puntos_conocidos():
    comunas = json.loads(td.CABA_COMUNAS_PATH.read_text(encoding="utf-8"))
    # Obelisco (CABA, Comuna 1) debe caer dentro de alguna comuna.
    encontrado = any(
        td._point_in_geometry(-58.3816, -34.6037, f["geometry"]) for f in comunas["features"]
    )
    assert encontrado is True
    # Rosario (Santa Fe) no debe caer en ninguna comuna de CABA.
    fuera = any(
        td._point_in_geometry(-60.6393, -32.9468, f["geometry"]) for f in comunas["features"]
    )
    assert fuera is False


def test_zona_norte_es_una_lista_documentada_no_una_cantidad_hardcodeada():
    """ZONA_NORTE_PARTIDOS es una DEFINICION (nombres de partido reales,
    verificables contra gba_partidos.geojson), no una cantidad de estaciones
    -- el conteo siempre se recalcula por cruce punto-en-poligono."""
    partidos = json.loads(td.GBA_PARTIDOS_PATH.read_text(encoding="utf-8"))
    nombres_reales = {f["properties"]["nombre"] for f in partidos["features"]}
    assert td.ZONA_NORTE_PARTIDOS.issubset(nombres_reales)
    assert len(td.ZONA_NORTE_PARTIDOS) >= 8


def test_clasificacion_caba_zona_norte_produccion(production_result):
    mapa = production_result["data"]["mapa"]
    czn = mapa["caba_zona_norte"]
    assert czn["n_caba"] == 16
    assert czn["n_zona_norte"] == 18
    puntos_amba = [p for p in mapa["puntos"] if p["amba"]]
    assert sum(1 for p in puntos_amba if p["zona"] == "CABA") == czn["n_caba"]
    assert sum(1 for p in puntos_amba if p["zona"] == "Zona Norte") == czn["n_zona_norte"]
    # cada punto CABA debe tener numero de comuna real (1-15)
    for p in puntos_amba:
        if p["zona"] == "CABA":
            assert p["comuna"] in range(1, 16)
        if p["zona"] == "Zona Norte":
            assert p["partido"] in td.ZONA_NORTE_PARTIDOS


def test_template_declara_dual_amba_general_y_zoom():
    assert "drawAmbaGeneral" in _TEMPLATE_SRC
    assert "drawAmbaZoom" in _TEMPLATE_SRC
    assert "drawPoliticalLayers" in _TEMPLATE_SRC
    assert "data-amba-badge-caba" in _TEMPLATE_SRC
    assert "data-amba-badge-zn" in _TEMPLATE_SRC
    assert "CABA_COMUNAS_GEOJSON" in _TEMPLATE_SRC
    assert "GBA_PARTIDOS_GEOJSON" in _TEMPLATE_SRC


def test_placeholders_geojson_amba_reemplazados_en_html(production_html):
    assert "{{CABA_COMUNAS_GEOJSON}}" not in production_html
    assert "{{GBA_PARTIDOS_GEOJSON}}" not in production_html
    assert '"type":"FeatureCollection"' in production_html or '"type": "FeatureCollection"' in production_html


def test_badges_dinamicos_leen_del_payload_no_hardcodeados():
    fn = _extract_js_function(_TEMPLATE_SRC, "drawAmbaZoom")
    assert "cabaZN.n_caba" in fn
    assert "cabaZN.n_zona_norte" in fn


def test_zoom_maximo_5_etiquetas():
    fn = _extract_js_function(_TEMPLATE_SRC, "drawAmbaZoom")
    assert "maxLabels:5" in fn.replace(" ", "")


def test_5_todas_las_provincias_con_campanas_aparecen_en_cobertura(production_result):
    _, semantic_result, engine = __import__("export_data").load_pipeline(PRODUCTION_FILE)
    universe = td.build_tv4_universe(semantic_result)
    mapa = production_result["data"]["mapa"]
    nombres_payload = {p["nombre"] for p in mapa["provincias"]}
    # Recalculamos de forma independiente el set de provincias activas
    # (mismo metodo que build_tv4_dashboard.compute_mapa) y verificamos que
    # el payload no omite ninguna.
    apie_prov = universe["apie_provincia_map"]
    element_ids = universe["element_ids_all"]
    apie_map = universe["apie_map"]
    activos = set()
    for m_ in range(1, 8):
        start, end = td._period_bounds(2026, m_)
        overlap = td._campanas_overlap_validas(engine, element_ids, start, end)
        if overlap.empty:
            continue
        activ = overlap.drop_duplicates(subset=["IDCampaña", "ElementoID"]).copy()
        activ["_apie"] = activ["ElementoID"].map(apie_map)
        activos.update(a for a in activ["_apie"] if a is not None)
    esperado = {apie_prov[a] for a in activos if apie_prov.get(a)}
    assert nombres_payload == esperado


def test_6_mendoza_aparece_con_campanas_confirmadas(production_result):
    mapa = production_result["data"]["mapa"]
    mendoza = next((p for p in mapa["provincias"] if p["nombre"] == "Mendoza"), None)
    assert mendoza is not None, "Mendoza tiene campañas confirmadas en producción y debe aparecer en la cobertura"
    assert mendoza["estaciones_activas"] == 26
    assert mendoza["campanas_unicas"] == 6
    assert mendoza["estaciones_con_punto"] == 0  # sin coordenada -> cobertura territorial, no puntual
    assert mendoza["estaciones_solo_provincia"] == 26


def test_6b_provincia_sin_campanas_no_aparece_en_cobertura(production_result):
    mapa = production_result["data"]["mapa"]
    nombres = {p["nombre"] for p in mapa["provincias"]}
    # Jujuy no figura entre las provincias con actividad Ene-Jul en la base
    # vigente (auditado): no debe forzarse su aparicion.
    assert "Jujuy" not in nombres


def test_7_ninguna_coordenada_de_punto_es_inventada(production_result):
    mapa = production_result["data"]["mapa"]
    apie_con_punto = {p["apie"] for p in mapa["puntos"]}
    provincias_con_punto = {p["nombre"] for p in mapa["provincias"] if p["estaciones_con_punto"] > 0}
    # Ninguna provincia "solo territorial" (sin coordenadas) inventa un punto.
    solo_territorio = {p["nombre"] for p in mapa["provincias"] if p["estaciones_con_punto"] == 0}
    assert solo_territorio.isdisjoint(provincias_con_punto)
    assert "Mendoza" in solo_territorio


def test_7b_estaciones_sin_ubicacion_no_se_fuerzan_a_una_provincia(production_result):
    mapa = production_result["data"]["mapa"]
    assert mapa["n_estaciones_sin_ubicacion"] == 128
    total = mapa["n_estaciones_con_geo"] + mapa["n_estaciones_solo_provincia"] + mapa["n_estaciones_sin_ubicacion"]
    assert total == mapa["n_estaciones_con_actividad"]


def test_provincia_no_forzada_por_localidad_ambigua():
    """'Saladillo' (nombre de localidad, no de provincia) sin corroboracion
    de otro elemento de la misma estacion NO se resuelve a Buenos Aires por
    adivinanza: queda sin provincia."""
    row = _ypf_row(
        "700 - TT - 1", "700", medio="Digital",
        Observaciones="Provincia: Saladillo | Barrio: Centro | Dirección: Test 123",
    )
    _, _, universe = _build([row])
    assert universe["apie_provincia_map"].get("700") is None


def test_provincia_resuelta_por_corroboracion_de_otro_elemento_misma_estacion():
    """Si UN elemento de la estacion trae 'Provincia: Saladillo' y OTRO trae
    'Provincia: BUENOS AIRES', se resuelve a Buenos Aires (dato de la propia
    estacion, no una adivinanza externa)."""
    r1 = _ypf_row("701 - TT - 1", "701", medio="Digital", Observaciones="Provincia: Saladillo | X")
    r2 = _ypf_row("701 - TT - 2", "701", medio="Digital", Observaciones="Provincia: BUENOS AIRES | X")
    _, _, universe = _build([r1, r2])
    assert universe["apie_provincia_map"].get("701") == "Buenos Aires"


def test_gba_norte_sur_oeste_resuelven_a_buenos_aires():
    for zona in ("GBA Norte", "GBA Sur", "GBA Oeste"):
        row = _ypf_row("800 - TT - 1", "800", medio="Digital", Observaciones=f"Provincia: {zona} | X")
        _, _, universe = _build([row])
        assert universe["apie_provincia_map"].get("800") == "Buenos Aires", zona


def test_gba_a_revisar_no_se_resuelve():
    row = _ypf_row("801 - TT - 1", "801", medio="Digital", Observaciones="Provincia: GBA - Revisar | X")
    _, _, universe = _build([row])
    assert universe["apie_provincia_map"].get("801") is None


# ---------------------------------------------------------------------------
# Payload / reconciliacion / calidad
# ---------------------------------------------------------------------------


def test_payload_top_level_keys(production_result):
    assert set(production_result["data"].keys()) == {
        "meta", "universo", "kpis", "historico", "mapa", "calidad", "reconciliacion", "insights",
    }


def test_payload_contains_no_other_tv_datasets(production_json):
    for token in ("tv1_data", "tv2_data", "tv3_data", "tv5_data", "tv6_data"):
        assert token not in production_json.lower()


def test_reconciliacion_incluye_control_preliminar_conocido(production_result):
    filas = production_result["data"]["reconciliacion"]["filas"]
    metricas = {f["metrica"] for f in filas}
    assert metricas == set(td.CONTROL_PRELIMINAR_CONOCIDO)
    # 9 de 10 metricas deben reconciliar exacto contra el control conocido
    coincidencias = sum(1 for f in filas if f["coincide"])
    assert coincidencias >= 9


def test_calidad_reporta_filas_excluidas_por_revision_nula(production_result):
    avisos = " ".join(production_result["data"]["calidad"]["avisos"])
    assert "RevisionMaestro" in avisos or "revisión" in avisos.lower()


def test_apie_sin_derivar_bloquea_el_build():
    row = _maestro_row(
        "999 - TT - 1", CircuitoDashboard="YPF Digital", Subcircuito=None, Ubicacion="X",
        Medio="Digital", TipoCatalogo="Abierto", TipoInventario="Digital",
        CapacidadSlotsReel=0, SegundosDia=0, RevisionMaestro="TEST",
    )
    semantic_result = _semantic([row])
    with pytest.raises(td.BuildError, match="APIE"):
        td.build_tv4_universe(semantic_result)


def test_token_desconocido_bloquea_el_build():
    row = _ypf_row("500 - XX - 1", "500")
    semantic_result = _semantic([row])
    with pytest.raises(td.BuildError, match="token"):
        td.build_tv4_universe(semantic_result)


def test_revision_nula_excluye_del_universo_vigente():
    vigente = _digital("500", 1, "TT")
    obsoleta = _digital("501", 1, "TT", RevisionMaestro=None)
    _, _, universe = _build([vigente, obsoleta])
    assert universe["estaciones_catalogo"] == 1
    assert universe["calidad_filtro"]["filas_excluidas_revision_nula"] == 1
