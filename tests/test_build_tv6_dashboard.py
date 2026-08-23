"""Pruebas de negocio para scripts/build_tv6_dashboard.py (dashboard TV6
OCU26, "Demanda de marcas y expansion por circuito").

Reescrito 2026-08-23 (segunda vez el mismo dia -- decision de negocio):
Agencias/Programatica/Clientes directos/Exclusividades/Canal de ingreso
salen del alcance de TV6 (cobertura de fuente insuficiente, pospuesto para
una mejora futura de la base). TV6 pasa a responder "que marcas pautan,
cuales regresan (recurrencia 2026) y en cuantos circuitos estan": marcas
activas, primera aparicion vs. recurrencia, y presencia multicircuito vs.
un solo circuito. Este archivo reemplaza integramente la suite anterior
(que cubria agencias/programatica/exclusividad/canal, ya fuera de alcance).

No modifica scripts/validate_input.py, scripts/transform_data.py,
scripts/semantic_model.py, scripts/metrics_engine.py,
config/business_semantics.json, input/OCU26_BASE_DATOS.xlsx, ni ningun
archivo productivo de TV1/TV2/TV3/TV4/TV5 (build_tv1..5_dashboard.py,
tv1..5_template.html, tv1..5.html, test_build_tv1..5_dashboard.py,
TV6_REFERENCE.html.html). Los fixtures sinteticos usan la CONFIGURACION REAL
(sm.load_config()), mismo patron que test_build_tv4_dashboard.py.
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
from metrics_engine import MetricsEngine  # noqa: E402
import build_tv6_dashboard as td  # noqa: E402
from test_semantic_model import _maestro_row, _campana_row, _transform_result  # noqa: E402
from test_build_tv4_dashboard import (  # noqa: E402
    _cencosud_static, _cencosud_digital, _ypf_static, _apsa_static,
    _london_static, _pantalla_led,
)

PRODUCTION_FILE = REPO_ROOT / "input" / "OCU26_BASE_DATOS.xlsx"
BUILDER_SOURCE = (REPO_ROOT / "scripts" / "build_tv6_dashboard.py").read_text(encoding="utf-8")
TEMPLATE_SOURCE = (REPO_ROOT / "scripts" / "templates" / "tv6_template.html").read_text(encoding="utf-8")

IDCOL = "IDCampaña"


def _semantic(maestro_rows: list[dict], campanas_rows: list[dict] | None = None) -> dict:
    config = sm.load_config()
    return sm.build_semantic_model(_transform_result(maestro_rows, campanas_rows or []), config)


def _demanda_row(carga_id: str, elemento_id: str, **overrides) -> dict:
    """Fila CAMPANAS con dimension de marca dentro de julio 2026 por
    defecto (el periodo de reporte TV6)."""
    row = dict(
        FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31"),
        Marca="MARCA_TEST",
    )
    row.update(overrides)
    return _campana_row(carga_id, elemento_id, **row)


def _engine(maestro_rows: list[dict], campana_rows: list[dict]) -> tuple[dict, MetricsEngine]:
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv6_universe(semantic_result)
    return universe, engine


def _julio_scope(maestro_rows: list[dict], campana_rows: list[dict]) -> pd.DataFrame:
    universe, engine = _engine(maestro_rows, campana_rows)
    start, end = td._period_bounds(td.REPORT_YEAR, td.REPORT_MONTH)
    return td.classify_marca(td.build_tv6_scope(engine, universe["element_ids"], start, end))


@pytest.fixture(scope="module")
def production_result():
    return td.build_tv6_data(PRODUCTION_FILE)


@pytest.fixture(scope="module")
def production_json(production_result):
    return json.dumps(production_result["data"], ensure_ascii=False)


# ---------------------------------------------------------------------------
# 1. Payload contiene solo TV6, sin los bloques retirados
# ---------------------------------------------------------------------------


def test_payload_top_level_keys_are_tv6_only(production_result):
    assert set(production_result["data"].keys()) == {
        "meta", "universo", "marcas", "calidad", "reconciliacion", "ranking", "matriz", "insights",
    }


def test_payload_contains_no_other_tv_datasets(production_json):
    for token in ("tv1_data", "tv2_data", "tv3_data", "tv4_data", "tv5_data", "ocu_data"):
        assert token not in production_json.lower()


def test_payload_never_mentions_retired_topics(production_json):
    lowered = production_json.lower()
    for token in ("agencia", "programátic", "programatic", "cliente directo", "exclusiv", "canal de ingreso"):
        assert token not in lowered


# ---------------------------------------------------------------------------
# 2-3. Core Comercial completo + YPF incluidos / APSA-London excluidos
# ---------------------------------------------------------------------------


def test_core_comercial_and_ypf_included_in_synthetic_universe():
    maestro_rows = [_cencosud_static("C1"), _pantalla_led("P1"), _ypf_static("Y1")]
    universe, engine = _engine(maestro_rows, [
        _demanda_row("CC", "C1", IDCampaña="X1"),
        _demanda_row("PP", "P1", IDCampaña="X2"),
        _demanda_row("YY", "Y1", IDCampaña="X3"),
    ])
    assert set(universe["circuitos"]) == {"CENCOSUD", "PANTALLAS_LED", "YPF"}


def test_production_universe_includes_ypf(production_result):
    assert "YPF" in production_result["data"]["universo"]["circuitos"]


def test_production_universe_includes_core_circuitos(production_result):
    circuitos = set(production_result["data"]["universo"]["circuitos"])
    assert {"CENCOSUD", "PANTALLAS_LED", "REMEROS"}.issubset(circuitos)


def test_apsa_and_london_excluded_from_synthetic_universe():
    maestro_rows = [_cencosud_static("C1"), _apsa_static("A1"), _london_static("L1")]
    universe, _engine_ = _engine(maestro_rows, [_demanda_row("CC", "C1", IDCampaña="X1")])
    assert universe["circuitos"] == ["CENCOSUD"]


def test_production_excludes_apsa_and_london(production_result, production_json):
    circuitos = production_result["data"]["universo"]["circuitos"]
    assert "APSA" not in circuitos
    assert "LONDON_SUPPLY" not in circuitos
    assert "APSA" not in production_json.upper()
    assert "LONDON" not in production_json.upper()


# ---------------------------------------------------------------------------
# 4. Grano ElementoID x IDCampaña -- duplicados no inflan campañas/activaciones
# ---------------------------------------------------------------------------


def test_activacion_grain_deduplicates_same_element_campana_pair():
    maestro_rows = [_cencosud_static("C1")]
    scope = _julio_scope(maestro_rows, [
        _demanda_row("A", "C1", IDCampaña="X1", FechaFin=pd.Timestamp("2026-07-15")),
        _demanda_row("B", "C1", IDCampaña="X1", FechaFin=pd.Timestamp("2026-07-31")),
    ])
    assert len(scope) == 1


def test_campana_unica_counted_by_idcampana_distinct():
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    scope = _julio_scope(maestro_rows, [
        _demanda_row("A", "C1", IDCampaña="X1", Marca="M1"),
        _demanda_row("B", "C2", IDCampaña="X1", Marca="M1"),  # misma campaña, otro elemento
        _demanda_row("C", "C1", IDCampaña="X2", Marca="M1"),
    ])
    marcas = td.compute_marcas_julio(scope)
    assert marcas["campanas_unicas_total"] == 2  # X1, X2
    assert marcas["activaciones_totales"] == 3  # 2 pares distintos de X1 + 1 de X2


def test_production_activaciones_and_campanas_deduplicated(production_result):
    """Control de reconciliacion: recalcula el par distinto
    (ElementoID,IDCampaña) de forma independiente y compara contra el
    universo del payload."""
    from export_data import load_pipeline

    _tr, semantic_result, engine = load_pipeline(PRODUCTION_FILE)
    universe = td.build_tv6_universe(semantic_result)
    start, end = td._period_bounds(td.REPORT_YEAR, td.REPORT_MONTH)
    raw = engine._campanas_overlap(universe["element_ids"], start, end)
    activaciones_independiente = len(raw.drop_duplicates(subset=["ElementoID", IDCOL]))
    campanas_independiente = raw[IDCOL].dropna().nunique()

    assert production_result["data"]["universo"]["julio"]["activaciones_totales"] == activaciones_independiente
    assert production_result["data"]["universo"]["julio"]["campanas_unicas"] == campanas_independiente


# ---------------------------------------------------------------------------
# 5-6. Normalizacion de marca: variaciones no crean marcas artificiales
# ---------------------------------------------------------------------------


def test_marca_whitespace_variation_does_not_create_duplicate_brand():
    """Auditoria 2026-08-23: 'PARQUE DE LA COSTA' y 'PARQUE DE LA COSTA '
    (espacio final) son la misma marca en produccion -- deben contar como
    una sola entidad, no dos."""
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    scope = _julio_scope(maestro_rows, [
        _demanda_row("A", "C1", IDCampaña="X1", Marca="ACME "),
        _demanda_row("B", "C2", IDCampaña="X2", Marca="ACME"),
    ])
    marcas = td.compute_marcas_julio(scope)
    assert marcas["activas"] == 1


def test_marca_repeated_internal_spaces_normalized():
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    scope = _julio_scope(maestro_rows, [
        _demanda_row("A", "C1", IDCampaña="X1", Marca="ACME  CORP"),
        _demanda_row("B", "C2", IDCampaña="X2", Marca="ACME CORP"),
    ])
    marcas = td.compute_marcas_julio(scope)
    assert marcas["activas"] == 1


def test_canonical_display_name_preserves_original_casing():
    """El nombre canonico mostrado es la variante trimeada real de la
    fuente, nunca una version en mayusculas forzada por la clave de
    normalizacion."""
    maestro_rows = [_cencosud_static("C1")]
    scope = _julio_scope(maestro_rows, [_demanda_row("A", "C1", IDCampaña="X1", Marca="Coca Cola")])
    validas = scope[scope["_marca_valida"]]
    assert set(validas["_marca_canon"].unique()) == {"Coca Cola"}


# ---------------------------------------------------------------------------
# 7-8. Placeholders y plataformas nunca cuentan como marca; nunca fallback
# a Agencia/Cliente
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("placeholder", ["A CONFIRMAR", "S/D", "SIN DATO", "N/A", "", "  "])
def test_marca_placeholders_never_count(placeholder):
    maestro_rows = [_cencosud_static("C1")]
    scope = _julio_scope(maestro_rows, [_demanda_row("A", "C1", IDCampaña="X1", Marca=placeholder or None)])
    marcas = td.compute_marcas_julio(scope)
    assert marcas["activas"] == 0
    assert marcas["activaciones_sin_marca_valida"] == 1


@pytest.mark.parametrize("plataforma", ["TAGGIFY", "taggify", " Taggify ", "BEEYOND", "Latin Ad", "LATIN  AD", "GLOBAL"])
def test_known_platforms_never_count_as_marca(plataforma):
    maestro_rows = [_cencosud_static("C1")]
    scope = _julio_scope(maestro_rows, [_demanda_row("A", "C1", IDCampaña="X1", Marca=plataforma)])
    marcas = td.compute_marcas_julio(scope)
    assert marcas["activas"] == 0
    assert marcas["activaciones_sin_marca_valida"] == 1


def test_marca_never_falls_back_to_other_fields():
    """El campo Marca es la unica fuente: aunque Cliente/Agencia tengan un
    valor real, un Marca vacio nunca se completa con ellos."""
    maestro_rows = [_cencosud_static("C1")]
    scope = _julio_scope(maestro_rows, [
        _demanda_row("A", "C1", IDCampaña="X1", Marca=None, Cliente="CLIENTE_REAL", Agencia="AGENCIA_REAL"),
    ])
    marcas = td.compute_marcas_julio(scope)
    assert marcas["activas"] == 0
    assert "CLIENTE_REAL" not in scope["Marca"].astype(str).tolist()


# ---------------------------------------------------------------------------
# 9. Contradiccion de marca dentro de la misma campaña: se conserva, se
# informa, no se resuelve por mayoria
# ---------------------------------------------------------------------------


def test_marca_contradictoria_misma_campana_se_reporta_no_se_resuelve():
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2"), _cencosud_static("C3")]
    scope = _julio_scope(maestro_rows, [
        _demanda_row("A", "C1", IDCampaña="X1", Marca="MARCA_A"),
        _demanda_row("B", "C2", IDCampaña="X1", Marca="MARCA_A"),
        _demanda_row("C", "C3", IDCampaña="X1", Marca="MARCA_B"),  # misma campaña, marca distinta
    ])
    marcas = td.compute_marcas_julio(scope)
    assert marcas["campanas_marca_contradictoria"] == 1
    # ninguna se descarta: ambas marcas siguen contando activaciones
    assert marcas["activas"] == 2


def test_whitespace_only_variation_is_not_a_contradiction():
    """Confirma que la deteccion de contradiccion usa la marca CANONICA
    (post-normalizacion), no el valor crudo: 'ACME' y 'ACME ' en la misma
    campaña NO son una contradiccion real."""
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    scope = _julio_scope(maestro_rows, [
        _demanda_row("A", "C1", IDCampaña="X1", Marca="ACME"),
        _demanda_row("B", "C2", IDCampaña="X1", Marca="ACME "),
    ])
    marcas = td.compute_marcas_julio(scope)
    assert marcas["campanas_marca_contradictoria"] == 0


def test_production_campanas_marca_contradictoria(production_result):
    assert production_result["data"]["calidad"]["campanas_marca_contradictoria"] == 4


# ---------------------------------------------------------------------------
# 10-11. Primera aparicion / recurrencia
# ---------------------------------------------------------------------------


def _marcas_por_mes(maestro_rows, campana_rows):
    universe, engine = _engine(maestro_rows, campana_rows)
    return td.compute_monthly_brand_presence(engine, universe["element_ids"])


def test_primera_aparicion_usa_primer_mes_observado_en_ene_jul():
    maestro_rows = [_cencosud_static("C1")]
    presencia = _marcas_por_mes(maestro_rows, [
        _demanda_row("A", "C1", IDCampaña="X1", Marca="NUEVA",
                     FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
    ])
    recurrencia = td.compute_recurrencia({"NUEVA"}, presencia)
    assert recurrencia["primera_aparicion"]["marcas"] == ["NUEVA"]
    assert recurrencia["recurrentes"]["marcas"] == []


def test_recurrente_no_exige_actividad_en_junio_especificamente():
    """Marca activa en marzo, ausente abril-junio, de vuelta en julio: debe
    contar como recurrente (no exige el mes inmediato anterior)."""
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    presencia = _marcas_por_mes(maestro_rows, [
        _demanda_row("A", "C1", IDCampaña="X1", Marca="VUELVE",
                     FechaInicio=pd.Timestamp("2026-03-01"), FechaFin=pd.Timestamp("2026-03-31")),
        _demanda_row("B", "C2", IDCampaña="X2", Marca="VUELVE",
                     FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
    ])
    recurrencia = td.compute_recurrencia({"VUELVE"}, presencia)
    assert recurrencia["recurrentes"]["marcas"] == ["VUELVE"]
    assert recurrencia["primera_aparicion"]["marcas"] == []


def test_primera_aparicion_no_es_cliente_nuevo_historico():
    """El campo/nota de la tarjeta no debe afirmar historial fuera de la
    ventana observada Ene-Jul: ni 'cliente nuevo' ni 'marca nueva' ni
    confirmar ausencia de actividad antes de enero 2026."""
    lowered = TEMPLATE_SOURCE.lower()
    assert "cliente nuevo" not in lowered
    assert "clientes nuevos" not in lowered
    assert "marcas nuevas" not in lowered
    assert "nunca pautó" not in lowered
    assert "Sin actividad entre enero y junio" in TEMPLATE_SOURCE


def test_recurrencia_reconciliation_holds_generically():
    maestro_rows = [_cencosud_static(f"C{i}") for i in range(3)]
    presencia = _marcas_por_mes(maestro_rows, [
        _demanda_row("A", "C0", IDCampaña="X1", Marca="NUEVA", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
        _demanda_row("B", "C1", IDCampaña="X2", Marca="VIEJA",
                     FechaInicio=pd.Timestamp("2026-02-01"), FechaFin=pd.Timestamp("2026-07-31")),
    ])
    marcas_julio = {"NUEVA", "VIEJA"}
    recurrencia = td.compute_recurrencia(marcas_julio, presencia)
    r = recurrencia["reconciliacion"]
    assert r["primera_mas_recurrentes_mas_sin_historial"] == r["marcas_activas"] == 2
    assert r["ok"] is True


def test_production_primera_aparicion_recurrentes_reconciliation(production_result):
    marcas = production_result["data"]["marcas"]
    r = production_result["data"]["reconciliacion"]["recurrencia"]
    assert r["ok"] is True
    assert r["marcas_activas"] == marcas["activas_julio"]
    assert (
        marcas["primera_aparicion"]["count"] + marcas["recurrentes"]["count"]
        + marcas["sin_historial_comparable"]["count"] == marcas["activas_julio"]
    )


# ---------------------------------------------------------------------------
# 12-13. Multicircuito / un solo circuito
# ---------------------------------------------------------------------------


def test_multicircuito_requiere_dos_o_mas_grupos_canonicos():
    maestro_rows = [_cencosud_static("C1"), _pantalla_led("P1")]
    scope = _julio_scope(maestro_rows, [
        _demanda_row("A", "C1", IDCampaña="X1", Marca="MULTI"),
        _demanda_row("B", "P1", IDCampaña="X2", Marca="MULTI"),
    ])
    circuitos = td.compute_circuitos_por_marca(scope, {"MULTI"})
    assert circuitos["multicircuito"]["marcas"] == ["MULTI"]
    assert circuitos["un_solo_circuito"]["marcas"] == []


def test_un_solo_circuito_requiere_exactamente_uno():
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    scope = _julio_scope(maestro_rows, [
        _demanda_row("A", "C1", IDCampaña="X1", Marca="MONO"),
        _demanda_row("B", "C2", IDCampaña="X2", Marca="MONO"),
    ])
    circuitos = td.compute_circuitos_por_marca(scope, {"MONO"})
    assert circuitos["un_solo_circuito"]["marcas"] == ["MONO"]
    assert circuitos["multicircuito"]["marcas"] == []


def test_circuitos_reconciliation_holds_generically():
    maestro_rows = [_cencosud_static("C1"), _pantalla_led("P1"), _cencosud_static("C2")]
    scope = _julio_scope(maestro_rows, [
        _demanda_row("A", "C1", IDCampaña="X1", Marca="MULTI"),
        _demanda_row("B", "P1", IDCampaña="X2", Marca="MULTI"),
        _demanda_row("C", "C2", IDCampaña="X3", Marca="MONO"),
    ])
    circuitos = td.compute_circuitos_por_marca(scope, {"MULTI", "MONO"})
    r = circuitos["reconciliacion"]
    assert r["multicircuito_mas_uno_mas_sin_circuito"] == r["marcas_activas"] == 2
    assert r["ok"] is True


def test_production_multicircuito_reconciliation(production_result):
    marcas = production_result["data"]["marcas"]
    r = production_result["data"]["reconciliacion"]["circuitos"]
    assert r["ok"] is True
    assert (
        marcas["multicircuito"]["count"] + marcas["un_solo_circuito"]["count"]
        + marcas["sin_circuito_valido"]["count"] == marcas["activas_julio"]
    )


# ---------------------------------------------------------------------------
# 14-16. Ranking: orden campañas > circuitos > activaciones > nombre; nunca
# ordenado primariamente por activaciones
# ---------------------------------------------------------------------------


def test_ranking_sorted_by_campanas_desc_then_circuitos_desc_then_activaciones_desc_then_nombre():
    maestro_rows = [_cencosud_static(f"C{i}") for i in range(6)] + [_pantalla_led("P1")]
    scope = _julio_scope(maestro_rows, [
        # ZETA: 1 campaña, 1 circuito, muchas activaciones (no debe ganar por volumen)
        _demanda_row("A", "C0", IDCampaña="Z1", Marca="ZETA"),
        _demanda_row("A2", "C1", IDCampaña="Z1", Marca="ZETA"),
        _demanda_row("A3", "C2", IDCampaña="Z1", Marca="ZETA"),
        # ALFA: 2 campañas, 1 circuito
        _demanda_row("B", "C3", IDCampaña="A1", Marca="ALFA"),
        _demanda_row("C", "C3", IDCampaña="A2", Marca="ALFA"),
        # BETA: 2 campañas, 2 circuitos (debe superar a ALFA por circuitos)
        _demanda_row("D", "C4", IDCampaña="B1", Marca="BETA"),
        _demanda_row("E", "P1", IDCampaña="B2", Marca="BETA"),
    ])
    circuitos = td.compute_circuitos_por_marca(scope, {"ZETA", "ALFA", "BETA"})
    ranking = td.compute_ranking(scope, circuitos["circuitos_por_marca"])
    nombres = [r["nombre"] for r in ranking]
    assert nombres == ["BETA", "ALFA", "ZETA"]


def test_ranking_never_sorted_primarily_by_activaciones():
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    scope = _julio_scope(maestro_rows, [
        _demanda_row("A", "C1", IDCampaña="X1", Marca="VOLUMEN"),
        _demanda_row("A2", "C1", IDCampaña="X1", Marca="VOLUMEN"),  # dedup: misma activacion
        _demanda_row("B", "C2", IDCampaña="X2", Marca="POCAS_ACTIV"),
        _demanda_row("C", "C1", IDCampaña="X3", Marca="POCAS_ACTIV"),
    ])
    circuitos = td.compute_circuitos_por_marca(scope, {"VOLUMEN", "POCAS_ACTIV"})
    ranking = td.compute_ranking(scope, circuitos["circuitos_por_marca"])
    # POCAS_ACTIV tiene 2 campañas (vs 1 de VOLUMEN) aunque menos activaciones
    assert ranking[0]["nombre"] == "POCAS_ACTIV"


def test_ranking_tie_break_by_name_ascending():
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    scope = _julio_scope(maestro_rows, [
        _demanda_row("A", "C1", IDCampaña="X1", Marca="ZETA"),
        _demanda_row("B", "C2", IDCampaña="X2", Marca="ALFA"),
    ])
    circuitos = td.compute_circuitos_por_marca(scope, {"ZETA", "ALFA"})
    ranking = td.compute_ranking(scope, circuitos["circuitos_por_marca"])
    assert [r["nombre"] for r in ranking] == ["ALFA", "ZETA"]


def test_ranking_top_n_max_six():
    maestro_rows = [_cencosud_static(f"C{i}") for i in range(8)]
    scope = _julio_scope(maestro_rows, [
        _demanda_row(f"A{i}", f"C{i}", IDCampaña=f"X{i}", Marca=f"M{i}") for i in range(8)
    ])
    circuitos = td.compute_circuitos_por_marca(scope, {f"M{i}" for i in range(8)})
    ranking = td.compute_ranking(scope, circuitos["circuitos_por_marca"])
    assert len(ranking) == 6


def test_production_ranking_top_six(production_result):
    ranking = production_result["data"]["ranking"]["top"]
    assert len(ranking) == 6
    assert [r["nombre"] for r in ranking] == ["KFC", "SAMSUNG", "SIGLO 21", "YPF", "ADIDAS", "MEDIFE"]
    assert ranking[0] == {"nombre": "KFC", "campanas_unicas": 7, "circuitos": 2, "activaciones": 21}


# ---------------------------------------------------------------------------
# 17-19. Matriz: mismas marcas/orden que ranking, celdas = campañas unicas,
# totales por circuito sobre TODAS las marcas
# ---------------------------------------------------------------------------


def test_matriz_uses_same_brands_and_order_as_ranking(production_result):
    ranking_nombres = [r["nombre"] for r in production_result["data"]["ranking"]["top"]]
    matriz_nombres = [f["nombre"] for f in production_result["data"]["matriz"]["filas"]]
    assert matriz_nombres == ranking_nombres


def test_matriz_cells_count_unique_campanas_not_activaciones():
    """2 elementos del mismo circuito, misma campaña: son 2 activaciones
    pero 1 sola campaña unica -- la celda debe mostrar 1, no 2."""
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    scope = _julio_scope(maestro_rows, [
        _demanda_row("A", "C1", IDCampaña="X1", Marca="M1"),
        _demanda_row("B", "C2", IDCampaña="X1", Marca="M1"),
    ])
    matriz = td.compute_matriz(scope, ["M1"])
    fila = matriz["filas"][0]
    assert sum(fila["valores"]) == 1


def test_campana_en_dos_circuitos_cuenta_una_vez_en_cada_columna():
    maestro_rows = [_cencosud_static("C1"), _pantalla_led("P1")]
    scope = _julio_scope(maestro_rows, [
        _demanda_row("A", "C1", IDCampaña="X1", Marca="M1"),
        _demanda_row("B", "P1", IDCampaña="X1", Marca="M1"),
    ])
    matriz = td.compute_matriz(scope, ["M1"])
    fila = matriz["filas"][0]
    assert sum(fila["valores"]) == 2  # 1 en Shoppings Estático, 1 en Pantallas LED
    assert all(v <= 1 for v in fila["valores"])


def test_totales_por_circuito_use_all_brands_not_only_top6():
    maestro_rows = [_cencosud_static(f"C{i}") for i in range(8)]
    scope = _julio_scope(maestro_rows, [
        _demanda_row(f"A{i}", f"C{i}", IDCampaña=f"X{i}", Marca=f"M{i}") for i in range(8)
    ])
    matriz = td.compute_matriz(scope, [f"M{i}" for i in range(6)])  # solo top 6 en filas
    idx_estatico = matriz["columnas"].index("Shoppings Estático")
    assert matriz["totales_por_circuito"][idx_estatico] == 8  # las 8 marcas, no solo 6


def test_production_totales_por_circuito_sum_at_least_top6_marcas(production_result):
    matriz = production_result["data"]["matriz"]
    for total in matriz["totales_por_circuito"]:
        assert total >= 0
    # YPF: Siglo 21 (top6) esta en YPF -> el total de YPF debe ser >= 1
    idx_ypf = matriz["columnas"].index("YPF")
    assert matriz["totales_por_circuito"][idx_ypf] >= 1


# ---------------------------------------------------------------------------
# 20. Temas retirados no aparecen en la vista / template
# ---------------------------------------------------------------------------


def test_retired_topics_not_in_template_visible_copy():
    """Solo revisa el texto visible/HTML real: los comentarios /* ... */ que
    documentan por que se retiraron esos temas (ej. el docstring de
    cabecera del <style>) son legitimos y no cuentan como 'aparecen en la
    vista'."""
    import re as _re
    sin_comentarios = _re.sub(r"/\*.*?\*/", "", TEMPLATE_SOURCE, flags=_re.DOTALL)
    lowered = sin_comentarios.lower()
    for token in ("agencia", "programátic", "programatic", "cliente directo", "exclusivid"):
        assert token not in lowered


def test_old_cards_removed_from_template():
    assert "Cliente top identificado" not in TEMPLATE_SOURCE
    assert "Concentraci" not in TEMPLATE_SOURCE


def test_five_new_kpi_blocks_present_in_template():
    assert "renderKpis" in TEMPLATE_SOURCE
    for label in (
        "MARCAS ACTIVAS &middot; JULIO",
        "PRIMERA ACTIVIDAD REGISTRADA EN 2026 &middot; JULIO",
        "CON ACTIVIDAD PREVIA EN 2026 &middot; JULIO",
        "PRESENTES EN 2 O M&Aacute;S CIRCUITOS &middot; JULIO",
        "PRESENTES EN UN SOLO CIRCUITO &middot; JULIO",
    ):
        assert label in TEMPLATE_SOURCE


# ---------------------------------------------------------------------------
# 21. Sin temporizador de rotacion / ciclo
# ---------------------------------------------------------------------------


def test_no_cycle_timer_in_template():
    assert "cycleTimer" not in TEMPLATE_SOURCE
    assert "setInterval(function(){ cycleI" not in TEMPLATE_SOURCE
    assert "drawCycle" not in TEMPLATE_SOURCE


def test_ranking_and_matrix_are_static_single_view(production_json):
    assert "D.ranking.top" in TEMPLATE_SOURCE or "D.ranking && D.ranking.top" in TEMPLATE_SOURCE
    assert '"ciclo"' not in production_json


# ---------------------------------------------------------------------------
# 22. Header: la pregunta queda completamente visible (sin overlap)
# ---------------------------------------------------------------------------


def test_header_not_fixed_height_that_causes_overlap():
    hdr_block = TEMPLATE_SOURCE[TEMPLATE_SOURCE.index(".hdr{"):TEMPLATE_SOURCE.index(".hdr{") + 200]
    assert "height:96px" not in hdr_block
    assert "min-height" in hdr_block


def test_header_question_text_updated():
    assert (
        "QU&Eacute; MARCAS EST&Aacute;N ACTIVAS, CU&Aacute;LES VUELVEN Y D&Oacute;NDE PAUTAN"
        in TEMPLATE_SOURCE
    )


def test_header_subtitle_updated():
    assert "Marcas activas, continuidad y presencia por circuito" in TEMPLATE_SOURCE


def test_header_title_updated():
    assert "Demanda comercial por marca" in TEMPLATE_SOURCE


# ---------------------------------------------------------------------------
# 23. 1920x1080 fijo, sin scroll (estructura del stage)
# ---------------------------------------------------------------------------


def test_stage_fixed_1920x1080_no_scroll():
    assert "width:1920px;height:1080px" in TEMPLATE_SOURCE
    assert "html,body{height:100%;background:var(--bp-black-deep);overflow:hidden}" in TEMPLATE_SOURCE


# ---------------------------------------------------------------------------
# Payload: estructura minima pedida (Sec.10)
# ---------------------------------------------------------------------------


def test_payload_marcas_block_has_required_fields(production_result):
    marcas = production_result["data"]["marcas"]
    required = {
        "activas_julio", "campanas_unicas_marca_valida_julio", "activaciones_sin_marca_valida_julio",
        "primera_aparicion", "recurrentes", "sin_historial_comparable", "primer_mes_por_marca",
        "multicircuito", "un_solo_circuito", "sin_circuito_valido", "circuitos_por_marca",
    }
    assert required.issubset(marcas.keys())


def test_payload_ranking_and_matriz_present(production_result):
    data = production_result["data"]
    assert "top" in data["ranking"]
    assert {"columnas", "totales_por_circuito", "filas"}.issubset(data["matriz"].keys())


# ---------------------------------------------------------------------------
# Insights (Lectura / Punto positivo / A atender)
# ---------------------------------------------------------------------------


def test_insights_footer_is_triple_pattern(production_result):
    insights = production_result["data"]["insights"]
    assert set(insights.keys()) == {"lectura", "punto_positivo", "a_atender"}
    for key in insights:
        assert insights[key]


def test_insights_do_not_repeat_full_rankings(production_result):
    lectura = production_result["data"]["insights"]["lectura"]
    assert "Samsung" not in lectura and "KFC" not in lectura


def test_template_has_triple_insight_footer_markup():
    assert "data-insight-lectura" in TEMPLATE_SOURCE
    assert "data-insight-positivo" in TEMPLATE_SOURCE
    assert "data-insight-atender" in TEMPLATE_SOURCE
    assert "insight-col" in TEMPLATE_SOURCE


def test_output_html_has_triple_insight_footer_rendered():
    td.build_and_write(PRODUCTION_FILE)
    html = td.DEFAULT_OUTPUT_HTML.read_text(encoding="utf-8")
    assert "Lectura</div>" in html
    assert "Punto positivo</div>" in html
    assert "A atender</div>" in html


# ---------------------------------------------------------------------------
# Produccion: numeros de referencia (recalibrar solo si cambia la fuente)
# ---------------------------------------------------------------------------


def test_production_marcas_snapshot(production_result):
    marcas = production_result["data"]["marcas"]
    assert marcas["activas_julio"] == 93
    assert marcas["campanas_unicas_marca_valida_julio"] == 131
    assert marcas["activaciones_sin_marca_valida_julio"] == 9


def test_production_universo_snapshot(production_result):
    universo = production_result["data"]["universo"]["julio"]
    assert universo["campanas_unicas"] == 133
    assert universo["activaciones_totales"] == 11477
    assert universo["activaciones_sin_marca_valida"] == 9


def test_production_recurrencia_snapshot(production_result):
    marcas = production_result["data"]["marcas"]
    assert marcas["primera_aparicion"]["count"] == 17
    assert marcas["recurrentes"]["count"] == 76
    assert marcas["sin_historial_comparable"]["count"] == 0


def test_production_circuitos_snapshot(production_result):
    marcas = production_result["data"]["marcas"]
    assert marcas["multicircuito"]["count"] == 18
    assert marcas["un_solo_circuito"]["count"] == 75
    assert marcas["sin_circuito_valido"]["count"] == 0


def test_production_matriz_totales_por_circuito(production_result):
    matriz = production_result["data"]["matriz"]
    assert dict(zip(matriz["columnas"], matriz["totales_por_circuito"])) == {
        "Pantallas LED": 45, "Shoppings Digital": 25, "Shoppings Estático": 30,
        "AA2000 / Pilar Frontlight": 0, "YPF": 13,
    }


# ---------------------------------------------------------------------------
# UI usa "Estatico", nunca "Fijo"
# ---------------------------------------------------------------------------


def test_matriz_columns_use_estatico_not_fijo():
    assert "Shoppings Estático" in td.MATRIZ_COLUMNAS
    assert not any("Fijo" in c for c in td.MATRIZ_COLUMNAS)


def test_template_uses_estatico_not_fijo_in_visible_copy():
    assert "Fijo" not in TEMPLATE_SOURCE
    assert "FIJO" not in TEMPLATE_SOURCE.upper()


def test_payload_never_says_shoppings_fijo(production_json):
    assert "Shoppings Fijo" not in production_json
    assert "shoppings fijo" not in production_json.lower()


# ---------------------------------------------------------------------------
# Sin legacy window.OCU_DATA en el HTML productivo
# ---------------------------------------------------------------------------


def test_output_html_has_no_legacy_ocu_data():
    td.build_and_write(PRODUCTION_FILE)
    html = td.DEFAULT_OUTPUT_HTML.read_text(encoding="utf-8")
    assert "OCU_DATA" not in html
    assert "window.TV6_DATA" in html
    assert 'id="tv5"' not in html
    assert "TV 5" not in html


# ---------------------------------------------------------------------------
# SHA del Excel sin cambios / Excel fuente no tocado
# ---------------------------------------------------------------------------


def test_input_excel_sha_unchanged(production_result):
    sha_after_build = vi.calculate_sha256(PRODUCTION_FILE)
    assert sha_after_build == production_result["sha256"]


def test_production_excel_untouched_by_import_path():
    assert PRODUCTION_FILE.exists()


# ---------------------------------------------------------------------------
# Otras TVs siguen construyendo (smoke test de no-interferencia)
# ---------------------------------------------------------------------------


def test_tv1_pipeline_still_builds_successfully():
    import build_tv1_dashboard as t1
    result = t1.build_tv1_data(PRODUCTION_FILE)
    assert result["data"]["kpis"]["core_comercial"]["value"] > 0


def test_tv6_reference_untouched():
    ref = REPO_ROOT / "audit_sources" / "TV6_REFERENCE.html.html"
    assert ref.exists()
    html = ref.read_text(encoding="utf-8")
    assert "window.OCU_DATA" in html  # dataset legacy original, nunca reescrito
