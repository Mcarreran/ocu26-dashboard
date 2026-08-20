"""Pruebas de negocio para scripts/build_tv6_dashboard.py (dashboard TV6
OCU26, Demanda Comercial = Core Comercial completo + YPF).

No modifica scripts/validate_input.py, scripts/transform_data.py,
scripts/semantic_model.py, scripts/metrics_engine.py,
config/business_semantics.json, input/OCU26_BASE_DATOS.xlsx, ni ningun
archivo productivo de TV1/TV2/TV3/TV4 (build_tv1..4_dashboard.py,
tv1..4_template.html, tv1..4.html, test_build_tv1..4_dashboard.py,
TV1..4/TV6_REFERENCE.html.html). Los fixtures sinteticos usan la
CONFIGURACION REAL (sm.load_config()), mismo patron que
test_build_tv4_dashboard.py.
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
from export_data import load_pipeline  # noqa: E402
import build_tv6_dashboard as td  # noqa: E402
from test_semantic_model import _maestro_row, _campana_row, _transform_result  # noqa: E402
from test_build_tv4_dashboard import (  # noqa: E402
    _cencosud_static, _cencosud_digital, _ypf_static, _apsa_static,
    _london_static, _pantalla_led,
)

PRODUCTION_FILE = REPO_ROOT / "input" / "OCU26_BASE_DATOS.xlsx"
BUILDER_SOURCE = (REPO_ROOT / "scripts" / "build_tv6_dashboard.py").read_text(encoding="utf-8")


def _semantic(maestro_rows: list[dict], campanas_rows: list[dict] | None = None) -> dict:
    config = sm.load_config()
    return sm.build_semantic_model(_transform_result(maestro_rows, campanas_rows or []), config)


def _demanda_row(carga_id: str, elemento_id: str, **overrides) -> dict:
    """Fila CAMPANAS con dimensiones de demanda (Cliente/Marca/Agencia/
    PROGRAMATICA) dentro del periodo de reporte TV6 (julio 2026)."""
    row = dict(
        FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31"),
        Cliente="CLIENTE_TEST", Marca="MARCA_TEST", Agencia="No", PROGRAMATICA="No",
    )
    row.update(overrides)
    return _campana_row(carga_id, elemento_id, **row)


def _build_scope(maestro_rows: list[dict], campana_rows: list[dict]):
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv6_universe(semantic_result)
    start, end = td._period_bounds(2026, 7)
    scope = td.build_tv6_scope(engine, universe["element_ids"], start, end)
    return universe, scope


@pytest.fixture(scope="module")
def production_result():
    return td.build_tv6_data(PRODUCTION_FILE)


@pytest.fixture(scope="module")
def production_json(production_result):
    return json.dumps(production_result["data"], ensure_ascii=False)


# ---------------------------------------------------------------------------
# 1. Payload contiene solo TV6
# ---------------------------------------------------------------------------


def test_payload_top_level_keys_are_tv6_only(production_result):
    assert set(production_result["data"].keys()) == {
        "meta", "universo", "kpis", "ranking", "matriz", "pendientes", "insights",
    }


def test_payload_contains_no_other_tv_datasets(production_json):
    for token in ("tv1_data", "tv2_data", "tv3_data", "tv4_data", "ocu_data"):
        assert token not in production_json.lower()


# ---------------------------------------------------------------------------
# 2-3. Core Comercial completo + YPF incluidos
# ---------------------------------------------------------------------------


def test_core_comercial_and_ypf_included_in_synthetic_universe():
    maestro_rows = [_cencosud_static("C1"), _pantalla_led("P1"), _ypf_static("Y1")]
    universe, scope = _build_scope(maestro_rows, [
        _demanda_row("CC", "C1", IDCampaña="X1"),
        _demanda_row("PP", "P1", IDCampaña="X2"),
        _demanda_row("YY", "Y1", IDCampaña="X3"),
    ])
    assert set(universe["circuitos"]) == {"CENCOSUD", "PANTALLAS_LED", "YPF"}
    assert set(scope["CircuitoNegocio"].unique()) == {"CENCOSUD", "PANTALLAS_LED", "YPF"}


def test_production_universe_includes_ypf(production_result):
    assert "YPF" in production_result["data"]["universo"]["circuitos"]


def test_production_universe_includes_core_circuitos(production_result):
    circuitos = set(production_result["data"]["universo"]["circuitos"])
    assert {"CENCOSUD", "PANTALLAS_LED", "REMEROS"}.issubset(circuitos)


# ---------------------------------------------------------------------------
# 4-5. APSA / London excluidos
# ---------------------------------------------------------------------------


def test_apsa_and_london_excluded_from_synthetic_universe():
    maestro_rows = [_cencosud_static("C1"), _apsa_static("A1"), _london_static("L1")]
    universe, _ = _build_scope(maestro_rows, [_demanda_row("CC", "C1", IDCampaña="X1")])
    assert universe["circuitos"] == ["CENCOSUD"]


def test_production_excludes_apsa_and_london(production_result, production_json):
    circuitos = production_result["data"]["universo"]["circuitos"]
    assert "APSA" not in circuitos
    assert "LONDON_SUPPLY" not in circuitos
    assert "APSA" not in production_json.upper()
    assert "LONDON" not in production_json.upper()


# ---------------------------------------------------------------------------
# 6-7. Marcas y agencias deduplicadas Core + YPF
# ---------------------------------------------------------------------------


def test_marca_present_in_core_and_ypf_counts_once():
    maestro_rows = [_cencosud_static("C1"), _ypf_static("Y1")]
    _, scope = _build_scope(maestro_rows, [
        _demanda_row("CC", "C1", IDCampaña="X1", Marca="MISMA_MARCA"),
        _demanda_row("YY", "Y1", IDCampaña="X2", Marca="MISMA_MARCA"),
    ])
    marcas = td.compute_marcas(scope)
    assert marcas["activas"] == 1
    assert marcas["ranking_top"][0]["activaciones"] == 2  # dedup de entidad, no de actividad


def test_agencia_present_in_core_and_ypf_counts_once():
    maestro_rows = [_cencosud_static("C1"), _ypf_static("Y1")]
    _, scope = _build_scope(maestro_rows, [
        _demanda_row("CC", "C1", IDCampaña="X1", Agencia="MISMA_AGENCIA"),
        _demanda_row("YY", "Y1", IDCampaña="X2", Agencia="MISMA_AGENCIA"),
    ])
    agencias = td.compute_agencias(scope)
    assert agencias["activas"] == 1
    assert agencias["ranking_top"][0]["activaciones"] == 2


# ---------------------------------------------------------------------------
# 8-9. Programatica es subset de Agencias, sin condicion inventada
# ---------------------------------------------------------------------------


def test_programatica_is_subset_of_agencias(production_result):
    agencias_nombres = {a["nombre"] for a in production_result["data"]["ranking"]["agencias"]}
    prog_nombres = {a["nombre"] for a in production_result["data"]["ranking"]["programatica"]["ranking_top"]}
    assert prog_nombres.issubset(agencias_nombres) or not prog_nombres


def test_agencia_with_programatica_no_never_ranked_as_programatica():
    maestro_rows = [_cencosud_static("C1")]
    _, scope = _build_scope(maestro_rows, [
        _demanda_row("CC", "C1", IDCampaña="X1", Agencia="SOLO_TRADICIONAL", PROGRAMATICA="No"),
    ])
    agencias = td.compute_agencias(scope)
    prog = td.compute_programatica(scope, agencias["nombres_identificados"])
    assert prog["estado"] == "A_VALIDAR"
    assert prog["ranking_top"] == []


def test_agencia_with_programatica_si_is_ranked():
    maestro_rows = [_cencosud_static("C1")]
    _, scope = _build_scope(maestro_rows, [
        _demanda_row("CC", "C1", IDCampaña="X1", Agencia="AGENCIA_PROG", PROGRAMATICA="Si"),
    ])
    agencias = td.compute_agencias(scope)
    prog = td.compute_programatica(scope, agencias["nombres_identificados"])
    assert prog["estado"] == "OK"
    assert prog["ranking_top"] == [{"nombre": "AGENCIA_PROG", "activaciones": 1}]


def test_programatica_blank_or_a_confirmar_never_treated_as_positive():
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    _, scope = _build_scope(maestro_rows, [
        _demanda_row("CC", "C1", IDCampaña="X1", Agencia="AGENCIA_X", PROGRAMATICA=None),
        _demanda_row("CC2", "C2", IDCampaña="X2", Agencia="AGENCIA_X", PROGRAMATICA="A confirmar"),
    ])
    agencias = td.compute_agencias(scope)
    prog = td.compute_programatica(scope, agencias["nombres_identificados"])
    assert prog["ranking_top"] == []  # ni vacio ni "A confirmar" se convierten en Si


# ---------------------------------------------------------------------------
# 10. No inferencia manual / lista hardcodeada de agencias programaticas
# ---------------------------------------------------------------------------


def test_no_hardcoded_agency_name_classification_in_builder_source():
    """El nombre de una agencia real solo puede aparecer en un comentario/
    docstring documentando el hallazgo de auditoria, nunca como literal de
    codigo (lista/set/comparacion) que clasificaria Programatica por nombre."""
    for real_agency_name in ("GROUPM", "CARAT"):
        for quoted in (f'"{real_agency_name}"', f"'{real_agency_name}'"):
            assert quoted not in BUILDER_SOURCE


def test_programatica_uses_only_canonical_field():
    assert '"PROGRAMATICA"' in BUILDER_SOURCE or "'PROGRAMATICA'" in BUILDER_SOURCE


# ---------------------------------------------------------------------------
# 11-12. Clientes directos excluyen placeholders / via agencia; top valido
# ---------------------------------------------------------------------------


def test_clientes_directos_exclude_via_agencia_and_a_confirmar():
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2"), _cencosud_static("C3")]
    _, scope = _build_scope(maestro_rows, [
        _demanda_row("D", "C1", IDCampaña="X1", Cliente="CLIENTE_DIRECTO", Agencia="No"),
        _demanda_row("V", "C2", IDCampaña="X2", Cliente="AGENCIA", Agencia="AGENCIA_REAL"),
        _demanda_row("A", "C3", IDCampaña="X3", Cliente="A CONFIRMAR", Agencia="A confirmar"),
    ])
    clientes = td.compute_clientes_directos(scope)
    assert clientes["activos"] == 1
    assert clientes["ranking"] == [{"nombre": "CLIENTE_DIRECTO", "activaciones": 1}]


def test_cliente_directo_requires_agencia_no_not_just_blank():
    """Agencia vacia (NaN) no es lo mismo que Agencia='No' confirmado: no se
    puede asumir que un cliente sin dato de agencia es directo."""
    maestro_rows = [_cencosud_static("C1")]
    _, scope = _build_scope(maestro_rows, [
        _demanda_row("X", "C1", IDCampaña="X1", Cliente="CLIENTE_AMBIGUO", Agencia=None),
    ])
    clientes = td.compute_clientes_directos(scope)
    assert clientes["activos"] == 0


def test_cliente_top_identificado_excludes_placeholder_even_if_largest():
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    _, scope = _build_scope(maestro_rows, [
        _demanda_row(f"A{i}", "C1", IDCampaña=f"BIG{i}", Cliente="A CONFIRMAR", Agencia="No")
        for i in range(10)
    ] + [_demanda_row("B", "C2", IDCampaña="SMALL", Cliente="CLIENTE_REAL", Agencia="No")])
    clientes = td.compute_clientes_directos(scope)
    assert clientes["top_identificado"] == {"nombre": "CLIENTE_REAL", "activaciones": 1}


def test_production_cliente_top_identificado_is_valid(production_result):
    top = production_result["data"]["kpis"]["cliente_top_identificado"]
    assert top is not None
    assert top["nombre"] not in ("A CONFIRMAR", "AGENCIA")
    assert top["activaciones"] > 0


# ---------------------------------------------------------------------------
# 13-15. Rankings correctos (orden actividad desc, nombre asc) y Top5
# ---------------------------------------------------------------------------


def test_ranking_marcas_sorted_desc_then_name_asc():
    maestro_rows = [_cencosud_static(f"C{i}") for i in range(4)]
    rows = [
        _demanda_row("A1", "C0", IDCampaña="A1", Marca="ZETA"),
        _demanda_row("A2", "C1", IDCampaña="A2", Marca="ALFA"),
        _demanda_row("A3", "C2", IDCampaña="A3", Marca="ALFA"),
        _demanda_row("A4", "C3", IDCampaña="A4", Marca="BETA"),
    ]
    _, scope = _build_scope(maestro_rows, rows)
    marcas = td.compute_marcas(scope)
    nombres = [m["nombre"] for m in marcas["ranking_top"]]
    assert nombres == ["ALFA", "BETA", "ZETA"]


def test_ranking_agencias_sorted_desc_then_name_asc():
    maestro_rows = [_cencosud_static(f"C{i}") for i in range(3)]
    rows = [
        _demanda_row("A1", "C0", IDCampaña="A1", Agencia="ZETA_AG"),
        _demanda_row("A2", "C1", IDCampaña="A2", Agencia="ALFA_AG"),
        _demanda_row("A3", "C2", IDCampaña="A3", Agencia="ALFA_AG"),
    ]
    _, scope = _build_scope(maestro_rows, rows)
    agencias = td.compute_agencias(scope)
    nombres = [a["nombre"] for a in agencias["ranking_top"]]
    assert nombres == ["ALFA_AG", "ZETA_AG"]


def test_concentracion_top5_numerador_denominador_explicit():
    maestro_rows = [_cencosud_static(f"C{i}") for i in range(6)]
    rows = [_demanda_row(f"R{i}", f"C{i}", IDCampaña=f"X{i}", Marca=f"M{i}") for i in range(5)]
    rows += [_demanda_row("R5", "C5", IDCampaña="X5", Marca="M5")]
    _, scope = _build_scope(maestro_rows, rows)
    marcas = td.compute_marcas(scope)
    conc = td.compute_concentracion_top5(scope, marcas["ranking_full"])
    assert conc["denominador"] == 6
    assert conc["numerador"] == 5  # top 5 de 6 marcas de 1 activacion cada una
    assert conc["pct"] == pytest.approx(5 / 6 * 100, abs=0.05)
    assert conc["unidad"] == "activaciones"


def test_production_concentracion_top5_matches_manual_calc(production_result):
    """La card Concentracion Top 5 sigue siendo JULIO (Sec.5 del ajuste
    hibrido): se recalcula de forma independiente sobre el scope de julio
    (no el acumulado Ene-Jul de ranking.marcas) y debe coincidir exacto."""
    _transform_result, semantic_result, engine = load_pipeline(PRODUCTION_FILE)
    universe = td.build_tv6_universe(semantic_result)
    start, end = td._period_bounds(td.REPORT_YEAR, td.REPORT_MONTH)
    scope_julio = td.build_tv6_scope(engine, universe["element_ids"], start, end)
    marcas_julio = td.compute_marcas(scope_julio)
    conc_manual = td.compute_concentracion_top5(scope_julio, marcas_julio["ranking_full"])
    assert conc_manual == production_result["data"]["kpis"]["concentracion_top5"]


# ---------------------------------------------------------------------------
# 16-17. Ranking Programatica correcto + ciclo exacto
# ---------------------------------------------------------------------------


def test_ranking_programatica_correct_for_flagged_agency():
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2")]
    _, scope = _build_scope(maestro_rows, [
        _demanda_row("P1", "C1", IDCampaña="P1", Agencia="AG_PROG", PROGRAMATICA="Si"),
        _demanda_row("P2", "C2", IDCampaña="P2", Agencia="AG_NORMAL", PROGRAMATICA="No"),
    ])
    agencias = td.compute_agencias(scope)
    prog = td.compute_programatica(scope, agencias["nombres_identificados"])
    assert prog["ranking_top"] == [{"nombre": "AG_PROG", "activaciones": 1}]


def test_ciclo_is_exactly_marcas_agencias_programatica(production_result):
    assert production_result["data"]["ranking"]["ciclo"] == ["marcas", "agencias", "programatica"]


def test_ciclo_never_includes_clientes(production_json):
    ciclo = json.loads(production_json)["ranking"]["ciclo"]
    assert "clientes" not in ciclo


# ---------------------------------------------------------------------------
# 18. Matriz reconcilia con la medida de demanda elegida (activaciones)
# ---------------------------------------------------------------------------


def test_matriz_reconciles_with_scope_activaciones():
    maestro_rows = [_cencosud_static("C1"), _pantalla_led("P1"), _ypf_static("Y1")]
    _, scope = _build_scope(maestro_rows, [
        _demanda_row("CC", "C1", IDCampaña="X1", Marca="M1"),
        _demanda_row("PP", "P1", IDCampaña="X2", Marca="M1"),
        _demanda_row("YY", "Y1", IDCampaña="X3", Marca="M1"),
    ])
    matriz = td.compute_matriz(scope, ["M1"])
    fila = matriz["filas"][0]
    assert sum(fila["valores"]) == 3
    assert matriz["columnas"] == td.MATRIZ_COLUMNAS


def test_production_matriz_row_totals_never_exceed_marca_activaciones(production_result):
    marcas_by_name = {m["nombre"]: m["activaciones"] for m in production_result["data"]["ranking"]["marcas"]}
    for fila in production_result["data"]["matriz"]["filas"]:
        assert sum(fila["valores"]) == marcas_by_name[fila["nombre"]]


# ---------------------------------------------------------------------------
# 19. UI usa "Estatico", nunca "Fijo"
# ---------------------------------------------------------------------------


def test_matriz_columns_use_estatico_not_fijo():
    assert "Shoppings Estático" in td.MATRIZ_COLUMNAS
    assert not any("Fijo" in c for c in td.MATRIZ_COLUMNAS)


def test_template_uses_estatico_not_fijo_in_visible_copy():
    template = (REPO_ROOT / "scripts" / "templates" / "tv6_template.html").read_text(encoding="utf-8")
    assert "Fijo" not in template
    assert "FIJO" not in template.upper()


def test_payload_never_says_shoppings_fijo(production_json):
    assert "Shoppings Fijo" not in production_json
    assert "shoppings fijo" not in production_json.lower()


# ---------------------------------------------------------------------------
# 20. Sin legacy window.OCU_DATA en el HTML productivo
# ---------------------------------------------------------------------------


def test_output_html_has_no_legacy_ocu_data():
    td.build_and_write(PRODUCTION_FILE)
    html = td.DEFAULT_OUTPUT_HTML.read_text(encoding="utf-8")
    assert "OCU_DATA" not in html
    assert "window.TV6_DATA" in html
    assert 'id="tv5"' not in html
    assert "TV 5" not in html


# ---------------------------------------------------------------------------
# 21. SHA del Excel sin cambios
# ---------------------------------------------------------------------------


def test_input_excel_sha_unchanged(production_result):
    sha_after_build = vi.calculate_sha256(PRODUCTION_FILE)
    assert sha_after_build == production_result["sha256"]


# ---------------------------------------------------------------------------
# 22-28. Enfoque hibrido: cards = Julio, paneles = acumulado Ene-Jul
# ---------------------------------------------------------------------------


def test_kpi_cards_unchanged_julio_scope(production_result):
    """Las 5 KPI cards siguen la foto de julio.
    Recalibrado 2026-08-18 (promocion base OCU26+YPF, ver
    docs/CM2_CIERRE_BASE_OCU26_YPF_2026-08-18.md): TV6 incluye YPF en su
    universo. El reemplazo integro del bloque legacy YPF (10000-10009) por
    YPF Etapa 2 -- que usa Marca=Cliente=Agencia=Campaña, convencion ya
    autorizada -- cambia que entidades aparecen como cliente/marca/agencia
    top; 'GCBA' (solo existia atado al bloque legacy retirado) ya no es el
    cliente top, ahora es 'TAGGIFY'."""
    kpis = production_result["data"]["kpis"]
    assert kpis["clientes_directos_activos"] == 10
    assert kpis["marcas_activas"] == 95
    assert kpis["agencias_activas"] == 15
    assert kpis["cliente_top_identificado"] == {"nombre": "TAGGIFY", "activaciones": 7}
    assert kpis["concentracion_top5"]["pct"] == 67.4
    assert kpis["concentracion_top5"]["denominador"] == 11497


def test_universo_exposes_julio_and_hist_separately(production_result):
    # Recalibrado 2026-08-18: 8203 -> 11497 (ver test_kpi_cards_unchanged_julio_scope).
    universo = production_result["data"]["universo"]
    assert set(universo.keys()) >= {"circuitos", "elementos", "julio", "hist"}
    assert universo["julio"]["activaciones_totales"] == 11497
    assert universo["hist"]["activaciones_totales"] > universo["julio"]["activaciones_totales"]
    assert universo["hist"]["periodo_label"] == td.HIST_LABEL


def test_rankings_use_accumulated_hist_scope_not_julio(production_result):
    """El ranking Marcas/Agencias debe reflejar Ene-Jul (HIST), no la foto
    de julio.

    Recalibrado 2026-08-18 (promocion base OCU26+YPF, ver
    docs/CM2_CIERRE_BASE_OCU26_YPF_2026-08-18.md):
    - 'GCBA' ya no es la marca top: solo existia como Cliente del bloque
      legacy YPF (IDCampaña 10000-10009), retirado. YPF Etapa 2 usa
      Marca=Cliente=Agencia=Campaña (convencion ya autorizada); la marca
      top pasa a ser 'SEGURIDAD VIAL'.
    - La antigua aserción `len(nombres) > agencias_activas` dejó de ser un
      invariante válido: comparaba magnitudes de distinto alcance y forma
      (ranking_top de HIST, tamaño fijo RANK_TOP_N, contra la cantidad de
      agencias DISTINTAS de JULIO). YPF Etapa 2 infla agencias_activas de
      julio (cada campaña YPF aporta un nombre de "agencia" propio via la
      convención ya autorizada) sin relación con el tamaño del ranking. Se
      reemplaza por controles verificables e independientes en vez de otra
      desigualdad arbitraria.
    """
    marcas = production_result["data"]["ranking"]["marcas"]
    assert marcas[0] == {"nombre": "SEGURIDAD VIAL", "activaciones": 3065}

    agencias_top = production_result["data"]["ranking"]["agencias"]
    nombres = [a["nombre"] for a in agencias_top]

    # Recalculo independiente de julio/hist (mismas funciones de scope del
    # builder, invocadas aca de forma independiente de las aserciones sobre
    # el payload) para verificar, sin asumir nada, que:
    _tr, semantic_result, engine = td.load_pipeline(PRODUCTION_FILE)
    universe = td.build_tv6_universe(semantic_result)
    start_jul, end_jul = td._period_bounds(td.REPORT_YEAR, td.REPORT_MONTH)
    scope_julio = td.build_tv6_scope(engine, universe["element_ids"], start_jul, end_jul)
    scope_hist = td.build_tv6_scope(engine, universe["element_ids"], td.HIST_START_ISO, td.HIST_END_ISO)
    agencias_julio = td.compute_agencias(scope_julio)
    agencias_hist = td.compute_agencias(scope_hist)

    # 1. agencias_activas (KPI de julio) coincide con la cantidad real de
    # agencias distintas identificadas en julio, recalculada de forma
    # independiente.
    assert agencias_julio["activas"] == production_result["data"]["kpis"]["agencias_activas"] == 15

    # 2. El ranking (HIST) no incluye agencias vacías/no identificadas.
    assert all(n and n not in td.AGENCIA_VALORES_NO_IDENTIFICADOS for n in nombres)

    # 3. Julio y el histórico no están mezclados: el total de activaciones
    # identificadas en HIST es estrictamente mayor al de JULIO (HIST
    # contiene a julio por construcción de fechas, más lo que solo estuvo
    # activo antes de julio) -- prueba de que el ranking no es la foto de
    # julio.
    assert agencias_hist["activaciones_totales"] > agencias_julio["activaciones_totales"]


def test_matriz_row_sums_match_hist_not_julio(production_result):
    """La matriz "Demanda por circuito" se construye sobre el acumulado
    Ene-Jul (HIST). Recalibrado 2026-08-18: 'GCBA' ya no es una fila (se
    retiró con el bloque legacy YPF); se verifica de forma general que las
    filas son exactamente las del ranking HIST (sin filas fantasma) y que
    cada una suma su propio total HIST."""
    marcas_hist = {m["nombre"]: m["activaciones"] for m in production_result["data"]["ranking"]["marcas"]}
    filas = production_result["data"]["matriz"]["filas"]

    # 1. Las entidades de la matriz son exactamente las del ranking HIST
    # (misma fuente y universo temporal): ninguna fila fantasma del bloque legacy.
    assert [f["nombre"] for f in filas] == list(marcas_hist.keys())
    assert "GCBA" not in [f["nombre"] for f in filas]

    # 2. Cada fila suma exactamente el total HIST de esa marca.
    for fila in filas:
        assert sum(fila["valores"]) == marcas_hist[fila["nombre"]]

    # 3. Verificación puntual con cálculo independiente desde el Excel: se
    # confirma primero que "SEGURIDAD VIAL" existe en la base canónica antes
    # de usarla, y se recalcula su total de forma independiente del builder.
    assert "SEGURIDAD VIAL" in marcas_hist
    _tr, semantic_result, engine = td.load_pipeline(PRODUCTION_FILE)
    universe = td.build_tv6_universe(semantic_result)
    scope_hist = td.build_tv6_scope(engine, universe["element_ids"], td.HIST_START_ISO, td.HIST_END_ISO)
    sv_independiente = int((scope_hist["Marca"] == "SEGURIDAD VIAL").sum())
    sv_fila = next(f for f in filas if f["nombre"] == "SEGURIDAD VIAL")
    assert sum(sv_fila["valores"]) == sv_independiente == 3065


def test_programatica_pendientes_preserved_not_dropped_or_assigned():
    """Las activaciones PROGRAMATICA=Si sin agencia identificada NO se
    eliminan del analisis ni se asignan artificialmente: quedan expuestas
    integras en pendientes_sin_agencia (spec Sec.3)."""
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2"), _cencosud_static("C3")]
    _, scope = _build_scope(maestro_rows, [
        _demanda_row("P1", "C1", IDCampaña="P1", Agencia="No", PROGRAMATICA="Si", Cliente="TAGGIFY"),
        _demanda_row("P2", "C2", IDCampaña="P2", Agencia="A confirmar", PROGRAMATICA="Si", Cliente="LATIN AD"),
        _demanda_row("P3", "C3", IDCampaña="P3", Agencia="AG_REAL", PROGRAMATICA="No"),
    ])
    agencias = td.compute_agencias(scope)
    prog = td.compute_programatica(scope, agencias["nombres_identificados"])
    assert prog["activaciones_flag_si_total"] == 2
    assert prog["pendientes_sin_agencia"] == 2  # ninguna se pierde ni se asigna a AG_REAL
    assert prog["ranking_top"] == []
    assert "2" in prog["nota"] and "pendientes de imputación" in prog["nota"]


def test_production_programatica_pendientes_sin_agencia(production_result):
    prog = production_result["data"]["ranking"]["programatica"]
    assert prog["activaciones_flag_si_total"] == 112
    assert prog["pendientes_sin_agencia"] == 112  # ninguna coincide con agencia identificada Ene-Jul
    assert prog["ranking_top"] == []


def test_insights_footer_is_triple_pattern(production_result):
    insights = production_result["data"]["insights"]
    assert set(insights.keys()) == {"lectura", "punto_positivo", "a_atender"}
    for key in insights:
        assert insights[key]  # no vacio


def test_a_atender_prioritizes_programatica_pendiente(production_result):
    """Prioridad 1 (spec Sec.6): activaciones programaticas sin agencia
    imputada, dado que Ene-Jul tiene 112 pendientes."""
    assert "programátic" in production_result["data"]["insights"]["a_atender"].lower()


def test_template_has_triple_insight_footer_markup():
    template = (REPO_ROOT / "scripts" / "templates" / "tv6_template.html").read_text(encoding="utf-8")
    assert "data-insight-lectura" in template
    assert "data-insight-positivo" in template
    assert "data-insight-atender" in template
    assert "insight-col" in template


def test_output_html_has_triple_insight_footer_rendered():
    td.build_and_write(PRODUCTION_FILE)
    html = td.DEFAULT_OUTPUT_HTML.read_text(encoding="utf-8")
    assert "Lectura</div>" in html
    assert "Punto positivo</div>" in html
    assert "A atender</div>" in html


# ---------------------------------------------------------------------------
# 22-26 (numeracion original). TV1/TV2/TV3/TV4/TV6_REFERENCE intactas
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


def test_tv4_pipeline_still_builds_successfully():
    import build_tv4_dashboard as t4
    result = t4.build_tv4_data(PRODUCTION_FILE)
    # Recalibrado 2026-08-18: 77 -> 82. TV4_CIRCUITOS incluye CENCOSUD; las
    # filas CENCOSUD (KFC x3 + MISHKA) ya autorizadas en FINAL_V2 y nunca
    # antes promovidas explican el delta (misma causa que TV3).
    assert result["data"]["kpis"]["actividad_actual"]["campanas_unicas"] == 82


def test_tv6_reference_untouched():
    ref = REPO_ROOT / "audit_sources" / "TV6_REFERENCE.html.html"
    assert ref.exists()
    html = ref.read_text(encoding="utf-8")
    assert "window.OCU_DATA" in html  # dataset legacy original, nunca reescrito
    assert "TV5" in html or "tv5" in html  # legacy numbering original, intacto


def test_production_excel_untouched_by_import_path():
    assert PRODUCTION_FILE.exists()
