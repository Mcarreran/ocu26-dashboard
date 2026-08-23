"""Pruebas de negocio para scripts/build_tv5_dashboard.py (dashboard TV5
OCU26, Pipeline Comercial = Core Comercial en espacios, union TV2 Digital +
TV3 Estatico).

No modifica scripts/validate_input.py, scripts/transform_data.py,
scripts/semantic_model.py, scripts/metrics_engine.py,
config/business_semantics.json, input/OCU26_BASE_DATOS.xlsx, ni ningun
archivo productivo de TV1/TV2/TV3/TV4/TV6 (build_tv1/2/3/4/6_dashboard.py,
tv1/2/3/4/6_template.html, tv1/2/3/4/6.html, test_build_tv1/2/3/4/6_dashboard.py,
audit_sources/*). Los fixtures sinteticos usan la CONFIGURACION REAL
(sm.load_config()), mismo patron que test_build_tv1_dashboard.py.

IMPORTANTE: ninguna de estas pruebas invoca build_and_write() de otro TV ni
scripts/build_tv1..4/6_dashboard.py como subproceso: solo build_tvX_data()
(no escribe archivos), para no regenerar timestamps de otros tableros.
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
import build_tv5_dashboard as td  # noqa: E402
import build_tv1_dashboard as td1  # noqa: E402
from export_data import load_pipeline  # noqa: E402
from metrics_engine import MetricsEngine  # noqa: E402
from test_semantic_model import _maestro_row, _campana_row, _transform_result  # noqa: E402
from test_build_tv1_dashboard import (  # noqa: E402
    _static_cencosud, _digital_cencosud, _apsa_static, _london_static,
    _ypf_digital, _pantalla_led,
)
from test_build_tv3_dashboard import _aa2000_static, _pilar_frontlight, _cencomedia, _mab_static  # noqa: E402

PRODUCTION_FILE = REPO_ROOT / "input" / "OCU26_BASE_DATOS.xlsx"
BUILDER_SOURCE = (REPO_ROOT / "scripts" / "build_tv5_dashboard.py").read_text(encoding="utf-8")
TEMPLATE_SOURCE = (REPO_ROOT / "scripts" / "templates" / "tv5_template.html").read_text(encoding="utf-8")

CUTOFF = pd.Timestamp("2026-07-31")


def _semantic(maestro_rows: list[dict], campanas_rows: list[dict] | None = None) -> dict:
    config = sm.load_config()
    return sm.build_semantic_model(_transform_result(maestro_rows, campanas_rows or []), config)


def _digital_aa2000(elemento_id: str, ubicacion: str = "EZE", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="AA2000", Subcircuito="X", Ubicacion=ubicacion,
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        CapacidadSlotsReel=8, SegundosDia=72000,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


# ---------------------------------------------------------------------------
# Fixtures de produccion (Excel real)
# ---------------------------------------------------------------------------


@pytest.fixture(scope="module")
def production_result():
    return td.build_tv5_data(PRODUCTION_FILE)


@pytest.fixture(scope="module")
def production_json(production_result):
    return json.dumps(production_result["data"], ensure_ascii=False)


@pytest.fixture(scope="module")
def production_html(production_result):
    return td.render_html(production_result["data"])


# ---------------------------------------------------------------------------
# 1. Badge visible TV5 / sin referencias a TV4 en el tablero (test 1, 3)
# ---------------------------------------------------------------------------


def test_template_badge_is_tv5():
    assert '<span class="tvno">TV 5</span>' in TEMPLATE_SOURCE
    assert "Pipeline Comercial" in TEMPLATE_SOURCE


def test_rendered_html_has_no_visible_tv4_reference(production_html):
    # El unico "TV4" aceptable es dentro del nombre interno de variable JS
    # (window.TV5_DATA no lo contiene); se verifica que no aparezca el
    # rotulo visible "TV 4" ni "TV4" en el marcado.
    assert "TV 4" not in production_html
    assert ">TV4<" not in production_html
    assert '<span class="tvno">TV 5</span>' in production_html


def test_default_output_paths_target_only_tv5():
    assert td.DEFAULT_OUTPUT_HTML.name == "tv5.html"
    assert td.DEFAULT_OUTPUT_JSON.name == "tv5_data.json"


def test_tv5_builder_does_not_import_tv4_builder():
    # build_tv4_dashboard.py ya importa de build_tv5_dashboard.py (geo YPF):
    # un import en sentido inverso crearia un ciclo. Se verifica que el
    # codigo fuente de TV5 nunca tiene una sentencia de import real hacia
    # ese modulo (el nombre puede aparecer en comentarios/docstrings
    # explicando por que no se importa).
    assert "import build_tv4_dashboard" not in BUILDER_SOURCE
    assert "from build_tv4_dashboard" not in BUILDER_SOURCE


# ---------------------------------------------------------------------------
# 2. Payload contiene solo TV5
# ---------------------------------------------------------------------------


def test_payload_top_level_keys_are_tv5_only(production_result):
    assert set(production_result["data"].keys()) == {
        "meta", "kpis", "estado_pipeline", "distribucion_grupo", "timeline", "calidad", "insights",
    }


def test_payload_contains_no_other_tv_datasets(production_json):
    for token in ("tv1_data", "tv2_data", "tv3_data", "tv4_data", "tv6_data", "ocu_data"):
        assert token not in production_json.lower()


# ---------------------------------------------------------------------------
# 3-6. Campanas / activaciones / duplicados (pedido Sec.5, tests 4-6)
# ---------------------------------------------------------------------------


def _two_digital_elements_one_campaign_each():
    maestro_rows = [_digital_cencosud("D1"), _digital_cencosud("D2")]
    campana_rows = [
        _campana_row("C1", "D1", IDCampaña="X1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
        _campana_row("C2", "D2", IDCampaña="X2", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
    ]
    return maestro_rows, campana_rows


def test_campanas_unicas_cuenta_idcampana_distintos():
    maestro_rows, campana_rows = _two_digital_elements_one_campaign_each()
    semantic_result = _semantic(maestro_rows, campana_rows)
    universe = td.build_tv5_universe(semantic_result)
    windows = td.compute_pipeline_windows(universe["scope_campanas"], CUTOFF)
    snap = td._espacios_snapshot(windows["en_curso"], universe["digital_ids"], universe["static_ids"])
    assert snap["campanas_unicas"] == 2


def test_activaciones_son_pares_idcampana_elemento():
    maestro_rows, campana_rows = _two_digital_elements_one_campaign_each()
    semantic_result = _semantic(maestro_rows, campana_rows)
    universe = td.build_tv5_universe(semantic_result)
    windows = td.compute_pipeline_windows(universe["scope_campanas"], CUTOFF)
    snap = td._espacios_snapshot(windows["en_curso"], universe["digital_ids"], universe["static_ids"])
    assert snap["activaciones"] == 2


def test_duplicados_no_inflan_espacios_digitales():
    maestro_rows = [_digital_cencosud("D1")]
    campana_rows = [
        _campana_row("C1", "D1", IDCampaña="X1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
        _campana_row("C1_dup", "D1", IDCampaña="X1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    universe = td.build_tv5_universe(semantic_result)
    windows = td.compute_pipeline_windows(universe["scope_campanas"], CUTOFF)
    snap = td._espacios_snapshot(windows["en_curso"], universe["digital_ids"], universe["static_ids"])
    assert snap["espacios"] == 1
    assert snap["activaciones"] == 1


# ---------------------------------------------------------------------------
# 7,10. Digital: slots canonicos TV2 (pares IDCampaña x ElementoID)
# ---------------------------------------------------------------------------


def test_digital_espacio_es_par_campana_elemento():
    maestro_rows, campana_rows = _two_digital_elements_one_campaign_each()
    semantic_result = _semantic(maestro_rows, campana_rows)
    universe = td.build_tv5_universe(semantic_result)
    windows = td.compute_pipeline_windows(universe["scope_campanas"], CUTOFF)
    snap = td._espacios_snapshot(windows["en_curso"], universe["digital_ids"], universe["static_ids"])
    assert snap["espacios_digitales"] == 2
    assert snap["espacios_estaticos"] == 0


def test_una_campana_digital_en_varios_elementos_produce_varios_espacios():
    maestro_rows = [_digital_cencosud("D1"), _digital_cencosud("D2"), _digital_cencosud("D3")]
    campana_rows = [
        _campana_row("C1", "D1", IDCampaña="X1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
        _campana_row("C2", "D2", IDCampaña="X1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
        _campana_row("C3", "D3", IDCampaña="X1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    universe = td.build_tv5_universe(semantic_result)
    windows = td.compute_pipeline_windows(universe["scope_campanas"], CUTOFF)
    snap = td._espacios_snapshot(windows["en_curso"], universe["digital_ids"], universe["static_ids"])
    assert snap["espacios_digitales"] == 3
    assert snap["campanas_unicas"] == 1


# ---------------------------------------------------------------------------
# 8,9. Estatico: ElementoID distintos (TV3) -- dos campanas en el mismo
# elemento producen 1 espacio ocupado, no 2.
# ---------------------------------------------------------------------------


def test_dos_campanas_activas_en_mismo_elemento_estatico_producen_un_espacio():
    maestro_rows = [_static_cencosud("S1")]
    campana_rows = [
        _campana_row("C1", "S1", IDCampaña="X1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
        _campana_row("C2", "S1", IDCampaña="X2", FechaInicio=pd.Timestamp("2026-07-20"), FechaFin=pd.Timestamp("2026-08-15")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    universe = td.build_tv5_universe(semantic_result)
    windows = td.compute_pipeline_windows(universe["scope_campanas"], CUTOFF)
    snap = td._espacios_snapshot(windows["en_curso"], universe["digital_ids"], universe["static_ids"])
    assert snap["espacios_estaticos"] == 1
    assert snap["campanas_unicas"] == 2
    assert snap["activaciones"] == 2


# ---------------------------------------------------------------------------
# 11. Espacios ocupados Core = suma de los cuatro grupos
# ---------------------------------------------------------------------------


def test_distribucion_grupo_reconcilia_con_core(production_result):
    kpis = production_result["data"]["kpis"]
    distribucion = production_result["data"]["distribucion_grupo"]
    assert {g["grupo"] for g in distribucion} == set(td.GRUPOS_CORE)
    assert sum(g["espacios"] for g in distribucion) == kpis["actividad_actual"]["espacios"]


def test_distribucion_grupo_synthetic_reconcilia():
    maestro_rows = [
        _digital_cencosud("SD1"), _static_cencosud("SE1"),
        _pantalla_led("PL1"), _digital_aa2000("AAD1"), _aa2000_static("AAE1"), _pilar_frontlight("PF1"),
    ]
    campana_rows = [
        _campana_row("C1", eid, IDCampaña=f"X{i}", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31"))
        for i, eid in enumerate(["SD1", "SE1", "PL1", "AAD1", "AAE1", "PF1"])
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    universe = td.build_tv5_universe(semantic_result)
    windows = td.compute_pipeline_windows(universe["scope_campanas"], CUTOFF)
    en_curso = windows["en_curso"]
    snap = td._espacios_snapshot(en_curso, universe["digital_ids"], universe["static_ids"])
    distribucion = td.compute_distribucion_grupo(en_curso, universe["digital_ids"], universe["static_ids"], universe["grupo_map"])
    assert sum(g["espacios"] for g in distribucion) == snap["espacios"] == 6
    by_grupo = {g["grupo"]: g["espacios"] for g in distribucion}
    assert by_grupo["Shoppings Digital"] == 1
    assert by_grupo["Shoppings Estático"] == 1
    assert by_grupo["Pantallas LED"] == 1
    assert by_grupo["AA2000 / Pilar Frontlight"] == 3


# ---------------------------------------------------------------------------
# 12-14. Ventanas temporales (reservas / inician / finalizan 30 dias)
# ---------------------------------------------------------------------------


def test_reservas_futuras_cuenta_todos_los_espacios_de_una_campana():
    maestro_rows = [_digital_cencosud("D1"), _digital_cencosud("D2")]
    campana_rows = [
        _campana_row("C1", "D1", IDCampaña="FUT1", FechaInicio=pd.Timestamp("2026-09-01"), FechaFin=pd.Timestamp("2026-09-30")),
        _campana_row("C2", "D2", IDCampaña="FUT1", FechaInicio=pd.Timestamp("2026-09-01"), FechaFin=pd.Timestamp("2026-09-30")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    universe = td.build_tv5_universe(semantic_result)
    windows = td.compute_pipeline_windows(universe["scope_campanas"], CUTOFF)
    reservas = td._espacios_evento(windows["futuras"])
    assert reservas["espacios"] == 2
    assert reservas["campanas_unicas"] == 1


def test_inician_30d_respeta_ventana_temporal():
    maestro_rows = [_digital_cencosud("D1"), _digital_cencosud("D2")]
    campana_rows = [
        _campana_row("C1", "D1", IDCampaña="IN29", FechaInicio=CUTOFF + pd.Timedelta(days=29), FechaFin=pd.Timestamp("2026-12-31")),
        _campana_row("C2", "D2", IDCampaña="IN31", FechaInicio=CUTOFF + pd.Timedelta(days=31), FechaFin=pd.Timestamp("2026-12-31")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    universe = td.build_tv5_universe(semantic_result)
    windows = td.compute_pipeline_windows(universe["scope_campanas"], CUTOFF)
    inician = td._espacios_evento(windows["inician_30d"])
    assert inician["espacios"] == 1
    assert inician["campanas_unicas"] == 1


def test_finalizan_30d_respeta_ventana_temporal():
    maestro_rows = [_digital_cencosud("D1"), _digital_cencosud("D2")]
    campana_rows = [
        _campana_row("C1", "D1", IDCampaña="FIN29", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=CUTOFF + pd.Timedelta(days=29)),
        _campana_row("C2", "D2", IDCampaña="FIN31", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=CUTOFF + pd.Timedelta(days=31)),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    universe = td.build_tv5_universe(semantic_result)
    windows = td.compute_pipeline_windows(universe["scope_campanas"], CUTOFF)
    finalizan = td._espacios_evento(windows["finalizan_30d"])
    assert finalizan["espacios"] == 1
    assert finalizan["campanas_unicas"] == 1


# ---------------------------------------------------------------------------
# 15. Finalizados historicos no entran en ocupacion actual
# ---------------------------------------------------------------------------


def test_finalizados_historicos_no_cuentan_como_actividad_actual():
    maestro_rows = [_digital_cencosud("D1")]
    campana_rows = [
        _campana_row("C1", "D1", IDCampaña="HIST1", FechaInicio=pd.Timestamp("2026-01-01"), FechaFin=pd.Timestamp("2026-06-01")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    universe = td.build_tv5_universe(semantic_result)
    windows = td.compute_pipeline_windows(universe["scope_campanas"], CUTOFF)
    snap = td._espacios_snapshot(windows["en_curso"], universe["digital_ids"], universe["static_ids"])
    assert snap["espacios"] == 0
    historico = td._espacios_evento(windows["historico"])
    assert historico["espacios"] == 1


# ---------------------------------------------------------------------------
# 16. Pipeline en Core usa espacios (no campanas) como metrica principal
# ---------------------------------------------------------------------------


def test_pipeline_core_pct_se_calcula_sobre_espacios(production_result):
    core = production_result["data"]["kpis"]["pipeline_core"]
    assert core["espacios_general"] > 0
    esperado = round(core["espacios_core"] / core["espacios_general"] * 100.0, 1)
    assert core["pct"] == esperado
    # No debe coincidir con el calculo legacy "82 de 93 campañas" como
    # porcentaje principal: el pct reportado usa espacios, no campanas.
    pct_legacy_campanas = round(core["campanas_core"] / core["campanas_general"] * 100.0, 1)
    if core["campanas_general"] and core["espacios_general"] != core["campanas_general"]:
        assert core["pct"] != pct_legacy_campanas or core["espacios_core"] == core["campanas_core"]


# ---------------------------------------------------------------------------
# 17. Circuitos sin regla de espacio no entran como cero
# ---------------------------------------------------------------------------


def test_circuitos_sin_regla_de_espacio_se_reportan_pendientes_no_cero():
    # APSA tiene incluye_conteo_general=false (business_semantics.json): ya
    # queda fuera del universo OPERATIVO_GENERAL aguas arriba (semantic_model
    # filter_universe), antes de llegar a TV5 -- por eso no se espera en
    # "pendientes" (nunca llega a existir en `op`). LONDON_SUPPLY/CENCOMEDIA/
    # MAB si son visibles en OPERATIVO_GENERAL (incluye_conteo_general=true)
    # pero no tienen regla de espacio Core ni YPF: esos SI deben listarse.
    maestro_rows = [_digital_cencosud("D1"), _apsa_static("A1"), _london_static("L1"), _cencomedia("CM1"), _mab_static("M1")]
    campana_rows = [
        _campana_row("C1", "D1", IDCampaña="X1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    universe = td.build_tv5_universe(semantic_result)
    core_circuitos = set(universe["circuitos_digital"]) | set(universe["circuitos_estatico"])
    pendientes = td.compute_circuitos_pendientes(universe["op"], core_circuitos)
    pendientes_nombres = {p["circuito"] for p in pendientes}
    assert {"LONDON_SUPPLY", "CENCOMEDIA", "MAB"} <= pendientes_nombres
    assert "APSA" not in set(universe["op"]["CircuitoNegocio"].unique())
    for p in pendientes:
        assert p["elementos_catalogo"] > 0  # nunca cero/inventado: cuenta real de catalogo
    # APSA/London/Cencomedia/MAB nunca entran al universo Core (Digital+Estatico)
    assert universe["digital_ids"].isdisjoint({"A1", "L1", "CM1", "M1"})
    assert universe["static_ids"].isdisjoint({"A1", "L1", "CM1", "M1"})


def test_production_pipeline_core_lista_circuitos_pendientes(production_result):
    pendientes = production_result["data"]["calidad"]["circuitos_pendientes"]
    assert isinstance(pendientes, list)
    for p in pendientes:
        assert p["elementos_catalogo"] > 0


# ---------------------------------------------------------------------------
# 18. Proximos inicios no estan hardcodeados
# ---------------------------------------------------------------------------


def test_timeline_no_hardcodea_nombres_de_campana():
    for token in ("Netflix", "Nintendo", "Oak Street", "PlayStation"):
        assert token not in BUILDER_SOURCE
        assert token not in TEMPLATE_SOURCE


def test_timeline_refleja_datos_sinteticos_no_hardcodeados():
    maestro_rows = [_digital_cencosud("D1")]
    campana_rows = [
        _campana_row("C1", "D1", IDCampaña="ACME1", Campaña="Campaña Acme de Prueba",
                      FechaInicio=CUTOFF + pd.Timedelta(days=5), FechaFin=pd.Timestamp("2026-12-31")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    universe = td.build_tv5_universe(semantic_result)
    windows = td.compute_pipeline_windows(universe["scope_campanas"], CUTOFF)
    rows, total = td.compute_proximos_inicios(windows["inician_30d"], universe["grupo_map"])
    assert total == 1
    assert rows[0]["campana"] == "Campaña Acme de Prueba"
    assert rows[0]["espacios"] == 1
    assert rows[0]["badge"] == "Inicio"


def test_timeline_ordena_por_fecha_luego_por_nombre():
    maestro_rows = [_digital_cencosud("D1"), _digital_cencosud("D2"), _digital_cencosud("D3")]
    campana_rows = [
        _campana_row("C1", "D1", IDCampaña="B", Campaña="Beta", FechaInicio=CUTOFF + pd.Timedelta(days=10), FechaFin=pd.Timestamp("2026-12-31")),
        _campana_row("C2", "D2", IDCampaña="A", Campaña="Alfa", FechaInicio=CUTOFF + pd.Timedelta(days=5), FechaFin=pd.Timestamp("2026-12-31")),
        _campana_row("C3", "D3", IDCampaña="C", Campaña="Charlie", FechaInicio=CUTOFF + pd.Timedelta(days=5), FechaFin=pd.Timestamp("2026-12-31")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    universe = td.build_tv5_universe(semantic_result)
    windows = td.compute_pipeline_windows(universe["scope_campanas"], CUTOFF)
    rows, _total = td.compute_proximos_inicios(windows["inician_30d"], universe["grupo_map"])
    assert [r["campana"] for r in rows] == ["Alfa", "Charlie", "Beta"]


# ---------------------------------------------------------------------------
# 19-20. Insights usan espacios como jerarquia; activaciones nunca son la
# unidad principal de ocupacion.
# ---------------------------------------------------------------------------


def test_insights_lectura_usa_espacios_como_jerarquia_principal(production_result):
    lectura = production_result["data"]["insights"]["lectura"]
    assert "espacios comerciales ocupados" in lectura
    idx_espacios = lectura.find("espacios comerciales ocupados")
    idx_campanas = lectura.find("campañas activas")
    assert idx_espacios != -1 and idx_campanas != -1 and idx_espacios < idx_campanas


def test_kpis_principales_son_espacios_no_activaciones(production_result):
    kpis = production_result["data"]["kpis"]
    assert "espacios" in kpis["actividad_actual"]
    assert "espacios" in kpis["reservas_futuras"]
    assert "espacios" in kpis["inician_30d"]
    assert "espacios" in kpis["finalizan_30d"]
    # activaciones existe solo como referencia menor en la tarjeta 1 (payload)
    assert "activaciones" not in kpis["reservas_futuras"]
    assert "activaciones" not in kpis["inician_30d"]
    assert "activaciones" not in kpis["finalizan_30d"]


def test_legacy_activaciones_448_no_se_renombra_como_espacios():
    # Guardia textual: el valor legacy no debe aparecer renombrado sin
    # recalcular (pedido Sec.5).
    assert "448 espacios" not in BUILDER_SOURCE
    assert "448 espacios" not in TEMPLATE_SOURCE


# ---------------------------------------------------------------------------
# 21. HTML 1920x1080 sin scroll ni overflow
# ---------------------------------------------------------------------------


def test_template_stage_is_1920x1080_no_scroll():
    assert "width:1920px;height:1080px" in TEMPLATE_SOURCE.replace(" ", "").replace("\n", "")
    assert "overflow:hidden" in TEMPLATE_SOURCE


def test_rendered_html_is_well_formed(production_html):
    assert production_html.count("<html") == 1
    assert production_html.count("</html>") == 1
    assert "{{TV5_DATA_JSON}}" not in production_html
    assert "{{LOGO_IMG_TAG}}" not in production_html


# ---------------------------------------------------------------------------
# 22. No se modifican timestamps de otras TVs: build_tv5_data() no toca
# archivos de disco (solo build_and_write lo hace, y solo sobre tv5.*).
# ---------------------------------------------------------------------------


def test_build_tv5_data_does_not_write_other_tv_outputs(tmp_path):
    before = {
        p.name: p.stat().st_mtime
        for p in (REPO_ROOT / "output").glob("tv*_data.json")
        if p.name != "tv5_data.json"
    }
    td.build_tv5_data(PRODUCTION_FILE)
    after = {
        p.name: p.stat().st_mtime
        for p in (REPO_ROOT / "output").glob("tv*_data.json")
        if p.name != "tv5_data.json"
    }
    assert before == after


# ---------------------------------------------------------------------------
# Universo Core = union TV2 (Digital) + TV3 (Estatico); YPF fuera del Core
# ---------------------------------------------------------------------------


def test_universo_core_excluye_ypf():
    maestro_rows = [_digital_cencosud("D1"), _ypf_digital("500 - TT - 1", "500 - TESTVILLE - Calle Falsa 123")]
    campana_rows = [
        _campana_row("C1", "D1", IDCampaña="X1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    universe = td.build_tv5_universe(semantic_result)
    assert "500 - TT - 1" not in universe["digital_ids"]
    assert "D1" in universe["digital_ids"]


def test_universo_core_es_union_tv2_digital_tv3_estatico():
    from build_tv2_dashboard import TV2_CIRCUITOS
    from build_tv3_dashboard import TV3_CIRCUITOS
    assert set(td.CORE_DIGITAL_CIRCUITOS) == set(TV2_CIRCUITOS)
    assert set(td.CORE_ESTATICO_CIRCUITOS) == set(TV3_CIRCUITOS)


# ---------------------------------------------------------------------------
# Auditoria TV1 (A, mensual) vs TV5 (B/C, al corte) -- 2026-08-23.
# Reconciliacion: A=133, B=93, C=82, A-B=40 (todas finalizaron antes del
# corte), B-A=0, B-C=11 (todas YPF), C-B=0. Ver informe de entrega para el
# detalle campana-a-campana.
# ---------------------------------------------------------------------------


def _set_a_tv1_julio(semantic_result: dict, engine) -> set[str]:
    """Set A: campanas TV1 (universo general, excluye solo APSA/London/
    Cencomedia) con solapamiento en julio 2026 completo."""
    tv1_universe = td1.build_tv1_universe(semantic_result)
    overlap = engine._campanas_overlap(tv1_universe["element_ids"], "2026-07-01", "2026-07-31")
    overlap = overlap[overlap["IDCampaña"].notna() & (overlap["IDCampaña"].astype(str).str.strip() != "")]
    return set(overlap["IDCampaña"].astype(str).str.strip().unique())


def _set_bc_tv5_corte(semantic_result: dict) -> tuple[set[str], set[str]]:
    """Sets B (general = Core+YPF) y C (Core) al corte, mismo criterio que
    kpis.pipeline_core del builder productivo."""
    universe = td.build_tv5_universe(semantic_result)
    windows = td.compute_pipeline_windows(universe["scope_campanas"], CUTOFF)
    en_curso_core = windows["en_curso"]
    set_c = set(en_curso_core["IDCampaña"].astype(str).str.strip().unique())
    ypf_snap = td.build_tv5_ypf_snapshot(semantic_result, CUTOFF)
    set_b = set_c | {str(x).strip() for x in ypf_snap["campanas_ids"]}
    return set_b, set_c


@pytest.fixture(scope="module")
def audit_sets_production():
    _transform_result, semantic_result, engine = load_pipeline(PRODUCTION_FILE)
    set_a = _set_a_tv1_julio(semantic_result, engine)
    set_b, set_c = _set_bc_tv5_corte(semantic_result)
    return {"A": set_a, "B": set_b, "C": set_c}


def test_reglas_temporales_tv1_mensual_vs_tv5_al_corte_son_distintas():
    # Una campana activa DURANTE julio pero finalizada antes del corte debe
    # aparecer en el conjunto mensual (TV1-like) y NO en el conjunto al
    # corte (TV5): las reglas temporales son deliberadamente distintas.
    maestro_rows = [_digital_cencosud("D1")]
    campana_rows = [
        _campana_row("C1", "D1", IDCampaña="MIDJUL", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-20")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    overlap_mensual = engine._campanas_overlap(["D1"], "2026-07-01", "2026-07-31")
    overlap_mensual = overlap_mensual[overlap_mensual["IDCampaña"].notna()]
    assert "MIDJUL" in set(overlap_mensual["IDCampaña"].astype(str).str.strip())

    universe = td.build_tv5_universe(semantic_result)
    windows = td.compute_pipeline_windows(universe["scope_campanas"], CUTOFF)
    assert "MIDJUL" not in set(windows["en_curso"]["IDCampaña"].astype(str).str.strip())


def test_b_es_subconjunto_de_a_en_produccion(audit_sets_production):
    sets = audit_sets_production
    assert sets["B"] <= sets["A"]
    assert sets["B"] - sets["A"] == set()


def test_c_es_subconjunto_de_b_en_produccion(audit_sets_production):
    sets = audit_sets_production
    assert sets["C"] <= sets["B"]
    assert sets["C"] - sets["B"] == set()


def test_reconciliacion_produccion_coincide_con_auditoria(audit_sets_production):
    sets = audit_sets_production
    assert len(sets["A"]) == 133
    assert len(sets["B"]) == 93
    assert len(sets["C"]) == 82
    assert len(sets["A"] - sets["B"]) == 40
    assert len(sets["B"] - sets["C"]) == 11


def test_campanas_se_cuentan_por_idcampana_no_por_fila(audit_sets_production):
    # Set A/B/C ya son sets de IDCampaña normalizado: verificamos que el
    # tamano del set es estrictamente menor o igual a la cantidad de filas
    # crudas (nunca se cuentan filas duplicadas como campanas distintas).
    _transform_result, semantic_result, engine = load_pipeline(PRODUCTION_FILE)
    tv1_universe = td1.build_tv1_universe(semantic_result)
    overlap = engine._campanas_overlap(tv1_universe["element_ids"], "2026-07-01", "2026-07-31")
    overlap = overlap[overlap["IDCampaña"].notna() & (overlap["IDCampaña"].astype(str).str.strip() != "")]
    assert len(overlap) >= overlap["IDCampaña"].nunique()
    assert overlap["IDCampaña"].nunique() == 133


def test_duplicados_no_inflan_conjunto_de_campanas_general():
    maestro_rows = [_digital_cencosud("D1"), _digital_cencosud("D2")]
    campana_rows = [
        _campana_row("C1", "D1", IDCampaña="X1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
        _campana_row("C1_dup", "D1", IDCampaña="X1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
        _campana_row("C2", "D2", IDCampaña="X1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-31")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    set_b, set_c = _set_bc_tv5_corte(semantic_result)
    assert set_c == {"X1"}
    assert set_b == {"X1"}


def test_etiquetas_visibles_incluyen_core_al_corte_y_fecha():
    combined = BUILDER_SOURCE + TEMPLATE_SOURCE
    assert "Core" in combined
    assert "al corte" in combined
    assert "31/07" in combined
    # Verificacion puntual del pedido Sec.4: label secundaria de Tarjeta 1
    assert "Campa&ntilde;as activas del Core al corte" in TEMPLATE_SOURCE
    assert "campa&ntilde;as activas del universo general al 31/07" in TEMPLATE_SOURCE
    assert "campa&ntilde;as activas al corte" in TEMPLATE_SOURCE


def test_footer_lectura_usa_frase_precisa_core_vs_general(production_result):
    lectura = production_result["data"]["insights"]["lectura"]
    assert "31/07" in lectura
    assert "campañas activas dentro del Core" in lectura
    assert "campañas activas del universo general" in lectura


def test_tv1_no_se_modifica_por_esta_auditoria():
    # Esta suite de auditoria solo usa funciones puras de lectura de TV1
    # (build_tv1_universe, engine._campanas_overlap): verificamos por mtime
    # real que tv1.html no se toca al ejecutar la auditoria.
    before = (REPO_ROOT / "tv1.html").stat().st_mtime
    _transform_result, semantic_result, _engine = load_pipeline(PRODUCTION_FILE)
    td1.build_tv1_universe(semantic_result)  # funcion pura, no escribe archivos
    after = (REPO_ROOT / "tv1.html").stat().st_mtime
    assert before == after


# ---------------------------------------------------------------------------
# Correccion visual 2026-08-23: la pregunta ejecutiva del header quedaba
# detras de la tarjeta "Pipeline en Core" (.hdr con height fijo mas chico
# que su contenido real -- el contenido se pintaba por fuera de la caja sin
# que .kpis lo supiera, y al ser .kpis posterior en el DOM lo tapaba). Fix:
# .hdr paso de height fijo a min-height (deja que la caja crezca con su
# contenido real), mas un recorte de paddings/margenes (.tv, .kpis, .body,
# .insight) para compensar el alto extra del header sin generar overflow ni
# recortar contenido de los paneles. Este test mide geometria RENDERIZADA
# real (Selenium + Chrome headless), no solo el CSS fuente: es la unica
# forma de detectar una interseccion visual real.
# ---------------------------------------------------------------------------

selenium = pytest.importorskip("selenium", reason="selenium no instalado: test de geometria renderizada se omite")
from selenium import webdriver  # noqa: E402
from selenium.webdriver.chrome.options import Options  # noqa: E402

_CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
]


def _find_chrome_binary() -> str | None:
    for candidate in _CHROME_CANDIDATES:
        if Path(candidate).exists():
            return candidate
    return None


@pytest.fixture(scope="module")
def rendered_geometry(production_html, tmp_path_factory):
    """Renderiza el HTML productivo de TV5 en Chrome headless a 1920x1080
    real (Emulation.setDeviceMetricsOverride, no el tamano de ventana --
    evita el offset de chrome/DPI del SO) y devuelve los bounding rects
    reales de header/pregunta/KPIs/paneles/footer, mas flags de overflow."""
    binary = _find_chrome_binary()
    if binary is None:
        pytest.skip("No se encontro un binario de Chrome/Edge para renderizar TV5")

    html_path = tmp_path_factory.mktemp("tv5_render") / "tv5.html"
    html_path.write_text(production_html, encoding="utf-8")

    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--force-device-scale-factor=1")
    opts.binary_location = binary
    driver = webdriver.Chrome(options=opts)
    try:
        driver.execute_cdp_cmd("Emulation.setDeviceMetricsOverride", {
            "width": 1920, "height": 1080, "deviceScaleFactor": 1, "mobile": False,
        })
        driver.get(html_path.as_uri())

        def rect(selector: str) -> dict:
            el = driver.find_element("css selector", selector)
            return driver.execute_script(
                "var r = arguments[0].getBoundingClientRect();"
                "return {top:r.top, bottom:r.bottom, left:r.left, right:r.right, height:r.height, width:r.width};",
                el,
            )

        rects = {
            sel: rect(sel)
            for sel in (".hdr", ".hdr-right", ".question", ".kpis", ".body",
                        ".panel:nth-of-type(1) .panel-body", ".panel:nth-of-type(2) .panel-body",
                        ".insight", "#hint")
        }
        page = driver.execute_script(
            "return {sw: document.body.scrollWidth, sh: document.body.scrollHeight, "
            "iw: window.innerWidth, ih: window.innerHeight};"
        )
        clip = driver.execute_script(
            "var pipe = document.querySelector('[data-pipe]');"
            "var tl = document.querySelector('[data-timeline]');"
            "var pb = document.querySelectorAll('.panel-body');"
            "return {pipe_scrollH: pipe.scrollHeight, pb0_clientH: pb[0].clientHeight,"
            "tl_scrollH: tl.scrollHeight, pb1_clientH: pb[1].clientHeight};"
        )
        return {"rects": rects, "page": page, "clip": clip}
    finally:
        driver.quit()


def _rects_intersect(a: dict, b: dict) -> bool:
    return a["left"] < b["right"] and a["right"] > b["left"] and a["top"] < b["bottom"] and a["bottom"] > b["top"]


def test_pregunta_header_no_intersecta_fila_kpis(rendered_geometry):
    r = rendered_geometry["rects"]
    assert not _rects_intersect(r[".question"], r[".kpis"]), (
        f"La pregunta del header {r['.question']} intersecta la fila de KPIs {r['.kpis']}"
    )
    assert not _rects_intersect(r[".hdr"], r[".kpis"]), (
        f"El header {r['.hdr']} intersecta la fila de KPIs {r['.kpis']}"
    )


def test_gap_pregunta_a_kpis_es_al_menos_12px(rendered_geometry):
    r = rendered_geometry["rects"]
    gap = r[".kpis"]["top"] - r[".question"]["bottom"]
    assert gap >= 12, f"Separacion pregunta->KPIs de {gap}px, se requiere >=12px"


def test_pregunta_contenida_dentro_del_header(rendered_geometry):
    r = rendered_geometry["rects"]
    # La pregunta debe quedar dentro de la caja real del header (nunca
    # sobresalir por debajo de .hdr, que es lo que causaba el solapamiento).
    assert r[".question"]["bottom"] <= r[".hdr"]["bottom"] + 0.5
    assert r[".question"]["top"] >= r[".hdr"]["top"] - 0.5
    assert r[".question"]["right"] <= r[".hdr-right"]["right"] + 0.5


def test_paneles_no_recortan_contenido_tras_el_ajuste(rendered_geometry):
    clip = rendered_geometry["clip"]
    assert clip["pipe_scrollH"] <= clip["pb0_clientH"] + 1, (
        f"Panel 'Estado del pipeline' recorta contenido: necesita {clip['pipe_scrollH']}px, "
        f"tiene {clip['pb0_clientH']}px"
    )
    assert clip["tl_scrollH"] <= clip["pb1_clientH"] + 1, (
        f"Panel 'Proximos inicios' recorta contenido: necesita {clip['tl_scrollH']}px, "
        f"tiene {clip['pb1_clientH']}px"
    )


def test_1920x1080_sin_scroll_y_footer_visible(rendered_geometry):
    page = rendered_geometry["page"]
    assert page["sw"] <= 1920 and page["sh"] <= 1080, f"Overflow de pagina detectado: {page}"
    hint = rendered_geometry["rects"]["#hint"]
    assert hint["bottom"] <= 1080, "El footer/hint quedo desplazado fuera de la pantalla 1920x1080"
    insight = rendered_geometry["rects"][".insight"]
    assert insight["bottom"] <= 1080
