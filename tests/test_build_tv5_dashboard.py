"""Pruebas de negocio para scripts/build_tv5_dashboard.py (dashboard TV5
OCU26, YPF).

No modifica scripts/validate_input.py, scripts/transform_data.py,
scripts/semantic_model.py, scripts/metrics_engine.py,
config/business_semantics.json, input/OCU26_BASE_DATOS.xlsx, ni ningun
archivo productivo de TV1/TV2/TV3/TV4/TV6 (build_tv1..4/6_dashboard.py,
tv1..4/6_template.html, tv1..4/6.html, test_build_tv1..4/6_dashboard.py,
audit_sources/*). Los fixtures sinteticos usan la CONFIGURACION REAL
(sm.load_config()), mismo patron que test_build_tv1_dashboard.py.
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
from metrics_engine import MetricsEngine  # noqa: E402
from test_semantic_model import _maestro_row, _campana_row, _transform_result  # noqa: E402
from test_build_tv1_dashboard import _static_cencosud, _apsa_static, _london_static  # noqa: E402

PRODUCTION_FILE = REPO_ROOT / "input" / "OCU26_BASE_DATOS.xlsx"
BUILDER_SOURCE = (REPO_ROOT / "scripts" / "build_tv5_dashboard.py").read_text(encoding="utf-8")


def _semantic(maestro_rows: list[dict], campanas_rows: list[dict] | None = None) -> dict:
    config = sm.load_config()
    return sm.build_semantic_model(_transform_result(maestro_rows, campanas_rows or []), config)


def _ypf_digital(elemento_id: str, ubicacion: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="YPF Digital", Subcircuito="X", Ubicacion=ubicacion,
        Medio="Digital", TipoCatalogo="Abierto", TipoInventario="Digital",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _one_station_three_elements():
    """Una estacion YPF (mismo prefijo + misma localidad) con 3 ElementoID
    de formatos distintos (TT, TT, MB), cada uno con una campana propia en
    julio 2026: permite distinguir estaciones (1) de elementos (3) de
    activaciones (3) de campanas unicas (3)."""
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
# 1. Payload contiene solo TV5
# ---------------------------------------------------------------------------


def test_payload_top_level_keys_are_tv5_only(production_result):
    assert set(production_result["data"].keys()) == {
        "meta", "universo", "kpis", "historico", "mapa", "calidad", "insights",
    }


def test_payload_contains_no_other_tv_datasets(production_json):
    for token in ("tv1_data", "tv2_data", "tv3_data", "tv4_data", "tv6_data", "ocu_data"):
        assert token not in production_json.lower()


# ---------------------------------------------------------------------------
# 2-4. Scope exclusivamente YPF; APSA/London excluidos
# ---------------------------------------------------------------------------


def test_universe_scope_is_ypf_only_even_with_other_circuitos_present():
    maestro_rows, campana_rows = _one_station_three_elements()
    mixed = [_static_cencosud("C1"), _apsa_static("A1"), _london_static("L1")] + maestro_rows
    semantic_result = _semantic(mixed, campana_rows)
    universe = td.build_tv5_universe(semantic_result)
    assert set(universe["maestro"]["CircuitoNegocio"].unique()) == {"YPF"}
    assert universe["elementos_catalogo"] == 3


def test_apsa_only_universe_raises_build_error():
    semantic_result = _semantic([_apsa_static("A1")])
    with pytest.raises(td.BuildError):
        td.build_tv5_universe(semantic_result)


def test_london_only_universe_raises_build_error():
    semantic_result = _semantic([_london_static("L1")])
    with pytest.raises(td.BuildError):
        td.build_tv5_universe(semantic_result)


def test_production_universe_is_exactly_ypf(production_result):
    assert production_result["universe"]["circuito"] == "YPF"


def test_apsa_and_london_absent_from_production_payload(production_json):
    assert "APSA" not in production_json
    assert "LONDON" not in production_json.upper()


# ---------------------------------------------------------------------------
# 5-6. Estaciones/elementos catalogo deduplicados
# ---------------------------------------------------------------------------


def test_catalog_stations_deduplicated_across_multiple_elements():
    maestro_rows, _ = _one_station_three_elements()
    semantic_result = _semantic(maestro_rows)
    universe = td.build_tv5_universe(semantic_result)
    assert universe["estaciones_catalogo"] == 1
    assert universe["elementos_catalogo"] == 3


def test_station_without_derivable_locality_blocks_build():
    bad_row = _ypf_digital("999 - TT - 1", ubicacion="SIN_FORMATO_DE_LOCALIDAD")
    semantic_result = _semantic([bad_row])
    with pytest.raises(td.BuildError, match="sin estacion"):
        td.build_tv5_universe(semantic_result)


# ---------------------------------------------------------------------------
# 7-11. Estaciones activas / campanas unicas / activaciones / elementos
# activos no se confunden entre si
# ---------------------------------------------------------------------------


def test_one_campaign_per_element_distinguishes_all_four_metrics():
    """1 estacion, 3 elementos, 3 campanas propias -> estaciones=1,
    elementos=3, campanas_unicas=3, activaciones=3 (nunca se confunden)."""
    maestro_rows, campana_rows = _one_station_three_elements()
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv5_universe(semantic_result)
    period, previous = ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30")

    estaciones = td.compute_estaciones_activas(engine, universe, period, previous)
    campanas = td.compute_campanas(engine, universe, period, previous)
    elementos = td.compute_elementos_activos(engine, period, previous)

    assert estaciones["actual"] == 1
    assert campanas["campanas_unicas_actual"] == 3
    assert campanas["activaciones_actual"] == 3
    assert elementos["actual"] == 3


def test_one_campaign_across_many_elements_counts_one_campana_many_activaciones():
    """1 sola IDCampaña aplicada a 3 ElementoID de 3 estaciones distintas:
    campanas_unicas=1, activaciones=3, estaciones_activas=3, elementos=3."""
    ubic_a, ubic_b, ubic_c = "1 - ALFA - X", "2 - BETA - X", "3 - GAMMA - X"
    maestro_rows = [
        _ypf_digital("1 - TT - 1", ubic_a),
        _ypf_digital("2 - TT - 1", ubic_b),
        _ypf_digital("3 - TT - 1", ubic_c),
    ]
    campana_rows = [
        _campana_row("YC1", "1 - TT - 1", IDCampaña="SAME", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10")),
        _campana_row("YC2", "2 - TT - 1", IDCampaña="SAME", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10")),
        _campana_row("YC3", "3 - TT - 1", IDCampaña="SAME", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv5_universe(semantic_result)
    period, previous = ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30")

    campanas = td.compute_campanas(engine, universe, period, previous)
    estaciones = td.compute_estaciones_activas(engine, universe, period, previous)

    assert campanas["campanas_unicas_actual"] == 1
    assert campanas["activaciones_actual"] == 3
    assert estaciones["actual"] == 3


def test_two_campaigns_same_element_counts_once_as_active_element():
    ubic = "500 - TESTVILLE - X"
    maestro_rows = [_ypf_digital("500 - TT - 1", ubic)]
    campana_rows = [
        _campana_row("YC1", "500 - TT - 1", IDCampaña="C1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-10")),
        _campana_row("YC2", "500 - TT - 1", IDCampaña="C2", FechaInicio=pd.Timestamp("2026-07-15"), FechaFin=pd.Timestamp("2026-07-20")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    period, previous = ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30")
    elementos = td.compute_elementos_activos(engine, period, previous)
    assert elementos["actual"] == 1  # 1 elemento, aunque tuvo 2 campanas


def test_production_metrics_never_collapse_into_each_other(production_result):
    k = production_result["data"]["kpis"]
    campanas = k["campanas"]["campanas_unicas_actual"]
    activaciones = k["campanas"]["activaciones_actual"]
    elementos = k["elementos_activos"]["actual"]
    estaciones = k["estaciones_activas"]["actual"]
    # En produccion, con esta base, las 4 metricas tienen valores distintos.
    assert len({campanas, activaciones, elementos, estaciones}) == 4
    assert activaciones > elementos > estaciones > campanas


# ---------------------------------------------------------------------------
# 12-13. Junio correcto y deltas correctos
# ---------------------------------------------------------------------------


def test_previous_period_computed_independently_from_current():
    maestro_rows = [_ypf_digital("500 - TT - 1", "500 - TESTVILLE - X")]
    campana_rows = [
        _campana_row("YC1", "500 - TT - 1", IDCampaña="C1", FechaInicio=pd.Timestamp("2026-06-05"), FechaFin=pd.Timestamp("2026-06-10")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv5_universe(semantic_result)
    period, previous = ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30")

    estaciones = td.compute_estaciones_activas(engine, universe, period, previous)
    assert estaciones["actual"] == 0
    assert estaciones["anterior"] == 1
    assert estaciones["delta_abs"] == -1


def test_production_deltas_are_actual_minus_anterior(production_result):
    k = production_result["data"]["kpis"]
    est = k["estaciones_activas"]
    assert est["delta_abs"] == est["actual"] - est["anterior"]
    camp = k["campanas"]
    assert camp["delta_campanas"] == camp["campanas_unicas_actual"] - camp["campanas_unicas_anterior"]
    assert camp["delta_activaciones"] == camp["activaciones_actual"] - camp["activaciones_anterior"]
    elem = k["elementos_activos"]
    assert elem["delta"] == elem["actual"] - elem["anterior"]


def test_production_pct_sobre_catalogo_uses_catalog_denominator(production_result):
    est = production_result["data"]["kpis"]["estaciones_activas"]
    esperado = round(est["actual"] / est["catalogo"] * 100.0, 1)
    assert est["pct_actual"] == esperado


# ---------------------------------------------------------------------------
# 14. Formato lider usa campanas unicas, nunca activaciones; desempate
# explicito por elementos activos (nunca por activaciones).
# ---------------------------------------------------------------------------


def test_formato_lider_uses_campanas_unicas_not_activaciones():
    """Formato TORRE: 1 sola campana repartida en 5 elementos (1 campana
    unica, 5 activaciones). Formato PUNTERA: 2 campanas, cada una en 1
    elemento (2 campanas unicas, 2 activaciones). Si se usara activaciones,
    TORRE ganaria (5 > 2); usando campanas unicas, PUNTERA gana (2 > 1)."""
    torre_rows = [_ypf_digital(f"1 - TT - {i}", "1 - ALFA - X") for i in range(5)]
    puntera_rows = [_ypf_digital("2 - PPUNTER - 1", "2 - BETA - X"), _ypf_digital("2 - PPUNTER - 2", "2 - BETA - X")]
    maestro_rows = torre_rows + puntera_rows
    campana_rows = [
        _campana_row(f"T{i}", f"1 - TT - {i}", IDCampaña="TORRE_UNICA", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10"))
        for i in range(5)
    ] + [
        _campana_row("P1", "2 - PPUNTER - 1", IDCampaña="PUNT_A", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10")),
        _campana_row("P2", "2 - PPUNTER - 2", IDCampaña="PUNT_B", FechaInicio=pd.Timestamp("2026-07-05"), FechaFin=pd.Timestamp("2026-07-10")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv5_universe(semantic_result)
    period, previous = ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30")

    fl = td.compute_formato_lider(engine, universe, period, previous)
    assert fl["actual"]["formato"] == "YPF_PUNTERA"
    assert fl["actual"]["campanas_unicas"] == 2


def test_formato_lider_tie_resolved_by_elementos_activos_not_activaciones():
    """Empate 2 vs 2 campanas unicas entre TORRE y PUNTERA. TORRE tiene mas
    ELEMENTOS activos (3) que PUNTERA (1), pero PUNTERA tiene mas
    ACTIVACIONES por repetirse la misma campana en el tiempo... en este
    caso ambos formatos tienen 2 campanas y se decide por elementos
    activos: TORRE (3 elementos) gana sobre PUNTERA (1 elemento)."""
    torre_rows = [_ypf_digital(f"1 - TT - {i}", "1 - ALFA - X") for i in range(3)]
    puntera_rows = [_ypf_digital("2 - PPUNTER - 1", "2 - BETA - X")]
    maestro_rows = torre_rows + puntera_rows
    campana_rows = [
        _campana_row("T0", "1 - TT - 0", IDCampaña="TC1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-05")),
        _campana_row("T1", "1 - TT - 1", IDCampaña="TC2", FechaInicio=pd.Timestamp("2026-07-06"), FechaFin=pd.Timestamp("2026-07-10")),
        _campana_row("P0", "2 - PPUNTER - 1", IDCampaña="PC1", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-05")),
        _campana_row("P1", "2 - PPUNTER - 1", IDCampaña="PC2", FechaInicio=pd.Timestamp("2026-07-06"), FechaFin=pd.Timestamp("2026-07-10")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv5_universe(semantic_result)
    period, previous = ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30")

    fl = td.compute_formato_lider(engine, universe, period, previous)
    assert fl["actual"]["campanas_unicas"] == 2
    assert fl["actual"]["formato"] == "YPF_TORRE"
    assert fl["actual"]["empate_resuelto_por_elementos_activos"] is True


def test_production_formato_lider_julio_is_punteras(production_result):
    # Recalibrado 2026-08-18 (promocion base OCU26+YPF): campanas_unicas 9 -> 17.
    # TV5 es exclusivamente circuito YPF; el reemplazo integro de BASE CAMPAÑAS
    # (YPF Etapa 2) explica el delta en su totalidad.
    fl = production_result["data"]["kpis"]["formato_lider"]["actual"]
    assert fl["label"] == "Punteras"
    assert fl["campanas_unicas"] == 17
    assert fl["empate_resuelto_por_elementos_activos"] is False


def test_production_formato_lider_junio_no_longer_ties(production_result):
    """Recalibrado 2026-08-18 (antes: test_production_formato_lider_junio_ties_resolved_by_elementos_activos).
    Con la base PRE_YPF, junio 2026 tenia un empate real 4 vs 4 entre Torres y
    Punteras, resuelto por elementos_activos (Torres ganaba). YPF Etapa 2
    reconstruyo las campanas de junio desde cero (mas granular, ver
    docs/CM2_CIERRE_BASE_OCU26_YPF_2026-08-18.md): con los datos productivos
    reales ya NO hay empate (Punteras lidera con margen propio en campanas
    unicas), asi que 'empate_resuelto_por_elementos_activos' pasa a False.
    La cobertura de la logica de desempate en si (para cuando SI hay empate)
    se conserva integra en el test sintetico
    test_formato_lider_tie_resolved_by_elementos_activos_not_activaciones
    (datos controlados, independiente del Excel productivo)."""
    fl = production_result["data"]["kpis"]["formato_lider"]["anterior"]
    assert fl["label"] == "Punteras"
    assert fl["campanas_unicas"] == 7
    assert fl["empate_resuelto_por_elementos_activos"] is False


# ---------------------------------------------------------------------------
# 15-16. Historico Ene-Jul, nunca Ago-Dic
# ---------------------------------------------------------------------------


def test_production_historico_covers_exactly_ene_jul(production_result):
    hist = production_result["data"]["historico"]
    assert hist["meses"] == ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul"]
    for label, serie in hist["series"].items():
        assert len(serie) == 7


def test_production_historico_never_invents_ago_dic(production_json):
    assert "Ago" not in production_json
    assert "Sep" not in production_json
    assert "Oct" not in production_json
    assert "Nov" not in production_json
    assert "Dic" not in production_json


def test_production_historico_series_are_the_three_named_formats(production_result):
    assert set(production_result["data"]["historico"]["series"].keys()) == {"Menu Board", "Torres", "Punteras"}


# ---------------------------------------------------------------------------
# 17. Mapa no inventa coordenadas
# ---------------------------------------------------------------------------


def test_mapa_payload_never_contains_lat_lon_keys(production_json):
    payload = json.loads(production_json)
    mapa = payload["mapa"]
    forbidden_keys = {"lat", "lon", "lat_lon", "latitud", "longitud", "latitude", "longitude", "coordinates", "lng"}
    assert forbidden_keys.isdisjoint(mapa.keys())
    for row in mapa["top_localidades"]:
        assert set(row.keys()) == {"nombre", "activaciones", "pct"}
        assert forbidden_keys.isdisjoint(row.keys())


def test_mapa_marked_as_a_validar_with_metric_status(production_result):
    mapa = production_result["data"]["mapa"]
    assert mapa["estado"] == "A_VALIDAR"
    assert mapa["metric_status"] == "REQUIERE_CONFIRMACION"
    assert mapa["nota"]


def test_mapa_top_localidades_use_real_ciudad_field_only():
    ubic_a = "1 - ALFA - X"
    maestro_rows = [_ypf_digital("1 - TT - 1", ubic_a, Ciudad="ALFA")]
    campana_rows = [_campana_row("C1", "1 - TT - 1", IDCampaña="C1", FechaInicio=pd.Timestamp("2026-01-05"), FechaFin=pd.Timestamp("2026-01-10"))]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    universe = td.build_tv5_universe(semantic_result)
    mapa = td.compute_mapa(engine, universe, 2026, 1)
    assert mapa["top_localidades"] == [{"nombre": "ALFA", "activaciones": 1, "pct": 100.0}]


# ---------------------------------------------------------------------------
# 18. MetricStatus preservado
# ---------------------------------------------------------------------------


def test_metric_status_present_on_catalogo_elementos_and_mapa(production_result):
    k = production_result["data"]["kpis"]
    assert k["catalogo"]["metric_status"] == "NO_APLICA"
    assert k["elementos_activos"]["metric_status"] in {"OK", "PARTIAL", "NO_APLICA", "REQUIERE_CONFIRMACION"}
    assert production_result["data"]["mapa"]["metric_status"] == "REQUIERE_CONFIRMACION"


# ---------------------------------------------------------------------------
# 19. No window.OCU_DATA legacy
# ---------------------------------------------------------------------------


def test_output_html_uses_tv5_data_not_legacy_ocu_data(production_html):
    assert "window.TV5_DATA" in production_html
    assert "OCU_DATA" not in production_html


# ---------------------------------------------------------------------------
# 20. Excel intacto (SHA-256 antes/despues)
# ---------------------------------------------------------------------------


def test_input_excel_sha_unchanged(production_result):
    sha_now = vi.calculate_sha256(PRODUCTION_FILE)
    assert sha_now == production_result["sha256"]


# ---------------------------------------------------------------------------
# 21-22. Archivos protegidos existen y no fueron tocados por este builder
# ---------------------------------------------------------------------------


def test_protected_reference_file_untouched_and_present():
    assert (REPO_ROOT / "audit_sources" / "TV5_REFERENCE.html").exists()


def test_other_tv_outputs_still_present():
    for name in ("tv1.html", "tv2.html", "tv3.html", "tv4.html", "tv6.html"):
        assert (REPO_ROOT / name).exists(), f"{name} no deberia haber sido eliminado por el build de TV5"


def test_builder_does_not_import_other_tv_builders():
    for token in ("import build_tv1_dashboard", "import build_tv2_dashboard", "import build_tv3_dashboard", "import build_tv4_dashboard", "import build_tv6_dashboard"):
        assert token not in BUILDER_SOURCE


# ---------------------------------------------------------------------------
# 23. Sin datos dummy hardcodeados en el builder
# ---------------------------------------------------------------------------


def test_builder_source_has_no_hardcoded_production_values():
    for literal in ("451", "3082", "305", "263", "7681", "4286", "2651", "2371"):
        assert literal not in BUILDER_SOURCE


def test_builder_source_has_no_hardcoded_dummy_business_values():
    # No debe existir logica condicional sobre nombres de estacion/ciudad
    # especificos (senal de hardcode de negocio en vez de derivacion real).
    for token in ("Martinez", "610 activaciones"):
        assert token not in BUILDER_SOURCE
