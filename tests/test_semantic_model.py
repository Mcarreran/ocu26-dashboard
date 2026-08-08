"""Pruebas para scripts/semantic_model.py (Gate 3B - capa semantica OCU26).

No modifica scripts/validate_input.py, scripts/transform_data.py, sus tests,
ni input/OCU26_BASE_DATOS.xlsx. Reutiliza transform_data() como unica fuente
de datos reales; los tests de jerarquia/validacion de config usan fixtures
sinteticos construidos en memoria (nunca escriben un .xlsx).

Ninguna cifra de negocio actual (cantidad de elementos por circuito, etc.)
se hardcodea como valor esperado fijo: las comparaciones contra datos reales
se hacen contra un conteo derivado dinamicamente del propio maestro crudo.
"""

from __future__ import annotations

import copy
import sys
from pathlib import Path

import pandas as pd
import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT / "scripts"))

import semantic_model as sm  # noqa: E402
import validate_input as vi  # noqa: E402
from transform_data import transform_data  # noqa: E402

PRODUCTION_FILE = REPO_ROOT / "input" / "OCU26_BASE_DATOS.xlsx"


# ---------------------------------------------------------------------------
# Fixtures sinteticos (no dependen del Excel)
# ---------------------------------------------------------------------------


def _maestro_row(elemento_id: str, **overrides) -> dict:
    row = dict.fromkeys(vi.MAESTRO_HEADERS)
    row.update(
        {
            "ElementoID": elemento_id,
            "TipoCatalogo": "Cerrado",
            "Ciudad": "CABA",
            "Medio": "Digital",
            "CircuitoDashboard": "TEST_CIRCUITO",
            "Subcircuito": "TEST_SUB",
            "Ubicacion": "TEST_UBI",
            "Nivel": "",
            "Descripcion": "",
            "TipoInventario": "Digital",
            "AplicaCantidad": "NO",
            "CapacidadSlotsReel": 20,
            "SegundosDia": 100800,
        }
    )
    row.update(overrides)
    return row


def _maestro_df(rows: list[dict]) -> pd.DataFrame:
    return pd.DataFrame(rows, columns=vi.MAESTRO_HEADERS)


def _campana_row(carga_id: str, elemento_id: str, **overrides) -> dict:
    row = dict.fromkeys(vi.CAMPANAS_HEADERS)
    row.update(
        {
            "CargaID": carga_id,
            "ElementoID": elemento_id,
            "FechaInicio": pd.Timestamp("2026-01-01"),
            "FechaFin": pd.Timestamp("2026-01-31"),
            "FechaIndefinida": "No",
            "Estado": "Activa",
            "DuracionSpotSeg": 10.0,
            "SalidasVendidas": 2.0,
            "ModalidadPauta": "Slot / Reel normal",
            "Cliente": "CLIENTE_TEST",
            "Marca": "MARCA_TEST",
        }
    )
    row.update(overrides)
    return row


def _campanas_df(rows: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(rows, columns=vi.CAMPANAS_HEADERS) if rows else pd.DataFrame(columns=vi.CAMPANAS_HEADERS)
    df["FechaInicio"] = pd.to_datetime(df["FechaInicio"])
    df["FechaFin"] = pd.to_datetime(df["FechaFin"])
    return df


def _transform_result(maestro_rows: list[dict], campanas_rows: list[dict] | None = None) -> dict:
    return {
        "maestro": _maestro_df(maestro_rows),
        "campanas": _campanas_df(campanas_rows or []),
        "parametros": pd.DataFrame(columns=["Categoria", "Valor"]),
        "warnings": [],
    }


_BUNDLE_BASE = {
    "circuito_negocio": "X",
    "portfolio_tier": "CORE",
    "incluye_performance_core": True,
    "incluye_conteo_general": True,
    "visible_por_defecto": True,
    "cobertura_catalogo": "COMPLETO",
    "completitud_maestro": "COMPLETO",
    "modo_disponibilidad": "CALCULABLE",
    "certeza_dato_regla": "cerrado_gobernado",
}


def _minimal_config(elemento_overrides=None, circuito_subcircuito_rules=None, circuito_dashboard_rules=None, generic_rules=None) -> dict:
    return {
        "schema_version": 1,
        "resolution_order": [
            "elemento_overrides",
            "circuito_subcircuito_rules",
            "circuito_dashboard_rules",
            "generic_rules",
            "defaults",
        ],
        "defaults": {
            "circuito_negocio": "NO_CLASIFICADO",
            "portfolio_tier": "NO_CLASIFICADO",
            "incluye_performance_core": False,
            "incluye_conteo_general": True,
            "visible_por_defecto": True,
            "cobertura_catalogo": "DESCONOCIDO",
            "completitud_maestro": "NO_APLICA",
            "modo_disponibilidad": "CONSULTA",
            "certeza_dato_regla": "requiere_revision",
        },
        "elemento_overrides": elemento_overrides or [],
        "circuito_subcircuito_rules": circuito_subcircuito_rules or [],
        "circuito_dashboard_rules": circuito_dashboard_rules or [],
        "generic_rules": generic_rules or [],
        "certeza_dato": {
            "rule_outputs": {
                "cerrado_gobernado": "CONFIRMADO",
                "requiere_revision": "REQUIERE_REVISION",
                "campana_confirma_con_actividad": "CONFIRMADO",
                "campana_confirma_sin_actividad": "REGISTRADO_NO_CONFIRMADO",
            },
            "technical_override_value": "PROVISORIO_REQUIERE_REVISION",
        },
        "sitio_negocio": {"default_field": "Ubicacion", "circuito_field_overrides": {}, "ubicacion_lookup": {}},
        "formato_negocio": {
            "ypf_elemento_id_token_map": {},
            "circuito_dashboard_formato_fijo": {},
            "descripcion_keyword_rules": [],
            "default": "OTRO",
        },
        "digital_capacity": {
            "default_segundos_comerciales": 72000,
            "spot_duracion_base_seg": 10,
            "spot_segundos_por_salida": 1800,
            "slots_profiles": {},
            "use_legacy_source_if_positive_for_circuitos": [],
            "requiere_confirmacion_sentinel": "REQUIERE_CONFIRMACION",
        },
        "universes": {
            "OPERATIVO_GENERAL": {"require_flag": "IncluyeConteoGeneral"},
            "PERFORMANCE_CORE": {"require_flag": "IncluyePerformanceCore"},
            "COMPLETO_HISTORICO": {"require_flag": None},
        },
        "metric_policies": {
            "slot_seconds_no_aplica_circuitos": [],
            "slot_seconds_no_aplica_metrics": [
                "slots_ocupados", "slots_disponibles", "fill_rate_slots",
                "segundos_vendidos", "segundos_disponibles", "fill_rate_segundos",
            ],
            "slot_seconds_no_aplica_warning": "politica de prueba",
        },
    }


@pytest.fixture
def base_config() -> dict:
    return copy.deepcopy(sm.load_config())


@pytest.fixture(scope="module")
def semantic_result() -> dict:
    return sm.build_semantic_model(transform_data(PRODUCTION_FILE))


@pytest.fixture(scope="module")
def raw_maestro() -> pd.DataFrame:
    return transform_data(PRODUCTION_FILE)["maestro"]


# ---------------------------------------------------------------------------
# Carga y validacion de configuracion
# ---------------------------------------------------------------------------


def test_load_config_real_file_is_valid():
    config = sm.load_config()
    assert config["schema_version"] == 1


def test_load_config_missing_file_raises(tmp_path):
    with pytest.raises(sm.SemanticConfigError):
        sm.load_config(tmp_path / "no_existe.json")


def test_load_config_invalid_json_raises(tmp_path):
    path = tmp_path / "bad.json"
    path.write_text("{esto no es json valido", encoding="utf-8")
    with pytest.raises(sm.SemanticConfigError):
        sm.load_config(path)


def test_validate_config_rejects_wrong_schema_version(base_config):
    base_config["schema_version"] = 2
    with pytest.raises(sm.SemanticConfigError):
        sm.validate_config(base_config)


def test_validate_config_rejects_missing_top_level_key(base_config):
    del base_config["universes"]
    with pytest.raises(sm.SemanticConfigError):
        sm.validate_config(base_config)


def test_validate_config_rejects_ambiguous_circuito_subcircuito_rule(base_config):
    base_config["circuito_subcircuito_rules"].append(dict(base_config["circuito_subcircuito_rules"][0]))
    with pytest.raises(sm.SemanticConfigError, match="ambigua"):
        sm.validate_config(base_config)


def test_validate_config_rejects_ambiguous_circuito_dashboard_rule(base_config):
    base_config["circuito_dashboard_rules"].append(dict(base_config["circuito_dashboard_rules"][0]))
    with pytest.raises(sm.SemanticConfigError, match="ambiguo"):
        sm.validate_config(base_config)


def test_validate_config_rejects_bad_portfolio_tier_value(base_config):
    base_config["defaults"]["portfolio_tier"] = "TIER_INVENTADO"
    with pytest.raises(sm.SemanticConfigError):
        sm.validate_config(base_config)


def test_validate_config_rejects_missing_bundle_field(base_config):
    del base_config["circuito_dashboard_rules"][0]["modo_disponibilidad"]
    with pytest.raises(sm.SemanticConfigError):
        sm.validate_config(base_config)


def test_validate_config_rejects_campana_confirma_without_required_outputs(base_config):
    del base_config["certeza_dato"]["rule_outputs"]["campana_confirma_sin_actividad"]
    with pytest.raises(sm.SemanticConfigError):
        sm.validate_config(base_config)


def test_validate_config_rejects_non_boolean_flag(base_config):
    base_config["defaults"]["incluye_performance_core"] = "true"
    with pytest.raises(sm.SemanticConfigError):
        sm.validate_config(base_config)


def test_validate_config_rejects_duplicate_elemento_override(base_config):
    override = {"elemento_id": "DUP-1", **_BUNDLE_BASE}
    base_config["elemento_overrides"] = [dict(override), dict(override)]
    with pytest.raises(sm.SemanticConfigError, match="duplicado"):
        sm.validate_config(base_config)


def test_validate_config_rejects_bad_digital_capacity_profile(base_config):
    base_config["digital_capacity"]["slots_profiles"]["ROTO"] = -5
    with pytest.raises(sm.SemanticConfigError):
        sm.validate_config(base_config)


# ---------------------------------------------------------------------------
# Jerarquia de resolucion (5 niveles), con config minimal sintetica
# ---------------------------------------------------------------------------


def test_hierarchy_elemento_override_wins_over_all_other_levels():
    config = _minimal_config(
        elemento_overrides=[{"elemento_id": "E1", **{**_BUNDLE_BASE, "circuito_negocio": "OVERRIDE"}}],
        circuito_subcircuito_rules=[
            {"circuito_dashboard": ["CD1"], "subcircuito": ["S1"], **{**_BUNDLE_BASE, "circuito_negocio": "SUBCIRCUITO"}}
        ],
        circuito_dashboard_rules=[{"circuito_dashboard": ["CD1"], **{**_BUNDLE_BASE, "circuito_negocio": "DASHBOARD"}}],
    )
    rows = [_maestro_row("E1", CircuitoDashboard="CD1", Subcircuito="S1")]
    result = sm.build_semantic_model(_transform_result(rows), config)
    assert result["maestro"].iloc[0]["CircuitoNegocio"] == "OVERRIDE"
    assert result["stats"]["matched_level_counts"] == {"elemento_override": 1}


def test_hierarchy_circuito_subcircuito_wins_over_circuito_dashboard():
    config = _minimal_config(
        circuito_subcircuito_rules=[
            {"circuito_dashboard": ["CD1"], "subcircuito": ["S1"], **{**_BUNDLE_BASE, "circuito_negocio": "SUBCIRCUITO"}}
        ],
        circuito_dashboard_rules=[{"circuito_dashboard": ["CD1"], **{**_BUNDLE_BASE, "circuito_negocio": "DASHBOARD"}}],
    )
    rows = [
        _maestro_row("E1", CircuitoDashboard="CD1", Subcircuito="S1"),
        _maestro_row("E2", CircuitoDashboard="CD1", Subcircuito="OTRO_SUB"),
    ]
    result = sm.build_semantic_model(_transform_result(rows), config)
    m = result["maestro"].set_index("ElementoID")
    assert m.loc["E1", "CircuitoNegocio"] == "SUBCIRCUITO"
    assert m.loc["E2", "CircuitoNegocio"] == "DASHBOARD"


def test_hierarchy_generic_rule_applies_when_no_specific_rule_matches():
    config = _minimal_config(
        generic_rules=[
            {
                "id": "g1",
                "when": {"field": "TipoInventario", "equals": "Flexible gráfico"},
                **{**_BUNDLE_BASE, "circuito_negocio": "GENERICO"},
            }
        ]
    )
    rows = [_maestro_row("E1", CircuitoDashboard="Cualquiera", TipoInventario="Flexible gráfico")]
    result = sm.build_semantic_model(_transform_result(rows), config)
    assert result["maestro"].iloc[0]["CircuitoNegocio"] == "GENERICO"


def test_hierarchy_default_fallback_is_safe_and_warns():
    config = _minimal_config()
    rows = [_maestro_row("E1", CircuitoDashboard="CIRCUITO_DESCONOCIDO")]
    result = sm.build_semantic_model(_transform_result(rows), config)
    row = result["maestro"].iloc[0]
    assert row["CircuitoNegocio"] == "NO_CLASIFICADO"
    assert row["PortfolioTier"] == "NO_CLASIFICADO"
    assert row["IncluyePerformanceCore"] is False or row["IncluyePerformanceCore"] == False  # noqa: E712
    assert any("no coincide con ninguna regla" in w for w in result["warnings"])


def test_original_columns_are_never_modified():
    config = sm.load_config()
    rows = [_maestro_row("E1", CircuitoDashboard="Pantalla Led", Subcircuito="PLED", Ubicacion="TEST", CapacidadSlotsReel=40, SegundosDia=75000)]
    tr = _transform_result(rows)
    result = sm.build_semantic_model(tr, config)
    pd.testing.assert_frame_equal(result["maestro"][vi.MAESTRO_HEADERS], tr["maestro"][vi.MAESTRO_HEADERS])


def test_campanas_join_preserves_row_count_and_adds_circuito_negocio():
    config = sm.load_config()
    maestro_rows = [_maestro_row("E1", CircuitoDashboard="Pantalla Led", Subcircuito="PLED", Ubicacion="CABILDO")]
    campanas_rows = [_campana_row("C1", "E1"), _campana_row("C2", "E1")]
    result = sm.build_semantic_model(_transform_result(maestro_rows, campanas_rows), config)
    assert len(result["campanas"]) == 2
    assert (result["campanas"]["CircuitoNegocio"] == "PANTALLAS_LED").all()


def test_pantalla_led_new_location_classified_generically_without_config_change():
    """No hardcodear '11 pantallas': un elemento futuro nuevo debe clasificarse
    CORE automaticamente por CircuitoDashboard, sin tocar config ni codigo."""
    config = sm.load_config()
    rows = [
        _maestro_row(
            "C99 - NEW", CircuitoDashboard="Pantalla Led", Subcircuito="PLED", Ubicacion="ROSARIO",
            Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
            CapacidadSlotsReel=40, SegundosDia=75000,
        )
    ]
    result = sm.build_semantic_model(_transform_result(rows), config)
    row = result["maestro"].iloc[0]
    assert row["CircuitoNegocio"] == "PANTALLAS_LED"
    assert row["PortfolioTier"] == "CORE"
    assert row["SlotsComerciales"] == 20


# ---------------------------------------------------------------------------
# Reglas de negocio reales (contra produccion, config real)
# ---------------------------------------------------------------------------


def test_production_file_no_elementoid_lost_or_duplicated(semantic_result, raw_maestro):
    m = semantic_result["maestro"]
    assert len(m) == len(raw_maestro)
    assert set(m["ElementoID"]) == set(raw_maestro["ElementoID"])
    assert m["ElementoID"].duplicated().sum() == raw_maestro["ElementoID"].duplicated().sum()


def test_cencosud_confirmado_regardless_of_activity(semantic_result):
    m = semantic_result["maestro"]
    cencosud = m[m["CircuitoNegocio"] == "CENCOSUD"]
    assert len(cencosud) > 0
    assert (cencosud["PortfolioTier"] == "CORE").all()
    assert cencosud["IncluyePerformanceCore"].all()
    assert (cencosud["CertezaDato"] == "CONFIRMADO").all()


def test_remeros_distinct_from_its_led_screen_but_shares_sitio(semantic_result):
    m = semantic_result["maestro"]
    remeros_shopping = m[m["CircuitoNegocio"] == "REMEROS"]
    assert len(remeros_shopping) > 0
    assert (remeros_shopping["PortfolioTier"] == "CORE").all()

    pled_remeros = m[(m["CircuitoDashboard"] == "Pantalla Led") & (m["Ubicacion"] == "REMEROS")]
    assert len(pled_remeros) == 1
    assert pled_remeros.iloc[0]["CircuitoNegocio"] == "PANTALLAS_LED"
    assert pled_remeros.iloc[0]["SitioNegocio"] == remeros_shopping.iloc[0]["SitioNegocio"]


def test_apsa_legacy_hidden_but_physically_present(semantic_result, raw_maestro):
    m = semantic_result["maestro"]
    apsa = m[m["CircuitoNegocio"] == "APSA"]
    raw_apsa = raw_maestro[(raw_maestro["CircuitoDashboard"] == "Shoppings Estático") & (raw_maestro["Subcircuito"] == "APSA")]
    assert len(apsa) == len(raw_apsa) > 0
    assert (apsa["PortfolioTier"] == "LEGACY").all()
    assert not apsa["IncluyePerformanceCore"].any()
    assert not apsa["IncluyeConteoGeneral"].any()
    assert (apsa["CertezaDato"] == "LEGACY").all()


def test_london_supply_complementario_confirmado_sin_depender_de_actividad(semantic_result):
    m = semantic_result["maestro"]
    ls = m[m["CircuitoNegocio"] == "LONDON_SUPPLY"]
    assert len(ls) > 0
    assert (ls["PortfolioTier"] == "COMPLEMENTARIO").all()
    assert not ls["IncluyePerformanceCore"].any()
    assert ls["IncluyeConteoGeneral"].all()
    assert (ls["CoberturaCatalogo"] == "COMPLETO").all()
    con_capacidad = ls[(ls["Medio"] == "Digital") & (ls["CapacidadSlotsReel"] > 0)]
    assert len(con_capacidad) > 0
    assert (con_capacidad["CertezaDato"] == "CONFIRMADO").all()


def test_aa2000_completitud_parcial_y_override_tecnico(semantic_result):
    m = semantic_result["maestro"]
    aa = m[m["CircuitoNegocio"] == "AA2000"]
    assert len(aa) > 0
    assert (aa["CoberturaCatalogo"] == "COMPLETO").all()
    assert (aa["CompletitudMaestro"] == "PARCIAL").all()
    assert (aa["ModoDisponibilidad"] == "MIXTO").all()

    zero_cap_digital = aa[(aa["Medio"] == "Digital") & (aa["CapacidadSlotsReel"] == 0)]
    assert len(zero_cap_digital) > 0
    assert (zero_cap_digital["CertezaDato"] == "PROVISORIO_REQUIERE_REVISION").all()


def test_ypf_certeza_depende_de_actividad_no_de_catalogo(semantic_result):
    m = semantic_result["maestro"]
    ypf = m[m["CircuitoNegocio"] == "YPF"]
    con_actividad = ypf[ypf["TieneActividadComercial"]]
    sin_actividad = ypf[~ypf["TieneActividadComercial"]]
    assert len(con_actividad) > 0 and (con_actividad["CertezaDato"] == "CONFIRMADO").all()
    assert len(sin_actividad) > 0 and (sin_actividad["CertezaDato"] == "REGISTRADO_NO_CONFIRMADO").all()


def test_ypf_formato_tokens_match_elemento_id(semantic_result):
    m = semantic_result["maestro"]
    ypf = m[m["CircuitoNegocio"] == "YPF"].copy()
    token_map = {"MB": "YPF_MENU_BOARD", "TT": "YPF_TORRE", "PPUNTER": "YPF_PUNTERA", "FB": "YPF_MUPI_FOTOBOX"}
    ypf["_token"] = ypf["ElementoID"].str.split(" - ").str[1]
    expected = ypf["_token"].map(token_map)
    assert (ypf["FormatoNegocio"] == expected).all()
    assert set(ypf["FormatoNegocio"].unique()) == set(token_map.values())


def test_cencomedia_via_generic_rule_not_hardcoded_store_list(semantic_result):
    m = semantic_result["maestro"]
    cenco = m[m["CircuitoNegocio"] == "CENCOMEDIA"]
    assert len(cenco) > 0
    assert (cenco["TipoInventario"] == "Flexible gráfico").all()
    assert (cenco["PortfolioTier"] == "COMPLEMENTARIO").all()
    assert not cenco["IncluyePerformanceCore"].any()
    assert (cenco["CertezaDato"] == "NO_EVALUADO").all()
    assert cenco["CircuitoDashboard"].nunique() > 1  # multiples tiendas Jumbo/Disco


def test_mab_flexible_no_permanente(semantic_result):
    m = semantic_result["maestro"]
    mab = m[m["CircuitoNegocio"] == "MAB"]
    assert len(mab) > 0
    assert (mab["PortfolioTier"] == "COMPLEMENTARIO").all()
    assert (mab["CoberturaCatalogo"] == "DESCONOCIDO").all()
    assert (mab["CertezaDato"] == "REGISTRADO").all()


def test_pilar_frontlight_shares_sitio_with_pantalla_led_pilar(semantic_result):
    m = semantic_result["maestro"]
    pilar_front = m[m["CircuitoNegocio"] == "PILAR_FRONTLIGHT"]
    assert len(pilar_front) == 1
    pled_pilar = m[(m["CircuitoDashboard"] == "Pantalla Led") & (m["Ubicacion"] == "PILAR")]
    assert len(pled_pilar) == 1
    assert pilar_front.iloc[0]["SitioNegocio"] == pled_pilar.iloc[0]["SitioNegocio"]
    assert pilar_front.iloc[0]["CircuitoNegocio"] != pled_pilar.iloc[0]["CircuitoNegocio"]
    assert pilar_front.iloc[0]["PortfolioTier"] == "CORE"


def test_pantalla_led_cordoba_is_core_and_matches_dynamic_count(semantic_result, raw_maestro):
    m = semantic_result["maestro"]
    cordoba = m[(m["CircuitoDashboard"] == "Pantalla Led") & (m["Ubicacion"] == "CORDOBA")]
    assert len(cordoba) == 1
    assert cordoba.iloc[0]["CircuitoNegocio"] == "PANTALLAS_LED"
    assert cordoba.iloc[0]["PortfolioTier"] == "CORE"
    total_enriquecido = m[m["CircuitoNegocio"] == "PANTALLAS_LED"]
    total_crudo = raw_maestro[raw_maestro["CircuitoDashboard"] == "Pantalla Led"]
    assert len(total_enriquecido) == len(total_crudo)


def test_circuitos_cerrados_confirmado_con_cero_campanas(semantic_result):
    m = semantic_result["maestro"]
    cerrados = m[m["CircuitoNegocio"].isin(["CENCOSUD", "REMEROS", "PANTALLAS_LED", "PILAR_FRONTLIGHT", "AA2000"])]
    sin_actividad = cerrados[~cerrados["TieneActividadComercial"]]
    normal = sin_actividad[sin_actividad["CertezaDato"] != "PROVISORIO_REQUIERE_REVISION"]
    if len(normal):
        assert (normal["CertezaDato"] == "CONFIRMADO").all()


# ---------------------------------------------------------------------------
# Capacidad digital: legacy vs. comercial efectiva
# ---------------------------------------------------------------------------


def test_puente_led_capacidad_comercial_13_legacy_10(semantic_result):
    m = semantic_result["maestro"]
    puente = m[m["FormatoNegocio"] == "PUENTE_LED"]
    assert len(puente) > 0
    assert (puente["SlotsComerciales"] == 13).all()
    assert (puente["CapacidadSlotsReel"] == 10).all()


def test_pantalla_led_capacidad_comercial_20_legacy_40(semantic_result):
    m = semantic_result["maestro"]
    pled = m[m["FormatoNegocio"] == "PANTALLA_LED"]
    assert len(pled) > 0
    assert (pled["SlotsComerciales"] == 20).all()
    assert (pled["CapacidadSlotsReel"] == 40).all()


def test_segundos_comerciales_default_72000_no_existe_en_legacy(semantic_result):
    m = semantic_result["maestro"]
    digital = m[m["Medio"] == "Digital"]
    assert (digital["SegundosComerciales"] == 72000).all()
    assert not (digital["SegundosDia"] == 72000).any()


def test_capacidad_fuente_cero_marca_requiere_confirmacion(semantic_result):
    m = semantic_result["maestro"]
    zero_cap = m[(m["Medio"] == "Digital") & (m["CapacidadSlotsReel"] == 0)]
    assert len(zero_cap) > 0
    sentinel = semantic_result["config"]["digital_capacity"]["requiere_confirmacion_sentinel"]
    assert (zero_cap["SlotsComerciales"] == sentinel).all()


# ---------------------------------------------------------------------------
# Universos de reporte
# ---------------------------------------------------------------------------


def test_performance_core_excludes_complementario_and_legacy(semantic_result):
    df = sm.filter_universe(semantic_result["maestro"], "PERFORMANCE_CORE", semantic_result["config"])
    circuitos = set(df["CircuitoNegocio"].unique())
    for excluido in ("APSA", "LONDON_SUPPLY", "CENCOMEDIA", "MAB"):
        assert excluido not in circuitos
    for incluido in ("CENCOSUD", "REMEROS", "PANTALLAS_LED", "PILAR_FRONTLIGHT", "AA2000", "YPF"):
        assert incluido in circuitos


def test_operativo_general_excludes_only_apsa(semantic_result):
    df = sm.filter_universe(semantic_result["maestro"], "OPERATIVO_GENERAL", semantic_result["config"])
    circuitos = set(df["CircuitoNegocio"].unique())
    assert "APSA" not in circuitos
    for incluido in ("LONDON_SUPPLY", "CENCOMEDIA", "MAB", "CENCOSUD", "REMEROS", "PANTALLAS_LED", "PILAR_FRONTLIGHT", "AA2000", "YPF"):
        assert incluido in circuitos


def test_completo_historico_includes_apsa_and_everything_else(semantic_result):
    df = sm.filter_universe(semantic_result["maestro"], "COMPLETO_HISTORICO", semantic_result["config"])
    assert "APSA" in set(df["CircuitoNegocio"].unique())
    assert len(df) == len(semantic_result["maestro"])


def test_filter_universe_unknown_raises(semantic_result):
    with pytest.raises(sm.SemanticModelError):
        sm.filter_universe(semantic_result["maestro"], "NO_EXISTE", semantic_result["config"])


# ---------------------------------------------------------------------------
# Gate3B.1 Sec.6 (item M/N): resolution_order gobierna realmente la jerarquia
# ---------------------------------------------------------------------------


def test_resolution_order_realmente_gobierna_precedencia():
    """Invirtiendo el orden (circuito_dashboard_rules antes que
    circuito_subcircuito_rules) debe ganar la regla menos especifica: prueba
    que el resolver lee config["resolution_order"], no una secuencia fija."""
    config = _minimal_config(
        circuito_subcircuito_rules=[
            {"circuito_dashboard": ["CD1"], "subcircuito": ["S1"], **{**_BUNDLE_BASE, "circuito_negocio": "SUBCIRCUITO"}}
        ],
        circuito_dashboard_rules=[{"circuito_dashboard": ["CD1"], **{**_BUNDLE_BASE, "circuito_negocio": "DASHBOARD"}}],
    )
    config["resolution_order"] = [
        "circuito_dashboard_rules", "circuito_subcircuito_rules", "elemento_overrides", "generic_rules", "defaults",
    ]
    rows = [_maestro_row("E1", CircuitoDashboard="CD1", Subcircuito="S1")]
    result = sm.build_semantic_model(_transform_result(rows), config)
    assert result["maestro"].iloc[0]["CircuitoNegocio"] == "DASHBOARD"
    assert result["stats"]["matched_level_counts"] == {"circuito_dashboard": 1}


def test_validate_config_rejects_resolution_order_missing_defaults(base_config):
    base_config["resolution_order"] = [
        "elemento_overrides", "circuito_subcircuito_rules", "circuito_dashboard_rules", "generic_rules",
    ]
    with pytest.raises(sm.SemanticConfigError):
        sm.validate_config(base_config)


def test_validate_config_rejects_resolution_order_duplicate_level(base_config):
    base_config["resolution_order"] = [
        "elemento_overrides", "elemento_overrides", "circuito_subcircuito_rules",
        "circuito_dashboard_rules", "generic_rules", "defaults",
    ]
    with pytest.raises(sm.SemanticConfigError):
        sm.validate_config(base_config)


def test_validate_config_rejects_resolution_order_unknown_level(base_config):
    base_config["resolution_order"] = [
        "elemento_overrides", "circuito_subcircuito_rules", "circuito_dashboard_rules",
        "generic_rules", "nivel_inventado",
    ]
    with pytest.raises(sm.SemanticConfigError):
        sm.validate_config(base_config)


def test_validate_config_rejects_resolution_order_defaults_not_last(base_config):
    base_config["resolution_order"] = [
        "defaults", "elemento_overrides", "circuito_subcircuito_rules", "circuito_dashboard_rules", "generic_rules",
    ]
    with pytest.raises(sm.SemanticConfigError):
        sm.validate_config(base_config)


# ---------------------------------------------------------------------------
# Gate3B.1 Sec.7 (item O/P/Q): overrides de capacidad por ElementoID
# ---------------------------------------------------------------------------


def test_elemento_override_slots_comerciales_pisa_perfil_formato():
    config = copy.deepcopy(sm.load_config())
    config["elemento_overrides"] = [{"elemento_id": "E1", **_BUNDLE_BASE, "slots_comerciales": 15}]
    rows = [_maestro_row(
        "E1", CircuitoDashboard="Pantalla Led", Subcircuito="PLED", Ubicacion="CABILDO",
        Medio="Digital", CapacidadSlotsReel=40, SegundosDia=75000,
    )]
    result = sm.build_semantic_model(_transform_result(rows), config)
    row = result["maestro"].iloc[0]
    assert row["SlotsComerciales"] == 15  # no 20 (perfil PANTALLA_LED)
    assert row["FormatoNegocio"] == "PANTALLA_LED"


def test_elemento_override_segundos_comerciales_pisa_default():
    config = copy.deepcopy(sm.load_config())
    config["elemento_overrides"] = [{"elemento_id": "E1", **_BUNDLE_BASE, "segundos_comerciales": 65000}]
    rows = [_maestro_row(
        "E1", CircuitoDashboard="Pantalla Led", Subcircuito="PLED", Ubicacion="CABILDO",
        Medio="Digital", CapacidadSlotsReel=40, SegundosDia=75000,
    )]
    result = sm.build_semantic_model(_transform_result(rows), config)
    row = result["maestro"].iloc[0]
    assert row["SegundosComerciales"] == 65000  # no 72000 (default)
    assert row["SlotsComerciales"] == 20  # perfil normal, no afectado por el override de segundos


def test_capacity_override_no_modifica_columnas_legacy():
    config = copy.deepcopy(sm.load_config())
    config["elemento_overrides"] = [
        {"elemento_id": "E1", **_BUNDLE_BASE, "slots_comerciales": 15, "segundos_comerciales": 65000}
    ]
    rows = [_maestro_row(
        "E1", CircuitoDashboard="Pantalla Led", Subcircuito="PLED", Ubicacion="CABILDO",
        Medio="Digital", CapacidadSlotsReel=40, SegundosDia=75000,
    )]
    result = sm.build_semantic_model(_transform_result(rows), config)
    row = result["maestro"].iloc[0]
    assert row["CapacidadSlotsReel"] == 40
    assert row["SegundosDia"] == 75000
    assert row["SlotsComerciales"] == 15
    assert row["SegundosComerciales"] == 65000


def test_validate_config_rejects_negative_slots_comerciales_override(base_config):
    base_config["elemento_overrides"] = [{"elemento_id": "X", **_BUNDLE_BASE, "slots_comerciales": -1}]
    with pytest.raises(sm.SemanticConfigError):
        sm.validate_config(base_config)


def test_validate_config_rejects_non_numeric_segundos_comerciales_override(base_config):
    base_config["elemento_overrides"] = [{"elemento_id": "X", **_BUNDLE_BASE, "segundos_comerciales": "mucho"}]
    with pytest.raises(sm.SemanticConfigError):
        sm.validate_config(base_config)


def test_validate_config_rejects_bad_metric_policies(base_config):
    base_config["metric_policies"]["slot_seconds_no_aplica_metrics"] = []
    with pytest.raises(sm.SemanticConfigError):
        sm.validate_config(base_config)
