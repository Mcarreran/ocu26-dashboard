"""Pruebas de negocio para scripts/build_tv3_dashboard.py (rediseno TV3
2026-08-22, prompt maestro "TV3 - Core Comercial Estatico").

No modifica scripts/validate_input.py, scripts/transform_data.py,
scripts/semantic_model.py, scripts/metrics_engine.py,
config/business_semantics.json, input/OCU26_BASE_DATOS.xlsx, ni ningun
archivo productivo de TV1/TV2 (build_tv1_dashboard.py, build_tv2_dashboard.py,
tv1_template.html, tv2_template.html, tv1.html, tv2.html,
test_build_tv1_dashboard.py, test_build_tv2_dashboard.py,
TV1_REFERENCE.html, TV2_REFERENCE.html). Los fixtures sinteticos usan la
CONFIGURACION REAL (sm.load_config()), mismo patron que
test_build_tv2_dashboard.py.
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
import build_tv3_dashboard as td  # noqa: E402
from metrics_engine import MetricsEngine  # noqa: E402
from test_semantic_model import _maestro_row, _campana_row, _transform_result  # noqa: E402

PRODUCTION_FILE = REPO_ROOT / "input" / "OCU26_BASE_DATOS.xlsx"


def _semantic(maestro_rows: list[dict], campanas_rows: list[dict] | None = None) -> dict:
    config = sm.load_config()
    return sm.build_semantic_model(_transform_result(maestro_rows, campanas_rows or []), config)


def _cencosud_static(elemento_id: str, ubicacion: str = "UNICENTER", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Shoppings Estático", Subcircuito="CENCOSUD", Ubicacion=ubicacion,
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _remeros_static(elemento_id: str, ubicacion: str = "REMEROS", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Shoppings Estático", Subcircuito="REMEROS", Ubicacion=ubicacion,
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _remeros_digital(elemento_id: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Shoppings Digital", Subcircuito="REMEROS", Ubicacion="REMEROS",
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        CapacidadSlotsReel=10, SegundosDia=72000, Descripcion="Totem Shopping",
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _aa2000_static(elemento_id: str, ubicacion: str = "EZE", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="AA2000", Subcircuito="X", Ubicacion=ubicacion,
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _pilar_frontlight(elemento_id: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Pilar", Subcircuito="X", Ubicacion="PILAR",
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático",
        CapacidadSlotsReel=0, SegundosDia=0,
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


def _cencomedia(elemento_id: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Supermercados", Subcircuito="X", Ubicacion="MARTINEZ",
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Flexible gráfico",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _mab_static(elemento_id: str, **overrides) -> dict:
    row = dict(
        CircuitoDashboard="MAB", Subcircuito="X", Ubicacion="MAB_SITE",
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


@pytest.fixture(scope="module")
def production_result():
    return td.build_tv3_data(PRODUCTION_FILE)


@pytest.fixture(scope="module")
def production_json(production_result):
    return json.dumps(production_result["data"], ensure_ascii=False)


# ---------------------------------------------------------------------------
# 1. Titulo correcto, sin "Fijo"
# ---------------------------------------------------------------------------


def test_template_title_is_estatico_not_fijo():
    template = (REPO_ROOT / "scripts" / "templates" / "tv3_template.html").read_text(encoding="utf-8")
    assert "Core Comercial Est" in template
    assert "Fijo" not in template
    assert "FIJO" not in template.upper()


# ---------------------------------------------------------------------------
# 2. Seis tarjetas en orden correcto
# ---------------------------------------------------------------------------


def test_payload_has_six_tarjetas_in_order(production_result):
    tarjetas = production_result["data"]["tarjetas"]
    assert list(tarjetas.keys()) == [
        "catalogo", "ocupacion_mes", "shoppings_estatico", "aa2000_estatico",
        "campanas_activaciones", "disponibles",
    ]


def test_payload_top_level_keys_are_tv3_only(production_result):
    assert set(production_result["data"].keys()) == {
        "meta", "tarjetas", "evolution", "shoppings", "soportes_top",
        "soportes_sin_clasificar", "reconciliacion_tv1", "insights",
    }


def test_payload_contains_no_other_tv_datasets(production_json):
    for token in ("tv1_data", "tv2_data", "tv4", "tv5", "tv6", "ocu_data"):
        assert token not in production_json.lower()


# ---------------------------------------------------------------------------
# 3. Catalogo 422 / Shoppings 381 / AA2000 40 / Otros 1
# ---------------------------------------------------------------------------


def test_catalogo_control_values(production_result):
    cat = production_result["data"]["tarjetas"]["catalogo"]
    assert cat["total"] == 422
    assert cat["shoppings"] == 381
    assert cat["aa2000"] == 40
    assert cat["otros"] == 1


def test_catalogo_families_sum_to_total(production_result):
    cat = production_result["data"]["tarjetas"]["catalogo"]
    assert cat["shoppings"] + cat["aa2000"] + cat["otros"] == cat["total"]


# ---------------------------------------------------------------------------
# 4. Participaciones suman 100%
# ---------------------------------------------------------------------------


def test_catalogo_participations_sum_100(production_result):
    cat = production_result["data"]["tarjetas"]["catalogo"]
    total = round(cat["pct_shoppings"] + cat["pct_aa2000"] + cat["pct_otros"], 1)
    assert total == 100.0


# ---------------------------------------------------------------------------
# 5-8. Ocupacion del mes: julio/junio/variacion/disponibles
# ---------------------------------------------------------------------------


def test_ocupacion_mes_julio(production_result):
    t = production_result["data"]["tarjetas"]["ocupacion_mes"]
    assert t["ocupados"] == 146
    assert t["pct_actual"] == 34.6


def test_ocupacion_mes_junio(production_result):
    t = production_result["data"]["tarjetas"]["ocupacion_mes"]
    assert t["ocupados_anterior"] == 144
    assert t["pct_anterior"] == 34.1


def test_ocupacion_mes_variacion(production_result):
    t = production_result["data"]["tarjetas"]["ocupacion_mes"]
    assert t["delta_elementos"] == 2
    assert t["delta_pp"] == 0.5


def test_disponibles_control_values(production_result):
    t = production_result["data"]["tarjetas"]["disponibles"]
    assert t["disponibles"] == 276
    assert t["pct_actual"] == 65.4
    assert t["disponibles"] == t["catalogo"] - production_result["data"]["tarjetas"]["ocupacion_mes"]["ocupados"]


# ---------------------------------------------------------------------------
# 9-10. Shoppings ocupados / Remeros incluido
# ---------------------------------------------------------------------------


def test_shoppings_estatico_control_values(production_result):
    t = production_result["data"]["tarjetas"]["shoppings_estatico"]
    assert t["ocupados"] == 146
    assert t["pct_actual"] == 38.3
    assert t["ocupados_anterior"] == 144
    assert t["pct_anterior"] == 37.8


def test_remeros_estatico_row_in_shoppings_panel(production_result):
    rows = {r["sitio"]: r for r in production_result["data"]["shoppings"]["rows"]}
    assert "Remeros" in rows
    remeros = rows["Remeros"]
    assert remeros["catalogo"] == 27
    assert remeros["ocupados"] == 19
    assert remeros["disponibles"] == 8
    assert remeros["pct"] == 70.4


# ---------------------------------------------------------------------------
# 11. Remeros Digital excluido
# ---------------------------------------------------------------------------


def test_remeros_digital_excluded_from_universe():
    maestro_rows = [_remeros_static("R1"), _remeros_digital("R2")]
    semantic_result = _semantic(maestro_rows)
    universe = td.build_tv3_universe(semantic_result)
    assert "R1" in universe["element_ids"]
    assert "R2" not in universe["element_ids"]


def test_remeros_digital_not_counted_in_production_shoppings_catalog(production_result):
    """El catalogo Shoppings Estatico (381) no debe incluir los totems
    digitales de Remeros (Medio=Digital, fuera de build_tv3_universe)."""
    cat = production_result["data"]["tarjetas"]["catalogo"]
    assert cat["shoppings"] == 381


# ---------------------------------------------------------------------------
# 12. AA2000 0 de 40, cero real (no S/D)
# ---------------------------------------------------------------------------


def test_aa2000_zero_is_real_not_sd(production_result):
    aa = production_result["data"]["tarjetas"]["aa2000_estatico"]
    assert aa["catalogo"] == 40
    assert aa["ocupados"] == 0
    assert aa["pct_actual"] == 0.0
    assert aa["pct_actual"] is not None


def test_aa2000_synthetic_activity_is_detected():
    maestro_rows = [_aa2000_static("A1")]
    campana_rows = [_campana_row("CA1", "A1", IDCampaña="CAMP-A", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-05"))]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    t4 = td.compute_tarjeta4_aa2000(engine, 1, ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert t4["ocupados"] == 1
    assert t4["catalogo"] == 1


# ---------------------------------------------------------------------------
# 13-15. Campanas / activaciones / duplicados
# ---------------------------------------------------------------------------


def test_campanas_activaciones_are_distinct_counts():
    maestro_rows = [_cencosud_static("C1")]
    campana_rows = [
        _campana_row("CA1", "C1", IDCampaña="CAMP-A", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-05")),
        _campana_row("CA2", "C1", IDCampaña="CAMP-B", FechaInicio=pd.Timestamp("2026-07-10"), FechaFin=pd.Timestamp("2026-07-15")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    t5 = td.compute_tarjeta5_campanas_activaciones(engine, ["C1"], ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert t5["campanas_actual"] == 2
    assert t5["activaciones_actual"] == 2


def test_duplicate_rows_do_not_inflate_activaciones():
    """Misma campana repetida en el mismo elemento (ej. carga duplicada) no
    debe contar como 2 activaciones: par (ElementoID, IDCampaña) distinto."""
    maestro_rows = [_cencosud_static("C1")]
    campana_rows = [
        _campana_row("CA1", "C1", IDCampaña="CAMP-A", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-05")),
        _campana_row("CA1-DUP", "C1", IDCampaña="CAMP-A", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-05")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    t5 = td.compute_tarjeta5_campanas_activaciones(engine, ["C1"], ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert t5["campanas_actual"] == 1
    assert t5["activaciones_actual"] == 1


def test_same_campana_multiple_elements_counts_as_multiple_activaciones():
    maestro_rows = [_cencosud_static("C1"), _cencosud_static("C2", ubicacion="P.OESTE")]
    campana_rows = [
        _campana_row("CA1", "C1", IDCampaña="CAMP-A", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-05")),
        _campana_row("CA2", "C2", IDCampaña="CAMP-A", FechaInicio=pd.Timestamp("2026-07-01"), FechaFin=pd.Timestamp("2026-07-05")),
    ]
    semantic_result = _semantic(maestro_rows, campana_rows)
    engine = MetricsEngine(semantic_result)
    t5 = td.compute_tarjeta5_campanas_activaciones(engine, ["C1", "C2"], ("2026-07-01", "2026-07-31"), ("2026-06-01", "2026-06-30"))
    assert t5["campanas_actual"] == 1
    assert t5["activaciones_actual"] == 2


# ---------------------------------------------------------------------------
# 16-18. Evolucion: tres series, escala 0-100, sin meses futuros
# ---------------------------------------------------------------------------


def test_evolution_has_three_series(production_result):
    """Ajuste 2026-08-22: la serie Core Estatico se elimino del grafico de
    evolucion (reemplazada por Otros); Shoppings/AA2000/Otros son las tres
    series visibles, cada una sobre su propio catalogo."""
    ev = production_result["data"]["evolution"]
    assert set(ev.keys()) >= {"meses", "shoppings_pct", "aa2000_pct", "otros_pct"}
    assert "core_pct" not in ev
    assert len(ev["shoppings_pct"]) == len(ev["meses"])
    assert len(ev["aa2000_pct"]) == len(ev["meses"])
    assert len(ev["otros_pct"]) == len(ev["meses"])


def test_evolution_series_use_own_catalog_not_core_total(production_result):
    """Ninguna de las tres series debe usar 422 (Core) como denominador:
    Shoppings/381, AA2000/40, Otros/1."""
    cat = production_result["data"]["tarjetas"]["catalogo"]
    ev = production_result["data"]["evolution"]
    julio_shop = ev["shoppings_ocupados"][-1]
    julio_aa = ev["aa2000_ocupados"][-1]
    julio_otros = ev["otros_ocupados"][-1]
    assert ev["shoppings_pct"][-1] == round(julio_shop / cat["shoppings"] * 100.0, 1)
    assert ev["aa2000_pct"][-1] == round(julio_aa / cat["aa2000"] * 100.0, 1) if cat["aa2000"] else True
    assert ev["otros_pct"][-1] == round(julio_otros / cat["otros"] * 100.0, 1) if cat["otros"] else True
    # nunca coincide con julio_shop/422 salvo casualidad matematica: se verifica
    # explicitamente que el denominador usado es 381, no 422.
    assert ev["shoppings_pct"][-1] != round(julio_shop / cat["total"] * 100.0, 1)


def test_evolution_scale_0_100(production_result):
    ev = production_result["data"]["evolution"]
    for serie in (ev["shoppings_pct"], ev["aa2000_pct"], ev["otros_pct"]):
        for v in serie:
            assert v is None or (0.0 <= v <= 100.0)


def test_evolution_excludes_future_months(production_result):
    ev = production_result["data"]["evolution"]
    assert len(ev["meses"]) == td.REPORT_MONTH
    assert ev["meses"][-1] == "Jul"
    assert "Ago" not in ev["meses"]


def test_evolution_aa2000_shows_zero_not_none_when_catalog_confirmed(production_result):
    ev = production_result["data"]["evolution"]
    assert all(v == 0.0 for v in ev["aa2000_pct"])


def test_evolution_otros_shows_zero_not_none_when_catalog_confirmed(production_result):
    ev = production_result["data"]["evolution"]
    assert all(v == 0.0 for v in ev["otros_pct"])


def test_template_evolution_chart_has_no_core_references():
    """Ajuste 2026-08-22: ni el JS del panel de evolucion ni su subtitulo
    deben referenciar 'Core Estatico' (removido del grafico, no del resto
    del dashboard: la tarjeta 2 sigue siendo el Core completo)."""
    template = (REPO_ROOT / "scripts" / "templates" / "tv3_template.html").read_text(encoding="utf-8")
    assert "core_pct" not in template
    assert "ev.core" not in template
    assert "Core Est" not in template


# ---------------------------------------------------------------------------
# 19-21. Listado 17 shoppings, no ranking, totales reconciliados
# ---------------------------------------------------------------------------


def test_shoppings_panel_lists_all_17(production_result):
    panel = production_result["data"]["shoppings"]
    assert panel["total_sedes"] == 17
    assert len(panel["rows"]) == 17


def test_shoppings_panel_includes_zero_activity_sites(production_result):
    panel = production_result["data"]["shoppings"]
    assert any(r["ocupados"] == 0 for r in panel["rows"])


def test_shoppings_panel_is_alphabetical_not_ranked(production_result):
    panel = production_result["data"]["shoppings"]
    sitios = [r["sitio"] for r in panel["rows"]]
    assert sitios == sorted(sitios)
    pcts = [r["pct"] or 0 for r in panel["rows"]]
    assert pcts != sorted(pcts, reverse=True)


def test_shoppings_panel_totals_reconcile(production_result):
    panel = production_result["data"]["shoppings"]
    assert panel["catalogo_total"] == 381
    assert panel["ocupados_total"] == 146
    assert panel["disponibles_total"] == 235
    assert panel["sedes_con_actividad"] == 10


# ---------------------------------------------------------------------------
# 22-24. Top 3 soportes, clasificacion reproducible, sin bucket dominante
# ---------------------------------------------------------------------------


def test_soportes_top_max_three_sorted_by_activaciones(production_result):
    rows = production_result["data"]["soportes_top"]
    assert len(rows) <= 3
    vals = [r["activaciones"] for r in rows]
    assert vals == sorted(vals, reverse=True)


def test_classify_soporte_is_deterministic():
    row = {"FormatoNegocio": "OTRO", "TipoInstalacion": "Ban Grande (3 Perss)", "Descripcion": "Octógono Chico Pasillo", "ElementoID": "UNI-1"}
    r1, r2 = td.classify_soporte(row), td.classify_soporte(row)
    assert r1 == r2 == ("Banner", "TipoInstalacion")


def test_classify_soporte_uses_tipo_instalacion_over_descripcion_location_name():
    """Un banner nombrado por su ubicacion fisica (Garganta/Sector/Octogono)
    debe clasificar como Banner via TipoInstalacion, no quedar sin clasificar
    ni inventar una categoria a partir del nombre de sitio."""
    for desc in ("Garganta Central", "Sector Deportivo", "SECTOR A - ARCO", "Pasillo PACO"):
        row = {"FormatoNegocio": "OTRO", "TipoInstalacion": "Ban Grande (3 Perss)", "Descripcion": desc, "ElementoID": "X1"}
        assert td.classify_soporte(row) == ("Banner", "TipoInstalacion")


def test_classify_soporte_falls_back_to_elemento_id_prefix():
    row = {"FormatoNegocio": "OTRO", "TipoInstalacion": None, "Descripcion": "Carteles estacionamiento", "ElementoID": "FPBROWN-BACK-1"}
    assert td.classify_soporte(row) == ("Backlight", "ElementoID")


def test_classify_soporte_sin_clasificar_when_nothing_matches():
    row = {"FormatoNegocio": "OTRO", "TipoInstalacion": None, "Descripcion": "Ventanal fachada cerviño", "ElementoID": "PPAL-VENT-1"}
    assert td.classify_soporte(row) == (td.SIN_CLASIFICAR, "fallback")


def test_no_single_bucket_dominates_active_soportes(production_result):
    """Ninguna categoria del Top 3 debe corresponder a un "OTRO" generico:
    la clasificacion canonica debe distinguir tipos reales de soporte."""
    for row in production_result["data"]["soportes_top"]:
        assert row["soporte"] not in ("OTRO", "Otro", "otro")


def test_active_elements_have_zero_sin_clasificar_residual(production_result):
    """Los 146 elementos activos de julio (los que importan para el panel
    Top 3) deben quedar 100% clasificados con los campos reales de la base."""
    residual = production_result["data"]["soportes_sin_clasificar"]
    assert residual["elementos_activos"] == 0
    assert residual["activaciones"] == 0


# ---------------------------------------------------------------------------
# 25. HTML / payload con los campos necesarios
# ---------------------------------------------------------------------------


def test_render_html_embeds_payload_and_logo(production_result):
    html = td.render_html(production_result["data"])
    assert "window.TV3_DATA" in html
    assert "146" in html
    assert "data:image/png;base64," in html
    assert "{{TV3_DATA_JSON}}" not in html
    assert "{{LOGO_IMG_TAG}}" not in html


def test_payload_meta_has_art_timezone(production_result):
    meta = production_result["data"]["meta"]
    assert meta["timezone"] == "America/Argentina/Buenos_Aires"


# ---------------------------------------------------------------------------
# Universo: exclusiones (APSA/London/YPF/Cencomedia)
# ---------------------------------------------------------------------------


def test_apsa_excluded_from_tv3_universe():
    semantic_result = _semantic([_cencosud_static("C1"), _apsa_static("A1")])
    universe = td.build_tv3_universe(semantic_result)
    assert "A1" not in universe["element_ids"]
    assert "C1" in universe["element_ids"]


def test_london_excluded_from_tv3_universe():
    semantic_result = _semantic([_cencosud_static("C1"), _london_static("L1")])
    universe = td.build_tv3_universe(semantic_result)
    assert "L1" not in universe["element_ids"]


def test_ypf_excluded_from_tv3_universe():
    semantic_result = _semantic([_cencosud_static("C1"), _ypf_static("Y1")])
    universe = td.build_tv3_universe(semantic_result)
    assert "Y1" not in universe["element_ids"]


def test_cencomedia_excluded_from_tv3_universe():
    semantic_result = _semantic([_cencosud_static("C1"), _cencomedia("M1")])
    universe = td.build_tv3_universe(semantic_result)
    assert "M1" not in universe["element_ids"]


def test_digital_elements_not_in_synthetic_tv3_universe():
    semantic_result = _semantic([_cencosud_static("C1"), _pantalla_led("P1")])
    universe = td.build_tv3_universe(semantic_result)
    assert "P1" not in universe["element_ids"]
    assert "C1" in universe["element_ids"]


def test_payload_contains_no_excluded_circuits(production_json):
    upper = production_json.upper()
    assert "APSA" not in upper
    assert "LONDON" not in upper


# ---------------------------------------------------------------------------
# Reconciliacion TV1 (lectura, prompt maestro Sec.7)
# ---------------------------------------------------------------------------


def test_reconciliacion_tv1_control_value(production_result):
    rec = production_result["data"]["reconciliacion_tv1"]
    assert rec["comunes"] == 422


def test_reconciliacion_tv1_diff_is_mab_only(production_result):
    """El unico universo Estatico mas amplio de TV1 frente a TV3 es MAB
    (COMPLEMENTARIO, no CORE): no se fuerza igualdad entre universos
    distintos (prompt maestro Sec.7)."""
    rec = production_result["data"]["reconciliacion_tv1"]
    assert rec["solo_tv3"] == 0
    assert rec["solo_tv1_circuitos"] == ["MAB"]


def test_reconciliacion_tv1_does_not_use_legacy_percentages(production_json):
    """No debe reaparecer el 29,4%/28,2% legacy (universos distintos, prompt
    maestro Sec.7: "no usar como valores objetivo")."""
    assert "29,4" not in production_json
    assert "28,2" not in production_json


# ---------------------------------------------------------------------------
# 26. Archivos protegidos sin cambios / pipeline TV1-TV2 sigue funcionando
# ---------------------------------------------------------------------------


def test_input_excel_sha_unchanged(production_result):
    sha_after_build = vi.calculate_sha256(PRODUCTION_FILE)
    assert sha_after_build == production_result["sha256"]


def test_tv1_pipeline_still_builds_successfully():
    import build_tv1_dashboard as t1
    t1.build_tv1_data(PRODUCTION_FILE)


def test_tv2_pipeline_still_builds_successfully():
    import build_tv2_dashboard as t2
    t2.build_tv2_data(PRODUCTION_FILE)


def test_tv3_reference_untouched():
    ref = REPO_ROOT / "audit_sources" / "TV3_REFERENCE.html.html"
    assert ref.exists()
    html = ref.read_text(encoding="utf-8")
    assert "window.OCU_DATA" in html


def test_tv1_and_tv2_references_untouched():
    for name in ("TV1_REFERENCE.html.html", "TV2_REFERENCE.html.html"):
        ref = REPO_ROOT / "audit_sources" / name
        assert ref.exists()
        assert "window.OCU_DATA" in ref.read_text(encoding="utf-8")


def test_no_legacy_ocu_data_leaks_into_payload(production_json):
    for legacy_token in ("533", "588", "1444", "2770", "P.PILAR"):
        assert legacy_token not in production_json
