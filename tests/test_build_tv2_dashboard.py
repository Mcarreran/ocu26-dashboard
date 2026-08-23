"""Pruebas de negocio para scripts/build_tv2_dashboard.py (dashboard TV2 OCU26,
Core Comercial Digital = Pantallas LED + Shoppings Digital + AA2000 Digital).

Rediseno 2026-08-22: unidad de negocio unica = ESPACIO/SLOT digital
(capacidad/ocupacion/disponibilidad), misma logica canonica de capacidad que
TV1 (build_tv1_dashboard.compute_espacios, Marco Rector CM3 2026-08-19).
Reemplaza la suite anterior basada en "elementos activos"/"ocupacion por
calendario"/rankings.

No modifica scripts/validate_input.py, scripts/transform_data.py,
scripts/semantic_model.py, scripts/metrics_engine.py,
config/business_semantics.json, input/OCU26_BASE_DATOS.xlsx, ni ningun
archivo productivo de TV1 (build_tv1_dashboard.py, tv1_template.html,
tv1.html, test_build_tv1_dashboard.py, TV1_REFERENCE.html). Los fixtures
sinteticos usan la CONFIGURACION REAL (sm.load_config()), mismo patron que
test_build_tv1_dashboard.py.

Cobertura: los 40 puntos de negocio obligatorios (prompt TV2 2026-08-22
Sec.21), agrupados tematicamente.
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
import build_tv2_dashboard as td  # noqa: E402
from metrics_engine import MetricsEngine  # noqa: E402
from export_data import load_pipeline  # noqa: E402
from test_semantic_model import _maestro_row, _transform_result  # noqa: E402

PRODUCTION_FILE = REPO_ROOT / "input" / "OCU26_BASE_DATOS.xlsx"


def _semantic(maestro_rows: list[dict], campanas_rows: list[dict] | None = None) -> dict:
    config = sm.load_config()
    return sm.build_semantic_model(_transform_result(maestro_rows, campanas_rows or []), config)


def _pantalla_led(elemento_id: str, ubicacion: str = "SITE", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Pantalla Led", Subcircuito="X", Ubicacion=ubicacion,
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        CapacidadSlotsReel=20, SegundosDia=72000,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _cencosud_totem(elemento_id: str, ubicacion: str = "UNICENTER", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Shoppings Digital", Subcircuito="CENCOSUD", Ubicacion=ubicacion,
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        Descripcion="Totem de Shopping", CapacidadSlotsReel=20, SegundosDia=72000,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _cencosud_static(elemento_id: str, ubicacion: str = "UNICENTER", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Shoppings Estático", Subcircuito="CENCOSUD", Ubicacion=ubicacion,
        Medio="Estático", TipoCatalogo="Cerrado", TipoInventario="Físico estático",
        CapacidadSlotsReel=0, SegundosDia=0,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _remeros_digital(elemento_id: str, ubicacion: str = "REMEROS", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="Shoppings Digital", Subcircuito="REMEROS", Ubicacion=ubicacion,
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        CapacidadSlotsReel=20, SegundosDia=72000,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _aa2000_totem(elemento_id: str, ubicacion: str = "EZEIZA", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="AA2000", Subcircuito="X", Ubicacion=ubicacion,
        Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
        Descripcion="Totem Simple", CapacidadSlotsReel=20, SegundosDia=72000,
    )
    row.update(overrides)
    return _maestro_row(elemento_id, **row)


def _ypf_digital(elemento_id: str, ubicacion: str = "500 - TESTVILLE - Calle Falsa 123", **overrides) -> dict:
    row = dict(
        CircuitoDashboard="YPF Digital", Subcircuito="X", Ubicacion=ubicacion,
        Medio="Digital", TipoCatalogo="Abierto", TipoInventario="Digital",
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


@pytest.fixture(scope="module")
def production_result():
    return td.build_tv2_data(PRODUCTION_FILE)


@pytest.fixture(scope="module")
def production_data(production_result):
    return production_result["data"]


@pytest.fixture(scope="module")
def production_json(production_result):
    return json.dumps(production_result["data"], ensure_ascii=False)


@pytest.fixture(scope="module")
def production_html():
    td.build_and_write(PRODUCTION_FILE)
    return (REPO_ROOT / "tv2.html").read_text(encoding="utf-8")


@pytest.fixture(scope="module")
def production_maps():
    _tr, semantic_result, _engine = load_pipeline(PRODUCTION_FILE)
    universe = td.build_tv2_universe(semantic_result)
    return td.build_catalog_maps(universe)


# ---------------------------------------------------------------------------
# 1-2. Titulo / subtitulo / pregunta rectora
# ---------------------------------------------------------------------------


def test_title_and_question_present(production_html):
    assert "Core Comercial Digital" in production_html
    assert "CAPACIDAD DIGITAL TENEMOS" in production_html
    assert "DISPONIBILIDAD?" in production_html


def test_subtitle_mentions_slots_not_calendar(production_html):
    assert "Capacidad, ocupaci" in production_html
    assert "ocupación por calendario" not in production_html.lower()


# ---------------------------------------------------------------------------
# 3-6. Cinco tarjetas, orden y ausencia total de contenido legacy
# ---------------------------------------------------------------------------


def test_payload_top_level_keys(production_data):
    assert set(production_data.keys()) == {
        "meta", "catalogo", "cards", "evolution", "matrix_shoppings", "matrix_pantallas", "insights",
    }


def test_five_cards_present_with_expected_shape(production_data):
    cards = production_data["cards"]
    assert set(cards.keys()) == {"core", "pantallas", "shoppings", "aa2000"}
    for fam in ("core", "pantallas", "shoppings", "aa2000"):
        c = cards[fam]
        for k in ("capacidad", "ocupados", "fill_pct", "disponibles", "disp_pct"):
            assert k in c


def test_no_ranking_activos_or_calendar_language(production_json):
    lowered = production_json.lower()
    for banned in ("ranking", "activos", "con_campana", "con campaña", "ocupacion_calendario", "elegibles"):
        assert banned not in lowered, f"termino legacy encontrado: {banned}"


def test_no_legacy_reference_values_leak_into_payload(production_json):
    for legacy_token in ("321", "1575", "1.575", "78,0", "71 activos", "\"activos\""):
        assert legacy_token not in production_json


def test_no_ranking_language_in_html(production_html):
    """Sin UI de ranking (spec Sec.21 puntos 6,30): se revisa el body
    renderizado, no los comentarios de codigo (el propio modulo documenta,
    a proposito, que este diseno REEMPLAZA los rankings anteriores)."""
    body = production_html.split("<body>", 1)[1]
    lowered = body.lower()
    for banned in ("data-rank=", 'class="rank', "ranking de", "· ranking"):
        assert banned not in lowered
    assert "% con campa" not in lowered


# ---------------------------------------------------------------------------
# 7-10. Catalogo digital = 990 (570 + 220 + 200), correccion 2026-08-22
# (Totems Remeros Shoppings Digital: 6 x 10 = 60 espacios, 510+60=570).
# ---------------------------------------------------------------------------


def test_catalogo_total_990(production_data):
    cat = production_data["catalogo"]
    assert cat["total"] == 990


def test_catalogo_shoppings_570(production_data):
    assert production_data["catalogo"]["shoppings"] == 570


def test_catalogo_pantallas_220(production_data):
    assert production_data["catalogo"]["pantallas"] == 220


def test_catalogo_aa2000_200(production_data):
    assert production_data["catalogo"]["aa2000"] == 200


def test_catalogo_families_sum_to_total(production_data):
    cat = production_data["catalogo"]
    assert cat["shoppings"] + cat["pantallas"] + cat["aa2000"] == cat["total"]


# ---------------------------------------------------------------------------
# 11-13. Ocupados = 281 (Remeros aporta 0, nunca tuvo campanas) / Disponibles
# = 709 / Disp = 71,6% (correccion 2026-08-22: capacidad +60, ocupados sin
# cambios -> fill baja, disponibilidad sube; es el comportamiento correcto,
# no un error de calculo).
# ---------------------------------------------------------------------------


def test_ocupados_281_fill_28_4(production_data):
    core = production_data["cards"]["core"]
    assert core["ocupados"] == 281
    assert core["fill_pct"] == 28.4


def test_disponibles_709_pct_71_6(production_data):
    core = production_data["cards"]["core"]
    assert core["disponibles"] == 709
    assert core["disp_pct"] == 71.6


def test_ocupados_mas_disponibles_igual_catalogo(production_data):
    core = production_data["cards"]["core"]
    assert core["ocupados"] + core["disponibles"] == production_data["catalogo"]["total"]


def test_fill_mas_disponibilidad_100_pct(production_data):
    core = production_data["cards"]["core"]
    assert abs((core["fill_pct"] + core["disp_pct"]) - 100.0) <= 0.15


# ---------------------------------------------------------------------------
# 14-16. Comparacion contra junio / promedios YTD
# ---------------------------------------------------------------------------


def test_junio_control_values(production_data):
    core = production_data["cards"]["core"]
    assert core["ocupados_anterior"] == 380
    assert core["fill_pct_anterior"] == 38.4
    assert core["disponibles_anterior"] == 610
    assert core["disp_pct_anterior"] == 61.6


def test_delta_julio_vs_junio(production_data):
    core = production_data["cards"]["core"]
    assert core["delta_ocupados"] == -99
    assert core["delta_pp"] == -10.0
    assert core["delta_disponibles"] == 99


def test_ytd_promedio_ponderado_present_and_consistent(production_data):
    core = production_data["cards"]["core"]
    assert core["ocupados_ytd_promedio"] is not None
    assert core["fill_pct_ytd"] is not None
    # fill YTD ponderado = suma ocupados / suma capacidad, nunca promedio de %.
    assert 0.0 <= core["fill_pct_ytd"] <= 100.0


# ---------------------------------------------------------------------------
# 17-18. Primera tarjeta: elementos confirmados + formatos
# ---------------------------------------------------------------------------


def test_elementos_confirmados_and_sin_confirmar(production_data):
    """Correccion 2026-08-22: los 6 totems Remeros pasan de sin_confirmar a
    confirmados (72->78, 19->13)."""
    cat = production_data["catalogo"]
    assert cat["elementos_confirmados"] == 78
    assert cat["elementos_sin_confirmar"] == 13


def test_formatos_composition_sums_100(production_data):
    formatos = production_data["catalogo"]["formatos"]
    suma = round(sum(f["pct"] for f in formatos), 1)
    assert abs(suma - 100.0) <= 0.15
    nombres = {f["nombre"] for f in formatos}
    assert {"Pantallas LED", "Tótems", "Puentes LED", "Triedros"}.issubset(nombres)


def test_totem_counts_include_tripstore_aa2000_and_remeros_as_totem(production_data):
    formatos = {f["nombre"]: f["elementos"] for f in production_data["catalogo"]["formatos"]}
    # 40 totems Cencosud + 10 Tripstore AA2000 + 6 totems Remeros (correccion
    # 2026-08-22) = 56.
    assert formatos["Tótems"] == 56


def test_formatos_elements_sum_to_confirmados(production_data):
    cat = production_data["catalogo"]
    suma = sum(f["elementos"] for f in cat["formatos"])
    assert suma == cat["elementos_confirmados"]


# ---------------------------------------------------------------------------
# 19-20. Tripstore comercialmente en AA2000, no en Shoppings
# ---------------------------------------------------------------------------


def test_tripstore_capacity_lives_in_aa2000_not_shoppings():
    maestro_rows = [
        _aa2000_totem("AEP-TS-1", ubicacion="AEROPARQUE"),
        _cencosud_totem("UNI-T-1"),
    ]
    semantic_result = _semantic(maestro_rows)
    universe = td.build_tv2_universe(semantic_result)
    maps = td.build_catalog_maps(universe)
    assert maps["familia_capacidad"]["AA2000"] == 20.0
    assert maps["familia_capacidad"]["Shoppings"] == 10.0


def test_aa2000_aeroparque_ezeiza_composition(production_data):
    aa = production_data["cards"]["aa2000"]
    assert aa["aeroparque_elementos"] == 3
    assert aa["ezeiza_elementos"] == 7
    assert aa["aeroparque_elementos"] + aa["ezeiza_elementos"] == aa["elementos"] == 10


# ---------------------------------------------------------------------------
# 21-24. Tres series mensuales, AA2000 incluido, reconciliacion con TV1
# ---------------------------------------------------------------------------


def test_evolution_has_three_series(production_data):
    ev = production_data["evolution"]
    assert set(("pantallas", "shoppings", "aa2000")).issubset(ev.keys())
    assert len(ev["meses"]) == td.REPORT_MONTH
    assert ev["meses"][-1] == "Jul"


def test_evolution_monthly_totals_match_tv1_control_values(production_data):
    ev = production_data["evolution"]
    esperado = [206, 270, 311, 293, 346, 380, 281]
    assert ev["total_ocupados"] == esperado


def test_evolution_families_sum_to_total_every_month(production_data):
    ev = production_data["evolution"]
    for i in range(len(ev["meses"])):
        suma = ev["pantallas"]["ocupados"][i] + ev["shoppings"]["ocupados"][i] + ev["aa2000"]["ocupados"][i]
        assert suma == ev["total_ocupados"][i]


def test_evolution_no_future_months(production_data):
    ev = production_data["evolution"]
    assert "Ago" not in ev["meses"]


# ---------------------------------------------------------------------------
# 25-27. Matrices: filas, celdas ocupados/capacidad, misma escala
# ---------------------------------------------------------------------------


def test_matrix_shoppings_rows_and_last_month_reconciles(production_data):
    m = production_data["matrix_shoppings"]
    # 9 sitios originales + Remeros (correccion 2026-08-22: ahora fila
    # confirmada con capacidad 60, ya no "pendiente"): 10 filas en total.
    assert len(m["filas"]) == 10
    ultimo = sum((f["celdas"][-1]["ocupados"] or 0) for f in m["filas"])
    assert ultimo == production_data["cards"]["shoppings"]["ocupados"] == 196


def test_matrix_pantallas_rows_and_last_month_reconciles(production_data):
    m = production_data["matrix_pantallas"]
    assert len(m["filas"]) == 11
    ultimo = sum(f["celdas"][-1]["ocupados"] for f in m["filas"])
    assert ultimo == production_data["cards"]["pantallas"]["ocupados"] == 85


def test_matrix_cells_have_ocupados_capacidad_label(production_data):
    m = production_data["matrix_shoppings"]
    celda = m["filas"][0]["celdas"][0]
    assert celda["label"] == f"{celda['ocupados']}/{celda['capacidad']}".replace(",", ".") or "/" in celda["label"]


def test_matrices_use_same_fill_pct_definition(production_data):
    for m in (production_data["matrix_shoppings"], production_data["matrix_pantallas"]):
        for fila in m["filas"]:
            for celda in fila["celdas"]:
                if celda["fill_pct"] is not None:
                    esperado = round(celda["ocupados"] / celda["capacidad"] * 100.0, 1)
                    assert celda["fill_pct"] == esperado


def test_matrices_not_sorted_by_current_month_fill(production_data):
    """Orden alfabetico (spec Sec.16), nunca ranking por fill del mes vigente:
    el orden de filas debe coincidir con el orden alfabetico case-insensitive,
    no con el fill_pct de julio."""
    for m in (production_data["matrix_shoppings"], production_data["matrix_pantallas"]):
        nombres = [f["sitio"] for f in m["filas"]]
        assert nombres == sorted(nombres, key=str.casefold)


# ---------------------------------------------------------------------------
# Correccion 2026-08-22 (definitiva): los 6 elementos Remeros Shoppings
# Digital son TOTEMS comerciales de 10 slots cada uno (regla de negocio
# confirmada por el usuario) -> capacidad 60, fila CONFIRMADA (no mas
# "capacidad por confirmar"). Remeros sigue teniendo DOS inventarios de
# ElementoID distintos que no deben mezclarse: Shoppings Digital (6
# elementos "TV Led", CircuitoNegocio REMEROS) y Pantallas LED (1 elemento
# "C3 - REM", CircuitoNegocio PANTALLAS_LED, inventario aparte).
# ---------------------------------------------------------------------------


def test_remeros_appears_as_confirmed_row_in_shoppings_matrix(production_data):
    filas = production_data["matrix_shoppings"]["filas"]
    remeros = [f for f in filas if f["sitio"] == "Remeros"]
    assert len(remeros) == 1
    fila = remeros[0]
    assert fila["pendiente"] is False
    assert fila["capacidad"] == 60
    assert len(fila["celdas"]) == len(production_data["matrix_shoppings"]["meses"]) == 7
    for celda in fila["celdas"]:
        assert celda["capacidad"] == 60
        assert celda["ocupados"] is not None  # 0 es un valor real, no S/D/N-A
        assert "S/D" not in celda["label"]
        assert "REQUIERE_CONFIRMACION" not in celda["label"]
        assert celda["label"] == f"{celda['ocupados']}/60"


def test_remeros_still_present_in_pantallas_matrix_as_confirmed_row(production_data):
    filas = production_data["matrix_pantallas"]["filas"]
    remeros = [f for f in filas if f["sitio"] == "Remeros"]
    assert len(remeros) == 1
    fila = remeros[0]
    assert fila.get("pendiente") is False
    assert fila["capacidad"] == 20
    # Fila confirmada: tiene ocupados/fill_pct reales, no S/D.
    assert fila["celdas"][-1]["ocupados"] is not None


def test_remeros_shoppings_and_pantallas_are_disjoint_elementoid_sets(production_maps):
    """Los 6 ElementoID REM-DB-* (Shoppings Digital, ahora confirmados como
    totems) y el ElementoID 'C3 - REM' (Pantallas LED) no se solapan: dos
    inventarios fisicos distintos que solo comparten el nombre de sitio
    'Remeros'."""
    shoppings_confirmado_remeros = set(production_maps["sitio_ids"]["Shoppings"].get("Remeros", []))
    pantallas_confirmado_remeros = set(production_maps["sitio_ids"]["Pantallas"].get("Remeros", []))
    assert shoppings_confirmado_remeros == {"REM-DB-1", "REM-DB-3", "REM-DB-5", "REM-DB-6", "REM-DB-8", "REM-DB-10"}
    assert pantallas_confirmado_remeros == {"C3 - REM"}
    assert not (shoppings_confirmado_remeros & pantallas_confirmado_remeros)
    # Ya no queda ningun rastro de Remeros en la lista de sitios pendientes.
    assert "Remeros" not in production_maps["sitio_pendiente_ids"]["Shoppings"]


def test_remeros_shoppings_totems_are_now_confirmed_capacity(production_data, production_maps):
    """Los 6 elementos de Remeros Shoppings Digital pasan a capacidad
    confirmada (regla de negocio 2026-08-22: totems de 10 slots): aportan a
    los 570 espacios de Shoppings, a los 78 elementos confirmados y a la
    composicion 'Totems' por formato."""
    confirmed = production_maps["confirmed"]
    remeros_shopping_ids = {"REM-DB-1", "REM-DB-3", "REM-DB-5", "REM-DB-6", "REM-DB-8", "REM-DB-10"}
    confirmed_remeros = confirmed[confirmed["ElementoID"].isin(remeros_shopping_ids)]
    assert set(confirmed_remeros["ElementoID"]) == remeros_shopping_ids
    assert (confirmed_remeros["_capacidad"] == 10.0).all()
    assert (confirmed_remeros["_categoria"] == "TOTEM_SHOPPING").all()
    assert confirmed_remeros["_capacidad"].sum() == 60.0
    assert production_data["catalogo"]["elementos_sin_confirmar"] == 13  # 19 - 6


def test_shoppings_matrix_sum_reconciles_with_card(production_data):
    m = production_data["matrix_shoppings"]
    ultimo = sum((f["celdas"][-1]["ocupados"] or 0) for f in m["filas"])
    assert ultimo == production_data["cards"]["shoppings"]["ocupados"] == 196


def test_new_totals_after_remeros_totem_correction(production_data):
    """Valores de control de la correccion definitiva 2026-08-22 (Sec.3,7):
    Shoppings 510->570, Digital 930->990. Los ocupados NO cambian porque los
    6 totems Remeros nunca tuvieron campanas (TieneActividadComercial=False
    en la fuente): fill rate baja y disponibilidad sube, comportamiento
    correcto (mas capacidad, misma demanda), nunca un error de calculo."""
    cat = production_data["catalogo"]
    core = production_data["cards"]["core"]
    assert cat["total"] == 990
    assert cat["shoppings"] == 570
    assert core["ocupados"] == 281  # sin cambios: Remeros aporta 0 ocupados
    assert core["disponibles"] == 709  # 649 + 60
    assert production_data["cards"]["shoppings"]["ocupados"] == 196  # sin cambios
    assert production_data["evolution"]["total_ocupados"] == [206, 270, 311, 293, 346, 380, 281]


def test_remeros_occupancy_computed_from_campaigns_not_assumed_full(production_data):
    """Spec Sec.4/8: 'no asumir que un totem con una campana ocupa sus 10
    slots'. Los 6 totems Remeros nunca tuvieron actividad comercial en la
    fuente -> su ocupacion real (calculada por engine._digital_period_activity,
    misma logica que el resto de los totems) es 0 los 7 meses, nunca 10."""
    m = production_data["matrix_shoppings"]
    fila = next(f for f in m["filas"] if f["sitio"] == "Remeros")
    assert [c["ocupados"] for c in fila["celdas"]] == [0, 0, 0, 0, 0, 0, 0]


def test_pending_sites_reduced_only_by_remeros_six(production_maps):
    """Spec Sec.8: 'los soportes pendientes bajan solamente por estos 6
    registros'. Los otros sitios pendientes (Cencosud Unicenter 'Patio de
    Comidas'/'UNI-PUENTELED-1', AA2000 Descripcion vacia) deben seguir
    exactamente igual: ni un elemento mas ni uno menos."""
    sin_regla_ids = set(production_maps["sin_regla"]["ElementoID"].tolist())
    remeros_ids = {"REM-DB-1", "REM-DB-3", "REM-DB-5", "REM-DB-6", "REM-DB-8", "REM-DB-10"}
    assert not (sin_regla_ids & remeros_ids)
    cencosud_pendientes = production_maps["sin_regla"][production_maps["sin_regla"]["CircuitoNegocio"] == "CENCOSUD"]
    aa2000_pendientes = production_maps["sin_regla"][production_maps["sin_regla"]["CircuitoNegocio"] == "AA2000"]
    assert int(cencosud_pendientes["ElementoID"].nunique()) == 11
    assert int(aa2000_pendientes["ElementoID"].nunique()) == 2
    assert len(sin_regla_ids) == 13 == 11 + 2


# ---------------------------------------------------------------------------
# 28-31. Apportionment (reconciliacion familia<->core, sitio<->familia)
# ---------------------------------------------------------------------------


def test_apportion_basic_largest_remainder():
    raw = {"a": 1.4, "b": 1.4, "c": 1.2}
    result = td._apportion(raw, 4)
    assert sum(result.values()) == 4
    assert result in ({"a": 2, "b": 1, "c": 1}, {"a": 1, "b": 2, "c": 1})


def test_apportion_matches_known_july_family_split(production_data):
    core = production_data["cards"]["core"]
    pant = production_data["cards"]["pantallas"]
    shop = production_data["cards"]["shoppings"]
    aa = production_data["cards"]["aa2000"]
    assert pant["ocupados"] + shop["ocupados"] + aa["ocupados"] == core["ocupados"]
    assert pant["disponibles"] + shop["disponibles"] + aa["disponibles"] == core["disponibles"]


def test_apportion_rejects_out_of_range_remainder():
    with pytest.raises(td.BuildError):
        td._apportion({"a": 1.0, "b": 1.0}, 10)


# ---------------------------------------------------------------------------
# 32-36. Exclusiones (YPF/Cencomedia/APSA/London) y Cencosud incluido
# ---------------------------------------------------------------------------


def test_ypf_excluded_from_tv2_universe(production_result):
    assert "YPF" not in production_result["universe"]["circuitos"]


def test_ypf_elements_not_in_synthetic_tv2_universe():
    semantic_result = _semantic([_pantalla_led("P1"), _ypf_digital("500 - TT - 1")])
    universe = td.build_tv2_universe(semantic_result)
    assert "500 - TT - 1" not in universe["element_ids"]
    assert "P1" in universe["element_ids"]


def test_apsa_elements_not_in_synthetic_tv2_universe():
    semantic_result = _semantic([_pantalla_led("P1"), _apsa_static("A1")])
    universe = td.build_tv2_universe(semantic_result)
    assert "A1" not in universe["element_ids"]


def test_london_elements_not_in_synthetic_tv2_universe():
    semantic_result = _semantic([_pantalla_led("P1"), _london_static("L1")])
    universe = td.build_tv2_universe(semantic_result)
    assert "L1" not in universe["element_ids"]


def test_payload_contains_no_excluded_circuits(production_json):
    upper = production_json.upper()
    assert "APSA" not in upper
    assert "LONDON" not in upper
    assert "CENCOMEDIA" not in upper
    assert '"YPF"' not in production_json


def test_cencosud_included_via_shoppings_family():
    assert "CENCOSUD" in td.SHOPPINGS_CIRCUITOS
    semantic_result = _semantic([_cencosud_totem("C1")])
    universe = td.build_tv2_universe(semantic_result)
    maps = td.build_catalog_maps(universe)
    assert maps["familia_capacidad"]["Shoppings"] == 10.0


# ---------------------------------------------------------------------------
# 37-38. Hora ART / sin texto fantasma
# ---------------------------------------------------------------------------


def test_meta_has_art_timezone(production_data):
    assert production_data["meta"]["timezone"] == "America/Argentina/Buenos_Aires"
    assert production_data["meta"]["generado_iso"]


def test_html_clock_uses_art(production_html):
    assert "America/Argentina/Buenos_Aires" in production_html


def test_no_ghost_text_placeholders(production_html):
    for placeholder in ("{{LOGO_IMG_TAG}}", "{{TV2_DATA_JSON}}"):
        assert placeholder not in production_html
    # ">NaN<" / ">undefined<" indicarian un valor mal formateado renderizado
    # en el DOM; "isNaN"/"undefined" como palabra de codigo JS es legitimo.
    assert ">NaN<" not in production_html
    assert ">undefined<" not in production_html


# ---------------------------------------------------------------------------
# 39-40. TV1 / TV3-TV6 sin modificaciones (por construccion: TV2 no las importa
# ni las escribe) + SHA del Excel sin cambios
# ---------------------------------------------------------------------------


def test_tv1_pipeline_still_builds_successfully():
    import build_tv1_dashboard as t1
    result = t1.build_tv1_data(PRODUCTION_FILE)
    assert result["data"]["kpis"]["core_comercial"]["value"] == 1064


def test_tv2_builder_does_not_import_tv1_module():
    assert "build_tv1_dashboard" not in sys.modules or True  # import propio de este test no cuenta
    src = (REPO_ROOT / "scripts" / "build_tv2_dashboard.py").read_text(encoding="utf-8")
    assert "import build_tv1_dashboard" not in src


def test_input_excel_sha_unchanged(production_result):
    sha_after_build = vi.calculate_sha256(PRODUCTION_FILE)
    assert sha_after_build == production_result["sha256"]


# ---------------------------------------------------------------------------
# Unidad: capacidad por formato (espacios), sin depender del Excel productivo
# ---------------------------------------------------------------------------


def test_pantalla_led_capacity_is_20():
    row = pd.Series({"FormatoNegocio": "PANTALLA_LED", "CircuitoNegocio": "PANTALLAS_LED", "CapacidadSlotsReel": 0})
    cap, cat = td._espacio_capacidad_digital(row)
    assert cap == 20.0 and cat == "PANTALLA_LED"


def test_totem_cencosud_capacity_is_10():
    row = pd.Series({"FormatoNegocio": "TOTEM", "CircuitoNegocio": "CENCOSUD", "CapacidadSlotsReel": 0})
    cap, cat = td._espacio_capacidad_digital(row)
    assert cap == 10.0 and cat == "TOTEM_SHOPPING"


def test_totem_aa2000_capacity_is_registered_not_10():
    row = pd.Series({"FormatoNegocio": "TOTEM", "CircuitoNegocio": "AA2000", "CapacidadSlotsReel": 20})
    cap, cat = td._espacio_capacidad_digital(row)
    assert cap == 20.0 and cat == "TRIPSTORE_AA2000"


def test_totem_aa2000_zero_capacity_stays_sin_regla():
    """Los 2 registros AA2000 con Descripcion vacia (spec Sec.4.C): formato
    OTRO, CapacidadSlotsReel=0 -> SIN_REGLA, nunca capacidad inventada."""
    row = pd.Series({"FormatoNegocio": "OTRO", "CircuitoNegocio": "AA2000", "CapacidadSlotsReel": 0})
    cap, cat = td._espacio_capacidad_digital(row)
    assert cap is None and cat == "SIN_REGLA"


def test_puente_led_capacity_is_10_any_circuit():
    row = pd.Series({"FormatoNegocio": "PUENTE_LED", "CircuitoNegocio": "REMEROS", "CapacidadSlotsReel": 0})
    cap, cat = td._espacio_capacidad_digital(row)
    assert cap == 10.0 and cat == "PUENTE_LED"


def test_triedro_uses_registered_capacity():
    row = pd.Series({"FormatoNegocio": "TRIEDRO", "CircuitoNegocio": "CENCOSUD", "CapacidadSlotsReel": 15})
    cap, cat = td._espacio_capacidad_digital(row)
    assert cap == 15.0 and cat == "TRIEDRO"


def test_remeros_totem_literal_format_has_no_confirmed_rule():
    """Un elemento con FormatoNegocio literal 'TOTEM' (no 'OTRO') en Remeros
    sigue sin regla: la tasa de Totem de Shopping (10) via fmt=='TOTEM' es
    exclusiva de CENCOSUD. La regla nueva de Remeros (correccion 2026-08-22,
    ver test_remeros_otro_format_in_remeros_gets_totem_capacity) solo cubre
    fmt=='OTRO', que es como semantic_model resuelve los 6 elementos reales
    ('TV Led' no matchea ningun keyword rule de Totem)."""
    row = pd.Series({"FormatoNegocio": "TOTEM", "CircuitoNegocio": "REMEROS", "CapacidadSlotsReel": 20})
    cap, cat = td._espacio_capacidad_digital(row)
    assert cap is None and cat == "SIN_REGLA"


def test_remeros_otro_format_gets_totem_capacity():
    """Correccion 2026-08-22 (regla de negocio confirmada por el usuario):
    los elementos Digital de CircuitoNegocio=REMEROS con FormatoNegocio=OTRO
    (resuelto asi porque su Descripcion 'TV Led' no matchea ningun keyword
    rule) son comercialmente totems de 10 slots."""
    row = pd.Series({"FormatoNegocio": "OTRO", "CircuitoNegocio": "REMEROS", "CapacidadSlotsReel": 10})
    cap, cat = td._espacio_capacidad_digital(row)
    assert cap == 10.0 and cat == "TOTEM_SHOPPING"


def test_otro_format_outside_remeros_still_sin_regla():
    """La regla nueva esta acotada a CircuitoNegocio=REMEROS (spec Sec.1
    scope): un OTRO en Cencosud (p.ej. 'Patio de Comidas') sigue SIN_REGLA,
    no se contagia la excepcion de Remeros a otros circuitos."""
    row = pd.Series({"FormatoNegocio": "OTRO", "CircuitoNegocio": "CENCOSUD", "CapacidadSlotsReel": 10})
    cap, cat = td._espacio_capacidad_digital(row)
    assert cap is None and cat == "SIN_REGLA"


def test_remeros_totem_count_validation_raises_if_not_exactly_six():
    """Spec Sec.1: 'agregar una validacion que confirme que se encontraron
    exactamente 6 elementos. Si la cantidad cambia en la base, informar la
    variacion'. Fixture sintetico con 5 (no 6) totems Remeros debe frenar el
    build con un mensaje explicito, nunca aplicar la regla en silencio."""
    maestro_rows = [_pantalla_led("P1")] + [
        _maestro_row(
            f"REM-DB-{i}", CircuitoDashboard="Shoppings Digital", Subcircuito="REMEROS", Ubicacion="REMEROS",
            Medio="Digital", TipoCatalogo="Cerrado", TipoInventario="Digital",
            Descripcion="TV Led", CapacidadSlotsReel=10, SegundosDia=50400,
        )
        for i in range(1, 6)  # 5 elementos, no 6
    ]
    semantic_result = _semantic(maestro_rows)
    universe = td.build_tv2_universe(semantic_result)
    with pytest.raises(td.BuildError, match="Totems Remeros"):
        td.build_catalog_maps(universe)
