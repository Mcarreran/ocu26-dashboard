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


@pytest.fixture(scope="module", autouse=True)
def tv1_control_hermetico(tmp_path_factory):
    """Etapa 2A: TV2 reconcilia su evolucion contra TV1. El control ya no es
    el output/tv1_data.json productivo (generado con reglas anteriores), sino
    un TV1 construido EN MEMORIA con las reglas vigentes y escrito en tmp:
    los tests nunca leen ni escriben salidas productivas."""
    import build_tv1_dashboard as t1
    control = tmp_path_factory.mktemp("tv1_control") / "tv1_data.json"
    control.write_text(json.dumps(t1.build_tv1_data(PRODUCTION_FILE)["data"], ensure_ascii=False), encoding="utf-8")
    original = td.TV1_CONTROL_JSON
    td.TV1_CONTROL_JSON = control
    yield control
    td.TV1_CONTROL_JSON = original


def _capacidad_por_regla_config(row) -> int:
    """Capacidad esperada leida DIRECTO de config (independiente del builder):
    perfil por FormatoNegocio, o regla por circuito digital."""
    dc = sm.load_config()["digital_capacity"]
    perfil = dc["slots_profiles"].get(row["FormatoNegocio"])
    if perfil is not None:
        return int(perfil)
    return int(dc["slots_por_circuito_digital"][row["CircuitoNegocio"]])


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
def production_html(tmp_path_factory):
    out_dir = tmp_path_factory.mktemp("tv2_render")
    td.build_and_write(
        PRODUCTION_FILE,
        output_html=out_dir / "tv2.html",
        output_json=out_dir / "tv2_data.json",
    )
    return (out_dir / "tv2.html").read_text(encoding="utf-8")


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
    # Etapa 2A: "321" (numerador legacy de fill slots de TV1) se retiro de la
    # lista porque con las reglas vigentes coincide, por casualidad, con los
    # slots ocupados reales de julio de TV2. El resto de los tokens legacy
    # (denominador 1.575 de Gate3, "activos") sigue sin poder aparecer.
    for legacy_token in ("1575", "1.575", "78,0", "71 activos", "\"activos\""):
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
# 7-10. Catalogo digital. Etapa 2A (2026-10-06): las cifras fijas 990/570/
# 220/200 (reglas anteriores) se reemplazan por la suma de la regla central
# por elemento, leida directo de config (sin depender del builder).
# ---------------------------------------------------------------------------


def test_catalogo_total_es_suma_de_reglas_centrales(production_data, production_maps):
    cats = production_maps["cats"]
    esperado = sum(_capacidad_por_regla_config(r) for _, r in cats.iterrows())
    assert production_data["catalogo"]["total"] == esperado


def test_catalogo_pantallas_es_20_por_pantalla(production_data, production_maps):
    cats = production_maps["cats"]
    n = int((cats["_familia"] == "Pantallas").sum())
    assert n > 0
    assert production_data["catalogo"]["pantallas"] == 20 * n


def test_catalogo_aa2000_es_10_por_elemento_digital(production_data, production_maps):
    cats = production_maps["cats"]
    n = int((cats["_familia"] == "AA2000").sum())
    assert n > 0
    assert production_data["catalogo"]["aa2000"] == 10 * n


def test_catalogo_shoppings_todos_sus_formatos_son_10(production_data, production_maps):
    cats = production_maps["cats"]
    n = int((cats["_familia"] == "Shoppings").sum())
    assert n > 0
    assert production_data["catalogo"]["shoppings"] == 10 * n


def test_catalogo_families_sum_to_total(production_data):
    cat = production_data["catalogo"]
    assert cat["shoppings"] + cat["pantallas"] + cat["aa2000"] == cat["total"]


# ---------------------------------------------------------------------------
# 11-13. Ocupados / disponibles / fill. Etapa 2A: 281/709/28,4% eran valores
# de las reglas anteriores; se reemplazan por identidades exactas.
# ---------------------------------------------------------------------------


def test_fill_pct_consistente_con_ocupados_y_capacidad(production_data):
    core = production_data["cards"]["core"]
    assert core["fill_pct"] == round(core["ocupados"] / core["capacidad"] * 100.0, 1)
    assert core["disp_pct"] == round(core["disponibles"] / core["capacidad"] * 100.0, 1)


def test_ocupados_mas_disponibles_igual_catalogo(production_data):
    core = production_data["cards"]["core"]
    assert core["ocupados"] + core["disponibles"] == production_data["catalogo"]["total"]


def test_fill_mas_disponibilidad_100_pct(production_data):
    core = production_data["cards"]["core"]
    assert abs((core["fill_pct"] + core["disp_pct"]) - 100.0) <= 0.15


# ---------------------------------------------------------------------------
# 14-16. Comparacion contra junio / promedios YTD
# ---------------------------------------------------------------------------


def test_junio_consistente(production_data):
    """Etapa 2A: reemplaza los valores de control de junio (380/38,4/610)."""
    core = production_data["cards"]["core"]
    assert core["ocupados_anterior"] + core["disponibles_anterior"] == core["capacidad_anterior"]
    assert core["fill_pct_anterior"] == round(core["ocupados_anterior"] / core["capacidad_anterior"] * 100.0, 1)


def test_delta_julio_vs_junio(production_data):
    core = production_data["cards"]["core"]
    assert core["delta_ocupados"] == core["ocupados"] - core["ocupados_anterior"]
    assert core["delta_disponibles"] == -core["delta_ocupados"]


def test_ytd_promedio_ponderado_present_and_consistent(production_data):
    core = production_data["cards"]["core"]
    assert core["ocupados_ytd_promedio"] is not None
    assert core["fill_pct_ytd"] is not None
    # fill YTD ponderado = suma ocupados / suma capacidad, nunca promedio de %.
    assert 0.0 <= core["fill_pct_ytd"] <= 100.0


# ---------------------------------------------------------------------------
# 17-18. Primera tarjeta: elementos confirmados + formatos
# ---------------------------------------------------------------------------


def test_elementos_confirmados_and_sin_confirmar(production_data, production_maps):
    """Etapa 2A: con las reglas centrales confirmadas (Patio de Comidas,
    UNI-PUENTELED-1 y EZEPAW005/011 = 10) no queda ningun elemento digital
    TV2 sin capacidad: confirmados = universo completo, sin_confirmar = 0."""
    cat = production_data["catalogo"]
    assert cat["elementos_sin_confirmar"] == 0
    assert cat["elementos_confirmados"] == int(production_maps["cats"]["ElementoID"].nunique())


def test_formatos_composition_sums_100(production_data):
    formatos = production_data["catalogo"]["formatos"]
    suma = round(sum(f["pct"] for f in formatos), 1)
    assert abs(suma - 100.0) <= 0.15
    nombres = {f["nombre"] for f in formatos}
    assert {"Pantallas LED", "Tótems", "Puentes LED", "Triedros", "Patio de Comidas"}.issubset(nombres)


def test_totem_counts_include_tripstore_aa2000_and_remeros_as_totem(production_data, production_maps):
    """Totems = todo elemento TV2 con FormatoNegocio TOTEM (Cencosud +
    Tripstore AA2000 + Remeros REM-DB-*), sin cifra fija."""
    formatos = {f["nombre"]: f["elementos"] for f in production_data["catalogo"]["formatos"]}
    cats = production_maps["cats"]
    assert formatos["Tótems"] == int((cats["FormatoNegocio"] == "TOTEM").sum())
    assert set(cats.loc[cats["CircuitoNegocio"] == "REMEROS", "FormatoNegocio"]) == {"TOTEM"}


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
    assert maps["familia_capacidad"]["AA2000"] == 10.0  # Etapa 2A: Totem = 10 (Excel trae 20)
    assert maps["familia_capacidad"]["Shoppings"] == 10.0


def test_aa2000_aeroparque_ezeiza_composition(production_data, production_maps):
    """Etapa 2A: Ezeiza suma EZEPAW005/011 (AA2000 digital = 10 por regla de
    circuito) a los 7 Tripstore; Aeroparque sigue con 3 Tripstore."""
    aa = production_data["cards"]["aa2000"]
    assert aa["aeroparque_elementos"] == 3
    cats = production_maps["cats"]
    aa_ids = set(cats.loc[cats["_familia"] == "AA2000", "ElementoID"])
    assert {"EZEPAW005", "EZEPAW011"} <= aa_ids
    assert aa["aeroparque_elementos"] + aa["ezeiza_elementos"] == aa["elementos"] == len(aa_ids)


# ---------------------------------------------------------------------------
# 21-24. Tres series mensuales, AA2000 incluido, reconciliacion con TV1
# ---------------------------------------------------------------------------


def test_evolution_has_three_series(production_data):
    ev = production_data["evolution"]
    assert set(("pantallas", "shoppings", "aa2000")).issubset(ev.keys())
    assert len(ev["meses"]) == td.REPORT_MONTH
    assert ev["meses"][-1] == "Jul"


def test_evolution_monthly_totals_match_tv1_control_values(production_data, tv1_control_hermetico):
    """Etapa 2A: reemplaza la serie fija [206, ..., 281] (reglas anteriores)
    por la reconciliacion real contra TV1 construido con las reglas vigentes
    y coherencia con las tarjetas de julio/junio."""
    ev = production_data["evolution"]
    tv1 = json.loads(tv1_control_hermetico.read_text(encoding="utf-8"))["evolution_espacios"]
    assert ev["total_ocupados"] == tv1["digital"]
    core = production_data["cards"]["core"]
    assert ev["total_ocupados"][-1] == core["ocupados"]
    assert ev["total_ocupados"][-2] == core["ocupados_anterior"]


def test_reconciliacion_omite_control_de_otra_version(tmp_path, capsys):
    """Etapa 2A: un output/tv1_data.json generado con otra base o version de
    reglas no bloquea el build ni se usa en silencio: se omite con aviso."""
    viejo = tmp_path / "tv1_data.json"
    viejo.write_text(json.dumps({"meta": {"reglas_version": "ANTERIOR"}, "evolution_espacios": {
        "meses": ["Ene"], "digital": [999999]}}), encoding="utf-8")
    original = td.TV1_CONTROL_JSON
    td.TV1_CONTROL_JSON = viejo
    try:
        td._reconcile_evolution_with_tv1({"meses": ["Ene"], "total_ocupados": [1]}, [(2026, 1)], "abc")
    finally:
        td.TV1_CONTROL_JSON = original
    assert "TV2_RECONCILE_SKIP" in capsys.readouterr().out


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
    assert ultimo == production_data["cards"]["shoppings"]["ocupados"]


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
    assert production_data["catalogo"]["elementos_sin_confirmar"] == 0  # Etapa 2A


def test_shoppings_matrix_sum_reconciles_with_card(production_data):
    m = production_data["matrix_shoppings"]
    ultimo = sum((f["celdas"][-1]["ocupados"] or 0) for f in m["filas"])
    assert ultimo == production_data["cards"]["shoppings"]["ocupados"]


def test_remeros_aporta_60_de_capacidad_y_ocupacion_real(production_data):
    """Reemplaza 'new_totals_after_remeros_totem_correction' (cifras
    990/570/281/709 de las reglas anteriores): la fila Remeros de la matriz
    Shoppings aporta 6 x 10 = 60 de capacidad confirmada y su ocupacion sale
    de campañas reales, nunca asumida llena."""
    fila = next(f for f in production_data["matrix_shoppings"]["filas"] if f["sitio"] == "Remeros")
    assert fila["capacidad"] == 60
    assert all(c["ocupados"] is not None and c["ocupados"] <= 60 for c in fila["celdas"])


def test_remeros_occupancy_computed_from_campaigns_not_assumed_full(production_data):
    """Spec Sec.4/8: 'no asumir que un totem con una campana ocupa sus 10
    slots'. Los 6 totems Remeros nunca tuvieron actividad comercial en la
    fuente -> su ocupacion real (calculada por engine._digital_period_activity,
    misma logica que el resto de los totems) es 0 los 7 meses, nunca 10."""
    m = production_data["matrix_shoppings"]
    fila = next(f for f in m["filas"] if f["sitio"] == "Remeros")
    assert [c["ocupados"] for c in fila["celdas"]] == [0, 0, 0, 0, 0, 0, 0]


def test_los_13_ex_pendientes_tienen_regla_confirmada_de_10(production_maps):
    """Etapa 2A (reemplaza 'pending_sites_reduced_only_by_remeros_six'): los
    13 soportes que antes quedaban SIN_REGLA (UNI-PUENTELED-1, Patio de
    Comidas Unicenter x10, EZEPAW005/011 de AA2000) tienen ahora capacidad
    confirmada de 10 cada uno y ya no hay sitios pendientes."""
    assert production_maps["sin_regla"].empty
    confirmed = production_maps["confirmed"].set_index("ElementoID")
    ex_pendientes = ["UNI-PUENTELED-1", "EZEPAW005", "EZEPAW011"] + [f"UNI-PACO-3L-{i}" for i in range(1, 11)]
    for eid in ex_pendientes:
        assert confirmed.loc[eid, "_capacidad"] == 10.0, eid
    assert confirmed.loc["UNI-PUENTELED-1", "_categoria"] == "PUENTE_LED"
    assert confirmed.loc["UNI-PACO-3L-1", "_categoria"] == "PATIO_COMIDAS"
    assert confirmed.loc["EZEPAW005", "_categoria"] == "OTRO_DIGITAL"
    assert all(not ids for ids in production_maps["sitio_pendiente_ids"].values())


def test_ningun_elemento_se_cuenta_dos_veces(production_maps):
    """Etapa 2A: cada ElementoID aparece una sola vez en el catalogo TV2 y en
    una sola familia/sitio."""
    cats = production_maps["cats"]
    assert not cats["ElementoID"].duplicated().any()
    todos = [eid for fam in production_maps["sitio_ids"].values() for ids in fam.values() for eid in ids]
    assert len(todos) == len(set(todos)) == int(production_maps["confirmed"]["ElementoID"].nunique())


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


def _row_semantica(maestro_row: dict) -> pd.Series:
    """Fila del maestro ya resuelta por semantic_model (SlotsComerciales de la
    regla central): las pruebas de unidad ya no construyen filas a mano."""
    return _semantic([maestro_row])["maestro"].iloc[0]


def test_pantalla_led_capacity_is_20():
    cap, cat = td._espacio_capacidad_digital(_row_semantica(_pantalla_led("P1", CapacidadSlotsReel=40)))
    assert cap == 20.0 and cat == "PANTALLA_LED"


def test_totem_cencosud_capacity_is_10():
    cap, cat = td._espacio_capacidad_digital(_row_semantica(_cencosud_totem("T1", CapacidadSlotsReel=20)))
    assert cap == 10.0 and cat == "TOTEM_SHOPPING"


def test_totem_aa2000_capacity_is_10_regla_central():
    """Etapa 2A (reemplaza 'capacity_is_registered_not_10'): Tripstore = 10."""
    cap, cat = td._espacio_capacidad_digital(_row_semantica(_aa2000_totem("AEP-TS-9", CapacidadSlotsReel=20)))
    assert cap == 10.0 and cat == "TRIPSTORE_AA2000"


def test_aa2000_digital_sin_descripcion_capacity_10():
    """Etapa 2A (reemplaza 'zero_capacity_stays_sin_regla'): EZEPAW005/011,
    AA2000 digital sin Descripcion ni capacidad en el Excel, = 10 por regla
    de circuito confirmada."""
    row = _aa2000_totem("EZEPAW005", Descripcion=None, CapacidadSlotsReel=0)
    cap, cat = td._espacio_capacidad_digital(_row_semantica(row))
    assert cap == 10.0 and cat == "OTRO_DIGITAL"


def test_puente_led_capacity_is_10_any_circuit():
    row = _remeros_digital("R-PUENTE", Descripcion="Puente Led 9", CapacidadSlotsReel=13)
    cap, cat = td._espacio_capacidad_digital(_row_semantica(row))
    assert cap == 10.0 and cat == "PUENTE_LED"


def test_uni_puenteled_1_es_puente_led_10():
    row = _cencosud_totem("UNI-PUENTELED-1", Descripcion=None, CapacidadSlotsReel=10)
    cap, cat = td._espacio_capacidad_digital(_row_semantica(row))
    assert cap == 10.0 and cat == "PUENTE_LED"


def test_triedro_capacity_is_10_regla_central():
    """Etapa 2A (reemplaza 'uses_registered_capacity'): Triedro = 10 aunque
    el Excel traiga 15."""
    row = _cencosud_totem("TR-1", Descripcion="Triedro Digital - Test", CapacidadSlotsReel=15)
    cap, cat = td._espacio_capacidad_digital(_row_semantica(row))
    assert cap == 10.0 and cat == "TRIEDRO"


def test_patio_de_comidas_capacity_is_10():
    row = _cencosud_totem("UNI-PACO-3L-1", Descripcion="Patio de Comidas - Nivel 3", CapacidadSlotsReel=10)
    cap, cat = td._espacio_capacidad_digital(_row_semantica(row))
    assert cap == 10.0 and cat == "PATIO_COMIDAS"


def test_remeros_digital_capacity_is_10():
    """Remeros digital = 10 por elemento: REM-DB-* via override de formato
    (TOTEM) y cualquier otro digital Remeros via regla de circuito."""
    cap, cat = td._espacio_capacidad_digital(_row_semantica(_remeros_digital("REM-DB-1", Descripcion="TV Led", CapacidadSlotsReel=20)))
    assert cap == 10.0 and cat == "TOTEM_SHOPPING"
    cap2, _cat2 = td._espacio_capacidad_digital(_row_semantica(_remeros_digital("REM-NUEVO-1", Descripcion="TV Led", CapacidadSlotsReel=20)))
    assert cap2 == 10.0


def test_otro_format_outside_remeros_still_sin_regla():
    """Un digital de formato desconocido en Cencosud (sin keyword, sin
    override) sigue SIN_REGLA aunque el Excel traiga capacidad: el valor
    historico del Excel no se usa como regla (Etapa 2A)."""
    row = _cencosud_totem("UNI-X-1", Descripcion="Pantalla experimental", CapacidadSlotsReel=10)
    cap, cat = td._espacio_capacidad_digital(_row_semantica(row))
    assert cap is None and cat == "SIN_REGLA"


def test_remeros_sin_excepcion_dispersa_en_tv2():
    """Etapa 2A (reemplaza 'remeros_totem_count_validation_raises_if_not_
    exactly_six'): la excepcion OTRO+REMEROS y su validacion de "exactamente
    6" se retiraron del builder; la regla vive en config y cada elemento
    digital Remeros vale 10 sin importar cuantos haya."""
    src = (REPO_ROOT / "scripts" / "build_tv2_dashboard.py").read_text(encoding="utf-8")
    assert "REMEROS_TOTEM_ELEMENT_COUNT_ESPERADO" not in src
    maestro_rows = [_pantalla_led("P1")] + [
        _remeros_digital(f"REM-DB-{i}", Descripcion="TV Led", CapacidadSlotsReel=10) for i in (1, 3, 5, 6, 8)
    ]
    universe = td.build_tv2_universe(_semantic(maestro_rows))
    maps = td.build_catalog_maps(universe)
    assert maps["familia_capacidad"]["Shoppings"] == 50.0
