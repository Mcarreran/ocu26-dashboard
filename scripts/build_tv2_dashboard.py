"""Capa de datos para el dashboard TV2 - OCU26 (Core Comercial Digital).

Se ejecuta DESPUES de scripts/semantic_model.py y scripts/metrics_engine.py
(Gate 3B). Reutiliza export_data.load_pipeline (Gate 4A), mismo patron que
build_tv1_dashboard.py: no reabre el Excel, no reimplementa reglas de
negocio de Gate 3 (toda cifra sale de MetricsEngine.query()). No importa
build_tv1_dashboard.py a proposito: TV1 queda protegida/aislada, cada TV
es un builder independiente sobre el mismo pipeline compartido.

Universo TV2 = CORE COMERCIAL DIGITAL = Pantallas LED + Shoppings Digital
(Cencosud + Remeros) + AA2000 Digital. YPF/APSA/London Supply excluidos.
AA2000 alimenta los KPIs generales del Core pero no tiene tarjeta ni
ranking propio (spec TV2 Sec.4,19).

Uso:
    python scripts/build_tv2_dashboard.py
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from pathlib import Path
from typing import Any

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import validate_input as vi  # noqa: E402
from semantic_model import filter_universe  # noqa: E402
from metrics_engine import MetricsEngine  # noqa: E402
from export_data import load_pipeline  # noqa: E402

REPO_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = Path(__file__).resolve().parent / "templates" / "tv2_template.html"
REFERENCE_PATH = REPO_ROOT / "audit_sources" / "TV1_REFERENCE.html.html"
DEFAULT_OUTPUT_HTML = REPO_ROOT / "tv2.html"
DEFAULT_OUTPUT_JSON = REPO_ROOT / "output" / "tv2_data.json"

MESES_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]
MESES_ES_ABR = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

# Periodo de referencia TV2 (spec Sec.3): Julio 2026, vs Junio 2026, evolucion Ene-Jul.
REPORT_YEAR = 2026
REPORT_MONTH = 7

# Universo TV2 (spec Sec.4-6): Core Comercial Digital = Pantallas LED +
# Shoppings Digital (Cencosud+Remeros) + AA2000 Digital. YPF/APSA/London
# nunca entran (ni siquiera se listan aqui). Cada ElementoID pertenece a
# una sola familia funcional, determinada por CircuitoNegocio (nunca por
# nombre de sitio/ubicacion): evita que "Pantalla Led" ubicada en Remeros
# se confunda con "Shoppings Remeros" (son CircuitoNegocio distintos).
PANTALLAS_CIRCUITOS = ["PANTALLAS_LED"]
SHOPPINGS_CIRCUITOS = ["CENCOSUD", "REMEROS"]
AA2000_CIRCUITOS = ["AA2000"]
TV2_CIRCUITOS = PANTALLAS_CIRCUITOS + SHOPPINGS_CIRCUITOS + AA2000_CIRCUITOS
FAMILY_MAP = {
    "PANTALLAS_LED": "Pantallas",
    "CENCOSUD": "Shoppings",
    "REMEROS": "Shoppings",
    "AA2000": "AA2000",
}

RANKING_PANTALLAS_TOP_N = 5
RANKING_SHOPPINGS_TOP_N = 3


class BuildError(Exception):
    """Error bloqueante al construir el dashboard TV2."""


def _period_bounds(year: int, month: int) -> tuple[str, str]:
    start = pd.Timestamp(year=year, month=month, day=1)
    end = start + pd.offsets.MonthEnd(0)
    return str(start.date()), str(end.date())


def _previous_month(year: int, month: int) -> tuple[int, int]:
    ts = pd.Timestamp(year=year, month=month, day=1) - pd.DateOffset(months=1)
    return ts.year, ts.month


def _round1(value: Any) -> float | None:
    if value is None or (isinstance(value, float) and pd.isna(value)) or value is pd.NA:
        return None
    return round(float(value), 1)


def _fmt_es_int(n: int) -> str:
    return f"{n:,}".replace(",", ".")


# ---------------------------------------------------------------------------
# Universo TV2
# ---------------------------------------------------------------------------


def build_tv2_universe(semantic_result: dict[str, Any]) -> dict[str, Any]:
    maestro = semantic_result["maestro"]
    config = semantic_result["config"]

    op = filter_universe(maestro, "OPERATIVO_GENERAL", config)
    tv2_maestro = op[op["CircuitoNegocio"].isin(TV2_CIRCUITOS) & (op["Medio"] == "Digital")].copy()
    if tv2_maestro.empty:
        raise BuildError("Universo TV2 vacio: revisar business_semantics.json / Medio=Digital")

    return {
        "maestro": tv2_maestro,
        "circuitos": sorted(tv2_maestro["CircuitoNegocio"].unique().tolist()),
        "element_ids": tv2_maestro["ElementoID"].tolist(),
    }


# ---------------------------------------------------------------------------
# KPI 1 - Ocupacion por calendario (Core Digital)
# ---------------------------------------------------------------------------


def _distinct_activos_registrados(
    engine: MetricsEngine, circuitos: list[str], period: tuple[str, str], previous: tuple[str, str],
) -> dict[str, Any]:
    """COUNT DISTINCT ElementoID digital con campana / COUNT DISTINCT
    ElementoID digital elegible, scope=circuitos (spec Sec.8: misma
    definicion canonica que TV1 'Digital por calendario', recalculada con
    scope TV2)."""
    reg = engine.query("elementos_registrados", filters={"CircuitoNegocio": circuitos, "Medio": "Digital"})
    reg_value = int(reg["Value"].iloc[0]) if len(reg) else 0

    act = engine.query("elementos_con_actividad", filters={"CircuitoNegocio": circuitos, "Medio": "Digital"}, start_date=period[0], end_date=period[1])
    act_value = int(act["Value"].iloc[0]) if len(act) else 0

    act_prev = engine.query("elementos_con_actividad", filters={"CircuitoNegocio": circuitos, "Medio": "Digital"}, start_date=previous[0], end_date=previous[1])
    act_prev_value = int(act_prev["Value"].iloc[0]) if len(act_prev) else 0

    pct_actual = _round1(act_value / reg_value * 100.0) if reg_value else None
    pct_anterior = _round1(act_prev_value / reg_value * 100.0) if reg_value else None
    delta_pp = _round1(pct_actual - pct_anterior) if pct_actual is not None and pct_anterior is not None else None

    return {
        "activos": act_value,
        "elegibles": reg_value,
        "anterior_activos": act_prev_value,
        "pct_actual": pct_actual,
        "pct_anterior": pct_anterior,
        "delta_pp": delta_pp,
    }


def compute_kpi1_ocupacion_calendario(engine: MetricsEngine, period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    return _distinct_activos_registrados(engine, TV2_CIRCUITOS, period, previous)


# ---------------------------------------------------------------------------
# Fill rate (generico, reutilizado para KPI general / Pantallas / Shoppings / AA2000)
# ---------------------------------------------------------------------------


def _fill_family(engine: MetricsEngine, circuitos: list[str], period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    """slots ocupados / capacidad aplicable (spec Sec.9), scope=circuitos.
    fill_rate_slots ya escopea Medio=Digital internamente (metrics_engine
    _compute_digital_metric). Preserva MetricStatus, nunca convierte
    NO_APLICA/PARTIAL a cero."""
    actual = engine.query("fill_rate_slots", filters={"CircuitoNegocio": circuitos}, start_date=period[0], end_date=period[1])
    anterior = engine.query("fill_rate_slots", filters={"CircuitoNegocio": circuitos}, start_date=previous[0], end_date=previous[1])
    a = actual.iloc[0]
    p = anterior.iloc[0]

    pct_actual = _round1(a["Value"]) if pd.notna(a["Value"]) else None
    pct_anterior = _round1(p["Value"]) if pd.notna(p["Value"]) else None
    delta_pp = _round1(pct_actual - pct_anterior) if pct_actual is not None and pct_anterior is not None else None
    numerador = a["Numerator"] if pd.notna(a["Numerator"]) else None
    denominador = a["Denominator"] if pd.notna(a["Denominator"]) else None

    return {
        "numerador": int(round(numerador)) if numerador is not None else None,
        "numerador_raw": float(numerador) if numerador is not None else None,
        "denominador": int(denominador) if denominador is not None else None,
        "pct_actual": pct_actual,
        "pct_anterior": pct_anterior,
        "delta_pp": delta_pp,
        "status": a["MetricStatus"],
        "status_anterior": p["MetricStatus"],
    }


def compute_kpi2_fill_digital(engine: MetricsEngine, period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    return _fill_family(engine, TV2_CIRCUITOS, period, previous)


def compute_pantallas_fill(engine: MetricsEngine, period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    """Fill rate de Pantallas (uso interno: reconciliacion del Core, KPI5
    y su desglose de disponibilidad, e insights). No es la tarjeta KPI3
    (esa es ocupacion por calendario, ver compute_kpi3_pantallas_ocupacion)."""
    return _fill_family(engine, PANTALLAS_CIRCUITOS, period, previous)


def compute_shoppings_fill(engine: MetricsEngine, period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    """Fill rate de Shoppings (mismo uso interno que compute_pantallas_fill)."""
    return _fill_family(engine, SHOPPINGS_CIRCUITOS, period, previous)


def compute_kpi3_pantallas_ocupacion(engine: MetricsEngine, period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    """Tarjeta KPI3 'Pantallas LED': ocupacion por calendario (misma
    definicion canonica que KPI1, scope=Pantallas)."""
    return _distinct_activos_registrados(engine, PANTALLAS_CIRCUITOS, period, previous)


def compute_kpi4_shoppings_ocupacion(engine: MetricsEngine, period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    """Tarjeta KPI4 'Shoppings Digital': ocupacion por calendario (misma
    definicion canonica que KPI1, scope=Shoppings, incluye Remeros)."""
    return _distinct_activos_registrados(engine, SHOPPINGS_CIRCUITOS, period, previous)


def compute_aa2000_fill(engine: MetricsEngine, period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    """AA2000 no tiene tarjeta ni ranking propio (spec Sec.4,19): solo se
    calcula para reconciliar el Core general y alimentar la apertura de
    disponibilidad por familia."""
    return _fill_family(engine, AA2000_CIRCUITOS, period, previous)


def _assert_family_reconciles_with_core(kpi2: dict, pant: dict, shop: dict, aa2000: dict) -> None:
    """Pantallas + Shoppings + AA2000 deben reconciliar con el fill general
    del Core (spec Sec.4 tabla D): el universo TV2 es una particion
    disjunta por CircuitoNegocio, asi que numerador/denominador deben
    sumar exacto. Si algun dia esto deja de cumplirse (ej. un circuito
    nuevo mal mapeado en FAMILY_MAP) el build debe fallar, no mostrar un
    Core que no reconcilia."""
    num_sum = sum(x["numerador_raw"] or 0.0 for x in (pant, shop, aa2000))
    den_sum = sum(x["denominador"] or 0 for x in (pant, shop, aa2000))
    if kpi2["denominador"] is not None and den_sum != kpi2["denominador"]:
        raise BuildError(f"Capacidad no reconcilia: familias={den_sum} vs Core={kpi2['denominador']}")
    if kpi2["numerador_raw"] is not None and abs(num_sum - kpi2["numerador_raw"]) > 0.01:
        raise BuildError(f"Slots ocupados no reconcilian: familias={num_sum} vs Core={kpi2['numerador_raw']}")


# ---------------------------------------------------------------------------
# KPI 5 - Slots disponibles (Core) + apertura por familia
# ---------------------------------------------------------------------------


def compute_kpi5_slots_disponibles(
    kpi2: dict[str, Any], pant: dict[str, Any], shop: dict[str, Any], aa2000: dict[str, Any],
) -> dict[str, Any]:
    """slots_disponibles = capacidad_total - slots_ocupados (spec Sec.16).
    Disponibilidad es inversa a fill: MENOS disponible = mejor utilizacion."""

    def _disponible(fam: dict[str, Any]) -> dict[str, Any]:
        if fam["denominador"] is None or fam["numerador_raw"] is None:
            return {"disponibles": None, "pct": None, "status": fam["status"]}
        disp_raw = fam["denominador"] - fam["numerador_raw"]
        return {
            "disponibles": int(round(disp_raw)),
            "pct": _round1(disp_raw / fam["denominador"] * 100.0) if fam["denominador"] else None,
            "status": fam["status"],
        }

    core_actual_disp = kpi2["denominador"] - kpi2["numerador_raw"] if kpi2["denominador"] is not None and kpi2["numerador_raw"] is not None else None
    pct_actual = _round1(core_actual_disp / kpi2["denominador"] * 100.0) if core_actual_disp is not None and kpi2["denominador"] else None

    # anterior: recalculado a partir de pct_anterior de fill (ya redondeado
    # de forma independiente, mismo patron que TV1 kpi5 delta_pp).
    den = kpi2["denominador"]
    pct_anterior = _round1(100.0 - kpi2["pct_anterior"]) if kpi2["pct_anterior"] is not None else None
    delta_pp = _round1(pct_actual - pct_anterior) if pct_actual is not None and pct_anterior is not None else None

    return {
        "disponibles": int(round(core_actual_disp)) if core_actual_disp is not None else None,
        "capacidad_total": den,
        "pct_actual": pct_actual,
        "pct_anterior": pct_anterior,
        "delta_pp": delta_pp,
        "apertura": {
            "pantallas": _disponible(pant),
            "shoppings": _disponible(shop),
            "aa2000": _disponible(aa2000),
        },
    }


# ---------------------------------------------------------------------------
# Rankings (Pantallas Top5, Shoppings Top3) - por SitioNegocio
# ---------------------------------------------------------------------------


def _ranking_por_sitio(engine: MetricsEngine, circuitos: list[str], period: tuple[str, str], top_n: int) -> list[dict[str, Any]]:
    """Fill rate + ocupacion calendario por SitioNegocio (dimension de
    sitio fisico, ortogonal a CircuitoNegocio: config.sitio_negocio).
    fill_rate_slots incluye TODOS los sitios registrados (incluso con 0
    slots ocupados); elementos_con_actividad solo devuelve sitios con
    >=1 elemento activo -> se hace left-join y se rellena 0 (nunca se
    excluye un sitio con actividad real 0, ver spec Sec.17 'NO usar
    NO_APLICA = 0': aqui el 0 es real, no ausencia de dato)."""
    fill_df = engine.query("fill_rate_slots", group_by=["SitioNegocio"], filters={"CircuitoNegocio": circuitos}, start_date=period[0], end_date=period[1])
    act_df = engine.query("elementos_con_actividad", group_by=["SitioNegocio"], filters={"CircuitoNegocio": circuitos, "Medio": "Digital"}, start_date=period[0], end_date=period[1])
    reg_df = engine.query("elementos_registrados", group_by=["SitioNegocio"], filters={"CircuitoNegocio": circuitos, "Medio": "Digital"})

    act_map = dict(zip(act_df["SitioNegocio"], act_df["Value"])) if len(act_df) else {}
    reg_map = dict(zip(reg_df["SitioNegocio"], reg_df["Value"])) if len(reg_df) else {}

    rows = []
    for _, r in fill_df.iterrows():
        sitio = r["SitioNegocio"]
        if sitio is None or (isinstance(sitio, float) and pd.isna(sitio)):
            continue
        elegibles = int(reg_map.get(sitio, 0))
        activos = int(act_map.get(sitio, 0))
        ocup_pct = _round1(activos / elegibles * 100.0) if elegibles else None
        fill_pct = _round1(r["Value"]) if pd.notna(r["Value"]) else None
        rows.append({
            "sitio": sitio,
            "fill_pct": fill_pct,
            "ocup_pct": ocup_pct,
            "elegibles": elegibles,
            "activos": activos,
            "status": r["MetricStatus"],
        })

    # Orden: fill DESC, ocupacion DESC, nombre ASC (spec Sec.17-18). NO_APLICA
    # (fill_pct None) va al final, nunca se trata como 0.
    rows.sort(key=lambda x: (
        x["fill_pct"] is None,
        -(x["fill_pct"] or 0),
        -(x["ocup_pct"] or 0),
        x["sitio"],
    ))
    return rows[:top_n]


def compute_ranking_pantallas(engine: MetricsEngine, period: tuple[str, str]) -> list[dict[str, Any]]:
    return _ranking_por_sitio(engine, PANTALLAS_CIRCUITOS, period, RANKING_PANTALLAS_TOP_N)


def compute_ranking_shoppings(engine: MetricsEngine, period: tuple[str, str]) -> list[dict[str, Any]]:
    return _ranking_por_sitio(engine, SHOPPINGS_CIRCUITOS, period, RANKING_SHOPPINGS_TOP_N)


# ---------------------------------------------------------------------------
# Evolucion mensual (fill rate Pantallas vs Shoppings, Ene-Jul)
# ---------------------------------------------------------------------------


def compute_evolution(engine: MetricsEngine, report_year: int, report_month: int) -> dict[str, Any]:
    """Fill rate mensual Pantallas LED vs Shoppings Digital (spec Sec.20).
    AA2000 no tiene serie propia (queda solo en los KPIs generales)."""
    meses, pantallas, shoppings = [], [], []
    for m in range(1, report_month + 1):
        start, end = _period_bounds(report_year, m)
        p = engine.query("fill_rate_slots", filters={"CircuitoNegocio": PANTALLAS_CIRCUITOS}, start_date=start, end_date=end).iloc[0]
        s = engine.query("fill_rate_slots", filters={"CircuitoNegocio": SHOPPINGS_CIRCUITOS}, start_date=start, end_date=end).iloc[0]
        meses.append(MESES_ES_ABR[m - 1])
        pantallas.append(_round1(p["Value"]) if pd.notna(p["Value"]) else None)
        shoppings.append(_round1(s["Value"]) if pd.notna(s["Value"]) else None)
    return {"meses": meses, "pantallas": pantallas, "shoppings": shoppings}


# ---------------------------------------------------------------------------
# Insights (Lectura / Punto positivo / A atender) - seleccion programatica
# ---------------------------------------------------------------------------


def _fmt_es_pct(v: float | None) -> str:
    if v is None:
        return "S/D"
    return f"{v:.1f}".replace(".", ",")


def compute_insights(
    kpi1: dict[str, Any], kpi2: dict[str, Any], kpi5: dict[str, Any],
    pant: dict[str, Any], shop: dict[str, Any],
    report_month_label: str, previous_month_label: str,
) -> dict[str, str]:
    lectura = (
        f"{report_month_label} registra {_fmt_es_int(kpi1['activos'])} de {_fmt_es_int(kpi1['elegibles'])} "
        f"elementos del Core Digital con campaña (<b>{_fmt_es_pct(kpi1['pct_actual'])}%</b>). El fill rate alcanza "
        f"<b>{_fmt_es_pct(kpi2['pct_actual'])}%</b> y quedan <b>{_fmt_es_int(kpi5['disponibles'])} slots disponibles</b>."
    )

    # Punto positivo: prioridad 1-4 = alguna mejora vs junio; si nada mejora,
    # prioridad 5 = mejor desempeño relativo entre familias (spec Sec.23).
    if kpi2["delta_pp"] is not None and kpi2["delta_pp"] > 0:
        punto_positivo = f"El fill rate del Core mejora <b>{_fmt_es_pct(kpi2['delta_pp'])} pp</b> frente a {previous_month_label.lower()}."
    elif pant["delta_pp"] is not None and pant["delta_pp"] > 0:
        punto_positivo = f"Pantallas LED mejora <b>{_fmt_es_pct(pant['delta_pp'])} pp</b> de fill frente a {previous_month_label.lower()}."
    elif shop["delta_pp"] is not None and shop["delta_pp"] > 0:
        punto_positivo = f"Shoppings Digital mejora <b>{_fmt_es_pct(shop['delta_pp'])} pp</b> de fill frente a {previous_month_label.lower()}."
    elif kpi1["delta_pp"] is not None and kpi1["delta_pp"] > 0:
        punto_positivo = f"La ocupación por calendario mejora <b>{_fmt_es_pct(kpi1['delta_pp'])} pp</b> frente a {previous_month_label.lower()}."
    else:
        punto_positivo = (
            f"Shoppings Digital concentra la mayor parte de la actividad del Core "
            f"(<b>{_fmt_es_int(shop.get('_activos', 0) or 0)}</b> elementos activos), sosteniendo el negocio del mes."
        )

    # A atender: prioridad 1 = mayor caida de fill entre familias.
    caidas = [("Pantallas LED", pant["delta_pp"]), ("Shoppings Digital", shop["delta_pp"])]
    caidas = [(nombre, d) for nombre, d in caidas if d is not None and d < 0]
    if caidas:
        caidas.sort(key=lambda x: x[1])
        nombre, delta = caidas[0]
        a_atender = f"El fill rate de {nombre} cae <b>{_fmt_es_pct(abs(delta))} pp</b> frente a {previous_month_label.lower()}."
    elif kpi1["delta_pp"] is not None and kpi1["delta_pp"] < 0:
        a_atender = f"La ocupación por calendario cae <b>{_fmt_es_pct(abs(kpi1['delta_pp']))} pp</b> frente a {previous_month_label.lower()}."
    else:
        a_atender = f"El Core Digital conserva <b>{_fmt_es_int(kpi5['disponibles'])} slots disponibles</b> por vender."

    return {"lectura": lectura, "punto_positivo": punto_positivo, "a_atender": a_atender}


# ---------------------------------------------------------------------------
# Logo (reutilizado byte-a-byte, igual patron que TV1)
# ---------------------------------------------------------------------------


def extract_logo_img_tag(reference_path: Path) -> str:
    html = reference_path.read_text(encoding="utf-8")
    marker = '<img src="data:image/png;base64,'
    start = html.find(marker)
    if start == -1:
        raise BuildError(f"No se encontro el logo embebido en {reference_path}")
    end = html.find(">", start)
    if end == -1:
        raise BuildError(f"Tag <img> de logo mal formado en {reference_path}")
    return html[start : end + 1]


# ---------------------------------------------------------------------------
# Orquestacion
# ---------------------------------------------------------------------------


def build_tv2_data(path: str | Path = vi.DEFAULT_INPUT_PATH) -> dict[str, Any]:
    path = Path(path)
    sha_before = vi.calculate_sha256(path)

    _transform_result, semantic_result, engine = load_pipeline(path)
    universe = build_tv2_universe(semantic_result)

    period = _period_bounds(REPORT_YEAR, REPORT_MONTH)
    prev_year, prev_month = _previous_month(REPORT_YEAR, REPORT_MONTH)
    previous = _period_bounds(prev_year, prev_month)

    kpi1 = compute_kpi1_ocupacion_calendario(engine, period, previous)
    kpi2 = compute_kpi2_fill_digital(engine, period, previous)
    pant_fill = compute_pantallas_fill(engine, period, previous)
    shop_fill = compute_shoppings_fill(engine, period, previous)
    aa2000 = compute_aa2000_fill(engine, period, previous)
    _assert_family_reconciles_with_core(kpi2, pant_fill, shop_fill, aa2000)
    kpi5 = compute_kpi5_slots_disponibles(kpi2, pant_fill, shop_fill, aa2000)

    shop_activos = engine.query("elementos_con_actividad", filters={"CircuitoNegocio": SHOPPINGS_CIRCUITOS, "Medio": "Digital"}, start_date=period[0], end_date=period[1])
    shop_fill["_activos"] = int(shop_activos["Value"].iloc[0]) if len(shop_activos) else 0
    pant_activos_df = engine.query("elementos_con_actividad", filters={"CircuitoNegocio": PANTALLAS_CIRCUITOS, "Medio": "Digital"}, start_date=period[0], end_date=period[1])
    pant_fill["_activos"] = int(pant_activos_df["Value"].iloc[0]) if len(pant_activos_df) else 0
    aa2000_activos_df = engine.query("elementos_con_actividad", filters={"CircuitoNegocio": AA2000_CIRCUITOS, "Medio": "Digital"}, start_date=period[0], end_date=period[1])
    aa2000["_activos"] = int(aa2000_activos_df["Value"].iloc[0]) if len(aa2000_activos_df) else 0

    # Tarjetas KPI3/KPI4: ocupacion por calendario (no fill), spec ajuste TV2.
    kpi3 = compute_kpi3_pantallas_ocupacion(engine, period, previous)
    kpi4 = compute_kpi4_shoppings_ocupacion(engine, period, previous)

    ranking_pantallas = compute_ranking_pantallas(engine, period)
    ranking_shoppings = compute_ranking_shoppings(engine, period)
    evolution = compute_evolution(engine, REPORT_YEAR, REPORT_MONTH)
    # Insights inferiores: sin cambios de negocio, siguen basados en fill
    # (pant_fill/shop_fill), no en las nuevas tarjetas de ocupacion calendario.
    insights = compute_insights(kpi1, kpi2, kpi5, pant_fill, shop_fill, MESES_ES[REPORT_MONTH - 1], MESES_ES[prev_month - 1])

    aa2000_elegibles = engine.query("elementos_registrados", filters={"CircuitoNegocio": AA2000_CIRCUITOS, "Medio": "Digital"})

    core_digital = {
        "elegibles": {
            "pantallas": kpi3["elegibles"],
            "shoppings": kpi4["elegibles"],
            "aa2000": int(aa2000_elegibles["Value"].iloc[0]) if len(aa2000_elegibles) else 0,
            "total": kpi1["elegibles"],
        },
        "con_campana": {
            "pantallas": kpi3["activos"],
            "shoppings": kpi4["activos"],
            "aa2000": aa2000["_activos"],
            "total": kpi1["activos"],
        },
        "aa2000_fill": aa2000,
        "pantallas_fill": pant_fill,
        "shoppings_fill": shop_fill,
    }

    reconciled_elegibles = (
        core_digital["elegibles"]["pantallas"] + core_digital["elegibles"]["shoppings"] + core_digital["elegibles"]["aa2000"]
    )
    if reconciled_elegibles != core_digital["elegibles"]["total"]:
        raise BuildError(f"Elegibles por familia no reconcilian con el Core: {reconciled_elegibles} vs {core_digital['elegibles']['total']}")
    reconciled_activos = (
        core_digital["con_campana"]["pantallas"] + core_digital["con_campana"]["shoppings"] + core_digital["con_campana"]["aa2000"]
    )
    if reconciled_activos != core_digital["con_campana"]["total"]:
        raise BuildError(f"Con-campana por familia no reconcilia con el Core: {reconciled_activos} vs {core_digital['con_campana']['total']}")

    sha_after = vi.calculate_sha256(path)
    if sha_after != sha_before:
        raise BuildError(
            f"ERROR CRITICO: el SHA-256 del input cambio durante la construccion del dashboard "
            f"(antes={sha_before}, despues={sha_after})."
        )

    data = {
        "meta": {
            "generado": dt.datetime.now().strftime("%d/%m/%Y %H:%M"),
            "report_year": REPORT_YEAR,
            "report_month": REPORT_MONTH,
            "report_month_label": MESES_ES[REPORT_MONTH - 1],
            "previous_month": prev_month,
            "previous_month_label": MESES_ES[prev_month - 1],
            "period_start": period[0],
            "period_end": period[1],
            "fuente": "OCU26 · Base maestra + base campañas",
        },
        "kpis": {
            "ocupacion_calendario": kpi1,
            "fill_digital": kpi2,
            "pantallas": kpi3,
            "shoppings": kpi4,
            "slots_disponibles": kpi5,
        },
        "core_digital": core_digital,
        "ranking_pantallas": ranking_pantallas,
        "ranking_shoppings": ranking_shoppings,
        "evolution": evolution,
        "insights": insights,
    }

    return {
        "data": data,
        "sha256": sha_after,
        "universe": {"circuitos": universe["circuitos"]},
    }


def render_html(data: dict[str, Any]) -> str:
    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    logo_tag = extract_logo_img_tag(REFERENCE_PATH)
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    html = template.replace("{{LOGO_IMG_TAG}}", logo_tag)
    html = html.replace("{{TV2_DATA_JSON}}", payload)
    return html


def build_and_write(
    path: str | Path = vi.DEFAULT_INPUT_PATH,
    output_html: str | Path = DEFAULT_OUTPUT_HTML,
    output_json: str | Path = DEFAULT_OUTPUT_JSON,
) -> dict[str, Any]:
    result = build_tv2_data(path)
    html = render_html(result["data"])

    output_html = Path(output_html)
    output_html.write_text(html, encoding="utf-8")

    output_json = Path(output_json)
    output_json.parent.mkdir(parents=True, exist_ok=True)
    with open(output_json, "w", encoding="utf-8") as fh:
        json.dump(result["data"], fh, ensure_ascii=False, indent=2)

    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Construye tv2.html (dashboard TV2 OCU26) con datos reales.")
    parser.add_argument("--file", default=str(vi.DEFAULT_INPUT_PATH), help="Ruta al archivo .xlsx a leer")
    parser.add_argument("--output-html", default=str(DEFAULT_OUTPUT_HTML), help="Ruta del HTML productivo generado")
    parser.add_argument("--output-json", default=str(DEFAULT_OUTPUT_JSON), help="Ruta del snapshot JSON generado")
    args = parser.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    try:
        result = build_and_write(args.file, args.output_html, args.output_json)
    except BuildError as exc:
        print("TV2_BUILD_ERROR:", exc)
        return 1

    print("TV2_BUILD_OK")
    print(json.dumps(result["data"]["meta"], ensure_ascii=False, indent=2))
    print(json.dumps(result["data"]["kpis"], ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
