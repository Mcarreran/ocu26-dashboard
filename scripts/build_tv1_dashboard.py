"""Capa de datos para el dashboard TV1 - OCU26 (post Gate 4).

Se ejecuta DESPUES de scripts/semantic_model.py y scripts/metrics_engine.py
(Gate 3B). Reutiliza export_data.load_pipeline (Gate 4A) para no reabrir el
Excel dos veces ni reimplementar la cadena Gate 1/2/3. No reabre el Excel
salvo para el control de SHA-256 read-only ya usado en el resto del
pipeline. No reimplementa ninguna regla de negocio de Gate 3: toda cifra
sale de MetricsEngine.query() o de metodos internos ya aprobados
(MetricsEngine._campanas_overlap, igual patron que export_data.py).

Responsabilidad de este modulo: producir window.TV1_DATA (JSON) y el HTML
productivo tv1.html a partir de scripts/templates/tv1_template.html. El
HTML resultante no calcula reglas de negocio: solo consume, formatea y
dibuja lo que este modulo ya resolvio.

Uso:
    python scripts/build_tv1_dashboard.py
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
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
TEMPLATE_PATH = Path(__file__).resolve().parent / "templates" / "tv1_template.html"
REFERENCE_PATH = REPO_ROOT / "audit_sources" / "TV1_REFERENCE.html.html"
DEFAULT_OUTPUT_HTML = REPO_ROOT / "tv1.html"
DEFAULT_OUTPUT_JSON = REPO_ROOT / "output" / "tv1_data.json"

MESES_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]
MESES_ES_ABR = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

# Periodo de referencia de TV1. Definido UNA sola vez aqui (prompt Sec.18):
# cambiar estos dos numeros re-genera todo el dashboard para otro mes/anio.
REPORT_YEAR = 2026
REPORT_MONTH = 7

# Exclusion absoluta de negocio (prompt Sec.17): APSA y London Supply nunca
# entran a TV1. No es una regla de Gate3B (alli London Supply SI cuenta para
# IncluyeConteoGeneral) sino una decision de scope propia de este dashboard.
TV1_EXCLUDED_CIRCUITOS = {"APSA", "LONDON_SUPPLY"}

FAMILY_MAP = {
    "CENCOSUD": "Shoppings",
    "REMEROS": "Shoppings",
    "PANTALLAS_LED": "Pantallas",
    "YPF": "YPF",
    "AA2000": "AA2000",
    "CENCOMEDIA": "Cencomedia",
}
FAMILY_ORDER = ["Shoppings", "Pantallas", "YPF", "AA2000", "Cencomedia", "Otros"]


class BuildError(Exception):
    """Error bloqueante al construir el dashboard TV1."""


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


def _fmt_es_pct(v: float | None) -> str:
    if v is None:
        return "S/D"
    return f"{v:.1f}".replace(".", ",")


def _display_circuito(name: str) -> str:
    """Formato legible de un CircuitoNegocio real (nunca inventa un nombre
    nuevo, solo lo formatea): acronimos cortos (<=3 chars, ej. MAB) quedan
    igual; el resto pasa de SNAKE_CASE a Title Case (PILAR_FRONTLIGHT ->
    'Pilar Frontlight')."""
    if len(name) <= 3:
        return name
    return " ".join(w.capitalize() for w in name.split("_"))


def _normalize_token(s: Any) -> str | None:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return None
    s = str(s).strip().upper()
    return re.sub(r"\s+", " ", s)


def _ypf_station_key(elemento_id: Any, ubicacion: Any) -> str | None:
    """Grano comercial de venta YPF: 1 estacion = 1 unidad comercial (prompt
    correccion APIE Sec.1-2). No existe una columna APIE en la fuente (Excel
    ni maestro Gate3B resuelto) -- se audito explicitamente y se decidio con
    el usuario, tras evidencia, un surrogate: (prefijo numerico de
    ElementoID, localidad normalizada de Ubicacion). El prefijo solo no era
    seguro (824/440 con evidencia fuerte de colision real de dos sitios
    fisicos distintos bajo el mismo prefijo); agregar la localidad separa
    ese caso y ~10 mas sin usar fuzzy matching. Nunca cae a ElementoID en
    silencio: si no se puede derivar (Ubicacion vacia o sin token de
    localidad), devuelve None y el llamador debe contabilizarlo aparte."""
    if elemento_id is None or (isinstance(elemento_id, float) and pd.isna(elemento_id)):
        return None
    prefix = str(elemento_id).split(" - ")[0].strip()
    if not prefix:
        return None
    parts = str(ubicacion).split(" - ") if ubicacion is not None and not (isinstance(ubicacion, float) and pd.isna(ubicacion)) else []
    town = _normalize_token(parts[1]) if len(parts) >= 2 else None
    if town is None:
        return None
    return f"{prefix}|{town}"


# ---------------------------------------------------------------------------
# Universos TV1
# ---------------------------------------------------------------------------


def build_tv1_universe(semantic_result: dict[str, Any]) -> dict[str, Any]:
    maestro = semantic_result["maestro"]
    config = semantic_result["config"]

    op = filter_universe(maestro, "OPERATIVO_GENERAL", config)
    tv1_maestro = op[~op["CircuitoNegocio"].isin(TV1_EXCLUDED_CIRCUITOS)].copy()
    tv1_circuitos = sorted(tv1_maestro["CircuitoNegocio"].unique().tolist())
    digital_circuitos = sorted(set(tv1_circuitos) - {"YPF"})
    non_ypf_circuitos = sorted(set(tv1_circuitos) - {"YPF"})
    tv1_element_ids = tv1_maestro["ElementoID"].tolist()

    if not tv1_circuitos:
        raise BuildError("Universo TV1 vacio tras excluir APSA/London Supply: revisar business_semantics.json")

    ypf_maestro = tv1_maestro[tv1_maestro["CircuitoNegocio"] == "YPF"].copy()
    ypf_maestro["StationKey"] = [
        _ypf_station_key(eid, ubic) for eid, ubic in zip(ypf_maestro["ElementoID"], ypf_maestro["Ubicacion"])
    ]
    ypf_null_station = ypf_maestro[ypf_maestro["StationKey"].isna()]
    if len(ypf_null_station):
        raise BuildError(
            f"{len(ypf_null_station)} elemento(s) YPF sin estacion (surrogate temporal) derivable a partir de "
            f"ElementoID/Ubicacion: {sorted(ypf_null_station['ElementoID'].unique().tolist())[:10]}. "
            f"No se aplica fallback silencioso a ElementoID."
        )
    ypf_station_map = dict(zip(ypf_maestro["ElementoID"], ypf_maestro["StationKey"]))
    ypf_medio_map = dict(zip(ypf_maestro["ElementoID"], ypf_maestro["Medio"]))
    ypf_element_ids = ypf_maestro["ElementoID"].tolist()
    ypf_station_catalog_count = int(ypf_maestro["StationKey"].nunique())

    return {
        "maestro": tv1_maestro,
        "circuitos": tv1_circuitos,
        "digital_circuitos": digital_circuitos,
        "non_ypf_circuitos": non_ypf_circuitos,
        "element_ids": tv1_element_ids,
        "ypf_element_ids": ypf_element_ids,
        "ypf_station_map": ypf_station_map,
        "ypf_medio_map": ypf_medio_map,
        "ypf_station_catalog_count": ypf_station_catalog_count,
    }


def _ypf_active_stations(
    engine: MetricsEngine, ypf_element_ids: list[Any], ypf_station_map: dict[Any, str], start: str, end: str,
) -> set[str]:
    """Estaciones YPF (surrogate temporal) con >=1 ElementoID activo en el
    periodo. Reutiliza MetricsEngine._campanas_overlap (misma logica
    temporal aprobada de Gate3B) para resolver actividad a nivel
    ElementoID, y luego colapsa al grano de estacion."""
    if not ypf_element_ids:
        return set()
    overlap = engine._campanas_overlap(ypf_element_ids, start, end)
    active_ids = overlap["ElementoID"].dropna().unique().tolist()
    return {ypf_station_map[eid] for eid in active_ids if eid in ypf_station_map}


# ---------------------------------------------------------------------------
# KPI 1 - Core comercial
# ---------------------------------------------------------------------------


def compute_kpi1_core(universe: dict[str, Any]) -> dict[str, Any]:
    """Grano comercial mixto (correccion APIE Sec.3): No YPF cuenta
    ElementoID distinct; YPF cuenta estaciones (surrogate de estación YPF) distinct.
    Estructural: no depende de periodo."""
    maestro = universe["maestro"]
    core_no_ypf = maestro[(maestro["PortfolioTier"] == "CORE") & (maestro["CircuitoNegocio"] != "YPF")]
    non_ypf_value = int(core_no_ypf["ElementoID"].nunique())
    ypf_stations = universe["ypf_station_catalog_count"]
    return {
        "value": non_ypf_value + ypf_stations,
        "non_ypf": non_ypf_value,
        "ypf_estaciones": ypf_stations,
    }


# ---------------------------------------------------------------------------
# KPI 2 - Campanas unicas (IDCampana distinct, overlap temporal)
# ---------------------------------------------------------------------------


def _distinct_campanas(engine: MetricsEngine, element_ids: list[Any], start: str, end: str) -> int:
    overlap = engine._campanas_overlap(element_ids, start, end)
    overlap = overlap[overlap["IDCampaña"].notna() & (overlap["IDCampaña"].astype(str).str.strip() != "")]
    return int(overlap["IDCampaña"].nunique())


def compute_kpi2_campanas(
    engine: MetricsEngine, universe: dict[str, Any],
    ytd_start: str, period: tuple[str, str], previous: tuple[str, str],
) -> dict[str, Any]:
    element_ids = universe["element_ids"]
    ytd = _distinct_campanas(engine, element_ids, ytd_start, period[1])
    actual = _distinct_campanas(engine, element_ids, period[0], period[1])
    anterior = _distinct_campanas(engine, element_ids, previous[0], previous[1])
    return {
        "ytd": ytd,
        "mes_actual": actual,
        "mes_anterior": anterior,
        "label_kpi": "CAMPAÑAS ÚNICAS",
        "label_kpi_reason": (
            "Se implementó Campañas únicas en lugar de Campañas vendidas porque la fuente no "
            "contiene una fecha inequívoca de venta (FechaInicio/FechaFin son fechas de flight, "
            "FechaHoraCarga es fecha de carga en el sistema, no de venta)."
        ),
    }


# ---------------------------------------------------------------------------
# KPI 3 - Elementos con actividad
# ---------------------------------------------------------------------------


def compute_kpi3_actividad(
    engine: MetricsEngine, universe: dict[str, Any], period: tuple[str, str], previous: tuple[str, str], core_value: int,
) -> dict[str, Any]:
    """Unidades con campaña (correccion APIE Sec.4-5): No YPF cuenta
    ElementoID con campaña distinct (desglosado Digital/Estatico); YPF
    cuenta estaciones (surrogate temporal) con >=1 ElementoID con campaña
    valida. % del core usa el mismo grano mixto en numerador y
    denominador (nunca ElementoID YPF contra estacion, ni al reves)."""
    non_ypf_circuitos = universe["non_ypf_circuitos"]
    ypf_ids = universe["ypf_element_ids"]
    ypf_map = universe["ypf_station_map"]

    actual = engine.query("elementos_con_actividad", filters={"CircuitoNegocio": non_ypf_circuitos}, start_date=period[0], end_date=period[1])
    anterior = engine.query("elementos_con_actividad", filters={"CircuitoNegocio": non_ypf_circuitos}, start_date=previous[0], end_date=previous[1])
    non_ypf_actual = int(actual["Value"].iloc[0])
    non_ypf_anterior = int(anterior["Value"].iloc[0])

    dig_actual = engine.query("elementos_con_actividad", filters={"CircuitoNegocio": non_ypf_circuitos, "Medio": "Digital"}, start_date=period[0], end_date=period[1])
    dig_anterior = engine.query("elementos_con_actividad", filters={"CircuitoNegocio": non_ypf_circuitos, "Medio": "Digital"}, start_date=previous[0], end_date=previous[1])
    digital_actual = int(dig_actual["Value"].iloc[0])
    digital_anterior = int(dig_anterior["Value"].iloc[0])
    estatico_actual = non_ypf_actual - digital_actual
    estatico_anterior = non_ypf_anterior - digital_anterior

    ypf_estaciones_actual = len(_ypf_active_stations(engine, ypf_ids, ypf_map, period[0], period[1]))
    ypf_estaciones_anterior = len(_ypf_active_stations(engine, ypf_ids, ypf_map, previous[0], previous[1]))

    value = non_ypf_actual + ypf_estaciones_actual
    prev_value = non_ypf_anterior + ypf_estaciones_anterior
    pct_core = _round1(value / core_value * 100.0) if core_value else None
    return {
        "value": value,
        "anterior": prev_value,
        "delta": value - prev_value,
        "pct_core": pct_core,
        "non_ypf_activo": non_ypf_actual,
        "ypf_estaciones_activas": ypf_estaciones_actual,
        "ypf_estaciones_anterior": ypf_estaciones_anterior,
        "digital_activo": digital_actual,
        "digital_anterior": digital_anterior,
        "estatico_activo": estatico_actual,
        "estatico_anterior": estatico_anterior,
    }


# ---------------------------------------------------------------------------
# KPI YPF - estaciones con campaña + mix descriptivo de presencia por Medio
# ---------------------------------------------------------------------------


def compute_kpi_ypf(
    engine: MetricsEngine, universe: dict[str, Any], period: tuple[str, str], previous: tuple[str, str], unidades_campana_total: int,
) -> dict[str, Any]:
    """Estaciones YPF (surrogate temporal) con campaña valida en el
    periodo, mas una nota descriptiva de mix de presencia Estatico/Digital
    POR ESTACION (no por ElementoID): una estacion con ambos tipos de
    formato sigue contando 1 sola vez en 'estaciones_activas'; el mix es
    solo presencia (puede sumar >100% si hay solapamiento, nunca reparte
    el total)."""
    ypf_ids = universe["ypf_element_ids"]
    station_map = universe["ypf_station_map"]
    medio_map = universe["ypf_medio_map"]

    def _stations_and_mix(start: str, end: str) -> tuple[int, int, int]:
        if not ypf_ids:
            return 0, 0, 0
        overlap = engine._campanas_overlap(ypf_ids, start, end)
        active_ids = overlap["ElementoID"].dropna().unique().tolist()
        by_station: dict[str, set[str]] = {}
        for eid in active_ids:
            station = station_map.get(eid)
            if station is None:
                continue
            by_station.setdefault(station, set()).add(medio_map.get(eid))
        total = len(by_station)
        n_estatico = sum(1 for medios in by_station.values() if "Estático" in medios)
        n_digital = sum(1 for medios in by_station.values() if "Digital" in medios)
        return total, n_estatico, n_digital

    actual_total, actual_estatico, actual_digital = _stations_and_mix(period[0], period[1])
    anterior_total, _, _ = _stations_and_mix(previous[0], previous[1])

    return {
        "estaciones_activas": actual_total,
        "anterior": anterior_total,
        "delta": actual_total - anterior_total,
        "pct_sobre_unidades_campana": _round1(actual_total / unidades_campana_total * 100.0) if unidades_campana_total else None,
        "mix_estatico_pct": _round1(actual_estatico / actual_total * 100.0) if actual_total else None,
        "mix_digital_pct": _round1(actual_digital / actual_total * 100.0) if actual_total else None,
    }


# ---------------------------------------------------------------------------
# KPI 4 - Estatico (ocupacion calendario, solo universo elegible)
# ---------------------------------------------------------------------------


def _eligible_static_occupancy(
    engine: MetricsEngine, circuitos: list[str], start: str, end: str,
) -> tuple[float | None, int, int, str, list[str]]:
    """Ocupacion calendario agregada solo sobre el universo elegible: los
    CircuitoNegocio cuyo MetricStatus para ocupacion_calendario_pct no es
    NO_APLICA (prompt Sec.30 'solo sobre universo elegible'). Elegibilidad
    se deriva de la metrica canonica del motor (CoberturaCatalogo +
    CompletitudMaestro == COMPLETO), nunca de una lista fija hardcodeada."""
    per_circuito = engine.query(
        "ocupacion_calendario_pct", group_by=["CircuitoNegocio"],
        filters={"CircuitoNegocio": circuitos}, start_date=start, end_date=end,
    )
    eligible = per_circuito[per_circuito["MetricStatus"] != "NO_APLICA"]
    eligible_circuitos = sorted(eligible["CircuitoNegocio"].unique().tolist())
    if eligible.empty:
        return None, 0, 0, "NO_APLICA", eligible_circuitos
    numerator = int(eligible["Numerator"].sum())
    denominator = int(eligible["Denominator"].sum())
    status = "PARTIAL" if (eligible["MetricStatus"] == "PARTIAL").any() else "OK"
    value = (numerator / denominator * 100.0) if denominator else None
    return value, numerator, denominator, status, eligible_circuitos


def compute_kpi4_estatico(
    engine: MetricsEngine, universe: dict[str, Any],
    period: tuple[str, str], previous: tuple[str, str], ytd: tuple[str, str],
) -> dict[str, Any]:
    circuitos = universe["circuitos"]
    val_actual, num_actual, den_actual, status_actual, eligible = _eligible_static_occupancy(engine, circuitos, period[0], period[1])
    val_anterior, _, _, _, _ = _eligible_static_occupancy(engine, circuitos, previous[0], previous[1])
    val_ytd, _, _, _, _ = _eligible_static_occupancy(engine, circuitos, ytd[0], ytd[1])

    activos = engine.query(
        "elementos_con_actividad", filters={"CircuitoNegocio": eligible or circuitos, "Medio": "Estático"},
        start_date=period[0], end_date=period[1],
    )
    activos_val = int(activos["Value"].iloc[0]) if len(activos) else 0

    delta_pp = None
    if val_actual is not None and val_anterior is not None:
        delta_pp = _round1(val_actual - val_anterior)

    return {
        "activos": activos_val,
        "ocupacion_actual": _round1(val_actual),
        "ocupacion_anterior": _round1(val_anterior),
        "delta_pp": delta_pp,
        "ytd": _round1(val_ytd),
        "status": status_actual,
        "numerador": num_actual,
        "denominador": den_actual,
        "universo_elegible": eligible,
    }


# ---------------------------------------------------------------------------
# KPI 5 - Digital por calendario
# ---------------------------------------------------------------------------


def compute_kpi5_digital_calendario(
    engine: MetricsEngine, universe: dict[str, Any], period: tuple[str, str], previous: tuple[str, str],
) -> dict[str, Any]:
    digital_circuitos = universe["digital_circuitos"]
    registrados = engine.query("elementos_registrados", filters={"CircuitoNegocio": digital_circuitos, "Medio": "Digital"})
    reg_value = int(registrados["Value"].iloc[0]) if len(registrados) else 0

    actual = engine.query("elementos_con_actividad", filters={"CircuitoNegocio": digital_circuitos, "Medio": "Digital"}, start_date=period[0], end_date=period[1])
    anterior = engine.query("elementos_con_actividad", filters={"CircuitoNegocio": digital_circuitos, "Medio": "Digital"}, start_date=previous[0], end_date=previous[1])
    act_value = int(actual["Value"].iloc[0]) if len(actual) else 0
    ant_value = int(anterior["Value"].iloc[0]) if len(anterior) else 0

    pct_actual = _round1(act_value / reg_value * 100.0) if reg_value else None
    pct_anterior = _round1(ant_value / reg_value * 100.0) if reg_value else None
    delta_pp = _round1(pct_actual - pct_anterior) if pct_actual is not None and pct_anterior is not None else None

    return {
        "activos": act_value,
        "elegibles": reg_value,
        "pct_actual": pct_actual,
        "pct_anterior": pct_anterior,
        "delta_pp": delta_pp,
    }


# ---------------------------------------------------------------------------
# KPI 6 - Digital por fill rate (metrica canonica del motor, YPF excluido)
# ---------------------------------------------------------------------------


def compute_kpi6_digital_fill(
    engine: MetricsEngine, universe: dict[str, Any], period: tuple[str, str], previous: tuple[str, str],
) -> dict[str, Any]:
    digital_circuitos = universe["digital_circuitos"]
    actual = engine.query("fill_rate_slots", filters={"CircuitoNegocio": digital_circuitos}, start_date=period[0], end_date=period[1])
    anterior = engine.query("fill_rate_slots", filters={"CircuitoNegocio": digital_circuitos}, start_date=previous[0], end_date=previous[1])

    a = actual.iloc[0]
    p = anterior.iloc[0]
    pct_actual = _round1(a["Value"]) if pd.notna(a["Value"]) else None
    pct_anterior = _round1(p["Value"]) if pd.notna(p["Value"]) else None
    delta_pp = _round1(pct_actual - pct_anterior) if pct_actual is not None and pct_anterior is not None else None
    numerador = int(round(a["Numerator"])) if pd.notna(a["Numerator"]) else None

    return {
        "numerador": numerador,
        "denominador": int(a["Denominator"]) if pd.notna(a["Denominator"]) else None,
        "unidad": "slots ocupados",
        "pct_actual": pct_actual,
        "pct_anterior": pct_anterior,
        "delta_pp": delta_pp,
        "status": a["MetricStatus"],
    }


# ---------------------------------------------------------------------------
# Evolucion mensual (Digital / Estatico / YPF)
# ---------------------------------------------------------------------------


def compute_evolution(engine: MetricsEngine, universe: dict[str, Any], report_year: int, report_month: int) -> dict[str, Any]:
    """DIGITAL/ESTATICO: ElementoID Brand Plus activo (YPF excluido, sin
    cambios). YPF: estaciones (surrogate de estación YPF) activas por mes -- NUNCA
    ElementoID/formatos (correccion APIE Sec.6)."""
    digital_circuitos = universe["digital_circuitos"]
    ypf_ids = universe["ypf_element_ids"]
    ypf_map = universe["ypf_station_map"]
    meses, digital, estatico, ypf = [], [], [], []
    for m in range(1, report_month + 1):
        start, end = _period_bounds(report_year, m)
        d = engine.query("elementos_con_actividad", filters={"CircuitoNegocio": digital_circuitos, "Medio": "Digital"}, start_date=start, end_date=end)
        e = engine.query("elementos_con_actividad", filters={"CircuitoNegocio": digital_circuitos, "Medio": "Estático"}, start_date=start, end_date=end)
        ypf_estaciones = len(_ypf_active_stations(engine, ypf_ids, ypf_map, start, end))
        meses.append(MESES_ES_ABR[m - 1])
        digital.append(int(d["Value"].iloc[0]))
        estatico.append(int(e["Value"].iloc[0]))
        ypf.append(ypf_estaciones)
    return {"meses": meses, "digital": digital, "estatico": estatico, "ypf": ypf}


# ---------------------------------------------------------------------------
# Composicion del negocio (familias x Digital/Estatico)
# ---------------------------------------------------------------------------


def compute_catalogo_total(universe: dict[str, Any]) -> dict[str, Any]:
    """Catalogo comercial TV1 completo (estructural, no depende de
    periodo/actividad): TODOS los portfolio tiers (CORE + COMPLEMENTARIO),
    a diferencia de Core Comercial (KPI1) que es solo PortfolioTier=CORE.
    Necesario porque Cencomedia/MAB (COMPLEMENTARIO) deben poder aparecer
    en composicion sin forzarse dentro del Core (spec original Sec.20)."""
    maestro = universe["maestro"]
    non_ypf = maestro[maestro["CircuitoNegocio"] != "YPF"]
    non_ypf_value = int(non_ypf["ElementoID"].nunique())
    ypf_stations = universe["ypf_station_catalog_count"]

    complementario = non_ypf[non_ypf["PortfolioTier"] != "CORE"]
    detalle = complementario.groupby("CircuitoNegocio")["ElementoID"].nunique().sort_index()
    complementario_circuitos = sorted(detalle.index.tolist())
    complementario_count = int(detalle.sum())
    catalogo_value = non_ypf_value + ypf_stations
    core_value = catalogo_value - complementario_count

    nota_catalogo_ampliado = ""
    if complementario_count:
        partes = " + ".join(f"{_display_circuito(c)} {int(n)}" for c, n in detalle.items())
        nota_catalogo_ampliado = (
            f"Catálogo ampliado = Core Comercial {_fmt_es_int(core_value)} + {partes} "
            f"= {_fmt_es_int(catalogo_value)} unidades."
        )

    return {
        "value": catalogo_value,
        "non_ypf": non_ypf_value,
        "ypf_estaciones": ypf_stations,
        "complementario_circuitos": complementario_circuitos,
        "complementario_count": complementario_count,
        "nota_catalogo_ampliado": nota_catalogo_ampliado,
    }


def compute_composition(universe: dict[str, Any], grand_total: int) -> dict[str, Any]:
    """Composicion del CATALOGO comercial TV1 (estructural: cuenta lo que
    existe en el maestro, no lo que tuvo campana en el mes). Denominador =
    catalogo total TV1 (compute_catalogo_total, incluye COMPLEMENTARIO).
    Shoppings fusiona CENCOSUD+REMEROS (correccion Sec.7.1). No YPF:
    familias con split Digital/Estatico. YPF: UNA sola barra de estaciones
    (surrogate temporal) del CATALOGO, SIN split Digital/Estatico
    (correccion Sec.7.3): una estacion puede combinar formatos de ambos
    medios y dividirla falsearia su participacion."""
    maestro = universe["maestro"]
    non_ypf = maestro[maestro["CircuitoNegocio"] != "YPF"].copy()
    non_ypf["_familia"] = non_ypf["CircuitoNegocio"].map(lambda c: FAMILY_MAP.get(c, "Otros"))

    non_ypf_families = [name for name in FAMILY_ORDER if name != "YPF"]
    buckets: dict[str, dict[str, int]] = {name: {"Digital": 0, "Estático": 0} for name in non_ypf_families}
    otros_detalle: dict[str, int] = {}
    for (familia, medio), sub in non_ypf.groupby(["_familia", "Medio"]):
        buckets.setdefault(familia, {"Digital": 0, "Estático": 0})
        buckets[familia][medio] = buckets[familia].get(medio, 0) + int(sub["ElementoID"].nunique())
        if familia == "Otros":
            for circuito, csub in sub.groupby("CircuitoNegocio"):
                otros_detalle[circuito] = otros_detalle.get(circuito, 0) + int(csub["ElementoID"].nunique())

    familias = []
    for nombre in FAMILY_ORDER:
        if nombre == "YPF":
            ypf_total = universe["ypf_station_catalog_count"]
            familias.append({
                "nombre": "YPF",
                "split": False,
                "digital": 0,
                "estatico": 0,
                "total": ypf_total,
                "digital_pct": 0.0,
                "estatico_pct": 0.0,
                "pct_total": _round1(ypf_total / grand_total * 100.0) if grand_total else 0.0,
            })
            continue
        digital = buckets.get(nombre, {}).get("Digital", 0)
        estatico = buckets.get(nombre, {}).get("Estático", 0)
        total = digital + estatico
        familias.append({
            "nombre": nombre,
            "split": True,
            "digital": digital,
            "estatico": estatico,
            "total": total,
            "digital_pct": _round1(digital / grand_total * 100.0) if grand_total else 0.0,
            "estatico_pct": _round1(estatico / grand_total * 100.0) if grand_total else 0.0,
            "pct_total": _round1(total / grand_total * 100.0) if grand_total else 0.0,
        })

    reconciled_total = sum(f["total"] for f in familias)
    if reconciled_total != grand_total:
        raise BuildError(
            f"Composicion no reconcilia con catalogo comercial TV1: familias={reconciled_total} vs total={grand_total}"
        )

    nota_otros = ""
    if otros_detalle:
        partes = " + ".join(f"{_display_circuito(c)} {n}" for c, n in sorted(otros_detalle.items()))
        nota_otros = f"Otros = {partes}."

    return {
        "denominador": grand_total,
        "familias": familias,
        "otros_circuitos": sorted(otros_detalle.keys()),
        "nota_otros": nota_otros,
    }


# ---------------------------------------------------------------------------
# Insights (Lectura / Punto positivo / A atender) - seleccion programatica
# ---------------------------------------------------------------------------


def _signed_int(n: int) -> str:
    return ("+" if n >= 0 else "−") + _fmt_es_int(abs(n))


def compute_insights(
    kpi2: dict[str, Any], kpi3: dict[str, Any], kpi6: dict[str, Any],
    report_month_label: str, previous_month_label: str,
) -> dict[str, str]:
    """Lectura/Punto positivo/A atender: contenido factual, calculado a
    partir del desglose real de kpi3 (No YPF Digital/Estatico + estaciones
    YPF). Nunca inventa causas ni magnitudes."""
    lectura = (
        f"{report_month_label} registra <b>{_fmt_es_int(kpi3['value'])} unidades con campaña</b>: "
        f"{_fmt_es_int(kpi3['ypf_estaciones_activas'])} estaciones YPF, "
        f"{_fmt_es_int(kpi3['estatico_activo'])} unidades estáticas y "
        f"{_fmt_es_int(kpi3['digital_activo'])} digitales. "
        f"A la fecha se contabilizan <b>{_fmt_es_int(kpi2['ytd'])} campañas únicas</b>."
    )

    ypf_delta = kpi3["ypf_estaciones_activas"] - kpi3["ypf_estaciones_anterior"]
    digital_delta = kpi3["digital_activo"] - kpi3["digital_anterior"]
    estatico_delta = kpi3["estatico_activo"] - kpi3["estatico_anterior"]
    verbo = "crece" if kpi3["delta"] >= 0 else "cae"
    punto_positivo = (
        f"La actividad total {verbo} en <b>{_fmt_es_int(abs(kpi3['delta']))} unidades</b> respecto a "
        f"{previous_month_label.lower()}, impulsada por YPF ({_signed_int(ypf_delta)} estaciones); "
        f"Digital {_signed_int(digital_delta)} y Estático {_signed_int(estatico_delta)}."
    )

    if kpi6["delta_pp"] is not None and kpi6["delta_pp"] < 0:
        a_atender = (
            f"El fill rate digital cae <b>{_fmt_es_pct(abs(kpi6['delta_pp']))} pp</b> frente a "
            f"{previous_month_label.lower()} ({_fmt_es_pct(kpi6['pct_anterior'])}% → {_fmt_es_pct(kpi6['pct_actual'])}%)."
        )
    elif kpi6["delta_pp"] is not None and kpi6["delta_pp"] > 0:
        a_atender = (
            f"El fill rate digital mejora <b>{_fmt_es_pct(kpi6['delta_pp'])} pp</b> frente a "
            f"{previous_month_label.lower()}, aunque {_fmt_es_pct(kpi6['pct_actual'])}% de capacidad sigue sin vender."
        )
    else:
        a_atender = f"El fill rate digital se mantiene estable frente a {previous_month_label.lower()}."

    return {"lectura": lectura, "punto_positivo": punto_positivo, "a_atender": a_atender}


# ---------------------------------------------------------------------------
# Logo (reutilizado byte-a-byte desde la referencia, nunca retipeado a mano)
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


def build_tv1_data(path: str | Path = vi.DEFAULT_INPUT_PATH) -> dict[str, Any]:
    path = Path(path)
    sha_before = vi.calculate_sha256(path)

    _transform_result, semantic_result, engine = load_pipeline(path)
    universe = build_tv1_universe(semantic_result)

    period = _period_bounds(REPORT_YEAR, REPORT_MONTH)
    prev_year, prev_month = _previous_month(REPORT_YEAR, REPORT_MONTH)
    previous = _period_bounds(prev_year, prev_month)
    ytd = (f"{REPORT_YEAR}-01-01", period[1])

    kpi1 = compute_kpi1_core(universe)
    kpi2 = compute_kpi2_campanas(engine, universe, ytd[0], period, previous)
    kpi3 = compute_kpi3_actividad(engine, universe, period, previous, kpi1["value"])
    kpi_ypf = compute_kpi_ypf(engine, universe, period, previous, kpi3["value"])
    kpi4 = compute_kpi4_estatico(engine, universe, period, previous, ytd)
    kpi5 = compute_kpi5_digital_calendario(engine, universe, period, previous)
    kpi6 = compute_kpi6_digital_fill(engine, universe, period, previous)
    evolution = compute_evolution(engine, universe, REPORT_YEAR, REPORT_MONTH)
    catalogo = compute_catalogo_total(universe)
    composition = compute_composition(universe, catalogo["value"])
    insights = compute_insights(
        kpi2, kpi3, kpi6, MESES_ES[REPORT_MONTH - 1], MESES_ES[prev_month - 1],
    )

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
            "ytd_start": ytd[0],
            "fuente": "OCU26 · Base maestra + base campañas",
        },
        "kpis": {
            "core_comercial": kpi1,
            "campanas_unicas": kpi2,
            "unidades_actividad": kpi3,
            "ypf": kpi_ypf,
            "estatico": kpi4,
            "digital_calendario": kpi5,
            "digital_fill": kpi6,
        },
        "evolution": evolution,
        "catalogo_comercial": catalogo,
        "composition": composition,
        "insights": insights,
    }

    return {
        "data": data,
        "sha256": sha_after,
        "universe": {"circuitos": universe["circuitos"], "digital_circuitos": universe["digital_circuitos"]},
        "ypf_audit": {
            "estaciones_ypf_catalogo": universe["ypf_station_catalog_count"],
            "estaciones_ypf_activas": kpi3["ypf_estaciones_activas"],
            "stationkey_tv1_no_derivable": 0,  # build_tv1_universe levanta BuildError si hay alguno
            "grano": (
                "estacion YPF mediante surrogate temporal (prefijo ElementoID + localidad normalizada "
                "de Ubicacion); no existe columna APIE en la fuente"
            ),
        },
    }


def render_html(data: dict[str, Any]) -> str:
    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    logo_tag = extract_logo_img_tag(REFERENCE_PATH)
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    html = template.replace("{{LOGO_IMG_TAG}}", logo_tag)
    html = html.replace("{{TV1_DATA_JSON}}", payload)
    return html


def build_and_write(
    path: str | Path = vi.DEFAULT_INPUT_PATH,
    output_html: str | Path = DEFAULT_OUTPUT_HTML,
    output_json: str | Path = DEFAULT_OUTPUT_JSON,
) -> dict[str, Any]:
    result = build_tv1_data(path)
    html = render_html(result["data"])

    output_html = Path(output_html)
    output_html.write_text(html, encoding="utf-8")

    output_json = Path(output_json)
    output_json.parent.mkdir(parents=True, exist_ok=True)
    with open(output_json, "w", encoding="utf-8") as fh:
        json.dump(result["data"], fh, ensure_ascii=False, indent=2)

    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Construye tv1.html (dashboard TV1 OCU26) con datos reales.")
    parser.add_argument("--file", default=str(vi.DEFAULT_INPUT_PATH), help="Ruta al archivo .xlsx a leer")
    parser.add_argument("--output-html", default=str(DEFAULT_OUTPUT_HTML), help="Ruta del HTML productivo generado")
    parser.add_argument("--output-json", default=str(DEFAULT_OUTPUT_JSON), help="Ruta del snapshot JSON generado")
    args = parser.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    try:
        result = build_and_write(args.file, args.output_html, args.output_json)
    except BuildError as exc:
        print("TV1_BUILD_ERROR:", exc)
        return 1

    print("TV1_BUILD_OK")
    print(json.dumps(result["data"]["meta"], ensure_ascii=False, indent=2))
    print(json.dumps(result["data"]["kpis"], ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
