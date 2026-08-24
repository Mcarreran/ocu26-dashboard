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
from zoneinfo import ZoneInfo

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

# Zona horaria oficial de TV1 (prompt Sec.11 / CM3 Sec.7): toda metadata de
# generacion es timezone-aware en ART. No se suma/resta 3hs a mano.
ART_TZ = ZoneInfo("America/Argentina/Buenos_Aires")
ART_TZ_NAME = "America/Argentina/Buenos_Aires"

# Periodo de referencia de TV1. Definido UNA sola vez aqui (prompt Sec.18):
# cambiar estos dos numeros re-genera todo el dashboard para otro mes/anio.
REPORT_YEAR = 2026
REPORT_MONTH = 7

# Exclusion absoluta de negocio (prompt Sec.17): APSA y London Supply nunca
# entran a TV1. No es una regla de Gate3B (alli London Supply SI cuenta para
# IncluyeConteoGeneral) sino una decision de scope propia de este dashboard.
# CENCOMEDIA se agrego el 21/08/2026 (decision explicita posterior del
# usuario, ver docs/CM3_MARCO_RECTOR_ESPACIOS_PUBLICITARIOS_2026-08-19.md):
# excluida de TODOS los analisis de TV1 (catalogos, espacios, campanas,
# historicos, rankings, composiciones, insights). Sus filas NO se borran del
# Excel ni de Gates: solo se excluyen de este universo de reporte, igual que
# APSA/London. CENCOSUD permanece incluido (no confundir por nombre).
TV1_EXCLUDED_CIRCUITOS = {"APSA", "LONDON_SUPPLY", "CENCOMEDIA"}

FAMILY_MAP = {
    "CENCOSUD": "Shoppings",
    "REMEROS": "Shoppings",
    "PANTALLAS_LED": "Pantallas",
    "YPF": "YPF",
    "AA2000": "AA2000",
}
# CENCOMEDIA no tiene entrada aqui a proposito (decision 21/08/2026): esta
# excluida de TV1_EXCLUDED_CIRCUITOS, por lo que nunca llega una fila con
# CircuitoNegocio=CENCOMEDIA a este mapeo; no debe aparecer ni siquiera como
# familia en cero (prompt Sec.1: "Cencomedia no aparece en payload ni HTML").
FAMILY_ORDER = ["Shoppings", "Pantallas", "YPF", "AA2000", "Otros"]

# ---------------------------------------------------------------------------
# Espacios publicitarios (Marco Rector CM3 2026-08-19): capa nueva sobre el
# grano fisico existente (ElementoID/estacion). No reemplaza los KPIs de
# arriba (elementos/estaciones activos): esos se preservan intactos y se
# reusan como equivalencia fisica secundaria (MR.8 punto 2). Conversion
# NUEVA, deliberadamente distinta del perfil legacy de
# config/business_semantics.json ("digital_capacity.slots_profiles", Gate3:
# TOTEM=20, PUENTE_LED=13): por instruccion explicita del usuario esas cifras
# de Gate3 NO se reutilizan aqui. Vive solo en este builder, scope TV1.
ESPACIOS_POR_FORMATO_DIGITAL: dict[str, int] = {
    "PANTALLA_LED": 20,
    "TOTEM": 10,
    "PUENTE_LED": 10,
}
ESPACIOS_YPF_POR_ESTACION = 5

# Soportes físicos YPF (Tarjeta 1, correccion 21/08/2026): mapeo
# FormatoNegocio -> etiqueta de reporte, mismos 4 formatos ya resueltos por
# semantic_model._resolve_formato_negocio via
# config/business_semantics.json formato_negocio.ypf_elemento_id_token_map
# (MB/TT/PPUNTER/FB). Capa puramente fisica (conteo de unidades instaladas),
# nunca reemplaza la capacidad comercial de ESPACIOS_YPF_POR_ESTACION.
YPF_SOPORTE_FORMATO_LABELS: dict[str, str] = {
    "YPF_MENU_BOARD": "menu_board",
    "YPF_TORRE": "torres",
    "YPF_PUNTERA": "punteras",
    "YPF_MUPI_FOTOBOX": "fotobox",
}


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
# Soportes físicos (Tarjeta 1, correccion 21/08/2026): conteo de UNIDADES
# FISICAS instaladas, capa DISTINTA de "espacios comerciales" (compute_espacios
# mas abajo). Estructural (no depende de periodo/actividad), no toca ni
# recalcula Espacios Totales (3.958 desde la correccion 22/08/2026, antes 3.898).
# ---------------------------------------------------------------------------


def compute_soportes_fisicos(universe: dict[str, Any]) -> dict[str, Any]:
    """Soportes fisicos != espacios comerciales (prompt Sec.4 "IMPORTANTE").
    No YPF: 1 ElementoID fisico = 1 soporte, Estatico y Digital contados por
    separado (todo el universo TV1, incluidos los formatos digitales sin
    tasa de espacios confirmada: siguen siendo un soporte fisico instalado
    aunque no aporten espacios comerciales). YPF: se cuenta cada Puntera/
    Menu Board/Torre/Fotobox INDIVIDUALMENTE via FormatoNegocio -- nunca por
    estacion -- porque varios soportes digitales en una misma estacion NO
    aumentan su capacidad comercial de 5 espacios (ESPACIOS_YPF_POR_ESTACION)."""
    maestro = universe["maestro"]
    non_ypf = maestro[maestro["CircuitoNegocio"] != "YPF"]
    ypf = maestro[maestro["CircuitoNegocio"] == "YPF"]

    estatico_no_ypf = int(non_ypf.loc[non_ypf["Medio"] == "Estático", "ElementoID"].nunique())
    digital_no_ypf = int(non_ypf.loc[non_ypf["Medio"] == "Digital", "ElementoID"].nunique())

    ypf_por_formato = ypf.groupby("FormatoNegocio")["ElementoID"].nunique()
    detalle_ypf = {
        etiqueta: int(ypf_por_formato.get(formato, 0))
        for formato, etiqueta in YPF_SOPORTE_FORMATO_LABELS.items()
    }
    soportes_ypf = sum(detalle_ypf.values())

    total = estatico_no_ypf + digital_no_ypf + soportes_ypf
    if total == 0:
        raise BuildError("Soportes fisicos totales = 0: revisar universo TV1")

    pct_estatico = _round1(estatico_no_ypf / total * 100.0)
    pct_digital = _round1(digital_no_ypf / total * 100.0)
    pct_ypf = _round1(soportes_ypf / total * 100.0)
    suma_pct = round((pct_estatico or 0.0) + (pct_digital or 0.0) + (pct_ypf or 0.0), 1)
    if abs(suma_pct - 100.0) > 0.1:
        raise BuildError(
            f"Porcentajes de soportes fisicos no suman 100% (tolerancia +/-0.1pp): "
            f"estatico={pct_estatico} digital={pct_digital} ypf={pct_ypf} suma={suma_pct}"
        )

    return {
        "total": total,
        "estatico_no_ypf": estatico_no_ypf,
        "digital_no_ypf": digital_no_ypf,
        "ypf": soportes_ypf,
        "ypf_detalle": detalle_ypf,
        "pct_estatico": pct_estatico,
        "pct_digital": pct_digital,
        "pct_ypf": pct_ypf,
    }


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
# Espacios publicitarios (decision usuario 21/08/2026: catalogo completo,
# YA NO "nucleo exacto"). Cencomedia excluida via TV1_EXCLUDED_CIRCUITOS
# (build_tv1_universe), CENCOSUD permanece. YPF ahora suma Digital+Estatico
# con ocupacion propia (estacion activa x5 SI cuenta como ocupacion en TV1,
# reemplaza la restriccion anterior por decision explicita de este prompt).
# ---------------------------------------------------------------------------


def _espacio_capacidad_digital(row: pd.Series) -> tuple[float | None, str]:
    """Capacidad en espacios de un elemento Digital no-YPF, y su categoria
    para reporting. Nunca reutiliza config/business_semantics.json
    (Gate3/slots_profiles): tasas nuevas definidas en este builder.

    - PANTALLA_LED: 20 (regla confirmada).
    - TOTEM en CENCOSUD (Totem de Shopping): 10 (regla confirmada).
    - TOTEM (formato OTRO, Descripcion "TV Led") en REMEROS: 10 (correccion
      22/08/2026, regla de negocio confirmada por el usuario: los 6
      elementos digitales de Remeros Shoppings son totems comerciales de 10
      slots cada uno -> 60 espacios. semantic_model los resuelve como
      FormatoNegocio=OTRO porque su Descripcion "TV Led" no matchea ningun
      keyword rule de business_semantics.json; no se confunde con el
      elemento Remeros de Pantallas LED, que es CircuitoNegocio
      PANTALLAS_LED y conserva su propia regla PANTALLA_LED=20).
    - TOTEM en AA2000 ("Tripstore", confirmado con el usuario 21/08/2026):
      capacidad REGISTRADA (CapacidadSlotsReel), NUNCA la tasa de 10 de
      Shopping ni un valor inventado.
    - PUENTE_LED: 10 (regla confirmada).
    - TRIEDRO: capacidad registrada (CapacidadSlotsReel); no hay tasa
      generica confirmada por el usuario, se usa el valor fuente tal cual.
    - Cualquier otro formato ('Otro' fuera de Remeros, etc.): sin regla ->
      (None, 'SIN_REGLA'), nunca se inventa una cifra."""
    fmt = row["FormatoNegocio"]
    circuito = row["CircuitoNegocio"]
    if fmt == "PANTALLA_LED":
        return float(ESPACIOS_POR_FORMATO_DIGITAL["PANTALLA_LED"]), "PANTALLA_LED"
    if fmt == "PUENTE_LED":
        return float(ESPACIOS_POR_FORMATO_DIGITAL["PUENTE_LED"]), "PUENTE_LED"
    if fmt == "TOTEM" and circuito == "CENCOSUD":
        return float(ESPACIOS_POR_FORMATO_DIGITAL["TOTEM"]), "TOTEM_SHOPPING"
    if fmt == "OTRO" and circuito == "REMEROS":
        return float(ESPACIOS_POR_FORMATO_DIGITAL["TOTEM"]), "TOTEM_SHOPPING"
    if fmt == "TOTEM" and circuito == "AA2000":
        legacy = row["CapacidadSlotsReel"]
        return (float(legacy) if pd.notna(legacy) and legacy > 0 else None), "TRIPSTORE_AA2000"
    if fmt == "TRIEDRO":
        legacy = row["CapacidadSlotsReel"]
        return (float(legacy) if pd.notna(legacy) and legacy > 0 else None), "TRIEDRO"
    return None, "SIN_REGLA"


def _con_capacidad_digital(digital: pd.DataFrame) -> pd.DataFrame:
    """Agrega columnas _capacidad/_categoria via _espacio_capacidad_digital.
    pandas.DataFrame.apply(..., result_type='expand') sobre un DataFrame
    vacio devuelve el frame original sin las columnas nuevas (caso limite
    real, no documentado): se maneja aparte para no romper con fixtures
    pequeños/sin elementos digitales."""
    digital = digital.copy()
    if digital.empty:
        digital["_capacidad"] = pd.Series(dtype="float64")
        digital["_categoria"] = pd.Series(dtype="object")
        return digital
    capacidades = digital.apply(_espacio_capacidad_digital, axis=1, result_type="expand")
    # pandas infiere dtype=object en la columna 0 cuando hay una mezcla de
    # float y None: forzar float64 evita que .sum() sobre un subconjunto
    # vacio/filtrado devuelva '' (comportamiento real de pandas en columnas
    # object) en lugar de 0.0.
    digital["_capacidad"] = pd.to_numeric(capacidades[0], errors="coerce")
    digital["_categoria"] = capacidades[1]
    return digital


_ESPACIOS_DIGITAL_NUCLEO_CATEGORIAS = ["PANTALLA_LED", "TOTEM_SHOPPING", "PUENTE_LED", "TRIEDRO"]
_ESPACIOS_DIGITAL_TODAS_CATEGORIAS = _ESPACIOS_DIGITAL_NUCLEO_CATEGORIAS + ["TRIPSTORE_AA2000"]


def _max_overlap_count(intervals: list[tuple[Any, Any]]) -> int:
    """Maxima cantidad de intervalos [inicio,fin] (dias, ambos inclusive)
    simultaneamente activos en algun punto. Barrido de eventos (+1 en el
    inicio, -1 el dia siguiente al fin, porque el fin es inclusive): dos
    campañas que comparten aunque sea 1 dia cuentan como concurrentes; dos
    campañas consecutivas sin dia compartido no."""
    if not intervals:
        return 0
    events: list[tuple[Any, int]] = []
    one_day = pd.Timedelta(days=1)
    for s, e in intervals:
        events.append((s, 1))
        events.append((e + one_day, -1))
    events.sort(key=lambda ev: (ev[0], ev[1]))
    cur = 0
    best = 0
    for _t, delta in events:
        cur += delta
        best = max(best, cur)
    return best


def _estacion_ocupacion_ypf(
    intervals: list[tuple[Any, Any]], capacidad_slots: int = ESPACIOS_YPF_POR_ESTACION,
) -> dict[str, Any]:
    """Ocupacion de UNA estacion YPF digital (prompt "AGREGADO DE DISEÑO
    FUTURO — SOBRECARGA YPF", 21/08/2026): capacidad_slots = capacidad
    confirmada de la estacion si existiera una fuente (hoy no existe en la
    base -> siempre cae al default de 5, nunca se infla el catalogo segun
    la demanda); campanas_concurrentes = maximo de IDCampaña distintas
    activas SIMULTANEAMENTE (solapamiento real de fechas, no "activas en
    algun momento del periodo"); slots_ocupados = min(capacidad,
    concurrentes) (nunca > capacidad); campanas_excedentes = concurrentes
    por encima de la capacidad, se conservan/contabilizan como actividad
    pero NUNCA se convierten en mas espacios de catalogo/ocupados."""
    concurrentes = _max_overlap_count(intervals)
    slots_ocupados = min(capacidad_slots, concurrentes)
    excedentes = max(0, concurrentes - capacidad_slots)
    pct = _round1(slots_ocupados / capacidad_slots * 100.0) if capacidad_slots else None
    presion = _round1(concurrentes / capacidad_slots * 100.0) if capacidad_slots else None
    if concurrentes > capacidad_slots:
        estado = "SOBRECAPACIDAD"
    elif concurrentes == capacidad_slots and capacidad_slots > 0:
        estado = "COMPLETA"
    else:
        estado = "DISPONIBLE"
    return {
        "capacidad_slots": capacidad_slots,
        "campanas_concurrentes": concurrentes,
        "slots_ocupados": slots_ocupados,
        "campanas_excedentes": excedentes,
        "porcentaje_ocupacion": pct,
        "indice_presion": presion,
        "estado": estado,
    }


def _ypf_digital_ocupados_por_estacion(
    engine: MetricsEngine, ypf_dig: pd.DataFrame, start: str, end: str,
) -> dict[str, Any]:
    """Ocupacion YPF Digital, estacion por estacion, agregada a nivel red
    (corrección prioritaria + agregado de diseño futuro, 21/08/2026;
    reemplaza 'estacion activa x5'). Por estacion: capacidad 5 (default,
    sin fuente de capacidad confirmada por estacion todavia);
    campanas_concurrentes = maximo solapamiento REAL de fechas (no distinct
    en el periodo); slots_ocupados = min(5, concurrentes), nunca invade el
    catalogo. Deduplicacion por StationKey+IDCampaña: una misma campaña en
    varios ElementoID/filas de la misma estacion se une a UN solo intervalo
    [min(inicio), max(fin)] antes de barrer solapamientos (nunca cuenta
    filas/activaciones/ElementoID como slots separados). Elementos con
    fechas no resolubles (sin FechaInicio, o sin FechaFin/FechaIndefinida)
    ya quedan fuera de `_campanas_overlap`: se reportan aparte como
    'estaciones_requiere_confirmacion' (agregado, nunca se inventa un
    resultado para ellas)."""
    element_ids = ypf_dig["ElementoID"].tolist()
    vacio = {
        "ocupados": 0, "estaciones_activas": 0, "estaciones_completa": 0,
        "estaciones_sobrecapacidad": 0, "campanas_excedentes_total": 0,
        "estaciones_requiere_confirmacion": 0,
    }
    if not element_ids:
        return vacio

    incompletos = engine._incomplete_date_elements(element_ids)
    requiere_confirmacion = 0
    if incompletos:
        station_map_all = dict(zip(ypf_dig["ElementoID"], ypf_dig["StationKey"]))
        requiere_confirmacion = len({station_map_all[e] for e in incompletos if e in station_map_all})

    overlap = engine._campanas_overlap(element_ids, start, end)
    overlap = overlap[overlap["IDCampaña"].notna() & (overlap["IDCampaña"].astype(str).str.strip() != "")]
    if overlap.empty:
        return {**vacio, "estaciones_requiere_confirmacion": requiere_confirmacion}

    station_map = dict(zip(ypf_dig["ElementoID"], ypf_dig["StationKey"]))
    overlap = overlap.copy()
    overlap["StationKey"] = overlap["ElementoID"].map(station_map)
    overlap = overlap.dropna(subset=["StationKey"])
    if overlap.empty:
        return {**vacio, "estaciones_requiere_confirmacion": requiere_confirmacion}

    # Union por (estacion, campaña): si la misma campaña aparece en varios
    # ElementoID/filas de la estacion, se funde en UN intervalo antes de
    # barrer solapamientos (nunca cuenta cada fila por separado).
    union = overlap.groupby(["StationKey", "IDCampaña"]).agg(
        _eff_start=("_eff_start", "min"), _eff_end=("_eff_end", "max"),
    ).reset_index()

    ocupados_total = 0
    estaciones_activas = 0
    estaciones_completa = 0
    estaciones_sobrecapacidad = 0
    excedentes_total = 0
    for _station, g in union.groupby("StationKey"):
        intervals = list(zip(g["_eff_start"], g["_eff_end"]))
        r = _estacion_ocupacion_ypf(intervals, ESPACIOS_YPF_POR_ESTACION)
        ocupados_total += r["slots_ocupados"]
        excedentes_total += r["campanas_excedentes"]
        if r["campanas_concurrentes"] > 0:
            estaciones_activas += 1
        if r["estado"] == "COMPLETA":
            estaciones_completa += 1
        elif r["estado"] == "SOBRECAPACIDAD":
            estaciones_sobrecapacidad += 1

    return {
        "ocupados": int(ocupados_total),
        "estaciones_activas": estaciones_activas,
        "estaciones_completa": estaciones_completa,
        "estaciones_sobrecapacidad": estaciones_sobrecapacidad,
        "campanas_excedentes_total": int(excedentes_total),
        "estaciones_requiere_confirmacion": requiere_confirmacion,
    }


def _espacios_snapshot(engine: MetricsEngine, universe: dict[str, Any], start: str, end: str) -> dict[str, Any]:
    """Un corte (periodo dado) de todo el catalogo de espacios TV1: Estatico
    no YPF completo, Digital no YPF por categoria de capacidad, y YPF
    separado en Digital/Estatico (prompt Sec.5, Sec.11)."""
    maestro = universe["maestro"]
    non_ypf = maestro[maestro["CircuitoNegocio"] != "YPF"]
    ypf = maestro[maestro["CircuitoNegocio"] == "YPF"]

    estatico = non_ypf[non_ypf["Medio"] == "Estático"]
    est_ids = estatico["ElementoID"].tolist()
    est_totales = len(est_ids)
    est_ocupados = int(engine._campanas_overlap(est_ids, start, end)["ElementoID"].dropna().nunique())

    digital = _con_capacidad_digital(non_ypf[non_ypf["Medio"] == "Digital"])

    categorias: dict[str, dict[str, Any]] = {}
    for cat in _ESPACIOS_DIGITAL_TODAS_CATEGORIAS:
        sub = digital[(digital["_categoria"] == cat) & digital["_capacidad"].notna()]
        ids_cat = sub["ElementoID"].tolist()
        ocup_cat = 0.0
        if ids_cat:
            slots_s, _seg, _inc, _sal = engine._digital_period_activity(ids_cat, start, end)
            ocup_cat = float(slots_s.sum()) if len(slots_s) else 0.0
        categorias[cat] = {
            "elementos": len(ids_cat),
            "totales": int(sub["_capacidad"].sum()),
            "ocupados": ocup_cat,
        }

    sin_regla = digital[digital["_capacidad"].isna()]

    ypf_dig = ypf[ypf["Medio"] == "Digital"].copy()
    ypf_est = ypf[ypf["Medio"] == "Estático"]
    ypf_dig["StationKey"] = [_ypf_station_key(e, u) for e, u in zip(ypf_dig["ElementoID"], ypf_dig["Ubicacion"])]
    ypf_dig_stations_catalogo = int(ypf_dig["StationKey"].nunique())

    ypf_dig_ocup = _ypf_digital_ocupados_por_estacion(engine, ypf_dig, start, end)

    ypf_est_totales = len(ypf_est)
    ypf_est_ocupados = int(
        engine._campanas_overlap(ypf_est["ElementoID"].tolist(), start, end)["ElementoID"].dropna().nunique()
    )

    return {
        "estatico": {"totales": est_totales, "ocupados": est_ocupados},
        "digital": categorias,
        "digital_sin_regla_elementos": int(sin_regla["ElementoID"].nunique()),
        "ypf_digital": {
            "estaciones_catalogo": ypf_dig_stations_catalogo,
            "estaciones_activas": ypf_dig_ocup["estaciones_activas"],
            "estaciones_completa": ypf_dig_ocup["estaciones_completa"],
            "estaciones_sobrecapacidad": ypf_dig_ocup["estaciones_sobrecapacidad"],
            "campanas_excedentes_total": ypf_dig_ocup["campanas_excedentes_total"],
            "estaciones_requiere_confirmacion": ypf_dig_ocup["estaciones_requiere_confirmacion"],
            "totales": ypf_dig_stations_catalogo * ESPACIOS_YPF_POR_ESTACION,
            "ocupados": ypf_dig_ocup["ocupados"],
        },
        "ypf_estatico": {"totales": ypf_est_totales, "ocupados": ypf_est_ocupados},
    }


def _sumar_categorias(snap: dict[str, Any], categorias: list[str]) -> tuple[int, float]:
    tot = sum(snap["digital"][c]["totales"] for c in categorias)
    ocup = sum(snap["digital"][c]["ocupados"] for c in categorias)
    return tot, ocup


def _totales_ocupados_globales(snap: dict[str, Any]) -> tuple[int, float]:
    dig_tot, dig_ocup = _sumar_categorias(snap, _ESPACIOS_DIGITAL_TODAS_CATEGORIAS)
    tot = snap["estatico"]["totales"] + dig_tot + snap["ypf_digital"]["totales"] + snap["ypf_estatico"]["totales"]
    ocup = snap["estatico"]["ocupados"] + dig_ocup + snap["ypf_digital"]["ocupados"] + snap["ypf_estatico"]["ocupados"]
    return tot, ocup


def _pct_par(ocup: float, tot: int) -> float | None:
    return _round1(ocup / tot * 100.0) if tot else None


def compute_espacios(
    engine: MetricsEngine, universe: dict[str, Any], period: tuple[str, str], previous: tuple[str, str],
) -> dict[str, Any]:
    """Espacios publicitarios de TV1 completos (tarjetas 1-2-4-5-6, prompt
    Sec.4-11): catalogo cargado y autorizado, SIN restringirse a un "nucleo
    exacto" (decision 21/08/2026 reemplaza el diseño anterior). Excluye
    unicamente Cencomedia/APSA/London (ya fuera de `universe["maestro"]`
    via TV1_EXCLUDED_CIRCUITOS) y los formatos digitales sin conversion de
    espacios confirmada (nunca se inventa una tasa)."""
    actual = _espacios_snapshot(engine, universe, period[0], period[1])
    anterior = _espacios_snapshot(engine, universe, previous[0], previous[1])

    total_actual, ocup_actual = _totales_ocupados_globales(actual)
    total_anterior, ocup_anterior = _totales_ocupados_globales(anterior)
    ocup_actual_i = int(round(ocup_actual))
    ocup_anterior_i = int(round(ocup_anterior))
    pct = _pct_par(ocup_actual_i, total_actual)
    pct_anterior = _pct_par(ocup_anterior_i, total_anterior)
    delta_pp = _round1(pct - pct_anterior) if pct is not None and pct_anterior is not None else None

    # Auditoria 21/08/2026 (correccion controlada): el denominador de la
    # tarjeta Digital DEBE incluir Tripstore AA2000 (200 espacios), porque ya
    # esta incluido en Espacios Totales y en la barra AA2000 de la apertura
    # por circuito -- excluirlo solo de esta tarjeta rompia la reconciliacion
    # (435+730+2533=3.698 != 3.898). Se usa _ESPACIOS_DIGITAL_TODAS_CATEGORIAS
    # (incluye TRIPSTORE_AA2000), nunca _NUCLEO_CATEGORIAS, para que Digital
    # reconcilie exacto con Espacios Totales.
    dig_tot, dig_ocup = _sumar_categorias(actual, _ESPACIOS_DIGITAL_TODAS_CATEGORIAS)
    dig_tot_a, dig_ocup_a = _sumar_categorias(anterior, _ESPACIOS_DIGITAL_TODAS_CATEGORIAS)
    dig_ocup_i, dig_ocup_a_i = int(round(dig_ocup)), int(round(dig_ocup_a))
    dig_pct, dig_pct_a = _pct_par(dig_ocup_i, dig_tot), _pct_par(dig_ocup_a_i, dig_tot_a)
    dig_delta = _round1(dig_pct - dig_pct_a) if dig_pct is not None and dig_pct_a is not None else None

    est, est_a = actual["estatico"], anterior["estatico"]
    est_pct, est_pct_a = _pct_par(est["ocupados"], est["totales"]), _pct_par(est_a["ocupados"], est_a["totales"])
    est_delta = _round1(est_pct - est_pct_a) if est_pct is not None and est_pct_a is not None else None

    ypf_tot = actual["ypf_digital"]["totales"] + actual["ypf_estatico"]["totales"]
    ypf_ocup = actual["ypf_digital"]["ocupados"] + actual["ypf_estatico"]["ocupados"]
    ypf_tot_a = anterior["ypf_digital"]["totales"] + anterior["ypf_estatico"]["totales"]
    ypf_ocup_a = anterior["ypf_digital"]["ocupados"] + anterior["ypf_estatico"]["ocupados"]
    ypf_pct, ypf_pct_a = _pct_par(ypf_ocup, ypf_tot), _pct_par(ypf_ocup_a, ypf_tot_a)
    ypf_delta = _round1(ypf_pct - ypf_pct_a) if ypf_pct is not None and ypf_pct_a is not None else None
    ypf_ocup_i = int(round(ypf_ocup))

    # Participacion de cada familia sobre el total de espacios OCUPADOS
    # (prompt Sec. Objetivo 1.B): numerador y denominador propios, nunca un
    # "resto hasta 100%". Las tres familias explican el 100% de ocup_actual_i
    # porque Tripstore (ya sumado dentro de "digital") es la unica pieza
    # digital fuera de Estatico/Digital-nucleo/YPF, y su ocupado es 0.
    part_est = _pct_par(est["ocupados"], ocup_actual_i)
    part_dig = _pct_par(dig_ocup_i, ocup_actual_i)
    part_ypf = _pct_par(ypf_ocup_i, ocup_actual_i)

    tripstore, tripstore_a = actual["digital"]["TRIPSTORE_AA2000"], anterior["digital"]["TRIPSTORE_AA2000"]
    soportes_fisicos = (
        est["totales"]
        + sum(actual["digital"][c]["elementos"] for c in _ESPACIOS_DIGITAL_TODAS_CATEGORIAS)
        + actual["ypf_digital"]["estaciones_catalogo"]
        + actual["ypf_estatico"]["totales"]
    )

    return {
        "totales": total_actual,
        "ocupados": ocup_actual_i,
        "disponibles": total_actual - ocup_actual_i,
        "pct_ocupacion": pct,
        "ocupados_anterior": ocup_anterior_i,
        "pct_ocupacion_anterior": pct_anterior,
        "delta_pp": delta_pp,
        "status": "PARTIAL",
        "status_motivo": (
            "Incluye AA2000 (CompletitudMaestro=PARCIAL, catálogo con Mendoza/Córdoba sin cargar) y "
            "MAB (CoberturaCatalogo=DESCONOCIDO): catálogo cargado y autorizado, no se afirma que sea "
            "el 100% del universo físico real."
        ),
        "soportes_fisicos_equivalentes": soportes_fisicos,
        "estatico": {
            "totales": est["totales"], "ocupados": est["ocupados"], "ocupados_anterior": est_a["ocupados"],
            "pct_ocupacion": est_pct, "pct_ocupacion_anterior": est_pct_a, "delta_pp": est_delta,
            "participacion_ocupado_pct": part_est,
        },
        "digital": {
            "totales": dig_tot, "ocupados": dig_ocup_i, "ocupados_anterior": dig_ocup_a_i,
            "pct_ocupacion": dig_pct, "pct_ocupacion_anterior": dig_pct_a, "delta_pp": dig_delta,
            "participacion_ocupado_pct": part_dig,
            "por_categoria": {
                "PANTALLA_LED": {**actual["digital"]["PANTALLA_LED"], "capacidad_por_unidad": ESPACIOS_POR_FORMATO_DIGITAL["PANTALLA_LED"]},
                "TOTEM_SHOPPING": {**actual["digital"]["TOTEM_SHOPPING"], "capacidad_por_unidad": ESPACIOS_POR_FORMATO_DIGITAL["TOTEM"]},
                "PUENTE_LED": {**actual["digital"]["PUENTE_LED"], "capacidad_por_unidad": ESPACIOS_POR_FORMATO_DIGITAL["PUENTE_LED"]},
                "TRIEDRO": {**actual["digital"]["TRIEDRO"], "capacidad_por_unidad": "registrada (CapacidadSlotsReel)"},
                "TRIPSTORE_AA2000": {**actual["digital"]["TRIPSTORE_AA2000"], "capacidad_por_unidad": "registrada (CapacidadSlotsReel)"},
            },
        },
        "ypf": {
            "totales": ypf_tot, "ocupados": ypf_ocup_i, "ocupados_anterior": int(round(ypf_ocup_a)),
            "pct_ocupacion": ypf_pct, "pct_ocupacion_anterior": ypf_pct_a, "delta_pp": ypf_delta,
            "participacion_ocupado_pct": part_ypf,
            "digital": actual["ypf_digital"], "estatico": actual["ypf_estatico"],
            "espacios_por_estacion": ESPACIOS_YPF_POR_ESTACION,
            "nota": (
                "Corrección 21/08/2026: la ocupación digital YPF se calcula estación por estación como "
                "mín(5, máximo de IDCampaña distintas realmente CONCURRENTES —solapamiento de fechas, no "
                "solo activas alguna vez en el período— en esa estación) — NUNCA estaciones activas × 5, "
                "que sobreestimaba la ocupación. Estaciones con más de 5 campañas concurrentes quedan en "
                "5 ocupados y se registran como sobrecapacidad (excedentes ≠ espacios ocupados). YPF "
                "Estático (fotobox) no registra ninguna campaña en la base a la fecha (0 ocupados)."
            ),
        },
        "aa2000_tripstore": {
            "elementos": tripstore["elementos"], "totales": tripstore["totales"],
            "ocupados": int(round(tripstore["ocupados"])), "ocupados_anterior": int(round(tripstore_a["ocupados"])),
            "capacidad_fuente": "CapacidadSlotsReel (capacidad registrada; NO la tasa de 10 de Tótem de Shopping)",
            "nota": "Pertenece a AA2000 (barra 'AA2000' en la apertura por circuito); no se duplica en Shoppings.",
        },
        "sin_regla_confirmada": {
            "elementos": actual["digital_sin_regla_elementos"],
            "motivo": (
                "Formatos digitales sin conversión de espacios confirmada (fuera de Pantalla LED / Tótem de "
                "Shopping / Puente LED / Tríedro / Tripstore AA2000): no se incluyen en Espacios Totales ni en "
                "ninguna barra, para no inventar una capacidad."
            ),
        },
    }


def compute_campanas_catalogo(
    engine: MetricsEngine, universe: dict[str, Any], ytd_start: str, period: tuple[str, str], previous: tuple[str, str],
) -> dict[str, Any]:
    """Tarjeta 3 - Campañas únicas (correccion 21/08/2026): incluye YPF
    ademas del resto del universo autorizado (universe["element_ids"], que
    ya excluye Cencomedia/APSA/London). Campañas únicas = distinct
    IDCampaña sobre el universo TV1 COMPLETO: una IDCampaña presente en YPF
    y tambien en otro circuito se cuenta una sola vez en el total (nunca
    sumando subtotales por circuito). Presencias en elementos = distinct
    (IDCampaña, ElementoID) — nunca "activaciones"."""
    element_ids = universe["element_ids"]

    def _scope(start: str, end: str) -> tuple[int, int]:
        ov = engine._campanas_overlap(element_ids, start, end)
        ov = ov[ov["IDCampaña"].notna() & (ov["IDCampaña"].astype(str).str.strip() != "")]
        campanas = int(ov["IDCampaña"].nunique())
        presencias = int(ov.drop_duplicates(subset=["IDCampaña", "ElementoID"]).shape[0])
        return campanas, presencias

    campanas_actual, presencias_actual = _scope(period[0], period[1])
    campanas_anterior, _presencias_anterior = _scope(previous[0], previous[1])
    campanas_ytd, _presencias_ytd = _scope(ytd_start, period[1])

    return {
        "campanas_unicas": campanas_actual,
        "campanas_unicas_anterior": campanas_anterior,
        "delta_campanas": campanas_actual - campanas_anterior,
        "campanas_unicas_ytd": campanas_ytd,
        "presencias_en_elementos": presencias_actual,
        "nota": (
            "Presencias en elementos: cantidad de combinaciones únicas campaña–elemento. Una campaña cuenta "
            "más de una vez cuando está presente en distintos elementos."
        ),
    }


def compute_evolution_espacios(
    engine: MetricsEngine, universe: dict[str, Any], report_year: int, report_month: int,
) -> dict[str, Any]:
    """Evolución mensual de ESPACIOS OCUPADOS (prompt Sec.13), no de
    elementos/estaciones. Estático/Digital usan el mismo universo que las
    tarjetas 4/5 (Digital incluye Tripstore AA2000, igual que la tarjeta
    desde la auditoría del denominador 21/08/2026). YPF = ocupación estación
    por estación (mín(5, IDCampaña distintas) por estación, corrección
    prioritaria 21/08/2026 — NUNCA estaciones activas × 5) + elementos YPF
    estáticos con campaña del mes × 1, en una sola serie."""
    maestro = universe["maestro"]
    non_ypf = maestro[maestro["CircuitoNegocio"] != "YPF"]
    estatico_ids = non_ypf.loc[non_ypf["Medio"] == "Estático", "ElementoID"].tolist()

    digital = _con_capacidad_digital(non_ypf[non_ypf["Medio"] == "Digital"])
    digital_ids = digital.loc[
        digital["_categoria"].isin(_ESPACIOS_DIGITAL_TODAS_CATEGORIAS), "ElementoID"
    ].tolist()

    ypf = maestro[maestro["CircuitoNegocio"] == "YPF"]
    ypf_dig = ypf[ypf["Medio"] == "Digital"].copy()
    ypf_dig["StationKey"] = [_ypf_station_key(e, u) for e, u in zip(ypf_dig["ElementoID"], ypf_dig["Ubicacion"])]
    ypf_est_ids = ypf.loc[ypf["Medio"] == "Estático", "ElementoID"].tolist()

    meses, estatico_serie, digital_serie, ypf_serie = [], [], [], []
    for m in range(1, report_month + 1):
        start, end = _period_bounds(report_year, m)

        est_ov = engine._campanas_overlap(estatico_ids, start, end)
        estatico_serie.append(int(est_ov["ElementoID"].dropna().nunique()))

        if digital_ids:
            slots_s, _seg, _inc, _sal = engine._digital_period_activity(digital_ids, start, end)
            digital_serie.append(int(round(float(slots_s.sum()))) if len(slots_s) else 0)
        else:
            digital_serie.append(0)

        ypf_dig_ocup = _ypf_digital_ocupados_por_estacion(engine, ypf_dig, start, end)
        ov_est = engine._campanas_overlap(ypf_est_ids, start, end)
        n_est_activos = int(ov_est["ElementoID"].dropna().nunique())
        ypf_serie.append(ypf_dig_ocup["ocupados"] + n_est_activos)

        meses.append(MESES_ES_ABR[m - 1])

    return {"meses": meses, "estatico": estatico_serie, "digital": digital_serie, "ypf": ypf_serie}


def compute_apertura_circuito(universe: dict[str, Any], espacios: dict[str, Any]) -> dict[str, Any]:
    """Apertura de espacios por circuito (prompt Sec.14): Shoppings /
    Pantallas LED / AA2000 / YPF / Otros. Denominador = Espacios Totales
    (tarjeta 1) exacto: las barras deben reconciliar 1:1 (BuildError si no).
    Tótems/Puentes/Tríedros de Shopping van en Shoppings (desde 22/08/2026
    incluye los 6 totems Remeros, categoria TOTEM_SHOPPING igual que
    Cencosud); Tripstore va en AA2000 (no se duplica); YPF Estático va en
    YPF, no en Otros; Pilar y MAB van en Otros."""
    maestro = universe["maestro"]
    non_ypf = maestro[maestro["CircuitoNegocio"] != "YPF"]
    digital_all = _con_capacidad_digital(non_ypf[non_ypf["Medio"] == "Digital"])

    def _digital_espacios(circuitos: list[str], categorias: list[str]) -> int:
        sub = digital_all[
            digital_all["CircuitoNegocio"].isin(circuitos) & digital_all["_categoria"].isin(categorias)
        ]
        return int(sub["_capacidad"].sum())

    def _estatico_espacios(circuitos: list[str]) -> int:
        sub = non_ypf[non_ypf["CircuitoNegocio"].isin(circuitos) & (non_ypf["Medio"] == "Estático")]
        return len(sub)

    shoppings_est = _estatico_espacios(["CENCOSUD", "REMEROS"])
    # Correccion 22/08/2026: REMEROS se suma junto con CENCOSUD (antes solo
    # CENCOSUD) porque sus 6 totems ya tienen capacidad confirmada
    # (TOTEM_SHOPPING); si no se incluyera aca, la barra "Shoppings" quedaria
    # 60 espacios por debajo de Espacios Totales y compute_apertura_circuito
    # fallaria con BuildError de reconciliacion.
    shoppings_dig = _digital_espacios(["CENCOSUD", "REMEROS"], ["TOTEM_SHOPPING", "PUENTE_LED", "TRIEDRO"])
    pantallas_dig = _digital_espacios(["PANTALLAS_LED"], ["PANTALLA_LED"])
    aa2000_est = _estatico_espacios(["AA2000"])
    aa2000_tripstore = _digital_espacios(["AA2000"], ["TRIPSTORE_AA2000"])
    otros_est = _estatico_espacios(["PILAR_FRONTLIGHT", "MAB"])
    ypf_est = espacios["ypf"]["estatico"]["totales"]
    ypf_dig = espacios["ypf"]["digital"]["totales"]

    barras = [
        {"nombre": "Shoppings", "estatico": shoppings_est, "digital": shoppings_dig},
        {"nombre": "Pantallas LED", "estatico": 0, "digital": pantallas_dig},
        {"nombre": "AA2000", "estatico": aa2000_est, "digital": aa2000_tripstore},
        {"nombre": "YPF", "estatico": ypf_est, "digital": ypf_dig},
        {"nombre": "Otros", "estatico": otros_est, "digital": 0},
    ]
    for b in barras:
        b["total"] = b["estatico"] + b["digital"]

    grand_total = espacios["totales"]
    reconciled = sum(b["total"] for b in barras)
    if reconciled != grand_total:
        raise BuildError(
            f"Apertura por circuito no reconcilia con Espacios Totales: barras={reconciled} vs total={grand_total}"
        )

    for b in barras:
        b["pct_total"] = _round1(b["total"] / grand_total * 100.0) if grand_total else 0.0
        b["digital_pct"] = _round1(b["digital"] / grand_total * 100.0) if grand_total else 0.0
        b["estatico_pct"] = _round1(b["estatico"] / grand_total * 100.0) if grand_total else 0.0

    return {
        "denominador": grand_total,
        "barras": barras,
        "nota": (
            "Shoppings incluye Tótems, Puentes y Triedros; AA2000 incluye Tripstore; YPF incluye "
            "Estático; Otros incluye Pilar y MAB."
        ),
    }


def compute_insights_espacios(
    espacios: dict[str, Any], campanas: dict[str, Any],
    report_month_label: str, previous_month_label: str,
) -> dict[str, str]:
    """Lectura/Punto positivo/A atender del nuevo catálogo de espacios
    (prompt Sec.17): nunca menciona Cencomedia, "fuera del núcleo" ni el
    total 1.052 del diseño anterior; no afirma más de lo que el dato permite."""
    lectura = (
        f"{report_month_label} registra <b>{_fmt_es_int(espacios['ocupados'])} espacios publicitarios ocupados</b> "
        f"sobre {_fmt_es_int(espacios['totales'])} espacios de capacidad cargada y autorizada "
        f"({_fmt_es_int(espacios['soportes_fisicos_equivalentes'])} soportes/estaciones equivalentes). "
        f"<b>{_fmt_es_int(campanas['campanas_unicas'])} campañas únicas</b> (incluye YPF) y "
        f"{_fmt_es_int(campanas['presencias_en_elementos'])} presencias en elementos."
    )

    if espacios["delta_pp"] is not None:
        verbo = "crece" if espacios["delta_pp"] >= 0 else "cae"
        punto_positivo = (
            f"La ocupación de espacios {verbo} <b>{_fmt_es_pct(abs(espacios['delta_pp']))} pp</b> frente a "
            f"{previous_month_label.lower()} "
            f"({_fmt_es_pct(espacios['pct_ocupacion_anterior'])}% → {_fmt_es_pct(espacios['pct_ocupacion'])}%)."
        )
    else:
        punto_positivo = f"Ocupación de espacios sin comparación disponible frente a {previous_month_label.lower()}."

    pct_disp = _round1(100.0 - espacios["pct_ocupacion"]) if espacios["pct_ocupacion"] is not None else None
    a_atender = (
        f"{_fmt_es_int(espacios['disponibles'])} espacios disponibles "
        f"({_fmt_es_pct(pct_disp)}% de la capacidad). "
        f"{_fmt_es_int(espacios['sin_regla_confirmada']['elementos'])} soportes digitales quedan fuera del "
        f"cálculo por no tener una conversión de espacios confirmada."
    )
    return {"lectura": lectura, "punto_positivo": punto_positivo, "a_atender": a_atender}


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

    soportes_fisicos = compute_soportes_fisicos(universe)

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

    espacios = compute_espacios(engine, universe, period, previous)
    campanas_catalogo = compute_campanas_catalogo(engine, universe, ytd[0], period, previous)
    evolution_espacios = compute_evolution_espacios(engine, universe, REPORT_YEAR, REPORT_MONTH)
    apertura_circuito = compute_apertura_circuito(universe, espacios)
    insights = compute_insights_espacios(
        espacios, campanas_catalogo, MESES_ES[REPORT_MONTH - 1], MESES_ES[prev_month - 1],
    )

    sha_after = vi.calculate_sha256(path)
    if sha_after != sha_before:
        raise BuildError(
            f"ERROR CRITICO: el SHA-256 del input cambio durante la construccion del dashboard "
            f"(antes={sha_before}, despues={sha_after})."
        )

    generado_art = dt.datetime.now(ART_TZ)
    data = {
        "meta": {
            "generado": generado_art.strftime("%d/%m/%Y %H:%M"),
            "generado_iso": generado_art.isoformat(timespec="seconds"),
            "timezone": ART_TZ_NAME,
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
        "espacios": espacios,
        "soportes_fisicos": soportes_fisicos,
        "campanas_catalogo": campanas_catalogo,
        "evolution_espacios": evolution_espacios,
        "apertura_circuito": apertura_circuito,
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
