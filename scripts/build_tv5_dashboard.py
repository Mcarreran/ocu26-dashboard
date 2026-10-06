"""Capa de datos para el dashboard TV5 - OCU26 (Pipeline Comercial).

Migracion 2026-08-23 (pedido "TV5 Pipeline Comercial - Espacios"): TV5 pasa
a alojar el Pipeline Comercial que antes vivia en TV4 (ver historial git,
commit "checkpoint: migracion restaurada 2026-08-20", scripts/
build_tv4_dashboard.py de ese commit). TV4 paso a reportar exclusivamente
YPF ("Pulso comercial YPF", decision de negocio explicita del usuario, ver
docstring de build_tv4_dashboard.py) y este builder ocupa ahora el slot
TV5, con la MISMA logica de pipeline temporal (que corre / que viene) pero
con la unidad principal migrada de "activaciones" a "espacios comerciales":
campanas y activaciones quedan como metricas secundarias/de auditoria
(pedido Sec.5).

Se ejecuta DESPUES de scripts/semantic_model.py y scripts/metrics_engine.py
(Gate 3B). Reutiliza export_data.load_pipeline (Gate 4A), mismo patron que
build_tv1/tv2/tv3_dashboard.py: no reabre el Excel, no reimplementa reglas
de negocio de Gate 3. Importa constantes/funciones PUBLICAS (sin guion bajo)
de build_tv2_dashboard.py (universo/capacidad Digital) y build_tv3_dashboard.py
(universo Estatico): "reutilizar la logica canonica ya implementada para
TV2 y TV3" del pedido Sec.2/4, en vez de reimplementarla a mano. No importa
build_tv4_dashboard.py: ese modulo ya importa de build_tv5_dashboard.py
(load_geo_aux, para el mapa YPF) y un import en sentido inverso crearia un
ciclo de imports. Por eso el universo YPF (necesario solo para el
denominador de la Tarjeta 5, "Pipeline en Core") se resuelve aqui con una
version minima y autonoma (build_tv5_ypf_snapshot), documentada como
duplicado intencional -- mismo patron ya usado por TV3 para espejar la
lista de circuitos excluidos de TV1 (_TV1_ESTATICO_EXCLUDED_CIRCUITOS).

Universo Core Comercial (pedido Sec.3) = union de los universos canonicos
de TV2 (Digital: Pantallas LED + Shoppings Digital + AA2000 Digital) y TV3
(Estatico: Shoppings Estatico + AA2000 Estatico + Pilar Frontlight). YPF,
MAB, Cencomedia, APSA y London Supply quedan fuera del Core (no tienen
regla de espacio Core); YPF tiene su propia regla (TV4) y se usa solo para
el universo general comparativo de la Tarjeta 5. Los circuitos restantes
(APSA, LONDON_SUPPLY, MAB, CENCOMEDIA en el snapshot auditado) no tienen
regla de espacio confirmada: nunca se cuentan como cero, se listan aparte
como "pendientes" en calidad.

Definicion de espacio comercial (pedido Sec.4-5):
- Digital: cada campana activa en un ElementoID digital consume un slot de
  ese elemento -- IDCampaña x ElementoID, pares distintos (duplicados de un
  mismo par cuentan una sola vez). La capacidad confirmada por elemento
  (regla central semantic_model.capacidad_espacio_digital, la misma de
  TV1/TV2, Etapa 2A) se usa solo para detectar
  y declarar sobrecapacidad (mas campanas concurrentes que slots
  confirmados en un mismo elemento), nunca para capar el conteo real.
- Estatico: un espacio ocupado es un ElementoID fisico distinto con al
  menos una campana activa (igual TV3): dos campanas activas en el mismo
  elemento estatico cuentan 1 espacio ocupado (no 2).
- Eventos temporales (reservas futuras / inician 30d / finalizan 30d /
  finalizados historico): "asignaciones distintas de campana a espacio" =
  pares (IDCampaña, ElementoID) distintos, sea el elemento Digital o
  Estatico -- cada asignacion futura/pasada es por construccion un evento
  puntual sobre un elemento fisico especifico, no hay colapso de
  simultaneidad que aplicar (ese colapso solo corresponde a la foto de
  ocupacion AL CORTE, pedido Sec.5 "una activacion no equivale
  automaticamente a un espacio estatico distinto").

Corte operativo (pedido Sec.2): 31/07/2026, mismo "corte" que julio 2026
como mes de reporte vigente en TV1-3. Toda clasificacion temporal (activa/
futura/finalizada) se resuelve por FechaInicio/FechaFin contra el corte,
nunca por el campo Estado (igual razonamiento que el Pipeline historico:
Estado no es fecha-consistente en la fuente).

Uso:
    python scripts/build_tv5_dashboard.py
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import validate_input as vi  # noqa: E402
from semantic_model import filter_universe  # noqa: E402
import semantic_model as sm  # noqa: E402
from metrics_engine import MetricsEngine, ocupacion_simultanea_por_estacion  # noqa: E402
from export_data import load_pipeline  # noqa: E402
from build_tv2_dashboard import (  # noqa: E402
    PANTALLAS_CIRCUITOS as _TV2_PANTALLAS_CIRCUITOS,
    SHOPPINGS_CIRCUITOS as _TV2_SHOPPINGS_CIRCUITOS,
    AA2000_CIRCUITOS as _TV2_AA2000_CIRCUITOS,
    TV2_CIRCUITOS as CORE_DIGITAL_CIRCUITOS,
)
from build_tv3_dashboard import (  # noqa: E402
    OTROS_CIRCUITOS as _TV3_OTROS_CIRCUITOS,
    TV3_CIRCUITOS as CORE_ESTATICO_CIRCUITOS,
    MEDIO_ESTATICO,
)

REPO_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = Path(__file__).resolve().parent / "templates" / "tv5_template.html"
REFERENCE_PATH = REPO_ROOT / "audit_sources" / "TV5_REFERENCE.html"
DEFAULT_OUTPUT_HTML = REPO_ROOT / "tv5.html"
DEFAULT_OUTPUT_JSON = REPO_ROOT / "output" / "tv5_data.json"

# Cartografia real de Argentina (Natural Earth) y fuente geografica auxiliar
# YPF: ya NO las usa el propio TV5 (el Pipeline Comercial no tiene mapa),
# pero build_tv4_dashboard.py (Pulso comercial YPF) las importa de este
# modulo para su mapa de intensidad comercial (`from build_tv5_dashboard
# import load_geo_aux, ARG_PAIS_GEOJSON_PATH, ARG_PROVINCIAS_GEOJSON_PATH`).
# Se mantienen aqui SOLO por compatibilidad hacia atras con ese import: TV4
# debe seguir construyendo sin cambios (protocolo "TV4 YPF byte-identica").
ARG_PAIS_GEOJSON_PATH = Path(__file__).resolve().parent / "templates" / "assets" / "argentina_pais.geojson"
ARG_PROVINCIAS_GEOJSON_PATH = Path(__file__).resolve().parent / "templates" / "assets" / "argentina_provincias.geojson"
GEO_AUX_PATH = REPO_ROOT / "input_aux" / "YPF_GEO_COORDENADAS.xlsx.xlsx"
_GEO_PREFIX_RE = re.compile(r"^\s*(\d+)\s*-")

MESES_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]
MESES_ES_ABR = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

# Corte operativo TV5 (pedido Sec.2): cierre del mismo periodo de reporte
# vigente en TV1-3 (Julio 2026), no la fecha real del sistema.
REPORT_YEAR = 2026
REPORT_MONTH = 7
WINDOW_DAYS = 30

MEDIO_DIGITAL = "Digital"

TIMELINE_TOP_N = 6

# Grupos comerciales de apertura (pedido Sec.7): 4 grupos que particionan
# exactamente el Core Digital (TV2) + Core Estatico (TV3).
GRUPOS_CORE = ["Shoppings Digital", "Shoppings Estático", "Pantallas LED", "AA2000 / Pilar Frontlight"]

# Universo YPF minimo (solo para el denominador de la Tarjeta 5): mismo
# criterio de vigencia y misma tokenizacion de ElementoID que
# build_tv4_dashboard.py (duplicado intencional, ver docstring del modulo).
_YPF_DIGITAL_TOKENS = {"MB", "TT", "PPUNTER"}
_YPF_STATIC_TOKENS = {"FB"}
_YPF_ELEMENTO_TOKEN_RE = re.compile(r"^(.+) - ([A-Za-z]+) - (\d+)$")


class BuildError(Exception):
    """Error bloqueante al construir el dashboard TV5."""


def _period_bounds(year: int, month: int) -> tuple[pd.Timestamp, pd.Timestamp]:
    start = pd.Timestamp(year=year, month=month, day=1)
    end = start + pd.offsets.MonthEnd(0)
    return start, end


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


def _signed_int(v: int) -> str:
    return f"+{v}" if v > 0 else str(v)


# ---------------------------------------------------------------------------
# Universo Core (union TV2 Digital + TV3 Estatico) y grupos comerciales
# ---------------------------------------------------------------------------


def _grupo_comercial(circuito: str, medio: str) -> str:
    """Grupo comercial de apertura (pedido Sec.7): 4 grupos que particionan
    exacto el Core (Digital via TV2_CIRCUITOS, Estatico via TV3_CIRCUITOS)."""
    if circuito == "PANTALLAS_LED":
        return "Pantallas LED"
    if circuito in ("CENCOSUD", "REMEROS"):
        return "Shoppings Digital" if medio == MEDIO_DIGITAL else "Shoppings Estático"
    return "AA2000 / Pilar Frontlight"


def _espacio_capacidad_digital_tv5(row: pd.Series) -> float | None:
    """Capacidad confirmada de UN elemento digital Core, solo para detectar
    sobrecapacidad (pedido Sec.4). Delegado en la regla central
    (semantic_model.capacidad_espacio_digital, Etapa 2A), compartida con
    TV1/TV2: nunca una tabla propia de este builder."""
    capacidad, _categoria = sm.capacidad_espacio_digital(row)
    return capacidad


def build_tv5_universe(semantic_result: dict[str, Any]) -> dict[str, Any]:
    maestro = semantic_result["maestro"]
    config = semantic_result["config"]

    op = filter_universe(maestro, "OPERATIVO_GENERAL", config)
    digital_maestro = op[op["CircuitoNegocio"].isin(CORE_DIGITAL_CIRCUITOS) & (op["Medio"] == MEDIO_DIGITAL)].copy()
    static_maestro = op[op["CircuitoNegocio"].isin(CORE_ESTATICO_CIRCUITOS) & (op["Medio"] == MEDIO_ESTATICO)].copy()
    if digital_maestro.empty and static_maestro.empty:
        raise BuildError("Universo Core TV5 vacio: revisar business_semantics.json / TV2_CIRCUITOS / TV3_CIRCUITOS")

    digital_maestro["_grupo"] = [
        _grupo_comercial(c, MEDIO_DIGITAL) for c in digital_maestro["CircuitoNegocio"]
    ]
    static_maestro["_grupo"] = [
        _grupo_comercial(c, MEDIO_ESTATICO) for c in static_maestro["CircuitoNegocio"]
    ]

    digital_ids = digital_maestro["ElementoID"].tolist()
    static_ids = static_maestro["ElementoID"].tolist()
    grupo_map: dict[Any, str] = dict(zip(digital_maestro["ElementoID"], digital_maestro["_grupo"]))
    grupo_map.update(dict(zip(static_maestro["ElementoID"], static_maestro["_grupo"])))

    capacidad_map: dict[Any, float | None] = {
        row["ElementoID"]: _espacio_capacidad_digital_tv5(row)
        for _, row in digital_maestro.iterrows()
    }

    campanas = semantic_result["campanas"]
    op_campanas = filter_universe(campanas, "OPERATIVO_GENERAL", config)
    op_campanas = op_campanas[op_campanas["Estado"] != "Cancelado"]

    scope = op_campanas[op_campanas["ElementoID"].isin(set(digital_ids) | set(static_ids))].copy()

    return {
        "op": op,
        "op_campanas": op_campanas,
        "digital_ids": set(digital_ids),
        "static_ids": set(static_ids),
        "grupo_map": grupo_map,
        "capacidad_map": capacidad_map,
        "scope_campanas": scope,
        "circuitos_digital": sorted(digital_maestro["CircuitoNegocio"].unique().tolist()),
        "circuitos_estatico": sorted(static_maestro["CircuitoNegocio"].unique().tolist()),
        "elementos_digital": int(digital_maestro["ElementoID"].nunique()),
        "elementos_estatico": int(static_maestro["ElementoID"].nunique()),
    }


# ---------------------------------------------------------------------------
# Universo YPF minimo (solo denominador Tarjeta 5) -- duplicado intencional
# de build_tv4_dashboard.build_tv4_universe, ver docstring del modulo.
# ---------------------------------------------------------------------------


def _ypf_token(elemento_id: Any) -> str | None:
    m = _YPF_ELEMENTO_TOKEN_RE.match(str(elemento_id))
    return m.group(2).upper() if m else None


def build_tv5_ypf_snapshot(semantic_result: dict[str, Any], cutoff: pd.Timestamp) -> dict[str, Any]:
    maestro = semantic_result["maestro"]
    config = semantic_result["config"]

    op = filter_universe(maestro, "OPERATIVO_GENERAL", config)
    ypf = op[(op["CircuitoNegocio"] == "YPF") & op["RevisionMaestro"].notna()].copy()
    if ypf.empty:
        return {"digital_ids": set(), "static_ids": set(), "apie_map": {}, "espacios": 0, "campanas_ids": set()}

    ypf["_apie"] = ypf["Subcircuito"].astype(str).str.strip()
    ypf["_token"] = ypf["ElementoID"].apply(_ypf_token)
    digital_ids = set(ypf.loc[ypf["_token"].isin(_YPF_DIGITAL_TOKENS), "ElementoID"])
    static_ids = set(ypf.loc[ypf["_token"].isin(_YPF_STATIC_TOKENS), "ElementoID"])
    apie_map = dict(zip(ypf["ElementoID"], ypf["_apie"]))

    campanas = semantic_result["campanas"]
    op_campanas = filter_universe(campanas, "OPERATIVO_GENERAL", config)
    op_campanas = op_campanas[op_campanas["Estado"] != "Cancelado"]
    scope = op_campanas[op_campanas["ElementoID"].isin(digital_ids | static_ids)].copy()

    en_curso = scope[_en_curso_mask(scope, cutoff)]
    en_curso = en_curso[en_curso["IDCampaña"].notna() & (en_curso["IDCampaña"].astype(str).str.strip() != "")]

    # Etapa 2A: ocupacion YPF por estacion con la MISMA funcion central que
    # TV1/TV4 (maximo de campanas reales simultaneas, sin tope), evaluada en
    # el dia de corte: en un solo dia, simultaneas = campanas vigentes.
    dig = en_curso[en_curso["ElementoID"].isin(digital_ids)].copy()
    if not dig.empty:
        dig["_apie"] = dig["ElementoID"].map(apie_map)
        indefinida_ok = (dig["FechaIndefinida"] == "Si") & dig["FechaFin"].isna()
        dig["_fin_corte"] = dig["FechaFin"].where(~indefinida_ok, cutoff)
        por_estacion = ocupacion_simultanea_por_estacion(
            dig, cutoff, cutoff, sm.espacios_base_por_estacion("YPF", config),
            station_col="_apie", start_col="FechaInicio", end_col="_fin_corte",
        )
        n_dig = int(por_estacion["ocupacion"].sum()) if not por_estacion.empty else 0
    else:
        n_dig = 0

    stat = en_curso[en_curso["ElementoID"].isin(static_ids)]
    n_stat = int(stat["ElementoID"].nunique())

    return {
        "digital_ids": digital_ids,
        "static_ids": static_ids,
        "espacios": n_dig + n_stat,
        "campanas_ids": set(en_curso["IDCampaña"].dropna().unique()),
    }


# ---------------------------------------------------------------------------
# Clasificacion temporal (activa/futura/finalizada) resuelta por fecha
# contra el corte, nunca por Estado (pedido Sec.2).
# ---------------------------------------------------------------------------


def _en_curso_mask(df: pd.DataFrame, cutoff: pd.Timestamp) -> pd.Series:
    indefinida_ok = (df["FechaIndefinida"] == "Si") & df["FechaFin"].isna()
    eff_fin = df["FechaFin"].where(~indefinida_ok, pd.Timestamp("2100-01-01"))
    return df["FechaInicio"].notna() & (df["FechaInicio"] <= cutoff) & (eff_fin >= cutoff)


def compute_pipeline_windows(scope: pd.DataFrame, cutoff: pd.Timestamp) -> dict[str, Any]:
    window_end = cutoff + pd.Timedelta(days=WINDOW_DAYS)

    scope_validas = scope[scope["IDCampaña"].notna() & (scope["IDCampaña"].astype(str).str.strip() != "")]
    indefinida_ok_validas = (scope_validas["FechaIndefinida"] == "Si") & scope_validas["FechaFin"].isna()

    en_curso = scope_validas[_en_curso_mask(scope_validas, cutoff)]
    futuras = scope_validas[scope_validas["FechaInicio"].notna() & (scope_validas["FechaInicio"] > cutoff)]
    inician_30d = futuras[futuras["FechaInicio"] <= window_end]
    finalizan_30d = en_curso[
        en_curso["FechaFin"].notna() & (en_curso["FechaFin"] > cutoff)
        & (en_curso["FechaFin"] <= window_end) & (~indefinida_ok_validas.loc[en_curso.index])
    ]
    historico = scope_validas[
        scope_validas["FechaFin"].notna() & (scope_validas["FechaFin"] < cutoff) & (~indefinida_ok_validas)
    ]

    return {
        "window_end": window_end,
        "en_curso": en_curso,
        "futuras": futuras,
        "inician_30d": inician_30d,
        "finalizan_30d": finalizan_30d,
        "historico": historico,
    }


# ---------------------------------------------------------------------------
# Espacios ocupados (foto al corte, regla Digital/Estatico) y eventos
# temporales (pares campana-espacio, pedido Sec.5)
# ---------------------------------------------------------------------------


def _espacios_snapshot(en_curso: pd.DataFrame, digital_ids: set[Any], static_ids: set[Any]) -> dict[str, Any]:
    """Espacios ocupados AL CORTE (pedido Sec.4-5): Digital = pares
    (IDCampaña, ElementoID) distintos; Estatico = ElementoID distintos con
    >=1 campana activa (colapsa simultaneidad: 2 campanas en el mismo
    elemento estatico = 1 espacio ocupado)."""
    dig = en_curso[en_curso["ElementoID"].isin(digital_ids)]
    dig_espacios = dig.drop_duplicates(subset=["IDCampaña", "ElementoID"])
    n_dig = int(len(dig_espacios))

    stat = en_curso[en_curso["ElementoID"].isin(static_ids)]
    n_stat = int(stat["ElementoID"].nunique())

    campanas_unicas = int(en_curso["IDCampaña"].dropna().nunique())
    activaciones = int(en_curso.drop_duplicates(subset=["IDCampaña", "ElementoID"]).shape[0])

    return {
        "espacios": n_dig + n_stat,
        "espacios_digitales": n_dig,
        "espacios_estaticos": n_stat,
        "campanas_unicas": campanas_unicas,
        "activaciones": activaciones,
    }


def _espacios_evento(df: pd.DataFrame) -> dict[str, Any]:
    """Asignaciones distintas de campana a espacio (pedido Sec.5/6): pares
    (IDCampaña, ElementoID) distintos, sin distinguir Digital/Estatico (cada
    asignacion futura/pasada es un evento puntual sobre un elemento fisico
    especifico, no hay simultaneidad que colapsar)."""
    espacios_df = df.drop_duplicates(subset=["IDCampaña", "ElementoID"])
    return {
        "espacios": int(len(espacios_df)),
        "campanas_unicas": int(df["IDCampaña"].dropna().nunique()),
    }


# ---------------------------------------------------------------------------
# Tarjeta 5 - Pipeline en Core (espacios, no campanas)
# ---------------------------------------------------------------------------


def compute_pipeline_core(
    core_snapshot: dict[str, Any], en_curso_core: pd.DataFrame, ypf_snapshot: dict[str, Any],
) -> dict[str, Any]:
    espacios_core = core_snapshot["espacios"]
    espacios_general = espacios_core + ypf_snapshot["espacios"]
    pct = _round1(espacios_core / espacios_general * 100.0) if espacios_general else None

    campanas_core = set(en_curso_core["IDCampaña"].dropna().unique())
    campanas_general = campanas_core | ypf_snapshot["campanas_ids"]

    return {
        "pct": pct,
        "espacios_core": espacios_core,
        "espacios_general": espacios_general,
        "campanas_core": len(campanas_core),
        "campanas_general": len(campanas_general),
    }


# ---------------------------------------------------------------------------
# Apertura por grupo comercial (Sec.7): debe reconciliar exacto con el Core
# ---------------------------------------------------------------------------


def compute_distribucion_grupo(en_curso_core: pd.DataFrame, digital_ids: set[Any], static_ids: set[Any], grupo_map: dict[Any, str]) -> list[dict[str, Any]]:
    rows = []
    for grupo in GRUPOS_CORE:
        ids_grupo = {eid for eid, g in grupo_map.items() if g == grupo}
        sub = en_curso_core[en_curso_core["ElementoID"].isin(ids_grupo)]
        snap = _espacios_snapshot(sub, digital_ids, static_ids)
        rows.append({"grupo": grupo, "espacios": snap["espacios"]})
    return rows


# ---------------------------------------------------------------------------
# Panel izquierdo - Estado del pipeline (Sec.7)
# ---------------------------------------------------------------------------


def compute_estado_pipeline(core_snapshot: dict[str, Any], reservas: dict[str, Any], historico: dict[str, Any]) -> dict[str, Any]:
    return {
        "ocupados": core_snapshot["espacios"],
        "reservados": reservas["espacios"],
        "finalizados_historico": historico["espacios"],
    }


# ---------------------------------------------------------------------------
# Panel derecho - Proximos inicios (Sec.8)
# ---------------------------------------------------------------------------


def compute_proximos_inicios(inician_30d: pd.DataFrame, grupo_map: dict[Any, str]) -> tuple[list[dict[str, Any]], int]:
    if inician_30d.empty:
        return [], 0

    espacios_df = inician_30d.drop_duplicates(subset=["IDCampaña", "ElementoID"]).copy()
    espacios_df["_grupo"] = espacios_df["ElementoID"].map(grupo_map)

    grouped = (
        espacios_df.groupby(["IDCampaña", "FechaInicio"])
        .agg(campana=("Campaña", "first"), grupos=("_grupo", lambda s: sorted(set(s))), espacios=("ElementoID", "size"))
        .reset_index()
        .sort_values(["FechaInicio", "campana"])
    )
    total = int(len(grouped))
    rows = []
    for _, r in grouped.head(TIMELINE_TOP_N).iterrows():
        fecha: pd.Timestamp = r["FechaInicio"]
        campana = r["campana"]
        campana = campana if isinstance(campana, str) and campana.strip() else "Campaña sin nombre"
        rows.append({
            "fecha_iso": str(fecha.date()),
            "dia": fecha.strftime("%d"),
            "mes_abbr": MESES_ES_ABR[fecha.month - 1],
            "campana": campana,
            "grupo": " + ".join(r["grupos"]),
            "espacios": int(r["espacios"]),
            "badge": "Inicio",
        })
    return rows, total


# ---------------------------------------------------------------------------
# Calidad - sobrecapacidad digital y circuitos sin identidad de espacio
# ---------------------------------------------------------------------------


def compute_sobrecapacidad(en_curso_core: pd.DataFrame, digital_ids: set[Any], capacidad_map: dict[Any, float | None]) -> dict[str, Any]:
    dig = en_curso_core[en_curso_core["ElementoID"].isin(digital_ids)]
    if dig.empty:
        return {"elementos_sobrecapacidad": 0, "exceso_total": 0}
    concurrentes = dig.drop_duplicates(subset=["IDCampaña", "ElementoID"]).groupby("ElementoID").size()
    sobre = 0
    exceso = 0
    for elemento_id, n in concurrentes.items():
        cap = capacidad_map.get(elemento_id)
        if cap is not None and n > cap:
            sobre += 1
            exceso += int(n - cap)
    return {"elementos_sobrecapacidad": sobre, "exceso_total": exceso}


def compute_circuitos_pendientes(op: pd.DataFrame, core_circuitos: set[str]) -> list[dict[str, Any]]:
    """Circuitos del universo general sin regla de espacio confirmada
    (pedido Sec.3/6): ni Core (TV2+TV3) ni YPF. Nunca se cuentan como cero
    ni se les inventa capacidad: se listan con su cantidad de elementos de
    catalogo, calculada dinamicamente (nunca hardcodeada)."""
    todos = set(op["CircuitoNegocio"].dropna().unique())
    pendientes = sorted(todos - core_circuitos - {"YPF"})
    rows = []
    for c in pendientes:
        n = int(op.loc[op["CircuitoNegocio"] == c, "ElementoID"].nunique())
        rows.append({"circuito": c, "elementos_catalogo": n})
    return rows


# ---------------------------------------------------------------------------
# Insights (Lectura / Punto positivo / A atender) - pedido Sec.9
# ---------------------------------------------------------------------------


def compute_insights(
    core_snapshot: dict[str, Any], reservas: dict[str, Any], inician: dict[str, Any],
    finalizan: dict[str, Any], pipeline_core: dict[str, Any], distribucion: list[dict[str, Any]],
    circuitos_pendientes: list[dict[str, Any]],
) -> dict[str, str]:
    lectura = (
        f"Al 31/07 hay <b>{_fmt_es_int(core_snapshot['espacios'])} espacios comerciales ocupados</b> dentro del "
        f"Core, con <b>{_fmt_es_int(pipeline_core['campanas_core'])} campañas activas dentro del Core</b>, sobre "
        f"<b>{_fmt_es_int(pipeline_core['campanas_general'])} campañas activas del universo general</b> al corte. "
        f"El pipeline registra {_fmt_es_int(reservas['espacios'])} espacios reservados a futuro, con "
        f"{_fmt_es_int(inician['espacios'])} espacios que inician y {_fmt_es_int(finalizan['espacios'])} que "
        f"se liberan en los próximos 30 días."
    )

    # PUNTO POSITIVO (prioridad pedido Sec.9): 1) reservas futuras;
    # 2) crecimiento de espacios que ingresan; 3) concentracion favorable
    # en el Core; 4) fallback factual.
    if reservas["espacios"] > 0:
        punto_positivo = (
            f"Hay <b>{_fmt_es_int(reservas['espacios'])} espacios reservados</b> a futuro "
            f"({_fmt_es_int(reservas['campanas_unicas'])} campañas), que sumarán ocupación al pipeline."
        )
    elif inician["espacios"] > 0:
        punto_positivo = (
            f"Ingresan <b>{_fmt_es_int(inician['espacios'])} espacios nuevos</b> al Pipeline en los próximos "
            f"30 días ({_fmt_es_int(inician['campanas_unicas'])} campañas)."
        )
    elif pipeline_core["pct"] is not None and pipeline_core["pct"] >= 60:
        punto_positivo = (
            f"El Core concentra <b>{_fmt_es_pct(pipeline_core['pct'])}%</b> de los espacios del universo general "
            f"del Pipeline ({_fmt_es_int(pipeline_core['espacios_core'])} de {_fmt_es_int(pipeline_core['espacios_general'])})."
        )
    else:
        punto_positivo = f"El Core Comercial mantiene {_fmt_es_int(core_snapshot['espacios'])} espacios ocupados al corte."

    # A ATENDER (prioridad pedido Sec.9): 1) finalizan > inician;
    # 2) sobrecapacidad/circuitos sin identidad; 3) sin reservas; 4) fallback.
    if finalizan["espacios"] > 0 and finalizan["espacios"] > inician["espacios"]:
        a_atender = (
            f"Se liberan <b>{_fmt_es_int(finalizan['espacios'])} espacios</b> en los próximos 30 días frente a "
            f"{_fmt_es_int(inician['espacios'])} que inician: oportunidad de renovación."
        )
    elif circuitos_pendientes:
        nombres = ", ".join(c["circuito"] for c in circuitos_pendientes)
        a_atender = (
            f"{len(circuitos_pendientes)} circuito(s) del universo general ({nombres}) no tienen regla de "
            f"identidad de espacio confirmada: quedan fuera del denominador de la Tarjeta 5."
        )
    elif reservas["espacios"] == 0:
        a_atender = "No hay espacios reservados a futuro cargados: el pipeline depende únicamente de la ocupación ya en curso."
    else:
        a_atender = (
            f"El Core Comercial libera {_fmt_es_int(finalizan['espacios'])} espacios por finalización en los "
            f"próximos 30 días; conviene anticipar renovación."
        )

    return {"lectura": lectura, "punto_positivo": punto_positivo, "a_atender": a_atender}


# ---------------------------------------------------------------------------
# load_geo_aux -- SOLO compatibilidad hacia atras con build_tv4_dashboard.py
# (ver nota junto a GEO_AUX_PATH mas arriba). Logica sin cambios respecto a
# la version previa de este modulo cuando TV5 todavia era el mapa YPF.
# ---------------------------------------------------------------------------


def load_geo_aux(path: Path = GEO_AUX_PATH) -> tuple[dict[str, dict[str, Any]], dict[str, Any]]:
    """Lee la fuente geografica auxiliar YPF (aislada de Gate 1/2/3) y arma
    prefijo_numerico -> {lat, lon, ciudad, direccion}, quedandose con la
    primera fila valida por prefijo. Devuelve tambien el detalle de calidad
    del join. Usada exclusivamente por build_tv4_dashboard.py."""
    if not path.exists():
        raise BuildError(f"Fuente geografica auxiliar no encontrada: {path}")

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Hoja1"]
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    col_idx = {name: i for i, name in enumerate(headers)}
    for required in ("CODIGO", "LATITUD", "LONGITUD"):
        if required not in col_idx:
            raise BuildError(f"Columna requerida {required!r} no encontrada en {path} (Hoja1)")

    total_filas = 0
    con_prefijo = 0
    filas_validas = 0
    filas_invalidas = 0
    geo_map: dict[str, dict[str, Any]] = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        codigo = row[col_idx["CODIGO"]]
        if codigo is None:
            continue
        total_filas += 1
        m = _GEO_PREFIX_RE.match(str(codigo))
        if not m:
            continue
        con_prefijo += 1
        prefix = m.group(1)

        lat_raw = row[col_idx["LATITUD"]]
        lon_raw = row[col_idx["LONGITUD"]]
        try:
            lat, lon = float(lat_raw), float(lon_raw)
        except (TypeError, ValueError):
            filas_invalidas += 1
            continue
        if not (-90.0 <= lat <= 90.0 and -180.0 <= lon <= 180.0) or (lat == 0 and lon == 0):
            filas_invalidas += 1
            continue

        filas_validas += 1
        if prefix in geo_map:
            continue  # primera fila valida del prefijo gana (auditado: sin conflicto real)
        geo_map[prefix] = {
            "lat": lat,
            "lon": lon,
            "ciudad": row[col_idx.get("CIUDAD*", -1)] if "CIUDAD*" in col_idx else None,
            "direccion": row[col_idx.get("DIRECCION*", -1)] if "DIRECCION*" in col_idx else None,
        }

    calidad = {
        "fuente": str(path.relative_to(REPO_ROOT)),
        "filas_totales": total_filas,
        "filas_con_prefijo_numerico": con_prefijo,
        "filas_lat_lon_validas": filas_validas,
        "filas_lat_lon_invalidas_excluidas": filas_invalidas,
        "prefijos_distintos_con_coordenada_valida": len(geo_map),
    }
    return geo_map, calidad


# ---------------------------------------------------------------------------
# Logo (reutilizado byte-a-byte, igual patron que TV1-4/6)
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


def build_tv5_data(path: str | Path | None = None) -> dict[str, Any]:
    path = vi.resolve_input_path(path)
    sha_before = vi.calculate_sha256(path)

    _transform_result, semantic_result, _engine = load_pipeline(path)
    universe = build_tv5_universe(semantic_result)
    scope = universe["scope_campanas"]

    _, cutoff = _period_bounds(REPORT_YEAR, REPORT_MONTH)  # cierre del mes de reporte vigente (31/07/2026)

    windows = compute_pipeline_windows(scope, cutoff)
    en_curso_core = windows["en_curso"]

    core_snapshot = _espacios_snapshot(en_curso_core, universe["digital_ids"], universe["static_ids"])
    reservas = _espacios_evento(windows["futuras"])
    inician = _espacios_evento(windows["inician_30d"])
    finalizan = _espacios_evento(windows["finalizan_30d"])
    historico = _espacios_evento(windows["historico"])

    incomplete_rows = int(
        (scope["FechaInicio"].isna() | (scope["FechaFin"].isna() & ~((scope["FechaIndefinida"] == "Si") & scope["FechaFin"].isna()))).sum()
    )

    ypf_snapshot = build_tv5_ypf_snapshot(semantic_result, cutoff)
    pipeline_core = compute_pipeline_core(core_snapshot, en_curso_core, ypf_snapshot)

    distribucion = compute_distribucion_grupo(en_curso_core, universe["digital_ids"], universe["static_ids"], universe["grupo_map"])
    suma_grupos = sum(g["espacios"] for g in distribucion)
    if suma_grupos != core_snapshot["espacios"]:
        raise BuildError(f"Apertura por grupo comercial no reconcilia con el Core: {suma_grupos} vs {core_snapshot['espacios']}")

    estado_pipeline = compute_estado_pipeline(core_snapshot, reservas, historico)
    timeline_rows, timeline_total = compute_proximos_inicios(windows["inician_30d"], universe["grupo_map"])

    sobrecapacidad = compute_sobrecapacidad(en_curso_core, universe["digital_ids"], universe["capacidad_map"])
    core_circuitos = set(universe["circuitos_digital"]) | set(universe["circuitos_estatico"])
    circuitos_pendientes = compute_circuitos_pendientes(universe["op"], core_circuitos)

    insights = compute_insights(core_snapshot, reservas, inician, finalizan, pipeline_core, distribucion, circuitos_pendientes)

    sha_after = vi.calculate_sha256(path)
    if sha_after != sha_before:
        raise BuildError(
            f"ERROR CRITICO: el SHA-256 del input cambio durante la construccion del dashboard "
            f"(antes={sha_before}, despues={sha_after})."
        )

    data = {
        "meta": {
            "generado": dt.datetime.now().strftime("%d/%m/%Y %H:%M"),
            "cutoff_iso": str(cutoff.date()),
            "cutoff_label": cutoff.strftime("%d/%m/%Y"),
            "window_end_iso": str(windows["window_end"].date()),
            "window_end_label": windows["window_end"].strftime("%d/%m/%Y"),
            "window_days": WINDOW_DAYS,
            "fuente": "OCU26 · Base maestra + base campañas",
            "advertencia_fechas_incompletas": (
                f"{_fmt_es_int(incomplete_rows)} activaciones del Core Comercial quedan excluidas de la "
                f"clasificación temporal por fecha incompleta." if incomplete_rows else ""
            ),
        },
        "kpis": {
            "actividad_actual": core_snapshot,
            "reservas_futuras": reservas,
            "inician_30d": inician,
            "finalizan_30d": finalizan,
            "pipeline_core": pipeline_core,
        },
        "estado_pipeline": estado_pipeline,
        "distribucion_grupo": distribucion,
        "timeline": {"rows": timeline_rows, "total": timeline_total},
        "calidad": {
            "sobrecapacidad_digital": sobrecapacidad,
            "circuitos_pendientes": circuitos_pendientes,
        },
        "insights": insights,
    }

    return {
        "data": data,
        "sha256": sha_after,
        "universe": {
            "circuitos_digital": universe["circuitos_digital"],
            "circuitos_estatico": universe["circuitos_estatico"],
            "elementos_digital": universe["elementos_digital"],
            "elementos_estatico": universe["elementos_estatico"],
        },
    }


def render_html(data: dict[str, Any]) -> str:
    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    logo_tag = extract_logo_img_tag(REFERENCE_PATH)
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    html = template.replace("{{LOGO_IMG_TAG}}", logo_tag)
    html = html.replace("{{TV5_DATA_JSON}}", payload)
    return html


def build_and_write(
    path: str | Path | None = None,
    output_html: str | Path = DEFAULT_OUTPUT_HTML,
    output_json: str | Path = DEFAULT_OUTPUT_JSON,
) -> dict[str, Any]:
    result = build_tv5_data(path)
    html = render_html(result["data"])

    output_html = Path(output_html)
    output_html.write_text(html, encoding="utf-8")

    output_json = Path(output_json)
    output_json.parent.mkdir(parents=True, exist_ok=True)
    with open(output_json, "w", encoding="utf-8") as fh:
        json.dump(result["data"], fh, ensure_ascii=False, indent=2)

    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Construye tv5.html (dashboard TV5 OCU26, Pipeline Comercial) con datos reales.")
    parser.add_argument("--file", default=None, help="Ruta al .xlsx (si se omite: $OCU26_INPUT_PATH o input/OCU26_BASE_DATOS.xlsx)")
    parser.add_argument("--output-html", default=str(DEFAULT_OUTPUT_HTML), help="Ruta del HTML productivo generado")
    parser.add_argument("--output-json", default=str(DEFAULT_OUTPUT_JSON), help="Ruta del snapshot JSON generado")
    args = parser.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    try:
        result = build_and_write(args.file, args.output_html, args.output_json)
    except BuildError as exc:
        print("TV5_BUILD_ERROR:", exc)
        return 1

    print("TV5_BUILD_OK")
    print(json.dumps(result["data"]["meta"], ensure_ascii=False, indent=2))
    print(json.dumps(result["data"]["kpis"], ensure_ascii=False, indent=2))
    print("UNIVERSE:", json.dumps(result["universe"], ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
