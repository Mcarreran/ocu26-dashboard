"""Capa de datos para el dashboard TV3 - OCU26 (Core Comercial Estatico).

Rediseno completo 2026-08-22 (prompt maestro "TV3 - Core Comercial
Estatico", ver docs/CONTEXTO_MAESTRO TVS NUEVAS - HASTA LA TV3.md Sec.23):
reemplaza el diseno anterior (vocabulario "elementos activos/elegibles",
5 tarjetas). TV3 mide ELEMENTOS FISICOS distintos via ElementoID; nunca
slots ni CapacidadSlotsReel (eso es exclusivo de TV1/TV2 Digital). No usa
vocabulario "Espacios" (decision explicita del prompt maestro: TV3 no
copia la arquitectura de TV1).

Se ejecuta DESPUES de scripts/semantic_model.py y scripts/metrics_engine.py
(Gate 3B). Reutiliza export_data.load_pipeline (Gate 4A), mismo patron que
build_tv1_dashboard.py / build_tv2_dashboard.py: no reabre el Excel, no
reimplementa reglas de negocio de Gate 3 (toda cifra sale de
MetricsEngine.query() o de metodos internos ya aprobados como
MetricsEngine._campanas_overlap). No importa build_tv1_dashboard.py ni
build_tv2_dashboard.py a proposito (decision de arquitectura preexistente):
cada TV es un builder independiente sobre el mismo pipeline compartido. La
reconciliacion de lectura con TV1 (Sec.7 del prompt) replica localmente,
de forma minima y documentada, la misma definicion de universo Estatico de
TV1 (build_tv1_dashboard.TV1_EXCLUDED_CIRCUITOS) en vez de importarla.

Universo TV3 = CORE COMERCIAL ESTATICO = Shoppings Estatico (Cencosud +
Remeros) + AA2000 Estatico + Pilar Frontlight ("Otros"). YPF/APSA/London
Supply/Cencomedia/MAB excluidos (MAB es PortfolioTier COMPLEMENTARIO, no
CORE; ver reconciliacion_tv1 mas abajo, que documenta esta diferencia
frente al universo Estatico mas amplio de TV1).

Uso:
    python scripts/build_tv3_dashboard.py
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
TEMPLATE_PATH = Path(__file__).resolve().parent / "templates" / "tv3_template.html"
REFERENCE_PATH = REPO_ROOT / "audit_sources" / "TV1_REFERENCE.html.html"
DEFAULT_OUTPUT_HTML = REPO_ROOT / "tv3.html"
DEFAULT_OUTPUT_JSON = REPO_ROOT / "output" / "tv3_data.json"

MESES_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]
MESES_ES_ABR = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

# Zona horaria oficial (prompt maestro Sec.9 / Sec.16): toda metadata de
# generacion es timezone-aware en ART, igual que TV1.
ART_TZ = ZoneInfo("America/Argentina/Buenos_Aires")
ART_TZ_NAME = "America/Argentina/Buenos_Aires"

# Periodo de referencia TV3 (prompt maestro): Julio 2026 vs Junio 2026.
REPORT_YEAR = 2026
REPORT_MONTH = 7

# Universo TV3 (prompt maestro Sec.3): Core Comercial Estatico = Shoppings
# Estatico (Cencosud+Remeros) + AA2000 Estatico + Otros/Pilar Frontlight,
# siempre scope Medio=Estatico. YPF/APSA/London/Cencomedia/MAB nunca entran.
SHOPPINGS_CIRCUITOS = ["CENCOSUD", "REMEROS"]
AA2000_CIRCUITOS = ["AA2000"]
OTROS_CIRCUITOS = ["PILAR_FRONTLIGHT"]
TV3_CIRCUITOS = SHOPPINGS_CIRCUITOS + AA2000_CIRCUITOS + OTROS_CIRCUITOS
MEDIO_ESTATICO = "Estático"

# Universo Estatico de TV1 (mirror de solo lectura, prompt maestro Sec.7:
# "auditoria de lectura", nunca importar build_tv1_dashboard.py). Debe
# coincidir EXACTO con build_tv1_dashboard.TV1_EXCLUDED_CIRCUITOS: si TV1
# cambia esta lista, este mirror queda desactualizado a proposito (no hay
# import cruzado) y test_reconciliacion_tv1_diff_is_mab_only lo detectaria.
_TV1_ESTATICO_EXCLUDED_CIRCUITOS = {"APSA", "LONDON_SUPPLY", "CENCOMEDIA"}

SOPORTES_TOP_N = 3


class BuildError(Exception):
    """Error bloqueante al construir el dashboard TV3."""


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


# ---------------------------------------------------------------------------
# Universo TV3
# ---------------------------------------------------------------------------


def build_tv3_universe(semantic_result: dict[str, Any]) -> dict[str, Any]:
    maestro = semantic_result["maestro"]
    config = semantic_result["config"]

    op = filter_universe(maestro, "OPERATIVO_GENERAL", config)
    tv3_maestro = op[op["CircuitoNegocio"].isin(TV3_CIRCUITOS) & (op["Medio"] == MEDIO_ESTATICO)].copy()
    if tv3_maestro.empty:
        raise BuildError("Universo TV3 vacio: revisar business_semantics.json / Medio=Estático")

    return {
        "maestro": tv3_maestro,
        "op": op,
        "circuitos": sorted(tv3_maestro["CircuitoNegocio"].unique().tolist()),
        "element_ids": tv3_maestro["ElementoID"].tolist(),
    }


# ---------------------------------------------------------------------------
# Helpers de conteo (ElementoID distinto; nunca slots ni CapacidadSlotsReel)
# ---------------------------------------------------------------------------


def _catalogo(engine: MetricsEngine, circuitos: list[str]) -> int:
    r = engine.query("elementos_registrados", filters={"CircuitoNegocio": circuitos, "Medio": MEDIO_ESTATICO})
    return int(r["Value"].iloc[0]) if len(r) else 0


def _ocupados(engine: MetricsEngine, circuitos: list[str], start: str, end: str) -> int:
    """ElementoID distintos con campana vigente en [start,end] (regla temporal
    canonica de MetricsEngine._campanas_overlap). Un elemento con varias
    campanas en el mismo mes se cuenta una sola vez (prompt maestro Sec.3)."""
    r = engine.query(
        "elementos_con_actividad", filters={"CircuitoNegocio": circuitos, "Medio": MEDIO_ESTATICO},
        start_date=start, end_date=end,
    )
    return int(r["Value"].iloc[0]) if len(r) else 0


def _ocupacion_bundle(catalogo: int, ocupados_actual: int, ocupados_anterior: int) -> dict[str, Any]:
    pct_actual = _round1(ocupados_actual / catalogo * 100.0) if catalogo else None
    pct_anterior = _round1(ocupados_anterior / catalogo * 100.0) if catalogo else None
    delta_pp = _round1(pct_actual - pct_anterior) if pct_actual is not None and pct_anterior is not None else None
    return {
        "catalogo": catalogo,
        "ocupados": ocupados_actual,
        "ocupados_anterior": ocupados_anterior,
        "pct_actual": pct_actual,
        "pct_anterior": pct_anterior,
        "delta_elementos": ocupados_actual - ocupados_anterior,
        "delta_pp": delta_pp,
    }


# ---------------------------------------------------------------------------
# Tarjeta 1 - Elementos del catalogo estatico
# ---------------------------------------------------------------------------


def compute_tarjeta1_catalogo(engine: MetricsEngine) -> dict[str, Any]:
    shoppings = _catalogo(engine, SHOPPINGS_CIRCUITOS)
    aa2000 = _catalogo(engine, AA2000_CIRCUITOS)
    otros = _catalogo(engine, OTROS_CIRCUITOS)
    total = _catalogo(engine, TV3_CIRCUITOS)

    if shoppings + aa2000 + otros != total:
        raise BuildError(
            f"Catalogo TV3 no reconcilia: shoppings({shoppings})+aa2000({aa2000})+otros({otros}) "
            f"!= total({total})"
        )

    pct_shoppings = _round1(shoppings / total * 100.0) if total else None
    pct_aa2000 = _round1(aa2000 / total * 100.0) if total else None
    pct_otros = _round1(otros / total * 100.0) if total else None

    return {
        "total": total,
        "shoppings": shoppings,
        "aa2000": aa2000,
        "otros": otros,
        "pct_shoppings": pct_shoppings,
        "pct_aa2000": pct_aa2000,
        "pct_otros": pct_otros,
    }


# ---------------------------------------------------------------------------
# Tarjetas 2-4 y 6 - Ocupacion del mes / Shoppings / AA2000 / Disponibles
# ---------------------------------------------------------------------------


def compute_tarjeta2_ocupacion_mes(engine: MetricsEngine, catalogo_total: int, period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    actual = _ocupados(engine, TV3_CIRCUITOS, *period)
    anterior = _ocupados(engine, TV3_CIRCUITOS, *previous)
    return _ocupacion_bundle(catalogo_total, actual, anterior)


def compute_tarjeta3_shoppings(engine: MetricsEngine, catalogo_shoppings: int, period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    """Incluye Remeros Estatico (prompt maestro Sec.4, tarjeta 3). Remeros
    Digital/LED nunca entran: quedan fuera de MEDIO_ESTATICO desde
    build_tv3_universe, no hace falta excluirlos aqui de nuevo."""
    actual = _ocupados(engine, SHOPPINGS_CIRCUITOS, *period)
    anterior = _ocupados(engine, SHOPPINGS_CIRCUITOS, *previous)
    return _ocupacion_bundle(catalogo_shoppings, actual, anterior)


def compute_tarjeta4_aa2000(engine: MetricsEngine, catalogo_aa2000: int, period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    """AA2000 Estatico (prompt maestro Sec.4, tarjeta 4): catalogo=40 esta
    confirmado, asi que un resultado de 0 ocupados es un cero real (la
    formula catalogo>0 hace que pct_actual sea 0.0, nunca None/S-D)."""
    actual = _ocupados(engine, AA2000_CIRCUITOS, *period)
    anterior = _ocupados(engine, AA2000_CIRCUITOS, *previous)
    bundle = _ocupacion_bundle(catalogo_aa2000, actual, anterior)
    if catalogo_aa2000 and bundle["pct_actual"] is None:
        raise BuildError("AA2000 Estatico: catalogo>0 pero pct_actual=None (no debe mostrarse S/D)")
    return bundle


def compute_tarjeta6_disponibles(catalogo_total: int, ocupacion_mes: dict[str, Any]) -> dict[str, Any]:
    disp_actual = catalogo_total - ocupacion_mes["ocupados"]
    disp_anterior = catalogo_total - ocupacion_mes["ocupados_anterior"]
    pct_actual = _round1(disp_actual / catalogo_total * 100.0) if catalogo_total else None
    pct_anterior = _round1(disp_anterior / catalogo_total * 100.0) if catalogo_total else None
    delta_pp = _round1(pct_actual - pct_anterior) if pct_actual is not None and pct_anterior is not None else None
    return {
        "catalogo": catalogo_total,
        "disponibles": disp_actual,
        "disponibles_anterior": disp_anterior,
        "pct_actual": pct_actual,
        "pct_anterior": pct_anterior,
        "delta_elementos": disp_actual - disp_anterior,
        "delta_pp": delta_pp,
    }


# ---------------------------------------------------------------------------
# Tarjeta 5 - Campanas del mes / Activaciones
# ---------------------------------------------------------------------------


def _campanas_y_activaciones(engine: MetricsEngine, element_ids: list[Any], start: str, end: str) -> tuple[int, int]:
    """Campanas = IDCampaña distintos vigentes en el mes (prompt maestro
    Sec.6/11: identidad de campana). Activaciones = pares distintos
    IDCampaña x ElementoID (drop_duplicates explicito: filas duplicadas del
    mismo par nunca inflan el conteo, aunque la fuente ya deduplica por
    CargaID via _campanas_overlap)."""
    overlap = engine._campanas_overlap(element_ids, start, end)
    overlap = overlap[overlap["IDCampaña"].notna() & (overlap["IDCampaña"].astype(str).str.strip() != "")]
    campanas = int(overlap["IDCampaña"].nunique())
    activaciones = int(overlap.drop_duplicates(subset=["ElementoID", "IDCampaña"]).shape[0])
    return campanas, activaciones


def compute_tarjeta5_campanas_activaciones(
    engine: MetricsEngine, element_ids: list[Any], period: tuple[str, str], previous: tuple[str, str],
) -> dict[str, Any]:
    campanas_actual, activaciones_actual = _campanas_y_activaciones(engine, element_ids, *period)
    campanas_anterior, activaciones_anterior = _campanas_y_activaciones(engine, element_ids, *previous)
    return {
        "campanas_actual": campanas_actual,
        "campanas_anterior": campanas_anterior,
        "delta_campanas": campanas_actual - campanas_anterior,
        "activaciones_actual": activaciones_actual,
        "activaciones_anterior": activaciones_anterior,
        "delta_activaciones": activaciones_actual - activaciones_anterior,
    }


# ---------------------------------------------------------------------------
# Panel izquierdo - Evolucion mensual de la ocupacion estatica (3 series %)
# ---------------------------------------------------------------------------


def compute_evolution(
    engine: MetricsEngine, report_year: int, report_month: int,
    catalogo_shoppings: int, catalogo_aa2000: int, catalogo_otros: int,
) -> dict[str, Any]:
    """Tres series porcentuales Ene-report_month, sin meses futuros: Shoppings
    / AA2000 / Otros Estatico, cada una sobre SU PROPIO catalogo (nunca sobre
    el total Core=422). Ajuste 2026-08-22 (prompt maestro "ajustes visuales"):
    reemplaza la serie Core Estatico -- que no aportaba lectura nueva sobre
    las otras dos KPI cards de arriba -- por Otros/Pilar Frontlight (catalogo
    de 1 elemento, antes solo interno). No toca ninguna otra tarjeta/panel:
    la tarjeta "Ocupacion del mes" sigue representando el Core completo
    (146/422) de forma independiente, via compute_tarjeta2_ocupacion_mes."""
    meses, shop_pct, aa_pct, otros_pct = [], [], [], []
    shop_ocup, aa_ocup, otros_ocup = [], [], []
    for m in range(1, report_month + 1):
        start, end = _period_bounds(report_year, m)
        s = _ocupados(engine, SHOPPINGS_CIRCUITOS, start, end)
        a = _ocupados(engine, AA2000_CIRCUITOS, start, end)
        o = _ocupados(engine, OTROS_CIRCUITOS, start, end)
        meses.append(MESES_ES_ABR[m - 1])
        shop_ocup.append(s)
        aa_ocup.append(a)
        otros_ocup.append(o)
        shop_pct.append(_round1(s / catalogo_shoppings * 100.0) if catalogo_shoppings else None)
        # AA2000/Otros muestran 0.0 (no None) cuando hay catalogo confirmado y 0 ocupados.
        aa_pct.append(_round1(a / catalogo_aa2000 * 100.0) if catalogo_aa2000 else None)
        otros_pct.append(_round1(o / catalogo_otros * 100.0) if catalogo_otros else None)
    return {
        "meses": meses,
        "shoppings_pct": shop_pct, "aa2000_pct": aa_pct, "otros_pct": otros_pct,
        "shoppings_ocupados": shop_ocup, "aa2000_ocupados": aa_ocup, "otros_ocupados": otros_ocup,
    }


# ---------------------------------------------------------------------------
# Panel derecho superior - Capacidad y ocupacion por shopping (17, completo)
# ---------------------------------------------------------------------------


def compute_shoppings_panel(engine: MetricsEngine, period: tuple[str, str]) -> dict[str, Any]:
    """Listado COMPLETO de los 17 SitioNegocio de Shoppings Estatico (prompt
    maestro Sec.5: "no es un ranking"): orden alfabetico estable, incluye
    sedes con cero actividad. Reconciliacion exacta contra la tarjeta 3."""
    reg_df = engine.query("elementos_registrados", group_by=["SitioNegocio"], filters={"CircuitoNegocio": SHOPPINGS_CIRCUITOS, "Medio": MEDIO_ESTATICO})
    act_df = engine.query("elementos_con_actividad", group_by=["SitioNegocio"], filters={"CircuitoNegocio": SHOPPINGS_CIRCUITOS, "Medio": MEDIO_ESTATICO}, start_date=period[0], end_date=period[1])
    act_map = dict(zip(act_df["SitioNegocio"], act_df["Value"])) if len(act_df) else {}

    rows = []
    for _, r in reg_df.sort_values("SitioNegocio").iterrows():
        sitio = r["SitioNegocio"]
        if sitio is None or (isinstance(sitio, float) and pd.isna(sitio)):
            continue
        catalogo = int(r["Value"])
        ocupados = int(act_map.get(sitio, 0))
        disponibles = catalogo - ocupados
        pct = _round1(ocupados / catalogo * 100.0) if catalogo else None
        rows.append({"sitio": sitio, "catalogo": catalogo, "ocupados": ocupados, "disponibles": disponibles, "pct": pct})

    total_catalogo = sum(r["catalogo"] for r in rows)
    total_ocupados = sum(r["ocupados"] for r in rows)
    total_disponibles = sum(r["disponibles"] for r in rows)
    sedes_con_actividad = sum(1 for r in rows if r["ocupados"] > 0)

    return {
        "rows": rows,
        "total_sedes": len(rows),
        "sedes_con_actividad": sedes_con_actividad,
        "catalogo_total": total_catalogo,
        "ocupados_total": total_ocupados,
        "disponibles_total": total_disponibles,
    }


# ---------------------------------------------------------------------------
# Clasificacion canonica de soportes (prompt maestro Sec.6): reemplaza el
# bucket "OTRO" de FormatoNegocio, que en Estatico agrupa la inmensa mayoria
# del catalogo (394/422 filas) sin distinguir tipo fisico de soporte.
#
# El repositorio no tiene columnas TipoSoporte ni TipoElemento (auditado:
# scripts/transform_data.py solo expone AplicaCantidad, CapacidadSlotsReel,
# CircuitoDashboard, Ciudad, Descripcion, DimensionOptico, DimensionTotal,
# ElementoID, Material, Medio, Nivel, Observaciones, Original, Proveedor,
# Resolucion, RevisionMaestro, SegundosDia, Subcircuito, TipoCatalogo,
# TipoInstalacion, TipoInventario, Ubicacion, b/h/m2/q). Orden de auditoria
# real, en la jerarquia pedida (estructurado antes que texto libre):
#   1. FormatoNegocio ya resuelto por Gate3B (confiable solo para
#      FRONTLIGHT/TOTEM; el resto cae en OTRO y no aporta).
#   2. TipoInstalacion (estructurado, auditado contra datos reales: "Ban
#      Chico/Grande/Mediano/Lejano/Gigante" = Banner en TODOS los casos
#      verificados, incluso cuando Descripcion nombra la ubicacion fisica
#      del banner -ej. "Octogono Chico Pasillo", "Garganta Central", "Sector
#      Deportivo"- en vez del tipo de soporte; "PLACA"; "Plot Esc*" = vinilo
#      de escalera; "Plot Micro" = vinilo microperforado).
#   3. Descripcion, por palabra clave explicita y auditable (case-insensitive,
#      sin inferir del nombre del shopping/sitio).
#   4. ElementoID como ultimo recurso estructurado (el codigo interno del
#      elemento suele incluir el tipo de soporte como prefijo/token, ej.
#      "FPBROWN-BACK-1", "FRONTLIGHTS EN RAMPAS -1"): solo se usa cuando
#      Descripcion es generica ("Carteles estacionamiento") y no permite
#      distinguir Backlight de Frontlight.
# Si ningun campo permite clasificar, "Sin clasificar" (fallback, nunca
# bloqueante): se reporta como residual en el payload y en "A atender".
# ---------------------------------------------------------------------------

SIN_CLASIFICAR = "Sin clasificar"

_FRONT_WORD_RE = re.compile(r"\bFRONT\b")


def classify_soporte(row: dict[str, Any]) -> tuple[str, str]:
    """Devuelve (categoria_canonica, campo_que_decidio). Pura/determinista:
    misma fila siempre produce la misma categoria (prompt maestro Sec.6:
    "reproducible")."""

    def _text(value: Any) -> str:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return ""
        return str(value).strip()

    fmt = _text(row.get("FormatoNegocio")).upper()
    if fmt == "FRONTLIGHT":
        return "Frontlight", "FormatoNegocio"
    if fmt == "TOTEM":
        return "Tótem", "FormatoNegocio"

    tipo_soporte = _text(row.get("TipoSoporte"))
    tipo_elemento = _text(row.get("TipoElemento"))
    if tipo_soporte and tipo_soporte.upper() not in {"OTRO", "OTROS", "N/A", "S/D"}:
        return tipo_soporte, "TipoSoporte"
    if tipo_elemento and tipo_elemento.upper() not in {"OTRO", "OTROS", "N/A", "S/D"}:
        return tipo_elemento, "TipoElemento"

    ti_lower = _text(row.get("TipoInstalacion")).lower()
    if ti_lower.startswith("ban "):
        return "Banner", "TipoInstalacion"
    if ti_lower == "placa":
        return "Placa", "TipoInstalacion"
    if ti_lower.startswith("plot esc"):
        return "Vinilo de escalera", "TipoInstalacion"
    if ti_lower.startswith("plot micro"):
        return "Vinilo microperforado", "TipoInstalacion"

    desc = _text(row.get("Descripcion")).upper()
    if "CHUPETE" in desc:
        return "Chupete backlight", "Descripcion"
    if "BACKLIGHT" in desc or "BACKLIT" in desc:
        return "Backlight", "Descripcion"
    if "BANNER" in desc:
        return "Banner", "Descripcion"
    if "ESTANDARTE" in desc:
        return "Estandarte", "Descripcion"
    if "MUPI" in desc:
        return "Mupi", "Descripcion"
    if _FRONT_WORD_RE.search(desc):
        return "Frontlight", "Descripcion"

    eid = _text(row.get("ElementoID")).upper()
    if "BACK" in eid:
        return "Backlight", "ElementoID"
    if "FRONT" in eid:
        return "Frontlight", "ElementoID"

    return SIN_CLASIFICAR, "fallback"


def compute_soportes_top(engine: MetricsEngine, universe_maestro: pd.DataFrame, period: tuple[str, str]) -> dict[str, Any]:
    """Top 3 por ACTIVACIONES (no por filas de catalogo), con clasificacion
    canonica sobre los elementos con campana vigente en el mes. Desempate
    (prompt maestro Sec.5): mas campanas unicas, luego nombre canonico asc."""
    element_ids = universe_maestro["ElementoID"].tolist()
    overlap = engine._campanas_overlap(element_ids, period[0], period[1])
    overlap = overlap[overlap["IDCampaña"].notna() & (overlap["IDCampaña"].astype(str).str.strip() != "")]
    overlap = overlap.drop_duplicates(subset=["ElementoID", "IDCampaña"])

    class_map: dict[Any, str] = {}
    for _, row in universe_maestro.iterrows():
        categoria, _fuente = classify_soporte(row.to_dict())
        class_map[row["ElementoID"]] = categoria

    overlap = overlap.copy()
    overlap["_categoria"] = overlap["ElementoID"].map(class_map)

    grouped = overlap.groupby("_categoria").agg(
        activaciones=("ElementoID", "size"),
        campanas=("IDCampaña", "nunique"),
    ).reset_index()

    clasificadas = grouped[grouped["_categoria"] != SIN_CLASIFICAR].copy()
    clasificadas = clasificadas.sort_values(
        by=["activaciones", "campanas", "_categoria"], ascending=[False, False, True],
    )
    top = [
        {"soporte": r["_categoria"], "activaciones": int(r["activaciones"]), "campanas": int(r["campanas"])}
        for _, r in clasificadas.head(SOPORTES_TOP_N).iterrows()
    ]

    sin_clasificar_row = grouped[grouped["_categoria"] == SIN_CLASIFICAR]
    residual_activaciones = int(sin_clasificar_row["activaciones"].iloc[0]) if len(sin_clasificar_row) else 0
    residual_elementos = int(
        (pd.Series(class_map) == SIN_CLASIFICAR).sum()
    )
    residual_elementos_activos = int(
        overlap.loc[overlap["_categoria"] == SIN_CLASIFICAR, "ElementoID"].nunique()
    )

    return {
        "top": top,
        "sin_clasificar": {
            "elementos_catalogo": residual_elementos,
            "elementos_activos": residual_elementos_activos,
            "activaciones": residual_activaciones,
        },
    }


# ---------------------------------------------------------------------------
# Reconciliacion de lectura con TV1 (prompt maestro Sec.7): SOLO cuando
# coincide exactamente universo/mes/unidad. No fuerza igualdad entre
# universos distintos. No modifica TV1 ni TV2 (solo lee semantic_result,
# mismo objeto ya cargado para TV3).
# ---------------------------------------------------------------------------


def compute_reconciliacion_tv1(universe: dict[str, Any]) -> dict[str, Any]:
    op = universe["op"]
    tv1_estatico = op[
        ~op["CircuitoNegocio"].isin(_TV1_ESTATICO_EXCLUDED_CIRCUITOS)
        & (op["CircuitoNegocio"] != "YPF")
        & (op["Medio"] == MEDIO_ESTATICO)
    ]
    tv1_ids = set(tv1_estatico["ElementoID"])
    tv3_ids = set(universe["element_ids"])

    comunes = tv1_ids & tv3_ids
    solo_tv1 = tv1_ids - tv3_ids
    solo_tv3 = tv3_ids - tv1_ids

    solo_tv1_rows = tv1_estatico[tv1_estatico["ElementoID"].isin(solo_tv1)]
    solo_tv1_circuitos = sorted(solo_tv1_rows["CircuitoNegocio"].unique().tolist()) if len(solo_tv1_rows) else []

    if solo_tv3:
        explicacion = f"{len(solo_tv3)} ElementoID de TV3 no estan en el universo Estatico de TV1: revisar manualmente."
    elif solo_tv1_circuitos == ["MAB"]:
        explicacion = (
            f"TV1 incluye {len(solo_tv1)} elemento(s) Estatico adicionales de MAB (PortfolioTier=COMPLEMENTARIO) "
            "que TV3 excluye porque su scope es solo Core (CENCOSUD+REMEROS+AA2000+PILAR_FRONTLIGHT). "
            "No es una discrepancia: son universos distintos por diseno."
        )
    elif solo_tv1:
        explicacion = (
            f"TV1 incluye {len(solo_tv1)} elemento(s) Estatico adicionales fuera del scope Core de TV3 "
            f"(circuitos: {solo_tv1_circuitos}). No se fuerza igualdad entre universos distintos."
        )
    else:
        explicacion = "Universo Estatico de TV1 y TV3 coinciden exactamente."

    return {
        "comunes": len(comunes),
        "solo_tv1": len(solo_tv1),
        "solo_tv3": len(solo_tv3),
        "solo_tv1_circuitos": solo_tv1_circuitos,
        "explicacion": explicacion,
        "ids_solo_tv1_muestra": sorted(str(x) for x in solo_tv1)[:20],
        "ids_solo_tv3_muestra": sorted(str(x) for x in solo_tv3)[:20],
    }


# ---------------------------------------------------------------------------
# Insights (Lectura / Punto positivo / A atender)
# ---------------------------------------------------------------------------


def compute_insights(
    tarjeta1: dict[str, Any], tarjeta2: dict[str, Any], tarjeta4: dict[str, Any],
    tarjeta5: dict[str, Any], tarjeta6: dict[str, Any],
    shoppings_panel: dict[str, Any], soportes: dict[str, Any],
    report_month_label: str, previous_month_label: str,
) -> dict[str, str]:
    lectura = (
        f"{report_month_label} registra {_fmt_es_int(tarjeta2['ocupados'])} de {_fmt_es_int(tarjeta1['total'])} "
        f"elementos del Core Estático con campaña (<b>{_fmt_es_pct(tarjeta2['pct_actual'])}%</b>), con "
        f"{_fmt_es_int(tarjeta5['campanas_actual'])} campañas y {_fmt_es_int(tarjeta5['activaciones_actual'])} "
        f"activaciones vigentes en el mes."
    )

    if soportes["top"]:
        top = soportes["top"][0]
        punto_positivo = (
            f"<b>{top['soporte']}</b> es el soporte más vendido de {report_month_label.lower()} con "
            f"<b>{_fmt_es_int(top['activaciones'])} activaciones</b> en {_fmt_es_int(top['campanas'])} campañas; "
            f"{_fmt_es_int(shoppings_panel['sedes_con_actividad'])} de {_fmt_es_int(shoppings_panel['total_sedes'])} "
            f"shoppings tuvieron actividad."
        )
    else:
        punto_positivo = f"El Core Estático mantiene {_fmt_es_int(tarjeta2['ocupados'])} elementos con campaña en {report_month_label.lower()}."

    atender_parts = []
    if tarjeta4["catalogo"] and tarjeta4["ocupados"] == 0:
        atender_parts.append(
            f"AA2000 Estático no registra campañas en {report_month_label.lower()} "
            f"(0 de {_fmt_es_int(tarjeta4['catalogo'])} elegibles)"
        )
    if tarjeta2["delta_pp"] is not None and tarjeta2["delta_pp"] < 0:
        atender_parts.append(f"la ocupación del Core Estático cae <b>{_fmt_es_pct(abs(tarjeta2['delta_pp']))} pp</b> frente a {previous_month_label.lower()}")
    residual = soportes["sin_clasificar"]
    if residual["activaciones"] > 0:
        atender_parts.append(
            f"{_fmt_es_int(residual['activaciones'])} activaciones de {report_month_label.lower()} quedan "
            f"\"Sin clasificar\" en el panel de soportes"
        )
    if not atender_parts:
        atender_parts.append(f"el Core Estático conserva <b>{_fmt_es_int(tarjeta6['disponibles'])} elementos disponibles</b> por vender")
    joined = "; ".join(atender_parts)
    a_atender = joined[:1].upper() + joined[1:] + "."

    return {"lectura": lectura, "punto_positivo": punto_positivo, "a_atender": a_atender}


# ---------------------------------------------------------------------------
# Logo (reutilizado byte-a-byte, igual patron que TV1/TV2)
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


def build_tv3_data(path: str | Path = vi.DEFAULT_INPUT_PATH) -> dict[str, Any]:
    path = Path(path)
    sha_before = vi.calculate_sha256(path)

    _transform_result, semantic_result, engine = load_pipeline(path)
    universe = build_tv3_universe(semantic_result)

    period = _period_bounds(REPORT_YEAR, REPORT_MONTH)
    prev_year, prev_month = _previous_month(REPORT_YEAR, REPORT_MONTH)
    previous = _period_bounds(prev_year, prev_month)

    tarjeta1 = compute_tarjeta1_catalogo(engine)
    tarjeta2 = compute_tarjeta2_ocupacion_mes(engine, tarjeta1["total"], period, previous)
    tarjeta3 = compute_tarjeta3_shoppings(engine, tarjeta1["shoppings"], period, previous)
    tarjeta4 = compute_tarjeta4_aa2000(engine, tarjeta1["aa2000"], period, previous)
    tarjeta5 = compute_tarjeta5_campanas_activaciones(engine, universe["element_ids"], period, previous)
    tarjeta6 = compute_tarjeta6_disponibles(tarjeta1["total"], tarjeta2)

    if tarjeta2["ocupados"] != tarjeta6["catalogo"] - tarjeta6["disponibles"]:
        raise BuildError("Tarjeta 6 (disponibles) no reconcilia con tarjeta 2 (ocupados)")

    evolution = compute_evolution(engine, REPORT_YEAR, REPORT_MONTH, tarjeta1["shoppings"], tarjeta1["aa2000"], tarjeta1["otros"])
    shoppings_panel = compute_shoppings_panel(engine, period)
    if shoppings_panel["catalogo_total"] != tarjeta1["shoppings"] or shoppings_panel["ocupados_total"] != tarjeta3["ocupados"]:
        raise BuildError(
            f"Panel de shoppings no reconcilia con tarjetas: catalogo={shoppings_panel['catalogo_total']} "
            f"(esperado {tarjeta1['shoppings']}), ocupados={shoppings_panel['ocupados_total']} (esperado {tarjeta3['ocupados']})"
        )
    if shoppings_panel["total_sedes"] != 17:
        raise BuildError(f"Panel de shoppings: se esperaban 17 sedes, se obtuvieron {shoppings_panel['total_sedes']}")

    soportes = compute_soportes_top(engine, universe["maestro"], period)
    reconciliacion_tv1 = compute_reconciliacion_tv1(universe)
    insights = compute_insights(
        tarjeta1, tarjeta2, tarjeta4, tarjeta5, tarjeta6, shoppings_panel, soportes,
        MESES_ES[REPORT_MONTH - 1], MESES_ES[prev_month - 1],
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
            "timezone": ART_TZ_NAME,
            "report_year": REPORT_YEAR,
            "report_month": REPORT_MONTH,
            "report_month_label": MESES_ES[REPORT_MONTH - 1],
            "previous_month": prev_month,
            "previous_month_label": MESES_ES[prev_month - 1],
            "period_start": period[0],
            "period_end": period[1],
            "fuente": "OCU26 · Base maestra + base campañas",
        },
        "tarjetas": {
            "catalogo": tarjeta1,
            "ocupacion_mes": tarjeta2,
            "shoppings_estatico": tarjeta3,
            "aa2000_estatico": tarjeta4,
            "campanas_activaciones": tarjeta5,
            "disponibles": tarjeta6,
        },
        "evolution": evolution,
        "shoppings": shoppings_panel,
        "soportes_top": soportes["top"],
        "soportes_sin_clasificar": soportes["sin_clasificar"],
        "reconciliacion_tv1": {
            "comunes": reconciliacion_tv1["comunes"],
            "solo_tv1": reconciliacion_tv1["solo_tv1"],
            "solo_tv3": reconciliacion_tv1["solo_tv3"],
            "solo_tv1_circuitos": reconciliacion_tv1["solo_tv1_circuitos"],
            "explicacion": reconciliacion_tv1["explicacion"],
        },
        "insights": insights,
    }

    return {
        "data": data,
        "sha256": sha_after,
        "universe": {"circuitos": universe["circuitos"]},
        "reconciliacion_tv1_detalle": reconciliacion_tv1,
    }


def render_html(data: dict[str, Any]) -> str:
    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    logo_tag = extract_logo_img_tag(REFERENCE_PATH)
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    html = template.replace("{{LOGO_IMG_TAG}}", logo_tag)
    html = html.replace("{{TV3_DATA_JSON}}", payload)
    return html


def build_and_write(
    path: str | Path = vi.DEFAULT_INPUT_PATH,
    output_html: str | Path = DEFAULT_OUTPUT_HTML,
    output_json: str | Path = DEFAULT_OUTPUT_JSON,
) -> dict[str, Any]:
    result = build_tv3_data(path)
    html = render_html(result["data"])

    output_html = Path(output_html)
    output_html.write_text(html, encoding="utf-8")

    output_json = Path(output_json)
    output_json.parent.mkdir(parents=True, exist_ok=True)
    with open(output_json, "w", encoding="utf-8") as fh:
        json.dump(result["data"], fh, ensure_ascii=False, indent=2)

    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Construye tv3.html (dashboard TV3 OCU26) con datos reales.")
    parser.add_argument("--file", default=str(vi.DEFAULT_INPUT_PATH), help="Ruta al archivo .xlsx a leer")
    parser.add_argument("--output-html", default=str(DEFAULT_OUTPUT_HTML), help="Ruta del HTML productivo generado")
    parser.add_argument("--output-json", default=str(DEFAULT_OUTPUT_JSON), help="Ruta del snapshot JSON generado")
    args = parser.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    try:
        result = build_and_write(args.file, args.output_html, args.output_json)
    except BuildError as exc:
        print("TV3_BUILD_ERROR:", exc)
        return 1

    print("TV3_BUILD_OK")
    print(json.dumps(result["data"]["meta"], ensure_ascii=False, indent=2))
    print(json.dumps(result["data"]["tarjetas"], ensure_ascii=False, indent=2))
    print("SOPORTES_TOP:", json.dumps(result["data"]["soportes_top"], ensure_ascii=False))
    print("SIN_CLASIFICAR:", json.dumps(result["data"]["soportes_sin_clasificar"], ensure_ascii=False))
    print("RECONCILIACION_TV1:", json.dumps(result["data"]["reconciliacion_tv1"], ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
