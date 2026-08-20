"""Capa de datos para el dashboard TV4 - OCU26 (Pipeline Comercial).

Se ejecuta DESPUES de scripts/semantic_model.py y scripts/metrics_engine.py
(Gate 3B). Reutiliza export_data.load_pipeline (Gate 4A), mismo patron que
build_tv1/tv2/tv3_dashboard.py: no reabre el Excel, no reimplementa reglas
de negocio de Gate 3 (Medio/CircuitoNegocio/PortfolioTier/IncluyeConteoGeneral
ya vienen resueltos por semantic_model.py sobre semantic_result["campanas"],
que es CAMPANAS con el mismo join semantico que MAESTRO_ELEMENTOS).

Universo TV4 = "Core Comercial" (pipeline) = PortfolioTier CORE excluyendo
YPF: CENCOSUD + REMEROS (Shoppings, Digital+Estatico) + PANTALLAS_LED +
PILAR_FRONTLIGHT + AA2000. YPF queda fuera por decision explicita del
usuario (tendra su propia TV, no forma parte de este scope). APSA/London
Supply nunca entran (no son CORE). 513 elementos = mismo numero "no YPF"
ya validado en TV1 (CM1 Sec.1: 513 no YPF + 451 YPF = 964).

TV4 mide pipeline temporal (que corre / que viene), no un mes calendario:
no hay comparativa mes anterior. El "hoy" operativo se ancla al cierre del
periodo de reporte ya usado por TV1-3 (31/07/2026, mismo "corte" que Julio
2026 como mes vigente), no a la fecha real del sistema: los datos de
CAMPANAS llegan reservados hasta 2027-03-31 y el campo Estado (Activa/
Reservada/Finalizada) no es fecha-consistente por si solo (hay filas
Reservada con FechaInicio ya pasada respecto a hoy real, y filas Activa
con FechaInicio futura) -- por eso TODA clasificacion temporal (en curso/
futura/finalizada) se resuelve por FechaInicio/FechaFin contra el corte,
igual que el resto del motor (_campanas_overlap), nunca por Estado.

Conteos "campanas unicas" vs "activaciones": una campana (IDCampaña) puede
tener muchas filas (una por ElementoID reservado). Los KPIs de eventos
comerciales (reservas futuras, inician 30d, finalizan 30d) cuentan eventos
distintos de (IDCampaña, fecha relevante) para no colapsar inicios/fines
genuinamente distintos de la misma campana en distintas fechas/circuitos,
y para no duplicar el mismo inicio repetido en N elementos. La card de
actividad actual sigue el patron ya usado en TV1 KPI2 (_distinct_campanas):
COUNT DISTINCT IDCampaña sobre el scope/periodo dado.

Uso:
    python scripts/build_tv4_dashboard.py
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
TEMPLATE_PATH = Path(__file__).resolve().parent / "templates" / "tv4_template.html"
REFERENCE_PATH = REPO_ROOT / "audit_sources" / "TV4_REFERENCE.html.html"
DEFAULT_OUTPUT_HTML = REPO_ROOT / "tv4.html"
DEFAULT_OUTPUT_JSON = REPO_ROOT / "output" / "tv4_data.json"

MESES_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]
MESES_ES_ABR = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

# Corte operativo TV4 (spec Sec.3): cierre del mismo periodo de reporte
# vigente en TV1-3 (Julio 2026), no la fecha real del sistema.
REPORT_YEAR = 2026
REPORT_MONTH = 7
WINDOW_DAYS = 30

# Universo TV4 (Core Comercial, YPF excluido -- spec Sec.2/3): PortfolioTier
# CORE sin YPF. APSA/London Supply nunca forman parte de CORE (business_
# semantics.json: portfolio_tier LEGACY/COMPLEMENTARIO respectivamente).
TV4_CIRCUITOS = ["CENCOSUD", "REMEROS", "PANTALLAS_LED", "PILAR_FRONTLIGHT", "AA2000"]
MEDIO_DIGITAL = "Digital"
MEDIO_ESTATICO = "Estático"

TIMELINE_TOP_N = 5
FAMILIA_DIST_MAX = 4


class BuildError(Exception):
    """Error bloqueante al construir el dashboard TV4."""


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


# ---------------------------------------------------------------------------
# Universo TV4 y scope de campanas (Core Comercial, excluye YPF/APSA/London)
# ---------------------------------------------------------------------------


def build_tv4_universe(semantic_result: dict[str, Any]) -> dict[str, Any]:
    maestro = semantic_result["maestro"]
    config = semantic_result["config"]

    op = filter_universe(maestro, "OPERATIVO_GENERAL", config)
    tv4_maestro = op[op["CircuitoNegocio"].isin(TV4_CIRCUITOS)].copy()
    if tv4_maestro.empty:
        raise BuildError("Universo TV4 vacio: revisar business_semantics.json / TV4_CIRCUITOS")

    campanas = semantic_result["campanas"]
    op_campanas = filter_universe(campanas, "OPERATIVO_GENERAL", config)
    scope = op_campanas[op_campanas["CircuitoNegocio"].isin(TV4_CIRCUITOS)].copy()
    scope = scope[scope["Estado"] != "Cancelado"]

    return {
        "circuitos": sorted(tv4_maestro["CircuitoNegocio"].unique().tolist()),
        "elementos": int(tv4_maestro["ElementoID"].nunique()),
        "scope_campanas": scope,
        "general_campanas": op_campanas[op_campanas["Estado"] != "Cancelado"],
    }


def _familia(row: pd.Series) -> str:
    """Grupo comercial ejecutivo (spec Sec.2/6): nomenclatura vigente UI,
    nunca 'Fijo' (categoria interna del pipeline se mantiene igual, solo
    cambia el label de presentacion)."""
    circuito = row["CircuitoNegocio"]
    if circuito == "PANTALLAS_LED":
        return "Pantallas LED"
    if circuito in ("CENCOSUD", "REMEROS"):
        return "Shoppings Digital" if row["Medio"] == MEDIO_DIGITAL else "Shoppings Estático"
    return "AA2000 / Pilar Frontlight"


def _add_familia(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["_familia"] = df.apply(_familia, axis=1)
    return df


# ---------------------------------------------------------------------------
# Clasificacion temporal (en curso / futuras / inician / finalizan / historico)
# resuelta por fecha contra el corte, nunca por el campo Estado (docstring
# del modulo: Estado no es fecha-consistente en la fuente).
# ---------------------------------------------------------------------------


def _en_curso_mask(df: pd.DataFrame, cutoff: pd.Timestamp) -> pd.Series:
    indefinida_ok = (df["FechaIndefinida"] == "Si") & df["FechaFin"].isna()
    eff_fin = df["FechaFin"].where(~indefinida_ok, pd.Timestamp("2100-01-01"))
    return df["FechaInicio"].notna() & (df["FechaInicio"] <= cutoff) & (eff_fin >= cutoff)


def _count_events(df: pd.DataFrame, date_col: str) -> int:
    """Eventos comerciales distintos = (IDCampaña, fecha relevante) unicos
    (spec Sec.3.C/G): evita duplicar el mismo inicio/fin repetido en N
    elementos, pero conserva inicios/fines genuinamente distintos de la
    misma campana en fechas diferentes."""
    d = df[["IDCampaña", date_col]].dropna()
    return int(d.drop_duplicates().shape[0])


def compute_pipeline_windows(scope: pd.DataFrame, cutoff: pd.Timestamp) -> dict[str, Any]:
    window_end = cutoff + pd.Timedelta(days=WINDOW_DAYS)
    indefinida_ok = (scope["FechaIndefinida"] == "Si") & scope["FechaFin"].isna()

    en_curso = scope[_en_curso_mask(scope, cutoff)]
    futuras = scope[scope["FechaInicio"].notna() & (scope["FechaInicio"] > cutoff)]
    inician_30d = futuras[futuras["FechaInicio"] <= window_end]
    finalizan_30d = en_curso[
        en_curso["FechaFin"].notna() & (en_curso["FechaFin"] > cutoff)
        & (en_curso["FechaFin"] <= window_end) & (~indefinida_ok.loc[en_curso.index])
    ]
    historico = scope[
        scope["FechaFin"].notna() & (scope["FechaFin"] < cutoff) & (~indefinida_ok)
    ]

    return {
        "window_end": window_end,
        "en_curso": en_curso,
        "futuras": futuras,
        "inician_30d": inician_30d,
        "finalizan_30d": finalizan_30d,
        "historico": historico,
    }


def compute_kpi1_actividad_actual(en_curso: pd.DataFrame) -> dict[str, Any]:
    """Actividad actual (spec Sec.5 Card1): mismo patron que TV1 KPI2
    (_distinct_campanas) -- COUNT DISTINCT IDCampaña, sin agrupar por fecha
    (es una foto del momento, no una serie de eventos)."""
    campanas_unicas = int(en_curso["IDCampaña"].dropna().nunique())
    return {"campanas_unicas": campanas_unicas, "activaciones": int(len(en_curso))}


def compute_kpi_evento(df: pd.DataFrame, date_col: str) -> dict[str, Any]:
    return {"campanas_unicas": _count_events(df, date_col), "activaciones": int(len(df))}


def compute_kpi5_pipeline_core(en_curso_core: pd.DataFrame, general_campanas: pd.DataFrame, cutoff: pd.Timestamp) -> dict[str, Any]:
    """% Pipeline en Core (spec Sec.3.E): el propio Core como denominador de
    si mismo da 100% tautologico (asi salia en la referencia legacy,
    pctCore=100.0 -- no se reutiliza). Se usa como denominador el universo
    comercial general vigente (OPERATIVO_GENERAL, incluye YPF/Cencomedia/
    MAB, excluye solo APSA que no cuenta en conteo general), a grano
    campana unica (no activacion): a nivel activacion el % queda dominado
    por el volumen fisico de elementos YPF (~7%) y no es informativo; a
    nivel campana unica mide cuantas de las campanas activas del universo
    general tocan el Core, que es la pregunta relevante para este tablero."""
    core_campanas = set(en_curso_core["IDCampaña"].dropna().unique())
    general_en_curso = general_campanas[_en_curso_mask(general_campanas, cutoff)]
    general_campanas_set = set(general_en_curso["IDCampaña"].dropna().unique())
    denom = len(general_campanas_set)
    pct = _round1(len(core_campanas) / denom * 100.0) if denom else None
    return {
        "pct": pct,
        "campanas_core": len(core_campanas),
        "campanas_general": denom,
    }


def compute_distribucion_familia(en_curso_core: pd.DataFrame) -> list[dict[str, Any]]:
    """Activaciones activas hoy por grupo comercial (spec Sec.6, max 4
    grupos, sin inventar 'Otros'): las 4 familias reales del scope TV4."""
    fam = _add_familia(en_curso_core)
    counts = fam["_familia"].value_counts().to_dict()
    familias = ["Pantallas LED", "Shoppings Digital", "Shoppings Estático", "AA2000 / Pilar Frontlight"]
    rows = [{"familia": f, "activaciones": int(counts.get(f, 0))} for f in familias]
    rows.sort(key=lambda r: -r["activaciones"])
    return rows


def compute_estado_pipeline(windows: dict[str, Any]) -> dict[str, Any]:
    """Estado del pipeline (spec Sec.6, subtitulo 'activaciones cargadas'):
    grano activacion, igual unidad que el desglose por familia."""
    return {
        "activas": int(len(windows["en_curso"])),
        "reservas": int(len(windows["futuras"])),
        "finalizadas_historico": int(len(windows["historico"])),
    }


def compute_proximas_activaciones(inician_30d: pd.DataFrame) -> tuple[list[dict[str, Any]], int]:
    """Timeline de proximas activaciones (spec Sec.7): agrupa por (IDCampaña,
    FechaInicio) para no duplicar la misma campana por sus N ElementoID pero
    conservar inicios genuinamente distintos de la misma campana. Orden
    fecha asc -> campana asc; maximo 5 visibles, total real en el KPI."""
    if inician_30d.empty:
        return [], 0

    fam = _add_familia(inician_30d)
    grouped = (
        fam.groupby(["IDCampaña", "FechaInicio"])
        .agg(campana=("Campaña", "first"), familias=("_familia", lambda s: sorted(set(s))))
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
            "familia": " + ".join(r["familias"]),
            "badge": "Inicia",
        })
    return rows, total


# ---------------------------------------------------------------------------
# Insights (Lectura / Punto positivo / A atender) - seleccion programatica,
# mismo espiritu que TV3 compute_insights: prioridad declarada en el prompt
# maestro TV4 Sec.8, nunca causas inventadas ni comparativa mensual.
# ---------------------------------------------------------------------------


def compute_insights(
    actividad: dict[str, Any], reservas: dict[str, Any], inician: dict[str, Any],
    finalizan: dict[str, Any], distribucion: list[dict[str, Any]],
) -> dict[str, str]:
    lectura = (
        f"Hoy corren <b>{_fmt_es_int(actividad['campanas_unicas'])} campañas únicas activas</b> en "
        f"{_fmt_es_int(actividad['activaciones'])} activaciones del Core Comercial. El pipeline registra "
        f"{_fmt_es_int(reservas['campanas_unicas'])} reservas futuras, con {_fmt_es_int(inician['campanas_unicas'])} "
        f"campañas que inician y {_fmt_es_int(finalizan['campanas_unicas'])} que finalizan en los próximos 30 días."
    )

    # PUNTO POSITIVO (prioridad Sec.8): 1) volumen relevante de inicios;
    # 2) reservas futuras; 3) diversificacion por familia; 4) fallback factual.
    familias_activas = [f for f in distribucion if f["activaciones"] > 0]
    if inician["campanas_unicas"] >= 5:
        punto_positivo = (
            f"Inician <b>{_fmt_es_int(inician['campanas_unicas'])} campañas nuevas</b> en los próximos 30 días, "
            f"sumando {_fmt_es_int(inician['activaciones'])} activaciones al pipeline."
        )
    elif reservas["campanas_unicas"] > 0:
        punto_positivo = (
            f"Hay <b>{_fmt_es_int(reservas['campanas_unicas'])} reservas futuras</b> cargadas, que suman "
            f"{_fmt_es_int(reservas['activaciones'])} activaciones adicionales al pipeline."
        )
    elif len(familias_activas) >= 2:
        top2 = familias_activas[:2]
        detalle = " y ".join(f"{f['familia']} ({_fmt_es_int(f['activaciones'])})" for f in top2)
        punto_positivo = f"Buena diversificación de la actividad del Core: {detalle} mantienen actividad simultánea hoy."
    else:
        punto_positivo = f"El Core Comercial mantiene {_fmt_es_int(actividad['activaciones'])} activaciones en curso hoy."

    # A ATENDER (prioridad Sec.8): 1) finalizan > inician; 2) concentracion
    # excesiva en una familia; 3) sin reservas futuras; 4) fallback factual.
    total_familias = sum(f["activaciones"] for f in distribucion) or 1
    top_familia = max(distribucion, key=lambda f: f["activaciones"]) if distribucion else None
    top_share = _round1(top_familia["activaciones"] / total_familias * 100.0) if top_familia else None

    if finalizan["campanas_unicas"] > 0 and finalizan["campanas_unicas"] > inician["campanas_unicas"]:
        a_atender = (
            f"Finalizan <b>{_fmt_es_int(finalizan['campanas_unicas'])} campañas</b> en los próximos 30 días "
            f"frente a {_fmt_es_int(inician['campanas_unicas'])} que inician."
        )
    elif top_share is not None and top_share >= 70:
        a_atender = (
            f"<b>{top_familia['familia']}</b> concentra {str(top_share).replace('.', ',')}% de la actividad "
            f"actual del Core; el resto del pipeline queda poco diversificado."
        )
    elif reservas["campanas_unicas"] == 0:
        a_atender = "No hay reservas futuras cargadas: el pipeline depende únicamente de la actividad ya en curso."
    else:
        a_atender = (
            f"El Core Comercial libera {_fmt_es_int(finalizan['activaciones'])} activaciones por finalización "
            f"en los próximos 30 días; conviene anticipar renovación."
        )

    return {"lectura": lectura, "punto_positivo": punto_positivo, "a_atender": a_atender}


# ---------------------------------------------------------------------------
# Logo (reutilizado byte-a-byte, igual patron que TV1/TV2/TV3)
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


def build_tv4_data(path: str | Path = vi.DEFAULT_INPUT_PATH) -> dict[str, Any]:
    path = Path(path)
    sha_before = vi.calculate_sha256(path)

    _transform_result, semantic_result, _engine = load_pipeline(path)
    universe = build_tv4_universe(semantic_result)
    scope = universe["scope_campanas"]

    _, cutoff = _period_bounds(REPORT_YEAR, REPORT_MONTH)  # cierre del mes de reporte vigente

    windows = compute_pipeline_windows(scope, cutoff)

    kpi1_actividad = compute_kpi1_actividad_actual(windows["en_curso"])
    kpi2_reservas = compute_kpi_evento(windows["futuras"], "FechaInicio")
    kpi3_inician = compute_kpi_evento(windows["inician_30d"], "FechaInicio")
    kpi4_finalizan = compute_kpi_evento(windows["finalizan_30d"], "FechaFin")
    kpi5_core = compute_kpi5_pipeline_core(windows["en_curso"], universe["general_campanas"], cutoff)

    incomplete_rows = int(
        (scope["FechaInicio"].isna() | (scope["FechaFin"].isna() & ~((scope["FechaIndefinida"] == "Si") & scope["FechaFin"].isna()))).sum()
    )

    distribucion = compute_distribucion_familia(windows["en_curso"])
    estado_pipeline = compute_estado_pipeline(windows)
    timeline_rows, timeline_total = compute_proximas_activaciones(windows["inician_30d"])
    insights = compute_insights(kpi1_actividad, kpi2_reservas, kpi3_inician, kpi4_finalizan, distribucion)

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
            "actividad_actual": kpi1_actividad,
            "reservas_futuras": kpi2_reservas,
            "inician_30d": kpi3_inician,
            "finalizan_30d": kpi4_finalizan,
            "pipeline_core": kpi5_core,
        },
        "estado_pipeline": estado_pipeline,
        "distribucion_familia": distribucion,
        "timeline": {"rows": timeline_rows, "total": timeline_total},
        "insights": insights,
    }

    return {
        "data": data,
        "sha256": sha_after,
        "universe": {"circuitos": universe["circuitos"], "elementos": universe["elementos"]},
    }


def render_html(data: dict[str, Any]) -> str:
    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    logo_tag = extract_logo_img_tag(REFERENCE_PATH)
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    html = template.replace("{{LOGO_IMG_TAG}}", logo_tag)
    html = html.replace("{{TV4_DATA_JSON}}", payload)
    return html


def build_and_write(
    path: str | Path = vi.DEFAULT_INPUT_PATH,
    output_html: str | Path = DEFAULT_OUTPUT_HTML,
    output_json: str | Path = DEFAULT_OUTPUT_JSON,
) -> dict[str, Any]:
    result = build_tv4_data(path)
    html = render_html(result["data"])

    output_html = Path(output_html)
    output_html.write_text(html, encoding="utf-8")

    output_json = Path(output_json)
    output_json.parent.mkdir(parents=True, exist_ok=True)
    with open(output_json, "w", encoding="utf-8") as fh:
        json.dump(result["data"], fh, ensure_ascii=False, indent=2)

    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Construye tv4.html (dashboard TV4 OCU26) con datos reales.")
    parser.add_argument("--file", default=str(vi.DEFAULT_INPUT_PATH), help="Ruta al archivo .xlsx a leer")
    parser.add_argument("--output-html", default=str(DEFAULT_OUTPUT_HTML), help="Ruta del HTML productivo generado")
    parser.add_argument("--output-json", default=str(DEFAULT_OUTPUT_JSON), help="Ruta del snapshot JSON generado")
    args = parser.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    try:
        result = build_and_write(args.file, args.output_html, args.output_json)
    except BuildError as exc:
        print("TV4_BUILD_ERROR:", exc)
        return 1

    print("TV4_BUILD_OK")
    print(json.dumps(result["data"]["meta"], ensure_ascii=False, indent=2))
    print(json.dumps(result["data"]["kpis"], ensure_ascii=False, indent=2))
    print("UNIVERSE:", json.dumps(result["universe"], ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
