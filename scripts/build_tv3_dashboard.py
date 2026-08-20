"""Capa de datos para el dashboard TV3 - OCU26 (Core Comercial Estatico).

Se ejecuta DESPUES de scripts/semantic_model.py y scripts/metrics_engine.py
(Gate 3B). Reutiliza export_data.load_pipeline (Gate 4A), mismo patron que
build_tv1_dashboard.py / build_tv2_dashboard.py: no reabre el Excel, no
reimplementa reglas de negocio de Gate 3 (toda cifra sale de
MetricsEngine.query()). No importa build_tv1_dashboard.py ni
build_tv2_dashboard.py a proposito: cada TV es un builder independiente
sobre el mismo pipeline compartido.

Universo TV3 = CORE COMERCIAL ESTATICO = Shoppings Estatico (Cencosud +
Remeros) + AA2000 Estatico + Pilar Frontlight. YPF/APSA/London
Supply/Cencomedia/MAB excluidos (auditoria consolidada CM1 Sec.1: YPF
Estatico tiene 0% de actividad real -confirmado en TV1- y no forma parte
del scope comercial disponible; Cencomedia/MAB son PortfolioTier
COMPLEMENTARIO, no CORE).

KPI cards: mismo sistema que TV2 (COUNT DISTINCT ElementoID con campana /
COUNT DISTINCT ElementoID elegible, Medio=Estatico), no el ratio dia-based
de TV1 (ocupacion_calendario_pct): ver compute_reconciliacion_tv1 para la
validacion cruzada contra la metrica canonica de TV1 sobre el mismo scope.

Uso:
    python scripts/build_tv3_dashboard.py
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
TEMPLATE_PATH = Path(__file__).resolve().parent / "templates" / "tv3_template.html"
REFERENCE_PATH = REPO_ROOT / "audit_sources" / "TV1_REFERENCE.html.html"
DEFAULT_OUTPUT_HTML = REPO_ROOT / "tv3.html"
DEFAULT_OUTPUT_JSON = REPO_ROOT / "output" / "tv3_data.json"

MESES_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]
MESES_ES_ABR = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

# Periodo de referencia TV3 (mismo patron que TV1/TV2): Julio 2026 vs Junio 2026.
REPORT_YEAR = 2026
REPORT_MONTH = 7

# Universo TV3 (auditoria consolidada CM1 Sec.1): Core Comercial Estatico =
# Shoppings Estatico (Cencosud+Remeros) + AA2000 + Pilar Frontlight, siempre
# scope Medio=Estatico y PortfolioTier=CORE. YPF/APSA/London nunca entran
# (ni siquiera se listan aqui); Cencomedia/MAB son COMPLEMENTARIO, no CORE.
SHOPPINGS_CIRCUITOS = ["CENCOSUD", "REMEROS"]
AA2000_CIRCUITOS = ["AA2000"]
PILAR_CIRCUITOS = ["PILAR_FRONTLIGHT"]
TV3_CIRCUITOS = SHOPPINGS_CIRCUITOS + AA2000_CIRCUITOS + PILAR_CIRCUITOS
MEDIO_ESTATICO = "Estático"

RANKING_SHOPPINGS_TOP_N = 5
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
        "circuitos": sorted(tv3_maestro["CircuitoNegocio"].unique().tolist()),
        "element_ids": tv3_maestro["ElementoID"].tolist(),
    }


# ---------------------------------------------------------------------------
# Ocupacion por calendario (mismo sistema que TV2: COUNT DISTINCT
# ElementoID con campana / COUNT DISTINCT ElementoID elegible)
# ---------------------------------------------------------------------------


def _distinct_activos_registrados(
    engine: MetricsEngine, circuitos: list[str], period: tuple[str, str], previous: tuple[str, str],
) -> dict[str, Any]:
    """COUNT DISTINCT ElementoID Estatico con campana / COUNT DISTINCT
    ElementoID Estatico elegible, scope=circuitos (spec TV3 Sec.3: mismo
    sistema que TV2, recalculado con Medio=Estático)."""
    reg = engine.query("elementos_registrados", filters={"CircuitoNegocio": circuitos, "Medio": MEDIO_ESTATICO})
    reg_value = int(reg["Value"].iloc[0]) if len(reg) else 0

    act = engine.query("elementos_con_actividad", filters={"CircuitoNegocio": circuitos, "Medio": MEDIO_ESTATICO}, start_date=period[0], end_date=period[1])
    act_value = int(act["Value"].iloc[0]) if len(act) else 0

    act_prev = engine.query("elementos_con_actividad", filters={"CircuitoNegocio": circuitos, "Medio": MEDIO_ESTATICO}, start_date=previous[0], end_date=previous[1])
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
    return _distinct_activos_registrados(engine, TV3_CIRCUITOS, period, previous)


def compute_kpi2_shoppings_estatico(engine: MetricsEngine, period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    return _distinct_activos_registrados(engine, SHOPPINGS_CIRCUITOS, period, previous)


def compute_kpi3_aa2000_estatico(engine: MetricsEngine, period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    """AA2000 Estatico (spec TV3 Sec.3, preferencia explicita): CompletitudMaestro
    PARCIAL (faltan Mendoza/Cordoba) hace NO_APLICA la metrica dia-based
    (ocupacion_calendario_pct), pero elementos_con_actividad/registrados
    (conteo, no requiere cobertura completa) SI son validos y hoy devuelven
    0 real (no invento el cero: viene de la query, MetricStatus=OK)."""
    return _distinct_activos_registrados(engine, AA2000_CIRCUITOS, period, previous)


def compute_pilar_frontlight(engine: MetricsEngine, period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    """Pilar Frontlight: familia Core real, uso interno (reconciliacion +
    insights), no tiene tarjeta propia (universo de 1 elemento)."""
    return _distinct_activos_registrados(engine, PILAR_CIRCUITOS, period, previous)


def _assert_family_reconciles_with_core(core: dict, shop: dict, aa2000: dict, pilar: dict) -> None:
    elegibles_sum = shop["elegibles"] + aa2000["elegibles"] + pilar["elegibles"]
    activos_sum = shop["activos"] + aa2000["activos"] + pilar["activos"]
    if elegibles_sum != core["elegibles"]:
        raise BuildError(f"Elegibles no reconcilian: familias={elegibles_sum} vs Core={core['elegibles']}")
    if activos_sum != core["activos"]:
        raise BuildError(f"Activos no reconcilian: familias={activos_sum} vs Core={core['activos']}")


# ---------------------------------------------------------------------------
# KPI 4 - Cobertura de shoppings (sitios con actividad / sitios elegibles)
# ---------------------------------------------------------------------------


def compute_kpi4_cobertura_shoppings(engine: MetricsEngine, period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    """Metrica ejecutiva de actividad Core Estatico (spec TV3 Sec.3, Card4):
    cantidad de shoppings (SitioNegocio) con >=1 campana estatica en el mes
    sobre el total de shoppings con inventario Estatico registrado. Misma
    dimension ya usada en el ranking (SitioNegocio), sin logica de negocio
    nueva."""
    reg_df = engine.query("elementos_registrados", group_by=["SitioNegocio"], filters={"CircuitoNegocio": SHOPPINGS_CIRCUITOS, "Medio": MEDIO_ESTATICO})
    act_df = engine.query("elementos_con_actividad", group_by=["SitioNegocio"], filters={"CircuitoNegocio": SHOPPINGS_CIRCUITOS, "Medio": MEDIO_ESTATICO}, start_date=period[0], end_date=period[1])
    act_prev_df = engine.query("elementos_con_actividad", group_by=["SitioNegocio"], filters={"CircuitoNegocio": SHOPPINGS_CIRCUITOS, "Medio": MEDIO_ESTATICO}, start_date=previous[0], end_date=previous[1])

    total_sitios = int(len(reg_df))
    activos = int((act_df["Value"] > 0).sum())
    activos_prev = int((act_prev_df["Value"] > 0).sum())

    pct_actual = _round1(activos / total_sitios * 100.0) if total_sitios else None
    pct_anterior = _round1(activos_prev / total_sitios * 100.0) if total_sitios else None
    delta_pp = _round1(pct_actual - pct_anterior) if pct_actual is not None and pct_anterior is not None else None

    return {
        "activos": activos,
        "elegibles": total_sitios,
        "anterior_activos": activos_prev,
        "pct_actual": pct_actual,
        "pct_anterior": pct_anterior,
        "delta_pp": delta_pp,
    }


# ---------------------------------------------------------------------------
# KPI 5 - Disponibles (elegibles - activos)
# ---------------------------------------------------------------------------


def compute_kpi5_disponibles(kpi1: dict[str, Any]) -> dict[str, Any]:
    """disponibles = elegibles - activos (mismo patron que TV2 KPI5:
    capacidad_total - ocupado). Disponibilidad es inversa a ocupacion: MENOS
    disponible = mejor utilizacion."""
    elegibles = kpi1["elegibles"]
    disp_actual = elegibles - kpi1["activos"]
    disp_anterior = elegibles - kpi1["anterior_activos"]
    pct_actual = _round1(disp_actual / elegibles * 100.0) if elegibles else None
    pct_anterior = _round1(disp_anterior / elegibles * 100.0) if elegibles else None
    delta_pp = _round1(pct_actual - pct_anterior) if pct_actual is not None and pct_anterior is not None else None

    return {
        "disponibles": disp_actual,
        "elegibles": elegibles,
        "pct_actual": pct_actual,
        "pct_anterior": pct_anterior,
        "delta_pp": delta_pp,
    }


# ---------------------------------------------------------------------------
# Ranking Shoppings Estatico por SitioNegocio (Top 5)
# ---------------------------------------------------------------------------


def compute_ranking_shoppings(engine: MetricsEngine, period: tuple[str, str]) -> list[dict[str, Any]]:
    act_df = engine.query("elementos_con_actividad", group_by=["SitioNegocio"], filters={"CircuitoNegocio": SHOPPINGS_CIRCUITOS, "Medio": MEDIO_ESTATICO}, start_date=period[0], end_date=period[1])
    reg_df = engine.query("elementos_registrados", group_by=["SitioNegocio"], filters={"CircuitoNegocio": SHOPPINGS_CIRCUITOS, "Medio": MEDIO_ESTATICO})

    act_map = dict(zip(act_df["SitioNegocio"], act_df["Value"])) if len(act_df) else {}

    rows = []
    for _, r in reg_df.iterrows():
        sitio = r["SitioNegocio"]
        if sitio is None or (isinstance(sitio, float) and pd.isna(sitio)):
            continue
        elegibles = int(r["Value"])
        activos = int(act_map.get(sitio, 0))
        ocup_pct = _round1(activos / elegibles * 100.0) if elegibles else None
        disponibles = elegibles - activos
        rows.append({
            "sitio": sitio,
            "elegibles": elegibles,
            "activos": activos,
            "disponibles": disponibles,
            "ocup_pct": ocup_pct,
        })

    # Orden (spec TV3 Sec.4): ocupacion desc, activos desc, nombre asc.
    rows.sort(key=lambda x: (
        x["ocup_pct"] is None,
        -(x["ocup_pct"] or 0),
        -x["activos"],
        x["sitio"],
    ))
    return rows[:RANKING_SHOPPINGS_TOP_N]


# ---------------------------------------------------------------------------
# Soportes mas vendidos (FormatoNegocio, Top 3)
# ---------------------------------------------------------------------------


def compute_soportes_top(engine: MetricsEngine, period: tuple[str, str]) -> list[dict[str, Any]]:
    """Top soportes activados en el mes por FormatoNegocio (unica dimension
    de tipo de soporte ya resuelta por Gate3B semantic_model; no se agrega
    una taxonomia nueva). FormatoNegocio para Estatico solo distingue hoy
    FRONTLIGHT/TOTEM/etc. via palabra clave en Descripcion (business_semantics
    formato_negocio.descripcion_keyword_rules); el resto cae en OTRO."""
    df = engine.query("elementos_con_actividad", group_by=["FormatoNegocio"], filters={"CircuitoNegocio": TV3_CIRCUITOS, "Medio": MEDIO_ESTATICO}, start_date=period[0], end_date=period[1])
    df = df[df["Value"] > 0].copy()
    df = df.sort_values(["Value", "FormatoNegocio"], ascending=[False, True])
    return [{"soporte": r["FormatoNegocio"], "activos": int(r["Value"])} for _, r in df.head(SOPORTES_TOP_N).iterrows()]


# ---------------------------------------------------------------------------
# Evolucion mensual (Core Estatico, Ene-Jul)
# ---------------------------------------------------------------------------


def compute_evolution(engine: MetricsEngine, report_year: int, report_month: int) -> dict[str, Any]:
    meses, core = [], []
    for m in range(1, report_month + 1):
        start, end = _period_bounds(report_year, m)
        a = engine.query("elementos_con_actividad", filters={"CircuitoNegocio": TV3_CIRCUITOS, "Medio": MEDIO_ESTATICO}, start_date=start, end_date=end)
        meses.append(MESES_ES_ABR[m - 1])
        core.append(int(a["Value"].iloc[0]) if len(a) else 0)
    return {"meses": meses, "core": core}


# ---------------------------------------------------------------------------
# Reconciliacion con la metrica canonica de TV1 (dia-based, solo validacion,
# no se muestra en las tarjetas -- mismo sistema que TV2 exige conteo, no
# ratio dia-based; ver docstring del modulo).
# ---------------------------------------------------------------------------


def compute_reconciliacion_tv1(engine: MetricsEngine, period: tuple[str, str]) -> dict[str, Any]:
    per_circuito = engine.query(
        "ocupacion_calendario_pct", group_by=["CircuitoNegocio"],
        filters={"CircuitoNegocio": TV3_CIRCUITOS}, start_date=period[0], end_date=period[1],
    )
    eligible = per_circuito[per_circuito["MetricStatus"] != "NO_APLICA"]
    eligible_circuitos = sorted(eligible["CircuitoNegocio"].unique().tolist())
    if eligible.empty:
        return {"ocupacion_dia_based_pct": None, "circuitos_elegibles": [], "status": "NO_APLICA"}
    numerator = int(eligible["Numerator"].sum())
    denominator = int(eligible["Denominator"].sum())
    value = _round1(numerator / denominator * 100.0) if denominator else None
    status = "PARTIAL" if (eligible["MetricStatus"] == "PARTIAL").any() else "OK"
    return {
        "ocupacion_dia_based_pct": value,
        "numerador_dias": numerator,
        "denominador_dias": denominator,
        "circuitos_elegibles": eligible_circuitos,
        "status": status,
    }


# ---------------------------------------------------------------------------
# Insights (Lectura / Punto positivo / A atender) - seleccion programatica
# ---------------------------------------------------------------------------


def compute_insights(
    kpi1: dict[str, Any], kpi3: dict[str, Any], kpi5: dict[str, Any],
    ranking: list[dict[str, Any]],
    report_month_label: str, previous_month_label: str,
) -> dict[str, str]:
    lectura = (
        f"{report_month_label} registra {_fmt_es_int(kpi1['activos'])} de {_fmt_es_int(kpi1['elegibles'])} "
        f"elementos del Core Estático con campaña (<b>{_fmt_es_pct(kpi1['pct_actual'])}%</b>). Shoppings Estático "
        f"concentra toda la actividad del mes y quedan <b>{_fmt_es_int(kpi5['disponibles'])} elementos disponibles</b>."
    )

    if ranking:
        top = ranking[0]
        if len(ranking) > 1:
            second = ranking[1]
            punto_positivo = (
                f"<b>{top['sitio']}</b> lidera la ocupación de Shoppings Estático con "
                f"<b>{_fmt_es_pct(top['ocup_pct'])}%</b> en {report_month_label.lower()}, seguido por "
                f"{second['sitio']} ({_fmt_es_pct(second['ocup_pct'])}%)."
            )
        else:
            punto_positivo = (
                f"<b>{top['sitio']}</b> lidera la ocupación de Shoppings Estático con "
                f"<b>{_fmt_es_pct(top['ocup_pct'])}%</b> en {report_month_label.lower()}."
            )
    else:
        punto_positivo = f"El Core Estático mantiene {_fmt_es_int(kpi1['activos'])} elementos con campaña en {report_month_label.lower()}."

    if kpi3["activos"] == 0 and kpi3["elegibles"] > 0:
        a_atender = (
            f"AA2000 Estático no registra elementos con campaña en {report_month_label.lower()} "
            f"(0 de {_fmt_es_int(kpi3['elegibles'])} elegibles), igual que en {previous_month_label.lower()}."
        )
    elif kpi1["delta_pp"] is not None and kpi1["delta_pp"] < 0:
        a_atender = f"La ocupación por calendario del Core Estático cae <b>{_fmt_es_pct(abs(kpi1['delta_pp']))} pp</b> frente a {previous_month_label.lower()}."
    else:
        a_atender = f"El Core Estático conserva <b>{_fmt_es_int(kpi5['disponibles'])} elementos disponibles</b> por vender."

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

    kpi1 = compute_kpi1_ocupacion_calendario(engine, period, previous)
    kpi2 = compute_kpi2_shoppings_estatico(engine, period, previous)
    kpi3 = compute_kpi3_aa2000_estatico(engine, period, previous)
    pilar = compute_pilar_frontlight(engine, period, previous)
    _assert_family_reconciles_with_core(kpi1, kpi2, kpi3, pilar)
    kpi4 = compute_kpi4_cobertura_shoppings(engine, period, previous)
    kpi5 = compute_kpi5_disponibles(kpi1)

    ranking_shoppings = compute_ranking_shoppings(engine, period)
    soportes_top = compute_soportes_top(engine, period)
    evolution = compute_evolution(engine, REPORT_YEAR, REPORT_MONTH)
    reconciliacion_tv1 = compute_reconciliacion_tv1(engine, period)
    insights = compute_insights(kpi1, kpi3, kpi5, ranking_shoppings, MESES_ES[REPORT_MONTH - 1], MESES_ES[prev_month - 1])

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
            "shoppings_estatico": kpi2,
            "aa2000_estatico": kpi3,
            "cobertura_shoppings": kpi4,
            "disponibles": kpi5,
        },
        "familias": {
            "shoppings": kpi2,
            "aa2000": kpi3,
            "pilar_frontlight": pilar,
        },
        "ranking_shoppings": ranking_shoppings,
        "soportes_top": soportes_top,
        "evolution": evolution,
        "insights": insights,
    }

    return {
        "data": data,
        "sha256": sha_after,
        "universe": {"circuitos": universe["circuitos"]},
        "reconciliacion_tv1": reconciliacion_tv1,
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
    print(json.dumps(result["data"]["kpis"], ensure_ascii=False, indent=2))
    print("RECONCILIACION_TV1:", json.dumps(result["reconciliacion_tv1"], ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
