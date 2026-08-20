"""Capa de datos para el dashboard TV6 - OCU26 (Demanda Comercial).

Se ejecuta DESPUES de scripts/semantic_model.py y scripts/metrics_engine.py
(Gate 3B). Reutiliza export_data.load_pipeline (Gate 4A) y
MetricsEngine._campanas_overlap (mismo patron que build_tv1_dashboard.py):
no reabre el Excel, no reimplementa Gate 1/2/3.

Universo TV6 = "Core Comercial" (CENCOSUD + REMEROS + PANTALLAS_LED +
PILAR_FRONTLIGHT + AA2000) + YPF (CM1 Sec.B4): a diferencia de TV4 (pipeline,
excluye YPF por decision explicita), TV6 mide demanda/marcas/clientes, no
capacidad fisica, y por eso reincorpora YPF. APSA/London Supply nunca entran
(no tienen PortfolioTier CORE ni filas en CAMPANAS para este universo).

TV6 no mide "quien esta en curso hoy" (snapshot, como TV4) sino "quien pauto
en julio 2026": por eso el scope usa solapamiento contra el mes calendario
completo (mismo patron que TV1 clientes_activos/marcas_activas), no un
snapshot puntual al corte.

Programatica (CM1 Sec.B5): el campo canonico CAMPANAS.PROGRAMATICA (Si/No/
vacio) es la unica fuente valida. Auditoria (ver docstring de
compute_programatica) confirmo que, en el acumulado Ene-Jul, PROGRAMATICA=
"Si" nunca aparece sobre una fila con alguna de las 5 agencias identificadas
del periodo (CARAT/OMD/OSA/GROUPM/NEXT MEDIA): aparece solo con Agencia="No"
o "A confirmar". Como el campo no permite mapear con certeza la marca
"programatica" sobre el nombre de una agencia real, esas activaciones NO se
eliminan del analisis ni se reasignan artificialmente: quedan reportadas
como pendientes de imputacion de agencia. Nunca se infiere una agencia
programatica por nombre ni se reclasifica Cliente como Agencia para forzar
el subconjunto.

Enfoque hibrido de dos niveles temporales (ajuste post-entrega): las 5 KPI
cards superiores siguen siendo una foto de JULIO 2026 (sin cambios respecto
de la version anterior). Los paneles inferiores (ranking Marcas/Agencias/
Programatica y matriz "Demanda por circuito") pasan a calcularse sobre el
ACUMULADO 01/01/2026-31/07/2026: un unico mes puede ocultar actores
relevantes que si aparecen en la foto YTD (ej. las agencias OMD/OSA/NEXT
MEDIA no tienen actividad en julio pero si en el acumulado).

Uso:
    python scripts/build_tv6_dashboard.py
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
TEMPLATE_PATH = Path(__file__).resolve().parent / "templates" / "tv6_template.html"
REFERENCE_PATH = REPO_ROOT / "audit_sources" / "TV6_REFERENCE.html.html"
DEFAULT_OUTPUT_HTML = REPO_ROOT / "tv6.html"
DEFAULT_OUTPUT_JSON = REPO_ROOT / "output" / "tv6_data.json"

MESES_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

# Corte operativo TV6 para las KPI cards: mismo mes de reporte vigente en
# TV1-5 (Julio 2026). No cambia con el enfoque hibrido.
REPORT_YEAR = 2026
REPORT_MONTH = 7

# Acumulado historico para los paneles inferiores (ranking + matriz): mismo
# universo TV6, ventana Ene-Jul 2026 en vez de solo julio.
HIST_START_ISO = "2026-01-01"
HIST_END_ISO = "2026-07-31"
HIST_LABEL = "Ene–Jul 2026"

# Universo TV6 (CM1 Sec.B4): Core Comercial completo + YPF. APSA/London
# Supply quedan afuera porque no tienen PortfolioTier CORE (business_
# semantics.json: LEGACY/COMPLEMENTARIO) y de hecho no tienen filas en
# CAMPANAS dentro de OPERATIVO_GENERAL.
TV6_CIRCUITOS = ["CENCOSUD", "REMEROS", "PANTALLAS_LED", "PILAR_FRONTLIGHT", "AA2000", "YPF"]
MEDIO_DIGITAL = "Digital"

# Ciclo visual del ranking principal (CM1 Sec.B5): reemplaza Clientes por
# Marcas -> Agencias -> Programatica.
CICLO = ["marcas", "agencias", "programatica"]
RANK_TOP_N = 6
CONCENTRACION_TOP_N = 5
MATRIZ_COLUMNAS = [
    "Pantallas LED", "Shoppings Digital", "Shoppings Estático",
    "AA2000 / Pilar Frontlight", "YPF",
]

# Valores del campo Agencia que NO representan un nombre de agencia real
# (CM1 Sec.B5/B11: no convertir vacio/"A confirmar" en agencia identificada).
AGENCIA_VALORES_NO_IDENTIFICADOS = {"No", "A confirmar"}
# Valores del campo Cliente que son placeholders, no un cliente real
# (auditoria: "A CONFIRMAR" = pendiente, "AGENCIA" = vendido via agencia sin
# cliente directo identificado).
CLIENTE_PLACEHOLDERS = {"A CONFIRMAR", "AGENCIA"}


class BuildError(Exception):
    """Error bloqueante al construir el dashboard TV6."""


def _period_bounds(year: int, month: int) -> tuple[str, str]:
    start = pd.Timestamp(year=year, month=month, day=1)
    end = start + pd.offsets.MonthEnd(0)
    return str(start.date()), str(end.date())


def _fmt_es_int(n: int) -> str:
    return f"{n:,}".replace(",", ".")


def _fmt_es_pct(v: float) -> str:
    return f"{v:.1f}".replace(".", ",")


def _join_es(names) -> str:
    names = list(names)
    if not names:
        return "sin agencia identificada"
    if len(names) == 1:
        return names[0]
    return ", ".join(names[:-1]) + " y " + names[-1]


# ---------------------------------------------------------------------------
# Universo TV6 y scope de campanas (Core Comercial + YPF, julio 2026)
# ---------------------------------------------------------------------------


def build_tv6_universe(semantic_result: dict[str, Any]) -> dict[str, Any]:
    maestro = semantic_result["maestro"]
    config = semantic_result["config"]

    op = filter_universe(maestro, "OPERATIVO_GENERAL", config)
    tv6_maestro = op[op["CircuitoNegocio"].isin(TV6_CIRCUITOS)].copy()
    if tv6_maestro.empty:
        raise BuildError("Universo TV6 vacio: revisar business_semantics.json / TV6_CIRCUITOS")

    return {
        "circuitos": sorted(tv6_maestro["CircuitoNegocio"].unique().tolist()),
        "elementos": int(tv6_maestro["ElementoID"].nunique()),
        "element_ids": tv6_maestro["ElementoID"].tolist(),
    }


def _familia(row: pd.Series) -> str:
    """Grupo comercial ejecutivo (CM1 Sec.B8): YPF conserva columna propia
    (opcion explicitamente habilitada) en vez de mezclarse con Shoppings,
    para no diluir su semantica de circuito propio dentro de la matriz."""
    circuito = row["CircuitoNegocio"]
    if circuito == "YPF":
        return "YPF"
    if circuito == "PANTALLAS_LED":
        return "Pantallas LED"
    if circuito in ("CENCOSUD", "REMEROS"):
        return "Shoppings Digital" if row["Medio"] == MEDIO_DIGITAL else "Shoppings Estático"
    return "AA2000 / Pilar Frontlight"


def build_tv6_scope(engine: MetricsEngine, element_ids: list[Any], start: str, end: str) -> pd.DataFrame:
    """Activaciones de julio 2026 en el universo TV6: solapamiento de mes
    calendario completo (no snapshot puntual, a diferencia de TV4 pipeline),
    mismo patron que TV1 clientes_activos/marcas_activas."""
    scope = engine._campanas_overlap(element_ids, start, end)
    if scope.empty:
        raise BuildError("Scope TV6 vacio para el periodo de reporte: revisar corte/universo")
    scope = scope.copy()
    scope["_familia"] = scope.apply(_familia, axis=1)
    return scope


# ---------------------------------------------------------------------------
# Clientes directos (CM1 Sec.B7/B11): identificado = Agencia=="No" (explicito,
# sin agencia) y Cliente real (no vacio, no placeholder). Vía agencia y A
# confirmar quedan en buckets separados, nunca mezclados con "identificados".
# ---------------------------------------------------------------------------


def compute_clientes_directos(scope: pd.DataFrame) -> dict[str, Any]:
    directo = scope[scope["Agencia"] == "No"]
    identificado = directo[directo["Cliente"].notna() & ~directo["Cliente"].isin(CLIENTE_PLACEHOLDERS)]

    ranking = (
        identificado.groupby("Cliente").size().reset_index(name="activaciones")
        if not identificado.empty else pd.DataFrame(columns=["Cliente", "activaciones"])
    )
    ranking = ranking.sort_values(["activaciones", "Cliente"], ascending=[False, True])

    top = None
    if not ranking.empty:
        row = ranking.iloc[0]
        top = {"nombre": row["Cliente"], "activaciones": int(row["activaciones"])}

    return {
        "activos": int(identificado["Cliente"].nunique()),
        "activaciones_totales": int(len(identificado)),
        "top_identificado": top,
        "ranking": [
            {"nombre": r["Cliente"], "activaciones": int(r["activaciones"])}
            for _, r in ranking.iterrows()
        ],
    }


# ---------------------------------------------------------------------------
# Marcas (CM1 Sec.B4/B11): Core + YPF ya combinados en `scope`; distinct()
# sobre Marca aplica la deduplicacion pedida en un solo paso.
# ---------------------------------------------------------------------------


def compute_marcas(scope: pd.DataFrame) -> dict[str, Any]:
    marcas = scope[scope["Marca"].notna()]
    counts = marcas.groupby("Marca").size().reset_index(name="activaciones")
    counts = counts.sort_values(["activaciones", "Marca"], ascending=[False, True])

    return {
        "activas": int(marcas["Marca"].nunique()),
        "ranking_top": [
            {"nombre": r["Marca"], "activaciones": int(r["activaciones"])}
            for _, r in counts.head(RANK_TOP_N).iterrows()
        ],
        "ranking_full": [
            {"nombre": r["Marca"], "activaciones": int(r["activaciones"])}
            for _, r in counts.iterrows()
        ],
    }


# ---------------------------------------------------------------------------
# Agencias (CM1 Sec.B7/B11): identificada = nombre real, excluye "No" (sin
# agencia) y "A confirmar" (pendiente). Ambos buckets se reportan aparte.
# ---------------------------------------------------------------------------


def compute_agencias(scope: pd.DataFrame) -> dict[str, Any]:
    identificada_mask = scope["Agencia"].notna() & ~scope["Agencia"].isin(AGENCIA_VALORES_NO_IDENTIFICADOS)
    identificadas = scope[identificada_mask]

    counts = identificadas.groupby("Agencia").size().reset_index(name="activaciones")
    counts = counts.sort_values(["activaciones", "Agencia"], ascending=[False, True])

    return {
        "activas": int(identificadas["Agencia"].nunique()),
        "activaciones_totales": int(len(identificadas)),
        "pendientes_a_confirmar": int((scope["Agencia"] == "A confirmar").sum()),
        "ranking_top": [
            {"nombre": r["Agencia"], "activaciones": int(r["activaciones"])}
            for _, r in counts.head(RANK_TOP_N).iterrows()
        ],
        "nombres_identificados": set(identificadas["Agencia"].unique()),
    }


# ---------------------------------------------------------------------------
# Programatica (CM1 Sec.B5): subconjunto de Agencias con PROGRAMATICA=="Si".
# Nunca infiere por nombre ni convierte vacio/"A confirmar" en clasificacion
# positiva. Auditoria Ene-Jul confirmo que PROGRAMATICA="Si" nunca coincide
# con ninguna de las 5 agencias identificadas del periodo (CARAT/OMD/OSA/
# GROUPM/NEXT MEDIA) -- aparece unicamente sobre filas con Agencia="No" o
# "A confirmar" (ej. Cliente=TAGGIFY/LATIN AD/BEEYOND). Esas activaciones NO
# se eliminan del analisis general ni se reasignan artificialmente a una
# agencia: quedan expuestas como "pendientes_sin_agencia" para que el panel
# las muestre debajo del ranking (vacio en este caso) sin inventar clasificacion.
# ---------------------------------------------------------------------------


def compute_programatica(scope: pd.DataFrame, agencias_identificadas: set[Any]) -> dict[str, Any]:
    flag_si = scope[scope["PROGRAMATICA"] == "Si"]
    prog_identificadas = flag_si[flag_si["Agencia"].isin(agencias_identificadas)]

    counts = prog_identificadas.groupby("Agencia").size().reset_index(name="activaciones")
    counts = counts.sort_values(["activaciones", "Agencia"], ascending=[False, True])
    ranking = [
        {"nombre": r["Agencia"], "activaciones": int(r["activaciones"])}
        for _, r in counts.iterrows()
    ]
    pendientes_sin_agencia = int(len(flag_si) - len(prog_identificadas))

    return {
        "estado": "OK" if ranking else "A_VALIDAR",
        "ranking_top": ranking[:RANK_TOP_N],
        "activaciones_flag_si_total": int(len(flag_si)),
        "pendientes_sin_agencia": pendientes_sin_agencia,
        "nota": (
            f"{_fmt_es_int(pendientes_sin_agencia)} activaciones programáticas pendientes de "
            f"imputación de agencia."
            if pendientes_sin_agencia else ""
        ),
    }


# ---------------------------------------------------------------------------
# Concentracion Top 5 (CM1 Sec.B11.F): definicion explicita requerida por
# spec -- Top 5 MARCAS (unidad del ranking principal del ciclo) sobre el
# total de activaciones del universo TV6/julio.
# ---------------------------------------------------------------------------


def compute_concentracion_top5(scope: pd.DataFrame, marcas_ranking_full: list[dict[str, Any]]) -> dict[str, Any]:
    total = len(scope)
    top5 = marcas_ranking_full[:CONCENTRACION_TOP_N]
    numerador = sum(m["activaciones"] for m in top5)
    pct = round(numerador / total * 100.0, 1) if total else None
    return {
        "pct": pct,
        "numerador": numerador,
        "denominador": total,
        "unidad": "activaciones",
        "base": "Top 5 marcas",
        "marcas": [m["nombre"] for m in top5],
    }


# ---------------------------------------------------------------------------
# Matriz "Demanda por circuito" (CM1 Sec.B8): Top marcas (misma unidad que
# el ranking principal) x grupo comercial, activaciones del mes. YPF con
# columna propia.
# ---------------------------------------------------------------------------


def compute_matriz(scope: pd.DataFrame, top_marcas: list[str]) -> dict[str, Any]:
    sub = scope[scope["Marca"].isin(top_marcas)]
    pivot = sub.pivot_table(index="Marca", columns="_familia", values="ElementoID", aggfunc="count", fill_value=0)
    for col in MATRIZ_COLUMNAS:
        if col not in pivot.columns:
            pivot[col] = 0
    pivot = pivot.reindex(top_marcas)[MATRIZ_COLUMNAS].fillna(0)

    filas = [
        {"nombre": nombre, "valores": [int(pivot.loc[nombre, c]) for c in MATRIZ_COLUMNAS]}
        for nombre in top_marcas
    ]
    return {"columnas": MATRIZ_COLUMNAS, "filas": filas}


# ---------------------------------------------------------------------------
# Pendientes de imputacion (CM1 Sec.B7): buckets separados, nunca mezclados
# con identificados/directos.
# ---------------------------------------------------------------------------


def compute_pendientes(scope: pd.DataFrame, agencias: dict[str, Any]) -> dict[str, Any]:
    return {
        "clientes_via_agencia": agencias["activaciones_totales"],
        "agencias_a_confirmar": agencias["pendientes_a_confirmar"],
        "clientes_pendientes_imputacion": int(
            (scope["Agencia"].isna() & scope["Cliente"].notna()).sum()
        ),
    }


# ---------------------------------------------------------------------------
# Insights (Lectura / Punto positivo / A atender): combinan la foto de julio
# (KPI cards) con el acumulado Ene-Jul (paneles), mismo patron de 3 columnas
# que TV1-4. No inventa causas: cada frase cita un hecho ya calculado.
# ---------------------------------------------------------------------------


def compute_insights(
    kpis_julio: dict[str, Any], agencias_julio: dict[str, Any],
    marcas_hist: dict[str, Any], agencias_hist: dict[str, Any], programatica_hist: dict[str, Any],
    concentracion_hist: dict[str, Any], familia_hist: dict[str, int], pendientes_julio: dict[str, Any],
) -> dict[str, str]:
    top_marcas_txt = ", ".join(m["nombre"] for m in marcas_hist["ranking_top"][:3])
    top_agencias_txt = _join_es(a["nombre"] for a in agencias_hist["ranking_top"][:3])
    familia_top = max(familia_hist.items(), key=lambda kv: kv[1]) if familia_hist else None
    total_hist = sum(familia_hist.values()) or 1
    familia_txt = ""
    if familia_top and familia_top[1] > 0:
        share = round(familia_top[1] / total_hist * 100.0, 1)
        familia_txt = f", concentrada principalmente en <b>{familia_top[0]}</b> ({_fmt_es_pct(share)}%)"

    lectura = (
        f"Julio registra <b>{_fmt_es_int(kpis_julio['marcas_activas'])} marcas</b> y "
        f"<b>{_fmt_es_int(kpis_julio['agencias_activas'])} agencias</b> activas. En el acumulado "
        f"{HIST_LABEL}, las principales marcas son <b>{top_marcas_txt}</b> y las agencias "
        f"identificadas son <b>{top_agencias_txt}</b>{familia_txt}."
    )

    # PUNTO POSITIVO: prioridad -- mayor profundidad de agencias identificadas
    # en el acumulado vs. julio; si no aplica, diversidad de marcas del acumulado.
    if agencias_hist["activas"] > agencias_julio["activas"]:
        agencias_nombres = ", ".join(a["nombre"] for a in agencias_hist["ranking_top"])
        punto_positivo = (
            f"El acumulado {HIST_LABEL} identifica <b>{_fmt_es_int(agencias_hist['activas'])} agencias</b> "
            f"({agencias_nombres}) frente a las {_fmt_es_int(agencias_julio['activas'])} vistas en julio: "
            f"mayor profundidad de imputación de canal en el año."
        )
    else:
        punto_positivo = (
            f"El acumulado {HIST_LABEL} registra <b>{_fmt_es_int(marcas_hist['activas'])} marcas distintas</b> "
            f"activas, mostrando una base de demanda diversificada."
        )

    # A ATENDER: prioridad 1) activaciones programáticas sin agencia imputada;
    # 2) A confirmar/pendientes; 3) concentración elevada; 4) dependencia YPF.
    if programatica_hist["pendientes_sin_agencia"] > 0:
        a_atender = (
            f"<b>{_fmt_es_int(programatica_hist['pendientes_sin_agencia'])} activaciones programáticas</b> "
            f"del acumulado {HIST_LABEL} quedan sin agencia imputada; el campo PROGRAMATICA no permite "
            f"asociarlas a una agencia identificada."
        )
    elif pendientes_julio["agencias_a_confirmar"] > 0:
        a_atender = (
            f"Quedan <b>{_fmt_es_int(pendientes_julio['agencias_a_confirmar'])} activaciones</b> de julio "
            f"con agencia A confirmar, pendientes de imputación."
        )
    elif concentracion_hist["pct"] is not None and concentracion_hist["pct"] >= 70:
        a_atender = (
            f"El Top 5 de marcas concentra <b>{_fmt_es_pct(concentracion_hist['pct'])}%</b> de las "
            f"activaciones del acumulado {HIST_LABEL}: la demanda queda poco diversificada."
        )
    elif familia_top and familia_top[0] == "YPF" and familia_top[1] / total_hist >= 0.7:
        a_atender = (
            f"El acumulado {HIST_LABEL} depende fuertemente de <b>YPF</b> "
            f"({_fmt_es_pct(round(familia_top[1] / total_hist * 100.0, 1))}% de las activaciones)."
        )
    else:
        a_atender = (
            f"El acumulado {HIST_LABEL} suma {_fmt_es_int(sum(familia_hist.values()))} activaciones "
            f"sobre el universo Core Comercial + YPF."
        )

    return {"lectura": lectura, "punto_positivo": punto_positivo, "a_atender": a_atender}


# ---------------------------------------------------------------------------
# Logo (reutilizado byte-a-byte, igual patron que TV1-4)
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


def build_tv6_data(path: str | Path = vi.DEFAULT_INPUT_PATH) -> dict[str, Any]:
    path = Path(path)
    sha_before = vi.calculate_sha256(path)

    _transform_result, semantic_result, engine = load_pipeline(path)
    universe = build_tv6_universe(semantic_result)

    # Nivel 1: foto JULIO 2026 -- unica base de las 5 KPI cards (sin cambios
    # respecto de la version anterior).
    start_jul, end_jul = _period_bounds(REPORT_YEAR, REPORT_MONTH)
    scope_julio = build_tv6_scope(engine, universe["element_ids"], start_jul, end_jul)

    clientes = compute_clientes_directos(scope_julio)
    marcas_julio = compute_marcas(scope_julio)
    agencias_julio = compute_agencias(scope_julio)
    concentracion_julio = compute_concentracion_top5(scope_julio, marcas_julio["ranking_full"])
    pendientes_julio = compute_pendientes(scope_julio, agencias_julio)
    familia_julio = {k: int(v) for k, v in scope_julio["_familia"].value_counts().to_dict().items()}

    # Nivel 2: ACUMULADO Ene-Jul -- base de los paneles inferiores (ranking
    # ciclado + matriz "Demanda por circuito"), para no ocultar actores que
    # un unico mes no muestra (CM1 ajuste hibrido).
    scope_hist = build_tv6_scope(engine, universe["element_ids"], HIST_START_ISO, HIST_END_ISO)

    marcas_hist = compute_marcas(scope_hist)
    agencias_hist = compute_agencias(scope_hist)
    programatica_hist = compute_programatica(scope_hist, agencias_hist["nombres_identificados"])
    concentracion_hist = compute_concentracion_top5(scope_hist, marcas_hist["ranking_full"])
    top_marcas_hist_nombres = [m["nombre"] for m in marcas_hist["ranking_top"]]
    matriz = compute_matriz(scope_hist, top_marcas_hist_nombres)
    familia_hist = {k: int(v) for k, v in scope_hist["_familia"].value_counts().to_dict().items()}

    insights = compute_insights(
        {"marcas_activas": marcas_julio["activas"], "agencias_activas": agencias_julio["activas"]},
        agencias_julio, marcas_hist, agencias_hist, programatica_hist,
        concentracion_hist, familia_hist, pendientes_julio,
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
            "periodo_inicio_iso": start_jul,
            "periodo_fin_iso": end_jul,
            "hist_inicio_iso": HIST_START_ISO,
            "hist_fin_iso": HIST_END_ISO,
            "hist_label": HIST_LABEL,
            "fuente": "OCU26 · Base maestra + base campañas",
        },
        "universo": {
            "circuitos": universe["circuitos"],
            "elementos": universe["elementos"],
            "julio": {
                "campanas_unicas": int(scope_julio["IDCampaña"].dropna().nunique()),
                "activaciones_totales": int(len(scope_julio)),
                "activaciones_por_familia": familia_julio,
            },
            "hist": {
                "periodo_label": HIST_LABEL,
                "campanas_unicas": int(scope_hist["IDCampaña"].dropna().nunique()),
                "activaciones_totales": int(len(scope_hist)),
                "activaciones_por_familia": familia_hist,
            },
        },
        "kpis": {
            "clientes_directos_activos": clientes["activos"],
            "marcas_activas": marcas_julio["activas"],
            "agencias_activas": agencias_julio["activas"],
            "cliente_top_identificado": clientes["top_identificado"],
            "concentracion_top5": concentracion_julio,
        },
        "ranking": {
            "ciclo": CICLO,
            "periodo_label": HIST_LABEL,
            "marcas": marcas_hist["ranking_top"],
            "agencias": agencias_hist["ranking_top"],
            "programatica": programatica_hist,
        },
        "matriz": {**matriz, "periodo_label": HIST_LABEL},
        "pendientes": pendientes_julio,
        "insights": insights,
    }

    return {"data": data, "sha256": sha_after, "universe": {"circuitos": universe["circuitos"], "elementos": universe["elementos"]}}


def render_html(data: dict[str, Any]) -> str:
    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    logo_tag = extract_logo_img_tag(REFERENCE_PATH)
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    html = template.replace("{{LOGO_IMG_TAG}}", logo_tag)
    html = html.replace("{{TV6_DATA_JSON}}", payload)
    return html


def build_and_write(
    path: str | Path = vi.DEFAULT_INPUT_PATH,
    output_html: str | Path = DEFAULT_OUTPUT_HTML,
    output_json: str | Path = DEFAULT_OUTPUT_JSON,
) -> dict[str, Any]:
    result = build_tv6_data(path)
    html = render_html(result["data"])

    output_html = Path(output_html)
    output_html.write_text(html, encoding="utf-8")

    output_json = Path(output_json)
    output_json.parent.mkdir(parents=True, exist_ok=True)
    with open(output_json, "w", encoding="utf-8") as fh:
        json.dump(result["data"], fh, ensure_ascii=False, indent=2)

    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Construye tv6.html (dashboard TV6 OCU26) con datos reales.")
    parser.add_argument("--file", default=str(vi.DEFAULT_INPUT_PATH), help="Ruta al archivo .xlsx a leer")
    parser.add_argument("--output-html", default=str(DEFAULT_OUTPUT_HTML), help="Ruta del HTML productivo generado")
    parser.add_argument("--output-json", default=str(DEFAULT_OUTPUT_JSON), help="Ruta del snapshot JSON generado")
    args = parser.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    try:
        result = build_and_write(args.file, args.output_html, args.output_json)
    except BuildError as exc:
        print("TV6_BUILD_ERROR:", exc)
        return 1

    print("TV6_BUILD_OK")
    print(json.dumps(result["data"]["meta"], ensure_ascii=False, indent=2))
    print(json.dumps(result["data"]["kpis"], ensure_ascii=False, indent=2))
    print("UNIVERSE:", json.dumps(result["universe"], ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
