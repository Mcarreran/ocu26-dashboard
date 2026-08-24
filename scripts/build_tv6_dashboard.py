"""Capa de datos para el dashboard TV6 - OCU26 (Demanda de marcas y
expansion por circuito).

Se ejecuta DESPUES de scripts/semantic_model.py y scripts/metrics_engine.py
(Gate 3B). Reutiliza export_data.load_pipeline (Gate 4A) y
MetricsEngine._campanas_overlap (mismo patron que build_tv1_dashboard.py):
no reabre el Excel, no reimplementa Gate 1/2/3.

Reescrito 2026-08-23 (decision de negocio): Agencias/Programatica/Clientes
directos/Exclusividades/Canal de ingreso salen del alcance de TV6 -- la
fuente actual no tiene cobertura suficiente para mostrarlos de forma
confiable (ver historial de auditoria en git). Quedan pospuestos para una
futura mejora de la base; esos campos de CAMPANAS no se tocan ni se
eliminan, simplemente TV6 ya no los audita, calcula ni muestra. El unico
resabio que se conserva es la lista de plataformas/intermediarios conocidos
(PROGRAMATICA_PLATAFORMAS_EXCLUIR): sigue haciendo falta para que "TAGGIFY"/
"BEEYOND"/"LATIN AD"/"GLOBAL" no se cuenten como marca cuando aparecen en el
campo Marca por un problema de carga de origen (regla de negocio de Marca,
no un calculo de canal).

Nuevo objetivo: "Demanda de marcas y expansion por circuito". Responde
quien pauta, cuales marcas regresan (recurrencia 2026) y en cuantos
circuitos estan presentes -- informacion que no aparece en TV1-TV5.

Universo TV6 = "Core Comercial" (CENCOSUD + REMEROS + PANTALLAS_LED +
PILAR_FRONTLIGHT + AA2000) + YPF (CM1 Sec.B4): a diferencia de TV4 (pipeline,
excluye YPF por decision explicita), TV6 mide demanda de marcas, no
capacidad fisica, y por eso reincorpora YPF. APSA/London Supply nunca entran
(no tienen PortfolioTier CORE ni filas en CAMPANAS para este universo).

TV6 no mide "quien esta en curso hoy" (snapshot, como TV4) sino "quien pauto
en julio 2026": por eso el scope usa solapamiento contra el mes calendario
completo (mismo patron que TV1 clientes_activos/marcas_activas), no un
snapshot puntual al corte.

Grano (auditoria 2026-08-23): una "activacion" es un par distinto
ElementoID x IDCampaña. CAMPANAS puede tener mas de una fila (CargaID) para
el mismo par cuando el mismo placement se recargo con una fecha de fin
corregida -- nunca representan dos activaciones comerciales distintas.
build_tv6_scope deduplica por (ElementoID,IDCampaña) antes de calcular
cualquier metrica: "no contar filas crudas como activaciones". "Campaña
unica" es IDCampaña distinto dentro del scope.

Primera aparicion / recurrencia (Sec.4 Tarjetas 2-3): se resuelve mirando,
mes a mes, el conjunto de marcas validas activas de enero a julio 2026 (7
scopes mensuales independientes, mismo universo y misma regla de
solapamiento que el scope de julio). El primer mes con actividad de cada
marca es el minimo de esos meses. Una marca activa en julio es "primera
aparicion" si ese minimo es julio, o "recurrente" si es cualquier mes
anterior (no exige actividad en junio especificamente: una marca activa en
marzo, ausente en abril-junio y de vuelta en julio sigue siendo recurrente).
"Primera aparicion" nunca se presenta como "cliente nuevo" absoluto: solo se
observa la ventana Ene-Jul 2026, la marca puede haber pauteado antes.

Multicircuito / un solo circuito (Sec.4 Tarjetas 4-5): cuenta grupos
comerciales canonicos (_familia) distintos con >=1 activacion valida de la
marca en julio.

Uso:
    python scripts/build_tv6_dashboard.py
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
import unicodedata
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

IDCAMPANA_COL = "IDCampaña"

MESES_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

# Corte operativo TV6: periodo principal Julio 2026, ventana historica
# Ene-Jul 2026 para determinar primera aparicion/recurrencia.
REPORT_YEAR = 2026
REPORT_MONTH = 7
HIST_START_ISO = "2026-01-01"
HIST_END_ISO = "2026-07-31"
HIST_LABEL = "Ene–Jul 2026"

# Universo TV6 (CM1 Sec.B4): Core Comercial completo + YPF. APSA/London
# Supply quedan afuera porque no tienen PortfolioTier CORE (business_
# semantics.json: LEGACY/COMPLEMENTARIO) y de hecho no tienen filas en
# CAMPANAS dentro de OPERATIVO_GENERAL.
TV6_CIRCUITOS = ["CENCOSUD", "REMEROS", "PANTALLAS_LED", "PILAR_FRONTLIGHT", "AA2000", "YPF"]
MEDIO_DIGITAL = "Digital"

RANK_TOP_N = 6
MATRIZ_COLUMNAS = [
    "Pantallas LED", "Shoppings Digital", "Shoppings Estático",
    "AA2000 / Pilar Frontlight", "YPF",
]

# Placeholders del campo Marca (CM1 Sec.3): nunca cuentan como marca.
MARCA_PLACEHOLDERS = {"A CONFIRMAR", "S/D", "SIN DATO", "N/A", ""}

# Plataformas/intermediarios conocidos (auditoria 2026-08-23): si aparecen
# en el campo Marca por un problema de carga de origen, esa activacion
# queda "sin marca valida", nunca se cuenta como marca. TV6 ya no clasifica
# canal/programatica -- este set solo protege la regla de negocio de Marca.
PROGRAMATICA_PLATAFORMAS_EXCLUIR = {"TAGGIFY", "BEEYOND", "LATIN AD", "LATINAD", "GLOBAL"}


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


def _norm_key(value: Any) -> str | None:
    """Clave de comparacion normalizada (CM1 Sec.3): trim, colapso de
    espacios repetidos, sin acentos, mayusculas. Nunca decide que mostrar
    (eso conserva el valor original de la fuente), solo si dos valores son
    la misma entidad o coinciden con un termino canonico/placeholder."""
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    if value is pd.NA:
        return None
    s = str(value).strip()
    if not s:
        return None
    s = re.sub(r"\s+", " ", s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.upper()


# ---------------------------------------------------------------------------
# Universo TV6 y scope de campanas (Core Comercial + YPF)
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
    """Grupo comercial ejecutivo (CM1 Sec.2): YPF conserva columna propia en
    vez de mezclarse con Shoppings, para no diluir su semantica de circuito
    propio dentro de la matriz."""
    circuito = row["CircuitoNegocio"]
    if circuito == "YPF":
        return "YPF"
    if circuito == "PANTALLAS_LED":
        return "Pantallas LED"
    if circuito in ("CENCOSUD", "REMEROS"):
        return "Shoppings Digital" if row["Medio"] == MEDIO_DIGITAL else "Shoppings Estático"
    return "AA2000 / Pilar Frontlight"


def build_tv6_scope(
    engine: MetricsEngine, element_ids: list[Any], start: str, end: str, allow_empty: bool = False,
) -> pd.DataFrame:
    """Activaciones del periodo en el universo TV6: solapamiento de mes
    calendario completo, deduplicado a grano (ElementoID, IDCampaña) --
    ver docstring de modulo. "keep=first" es determinista porque CAMPANAS
    ya viene ordenado por CargaID (orden de carga).

    allow_empty=False (default, periodo principal Julio) falla explicito si
    el universo no tiene ninguna activacion: es una señal de corte/universo
    mal configurado. allow_empty=True se usa solo para los 7 scopes
    mensuales de compute_monthly_brand_presence: un mes puntual sin
    actividad es un dato valido (no un error de configuracion)."""
    scope = engine._campanas_overlap(element_ids, start, end)
    if scope.empty:
        if allow_empty:
            return scope.assign(_familia=pd.Series(dtype="object"))
        raise BuildError("Scope TV6 vacio para el periodo de reporte: revisar corte/universo")
    scope = scope.drop_duplicates(subset=["ElementoID", IDCAMPANA_COL], keep="first").copy()
    scope["_familia"] = scope.apply(_familia, axis=1)
    return scope


def _trim_ws(value: Any) -> str | None:
    """Trim + colapso de espacios repetidos, preservando mayusculas/
    minusculas y acentos originales (para mostrar). Ej.: 'PARQUE DE LA
    COSTA ' (espacio final) y 'PARQUE DE LA COSTA' ya quedan identicos."""
    if value is None or (isinstance(value, float) and pd.isna(value)) or value is pd.NA:
        return None
    s = re.sub(r"\s+", " ", str(value).strip())
    return s or None


def classify_marca(scope: pd.DataFrame) -> pd.DataFrame:
    """Agrega _marca_k (clave normalizada para detectar placeholders/
    plataformas), _marca_valida (CM1 Sec.3: excluye vacios, placeholders y
    nombres de plataformas/intermediarios conocidos -- nunca usa Agencia/
    Cliente como fallback) y _marca_canon (identidad canonica de marca:
    misma clave normalizada = misma marca, ej. 'PARQUE DE LA COSTA' y
    'PARQUE DE LA COSTA ' con espacio final cuentan como una sola marca).
    El nombre de display de cada _marca_k es la variante mas frecuente ya
    trimeada (_trim_ws), nunca la clave normalizada en mayusculas: preserva
    el nombre real de la fuente."""
    scope = scope.copy()
    scope["_marca_k"] = scope["Marca"].apply(_norm_key)
    placeholder = scope["_marca_k"].isna() | scope["_marca_k"].isin(MARCA_PLACEHOLDERS)
    es_plataforma = scope["_marca_k"].isin(PROGRAMATICA_PLATAFORMAS_EXCLUIR)
    scope["_marca_valida"] = ~placeholder & ~es_plataforma

    scope["_marca_trim"] = scope["Marca"].apply(_trim_ws)
    validas = scope[scope["_marca_valida"]]
    if not validas.empty:
        display_por_clave = (
            validas.groupby("_marca_k")["_marca_trim"]
            .agg(lambda s: s.value_counts().index[0])
        )
        scope["_marca_canon"] = scope["_marca_k"].map(display_por_clave)
    else:
        scope["_marca_canon"] = None
    return scope


# ---------------------------------------------------------------------------
# Marcas activas de julio (Tarjeta 1) + control de calidad de contradiccion
# ---------------------------------------------------------------------------


def compute_marcas_julio(scope: pd.DataFrame) -> dict[str, Any]:
    validas = scope[scope["_marca_valida"]]

    contradiccion = (
        validas.groupby(IDCAMPANA_COL)["_marca_canon"].nunique() if not validas.empty else pd.Series(dtype="int64")
    )
    campanas_marca_contradictoria = int((contradiccion > 1).sum())

    return {
        "activas": int(validas["_marca_canon"].nunique()),
        "campanas_unicas_total": int(scope[IDCAMPANA_COL].dropna().nunique()),
        "campanas_unicas_marca_valida": int(validas[IDCAMPANA_COL].dropna().nunique()),
        "activaciones_totales": int(len(scope)),
        "activaciones_sin_marca_valida": int((~scope["_marca_valida"]).sum()),
        "campanas_marca_contradictoria": campanas_marca_contradictoria,
    }


# ---------------------------------------------------------------------------
# Primera aparicion / recurrencia (Tarjetas 2-3): 7 scopes mensuales
# independientes Ene-Jul, mismo universo, mismo grano.
# ---------------------------------------------------------------------------


def compute_monthly_brand_presence(engine: MetricsEngine, element_ids: list[Any]) -> dict[int, set[str]]:
    presencia: dict[int, set[str]] = {}
    for month in range(1, REPORT_MONTH + 1):
        start, end = _period_bounds(REPORT_YEAR, month)
        scope_mes = classify_marca(build_tv6_scope(engine, element_ids, start, end, allow_empty=True))
        presencia[month] = set(scope_mes.loc[scope_mes["_marca_valida"], "_marca_canon"].unique())
    return presencia


def compute_recurrencia(marcas_julio: set[str], presencia_mensual: dict[int, set[str]]) -> dict[str, Any]:
    todas_las_marcas = set().union(*presencia_mensual.values()) if presencia_mensual else set()
    primer_mes: dict[str, int] = {
        marca: min(m for m in presencia_mensual if marca in presencia_mensual[m])
        for marca in todas_las_marcas
    }

    primera_aparicion = sorted(m for m in marcas_julio if primer_mes.get(m) == REPORT_MONTH)
    recurrentes = sorted(m for m in marcas_julio if primer_mes.get(m) is not None and primer_mes[m] < REPORT_MONTH)
    sin_historial = sorted(marcas_julio - set(primera_aparicion) - set(recurrentes))

    total = len(marcas_julio)

    def _pct(n: int) -> float | None:
        return round(n / total * 100.0, 1) if total else None

    return {
        "primer_mes_por_marca": {m: primer_mes[m] for m in sorted(marcas_julio) if m in primer_mes},
        "primera_aparicion": {
            "count": len(primera_aparicion), "pct": _pct(len(primera_aparicion)), "marcas": primera_aparicion,
        },
        "recurrentes": {
            "count": len(recurrentes), "pct": _pct(len(recurrentes)), "marcas": recurrentes,
        },
        "sin_historial_comparable": {
            "count": len(sin_historial), "pct": _pct(len(sin_historial)), "marcas": sin_historial,
        },
        "reconciliacion": {
            "primera_mas_recurrentes_mas_sin_historial": (
                len(primera_aparicion) + len(recurrentes) + len(sin_historial)
            ),
            "marcas_activas": total,
            "ok": len(primera_aparicion) + len(recurrentes) + len(sin_historial) == total,
        },
    }


# ---------------------------------------------------------------------------
# Multicircuito / un solo circuito (Tarjetas 4-5)
# ---------------------------------------------------------------------------


def compute_circuitos_por_marca(scope_julio: pd.DataFrame, marcas_julio: set[str]) -> dict[str, Any]:
    validas = scope_julio[scope_julio["_marca_valida"]]
    circuitos_count = validas.groupby("_marca_canon")["_familia"].nunique()

    multicircuito = sorted(circuitos_count[circuitos_count >= 2].index.tolist())
    un_solo_circuito = sorted(circuitos_count[circuitos_count == 1].index.tolist())
    sin_circuito_valido = sorted(marcas_julio - set(multicircuito) - set(un_solo_circuito))

    total = len(marcas_julio)

    def _pct(n: int) -> float | None:
        return round(n / total * 100.0, 1) if total else None

    return {
        "circuitos_por_marca": {m: int(circuitos_count.get(m, 0)) for m in sorted(marcas_julio)},
        "multicircuito": {"count": len(multicircuito), "pct": _pct(len(multicircuito)), "marcas": multicircuito},
        "un_solo_circuito": {
            "count": len(un_solo_circuito), "pct": _pct(len(un_solo_circuito)), "marcas": un_solo_circuito,
        },
        "sin_circuito_valido": {
            "count": len(sin_circuito_valido), "pct": _pct(len(sin_circuito_valido)), "marcas": sin_circuito_valido,
        },
        "reconciliacion": {
            "multicircuito_mas_uno_mas_sin_circuito": (
                len(multicircuito) + len(un_solo_circuito) + len(sin_circuito_valido)
            ),
            "marcas_activas": total,
            "ok": len(multicircuito) + len(un_solo_circuito) + len(sin_circuito_valido) == total,
        },
    }


# ---------------------------------------------------------------------------
# Ranking "Top marcas · Julio 2026" (panel izquierdo, estatico): orden
# campañas unicas desc -> circuitos desc -> activaciones desc -> nombre asc.
# Las activaciones NUNCA definen el orden (el volumen de elementos YPF
# distorsiona la lectura de demanda real).
# ---------------------------------------------------------------------------


def compute_ranking(scope_julio: pd.DataFrame, circuitos_por_marca: dict[str, int]) -> list[dict[str, Any]]:
    validas = scope_julio[scope_julio["_marca_valida"]]
    campanas = validas.groupby("_marca_canon")[IDCAMPANA_COL].nunique()
    activaciones = validas.groupby("_marca_canon").size()

    filas = pd.DataFrame({
        "campanas_unicas": campanas,
        "circuitos": pd.Series(circuitos_por_marca),
        "activaciones": activaciones,
    }).fillna(0).rename_axis("Marca").reset_index()
    filas = filas.sort_values(
        ["campanas_unicas", "circuitos", "activaciones", "Marca"],
        ascending=[False, False, False, True],
    )

    return [
        {
            "nombre": r["Marca"],
            "campanas_unicas": int(r["campanas_unicas"]),
            "circuitos": int(r["circuitos"]),
            "activaciones": int(r["activaciones"]),
        }
        for _, r in filas.head(RANK_TOP_N).iterrows()
    ]


# ---------------------------------------------------------------------------
# Matriz "Demanda por circuito" (panel derecho): celdas = campañas unicas
# por Marca x Circuito (no activaciones). Totales por circuito = todas las
# marcas validas (no solo el Top 6).
# ---------------------------------------------------------------------------


def compute_matriz(scope_julio: pd.DataFrame, top_nombres: list[str]) -> dict[str, Any]:
    validas = scope_julio[scope_julio["_marca_valida"]]

    sub = validas[validas["_marca_canon"].isin(top_nombres)]
    pivot = sub.pivot_table(
        index="_marca_canon", columns="_familia", values=IDCAMPANA_COL,
        aggfunc=lambda s: s.dropna().nunique(), fill_value=0,
    )
    for col in MATRIZ_COLUMNAS:
        if col not in pivot.columns:
            pivot[col] = 0
    pivot = pivot.reindex(top_nombres)[MATRIZ_COLUMNAS].fillna(0)

    filas = [
        {"nombre": nombre, "valores": [int(pivot.loc[nombre, c]) for c in MATRIZ_COLUMNAS]}
        for nombre in top_nombres
    ]

    totales_por_circuito = {
        col: int(validas.loc[validas["_familia"] == col, "_marca_canon"].nunique()) for col in MATRIZ_COLUMNAS
    }

    return {
        "columnas": MATRIZ_COLUMNAS,
        "totales_por_circuito": [totales_por_circuito[c] for c in MATRIZ_COLUMNAS],
        "filas": filas,
    }


# ---------------------------------------------------------------------------
# Insights (Lectura / Punto positivo / A atender): CM1 Sec.9, una oracion
# breve por columna, siempre con datos ya calculados.
# ---------------------------------------------------------------------------


def compute_insights(
    marcas_julio: dict[str, Any], recurrencia: dict[str, Any], circuitos: dict[str, Any],
) -> dict[str, str]:
    lectura = (
        f"Julio reúne <b>{_fmt_es_int(marcas_julio['activas'])} marcas activas</b>: "
        f"<b>{_fmt_es_int(recurrencia['primera_aparicion']['count'])}</b> registran su primera actividad de "
        f"2026 y <b>{_fmt_es_int(recurrencia['recurrentes']['count'])}</b> ya habían pautado entre enero y junio."
    )
    punto_positivo = (
        f"<b>{_fmt_es_int(circuitos['multicircuito']['count'])} marcas</b> están presentes en dos o más "
        f"circuitos."
    )
    a_atender = (
        f"<b>{_fmt_es_int(circuitos['un_solo_circuito']['count'])} marcas</b> pautan en un solo circuito: "
        f"oportunidad de expansión comercial."
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

    start_jul, end_jul = _period_bounds(REPORT_YEAR, REPORT_MONTH)
    scope_julio = classify_marca(build_tv6_scope(engine, universe["element_ids"], start_jul, end_jul))

    marcas_julio_kpi = compute_marcas_julio(scope_julio)
    marcas_julio_set = set(scope_julio.loc[scope_julio["_marca_valida"], "_marca_canon"].unique())

    presencia_mensual = compute_monthly_brand_presence(engine, universe["element_ids"])
    recurrencia = compute_recurrencia(marcas_julio_set, presencia_mensual)

    circuitos = compute_circuitos_por_marca(scope_julio, marcas_julio_set)

    ranking_top = compute_ranking(scope_julio, circuitos["circuitos_por_marca"])
    top_nombres = [r["nombre"] for r in ranking_top]
    matriz = compute_matriz(scope_julio, top_nombres)

    insights = compute_insights(marcas_julio_kpi, recurrencia, circuitos)

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
                "campanas_unicas": marcas_julio_kpi["campanas_unicas_total"],
                "activaciones_totales": marcas_julio_kpi["activaciones_totales"],
                "activaciones_sin_marca_valida": marcas_julio_kpi["activaciones_sin_marca_valida"],
            },
        },
        "marcas": {
            "activas_julio": marcas_julio_kpi["activas"],
            "campanas_unicas_marca_valida_julio": marcas_julio_kpi["campanas_unicas_marca_valida"],
            "activaciones_sin_marca_valida_julio": marcas_julio_kpi["activaciones_sin_marca_valida"],
            "primera_aparicion": recurrencia["primera_aparicion"],
            "recurrentes": recurrencia["recurrentes"],
            "sin_historial_comparable": recurrencia["sin_historial_comparable"],
            "primer_mes_por_marca": recurrencia["primer_mes_por_marca"],
            "multicircuito": circuitos["multicircuito"],
            "un_solo_circuito": circuitos["un_solo_circuito"],
            "sin_circuito_valido": circuitos["sin_circuito_valido"],
            "circuitos_por_marca": circuitos["circuitos_por_marca"],
        },
        "calidad": {
            "campanas_marca_contradictoria": marcas_julio_kpi["campanas_marca_contradictoria"],
        },
        "reconciliacion": {
            "recurrencia": recurrencia["reconciliacion"],
            "circuitos": circuitos["reconciliacion"],
        },
        "ranking": {
            "periodo_label": f"{MESES_ES[REPORT_MONTH - 1]} {REPORT_YEAR}",
            "top": ranking_top,
        },
        "matriz": {
            "periodo_label": f"{MESES_ES[REPORT_MONTH - 1]} {REPORT_YEAR}",
            **matriz,
        },
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
    print(json.dumps(result["data"]["marcas"], ensure_ascii=False, indent=2, default=str))
    print("UNIVERSE:", json.dumps(result["universe"], ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
