"""Capa de datos para el dashboard TV2 - OCU26 (Core Comercial Digital).

Rediseno completo 2026-08-22 (spec "TV2 - ESPACIOS/SLOTS", reemplaza el
diseno anterior basado en "elementos activos"/"ocupacion por calendario"/
rankings). Unidad de negocio unica: ESPACIO/SLOT digital, misma logica
canonica de capacidad que TV1. Desde Etapa 2A (2026-10-06) la capacidad
sale de la regla central (config/business_semantics.json -> semantic_model.
capacidad_espacio_digital), compartida con TV1/TV5: mismo Excel + misma
regla = mismos numeros, verificado contra output/tv1_data.json cuando existe.

Se ejecuta DESPUES de scripts/semantic_model.py y scripts/metrics_engine.py
(Gate 3B). Reutiliza export_data.load_pipeline (Gate 4A), mismo patron que
build_tv1_dashboard.py: no reabre el Excel, no reimplementa reglas de
negocio de Gate 3 (los slots ocupados salen de MetricsEngine._digital_period_activity,
motor ya aprobado). No importa build_tv1_dashboard.py a proposito: TV1 queda
protegida/aislada, cada TV es un builder independiente sobre el mismo
pipeline compartido.

Universo TV2 = CORE COMERCIAL DIGITAL = Pantallas LED + Shoppings Digital
(Cencosud + Remeros) + AA2000 Digital. YPF/Cencomedia/APSA/London Supply
excluidos (nunca entran, ni siquiera se listan). AA2000 alimenta capacidad/
ocupacion/disponibilidad y la evolucion mensual (3 series) pero no tiene
tarjeta propia (spec Sec.4,14).

Uso:
    python scripts/build_tv2_dashboard.py
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import math
import sys
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import validate_input as vi  # noqa: E402
import semantic_model as sm  # noqa: E402
from semantic_model import filter_universe  # noqa: E402
from metrics_engine import MetricsEngine  # noqa: E402
from export_data import load_pipeline  # noqa: E402

REPO_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = Path(__file__).resolve().parent / "templates" / "tv2_template.html"
REFERENCE_PATH = REPO_ROOT / "audit_sources" / "TV1_REFERENCE.html.html"
DEFAULT_OUTPUT_HTML = REPO_ROOT / "tv2.html"
DEFAULT_OUTPUT_JSON = REPO_ROOT / "output" / "tv2_data.json"
TV1_CONTROL_JSON = REPO_ROOT / "output" / "tv1_data.json"

MESES_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]
MESES_ES_ABR = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

ART_TZ = ZoneInfo("America/Argentina/Buenos_Aires")
ART_TZ_NAME = "America/Argentina/Buenos_Aires"

# Periodo de referencia TV2 (spec Sec.7): Julio 2026, vs Junio 2026, evolucion
# Ene-Jul. Dinamico: cambiar estos dos numeros re-genera todo el tablero.
REPORT_YEAR = 2026
REPORT_MONTH = 7

# Universo TV2 (spec Sec.4): Core Comercial Digital = Pantallas LED +
# Shoppings Digital (Cencosud+Remeros) + AA2000 Digital. YPF/Cencomedia/
# APSA/London nunca entran. Cada ElementoID pertenece a una sola familia
# funcional, determinada por CircuitoNegocio (nunca por nombre de sitio):
# evita que una Pantalla LED ubicada en "Remeros" se confunda con el
# shopping "Remeros" (son CircuitoNegocio distintos: PANTALLAS_LED vs
# REMEROS).
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
FAMILIAS = ["Pantallas", "Shoppings", "AA2000"]

# Capacidad en ESPACIOS/SLOTS (Etapa 2A, 2026-10-06): la regla vive SOLO en
# config/business_semantics.json (digital_capacity) y se resuelve en
# semantic_model (SlotsComerciales + capacidad_espacio_digital). Este builder
# ya no define tasas propias ni excepciones por circuito: Pantalla LED 20,
# Totem 10 (Shopping, Remeros y Tripstore AA2000), Triedro 10, Puente LED 10,
# Patio de Comidas 10, AA2000 digital sin formato propio 10.
# Composicion fisica (tarjeta 1, spec Sec.9): etiquetas centrales; Tripstore
# AA2000 se muestra FISICAMENTE como Totem aunque comercialmente sea AA2000.
_CATEGORIA_A_ETIQUETA: dict[str, str] = dict(sm.ETIQUETA_CATEGORIA_ESPACIO_DIGITAL)
_ETIQUETA_ORDEN = list(sm.ORDEN_ETIQUETAS_ESPACIO_DIGITAL)


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


def _fmt_es_pct(v: float | None) -> str:
    if v is None:
        return "S/D"
    return f"{v:.1f}".replace(".", ",")


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
# Capacidad en espacios/slots (identica a TV1 compute_espacios, Sec.6)
# ---------------------------------------------------------------------------


def _espacio_capacidad_digital(row: pd.Series) -> tuple[float | None, str]:
    """Capacidad en espacios de UN elemento digital TV2 y su categoria de
    reporting: delegado 1:1 en la API central (semantic_model.
    capacidad_espacio_digital), la misma que usan TV1 y TV5. Sin capacidad
    confirmada -> (None, 'SIN_REGLA'), nunca se inventa una cifra."""
    return sm.capacidad_espacio_digital(row)


def _con_capacidad_digital(digital: pd.DataFrame) -> pd.DataFrame:
    """Agrega columnas _capacidad/_categoria/_familia. Ver
    build_tv1_dashboard._con_capacidad_digital (mismo patron, incluye el
    manejo del caso limite de DataFrame vacio)."""
    digital = digital.copy()
    if digital.empty:
        digital["_capacidad"] = pd.Series(dtype="float64")
        digital["_categoria"] = pd.Series(dtype="object")
        digital["_familia"] = pd.Series(dtype="object")
        return digital
    capacidades = digital.apply(_espacio_capacidad_digital, axis=1, result_type="expand")
    digital["_capacidad"] = pd.to_numeric(capacidades[0], errors="coerce")
    digital["_categoria"] = capacidades[1]
    digital["_familia"] = digital["CircuitoNegocio"].map(FAMILY_MAP)
    return digital


def build_catalog_maps(universe: dict[str, Any]) -> dict[str, Any]:
    """Estructuras de capacidad/elementos por familia y por sitio (spec
    Sec.9,16,17): estructural, no depende de periodo/actividad. Devuelve
    tambien el catalogo `cats` completo (incluye SIN_REGLA) para poder
    contar los 19 soportes sin capacidad confirmada."""
    cats = _con_capacidad_digital(universe["maestro"])
    confirmed = cats[cats["_capacidad"].notna()].copy()
    sin_regla = cats[cats["_capacidad"].isna()].copy()

    # Etapa 2A: la validacion "exactamente 6 totems Remeros" se retiro junto
    # con la excepcion OTRO+REMEROS de este builder; la clasificacion de los
    # 6 elementos REM-DB-* como TOTEM vive ahora en config (formato_negocio.
    # elemento_formato_overrides) y se prueba en tests/test_reglas_capacidad.py.

    familia_ids = {fam: confirmed.loc[confirmed["_familia"] == fam, "ElementoID"].tolist() for fam in FAMILIAS}
    familia_capacidad = {fam: float(confirmed.loc[confirmed["_familia"] == fam, "_capacidad"].sum()) for fam in FAMILIAS}

    sitio_ids: dict[str, dict[str, list[Any]]] = {}
    sitio_capacidad: dict[str, dict[str, float]] = {}
    for fam in FAMILIAS:
        sub = confirmed[confirmed["_familia"] == fam]
        sitio_ids[fam] = {s: g["ElementoID"].tolist() for s, g in sub.groupby("SitioNegocio")}
        sitio_capacidad[fam] = {s: float(g["_capacidad"].sum()) for s, g in sub.groupby("SitioNegocio")}

    # Sitios PENDIENTES (Caso C, spec Sec.4.C): sitios con elementos
    # digitales de la familia que existen en el universo TV2 pero SIN regla
    # de conversion a espacios confirmada (p.ej. los 11 "Patio de Comidas"
    # OTRO de Cencosud en Unicenter). Remeros Shoppings YA NO cae aqui desde
    # la correccion 2026-08-22 (paso a `confirmed` via la regla Totem
    # arriba). Nunca se inventa capacidad para lo que queda pendiente: se
    # muestra en las matrices como REQUIERE_CONFIRMACION.
    sitio_pendiente_ids: dict[str, dict[str, list[Any]]] = {}
    for fam in FAMILIAS:
        sub = sin_regla[sin_regla["_familia"] == fam]
        sitio_pendiente_ids[fam] = {s: g["ElementoID"].tolist() for s, g in sub.groupby("SitioNegocio")}

    return {
        "cats": cats,
        "confirmed": confirmed,
        "sin_regla": sin_regla,
        "familia_ids": familia_ids,
        "familia_capacidad": familia_capacidad,
        "sitio_ids": sitio_ids,
        "sitio_capacidad": sitio_capacidad,
        "sitio_pendiente_ids": sitio_pendiente_ids,
    }


def compute_catalogo(maps: dict[str, Any]) -> dict[str, Any]:
    """Tarjeta 1 - Espacios del catalogo digital (spec Sec.9): capacidad
    comercial (930), N elementos con capacidad confirmada (72) y su
    composicion por formato fisico (Pantallas LED/Totems/Puentes/Triedros,
    Tripstore clasificado como Totem)."""
    confirmed = maps["confirmed"]
    familia_capacidad = maps["familia_capacidad"]
    total = sum(familia_capacidad.values())
    n_confirmados = int(confirmed["ElementoID"].nunique())
    n_sin_confirmar = int(maps["sin_regla"]["ElementoID"].nunique())

    etiquetas = confirmed["_categoria"].map(_CATEGORIA_A_ETIQUETA)
    conteo = etiquetas.value_counts()
    formatos = []
    for nombre in _ETIQUETA_ORDEN:
        n = int(conteo.get(nombre, 0))
        if n == 0 and nombre not in conteo.index:
            continue
        formatos.append({
            "nombre": nombre,
            "elementos": n,
            "pct": _round1(n / n_confirmados * 100.0) if n_confirmados else None,
        })
    # Categoria residual: cualquier _categoria confirmada que no mapee a las
    # 4 etiquetas canonicas (no deberia ocurrir con las reglas vigentes, pero
    # nunca se oculta un elemento confirmado de la composicion, spec Sec.9
    # "Otros digitales").
    etiquetados = sum(f["elementos"] for f in formatos)
    resto = n_confirmados - etiquetados
    if resto > 0:
        formatos.append({"nombre": "Otros digitales", "elementos": resto, "pct": _round1(resto / n_confirmados * 100.0)})

    suma_pct = round(sum(f["pct"] or 0.0 for f in formatos), 1)
    if abs(suma_pct - 100.0) > 0.15:
        raise BuildError(f"Porcentajes de composicion por formato no suman ~100% (tolerancia 0.1pp): {suma_pct}")

    if n_confirmados + n_sin_confirmar != int(maps["cats"]["ElementoID"].nunique()):
        raise BuildError("Confirmados + sin confirmar no reconcilia con el universo TV2 total")

    total_i = int(round(total))
    return {
        # Capacidad estructural (catalogo de elementos con formato/tasa
        # confirmada): no depende del periodo/actividad, por eso
        # total_anterior/total_ytd_promedio son iguales a total (spec Sec.9
        # "No mostrar una variación ficticia si la capacidad permaneció
        # estable"). Se exponen igual para que el template no necesite
        # asumir nada por su cuenta.
        "total": total_i,
        "shoppings": int(round(familia_capacidad["Shoppings"])),
        "pantallas": int(round(familia_capacidad["Pantallas"])),
        "aa2000": int(round(familia_capacidad["AA2000"])),
        "total_anterior": total_i,
        "total_ytd_promedio": float(total_i),
        "elementos_confirmados": n_confirmados,
        "elementos_sin_confirmar": n_sin_confirmar,
        "formatos": formatos,
    }


# ---------------------------------------------------------------------------
# Ocupacion mensual por familia/sitio (apportionment, spec Sec.6-7)
# ---------------------------------------------------------------------------


def _apportion(raw: dict[str, float], target: int) -> dict[str, int]:
    """Reparto entero de `target` unidades entre las claves de `raw` segun el
    metodo de mayores restos (Hamilton): cada clave recibe floor(valor), y
    las unidades restantes (target - suma de floors) se asignan una a una a
    las claves con mayor parte fraccionaria. Resuelve de forma determinista
    y auditable la diferencia entre "redondear cada familia por separado"
    (que puede no sumar el total ya redondeado) y "el total ya redondeado"
    (Sec.5: "no corregir manualmente...identificar la regla que genera la
    diferencia"): la regla es este reparto, no un ajuste manual. target debe
    ser >= floor(suma(raw)) siempre se cumple porque floor(a)+floor(b)<=
    floor(a+b)<=round(a+b)."""
    if not raw:
        if target != 0:
            raise BuildError(f"Apportion: target={target} sin claves para repartir")
        return {}
    floors = {k: math.floor(v) for k, v in raw.items()}
    remainder = target - sum(floors.values())
    if remainder < 0 or remainder > len(raw):
        raise BuildError(f"Apportion: remainder={remainder} fuera de rango para {len(raw)} claves (raw={raw}, target={target})")
    fracs = sorted(raw.keys(), key=lambda k: (-(raw[k] - floors[k]), k))
    result = dict(floors)
    for i in range(remainder):
        result[fracs[i]] += 1
    return result


def _family_raw_ocupados(engine: MetricsEngine, familia_ids: dict[str, list[Any]], period: tuple[str, str]) -> dict[str, float]:
    raw = {}
    for fam in FAMILIAS:
        ids = familia_ids[fam]
        if not ids:
            raw[fam] = 0.0
            continue
        slots_s, _seg, _inc, _sal = engine._digital_period_activity(ids, period[0], period[1])
        raw[fam] = float(slots_s.sum()) if len(slots_s) else 0.0
    return raw


def _site_raw_ocupados(engine: MetricsEngine, sitio_ids: dict[str, list[Any]], period: tuple[str, str]) -> dict[str, float]:
    raw = {}
    for sitio, ids in sitio_ids.items():
        if not ids:
            raw[sitio] = 0.0
            continue
        slots_s, _seg, _inc, _sal = engine._digital_period_activity(ids, period[0], period[1])
        raw[sitio] = float(slots_s.sum()) if len(slots_s) else 0.0
    return raw


def compute_monthly(engine: MetricsEngine, maps: dict[str, Any], months: list[tuple[int, int]]) -> dict[tuple[int, int], dict[str, Any]]:
    """Un snapshot por mes (spec Sec.6-7,15-17): ocupados por familia
    (apportion sobre el total del Core ya redondeado) y ocupados por sitio
    dentro de cada familia (apportion sobre el entero de esa familia).
    Garantiza, para TODOS los meses (no solo julio): suma(familias)=core,
    suma(sitios de una familia)=familia (spec Sec.20.G)."""
    out: dict[tuple[int, int], dict[str, Any]] = {}
    for year, month in months:
        period = _period_bounds(year, month)
        fam_raw = _family_raw_ocupados(engine, maps["familia_ids"], period)
        core_target = round(sum(fam_raw.values()))
        fam_int = _apportion(fam_raw, core_target)

        sitio_int: dict[str, dict[str, int]] = {}
        for fam in FAMILIAS:
            sraw = _site_raw_ocupados(engine, maps["sitio_ids"][fam], period)
            sitio_int[fam] = _apportion(sraw, fam_int[fam]) if sraw else {}
            if sraw and sum(sitio_int[fam].values()) != fam_int[fam]:
                raise BuildError(f"Matriz {fam} {year}-{month:02d} no reconcilia con la tarjeta: {sum(sitio_int[fam].values())} vs {fam_int[fam]}")

        out[(year, month)] = {"raw": fam_raw, "core_target": core_target, "fam_int": fam_int, "sitio_int": sitio_int}
    return out


# ---------------------------------------------------------------------------
# Tarjetas 2-5 (Ocupacion / Disponibles / Pantallas / Shoppings)
# ---------------------------------------------------------------------------


def _family_card(nombre: str, capacidad: int, ocup_actual: int, ocup_anterior: int, ytd_sum: int, n_meses_ytd: int) -> dict[str, Any]:
    disp_actual = capacidad - ocup_actual
    disp_anterior = capacidad - ocup_anterior
    fill_actual = _round1(ocup_actual / capacidad * 100.0) if capacidad else None
    fill_anterior = _round1(ocup_anterior / capacidad * 100.0) if capacidad else None
    disp_pct_actual = _round1(disp_actual / capacidad * 100.0) if capacidad else None
    disp_pct_anterior = _round1(disp_anterior / capacidad * 100.0) if capacidad else None
    cap_ytd = capacidad * n_meses_ytd
    return {
        "nombre": nombre,
        "capacidad": capacidad,
        "ocupados": ocup_actual,
        "fill_pct": fill_actual,
        "disponibles": disp_actual,
        "disp_pct": disp_pct_actual,
        "capacidad_anterior": capacidad,
        "ocupados_anterior": ocup_anterior,
        "fill_pct_anterior": fill_anterior,
        "disponibles_anterior": disp_anterior,
        "disp_pct_anterior": disp_pct_anterior,
        "delta_ocupados": ocup_actual - ocup_anterior,
        "delta_pp": _round1(fill_actual - fill_anterior) if fill_actual is not None and fill_anterior is not None else None,
        "delta_disponibles": disp_actual - disp_anterior,
        "delta_disp_pp": _round1(disp_pct_actual - disp_pct_anterior) if disp_pct_actual is not None and disp_pct_anterior is not None else None,
        "ocupados_ytd_promedio": _round1(ytd_sum / n_meses_ytd) if n_meses_ytd else None,
        "fill_pct_ytd": _round1(ytd_sum / cap_ytd * 100.0) if cap_ytd else None,
        "disponibles_ytd_promedio": _round1((cap_ytd - ytd_sum) / n_meses_ytd) if n_meses_ytd else None,
        "disp_pct_ytd": _round1((cap_ytd - ytd_sum) / cap_ytd * 100.0) if cap_ytd else None,
    }


def compute_cards(catalogo: dict[str, Any], monthly: dict[tuple[int, int], dict[str, Any]], ytd_months: list[tuple[int, int]], report_key: tuple[int, int], prev_key: tuple[int, int]) -> dict[str, Any]:
    n_ytd = len(ytd_months)
    ytd_sum_fam = {fam: sum(monthly[m]["fam_int"][fam] for m in ytd_months) for fam in FAMILIAS}

    capacidad = {"Pantallas": catalogo["pantallas"], "Shoppings": catalogo["shoppings"], "AA2000": catalogo["aa2000"]}
    cards = {
        fam: _family_card(fam, capacidad[fam], monthly[report_key]["fam_int"][fam], monthly[prev_key]["fam_int"][fam], ytd_sum_fam[fam], n_ytd)
        for fam in FAMILIAS
    }

    core_cap = catalogo["total"]
    core_ocup_actual = monthly[report_key]["core_target"]
    core_ocup_anterior = monthly[prev_key]["core_target"]
    core_ytd_sum = sum(monthly[m]["core_target"] for m in ytd_months)
    core_card = _family_card("Core Digital", core_cap, core_ocup_actual, core_ocup_anterior, core_ytd_sum, n_ytd)

    esperado_core = cards["Pantallas"]["ocupados"] + cards["Shoppings"]["ocupados"] + cards["AA2000"]["ocupados"]
    if esperado_core != core_card["ocupados"]:
        raise BuildError(f"Ocupados por familia no reconcilian con el Core: {esperado_core} vs {core_card['ocupados']}")
    esperado_disp = cards["Pantallas"]["disponibles"] + cards["Shoppings"]["disponibles"] + cards["AA2000"]["disponibles"]
    if esperado_disp != core_card["disponibles"]:
        raise BuildError(f"Disponibles por familia no reconcilian con el Core: {esperado_disp} vs {core_card['disponibles']}")

    return {"core": core_card, "pantallas": cards["Pantallas"], "shoppings": cards["Shoppings"], "aa2000": cards["AA2000"]}


# ---------------------------------------------------------------------------
# Evolucion mensual (3 series, spec Sec.15)
# ---------------------------------------------------------------------------


def compute_evolution(catalogo: dict[str, Any], monthly: dict[tuple[int, int], dict[str, Any]], months: list[tuple[int, int]]) -> dict[str, Any]:
    capacidad = {"Pantallas": catalogo["pantallas"], "Shoppings": catalogo["shoppings"], "AA2000": catalogo["aa2000"]}
    meses = [MESES_ES_ABR[m - 1] for _y, m in months]
    series = {}
    for fam in FAMILIAS:
        ocupados = [monthly[k]["fam_int"][fam] for k in months]
        fill = [_round1(o / capacidad[fam] * 100.0) if capacidad[fam] else None for o in ocupados]
        series[fam] = {"ocupados": ocupados, "fill": fill}

    total_ocupados = [monthly[k]["core_target"] for k in months]
    for i, (year, month) in enumerate(months):
        suma_familias = sum(series[fam]["ocupados"][i] for fam in FAMILIAS)
        if suma_familias != total_ocupados[i]:
            raise BuildError(f"Evolucion {year}-{month:02d}: suma de familias {suma_familias} != total {total_ocupados[i]}")

    return {
        "meses": meses,
        "pantallas": series["Pantallas"],
        "shoppings": series["Shoppings"],
        "aa2000": series["AA2000"],
        "total_ocupados": total_ocupados,
    }


def _reconcile_evolution_with_tv1(
    evolution: dict[str, Any], months: list[tuple[int, int]], input_sha256: str | None = None,
) -> None:
    """Reconciliacion obligatoria (spec Sec.15,20.F) contra la fuente
    READ-ONLY output/tv1_data.json: la suma mensual de las 3 familias TV2
    debe coincidir exacto con evolution_espacios.digital de TV1 (misma
    formula, mismo Excel). Si el archivo de control no existe todavia (TV1
    nunca se construyo en este checkout) se omite con un aviso: TV2 nunca
    escribe ni depende en tiempo de ejecucion de que TV1 se reconstruya.
    Etapa 2A: tambien se omite (con aviso, nunca en silencio) si el control
    fue generado con OTRA base u OTRA version de reglas (meta.input_sha256 /
    meta.reglas_version de TV1): comparar contra un control desactualizado
    no prueba nada y bloqueaba builds validos."""
    if not TV1_CONTROL_JSON.exists():
        print(f"TV2_RECONCILE_SKIP: no existe {TV1_CONTROL_JSON}, no se pudo reconciliar contra TV1")
        return
    tv1_data = json.loads(TV1_CONTROL_JSON.read_text(encoding="utf-8"))
    meta1 = tv1_data.get("meta", {})
    if meta1.get("reglas_version") != sm.reglas_version() or (
        input_sha256 is not None and meta1.get("input_sha256") != input_sha256
    ):
        print(
            f"TV2_RECONCILE_SKIP: {TV1_CONTROL_JSON} fue generado con otra base o version de reglas "
            f"(control: reglas={meta1.get('reglas_version')!r}, sha={str(meta1.get('input_sha256'))[:12]!r}; "
            f"actual: reglas={sm.reglas_version()!r}, sha={str(input_sha256)[:12]!r}); reconstruir TV1 primero"
        )
        return
    ev1 = tv1_data.get("evolution_espacios")
    if not ev1:
        return
    tv1_meses = ev1["meses"]
    tv1_digital = ev1["digital"]
    tv1_by_label = dict(zip(tv1_meses, tv1_digital))
    for i, label in enumerate(evolution["meses"]):
        if label not in tv1_by_label:
            continue
        esperado = tv1_by_label[label]
        obtenido = evolution["total_ocupados"][i]
        if esperado != obtenido:
            raise BuildError(
                f"Evolucion TV2 no reconcilia con TV1 (output/tv1_data.json) en {label}: "
                f"TV2={obtenido} vs TV1={esperado}"
            )


# ---------------------------------------------------------------------------
# Matrices mensuales por sitio (spec Sec.16-18)
# ---------------------------------------------------------------------------


def compute_matrix(fam: str, maps: dict[str, Any], monthly: dict[tuple[int, int], dict[str, Any]], months: list[tuple[int, int]]) -> dict[str, Any]:
    sitio_capacidad = maps["sitio_capacidad"][fam]
    # Un sitio con AL MENOS UN elemento confirmado (p.ej. Unicenter, que ya
    # tiene Totems/Puentes/Triedros confirmados y ADEMAS algunos "Patio de
    # Comidas" OTRO sin regla) mantiene su fila normal con la capacidad
    # confirmada. Solo se agrega una fila "pendiente" nueva para sitios que
    # hoy NO tienen ninguna presencia confirmada en la matriz (Remeros dejo
    # de ser un ejemplo de esto desde la correccion 2026-08-22: sus 6 totems
    # ya son `confirmed`, ver build_catalog_maps).
    sitio_pendiente = {s: ids for s, ids in maps["sitio_pendiente_ids"].get(fam, {}).items() if s not in sitio_capacidad}
    sitios = sorted(set(sitio_capacidad) | set(sitio_pendiente), key=str.casefold)
    meses = [MESES_ES_ABR[m - 1] for _y, m in months]

    filas = []
    for sitio in sitios:
        if sitio in sitio_pendiente:
            # Caso C (spec correccion Remeros 2026-08-22): sitio con elementos
            # digitales reales pero sin regla de conversion a espacios
            # confirmada. Nunca se inventan slots: REQUIERE_CONFIRMACION en
            # las 7 columnas, fuera de la capacidad/ocupacion/fill rate de la
            # tarjeta (no participa de la suma de reconciliacion de abajo).
            # Label corto para que la celda entre sin cortarse (spec Sec.23
            # "sin textos cortados"): el estado real (REQUIERE_CONFIRMACION)
            # queda en `status`, mas el asterisco de fila + el pie de matriz
            # explican el motivo completo.
            celdas = [{"ocupados": None, "capacidad": None, "fill_pct": None, "label": "S/D", "status": "REQUIERE_CONFIRMACION"} for _ in months]
            filas.append({
                "sitio": sitio, "capacidad": None, "celdas": celdas,
                "pendiente": True, "elementos_pendientes": len(sitio_pendiente[sitio]),
            })
            continue
        capacidad = int(round(sitio_capacidad[sitio]))
        celdas = []
        for key in months:
            ocup = monthly[key]["sitio_int"][fam].get(sitio)
            if ocup is None or capacidad == 0:
                celdas.append({"ocupados": None, "capacidad": capacidad, "fill_pct": None, "label": "N/A"})
                continue
            fill_pct = _round1(ocup / capacidad * 100.0)
            celdas.append({"ocupados": ocup, "capacidad": capacidad, "fill_pct": fill_pct, "label": f"{_fmt_es_int(ocup)}/{_fmt_es_int(capacidad)}"})
        filas.append({"sitio": sitio, "capacidad": capacidad, "celdas": celdas, "pendiente": False})

    ultimo = months[-1]
    suma_ultimo = sum((f["celdas"][-1]["ocupados"] or 0) for f in filas)
    esperado = monthly[ultimo]["fam_int"][fam]
    if suma_ultimo != esperado:
        raise BuildError(f"Matriz {fam}: suma del ultimo mes ({suma_ultimo}) no reconcilia con la tarjeta ({esperado})")

    return {"meses": meses, "filas": filas}


# ---------------------------------------------------------------------------
# Insights (Lectura / Punto positivo / A atender) - seleccion programatica
# ---------------------------------------------------------------------------


def compute_insights(
    catalogo: dict[str, Any], cards: dict[str, Any],
    report_month_label: str, previous_month_label: str,
) -> dict[str, str]:
    core, pant, shop, aa = cards["core"], cards["pantallas"], cards["shoppings"], cards["aa2000"]

    lectura = (
        f"{report_month_label} registra <b>{_fmt_es_int(core['ocupados'])} slots ocupados</b> sobre "
        f"{_fmt_es_int(catalogo['total'])} espacios de capacidad digital "
        f"(Pantallas LED {_fmt_es_int(catalogo['pantallas'])} + Shoppings Digital {_fmt_es_int(catalogo['shoppings'])} + "
        f"AA2000 {_fmt_es_int(catalogo['aa2000'])}). Fill rate <b>{_fmt_es_pct(core['fill_pct'])}%</b>, "
        f"disponibles <b>{_fmt_es_int(core['disponibles'])} slots</b> ({_fmt_es_pct(core['disp_pct'])}%)."
    )

    # Punto positivo: familia con MENOR caida de fill (o mejora, si la hay).
    familias_delta = [(nombre, c["delta_pp"]) for nombre, c in (("Pantallas LED", pant), ("Shoppings Digital", shop)) if c["delta_pp"] is not None]
    mejoras = [x for x in familias_delta if x[1] > 0]
    if mejoras:
        mejoras.sort(key=lambda x: -x[1])
        nombre, delta = mejoras[0]
        punto_positivo = f"{nombre} mejora <b>{_fmt_es_pct(delta)} pp</b> de fill rate frente a {previous_month_label.lower()}."
    elif familias_delta:
        familias_delta.sort(key=lambda x: -x[1])  # menos negativo primero = menor caida
        nombre, delta = familias_delta[0]
        punto_positivo = f"{nombre} registra la menor caída de fill rate (<b>{_fmt_es_pct(delta)} pp</b>) frente a {previous_month_label.lower()}."
    else:
        punto_positivo = f"El Core Digital sostiene <b>{_fmt_es_pct(core['fill_pct'])}%</b> de fill rate en {report_month_label.lower()}."

    # A atender: familia con MAYOR caida + AA2000 sin ocupacion + soportes sin confirmar.
    partes_atender = []
    if familias_delta:
        peor_nombre, peor_delta = min(familias_delta, key=lambda x: x[1])
        if peor_delta < 0:
            partes_atender.append(f"{peor_nombre} cae <b>{_fmt_es_pct(abs(peor_delta))} pp</b> de fill frente a {previous_month_label.lower()}")
    if aa["ocupados"] == 0:
        partes_atender.append(f"AA2000 sigue sin ocupación (0/{_fmt_es_int(aa['capacidad'])})")
    if catalogo["elementos_sin_confirmar"] > 0:
        partes_atender.append(f"{_fmt_es_int(catalogo['elementos_sin_confirmar'])} soportes digitales sin capacidad confirmada")
    a_atender = "; ".join(partes_atender) + "." if partes_atender else f"El Core Digital conserva <b>{_fmt_es_int(core['disponibles'])} slots disponibles</b> por vender."

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


def build_tv2_data(path: str | Path | None = None) -> dict[str, Any]:
    path = vi.resolve_input_path(path)
    sha_before = vi.calculate_sha256(path)

    _transform_result, semantic_result, engine = load_pipeline(path)
    universe = build_tv2_universe(semantic_result)
    maps = build_catalog_maps(universe)

    prev_year, prev_month = _previous_month(REPORT_YEAR, REPORT_MONTH)
    report_key = (REPORT_YEAR, REPORT_MONTH)
    prev_key = (prev_year, prev_month)
    ytd_months = [(REPORT_YEAR, m) for m in range(1, REPORT_MONTH + 1)]
    all_months = sorted(set(ytd_months) | {prev_key})

    monthly = compute_monthly(engine, maps, all_months)

    catalogo = compute_catalogo(maps)
    cards = compute_cards(catalogo, monthly, ytd_months, report_key, prev_key)
    evolution = compute_evolution(catalogo, monthly, ytd_months)
    _reconcile_evolution_with_tv1(evolution, ytd_months, sha_before)
    matrix_shoppings = compute_matrix("Shoppings", maps, monthly, ytd_months)
    matrix_pantallas = compute_matrix("Pantallas", maps, monthly, ytd_months)
    insights = compute_insights(catalogo, cards, MESES_ES[REPORT_MONTH - 1], MESES_ES[prev_month - 1])

    # AA2000 (Sec.14): sin tarjeta propia, apertura Aeroparque/Ezeiza para
    # auditoria + insights.
    aa2000_confirmed = maps["confirmed"][maps["confirmed"]["_familia"] == "AA2000"]
    aeroparque_n = int(aa2000_confirmed.loc[aa2000_confirmed["SitioNegocio"] == "AEROPARQUE", "ElementoID"].nunique())
    ezeiza_n = int(aa2000_confirmed.loc[aa2000_confirmed["SitioNegocio"] == "EZEIZA", "ElementoID"].nunique())
    if aeroparque_n + ezeiza_n != int(aa2000_confirmed["ElementoID"].nunique()):
        raise BuildError("AA2000: Aeroparque + Ezeiza no reconcilia con el total de elementos confirmados")
    cards["aa2000"]["aeroparque_elementos"] = aeroparque_n
    cards["aa2000"]["ezeiza_elementos"] = ezeiza_n
    cards["aa2000"]["elementos"] = aeroparque_n + ezeiza_n

    cards["pantallas"]["elementos"] = int(maps["confirmed"].loc[maps["confirmed"]["_familia"] == "Pantallas", "ElementoID"].nunique())
    cards["shoppings"]["elementos"] = int(maps["confirmed"].loc[maps["confirmed"]["_familia"] == "Shoppings", "ElementoID"].nunique())

    # Exclusiones absolutas (spec Sec.4,20.J): nunca deben aparecer en el
    # universo TV2 (guard explicito, ademas de la construccion por lista de
    # circuitos permitidos).
    for excluido in ("YPF", "CENCOMEDIA", "APSA", "LONDON_SUPPLY"):
        if excluido in universe["circuitos"]:
            raise BuildError(f"Circuito excluido {excluido} presente en el universo TV2")

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
            "period_start": _period_bounds(REPORT_YEAR, REPORT_MONTH)[0],
            "period_end": _period_bounds(REPORT_YEAR, REPORT_MONTH)[1],
            "fuente": "OCU26 · Base maestra + base campañas",
        },
        "catalogo": catalogo,
        "cards": cards,
        "evolution": evolution,
        "matrix_shoppings": matrix_shoppings,
        "matrix_pantallas": matrix_pantallas,
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
    path: str | Path | None = None,
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
    parser.add_argument("--file", default=None, help="Ruta al .xlsx (si se omite: $OCU26_INPUT_PATH o input/OCU26_BASE_DATOS.xlsx)")
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
    print(json.dumps(result["data"]["catalogo"], ensure_ascii=False, indent=2))
    print(json.dumps({k: {kk: vv for kk, vv in v.items() if kk not in ("nombre",)} for k, v in result["data"]["cards"].items()}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
