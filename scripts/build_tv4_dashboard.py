"""Capa de datos para el dashboard TV4 - OCU26 (Pulso comercial YPF).

TV4 reemplaza el scope anterior de este slot ("Pipeline Comercial") por
decision explicita del usuario: YPF pasa a reportarse como TV4 (no como
TV5). audit_sources/TV5_REFERENCE.html (alias "tv5(1).html" en el pedido)
se usa EXCLUSIVAMENTE como referencia visual historica: sus cifras
embebidas (551 estaciones, 4.048 elementos) son la version SIN corregir
del universo YPF y no se reutilizan aqui como verdad.

Se ejecuta DESPUES de scripts/semantic_model.py y scripts/metrics_engine.py
(Gate 3B). Reutiliza export_data.load_pipeline (Gate 4A) y
MetricsEngine._campanas_overlap, mismo patron que build_tv5_dashboard.py:
no reabre el Excel, no reimplementa Gate 1/2/3. Tambien reutiliza
build_tv5_dashboard.load_geo_aux (fuente geografica YPF ya auditada) para
no duplicar esa logica de validacion lat/lon.

Universo TV4 = exclusivamente CircuitoNegocio YPF (igual que TV5), pero con
una identidad de estacion distinta y mas precisa: la fuente actual SI trae
un codigo de estacion real (columna Subcircuito, identico por construccion
al prefijo numerico de ElementoID) -- se usa como APIE canonico, no el
surrogate "prefijo+localidad" que TV5 usa por no tener alternativa.

Filtro de vigencia (Sec.2/3 del pedido; auditado contra el control
preliminar conocido -- ver docs de entrega):
- CircuitoNegocio == "YPF".
- RevisionMaestro no nulo: 165 filas YPF sin ese sello son remanentes
  previos a la correccion estructural "BASE LIMPIA YPF ETAPA 1 CORREGIDA -
  13/08/2026" y se excluyen del catalogo vigente (quedan en calidad.avisos).
Con ese filtro el universo reconcilia EXACTO contra el control conocido:
3.883 elementos, 3.500 digitales, 383 estaticos, 412 estaciones con
inventario digital, 27 campanas, 13.616 asignaciones en CAMPANAS. La unica
cifra que no reconcilia 1:1 es estaciones (521 vs 603 del control
preliminar): se reporta el valor recalculado y se explica la diferencia en
calidad.avisos y en el informe de entrega, en vez de forzar 603.

Catalogo de espacios (Sec.4 del pedido): por estacion (APIE),
espacios_digitales = 5 si tiene >=1 elemento MB/TT/PPUNTER, si no 0;
espacios_estaticos = cantidad de ElementoID FB distintos de esa estacion.
Los elementos digitales fisicos (aunque haya mas de 5) NO amplian la
capacidad de 5 espacios por estacion.

Uso:
    python scripts/build_tv4_dashboard.py
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
from build_tv5_dashboard import (  # noqa: E402
    load_geo_aux,
    ARG_PAIS_GEOJSON_PATH,
    ARG_PROVINCIAS_GEOJSON_PATH,
)

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

# Periodo de referencia TV4 (pedido: julio 2026 vs junio 2026).
REPORT_YEAR = 2026
REPORT_MONTH = 7

# Capacidad canonica de espacios digitales por estacion (Sec.4 del pedido).
DIGITAL_CAPACITY_PER_STATION = 5

# Recuadro geografico AMBA (CABA + Gran Buenos Aires + La Plata/Ensenada),
# usado solo para el recorte "Detalle AMBA" del mapa (ajuste 2026-08-23).
# Auditado contra las coordenadas reales de produccion: de los 78 puntos con
# coordenada valida, 55 caen dentro de este recuadro (CABA y localidades de
# Zona Norte/Oeste/Sur + La Plata/Ensenada); los 23 restantes son localidades
# del interior bonaerense u otras provincias (ej. Junin, Lincoln, San Pedro,
# Pérez) correctamente excluidas. No es un promedio ni una suposicion: es un
# filtro geografico fijo sobre coordenadas reales ya validadas.
AMBA_BOUNDS = {"lat_min": -35.2, "lat_max": -34.25, "lon_min": -59.3, "lon_max": -57.7}


def _en_amba(lat: float, lon: float) -> bool:
    return AMBA_BOUNDS["lat_min"] <= lat <= AMBA_BOUNDS["lat_max"] and AMBA_BOUNDS["lon_min"] <= lon <= AMBA_BOUNDS["lon_max"]


# ---------------------------------------------------------------------------
# Clasificacion CABA / Zona Norte por cruce punto-en-poligono (Sec.4 del
# ajuste "Corrección del mapa AMBA", 2026-08-23) contra geometria politica
# REAL: comunas de CABA (data.buenosaires.gob.ar/dataset/comunas, GCBA,
# descargado 2026-08-23 con autorizacion explicita del usuario) y partidos
# de la Provincia de Buenos Aires (IGN, via infra.datos.gob.ar/georef/
# departamentos.geojson filtrado a provincia=06, mismo dia). No existe en la
# fuente YPF un campo "Zona Norte" oficial: por eso se resuelve por cruce
# geografico documentado contra poligonos reales, nunca a ojo ni
# hardcodeado. "Zona Norte" es la convencion estandar de mercado (primer/
# segundo/tercer cordon norte del GBA), la misma que usan camaras
# inmobiliarias y medios: Vicente Lopez, San Isidro, San Fernando, Tigre,
# General San Martin, Tres de Febrero, San Miguel, Jose C. Paz, Malvinas
# Argentinas, Escobar y Pilar.
ZONA_NORTE_PARTIDOS = {
    "Vicente López", "San Isidro", "San Fernando", "Tigre", "General San Martín",
    "Tres de Febrero", "San Miguel", "José C. Paz", "Malvinas Argentinas",
    "Escobar", "Pilar",
}

GEO_ASSETS_DIR = Path(__file__).resolve().parent / "templates" / "assets"
CABA_COMUNAS_PATH = GEO_ASSETS_DIR / "caba_comunas.geojson"
GBA_PARTIDOS_PATH = GEO_ASSETS_DIR / "gba_partidos.geojson"


def _point_in_ring(lon: float, lat: float, ring: list[list[float]]) -> bool:
    """Ray casting estandar (Jordan curve theorem). Sin dependencias
    externas (shapely no esta instalado en este entorno)."""
    inside = False
    n = len(ring)
    x1, y1 = ring[0]
    for i in range(1, n + 1):
        x2, y2 = ring[i % n]
        if ((y1 > lat) != (y2 > lat)) and (lon < (x2 - x1) * (lat - y1) / (y2 - y1) + x1):
            inside = not inside
        x1, y1 = x2, y2
    return inside


def _point_in_geometry(lon: float, lat: float, geometry: dict[str, Any]) -> bool:
    polys = geometry["coordinates"] if geometry["type"] == "MultiPolygon" else [geometry["coordinates"]]
    for rings in polys:
        if not rings:
            continue
        if _point_in_ring(lon, lat, rings[0]) and not any(_point_in_ring(lon, lat, hole) for hole in rings[1:]):
            return True
    return False


def _geometry_bbox(geometry: dict[str, Any]) -> tuple[float, float, float, float]:
    polys = geometry["coordinates"] if geometry["type"] == "MultiPolygon" else [geometry["coordinates"]]
    lons, lats = [], []
    for rings in polys:
        for ring in rings:
            for x, y in ring:
                lons.append(x)
                lats.append(y)
    return (min(lons), min(lats), max(lons), max(lats))


def _load_geo_features(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        raise BuildError(f"Geometria politica no encontrada: {path}")
    data = json.loads(path.read_text(encoding="utf-8"))
    features = []
    for f in data["features"]:
        features.append({"props": f["properties"], "geometry": f["geometry"], "bbox": _geometry_bbox(f["geometry"])})
    return features


def clasificar_caba_zona_norte(puntos_amba: list[dict[str, Any]]) -> dict[str, Any]:
    """Clasifica cada punto AMBA por cruce punto-en-poligono real contra
    comunas de CABA y partidos de GBA (nunca por texto libre ni suposicion
    visual). Devuelve el detalle por punto (comuna/partido/zona) mas los
    conteos agregados para las tarjetas 'CABA · N' / 'Zona Norte · N'."""
    comunas = _load_geo_features(CABA_COMUNAS_PATH)
    partidos = _load_geo_features(GBA_PARTIDOS_PATH)

    detalle: dict[Any, dict[str, Any]] = {}
    n_caba = 0
    n_zona_norte = 0
    for p in puntos_amba:
        lon, lat = p["lon"], p["lat"]
        comuna = None
        for c in comunas:
            bx0, by0, bx1, by1 = c["bbox"]
            if not (bx0 <= lon <= bx1 and by0 <= lat <= by1):
                continue
            if _point_in_geometry(lon, lat, c["geometry"]):
                comuna = c["props"]["comuna"]
                break
        partido = None
        if comuna is None:
            for pt in partidos:
                bx0, by0, bx1, by1 = pt["bbox"]
                if not (bx0 <= lon <= bx1 and by0 <= lat <= by1):
                    continue
                if _point_in_geometry(lon, lat, pt["geometry"]):
                    partido = pt["props"]["nombre"]
                    break
        zona = "CABA" if comuna is not None else ("Zona Norte" if partido in ZONA_NORTE_PARTIDOS else None)
        if zona == "CABA":
            n_caba += 1
        elif zona == "Zona Norte":
            n_zona_norte += 1
        detalle[p["apie"]] = {"comuna": comuna, "partido": partido, "zona": zona}

    return {
        "detalle": detalle,
        "n_caba": n_caba,
        "n_zona_norte": n_zona_norte,
        "zona_norte_partidos": sorted(ZONA_NORTE_PARTIDOS),
    }


DIGITAL_TOKENS = {"MB", "TT", "PPUNTER"}
STATIC_TOKENS = {"FB"}
TOKEN_LABELS = {"MB": "Menú Board", "TT": "Torres", "PPUNTER": "Punteras", "FB": "Fotobox"}
FORMATOS_DIGITALES = ["MB", "TT", "PPUNTER"]
_ELEMENTO_TOKEN_RE = re.compile(r"^(.+) - ([A-Za-z]+) - (\d+)$")

# ---------------------------------------------------------------------------
# Provincia por estacion (Sec.4 del pedido de ajustes de mapa): la fuente no
# tiene columna Provincia estructurada, pero MAESTRO_ELEMENTOS.Observaciones
# trae "Provincia: X | Barrio: ... | Direccion: ..." como texto libre para la
# mayoria de los elementos (auditado: 2.439 de 3.883, y a nivel estacion --
# con >=1 elemento con el dato -- 429 de 521). Se parsea ese token y se
# normaliza SOLO contra nombres de provincia argentina inequivocos (lista
# CANON abajo) mas "GBA Norte/Sur/Oeste" -> Buenos Aires (Gran Buenos Aires
# es, sin ambiguedad, parte de la provincia de Buenos Aires). Nombres de
# localidad sueltos (ej. "Saladillo", "Junin") que aparecen en el campo libre
# por error de carga NO se fuerzan a una provincia por adivinanza: quedan
# sin resolver, salvo que OTRO elemento de la MISMA estacion (misma
# ubicacion fisica) ya tenga un nombre de provincia canonico -- eso no es
# inventar, es usar el resto de datos de la propia estacion. "GBA - Revisar"
# tampoco se resuelve: la propia fuente lo marca como pendiente de revision.
_PROVINCIA_RAW_RE = re.compile(r"Provincia:\s*([^|]+)")

PROVINCIA_CANON = {
    "CORDOBA": "Córdoba", "SANTA FE": "Santa Fe", "MENDOZA": "Mendoza",
    "RIO NEGRO": "Río Negro", "CHUBUT": "Chubut", "TUCUMAN": "Tucumán",
    "BUENOS AIRES": "Buenos Aires", "NEUQUEN": "Neuquén", "SAN JUAN": "San Juan",
    "SALTA": "Salta", "ENTRE RIOS": "Entre Ríos", "MISIONES": "Misiones",
    "CORRIENTES": "Corrientes", "CHACO": "Chaco", "SANTA CRUZ": "Santa Cruz",
    "SAN LUIS": "San Luis", "CABA": "Ciudad de Buenos Aires", "LA PAMPA": "La Pampa",
    "LA RIOJA": "La Rioja", "CATAMARCA": "Catamarca", "JUJUY": "Jujuy",
    "FORMOSA": "Formosa", "SANTIAGO DEL ESTERO": "Santiago del Estero",
    "TIERRA DEL FUEGO": "Tierra del Fuego",
    "GBA NORTE": "Buenos Aires", "GBA SUR": "Buenos Aires", "GBA OESTE": "Buenos Aires",
}


def _parse_provincia_raw(observaciones: Any) -> str | None:
    if not isinstance(observaciones, str):
        return None
    m = _PROVINCIA_RAW_RE.search(observaciones)
    if not m:
        return None
    return m.group(1).strip()


def _canon_provincia(raw: str | None) -> str | None:
    if not isinstance(raw, str):
        return None
    return PROVINCIA_CANON.get(raw.strip().upper())


class BuildError(Exception):
    """Error bloqueante al construir el dashboard TV4."""


def _period_bounds(year: int, month: int) -> tuple[str, str]:
    start = pd.Timestamp(year=year, month=month, day=1)
    end = start + pd.offsets.MonthEnd(0)
    return str(start.date()), str(end.date())


def _previous_month(year: int, month: int) -> tuple[int, int]:
    ts = pd.Timestamp(year=year, month=month, day=1) - pd.DateOffset(months=1)
    return ts.year, ts.month


def _fmt_es_int(n: int) -> str:
    return f"{n:,}".replace(",", ".")


def _fmt_es_pct(v: float | None) -> str:
    if v is None:
        return "S/D"
    return f"{v:.1f}".replace(".", ",")


def _round1(value: Any) -> float | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    return round(float(value), 1)


def _signed_int(v: int) -> str:
    return f"+{v}" if v > 0 else str(v)


def _signed_pp(v: float | None) -> str:
    if v is None:
        return "S/D"
    s = "+" if v > 0 else ""
    return f"{s}{_fmt_es_pct(v)} pp"


def _elemento_token(elemento_id: Any) -> str | None:
    m = _ELEMENTO_TOKEN_RE.match(str(elemento_id))
    if not m:
        return None
    return m.group(2).upper()


# ---------------------------------------------------------------------------
# Universo TV4 (exclusivamente YPF, con APIE real y catalogo de espacios)
# ---------------------------------------------------------------------------


def build_tv4_universe(semantic_result: dict[str, Any]) -> dict[str, Any]:
    maestro = semantic_result["maestro"]
    config = semantic_result["config"]

    op = filter_universe(maestro, "OPERATIVO_GENERAL", config)
    ypf_all = op[op["CircuitoNegocio"] == "YPF"].copy()
    if ypf_all.empty:
        raise BuildError("Universo TV4 vacio: no hay elementos YPF en OPERATIVO_GENERAL")

    ypf_all["_APIE"] = ypf_all["Subcircuito"].apply(lambda v: str(v).strip() if v is not None and not (isinstance(v, float) and pd.isna(v)) else None)
    null_apie = ypf_all[ypf_all["_APIE"].isna() | (ypf_all["_APIE"] == "")]
    if len(null_apie):
        raise BuildError(
            f"{len(null_apie)} elemento(s) YPF sin APIE (Subcircuito) derivable: "
            f"{sorted(null_apie['ElementoID'].astype(str).unique().tolist())[:10]}."
        )

    ypf_all["_token"] = ypf_all["ElementoID"].apply(_elemento_token)
    unknown_token = ypf_all[~ypf_all["_token"].isin(DIGITAL_TOKENS | STATIC_TOKENS)]
    if len(unknown_token):
        raise BuildError(
            f"{len(unknown_token)} ElementoID YPF con token de formato no reconocido "
            f"(se esperaba MB/TT/PPUNTER/FB): {sorted(unknown_token['ElementoID'].astype(str).unique().tolist())[:10]}."
        )

    revision_excluidos = ypf_all[ypf_all["RevisionMaestro"].isna()]
    maestro_vigente = ypf_all[ypf_all["RevisionMaestro"].notna()].copy()
    if maestro_vigente.empty:
        raise BuildError("Universo TV4 vacio tras filtrar por RevisionMaestro vigente.")

    confirmado_no_existe = maestro_vigente[
        maestro_vigente["RevisionMaestro"].astype(str).str.contains("CONFIRMADO_NO_EXISTE", na=False)
    ]

    element_ids_all = maestro_vigente["ElementoID"].tolist()
    apie_map = dict(zip(maestro_vigente["ElementoID"], maestro_vigente["_APIE"]))
    token_map = dict(zip(maestro_vigente["ElementoID"], maestro_vigente["_token"]))
    ciudad_map = dict(zip(maestro_vigente["ElementoID"], maestro_vigente["Ciudad"]))
    apie_ciudad_map = dict(zip(maestro_vigente["_APIE"], maestro_vigente["Ciudad"]))

    maestro_vigente["_prov_canon"] = maestro_vigente["Observaciones"].apply(_parse_provincia_raw).apply(_canon_provincia)
    apie_provincia_groups = maestro_vigente.groupby("_APIE")["_prov_canon"].apply(lambda s: set(x for x in s if isinstance(x, str)))
    apie_provincia_conflictos = {k: sorted(v) for k, v in apie_provincia_groups.items() if len(v) > 1}
    apie_provincia_map = {k: (next(iter(v)) if len(v) == 1 else None) for k, v in apie_provincia_groups.items()}

    digital_ids = maestro_vigente.loc[maestro_vigente["_token"].isin(DIGITAL_TOKENS), "ElementoID"].tolist()
    static_ids = maestro_vigente.loc[maestro_vigente["_token"].isin(STATIC_TOKENS), "ElementoID"].tolist()

    apie_con_digital = set(maestro_vigente.loc[maestro_vigente["_token"].isin(DIGITAL_TOKENS), "_APIE"].unique())
    apie_con_estatico = set(maestro_vigente.loc[maestro_vigente["_token"].isin(STATIC_TOKENS), "_APIE"].unique())
    apie_catalogo = set(maestro_vigente["_APIE"].unique())

    estaciones_catalogo = len(apie_catalogo)
    elementos_catalogo = int(maestro_vigente["ElementoID"].nunique())
    elementos_digitales_catalogo = len(digital_ids)
    elementos_estaticos_catalogo = int(maestro_vigente.loc[maestro_vigente["_token"].isin(STATIC_TOKENS), "ElementoID"].nunique())
    estaciones_con_digital = len(apie_con_digital)

    espacios_digitales_catalogo = estaciones_con_digital * DIGITAL_CAPACITY_PER_STATION
    espacios_estaticos_catalogo = elementos_estaticos_catalogo
    espacios_totales_catalogo = espacios_digitales_catalogo + espacios_estaticos_catalogo

    return {
        "maestro": maestro_vigente,
        "element_ids_all": element_ids_all,
        "digital_ids": digital_ids,
        "static_ids": static_ids,
        "apie_map": apie_map,
        "token_map": token_map,
        "ciudad_map": ciudad_map,
        "apie_ciudad_map": apie_ciudad_map,
        "apie_provincia_map": apie_provincia_map,
        "apie_provincia_conflictos": apie_provincia_conflictos,
        "apie_con_digital": apie_con_digital,
        "apie_con_estatico": apie_con_estatico,
        "apie_catalogo": apie_catalogo,
        "estaciones_catalogo": estaciones_catalogo,
        "elementos_catalogo": elementos_catalogo,
        "elementos_digitales_catalogo": elementos_digitales_catalogo,
        "elementos_estaticos_catalogo": elementos_estaticos_catalogo,
        "estaciones_con_digital": estaciones_con_digital,
        "espacios_digitales_catalogo": espacios_digitales_catalogo,
        "espacios_estaticos_catalogo": espacios_estaticos_catalogo,
        "espacios_totales_catalogo": espacios_totales_catalogo,
        "calidad_filtro": {
            "filas_ypf_totales_pre_filtro": int(len(ypf_all)),
            "filas_excluidas_revision_nula": int(len(revision_excluidos)),
            "filas_incluidas_tag_confirmado_no_existe": int(len(confirmado_no_existe)),
            "elementos_confirmado_no_existe": sorted(confirmado_no_existe["ElementoID"].astype(str).unique().tolist()),
        },
    }


def _campanas_overlap_validas(engine: MetricsEngine, element_ids: list[Any], start: str, end: str) -> pd.DataFrame:
    """CAMPANAS con solapamiento valido contra [start,end] y IDCampana
    informado (no vacio): base comun para todas las metricas de periodo."""
    if not element_ids:
        return engine.campanas.iloc[0:0]
    overlap = engine._campanas_overlap(element_ids, start, end)
    if overlap.empty:
        return overlap
    return overlap[overlap["IDCampaña"].notna() & (overlap["IDCampaña"].astype(str).str.strip() != "")]


# ---------------------------------------------------------------------------
# Card 1 - Catalogo YPF (estructural, snapshot vigente)
# ---------------------------------------------------------------------------


def compute_catalogo(universe: dict[str, Any]) -> dict[str, Any]:
    return {
        "estaciones": universe["estaciones_catalogo"],
        "espacios": universe["espacios_totales_catalogo"],
        "elementos": universe["elementos_catalogo"],
        "espacios_digitales": universe["espacios_digitales_catalogo"],
        "espacios_estaticos": universe["espacios_estaticos_catalogo"],
        "elementos_digitales": universe["elementos_digitales_catalogo"],
        "elementos_estaticos": universe["elementos_estaticos_catalogo"],
        "estaciones_con_digital": universe["estaciones_con_digital"],
        "status": "catalogo_actual",
        "metric_status": "NO_APLICA",
        "nota": (
            "Catálogo actual (snapshot vigente). Sin histórico de altas/bajas de catálogo disponible: "
            "no se compara contra junio."
        ),
    }


# ---------------------------------------------------------------------------
# Card 2 - Ocupacion del mes (espacios, no elementos ni activaciones)
# ---------------------------------------------------------------------------


def _ocupacion_periodo(engine: MetricsEngine, universe: dict[str, Any], start: str, end: str) -> dict[str, Any]:
    apie_map = universe["apie_map"]

    overlap_dig = _campanas_overlap_validas(engine, universe["digital_ids"], start, end)
    if overlap_dig.empty:
        pares_digitales = overlap_dig.assign(_apie=[])
        counts_por_apie = pd.Series(dtype="int64")
    else:
        pares_digitales = overlap_dig.copy()
        pares_digitales["_apie"] = pares_digitales["ElementoID"].map(apie_map)
        pares_digitales = pares_digitales.drop_duplicates(subset=["_apie", "IDCampaña"])
        counts_por_apie = pares_digitales.groupby("_apie").size()

    ocupados_digitales = int(len(pares_digitales))

    overlap_static = _campanas_overlap_validas(engine, universe["static_ids"], start, end)
    ocupados_estaticos = int(overlap_static["ElementoID"].nunique()) if not overlap_static.empty else 0

    ocupados_totales = ocupados_digitales + ocupados_estaticos
    espacios_totales = universe["espacios_totales_catalogo"]
    ocupacion_pct = _round1(ocupados_totales / espacios_totales * 100.0) if espacios_totales else None

    sobre_capacidad = counts_por_apie[counts_por_apie > DIGITAL_CAPACITY_PER_STATION]
    estaciones_sobre_capacidad = int(len(sobre_capacidad))
    exceso_sobre_capacidad = int((sobre_capacidad - DIGITAL_CAPACITY_PER_STATION).sum()) if len(sobre_capacidad) else 0
    max_pct_estacion = _round1(float(counts_por_apie.max()) / DIGITAL_CAPACITY_PER_STATION * 100.0) if len(counts_por_apie) else None

    return {
        "ocupados_digitales": ocupados_digitales,
        "ocupados_estaticos": ocupados_estaticos,
        "ocupados_totales": ocupados_totales,
        "ocupacion_pct": ocupacion_pct,
        "estaciones_sobre_capacidad": estaciones_sobre_capacidad,
        "exceso_sobre_capacidad": exceso_sobre_capacidad,
        "max_pct_estacion": max_pct_estacion,
        "conteo_por_estacion": counts_por_apie,
    }


def compute_ocupacion(universe: dict[str, Any], engine: MetricsEngine, period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    actual = _ocupacion_periodo(engine, universe, period[0], period[1])
    anterior = _ocupacion_periodo(engine, universe, previous[0], previous[1])

    delta_abs = actual["ocupados_totales"] - anterior["ocupados_totales"]
    delta_pp = (
        _round1(actual["ocupacion_pct"] - anterior["ocupacion_pct"])
        if actual["ocupacion_pct"] is not None and anterior["ocupacion_pct"] is not None
        else None
    )

    return {
        "actual": {k: v for k, v in actual.items() if k != "conteo_por_estacion"},
        "anterior": {k: v for k, v in anterior.items() if k != "conteo_por_estacion"},
        "delta_abs": delta_abs,
        "delta_pp": delta_pp,
        "espacios_totales_catalogo": universe["espacios_totales_catalogo"],
    }


# ---------------------------------------------------------------------------
# Card 3 - Campanas / Activaciones (todo el universo YPF: digital + estatico)
# ---------------------------------------------------------------------------


def compute_campanas(engine: MetricsEngine, universe: dict[str, Any], period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    element_ids = universe["element_ids_all"]

    overlap_actual = _campanas_overlap_validas(engine, element_ids, period[0], period[1])
    overlap_anterior = _campanas_overlap_validas(engine, element_ids, previous[0], previous[1])

    campanas_actual = int(overlap_actual["IDCampaña"].nunique()) if not overlap_actual.empty else 0
    campanas_anterior = int(overlap_anterior["IDCampaña"].nunique()) if not overlap_anterior.empty else 0
    activaciones_actual = int(overlap_actual.drop_duplicates(subset=["IDCampaña", "ElementoID"]).shape[0]) if not overlap_actual.empty else 0
    activaciones_anterior = int(overlap_anterior.drop_duplicates(subset=["IDCampaña", "ElementoID"]).shape[0]) if not overlap_anterior.empty else 0

    promedio_actual = _round1(activaciones_actual / campanas_actual) if campanas_actual else None
    promedio_anterior = _round1(activaciones_anterior / campanas_anterior) if campanas_anterior else None

    return {
        "campanas_unicas_actual": campanas_actual,
        "campanas_unicas_anterior": campanas_anterior,
        "delta_campanas": campanas_actual - campanas_anterior,
        "activaciones_actual": activaciones_actual,
        "activaciones_anterior": activaciones_anterior,
        "delta_activaciones": activaciones_actual - activaciones_anterior,
        "promedio_actual": promedio_actual,
        "promedio_anterior": promedio_anterior,
    }


# ---------------------------------------------------------------------------
# Card 4 - Elementos activos (composicion por formato, reconcilia 100%)
# ---------------------------------------------------------------------------


def compute_elementos_activos(engine: MetricsEngine, universe: dict[str, Any], period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    element_ids = universe["element_ids_all"]
    token_map = universe["token_map"]

    def _one(start: str, end: str) -> dict[str, Any]:
        overlap = _campanas_overlap_validas(engine, element_ids, start, end)
        if overlap.empty:
            return {"total": 0, "por_formato": {t: 0 for t in list(TOKEN_LABELS)}}
        tagged = overlap.drop_duplicates(subset=["ElementoID"]).copy()
        tagged["_token"] = tagged["ElementoID"].map(token_map)
        total = int(tagged["ElementoID"].nunique())
        counts = tagged.groupby("_token")["ElementoID"].nunique().to_dict()
        return {"total": total, "por_formato": {t: int(counts.get(t, 0)) for t in TOKEN_LABELS}}

    actual = _one(period[0], period[1])
    anterior = _one(previous[0], previous[1])

    composicion = {}
    for token, label in TOKEN_LABELS.items():
        cnt = actual["por_formato"][token]
        pct = _round1(cnt / actual["total"] * 100.0) if actual["total"] else None
        composicion[label] = {"token": token, "elementos": cnt, "pct": pct}

    return {
        "actual": actual["total"],
        "anterior": anterior["total"],
        "delta": actual["total"] - anterior["total"],
        "composicion_actual": composicion,
        "composicion_anterior_total": anterior["total"],
    }


# ---------------------------------------------------------------------------
# Card 5 - Formato lider (por campanas unicas; Fotobox solo si tiene campanas)
# ---------------------------------------------------------------------------


_CANONICAL_ORDER = ["MB", "TT", "PPUNTER", "FB"]


def _formato_campanas_unicas(overlap: pd.DataFrame, token_map: dict[Any, str]) -> dict[str, int]:
    if overlap.empty:
        return {}
    tagged = overlap.copy()
    tagged["_token"] = tagged["ElementoID"].map(token_map)
    counts = tagged.groupby("_token")["IDCampaña"].nunique().to_dict()
    return {k: int(v) for k, v in counts.items() if v > 0}


def _formato_elementos_activos(overlap: pd.DataFrame, token_map: dict[Any, str]) -> dict[str, int]:
    if overlap.empty:
        return {}
    tagged = overlap.copy()
    tagged["_token"] = tagged["ElementoID"].map(token_map)
    counts = tagged.groupby("_token")["ElementoID"].nunique().to_dict()
    return {k: int(v) for k, v in counts.items()}


def _pick_formato_lider(campanas_unicas: dict[str, int], elementos_activos: dict[str, int]) -> dict[str, Any]:
    max_val = max(campanas_unicas.values()) if campanas_unicas else 0
    empatados = [f for f, v in campanas_unicas.items() if v == max_val and max_val > 0]
    empate = len(empatados) > 1
    if not empatados:
        ganador = None
    elif empate:
        def _key(f: str) -> tuple[int, int]:
            return (-elementos_activos.get(f, 0), _CANONICAL_ORDER.index(f) if f in _CANONICAL_ORDER else 99)
        ganador = sorted(empatados, key=_key)[0]
    else:
        ganador = empatados[0]
    return {
        "formato_token": ganador,
        "label": TOKEN_LABELS.get(ganador, "Sin dato") if ganador else "Sin dato",
        "campanas_unicas": campanas_unicas.get(ganador, 0) if ganador else 0,
        "elementos_activos": elementos_activos.get(ganador, 0) if ganador else 0,
        "empate_resuelto_por_elementos_activos": empate,
        "detalle_campanas_unicas": {TOKEN_LABELS[k]: v for k, v in campanas_unicas.items()},
    }


def compute_formato_lider(engine: MetricsEngine, universe: dict[str, Any], period: tuple[str, str], previous: tuple[str, str]) -> dict[str, Any]:
    element_ids = universe["element_ids_all"]
    token_map = universe["token_map"]

    overlap_actual = _campanas_overlap_validas(engine, element_ids, period[0], period[1])
    overlap_anterior = _campanas_overlap_validas(engine, element_ids, previous[0], previous[1])

    camp_actual = _formato_campanas_unicas(overlap_actual, token_map)
    camp_anterior = _formato_campanas_unicas(overlap_anterior, token_map)
    elem_actual = _formato_elementos_activos(overlap_actual, token_map)
    elem_anterior = _formato_elementos_activos(overlap_anterior, token_map)

    return {
        "actual": _pick_formato_lider(camp_actual, elem_actual),
        "anterior": _pick_formato_lider(camp_anterior, elem_anterior),
    }


# ---------------------------------------------------------------------------
# Panel izquierdo - Evolucion mensual de la ocupacion (Ene-Jul, 3 series)
# ---------------------------------------------------------------------------


def compute_historico(engine: MetricsEngine, universe: dict[str, Any], report_year: int, report_month: int) -> dict[str, Any]:
    meses = MESES_ES_ABR[:report_month]
    cap_dig = universe["espacios_digitales_catalogo"]
    cap_static = universe["espacios_estaticos_catalogo"]
    cap_total = universe["espacios_totales_catalogo"]

    serie_total: list[dict[str, Any]] = []
    serie_digital: list[dict[str, Any]] = []
    serie_static: list[dict[str, Any]] = []

    for m in range(1, report_month + 1):
        start, end = _period_bounds(report_year, m)
        ocup = _ocupacion_periodo(engine, universe, start, end)
        pct_dig = _round1(ocup["ocupados_digitales"] / cap_dig * 100.0) if cap_dig else None
        pct_static = _round1(ocup["ocupados_estaticos"] / cap_static * 100.0) if cap_static else None
        pct_total = _round1(ocup["ocupados_totales"] / cap_total * 100.0) if cap_total else None

        serie_digital.append({"ocupados": ocup["ocupados_digitales"], "capacidad": cap_dig, "pct": pct_dig})
        serie_static.append({"ocupados": ocup["ocupados_estaticos"], "capacidad": cap_static, "pct": pct_static})
        serie_total.append({"ocupados": ocup["ocupados_totales"], "capacidad": cap_total, "pct": pct_total})

    return {
        "meses": meses,
        "series": {"Total": serie_total, "Digital": serie_digital, "Estático": serie_static},
        "denominador_constante": True,
        "nota_denominador": (
            f"Denominador constante = catálogo vigente ({_fmt_es_int(cap_total)} espacios: "
            f"{_fmt_es_int(cap_dig)} digitales + {_fmt_es_int(cap_static)} estáticos). "
            "Sin histórico de catálogo por mes: no se reconstruyen altas/bajas pasadas."
        ),
        "leyenda_capacidad": (
            "Capacidad digital: 5 espacios por estación. Si una estación supera 5 campañas activas, "
            "todas se contabilizan y la ocupación puede superar el 100%."
        ),
    }


# ---------------------------------------------------------------------------
# Mapa de intensidad comercial (reutiliza build_tv5_dashboard.load_geo_aux)
# ---------------------------------------------------------------------------


def compute_mapa(engine: MetricsEngine, universe: dict[str, Any], report_year: int, report_month: int, geo_map: dict[str, dict[str, Any]]) -> dict[str, Any]:
    element_ids = universe["element_ids_all"]
    apie_map = universe["apie_map"]
    apie_ciudad_map = universe["apie_ciudad_map"]
    apie_provincia_map = universe["apie_provincia_map"]

    activ_frames = []
    for m in range(1, report_month + 1):
        start, end = _period_bounds(report_year, m)
        overlap = _campanas_overlap_validas(engine, element_ids, start, end)
        if overlap.empty:
            continue
        activ = overlap.drop_duplicates(subset=["IDCampaña", "ElementoID"]).copy()
        activ["_apie"] = activ["ElementoID"].map(apie_map)
        activ_frames.append(activ[["IDCampaña", "ElementoID", "_apie"]])

    if not activ_frames:
        acumulado = pd.DataFrame(columns=["IDCampaña", "ElementoID", "_apie"])
    else:
        acumulado = pd.concat(activ_frames, ignore_index=True)
    acumulado = acumulado[acumulado["_apie"].notna()].copy()

    by_station = acumulado.groupby("_apie").size()
    n_estaciones_con_actividad = int(by_station.shape[0])
    active_apies = set(by_station.index)

    puntos = []
    con_geo: set[Any] = set()
    for apie, activaciones in by_station.items():
        geo = geo_map.get(str(apie))
        if geo is None:
            continue
        con_geo.add(apie)
        puntos.append({
            "apie": apie,
            "lat": geo["lat"],
            "lon": geo["lon"],
            "activaciones": int(activaciones),
            "ciudad": geo.get("ciudad") or apie_ciudad_map.get(apie),
            "amba": _en_amba(geo["lat"], geo["lon"]),
        })

    n_con_geo = len(puntos)
    puntos_amba = [p for p in puntos if p["amba"]]

    # Clasificacion CABA / Zona Norte por cruce punto-en-poligono real
    # (comunas GCBA + partidos IGN, ver clasificar_caba_zona_norte). Se
    # anota en cada punto AMBA para el detalle de mapa, nunca se infiere.
    caba_zn = clasificar_caba_zona_norte(puntos_amba)
    for p in puntos_amba:
        info = caba_zn["detalle"][p["apie"]]
        p["comuna"] = info["comuna"]
        p["partido"] = info["partido"]
        p["zona"] = info["zona"]

    # Marcador agregado real (no un centroide provincial disfrazado de
    # estacion): promedio de las coordenadas REALES de las estaciones AMBA,
    # ponderado por activaciones, usado solo como ancla visual "hay zoom
    # aca" en el mapa nacional -- nunca se dibuja como un punto de estacion.
    amba_marker = None
    if puntos_amba:
        total_activ = sum(p["activaciones"] for p in puntos_amba) or 1
        amba_marker = {
            "lat": sum(p["lat"] * p["activaciones"] for p in puntos_amba) / total_activ,
            "lon": sum(p["lon"] * p["activaciones"] for p in puntos_amba) / total_activ,
            "estaciones": len(puntos_amba),
            "activaciones": sum(p["activaciones"] for p in puntos_amba),
        }
    pct_cobertura = _round1(n_con_geo / n_estaciones_con_actividad * 100.0) if n_estaciones_con_actividad else None

    n_apie_con_geo_catalogo = sum(1 for apie in universe["apie_catalogo"] if str(apie) in geo_map)
    pct_cobertura_catalogo = _round1(n_apie_con_geo_catalogo / universe["estaciones_catalogo"] * 100.0) if universe["estaciones_catalogo"] else None

    # Cobertura provincial (Sec.4 del ajuste de mapa): estaciones activas con
    # provincia canonica resuelta (apie_provincia_map) pero SIN coordenada
    # puntual valida. No se coloca un punto ficticio: se representa la
    # provincia como territorio (ver template JS drawMapa).
    con_provincia = {a for a in active_apies if apie_provincia_map.get(a)}
    solo_provincia = con_provincia - con_geo
    sin_ubicacion = active_apies - con_geo - con_provincia

    provincias: list[dict[str, Any]] = []
    provincias_presentes = {apie_provincia_map[a] for a in con_provincia}
    for nombre in sorted(provincias_presentes):
        apies_prov = {a for a in con_provincia if apie_provincia_map.get(a) == nombre}
        rows_prov = acumulado[acumulado["_apie"].isin(apies_prov)]
        provincias.append({
            "nombre": nombre,
            "estaciones_activas": len(apies_prov),
            "estaciones_con_punto": len(apies_prov & con_geo),
            "estaciones_solo_provincia": len(apies_prov - con_geo),
            "campanas_unicas": int(rows_prov["IDCampaña"].nunique()),
            "activaciones": int(rows_prov.drop_duplicates(subset=["IDCampaña", "ElementoID"]).shape[0]),
        })
    provincias.sort(key=lambda p: p["activaciones"], reverse=True)

    bounds = None
    if puntos:
        lats = [p["lat"] for p in puntos]
        lons = [p["lon"] for p in puntos]
        bounds = {"lat_min": min(lats), "lat_max": max(lats), "lon_min": min(lons), "lon_max": max(lons)}

    return {
        "titulo": "Mapa de intensidad comercial",
        "subtitulo": "Actividad Ene–Jul · puntos georreferenciados y cobertura por provincia",
        "periodo_label": f"Ene–{MESES_ES_ABR[report_month-1]} {report_year}",
        "estado": "PARTIAL",
        "metric_status": "PARTIAL",
        "nota": (
            f"{_fmt_es_int(n_con_geo)} de {_fmt_es_int(n_estaciones_con_actividad)} estaciones con actividad "
            f"Ene–{MESES_ES_ABR[report_month-1]} cuentan con coordenada puntual válida (join por APIE contra la "
            f"fuente auxiliar YPF_GEO_COORDENADAS); {_fmt_es_int(len(solo_provincia))} más tienen provincia "
            f"identificada (texto libre en el maestro) y se representan a nivel territorio; "
            f"{_fmt_es_int(len(sin_ubicacion))} no tienen ninguna ubicación utilizable en la fuente. "
            "No se inventan ni corrigen coordenadas por suposición."
        ),
        "nota_breve": "Puntos reales y cobertura provincial; no se inventan coordenadas.",
        "leyenda_min": "Menor actividad",
        "leyenda_max": "Mayor actividad",
        "puntos": puntos,
        "bounds": bounds,
        "provincias": provincias,
        "n_estaciones_con_actividad": n_estaciones_con_actividad,
        "n_estaciones_con_geo": n_con_geo,
        "n_estaciones_solo_provincia": len(solo_provincia),
        "n_estaciones_sin_ubicacion": len(sin_ubicacion),
        "pct_cobertura_geo": pct_cobertura,
        "n_apie_catalogo_con_geo": n_apie_con_geo_catalogo,
        "pct_cobertura_geo_catalogo": pct_cobertura_catalogo,
        "amba_bounds": AMBA_BOUNDS,
        "amba_marker": amba_marker,
        "n_puntos_amba": len(puntos_amba),
        "caba_zona_norte": {
            "n_caba": caba_zn["n_caba"],
            "n_zona_norte": caba_zn["n_zona_norte"],
            "zona_norte_partidos": caba_zn["zona_norte_partidos"],
        },
    }


# ---------------------------------------------------------------------------
# Calidad de datos y reconciliacion contra el control preliminar conocido
# ---------------------------------------------------------------------------


CONTROL_PRELIMINAR_CONOCIDO = {
    "estaciones": 603,
    "elementos": 3883,
    "elementos_digitales": 3500,
    "elementos_estaticos": 383,
    "estaciones_con_digital": 412,
    "campanas": 27,
    "asignaciones_base_campanas": 13616,
    "espacios_digitales": 2060,
    "espacios_estaticos": 383,
    "espacios_totales": 2443,
}


def compute_reconciliacion(universe: dict[str, Any], asignaciones_base_campanas: int, campanas_all_time: int) -> dict[str, Any]:
    calculado = {
        "estaciones": universe["estaciones_catalogo"],
        "elementos": universe["elementos_catalogo"],
        "elementos_digitales": universe["elementos_digitales_catalogo"],
        "elementos_estaticos": universe["elementos_estaticos_catalogo"],
        "estaciones_con_digital": universe["estaciones_con_digital"],
        "campanas": campanas_all_time,
        "asignaciones_base_campanas": asignaciones_base_campanas,
        "espacios_digitales": universe["espacios_digitales_catalogo"],
        "espacios_estaticos": universe["espacios_estaticos_catalogo"],
        "espacios_totales": universe["espacios_totales_catalogo"],
    }
    filas = []
    for k, control in CONTROL_PRELIMINAR_CONOCIDO.items():
        calc = calculado[k]
        filas.append({
            "metrica": k,
            "control_preliminar_conocido": control,
            "calculado_vigente": calc,
            "coincide": bool(calc == control),
            "diferencia": calc - control,
        })
    return {
        "filas": filas,
        "nota": (
            "El control preliminar conocido (documento previo) reconcilia EXACTO contra la base vigente "
            "(input/OCU26_BASE_DATOS.xlsx, MAESTRO_ELEMENTOS+CAMPANAS filtrado a CircuitoNegocio=YPF con "
            "RevisionMaestro no nulo) en 9 de 10 métricas. La única diferencia es 'estaciones': "
            f"{universe['estaciones_catalogo']} calculado (APIE=Subcircuito distinto) vs 603 del control "
            "preliminar. La base vigente no permite reconstruir 603 estaciones: el roster crudo "
            "'YPF DIGITAL + ESTATICO' (BASE ESTACIONES) trae 525 APIE distintos, y el maestro operativo "
            "vigente (tras excluir 165 filas sin sello de revisión) trae 521. No se fuerza 603: se reporta "
            "el valor recalculado y se deja esta diferencia como pendiente de aclaración con el dueño del dato."
        ),
    }


def compute_calidad(universe: dict[str, Any], geo_calidad: dict[str, Any], mapa: dict[str, Any], reconciliacion: dict[str, Any]) -> dict[str, Any]:
    cf = universe["calidad_filtro"]
    avisos = [
        (
            f"{_fmt_es_int(cf['filas_excluidas_revision_nula'])} elemento(s) YPF del snapshot crudo no tienen "
            "sello de revisión (RevisionMaestro nulo): se excluyeron del catálogo vigente por ser remanentes "
            "previos a la corrección estructural del 13/08/2026."
        ),
        (
            f"{_fmt_es_int(cf['filas_incluidas_tag_confirmado_no_existe'])} elemento(s) están tageados "
            "'EXCLUIDO - CONFIRMADO_NO_EXISTE' en RevisionMaestro pero SÍ se incluyeron en el catálogo vigente "
            "(ese filtro es el que reconcilia contra el control preliminar conocido). Requiere aclaración: si "
            "esos elementos están físicamente confirmados como inexistentes, el catálogo real sería menor."
        ),
    ]
    for row in reconciliacion["filas"]:
        if not row["coincide"]:
            avisos.append(
                f"Métrica '{row['metrica']}': calculado {_fmt_es_int(row['calculado_vigente'])} vs control "
                f"preliminar conocido {_fmt_es_int(row['control_preliminar_conocido'])} "
                f"(diferencia {_signed_int(row['diferencia'])})."
            )
    if mapa["n_estaciones_sin_ubicacion"] > 0:
        avisos.append(
            f"{_fmt_es_int(mapa['n_estaciones_sin_ubicacion'])} estación(es) con actividad Ene–Jul no tienen "
            "coordenada válida NI provincia identificable (ni en la fuente geográfica auxiliar ni en el texto "
            "libre de Observaciones): quedan fuera del mapa (puntual y territorial)."
        )
    conflictos = universe.get("apie_provincia_conflictos") or {}
    if conflictos:
        avisos.append(
            f"{_fmt_es_int(len(conflictos))} estación(es) tienen texto de provincia inconsistente entre sus "
            "propios elementos (más de un nombre de provincia distinto): no se resolvió por adivinanza, "
            "quedan sin provincia asignada."
        )
    return {
        "avisos": avisos,
        "geo_aux": {
            **geo_calidad,
            "estaciones_con_actividad_ene_jul": mapa["n_estaciones_con_actividad"],
            "estaciones_con_actividad_con_geo_valida": mapa["n_estaciones_con_geo"],
            "pct_cobertura_geo_sobre_activas": mapa["pct_cobertura_geo"],
            "pct_cobertura_geo_sobre_catalogo": mapa["pct_cobertura_geo_catalogo"],
        },
    }


# ---------------------------------------------------------------------------
# Insights (Lectura / Punto positivo / A atender)
# ---------------------------------------------------------------------------


def compute_insights(
    catalogo: dict[str, Any], ocupacion: dict[str, Any], campanas: dict[str, Any],
    elementos: dict[str, Any], formato_lider: dict[str, Any], mapa: dict[str, Any],
    report_month_label: str, previous_month_label: str,
) -> dict[str, str]:
    fl_actual = formato_lider["actual"]
    fl_anterior = formato_lider["anterior"]
    oc_actual = ocupacion["actual"]
    oc_anterior = ocupacion["anterior"]

    lectura = (
        f"{report_month_label} registra un catálogo YPF de <b>{_fmt_es_int(catalogo['estaciones'])} estaciones</b> "
        f"y <b>{_fmt_es_int(catalogo['espacios'])} espacios comerciales</b> ({_fmt_es_int(catalogo['elementos'])} "
        f"elementos físicos), con <b>{_fmt_es_int(oc_actual['ocupados_totales'])} espacios ocupados</b> "
        f"({_fmt_es_pct(oc_actual['ocupacion_pct'])}% de ocupación), <b>{_fmt_es_int(campanas['campanas_unicas_actual'])} "
        f"campañas únicas</b> y <b>{_fmt_es_int(campanas['activaciones_actual'])} activaciones</b> sobre "
        f"{_fmt_es_int(elementos['actual'])} elementos activos. El formato líder es <b>{fl_actual['label']}</b> "
        f"({_fmt_es_int(fl_actual['campanas_unicas'])} campañas únicas), frente a {previous_month_label.lower()} "
        f"({_fmt_es_int(oc_anterior['ocupados_totales'])} espacios ocupados, "
        f"{_fmt_es_pct(oc_anterior['ocupacion_pct'])}%, formato líder {fl_anterior['label']})."
    )

    if ocupacion["delta_abs"] > 0:
        punto_positivo = (
            f"Los espacios ocupados crecieron de <b>{_fmt_es_int(oc_anterior['ocupados_totales'])} a "
            f"{_fmt_es_int(oc_actual['ocupados_totales'])}</b> ({_signed_int(ocupacion['delta_abs'])}), con una "
            f"ocupación que pasó de {_fmt_es_pct(oc_anterior['ocupacion_pct'])}% a "
            f"{_fmt_es_pct(oc_actual['ocupacion_pct'])}% ({_signed_pp(ocupacion['delta_pp'])})."
        )
    elif campanas["delta_campanas"] > 0:
        punto_positivo = (
            f"Las campañas únicas crecieron de {_fmt_es_int(campanas['campanas_unicas_anterior'])} a "
            f"<b>{_fmt_es_int(campanas['campanas_unicas_actual'])}</b> respecto a {previous_month_label.lower()}."
        )
    elif elementos["delta"] > 0:
        punto_positivo = (
            f"Los elementos activos crecieron de {_fmt_es_int(elementos['anterior'])} a "
            f"<b>{_fmt_es_int(elementos['actual'])}</b> respecto a {previous_month_label.lower()}."
        )
    else:
        punto_positivo = (
            f"La red YPF mantiene <b>{_fmt_es_int(oc_actual['ocupados_totales'])} espacios ocupados</b> "
            f"sobre un catálogo de {_fmt_es_int(catalogo['espacios'])}."
        )

    if oc_actual["estaciones_sobre_capacidad"] > 0:
        a_atender = (
            f"<b>{_fmt_es_int(oc_actual['estaciones_sobre_capacidad'])} estaciones</b> superan la capacidad de "
            f"5 espacios digitales en {report_month_label.lower()} (exceso total: "
            f"{_fmt_es_int(oc_actual['exceso_sobre_capacidad'])} espacios; máximo observado: "
            f"{_fmt_es_pct(oc_actual['max_pct_estacion'])}% en una estación): esas estaciones concentran más "
            "campañas digitales que espacios físicos disponibles."
        )
    elif mapa["pct_cobertura_geo"] is not None and mapa["pct_cobertura_geo"] < 70:
        a_atender = (
            f"El mapa de intensidad comercial solo geolocaliza <b>{_fmt_es_pct(mapa['pct_cobertura_geo'])}%</b> "
            f"de las estaciones con actividad ({_fmt_es_int(mapa['n_estaciones_con_geo'])} de "
            f"{_fmt_es_int(mapa['n_estaciones_con_actividad'])}): falta ampliar la cobertura de la fuente "
            f"geográfica auxiliar sobre el resto de la red."
        )
    elif oc_actual["ocupacion_pct"] is not None and oc_actual["ocupacion_pct"] < 70:
        a_atender = (
            f"Solo el {_fmt_es_pct(oc_actual['ocupacion_pct'])}% del catálogo de espacios está ocupado en "
            f"{report_month_label.lower()}: queda disponibilidad comercial sin vender sobre la red YPF."
        )
    else:
        a_atender = (
            "El catálogo recalculado de estaciones (" + _fmt_es_int(catalogo["estaciones"]) + ") no reconcilia "
            "1:1 contra el control preliminar conocido (603): requiere aclaración con el dueño del dato sobre "
            "las estaciones del roster crudo que aún no están reflejadas en el maestro operativo vigente."
        )

    return {"lectura": lectura, "punto_positivo": punto_positivo, "a_atender": a_atender}


# ---------------------------------------------------------------------------
# Logo (reutilizado byte-a-byte, igual patron que TV1-6)
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

    _transform_result, semantic_result, engine = load_pipeline(path)
    universe = build_tv4_universe(semantic_result)

    period = _period_bounds(REPORT_YEAR, REPORT_MONTH)
    prev_year, prev_month = _previous_month(REPORT_YEAR, REPORT_MONTH)
    previous = _period_bounds(prev_year, prev_month)

    catalogo = compute_catalogo(universe)
    ocupacion = compute_ocupacion(universe, engine, period, previous)
    campanas = compute_campanas(engine, universe, period, previous)
    elementos = compute_elementos_activos(engine, universe, period, previous)
    formato_lider = compute_formato_lider(engine, universe, period, previous)
    historico = compute_historico(engine, universe, REPORT_YEAR, REPORT_MONTH)
    geo_map, geo_calidad = load_geo_aux()
    mapa = compute_mapa(engine, universe, REPORT_YEAR, REPORT_MONTH, geo_map)

    asignaciones_base_campanas = int(
        engine.campanas[engine.campanas["ElementoID"].isin(universe["element_ids_all"])].shape[0]
    )
    campanas_all_time = int(
        engine.campanas.loc[
            engine.campanas["ElementoID"].isin(universe["element_ids_all"]), "IDCampaña"
        ].nunique()
    )
    reconciliacion = compute_reconciliacion(universe, asignaciones_base_campanas, campanas_all_time)
    calidad = compute_calidad(universe, geo_calidad, mapa, reconciliacion)

    report_month_label = MESES_ES[REPORT_MONTH - 1]
    previous_month_label = MESES_ES[prev_month - 1]
    insights = compute_insights(catalogo, ocupacion, campanas, elementos, formato_lider, mapa, report_month_label, previous_month_label)

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
            "report_month_label": report_month_label,
            "previous_month": prev_month,
            "previous_month_label": previous_month_label,
            "periodo_inicio_iso": period[0],
            "periodo_fin_iso": period[1],
            "previous_inicio_iso": previous[0],
            "previous_fin_iso": previous[1],
            "hist_label": f"Ene–{MESES_ES_ABR[REPORT_MONTH-1]} {REPORT_YEAR}",
            "fuente": "OCU26 · Base maestra + base campañas (CircuitoNegocio=YPF, RevisionMaestro vigente)",
            "apie_method": "APIE = Subcircuito (código real de estación de la fuente, no un surrogate).",
        },
        "universo": {
            "circuito": "YPF",
            "estaciones_catalogo": universe["estaciones_catalogo"],
            "elementos_catalogo": universe["elementos_catalogo"],
            "espacios_totales_catalogo": universe["espacios_totales_catalogo"],
        },
        "kpis": {
            "catalogo": catalogo,
            "ocupacion": ocupacion,
            "campanas": campanas,
            "elementos_activos": elementos,
            "formato_lider": formato_lider,
        },
        "historico": historico,
        "mapa": mapa,
        "calidad": calidad,
        "reconciliacion": reconciliacion,
        "insights": insights,
    }

    return {
        "data": data, "sha256": sha_after,
        "universe": {
            "circuito": "YPF",
            "estaciones_catalogo": universe["estaciones_catalogo"],
            "elementos_catalogo": universe["elementos_catalogo"],
        },
    }


def render_html(data: dict[str, Any]) -> str:
    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    logo_tag = extract_logo_img_tag(REFERENCE_PATH)
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    if not ARG_PAIS_GEOJSON_PATH.exists() or not ARG_PROVINCIAS_GEOJSON_PATH.exists():
        raise BuildError(
            f"Cartografia real de Argentina no encontrada en {ARG_PAIS_GEOJSON_PATH} / "
            f"{ARG_PROVINCIAS_GEOJSON_PATH}."
        )
    if not CABA_COMUNAS_PATH.exists() or not GBA_PARTIDOS_PATH.exists():
        raise BuildError(
            f"Geometria politica AMBA no encontrada en {CABA_COMUNAS_PATH} / {GBA_PARTIDOS_PATH}."
        )
    arg_pais_geojson = ARG_PAIS_GEOJSON_PATH.read_text(encoding="utf-8")
    arg_provincias_geojson = ARG_PROVINCIAS_GEOJSON_PATH.read_text(encoding="utf-8")
    caba_comunas_geojson = CABA_COMUNAS_PATH.read_text(encoding="utf-8")
    gba_partidos_geojson = GBA_PARTIDOS_PATH.read_text(encoding="utf-8")
    html = template.replace("{{LOGO_IMG_TAG}}", logo_tag)
    html = html.replace("{{TV4_DATA_JSON}}", payload)
    html = html.replace("{{ARG_PAIS_GEOJSON}}", arg_pais_geojson)
    html = html.replace("{{ARG_PROVINCIAS_GEOJSON}}", arg_provincias_geojson)
    html = html.replace("{{CABA_COMUNAS_GEOJSON}}", caba_comunas_geojson)
    html = html.replace("{{GBA_PARTIDOS_GEOJSON}}", gba_partidos_geojson)
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
    parser = argparse.ArgumentParser(description="Construye tv4.html (dashboard TV4 OCU26, Pulso comercial YPF) con datos reales.")
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
