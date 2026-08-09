"""Capa de salida determinista OCU26 - Gate 4A.

Se ejecuta DESPUES de scripts/semantic_model.py y scripts/metrics_engine.py
(Gate 3B). Unica fuente de datos: el dict devuelto por
semantic_model.build_semantic_model() (que a su vez viene de
transform_data()). No reabre el Excel salvo para el control de SHA-256
read-only ya usado por Gate 1/2, no reimplementa ninguna regla de negocio de
Gate 3: reutiliza sus funciones internas (MetricsEngine._campanas_overlap,
MetricsEngine._incomplete_date_elements, metrics_engine._explode_days) para
garantizar fidelidad exacta con Gate 3, en vez de reescribir la logica de
solapamiento de fechas / exclusion de politica / capacidad desconocida.

Genera 5 tablas Parquet + un manifest de auditoria en output/ (ignorado por
git, 100% regenerable a partir del Excel fuente):

    DIM_ELEMENTOS        - 1 fila / ElementoID
    FACT_CAMPANAS        - 1 fila / CargaID
    DIM_CALENDARIO       - 1 fila / dia
    BRIDGE_CAMPANA_DIA   - 1 fila / (CargaID, Fecha) vigente
    FACT_METRICAS_DIARIA - 1 fila / (ElementoID, Fecha) con actividad (WIDE, sparse)

100% en memoria hasta el momento de escribir Parquet. Estrictamente
read-only sobre el Excel fuente: verifica SHA-256 antes y despues, igual que
Gate 2.

Uso:
    python scripts/export_data.py
    python scripts/export_data.py --file ruta/al/archivo.xlsx --output-dir output
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
from transform_data import transform_data  # noqa: E402
import semantic_model as sm  # noqa: E402
from metrics_engine import MetricsEngine, UnsupportedBusinessCaseError, _explode_days  # noqa: E402

DEFAULT_INPUT_PATH = vi.DEFAULT_INPUT_PATH
DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parent.parent / "output"

SCHEMA_VERSION = "gate4a-v1"

MESES_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]
DIAS_ES = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado", "Domingo"]


class ExportError(Exception):
    """Error bloqueante de Gate 4A: invariante de exportacion rota."""


def _sanear_columnas_texto_mixto(df: pd.DataFrame) -> pd.DataFrame:
    """Parquet (a diferencia de Excel/pandas 'object') no admite una columna
    que mezcle str/int/float en la misma columna. Gate 1/2 deliberadamente
    no normaliza estas columnas (son texto/dimension libre del Excel, fuera
    de las 4 columnas objetivo de Gate 2) - Gate 4 tampoco reinterpreta su
    contenido, solo fija una representacion fisica unica (texto, pd.NA
    preservado) para poder escribirlas. Generico: no depende de una lista
    fija de nombres de columna, absorbe cualquier columna nueva con el mismo
    problema sin cambios de codigo."""
    df = df.copy()
    for columna in df.columns:
        if df[columna].dtype != object:
            continue
        tipos = {type(v).__name__ for v in df[columna] if not pd.isna(v)}
        if len(tipos) > 1:
            df[columna] = df[columna].apply(lambda v: v if pd.isna(v) else str(v)).astype("string")
    return df


# ---------------------------------------------------------------------------
# Carga del pipeline (Gate 1 -> Gate 2 -> Gate 3B), sin reabrir ni reinterpretar
# ---------------------------------------------------------------------------


def load_pipeline(path: str | Path = DEFAULT_INPUT_PATH) -> tuple[dict[str, Any], dict[str, Any], MetricsEngine]:
    """Ejecuta transform_data() + build_semantic_model() y arma el MetricsEngine.

    No reimplementa nada de Gate 1/2/3: es exactamente la misma cadena de
    llamadas que ya usan los tests de Gate 3B.
    """
    transform_result = transform_data(path)
    semantic_result = sm.build_semantic_model(transform_result)
    engine = MetricsEngine(semantic_result)
    return transform_result, semantic_result, engine


# ---------------------------------------------------------------------------
# DIM_ELEMENTOS
# ---------------------------------------------------------------------------


def _capacidad_bi_safe(serie: pd.Series, sentinel: str) -> tuple[pd.Series, pd.Series]:
    """Separa una columna de capacidad (numero, pd.NA, o el sentinel de texto
    REQUIERE_CONFIRMACION) en (ValorNumerico nullable, FlagDesconocida bool).

    Reutiliza el mismo patron ya usado dentro de metrics_engine.py
    (`pd.to_numeric(serie.where(~es_texto), errors="coerce")`) para no
    inventar una conversion nueva.
    """
    es_texto = serie.apply(lambda v: isinstance(v, str))
    desconocida = serie.apply(lambda v: isinstance(v, str) and v == sentinel)
    valor = pd.to_numeric(serie.where(~es_texto), errors="coerce")
    return valor, desconocida.astype(bool)


def build_dim_elementos(semantic_result: dict[str, Any], engine: MetricsEngine) -> pd.DataFrame:
    """1 fila por ElementoID. ElementoID debe ser PK unica: si aparece un
    duplicado, la exportacion falla explicito (no se fusiona, no first()).

    Agrega, ademas de las 43 columnas ya resueltas por Gate 3B, los flags
    auxiliares de Gate 4 necesarios para reconstruir MetricStatus en Power BI
    sin logica por circuito (Gate 4A.0 Seccion 7-9): FechaIncompletaCalendario,
    FechaIncompletaDigital, PolicyBloqueadaSlotSeconds, y una representacion
    BI-safe de SlotsComerciales/SegundosComerciales (numero nullable +
    flag de "desconocida", nunca 0), y BValor/HValor/M2Valor (numero nullable
    derivado de b/h/m2, columnas originales preservadas como texto). q NO
    tiene columna numerica derivada: su semantica de negocio no esta
    confirmada (ver comentario mas abajo). Todos calculados reutilizando
    funciones y config ya aprobados por Gate 3, sin reglas nuevas.
    """
    maestro = semantic_result["maestro"].copy(deep=True)
    config = semantic_result["config"]

    dup_mask = maestro["ElementoID"].duplicated(keep=False)
    if dup_mask.any():
        duplicados = sorted(maestro.loc[dup_mask, "ElementoID"].unique().tolist())
        raise ExportError(
            f"DIM_ELEMENTOS: ElementoID duplicado(s), no se puede establecer PK unica: {duplicados}"
        )
    if maestro["ElementoID"].isna().any():
        raise ExportError("DIM_ELEMENTOS: existe(n) fila(s) con ElementoID vacio en el maestro enriquecido")

    todos_los_ids = maestro["ElementoID"].tolist()
    fecha_incompleta_calendario_ids = engine._incomplete_date_elements(todos_los_ids, exclude_no_aplica=False)
    fecha_incompleta_digital_ids = engine._incomplete_date_elements(todos_los_ids, exclude_no_aplica=True)
    policy_circuitos = set(config["metric_policies"]["slot_seconds_no_aplica_circuitos"])
    sentinel = config["digital_capacity"]["requiere_confirmacion_sentinel"]

    maestro["FechaIncompletaCalendario"] = maestro["ElementoID"].isin(fecha_incompleta_calendario_ids)
    maestro["FechaIncompletaDigital"] = maestro["ElementoID"].isin(fecha_incompleta_digital_ids)
    maestro["PolicyBloqueadaSlotSeconds"] = maestro["CircuitoNegocio"].isin(policy_circuitos)

    slots_valor, slots_desconocida = _capacidad_bi_safe(maestro["SlotsComerciales"], sentinel)
    segundos_valor, segundos_desconocida = _capacidad_bi_safe(maestro["SegundosComerciales"], sentinel)
    maestro["SlotsComercialesValor"] = slots_valor.astype("Int64")
    maestro["CapacidadSlotsDesconocida"] = slots_desconocida
    maestro["SegundosComercialesValor"] = segundos_valor.astype("Float64")
    maestro["CapacidadSegundosDesconocida"] = segundos_desconocida

    # b/h/m2: columnas fisicas de MAESTRO_ELEMENTOS (base/alto/area), mayormente
    # numericas (96%+) con un conjunto chico y semantico de valores no numericos
    # ("No aplica", espacio no separable) - se auditaron contra produccion antes
    # de aprobar este tratamiento. pd.to_numeric(errors="coerce") nunca convierte
    # "no aplica"/vacio a 0: queda pd.NA. Se calculan sobre la columna original
    # (antes del saneo de texto de mas abajo), que se preserva intacta para
    # trazabilidad.
    maestro["BValor"] = pd.to_numeric(maestro["b"], errors="coerce").astype("Float64")
    maestro["HValor"] = pd.to_numeric(maestro["h"], errors="coerce").astype("Float64")
    maestro["M2Valor"] = pd.to_numeric(maestro["m2"], errors="coerce").astype("Float64")

    # q: 86.77% convertible en produccion, pero su semantica de negocio no esta
    # confirmada (no documentada en Gate 3A/3B). NO se crea QValor: no se infiere
    # que "q" signifique cantidad ni ninguna otra cosa. Queda unicamente como
    # texto saneado, igual que DimensionOptico/DimensionTotal/Observaciones.
    # Significado de q = A VALIDAR con Brand Plus antes de exponerla como numero.

    maestro = _sanear_columnas_texto_mixto(maestro)
    maestro = maestro.sort_values("ElementoID", kind="stable").reset_index(drop=True)
    return maestro


# ---------------------------------------------------------------------------
# FACT_CAMPANAS
# ---------------------------------------------------------------------------


def build_fact_campanas(semantic_result: dict[str, Any], dim_elementos: pd.DataFrame) -> pd.DataFrame:
    """1 fila por CargaID. Mantiene exactamente las 29 columnas originales de
    CAMPANAS (Gate 2) - no duplica las 17 columnas semanticas de Gate 3B (se
    acceden en Power BI via ElementoID -> DIM_ELEMENTOS). IDCampaña queda
    como columna de agrupacion comercial, no como PK; las filas con
    IDCampaña vacio se conservan tal cual.
    """
    campanas = semantic_result["campanas"][vi.CAMPANAS_HEADERS].copy(deep=True)

    if campanas["CargaID"].isna().any():
        raise ExportError("FACT_CAMPANAS: existe(n) fila(s) con CargaID vacio")
    dup_mask = campanas["CargaID"].duplicated(keep=False)
    if dup_mask.any():
        duplicados = sorted(campanas.loc[dup_mask, "CargaID"].unique().tolist())
        raise ExportError(f"FACT_CAMPANAS: CargaID duplicado(s), no se puede establecer PK unica: {duplicados}")

    ids_validos = set(dim_elementos["ElementoID"])
    huerfanos = campanas.loc[campanas["ElementoID"].notna() & ~campanas["ElementoID"].isin(ids_validos), "ElementoID"]
    if len(huerfanos):
        raise ExportError(
            f"FACT_CAMPANAS: ElementoID sin correspondencia en DIM_ELEMENTOS: {sorted(huerfanos.unique().tolist())[:10]}"
        )

    campanas = _sanear_columnas_texto_mixto(campanas)
    campanas = campanas.sort_values("CargaID", kind="stable").reset_index(drop=True)
    return campanas


# ---------------------------------------------------------------------------
# DIM_CALENDARIO
# ---------------------------------------------------------------------------


def resolve_rango_temporal(fact_campanas: pd.DataFrame) -> tuple[pd.Timestamp, pd.Timestamp]:
    """Rango dinamico real: minima FechaInicio valida -> maxima fecha valida
    presente (FechaFin, incluye reservas futuras). Nunca hardcodea años."""
    inicio = fact_campanas["FechaInicio"].min()
    fin = fact_campanas["FechaFin"].max()
    if pd.isna(inicio) or pd.isna(fin):
        raise ExportError("No se pudo determinar un rango temporal valido a partir de FACT_CAMPANAS")
    return pd.Timestamp(inicio).normalize(), pd.Timestamp(fin).normalize()


def build_dim_calendario(rango_inicio: pd.Timestamp, rango_fin: pd.Timestamp) -> pd.DataFrame:
    """1 fila por dia, generada en Python. Rango dinamico, no hardcodeado."""
    if rango_fin < rango_inicio:
        raise ExportError("DIM_CALENDARIO: rango_fin es anterior a rango_inicio")

    fechas = pd.date_range(rango_inicio, rango_fin, freq="D")
    df = pd.DataFrame({"Fecha": fechas})
    periodo_mes = df["Fecha"].dt.to_period("M")

    df["Año"] = df["Fecha"].dt.year
    df["Mes"] = df["Fecha"].dt.month
    df["NombreMes"] = df["Mes"].apply(lambda m: MESES_ES[m - 1])
    df["MesAño"] = df["Fecha"].dt.strftime("%Y-%m")
    df["Trimestre"] = df["Fecha"].dt.quarter
    df["Semana"] = df["Fecha"].dt.isocalendar().week.astype(int)
    df["DiaSemana"] = df["Fecha"].dt.dayofweek.apply(lambda d: DIAS_ES[d])
    df["InicioMes"] = periodo_mes.dt.to_timestamp()
    df["FinMes"] = periodo_mes.dt.to_timestamp(how="end").dt.normalize()

    return df


# ---------------------------------------------------------------------------
# BRIDGE_CAMPANA_DIA
# ---------------------------------------------------------------------------


def _explode_by_carga(scope: pd.DataFrame) -> pd.DataFrame:
    """Igual patron que metrics_engine._explode_days pero por CargaID en vez
    de ElementoID (la bridge es minimalista: solo pertenencia temporal)."""
    if scope.empty:
        return pd.DataFrame(columns=["CargaID", "Fecha"])
    dias_series = [pd.date_range(s, e, freq="D") for s, e in zip(scope["_eff_start"], scope["_eff_end"])]
    exploded = scope[["CargaID"]].assign(_dias=dias_series).explode("_dias")
    return exploded.rename(columns={"_dias": "Fecha"}).reset_index(drop=True)


def build_bridge_campana_dia(
    engine: MetricsEngine, fact_campanas: pd.DataFrame, rango_inicio: pd.Timestamp, rango_fin: pd.Timestamp
) -> pd.DataFrame:
    """Grain (CargaID, Fecha): solo dias donde el placement esta realmente
    vigente segun la MISMA logica temporal ya aprobada de Gate 3
    (MetricsEngine._campanas_overlap: Estado!=Cancelado, solapamiento
    inclusivo, FechaIndefinida extendida hasta rango_fin, fechas incompletas
    excluidas - nunca inventadas). Minimalista: solo CargaID + Fecha."""
    todos_los_elemento_ids = fact_campanas["ElementoID"].dropna().unique().tolist()
    overlap = engine._campanas_overlap(todos_los_elemento_ids, str(rango_inicio.date()), str(rango_fin.date()))
    bridge = _explode_by_carga(overlap)
    bridge = bridge.drop_duplicates(subset=["CargaID", "Fecha"])
    bridge = bridge.sort_values(["CargaID", "Fecha"], kind="stable").reset_index(drop=True)
    return bridge


# ---------------------------------------------------------------------------
# FACT_METRICAS_DIARIA (WIDE, sparse)
# ---------------------------------------------------------------------------


def build_fact_metricas_diaria(
    engine: MetricsEngine, dim_elementos: pd.DataFrame, rango_inicio: pd.Timestamp, rango_fin: pd.Timestamp
) -> pd.DataFrame:
    """Grain (ElementoID, Fecha), solo pares con actividad real (sparse). Un
    dia sin fila equivale a actividad=0, NO a "dia inexistente" (Gate 4A.0
    Seccion 8): la reconstruccion posterior debe iterar DIM_CALENDARIO, no
    solo esta tabla.

    Reutiliza literalmente MetricsEngine._campanas_overlap (familia
    calendario: Medio indistinto, exclude_no_aplica=False, igual que
    _element_day_stats) y (familia digital: exclude_no_aplica=True, igual
    que _digital_period_activity) para garantizar fidelidad con Gate 3 sin
    reimplementar la resolucion de fechas. Los guardarraíles de exclusividad
    / duracion de spot no soportada se verifican de la misma forma que
    MetricsEngine._digital_period_activity: si aparecen, la exportacion
    falla explicito (Gate 4 nunca inventa la matematica que Gate 3 rechaza).
    """
    config = engine.config
    seg_por_salida = config["digital_capacity"]["spot_segundos_por_salida"]
    inicio_str, fin_str = str(rango_inicio.date()), str(rango_fin.date())

    # --- familia calendario: TODOS los elementos, sin distincion de Medio ---
    todos_los_ids = dim_elementos["ElementoID"].tolist()
    overlap_calendario = engine._campanas_overlap(todos_los_ids, inicio_str, fin_str)
    exploded_calendario = _explode_days(overlap_calendario)
    pares_ocupados = exploded_calendario[["ElementoID", "_day"]].drop_duplicates()
    pares_ocupados = pares_ocupados.rename(columns={"_day": "Fecha"})
    pares_ocupados["DiasOcupado"] = 1

    # --- familia digital: solo elementos Digital, exclude_no_aplica=True ---
    ids_digital = dim_elementos.loc[dim_elementos["Medio"] == "Digital", "ElementoID"].tolist()
    overlap_digital = engine._campanas_overlap(ids_digital, inicio_str, fin_str, exclude_no_aplica=True)

    if not overlap_digital.empty:
        exclusividad_mask = overlap_digital["TipoExclusividad"].notna() | overlap_digital["ModalidadPauta"].astype(
            str
        ).str.startswith("Exclusividad")
        if exclusividad_mask.any():
            offenders = sorted(overlap_digital.loc[exclusividad_mask, "ElementoID"].unique().tolist())[:10]
            raise ExportError(
                f"FACT_METRICAS_DIARIA: {int(exclusividad_mask.sum())} fila(s) con exclusividad dentro del rango "
                f"exportado. Gate 3 no calcula esa matematica todavia; Gate 4 no la inventa. "
                f"ElementoID afectados (muestra): {offenders}"
            )
        spot_base = config["digital_capacity"]["spot_duracion_base_seg"]
        spot_invalid_mask = overlap_digital["DuracionSpotSeg"].notna() & (
            overlap_digital["DuracionSpotSeg"] != spot_base
        )
        if spot_invalid_mask.any():
            offenders = (
                overlap_digital.loc[spot_invalid_mask, ["ElementoID", "DuracionSpotSeg"]]
                .drop_duplicates()
                .head(10)
                .to_dict("records")
            )
            raise ExportError(
                f"FACT_METRICAS_DIARIA: spot(s) con DuracionSpotSeg != {spot_base}s dentro del rango exportado. "
                f"Gate 3 no calcula ese escalado todavia; Gate 4 no lo inventa. Filas afectadas (muestra): {offenders}"
            )

    exploded_digital = _explode_days(overlap_digital, extra_cols=["SalidasVendidas"])
    if len(exploded_digital):
        agregado_digital = exploded_digital.groupby(["ElementoID", "_day"]).agg(
            SlotsOcupadosDia=("ElementoID", "size"),
            SegundosVendidosDia=("SalidasVendidas", lambda s: float((s.dropna() * seg_por_salida).sum())),
            HasSalidasIndeterminada=("SalidasVendidas", lambda s: bool(s.isna().any())),
        )
        agregado_digital = agregado_digital.reset_index().rename(columns={"_day": "Fecha"})
    else:
        agregado_digital = pd.DataFrame(
            columns=["ElementoID", "Fecha", "SlotsOcupadosDia", "SegundosVendidosDia", "HasSalidasIndeterminada"]
        )

    fact = pares_ocupados.merge(agregado_digital, on=["ElementoID", "Fecha"], how="left")

    medio_por_elemento = dim_elementos.set_index("ElementoID")["Medio"]
    es_digital = fact["ElementoID"].map(medio_por_elemento) == "Digital"

    fact["SlotsOcupadosDia"] = fact["SlotsOcupadosDia"].where(~es_digital | fact["SlotsOcupadosDia"].notna(), 0)
    fact["SegundosVendidosDia"] = fact["SegundosVendidosDia"].where(
        ~es_digital | fact["SegundosVendidosDia"].notna(), 0.0
    )
    fact["HasSalidasIndeterminada"] = fact["HasSalidasIndeterminada"].where(
        ~es_digital | fact["HasSalidasIndeterminada"].notna(), False
    )

    # Los dtypes nullable deben fijarse ANTES de escribir pd.NA por .loc: un
    # bloque numpy nativo (bool/float64) rechaza pd.NA con TypeError.
    fact["SlotsOcupadosDia"] = fact["SlotsOcupadosDia"].astype("Int64")
    fact["SegundosVendidosDia"] = fact["SegundosVendidosDia"].astype("Float64")
    fact["HasSalidasIndeterminada"] = fact["HasSalidasIndeterminada"].astype("boolean")
    fact["DiasOcupado"] = fact["DiasOcupado"].astype("Int64")

    # elementos Estatico: slots/segundos/salidas-indeterminada no aplican -> NA explicito, nunca 0
    fact.loc[~es_digital, ["SlotsOcupadosDia", "SegundosVendidosDia", "HasSalidasIndeterminada"]] = pd.NA

    fact = fact[["ElementoID", "Fecha", "DiasOcupado", "SlotsOcupadosDia", "SegundosVendidosDia", "HasSalidasIndeterminada"]]
    fact = fact.sort_values(["ElementoID", "Fecha"], kind="stable").reset_index(drop=True)
    return fact


# ---------------------------------------------------------------------------
# Manifest + escritura
# ---------------------------------------------------------------------------


def build_manifest(
    sha_before: str,
    sha_after: str,
    tablas: dict[str, pd.DataFrame],
    rango_inicio: pd.Timestamp,
    rango_fin: pd.Timestamp,
    warnings: list[str],
    source_path: Path,
) -> dict[str, Any]:
    return {
        "schema_version": SCHEMA_VERSION,
        "export_timestamp_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
        "source_file": str(source_path),
        "source_sha256_before": sha_before,
        "source_sha256_after": sha_after,
        "source_sha256_match": sha_before == sha_after,
        "rango_temporal": {"inicio": str(rango_inicio.date()), "fin": str(rango_fin.date())},
        "tablas": {
            nombre: {"rows": int(len(df)), "columns": list(df.columns)} for nombre, df in tablas.items()
        },
        "warnings_count": len(warnings),
        "warnings_sample": warnings[:50],
    }


def write_outputs(tablas: dict[str, pd.DataFrame], manifest: dict[str, Any], output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    nombre_archivo = {
        "DIM_ELEMENTOS": "dim_elementos.parquet",
        "FACT_CAMPANAS": "fact_campanas.parquet",
        "DIM_CALENDARIO": "dim_calendario.parquet",
        "BRIDGE_CAMPANA_DIA": "bridge_campana_dia.parquet",
        "FACT_METRICAS_DIARIA": "fact_metricas_diaria.parquet",
    }
    for nombre, df in tablas.items():
        df.to_parquet(output_dir / nombre_archivo[nombre], engine="pyarrow", index=False)

    with open(output_dir / "_export_manifest.json", "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, ensure_ascii=False, indent=2, default=str)


# ---------------------------------------------------------------------------
# Orquestacion
# ---------------------------------------------------------------------------


def export_data(path: str | Path = DEFAULT_INPUT_PATH, output_dir: str | Path = DEFAULT_OUTPUT_DIR) -> dict[str, Any]:
    """Punto de entrada de Gate 4A. Read-only sobre el Excel fuente (verifica
    SHA-256 antes/despues). No reabre ni reinterpreta Gate 1/2/3."""
    path = Path(path)
    output_dir = Path(output_dir)

    sha_before = vi.calculate_sha256(path)

    transform_result, semantic_result, engine = load_pipeline(path)

    dim_elementos = build_dim_elementos(semantic_result, engine)
    fact_campanas = build_fact_campanas(semantic_result, dim_elementos)
    rango_inicio, rango_fin = resolve_rango_temporal(fact_campanas)
    dim_calendario = build_dim_calendario(rango_inicio, rango_fin)
    bridge_campana_dia = build_bridge_campana_dia(engine, fact_campanas, rango_inicio, rango_fin)
    fact_metricas_diaria = build_fact_metricas_diaria(engine, dim_elementos, rango_inicio, rango_fin)

    sha_after = vi.calculate_sha256(path)
    if sha_after != sha_before:
        raise ExportError(
            f"ERROR CRITICO: el SHA-256 del input cambio durante la exportacion "
            f"(antes={sha_before}, despues={sha_after})."
        )

    tablas = {
        "DIM_ELEMENTOS": dim_elementos,
        "FACT_CAMPANAS": fact_campanas,
        "DIM_CALENDARIO": dim_calendario,
        "BRIDGE_CAMPANA_DIA": bridge_campana_dia,
        "FACT_METRICAS_DIARIA": fact_metricas_diaria,
    }
    manifest = build_manifest(
        sha_before, sha_after, tablas, rango_inicio, rango_fin, semantic_result["warnings"], path
    )
    write_outputs(tablas, manifest, output_dir)

    return {"tablas": tablas, "manifest": manifest, "output_dir": output_dir}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Gate 4A - exporta la capa de salida Parquet para Power BI.")
    parser.add_argument("--file", default=str(DEFAULT_INPUT_PATH), help="Ruta al archivo .xlsx a exportar")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="Carpeta de salida de los Parquet")
    args = parser.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    try:
        result = export_data(args.file, args.output_dir)
    except (ExportError, UnsupportedBusinessCaseError) as exc:
        print("GATE4A_EXPORT_ERROR:", exc)
        return 1

    print("GATE4A_EXPORT_OK")
    print(json.dumps(result["manifest"], ensure_ascii=False, indent=2, default=str))
    return 0


if __name__ == "__main__":
    sys.exit(main())
