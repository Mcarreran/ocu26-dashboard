"""Motor de cruce V2 (segundo pase, autorizado) OCU26.

Diferencia central respecto de V1 (merge_common.py, que NO se modifica ni se
reutiliza para las decisiones de fila — solo se reutiliza para lo que no
cambió: rutas base, is_blank, read_table_df, inspect_structure, vi):

  V1: ante un conflicto (ambos valores no vacíos y distintos) -> conservar
      ACTUAL, nunca sobrescribir.
  V2 (este archivo, por instrucción explícita del usuario): NUEVA representa
      la actualización más reciente -> ante un conflicto en un campo
      operativo, gana NUEVA. Un valor vacío en NUEVA nunca borra un dato
      existente en ACTUAL (regla 5 del pedido).

Alcance de la nueva precedencia: SOLO campos operativos/comerciales de
CAMPANAS, para filas con coincidencia exacta de CargaID. Nunca se aplica a
campos técnicos (CargaID, ElementoID, FechaHoraCarga, UsuarioCarga,
FuenteCarga) ni a campos de sistema (ClaveNegocio, EstadoValidacion,
ObservacionValidacion, FilaOrigen), que tienen sus propias reglas debajo.
MAESTRO_ELEMENTOS conserva la lógica conservadora de V1 (el pedido de este
pase es explícito y exclusivamente sobre CAMPANAS/CargaID); en la práctica
no tiene efecto porque MAESTRO_ELEMENTOS es idéntico entre ACTUAL y NUEVA.

--- Fórmula de ClaveNegocio ---
Reconstruida empíricamente y verificada contra las 9.503 filas de
CAMPANAS de ACTUAL (match exacto 9503/9503, ver auditoría de esta corrida):

    ClaveNegocio = f"{IDCampaña}|{ElementoID}|{FechaInicio:%Y-%m-%d}|{FechaFinSegment}|{HoraSegment(HoraInicio)}|{HoraSegment(HoraFin)}"

  - Cada componente vacío se serializa como cadena vacía.
  - IDCampaña: entero sin parte decimal si es numérico.
  - Fechas: 'YYYY-MM-DD'.
  - FechaFinSegment: 'INDET' si FechaFin está vacía Y FechaIndefinida=='Si';
    si no, la fecha formateada (o vacío).
  - HoraSegment: '1900-01-01' (literal) si la celda de hora tiene un valor;
    '' si está vacía. Esto reproduce fielmente una particularidad histórica
    ya presente en el 100% de los datos migrados (la hora nunca aparece
    como HH:MM:SS dentro de ClaveNegocio, siempre como la fecha época de
    Excel) — no se "corrige" porque el pedido es reproducir la lógica
    canónica EXISTENTE, no una versión mejorada.

Esta fórmula NO estaba escrita como función en ningún script del proyecto;
se dedujo por ingeniería inversa de los datos y se validó al 100% antes de
usarse aquí (ver docstring de compute_clave_negocio). Se reproduce de forma
literal, sin leer en runtime ningún archivo de auditoría histórica (mismo
patrón que scripts/transform_data.py con TipoInventario).

--- Regla EstadoValidacion para registros nuevos ---
`scripts/validate_input.py` únicamente valida que EstadoValidacion
pertenezca a {OK, PENDIENTE_HISTORICO, PENDIENTE_DUPLICADO,
PENDIENTE_COLISION}; no calcula ese valor a partir de otros campos. La
única evidencia documentada en el proyecto sobre cómo se asignó
originalmente (audit_sources/INFORME_RECONCILIACION_MIGRACION.md, tabla de
categorías de la sección de EstadoValidacion) dice:

    OK                  -> fila sin problema identificado
    PENDIENTE_HISTORICO -> FechaInicio vacía (u otro campo clave vacío)
    PENDIENTE_DUPLICADO -> ClaveNegocio duplica exactamente otra fila
    PENDIENTE_COLISION  -> ClaveNegocio coincide con otra fila con datos
                           distintos

Un registro nuevo que (a) tiene ElementoID válido, (b) tiene FechaInicio,
FechaFin/FechaIndefinida, IDCampaña y Estado completos, y (c) genera una
ClaveNegocio que no colisiona con ninguna existente, no cae en ninguna de
las tres categorías PENDIENTE_* documentadas -> EstadoValidacion="OK" es la
consecuencia directa de esa clasificación documentada, no una invención.
Si alguno de esos tres controles falla, el registro completo queda en
A_VALIDAR y no se incorpora (no se "adivina" un PENDIENTE_* sin evidencia
de cuál correspondería).
"""

from __future__ import annotations

import datetime as dt
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import merge_common as mc  # noqa: E402

vi = mc.vi
is_blank = mc.is_blank
read_table_df = mc.read_table_df
inspect_structure = mc.inspect_structure
calculate_sha256 = mc.calculate_sha256

ACTUAL_PATH = mc.ACTUAL_PATH
NUEVA_PATH = mc.NUEVA_PATH
CANDIDATE_V2_PATH = mc.UPDATE_ROOT / "output" / "OCU26_BASE_DATOS_INTEGRADA_CANDIDATA_V2_2026-08-18.xlsx"
AUDIT_V2_PATH = mc.UPDATE_ROOT / "output" / "OCU26_AUDITORIA_CRUCE_V2_2026-08-18.xlsx"

CAMPANAS_KEY = mc.CAMPANAS_KEY
MAESTRO_KEY = mc.MAESTRO_KEY

TECHNICAL_FIELDS_CAMPANAS = {"CargaID", "ElementoID", "FechaHoraCarga", "UsuarioCarga", "FuenteCarga"}
SYSTEM_FIELDS_CAMPANAS = {"ClaveNegocio", "EstadoValidacion", "ObservacionValidacion", "FilaOrigen"}
OPERATIONAL_FIELDS_CAMPANAS = [
    h for h in mc.CAMPANAS_HEADERS if h not in TECHNICAL_FIELDS_CAMPANAS and h not in SYSTEM_FIELDS_CAMPANAS
]
CLAVE_NEGOCIO_TRIGGER_FIELDS = {"IDCampaña", "FechaInicio", "FechaFin", "FechaIndefinida", "HoraInicio", "HoraFin"}

CARGA_ID_PREFIX = "HIST-"
CARGA_ID_WIDTH = 8


def _domain_from_parametros(parametros_df: pd.DataFrame, categoria: str) -> set[str]:
    sub = parametros_df[parametros_df["Categoria"] == categoria]
    return set(sub["Valor"].dropna())


def fmt_id_component(v: Any) -> str:
    if is_blank(v):
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)) and float(v).is_integer():
        return str(int(v))
    return str(v)


def fmt_date_component(v: Any) -> str:
    if is_blank(v):
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    return str(v)


def fmt_hora_component(v: Any) -> str:
    return "" if is_blank(v) else "1900-01-01"


def compute_clave_negocio(
    id_campana: Any, elemento_id: Any, fecha_inicio: Any, fecha_fin: Any, fecha_indefinida: Any, hora_inicio: Any, hora_fin: Any
) -> str:
    """Reconstrucción validada al 100% contra los datos reales (ver docstring del módulo)."""
    idc = fmt_id_component(id_campana)
    eid = elemento_id if not is_blank(elemento_id) else ""
    fi = fmt_date_component(fecha_inicio)
    ff = "INDET" if (is_blank(fecha_fin) and fecha_indefinida == "Si") else fmt_date_component(fecha_fin)
    hi = fmt_hora_component(hora_inicio)
    hf = fmt_hora_component(hora_fin)
    return f"{idc}|{eid}|{fi}|{ff}|{hi}|{hf}"


@dataclass
class ChangeRecord:
    hoja: str
    carga_id: Any
    elemento_id: Any
    columna: str
    valor_anterior: Any
    valor_nuevo: Any
    motivo: str
    resultado: str  # APLICADO | RECHAZADO_A_VALIDAR
    fuente: str  # NUEVA | CALCULADO | GENERADO


def _validate_field_value(col: str, value: Any, domains: dict[str, set[str]]) -> tuple[bool, str]:
    if col == "Estado":
        if value not in domains["Estado"]:
            return False, f"Estado fuera de dominio permitido ({sorted(domains['Estado'])})"
        return True, ""
    if col == "FechaIndefinida":
        if value not in domains["SiNo"]:
            return False, f"FechaIndefinida fuera de dominio permitido ({sorted(domains['SiNo'])})"
        return True, ""
    if col in ("FechaInicio", "FechaFin"):
        if not isinstance(value, (dt.date, dt.datetime)):
            return False, f"{col} no es una fecha válida (tipo={type(value).__name__})"
        return True, ""
    if col == "TipoCargaDeclarado":
        if value not in vi.MEDIO_VALUES:
            return False, f"TipoCargaDeclarado fuera de dominio permitido ({sorted(vi.MEDIO_VALUES)})"
        return True, ""
    if col in ("PROGRAMATICA", "CANJE"):
        if value not in vi.SI_NO_VALUES:
            return False, f"{col} fuera de dominio permitido ({sorted(vi.SI_NO_VALUES)})"
        return True, ""
    if col == "ModalidadPauta" and domains.get("ModalidadPauta"):
        if value not in domains["ModalidadPauta"]:
            return False, f"ModalidadPauta fuera de dominio permitido ({sorted(domains['ModalidadPauta'])})"
        return True, ""
    if col == "TipoExclusividad" and domains.get("TipoExclusividad"):
        if value not in domains["TipoExclusividad"]:
            return False, f"TipoExclusividad fuera de dominio permitido ({sorted(domains['TipoExclusividad'])})"
        return True, ""
    return True, ""


def process_campanas_common_row(
    actual_row: dict[str, Any], nueva_row: dict[str, Any], domains: dict[str, set[str]]
) -> tuple[list[ChangeRecord], dict[str, Any]]:
    """Aplica la precedencia V2 a una fila común (mismo CargaID). Devuelve
    (registros_de_cambio_para_auditoria, campos_aplicados_a_la_candidata)."""
    key = actual_row[CAMPANAS_KEY]
    elemento_id = actual_row["ElementoID"]
    changes: list[ChangeRecord] = []

    # --- Pre-check obligatorio: mismo ElementoID ---
    if actual_row["ElementoID"] != nueva_row["ElementoID"]:
        changes.append(
            ChangeRecord(
                "CAMPANAS", key, elemento_id, "ElementoID", actual_row["ElementoID"], nueva_row["ElementoID"],
                "ElementoID no coincide entre ACTUAL y NUEVA para el mismo CargaID: fila bloqueada por completo",
                "RECHAZADO_A_VALIDAR", "NUEVA",
            )
        )
        return changes, {}

    pending: dict[str, tuple[Any, str]] = {}  # col -> (valor, accion)

    for col in OPERATIONAL_FIELDS_CAMPANAS:
        v_a = actual_row[col]
        v_n = nueva_row[col]
        blank_a, blank_n = is_blank(v_a), is_blank(v_n)

        if blank_a and blank_n:
            continue
        if blank_n and not blank_a:
            continue  # regla 5: un vacío del nuevo nunca borra el dato actual
        if blank_a and not blank_n:
            pending[col] = (v_n, "COMPLETADO")
            continue
        if v_a == v_n:
            continue

        # conflicto real -> gana NUEVA si pasa los controles
        valid, reason = _validate_field_value(col, v_n, domains)
        if not valid:
            changes.append(
                ChangeRecord("CAMPANAS", key, elemento_id, col, v_a, v_n, f"Tipo/dominio inválido: {reason}", "RECHAZADO_A_VALIDAR", "NUEVA")
            )
            continue
        pending[col] = (v_n, "ACTUALIZADO")

    # --- Coherencia de fechas sobre el resultado tentativo ---
    final_fi = pending.get("FechaInicio", (actual_row["FechaInicio"], None))[0]
    final_ff = pending.get("FechaFin", (actual_row["FechaFin"], None))[0]
    final_findef = pending.get("FechaIndefinida", (actual_row["FechaIndefinida"], None))[0]

    coherence_reason = None
    if not is_blank(final_findef) and final_findef == "Si" and not is_blank(final_ff):
        coherence_reason = "FechaIndefinida='Si' pero FechaFin no vacía tras aplicar la actualización"
    elif (
        isinstance(final_fi, (dt.date, dt.datetime))
        and isinstance(final_ff, (dt.date, dt.datetime))
        and final_ff < final_fi
    ):
        coherence_reason = f"FechaFin ({final_ff}) quedaría anterior a FechaInicio ({final_fi}) tras aplicar la actualización"

    if coherence_reason:
        for col in ("FechaInicio", "FechaFin", "FechaIndefinida"):
            if col in pending:
                v_n, _ = pending.pop(col)
                changes.append(
                    ChangeRecord("CAMPANAS", key, elemento_id, col, actual_row[col], v_n, coherence_reason, "RECHAZADO_A_VALIDAR", "NUEVA")
                )

    # applied: solo campos ACEPTADOS a nivel de tipo/dominio/coherencia. Las
    # ChangeRecord "APLICADO" se generan recién en finalize_campanas_updates,
    # después de la pasada global de deduplicación de ClaveNegocio (un campo
    # aceptado acá puede terminar revertido si genera una ClaveNegocio
    # duplicada con otra fila).
    applied: dict[str, tuple[Any, str]] = dict(pending)

    return changes, applied


def finalize_campanas_updates(
    actual_campanas_df: pd.DataFrame,
    per_row: dict[Any, tuple[dict[str, Any], dict[str, tuple[Any, str]]]],
) -> tuple[dict[Any, dict[str, Any]], list[ChangeRecord]]:
    """Pasada global (orden determinístico por CargaID) que:

    1. Recalcula ClaveNegocio para toda fila cuyos campos aplicados incluyan
       algún componente de la clave (IDCampaña/fechas/horarios).
    2. Si la ClaveNegocio recalculada colisiona con OTRA fila (existente en
       ACTUAL o ya aceptada en esta misma pasada), revierte únicamente los
       campos componentes de esa fila puntual (vuelven al valor de ACTUAL) y
       deja constancia en A_VALIDAR. El resto de los campos no-componentes
       de la misma fila (p.ej. Proveedor) permanece aplicado.
    3. Genera los ChangeRecord "APLICADO" definitivos.

    Devuelve (final_applied: {CargaID: {columna: valor}}, change_records).
    """
    change_records: list[ChangeRecord] = []
    final_applied: dict[Any, dict[str, Any]] = {}
    clave_universe = set(actual_campanas_df["ClaveNegocio"].dropna())

    for key in sorted(per_row.keys(), key=str):
        actual_row, pending = per_row[key]
        elemento_id = actual_row["ElementoID"]
        applied_values = {col: v for col, (v, _action) in pending.items()}

        if any(col in applied_values for col in CLAVE_NEGOCIO_TRIGGER_FIELDS):
            merged = dict(actual_row)
            merged.update(applied_values)
            new_clave = compute_clave_negocio(
                merged["IDCampaña"], merged["ElementoID"], merged["FechaInicio"], merged["FechaFin"],
                merged["FechaIndefinida"], merged["HoraInicio"], merged["HoraFin"],
            )
            old_clave = actual_row["ClaveNegocio"]
            if new_clave != old_clave:
                probe_universe = clave_universe - {old_clave}
                if new_clave in probe_universe:
                    for col in list(applied_values.keys()):
                        if col in CLAVE_NEGOCIO_TRIGGER_FIELDS:
                            v = applied_values.pop(col)
                            change_records.append(
                                ChangeRecord(
                                    "CAMPANAS", key, elemento_id, col, actual_row[col], v,
                                    f"Revertido: aplicar este valor generaría ClaveNegocio duplicada ({new_clave!r}), "
                                    "ya utilizada por otra fila. No se sobrescribe silenciosamente.",
                                    "RECHAZADO_A_VALIDAR", "NUEVA",
                                )
                            )
                else:
                    clave_universe.discard(old_clave)
                    clave_universe.add(new_clave)
                    applied_values["ClaveNegocio"] = new_clave
                    change_records.append(
                        ChangeRecord(
                            "CAMPANAS", key, elemento_id, "ClaveNegocio", old_clave, new_clave,
                            "Recalculada con la fórmula canónica porque cambió un campo componente "
                            "(IDCampaña/fechas/horarios)",
                            "APLICADO", "CALCULADO",
                        )
                    )

        for col, v in applied_values.items():
            if col == "ClaveNegocio":
                continue
            action = pending.get(col, (None, "ACTUALIZADO"))[1]
            motivo = (
                "Completado: ACTUAL vacío, NUEVA con dato"
                if action == "COMPLETADO"
                else "Actualizado: NUEVA representa la versión más reciente"
            )
            change_records.append(ChangeRecord("CAMPANAS", key, elemento_id, col, actual_row[col], v, motivo, "APLICADO", "NUEVA"))

        if applied_values:
            final_applied[key] = applied_values

    return final_applied, change_records


def next_carga_id_sequence(actual_campanas_df: pd.DataFrame) -> int:
    ids = actual_campanas_df[CAMPANAS_KEY].dropna().astype(str)
    nums = ids.str.extract(rf"^{CARGA_ID_PREFIX}(\d{{{CARGA_ID_WIDTH}}})$")[0].dropna().astype(int)
    if nums.empty:
        raise ValueError("No se encontró ningún CargaID con el formato histórico esperado")
    return int(nums.max()) + 1


def build_new_campanas_records(
    unkeyed_new_rows: list[dict[str, Any]],
    actual_campanas_df: pd.DataFrame,
    maestro_ids: set[Any],
    next_num_start: int,
    run_timestamp: dt.datetime,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Intenta incorporar los registros nuevos de CAMPANAS sin CargaID.

    Orden determinístico: por ElementoID ascendente (única clave estable
    disponible en los propios datos nuevos).
    """
    incorporated: list[dict[str, Any]] = []
    rejected: list[dict[str, Any]] = []

    sorted_rows = sorted(unkeyed_new_rows, key=lambda r: str(r.get("ElementoID") or ""))
    existing_clave = set(actual_campanas_df["ClaveNegocio"].dropna())
    seen_clave: set[str] = set()
    next_num = next_num_start

    for row in sorted_rows:
        eid = row.get("ElementoID")
        reasons: list[str] = []

        if is_blank(eid):
            reasons.append("ElementoID vacío")
        elif eid not in maestro_ids:
            reasons.append(f"ElementoID '{eid}' no existe en MAESTRO_ELEMENTOS (integridad referencial)")

        for col in ("FechaInicio", "FechaFin"):
            v = row.get(col)
            if not is_blank(v) and not isinstance(v, (dt.date, dt.datetime)):
                reasons.append(f"{col} con tipo inválido: {v!r}")

        findef = row.get("FechaIndefinida")
        if not is_blank(findef) and findef not in {"Si", "No"}:
            reasons.append(f"FechaIndefinida fuera de dominio: {findef!r}")

        estado = row.get("Estado")
        # Estado no se valida contra PARAMETROS aquí porque puede llegar vacío;
        # se exige no-vacío como parte de "referencias completas".
        if is_blank(estado):
            reasons.append("Estado vacío (dato comercial incompleto, no se puede clasificar EstadoValidacion=OK)")

        if is_blank(row.get("IDCampaña")):
            reasons.append("IDCampaña vacío (dato comercial incompleto)")
        if is_blank(row.get("FechaInicio")):
            reasons.append("FechaInicio vacía (dato comercial incompleto)")

        clave = None
        if not reasons:
            clave = compute_clave_negocio(
                row.get("IDCampaña"), eid, row.get("FechaInicio"), row.get("FechaFin"),
                row.get("FechaIndefinida"), row.get("HoraInicio"), row.get("HoraFin"),
            )
            if clave in existing_clave:
                reasons.append(f"ClaveNegocio ya existe en ACTUAL: {clave!r} (posible duplicado/colisión)")
            elif clave in seen_clave:
                reasons.append(f"ClaveNegocio duplicada entre los propios registros nuevos: {clave!r}")

        if reasons:
            rejected.append({**row, "motivo_rechazo": "; ".join(reasons)})
            continue

        carga_id = f"{CARGA_ID_PREFIX}{next_num:0{CARGA_ID_WIDTH}d}"
        next_num += 1
        seen_clave.add(clave)

        record = dict(row)
        record["CargaID"] = carga_id
        record["ClaveNegocio"] = clave
        record["EstadoValidacion"] = "OK"
        record["ObservacionValidacion"] = None
        record["FechaHoraCarga"] = run_timestamp
        record["UsuarioCarga"] = "MIGRACION_OCU26"
        record["FuenteCarga"] = "Migración histórica"
        record["FilaOrigen"] = None
        incorporated.append(record)

    return incorporated, rejected


def load_domains(parametros_df: pd.DataFrame) -> dict[str, set[str]]:
    return {
        "Estado": _domain_from_parametros(parametros_df, "Estado"),
        "SiNo": _domain_from_parametros(parametros_df, "SiNo"),
        "ModalidadPauta": _domain_from_parametros(parametros_df, "ModalidadPauta"),
        "TipoExclusividad": _domain_from_parametros(parametros_df, "TipoExclusividad"),
    }


@dataclass
class V2Plan:
    actual_maestro_df: pd.DataFrame
    nueva_maestro_df: pd.DataFrame
    actual_campanas_df: pd.DataFrame
    nueva_campanas_df: pd.DataFrame
    actual_parametros_df: pd.DataFrame
    nueva_parametros_df: pd.DataFrame
    maestro_cls: Any  # mc.SheetClassification (lógica V1, sin cambios)
    campanas_final_applied: dict[Any, dict[str, Any]]
    campanas_change_records: list[ChangeRecord]
    new_records_incorporated: list[dict[str, Any]]
    new_records_rejected: list[dict[str, Any]]
    run_timestamp: dt.datetime


def compute_v2_plan(run_timestamp: dt.datetime | None = None) -> V2Plan:
    """Punto de entrada único y determinístico del cruce V2. Usado tanto por
    audit_ocu26_actualizacion_v2.py como por build_ocu26_integrada_v2.py para
    garantizar que ambos vean exactamente el mismo plan de cambios.

    `run_timestamp` es la única entrada no derivada de los datos (marca de
    tiempo de la corrida, usada para FechaHoraCarga de los registros nuevos).
    Si no se provee, se usa el instante actual.
    """
    if run_timestamp is None:
        run_timestamp = dt.datetime.now()

    actual_maestro_df = read_table_df(ACTUAL_PATH, "MAESTRO_ELEMENTOS", "tblElementos")
    nueva_maestro_df = read_table_df(NUEVA_PATH, "MAESTRO_ELEMENTOS", "tblElementos")
    actual_campanas_df = read_table_df(ACTUAL_PATH, "CAMPANAS", "tblCampanas")
    nueva_campanas_df = read_table_df(NUEVA_PATH, "CAMPANAS", "tblCampanas")
    actual_parametros_df = read_table_df(ACTUAL_PATH, "PARAMETROS", "tblParametros")
    nueva_parametros_df = read_table_df(NUEVA_PATH, "PARAMETROS", "tblParametros")

    # MAESTRO_ELEMENTOS: sin cambio de precedencia respecto de V1 (el pedido
    # de este pase es explícito y exclusivo sobre CAMPANAS/CargaID).
    maestro_cls = mc.classify_maestro(actual_maestro_df, nueva_maestro_df)

    domains = load_domains(actual_parametros_df)

    a_idx = actual_campanas_df.set_index(CAMPANAS_KEY, drop=False)
    n_idx = nueva_campanas_df.dropna(subset=[CAMPANAS_KEY]).set_index(CAMPANAS_KEY, drop=False)
    common_keys = a_idx.index.intersection(n_idx.index)

    per_row: dict[Any, tuple[dict[str, Any], dict[str, tuple[Any, str]]]] = {}
    change_records: list[ChangeRecord] = []
    for key in common_keys:
        actual_row = a_idx.loc[key].to_dict()
        nueva_row = n_idx.loc[key].to_dict()
        rejected_changes, pending = process_campanas_common_row(actual_row, nueva_row, domains)
        change_records.extend(rejected_changes)
        per_row[key] = (actual_row, pending)

    final_applied, aplicado_records = finalize_campanas_updates(actual_campanas_df, per_row)
    change_records.extend(aplicado_records)

    maestro_ids = set(actual_maestro_df["ElementoID"].dropna())
    unkeyed_new_rows = nueva_campanas_df[nueva_campanas_df[CAMPANAS_KEY].isna()].to_dict(orient="records")
    next_num = next_carga_id_sequence(actual_campanas_df)
    incorporated, rejected_new = build_new_campanas_records(
        unkeyed_new_rows, actual_campanas_df, maestro_ids, next_num, run_timestamp
    )

    return V2Plan(
        actual_maestro_df=actual_maestro_df,
        nueva_maestro_df=nueva_maestro_df,
        actual_campanas_df=actual_campanas_df,
        nueva_campanas_df=nueva_campanas_df,
        actual_parametros_df=actual_parametros_df,
        nueva_parametros_df=nueva_parametros_df,
        maestro_cls=maestro_cls,
        campanas_final_applied=final_applied,
        campanas_change_records=change_records,
        new_records_incorporated=incorporated,
        new_records_rejected=rejected_new,
        run_timestamp=run_timestamp,
    )
