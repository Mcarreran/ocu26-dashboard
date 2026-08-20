"""FASE 1-5: preflight, mapeo de columnas y auditoría de cruce OCU26.

Compara (solo lectura, nunca escribe ninguno de los dos archivos fuente):
  - input/OCU26_BASE_DATOS.xlsx                                  (ACTUAL, canónica)
  - Pendientes/OCU26_ACTUALIZACION/input/OCU26_BASE_NUEVA_RECIBIDA.xlsx (NUEVA)

Genera:
  Pendientes/OCU26_ACTUALIZACION/output/OCU26_AUDITORIA_CRUCE_2026-08-18.xlsx

con las hojas: RESUMEN, MAPEO_COLUMNAS, CAMBIOS_SEGUROS, REGISTROS_NUEVOS,
SOLO_BASE_ACTUAL, CONFLICTOS, FALTANTES, DUPLICADOS, INTEGRIDAD_REFERENCIAL,
COLUMNAS_NUEVAS, A_VALIDAR.

Uso:
    python audit_ocu26_actualizacion.py
    python audit_ocu26_actualizacion.py --json
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

import pandas as pd

import merge_common as mc

vi = mc.vi


def _fmt(v: Any) -> Any:
    """Normaliza valores para volcar en el Excel de auditoría (texto legible)."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return v


def build_resumen(
    actual_info: dict[str, Any],
    nueva_info: dict[str, Any],
    actual_validation: dict[str, Any],
    maestro_cls: mc.SheetClassification,
    campanas_cls: mc.SheetClassification,
    parametros_cls: dict[str, Any],
) -> pd.DataFrame:
    rows: list[tuple[str, Any]] = [
        ("Fecha de auditoría", "2026-08-18"),
        ("", ""),
        ("--- Archivos ---", ""),
        ("ACTUAL (canónica, no modificar)", str(mc.ACTUAL_PATH)),
        ("ACTUAL SHA-256", actual_info["sha256"]),
        ("ACTUAL SHA-256 esperado histórico (docs/CM1.md)", "2f165e12ad90c2f05963367a8e1717dcd1006e9ce669f976afa6e57470aca2cd"),
        ("ACTUAL SHA-256 coincide con histórico", actual_info["sha256"] == "2f165e12ad90c2f05963367a8e1717dcd1006e9ce669f976afa6e57470aca2cd"),
        ("ACTUAL size (bytes)", actual_info["size_bytes"]),
        ("NUEVA (solo lectura)", str(mc.NUEVA_PATH)),
        ("NUEVA SHA-256", nueva_info["sha256"]),
        ("NUEVA size (bytes)", nueva_info["size_bytes"]),
        ("", ""),
        ("--- Integridad ZIP/XLSX ---", ""),
        ("ACTUAL zip_ok", actual_info["zip_ok"]),
        ("ACTUAL vba", actual_info["vba"]),
        ("ACTUAL external_links", actual_info["external_links"]),
        ("ACTUAL macro_enabled", actual_info["macro_enabled"]),
        ("ACTUAL formula_cells", actual_info["formula_cells"]),
        ("ACTUAL error_cells", actual_info["error_cells"]),
        ("NUEVA zip_ok", nueva_info["zip_ok"]),
        ("NUEVA vba", nueva_info["vba"]),
        ("NUEVA external_links", nueva_info["external_links"]),
        ("NUEVA macro_enabled", nueva_info["macro_enabled"]),
        ("NUEVA formula_cells", nueva_info["formula_cells"]),
        ("NUEVA error_cells", nueva_info["error_cells"]),
        ("", ""),
        ("--- Hojas y tablas ---", ""),
        ("ACTUAL sheets (orden)", actual_info["sheets"]),
        ("NUEVA sheets (orden)", nueva_info["sheets"]),
        ("Sheets orden idéntico", actual_info["sheets"] == nueva_info["sheets"]),
        ("ACTUAL defined_names", actual_info["defined_names"]),
        ("NUEVA defined_names", nueva_info["defined_names"]),
        ("", ""),
        ("--- Validación Gate 1 sobre ACTUAL (validate_input.py) ---", ""),
        ("ACTUAL validate_input result", actual_validation["result"]),
        ("ACTUAL errors", len(actual_validation["errors"])),
        ("ACTUAL warnings", len(actual_validation["warnings"])),
        ("", ""),
        ("--- Filas por hoja: ACTUAL vs NUEVA ---", ""),
        ("MAESTRO_ELEMENTOS filas ACTUAL", actual_info["sheet_detail"]["MAESTRO_ELEMENTOS"]["tables"]["tblElementos"]["data_rows"]),
        ("MAESTRO_ELEMENTOS filas NUEVA", nueva_info["sheet_detail"]["MAESTRO_ELEMENTOS"]["tables"]["tblElementos"]["data_rows"]),
        ("Snapshot histórico maestro (docs/CM1.md)", mc.HISTORICAL_SNAPSHOT["maestro_rows"]),
        ("CAMPANAS filas ACTUAL", actual_info["sheet_detail"]["CAMPANAS"]["tables"]["tblCampanas"]["data_rows"]),
        ("CAMPANAS filas NUEVA", nueva_info["sheet_detail"]["CAMPANAS"]["tables"]["tblCampanas"]["data_rows"]),
        ("Snapshot histórico campañas (docs/CM1.md)", mc.HISTORICAL_SNAPSHOT["campanas_rows"]),
        ("PARAMETROS filas ACTUAL", actual_info["sheet_detail"]["PARAMETROS"]["tables"]["tblParametros"]["data_rows"]),
        ("PARAMETROS filas NUEVA", nueva_info["sheet_detail"]["PARAMETROS"]["tables"]["tblParametros"]["data_rows"]),
        ("Snapshot histórico parámetros (docs/CM1.md)", mc.HISTORICAL_SNAPSHOT["parametros_rows"]),
        ("", ""),
        ("--- MAESTRO_ELEMENTOS: resultado del cruce ---", ""),
        ("Filas comunes (ElementoID en ambas)", maestro_cls.rows_common),
        ("Sin cambios", maestro_cls.rows_no_change),
        ("Completados de forma segura (campo)", len(maestro_cls.completions)),
        ("Conflictos de valor (campo, no sobrescritos)", len(maestro_cls.conflicts)),
        ("Solo en ACTUAL (conservados)", len(maestro_cls.only_actual_keys)),
        ("Solo en NUEVA con ElementoID válido (no incorporados automáticamente)", len(maestro_cls.only_nueva_keyed_new)),
        ("Solo en NUEVA sin ElementoID (no incorporables)", len(maestro_cls.only_nueva_unkeyed_rows)),
        ("ElementoID duplicados en ACTUAL (excluidos del cruce por campo)", len(maestro_cls.duplicates_actual)),
        ("ElementoID duplicados en NUEVA (excluidos del cruce por campo)", len(maestro_cls.duplicates_nueva)),
        ("", ""),
        ("--- CAMPANAS: resultado del cruce ---", ""),
        ("Filas comunes (CargaID en ambas)", campanas_cls.rows_common),
        ("Sin cambios", campanas_cls.rows_no_change),
        ("Completados de forma segura (campo)", len(campanas_cls.completions)),
        ("Conflictos de valor (campo, no sobrescritos)", len(campanas_cls.conflicts)),
        ("Solo en ACTUAL (conservados)", len(campanas_cls.only_actual_keys)),
        ("Solo en NUEVA con CargaID válido (no incorporados automáticamente)", len(campanas_cls.only_nueva_keyed_new)),
        ("Solo en NUEVA sin CargaID (no incorporables sin generación de ID)", len(campanas_cls.only_nueva_unkeyed_rows)),
        ("CargaID duplicados en ACTUAL (excluidos del cruce por campo)", len(campanas_cls.duplicates_actual)),
        ("CargaID duplicados en NUEVA (excluidos del cruce por campo)", len(campanas_cls.duplicates_nueva)),
        ("", ""),
        ("--- PARAMETROS: resultado del cruce ---", ""),
        ("Idéntico ACTUAL vs NUEVA", parametros_cls["identical"]),
        ("Filas solo en ACTUAL", len(parametros_cls["only_actual"])),
        ("Filas solo en NUEVA", len(parametros_cls["only_nueva"])),
        ("", ""),
        ("--- Decisión de arquitectura ---", ""),
        (
            "CargaID para registros nuevos",
            "NO generado automáticamente: el proyecto no define una función de "
            "generación de CargaID, solo un patrón histórico documentado "
            "('HIST-00000001', ver docs/CM1.md sección 11). Inventar esa lógica "
            "sin evidencia de código está fuera de alcance. Los registros nuevos "
            "sin CargaID quedan en REGISTROS_NUEVOS/A_VALIDAR y NO se incorporan "
            "a la base candidata.",
        ),
    ]
    return pd.DataFrame(rows, columns=["Campo", "Valor"])


def build_mapeo_columnas() -> pd.DataFrame:
    frames = []
    specs = [
        ("MAESTRO_ELEMENTOS", mc.MAESTRO_HEADERS, mc.MAESTRO_HEADERS),
        ("CAMPANAS", mc.CAMPANAS_HEADERS, mc.CAMPANAS_HEADERS),
        ("PARAMETROS", mc.PARAMETROS_HEADERS, mc.PARAMETROS_HEADERS),
    ]
    for sheet, h_a, h_b in specs:
        rows = mc.column_mapping(h_a, h_b)
        df = pd.DataFrame(rows)
        df.insert(0, "hoja", sheet)
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def build_cambios_seguros(maestro_cls: mc.SheetClassification, campanas_cls: mc.SheetClassification) -> pd.DataFrame:
    rows = []
    for cls, sheet, key_name in [
        (maestro_cls, "MAESTRO_ELEMENTOS", "ElementoID"),
        (campanas_cls, "CAMPANAS", "CargaID"),
    ]:
        for c in cls.completions:
            rows.append(
                {
                    "hoja": sheet,
                    key_name: c.key,
                    "columna": c.column,
                    "valor_actual": _fmt(c.actual_value),
                    "valor_nueva": _fmt(c.nueva_value),
                    "accion": "COMPLETAR_VACIO_ACTUAL",
                }
            )
    if not rows:
        return pd.DataFrame(columns=["hoja", "clave", "columna", "valor_actual", "valor_nueva", "accion"])
    return pd.DataFrame(rows)


def build_conflictos(maestro_cls: mc.SheetClassification, campanas_cls: mc.SheetClassification) -> pd.DataFrame:
    rows = []
    for cls, sheet, key_name in [
        (maestro_cls, "MAESTRO_ELEMENTOS", "ElementoID"),
        (campanas_cls, "CAMPANAS", "CargaID"),
    ]:
        for c in cls.conflicts:
            rows.append(
                {
                    "hoja": sheet,
                    key_name: c.key,
                    "columna": c.column,
                    "valor_actual_CONSERVADO": _fmt(c.actual_value),
                    "valor_nueva_NO_APLICADO": _fmt(c.nueva_value),
                    "accion": "CONFLICTO_VALOR - no sobrescrito",
                }
            )
    if not rows:
        return pd.DataFrame(columns=["hoja", "clave", "columna", "valor_actual_CONSERVADO", "valor_nueva_NO_APLICADO", "accion"])
    return pd.DataFrame(rows)


def build_registros_nuevos(maestro_cls: mc.SheetClassification, campanas_cls: mc.SheetClassification) -> pd.DataFrame:
    rows = []
    for row in campanas_cls.only_nueva_unkeyed_rows:
        r = {"hoja": "CAMPANAS", "tipo": "SIN_CARGAID_NO_INCORPORADO"}
        r.update({k: _fmt(v) for k, v in row.items()})
        rows.append(r)
    for eid in campanas_cls.only_nueva_keyed_new:
        rows.append({"hoja": "CAMPANAS", "tipo": "CARGAID_NUEVO_NO_INCORPORADO_AUTOMATICO", "CargaID": eid})
    for eid in maestro_cls.only_nueva_unkeyed_rows:
        rows.append({"hoja": "MAESTRO_ELEMENTOS", "tipo": "SIN_ELEMENTOID_NO_INCORPORADO", **{k: _fmt(v) for k, v in eid.items()}})
    for eid in maestro_cls.only_nueva_keyed_new:
        rows.append({"hoja": "MAESTRO_ELEMENTOS", "tipo": "ELEMENTOID_NUEVO_A_VALIDAR", "ElementoID": eid})
    if not rows:
        return pd.DataFrame(columns=["hoja", "tipo"])
    return pd.DataFrame(rows)


def build_solo_base_actual(actual_maestro: pd.DataFrame, actual_campanas: pd.DataFrame, maestro_cls: mc.SheetClassification, campanas_cls: mc.SheetClassification) -> pd.DataFrame:
    rows = []
    if campanas_cls.only_actual_keys:
        sub = actual_campanas[actual_campanas["CargaID"].isin(campanas_cls.only_actual_keys)]
        for _, row in sub.iterrows():
            r = {"hoja": "CAMPANAS", "accion": "CONSERVADO_SIN_CAMBIOS"}
            r.update({k: _fmt(v) for k, v in row.items()})
            rows.append(r)
    if maestro_cls.only_actual_keys:
        sub = actual_maestro[actual_maestro["ElementoID"].isin(maestro_cls.only_actual_keys)]
        for _, row in sub.iterrows():
            r = {"hoja": "MAESTRO_ELEMENTOS", "accion": "CONSERVADO_SIN_CAMBIOS"}
            r.update({k: _fmt(v) for k, v in row.items()})
            rows.append(r)
    if not rows:
        return pd.DataFrame(columns=["hoja", "accion"])
    return pd.DataFrame(rows)


REQUIRED_MAESTRO = {"ElementoID", "TipoCatalogo", "Ciudad", "Medio", "Ubicacion", "TipoInventario", "AplicaCantidad"}
REQUIRED_CAMPANAS = {"CargaID", "ClaveNegocio", "ElementoID", "TipoCargaDeclarado", "EstadoValidacion"}


def build_faltantes(
    actual_maestro: pd.DataFrame,
    actual_campanas: pd.DataFrame,
    actual_parametros: pd.DataFrame,
    maestro_cls: mc.SheetClassification,
    campanas_cls: mc.SheetClassification,
) -> pd.DataFrame:
    from collections import Counter

    completions_by_col_m = Counter(c.column for c in maestro_cls.completions)
    conflicts_by_col_m = Counter(c.column for c in maestro_cls.conflicts)
    completions_by_col_c = Counter(c.column for c in campanas_cls.completions)
    conflicts_by_col_c = Counter(c.column for c in campanas_cls.conflicts)

    rows = []
    specs = [
        ("MAESTRO_ELEMENTOS", actual_maestro, REQUIRED_MAESTRO, completions_by_col_m, conflicts_by_col_m),
        ("CAMPANAS", actual_campanas, REQUIRED_CAMPANAS, completions_by_col_c, conflicts_by_col_c),
        ("PARAMETROS", actual_parametros, {"Categoria", "Valor"}, Counter(), Counter()),
    ]
    for sheet, df, required, completions_by_col, conflicts_by_col in specs:
        total = len(df)
        for col in df.columns:
            vacios = int(df[col].apply(mc.is_blank).sum())
            completos = total - vacios
            pct = round(100.0 * completos / total, 2) if total else 0.0
            rows.append(
                {
                    "hoja": sheet,
                    "columna": col,
                    "obligatoriedad": "OBLIGATORIO" if col in required else "OPCIONAL",
                    "registros_totales": total,
                    "completos": completos,
                    "vacios": vacios,
                    "pct_completo": pct,
                    "aportados_por_base_nueva": int(completions_by_col.get(col, 0)),
                    "conflictos_pendientes": int(conflicts_by_col.get(col, 0)),
                }
            )
    return pd.DataFrame(rows)


def build_duplicados(
    actual_maestro: pd.DataFrame,
    actual_campanas: pd.DataFrame,
    nueva_campanas: pd.DataFrame,
    maestro_cls: mc.SheetClassification,
    campanas_cls: mc.SheetClassification,
) -> pd.DataFrame:
    rows = []
    for k, v in maestro_cls.duplicates_actual.items():
        rows.append({"hoja": "MAESTRO_ELEMENTOS", "campo": "ElementoID", "archivo": "ACTUAL", "valor": _fmt(k), "ocurrencias": v})
    for k, v in maestro_cls.duplicates_nueva.items():
        rows.append({"hoja": "MAESTRO_ELEMENTOS", "campo": "ElementoID", "archivo": "NUEVA", "valor": _fmt(k), "ocurrencias": v})
    for k, v in campanas_cls.duplicates_actual.items():
        rows.append({"hoja": "CAMPANAS", "campo": "CargaID", "archivo": "ACTUAL", "valor": _fmt(k), "ocurrencias": v})
    for k, v in campanas_cls.duplicates_nueva.items():
        rows.append({"hoja": "CAMPANAS", "campo": "CargaID", "archivo": "NUEVA", "valor": _fmt(k), "ocurrencias": v})

    # ClaveNegocio: clave de negocio compuesta (no unicidad estricta garantizada, pero se reporta)
    for label, df in [("ACTUAL", actual_campanas), ("NUEVA", nueva_campanas)]:
        dup = df["ClaveNegocio"][~df["ClaveNegocio"].apply(mc.is_blank)]
        counts = dup.value_counts()
        for k, v in counts[counts > 1].items():
            rows.append({"hoja": "CAMPANAS", "campo": "ClaveNegocio", "archivo": label, "valor": _fmt(k), "ocurrencias": int(v)})

    if not rows:
        return pd.DataFrame(columns=["hoja", "campo", "archivo", "valor", "ocurrencias"])
    return pd.DataFrame(rows).sort_values(["hoja", "campo", "archivo", "valor"]).reset_index(drop=True)


def build_integridad_referencial(actual_maestro: pd.DataFrame, actual_campanas: pd.DataFrame, nueva_maestro: pd.DataFrame, nueva_campanas: pd.DataFrame) -> pd.DataFrame:
    rows = []
    actual_ids = set(actual_maestro["ElementoID"].dropna())
    nueva_ids = set(nueva_maestro["ElementoID"].dropna())
    for label, campanas_df, ids in [("ACTUAL", actual_campanas, actual_ids), ("NUEVA", nueva_campanas, nueva_ids)]:
        orphans = mc.maestro_referential_orphans(campanas_df, ids)
        if not orphans:
            rows.append({"archivo": label, "CargaID": "", "ElementoID_huerfano": "", "estado": "OK - sin huérfanos"})
        for carga_id, eid in orphans:
            rows.append({"archivo": label, "CargaID": _fmt(carga_id), "ElementoID_huerfano": eid, "estado": "ERROR - ElementoID no existe en MAESTRO_ELEMENTOS"})
    return pd.DataFrame(rows)


def build_columnas_nuevas(mapeo: pd.DataFrame) -> pd.DataFrame:
    nuevas = mapeo[mapeo["clasificacion"] == "SOLO_BASE_NUEVA"].copy()
    if nuevas.empty:
        return pd.DataFrame(
            [{"hoja": "-", "columna": "Ninguna", "nota": "La base nueva no aporta columnas fuera de la estructura canónica."}]
        )
    nuevas["nota"] = "No incorporada a la candidata (podría romper scripts/formulario/tableros); documentar antes de decidir."
    return nuevas[["hoja", "columna", "nota"]]


def build_a_validar(
    campanas_cls: mc.SheetClassification,
    maestro_cls: mc.SheetClassification,
    duplicados: pd.DataFrame,
    integridad: pd.DataFrame,
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []

    for row in campanas_cls.only_nueva_unkeyed_rows:
        rows.append(
            {
                "tema": "Registro nuevo en CAMPANAS sin CargaID",
                "detalle": (
                    f"IDCampaña={_fmt(row.get('IDCampaña'))} ElementoID={_fmt(row.get('ElementoID'))} "
                    f"FechaInicio={_fmt(row.get('FechaInicio'))} FechaFin={_fmt(row.get('FechaFin'))}"
                ),
                "motivo": (
                    "Faltan CargaID, ClaveNegocio, EstadoValidacion, FechaHoraCarga, UsuarioCarga y "
                    "FuenteCarga. El proyecto no define una función de generación de CargaID (solo un "
                    "patrón histórico documentado). No se incorpora a la candidata sin decisión humana."
                ),
                "accion_recomendada": "Asignar CargaID/ClaveNegocio/EstadoValidacion siguiendo el proceso vigente y volver a correr el cruce.",
            }
        )

    for c in campanas_cls.conflicts:
        rows.append(
            {
                "tema": "Conflicto de valor en CAMPANAS",
                "detalle": f"CargaID={c.key} columna={c.column} actual={_fmt(c.actual_value)!r} nueva={_fmt(c.nueva_value)!r}",
                "motivo": "Ambas bases tienen valores distintos y no vacíos; no se sobrescribe automáticamente.",
                "accion_recomendada": "Confirmar cuál valor es correcto con el área de negocio antes de aplicar.",
            }
        )

    for c in maestro_cls.conflicts:
        rows.append(
            {
                "tema": "Conflicto de valor en MAESTRO_ELEMENTOS",
                "detalle": f"ElementoID={c.key} columna={c.column} actual={_fmt(c.actual_value)!r} nueva={_fmt(c.nueva_value)!r}",
                "motivo": "Ambas bases tienen valores distintos y no vacíos; no se sobrescribe automáticamente.",
                "accion_recomendada": "Confirmar cuál valor es correcto con el área de negocio antes de aplicar.",
            }
        )

    if not duplicados.empty:
        rows.append(
            {
                "tema": "Duplicados de ClaveNegocio",
                "detalle": f"{len(duplicados[duplicados['campo'] == 'ClaveNegocio'])} grupo(s) reportado(s) en hoja DUPLICADOS",
                "motivo": "Preexistentes en ambas bases; no introducidos por este cruce. No se modifican.",
                "accion_recomendada": "Revisar si corresponde a un problema histórico ya conocido (ver docs/CM1.md).",
            }
        )

    error_rows = integridad[integridad["estado"].astype(str).str.startswith("ERROR")]
    if not error_rows.empty:
        rows.append(
            {
                "tema": "Integridad referencial",
                "detalle": f"{len(error_rows)} ElementoID huérfano(s) detectado(s)",
                "motivo": "CAMPANAS referencia ElementoID inexistente en MAESTRO_ELEMENTOS.",
                "accion_recomendada": "Bloqueante: no incorporar filas afectadas hasta resolver.",
            }
        )

    if not rows:
        return pd.DataFrame([{"tema": "Ninguno", "detalle": "-", "motivo": "-", "accion_recomendada": "-"}])
    return pd.DataFrame(rows)


def run_audit() -> dict[str, Any]:
    actual_info = mc.inspect_structure(mc.ACTUAL_PATH)
    nueva_info = mc.inspect_structure(mc.NUEVA_PATH)
    actual_validation = vi.validate_input(mc.ACTUAL_PATH)

    actual_maestro = mc.read_table_df(mc.ACTUAL_PATH, "MAESTRO_ELEMENTOS", "tblElementos")
    nueva_maestro = mc.read_table_df(mc.NUEVA_PATH, "MAESTRO_ELEMENTOS", "tblElementos")
    actual_campanas = mc.read_table_df(mc.ACTUAL_PATH, "CAMPANAS", "tblCampanas")
    nueva_campanas = mc.read_table_df(mc.NUEVA_PATH, "CAMPANAS", "tblCampanas")
    actual_parametros = mc.read_table_df(mc.ACTUAL_PATH, "PARAMETROS", "tblParametros")
    nueva_parametros = mc.read_table_df(mc.NUEVA_PATH, "PARAMETROS", "tblParametros")

    maestro_cls = mc.classify_maestro(actual_maestro, nueva_maestro)
    campanas_cls = mc.classify_campanas(actual_campanas, nueva_campanas)
    parametros_cls = mc.classify_parametros(actual_parametros, nueva_parametros)

    mapeo = build_mapeo_columnas()
    resumen = build_resumen(actual_info, nueva_info, actual_validation, maestro_cls, campanas_cls, parametros_cls)
    cambios_seguros = build_cambios_seguros(maestro_cls, campanas_cls)
    registros_nuevos = build_registros_nuevos(maestro_cls, campanas_cls)
    solo_base_actual = build_solo_base_actual(actual_maestro, actual_campanas, maestro_cls, campanas_cls)
    conflictos = build_conflictos(maestro_cls, campanas_cls)
    faltantes = build_faltantes(actual_maestro, actual_campanas, actual_parametros, maestro_cls, campanas_cls)
    duplicados = build_duplicados(actual_maestro, actual_campanas, nueva_campanas, maestro_cls, campanas_cls)
    integridad = build_integridad_referencial(actual_maestro, actual_campanas, nueva_maestro, nueva_campanas)
    columnas_nuevas = build_columnas_nuevas(mapeo)
    a_validar = build_a_validar(campanas_cls, maestro_cls, duplicados, integridad)

    mc.AUDIT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(mc.AUDIT_PATH, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False)
        mapeo.to_excel(writer, sheet_name="MAPEO_COLUMNAS", index=False)
        cambios_seguros.to_excel(writer, sheet_name="CAMBIOS_SEGUROS", index=False)
        registros_nuevos.to_excel(writer, sheet_name="REGISTROS_NUEVOS", index=False)
        solo_base_actual.to_excel(writer, sheet_name="SOLO_BASE_ACTUAL", index=False)
        conflictos.to_excel(writer, sheet_name="CONFLICTOS", index=False)
        faltantes.to_excel(writer, sheet_name="FALTANTES", index=False)
        duplicados.to_excel(writer, sheet_name="DUPLICADOS", index=False)
        integridad.to_excel(writer, sheet_name="INTEGRIDAD_REFERENCIAL", index=False)
        columnas_nuevas.to_excel(writer, sheet_name="COLUMNAS_NUEVAS", index=False)
        a_validar.to_excel(writer, sheet_name="A_VALIDAR", index=False)

    # Confirmar que las fuentes no fueron tocadas por esta auditoría.
    sha_actual_after = mc.calculate_sha256(mc.ACTUAL_PATH)
    sha_nueva_after = mc.calculate_sha256(mc.NUEVA_PATH)
    sources_intact = sha_actual_after == actual_info["sha256"] and sha_nueva_after == nueva_info["sha256"]

    return {
        "result": "AUDIT_OK",
        "audit_path": str(mc.AUDIT_PATH),
        "actual_sha256": actual_info["sha256"],
        "nueva_sha256": nueva_info["sha256"],
        "sources_intact": sources_intact,
        "actual_validate_input_result": actual_validation["result"],
        "rows": {
            "maestro_actual": len(actual_maestro),
            "maestro_nueva": len(nueva_maestro),
            "campanas_actual": len(actual_campanas),
            "campanas_nueva": len(nueva_campanas),
            "parametros_actual": len(actual_parametros),
            "parametros_nueva": len(nueva_parametros),
        },
        "maestro": {
            "common": maestro_cls.rows_common,
            "no_change": maestro_cls.rows_no_change,
            "completions": len(maestro_cls.completions),
            "conflicts": len(maestro_cls.conflicts),
            "only_actual": len(maestro_cls.only_actual_keys),
            "only_nueva_unkeyed": len(maestro_cls.only_nueva_unkeyed_rows),
            "only_nueva_keyed": len(maestro_cls.only_nueva_keyed_new),
        },
        "campanas": {
            "common": campanas_cls.rows_common,
            "no_change": campanas_cls.rows_no_change,
            "completions": len(campanas_cls.completions),
            "conflicts": len(campanas_cls.conflicts),
            "only_actual": len(campanas_cls.only_actual_keys),
            "only_nueva_unkeyed": len(campanas_cls.only_nueva_unkeyed_rows),
            "only_nueva_keyed": len(campanas_cls.only_nueva_keyed_new),
        },
        "parametros": parametros_cls,
        "duplicados_count": len(duplicados),
        "integridad_referencial_errores": int(integridad["estado"].astype(str).str.startswith("ERROR").sum()),
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Audita el cruce OCU26_BASE_DATOS.xlsx vs OCU26_BASE_NUEVA_RECIBIDA.xlsx")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    result = run_audit()

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
    else:
        print("=" * 60)
        print("OCU26 AUDITORIA DE CRUCE")
        print("=" * 60)
        print(f"Auditoría generada: {result['audit_path']}")
        print(f"Fuentes intactas: {result['sources_intact']}")
        print(f"ACTUAL validate_input: {result['actual_validate_input_result']}")
        print("Rows:", result["rows"])
        print("MAESTRO_ELEMENTOS:", result["maestro"])
        print("CAMPANAS:", result["campanas"])
        print("PARAMETROS:", result["parametros"])
        print("Duplicados:", result["duplicados_count"])
        print("Errores integridad referencial:", result["integridad_referencial_errores"])

    return 0 if result["sources_intact"] else 1


if __name__ == "__main__":
    sys.exit(main())
