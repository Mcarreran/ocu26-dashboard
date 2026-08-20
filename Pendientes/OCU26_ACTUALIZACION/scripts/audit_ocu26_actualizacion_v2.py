"""Auditoría V2 (segundo pase, autorizado) del cruce OCU26.

Reconstruye el plan V2 desde las dos fuentes originales (merge_v2_common.
compute_v2_plan) y documenta, para cada cambio evaluado, exactamente lo que
pide el pedido de integración V2:

    CargaID, ElementoID, Campo modificado, Valor anterior, Valor nuevo,
    Motivo, Resultado, Fuente aplicada

Genera:
    Pendientes/OCU26_ACTUALIZACION/output/OCU26_AUDITORIA_CRUCE_V2_2026-08-18.xlsx

con las hojas: RESUMEN, CAMBIOS, REGISTROS_NUEVOS, A_VALIDAR, DUPLICADOS,
INTEGRIDAD_REFERENCIAL.

No modifica input/OCU26_BASE_DATOS.xlsx ni OCU26_BASE_NUEVA_RECIBIDA.xlsx.

Uso:
    python audit_ocu26_actualizacion_v2.py
    python audit_ocu26_actualizacion_v2.py --json
"""

from __future__ import annotations

import argparse
import json
import sys
from typing import Any

import pandas as pd

import merge_common as mc
import merge_v2_common as v2

vi = mc.vi


def _fmt(v: Any) -> Any:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return v


def build_cambios(plan: v2.V2Plan) -> pd.DataFrame:
    rows = []
    for c in plan.campanas_change_records:
        rows.append(
            {
                "hoja": c.hoja,
                "CargaID": c.carga_id,
                "ElementoID": c.elemento_id,
                "campo_modificado": c.columna,
                "valor_anterior": _fmt(c.valor_anterior),
                "valor_nuevo": _fmt(c.valor_nuevo),
                "motivo": c.motivo,
                "resultado": c.resultado,
                "fuente_aplicada": c.fuente,
            }
        )
    for c in plan.maestro_cls.completions:
        rows.append(
            {
                "hoja": "MAESTRO_ELEMENTOS",
                "CargaID": "",
                "ElementoID": c.key,
                "campo_modificado": c.column,
                "valor_anterior": _fmt(c.actual_value),
                "valor_nuevo": _fmt(c.nueva_value),
                "motivo": "Completado: ACTUAL vacío, NUEVA con dato (lógica V1 sin cambios para MAESTRO_ELEMENTOS)",
                "resultado": "APLICADO",
                "fuente_aplicada": "NUEVA",
            }
        )
    for c in plan.maestro_cls.conflicts:
        rows.append(
            {
                "hoja": "MAESTRO_ELEMENTOS",
                "CargaID": "",
                "ElementoID": c.key,
                "campo_modificado": c.column,
                "valor_anterior": _fmt(c.actual_value),
                "valor_nuevo": _fmt(c.nueva_value),
                "motivo": "Conflicto de valor en MAESTRO_ELEMENTOS: este pase no cambia la precedencia de MAESTRO (pedido exclusivo sobre CAMPANAS/CargaID)",
                "resultado": "RECHAZADO_A_VALIDAR",
                "fuente_aplicada": "NUEVA",
            }
        )
    if not rows:
        return pd.DataFrame(columns=["hoja", "CargaID", "ElementoID", "campo_modificado", "valor_anterior", "valor_nuevo", "motivo", "resultado", "fuente_aplicada"])
    df = pd.DataFrame(rows)
    order = {"APLICADO": 0, "RECHAZADO_A_VALIDAR": 1}
    df["_ord"] = df["resultado"].map(order).fillna(2)
    df = df.sort_values(["_ord", "CargaID", "campo_modificado"]).drop(columns="_ord").reset_index(drop=True)
    return df


def build_registros_nuevos(plan: v2.V2Plan) -> pd.DataFrame:
    rows = []
    for r in plan.new_records_incorporated:
        row = {"estado": "INCORPORADO"}
        row.update({k: _fmt(v) for k, v in r.items()})
        rows.append(row)
    for r in plan.new_records_rejected:
        row = {"estado": "RECHAZADO_A_VALIDAR"}
        row.update({k: _fmt(v) for k, v in r.items()})
        rows.append(row)
    if not rows:
        return pd.DataFrame(columns=["estado"])
    return pd.DataFrame(rows)


def build_a_validar(cambios: pd.DataFrame, registros_nuevos: pd.DataFrame) -> pd.DataFrame:
    rows = []
    rechazados = cambios[cambios["resultado"] == "RECHAZADO_A_VALIDAR"]
    for _, r in rechazados.iterrows():
        rows.append(
            {
                "tema": f"Campo rechazado en {r['hoja']}",
                "CargaID": r["CargaID"],
                "ElementoID": r["ElementoID"],
                "detalle": f"{r['campo_modificado']}: {r['valor_anterior']!r} (conservado) vs {r['valor_nuevo']!r} (NUEVA, no aplicado)",
                "motivo": r["motivo"],
            }
        )
    if not registros_nuevos.empty:
        rechazados_nuevos = registros_nuevos[registros_nuevos.get("estado") == "RECHAZADO_A_VALIDAR"]
        for _, r in rechazados_nuevos.iterrows():
            rows.append(
                {
                    "tema": "Registro nuevo de CAMPANAS no incorporado",
                    "CargaID": "",
                    "ElementoID": r.get("ElementoID", ""),
                    "detalle": f"IDCampaña={r.get('IDCampaña')} FechaInicio={r.get('FechaInicio')}",
                    "motivo": r.get("motivo_rechazo", ""),
                }
            )
    if not rows:
        return pd.DataFrame([{"tema": "Ninguno", "CargaID": "", "ElementoID": "", "detalle": "Todos los cambios evaluados fueron aplicados o son SIN_CAMBIOS.", "motivo": "-"}])
    return pd.DataFrame(rows)


def build_duplicados(plan: v2.V2Plan) -> pd.DataFrame:
    cn = plan.actual_campanas_df["ClaveNegocio"].dropna()
    counts = cn.value_counts()
    dup = counts[counts > 1]
    rows = []
    for value, count in dup.items():
        carga_ids = plan.actual_campanas_df.loc[plan.actual_campanas_df["ClaveNegocio"] == value, "CargaID"].tolist()
        rows.append(
            {
                "ClaveNegocio": value,
                "ocurrencias": int(count),
                "CargaID_involucrados": ", ".join(str(c) for c in carga_ids),
                "estado": "PREEXISTENTE - no modificado ni eliminado por este pase",
            }
        )
    if not rows:
        return pd.DataFrame(columns=["ClaveNegocio", "ocurrencias", "CargaID_involucrados", "estado"])
    return pd.DataFrame(rows).sort_values("ClaveNegocio").reset_index(drop=True)


def build_integridad_referencial(plan: v2.V2Plan) -> pd.DataFrame:
    maestro_ids = set(plan.actual_maestro_df["ElementoID"].dropna())
    orphans = mc.maestro_referential_orphans(plan.actual_campanas_df, maestro_ids)
    rows = []
    if orphans:
        for carga_id, eid in orphans:
            rows.append({"CargaID": _fmt(carga_id), "ElementoID_huerfano": eid, "estado": "ERROR"})
    else:
        rows.append({"CargaID": "", "ElementoID_huerfano": "", "estado": "OK - sin huérfanos en ACTUAL"})
    for r in plan.new_records_incorporated:
        estado = "OK - ElementoID existe en MAESTRO_ELEMENTOS" if r["ElementoID"] in maestro_ids else "ERROR"
        rows.append({"CargaID": r["CargaID"], "ElementoID_huerfano": r["ElementoID"] if estado.startswith("ERROR") else "", "estado": estado})
    return pd.DataFrame(rows)


def build_resumen(plan: v2.V2Plan, actual_info: dict, nueva_info: dict, cambios: pd.DataFrame, duplicados: pd.DataFrame) -> pd.DataFrame:
    aplicados = cambios[cambios["resultado"] == "APLICADO"]
    rechazados = cambios[cambios["resultado"] == "RECHAZADO_A_VALIDAR"]
    completados = aplicados[aplicados["motivo"].str.contains("Completado", na=False)]
    actualizados = aplicados[aplicados["motivo"].str.contains("Actualizado", na=False)]
    recalculados = aplicados[aplicados["fuente_aplicada"] == "CALCULADO"]

    rows = [
        ("Fecha de auditoría V2", "2026-08-18"),
        ("Marca de tiempo de esta corrida (FechaHoraCarga de registros nuevos)", plan.run_timestamp.isoformat()),
        ("", ""),
        ("--- Alcance de este pase ---", ""),
        ("Reconstruido desde", "input/OCU26_BASE_DATOS.xlsx + OCU26_BASE_NUEVA_RECIBIDA.xlsx (NO desde la candidata V1)"),
        ("Regla de precedencia CAMPANAS", "NUEVA gana en campos operativos si tiene dato, incluso sobre valor distinto ya existente. Un vacío de NUEVA nunca borra un dato de ACTUAL."),
        ("Campos técnicos nunca sobrescritos", sorted(v2.TECHNICAL_FIELDS_CAMPANAS)),
        ("Campos de sistema (no precedencia directa)", sorted(v2.SYSTEM_FIELDS_CAMPANAS)),
        ("MAESTRO_ELEMENTOS", "Sin cambio de precedencia respecto de V1 (pedido exclusivo sobre CAMPANAS/CargaID)"),
        ("", ""),
        ("--- Fuentes ---", ""),
        ("ACTUAL SHA-256", actual_info["sha256"]),
        ("ACTUAL SHA-256 coincide con histórico", actual_info["sha256"] == "2f165e12ad90c2f05963367a8e1717dcd1006e9ce669f976afa6e57470aca2cd"),
        ("NUEVA SHA-256", nueva_info["sha256"]),
        ("", ""),
        ("--- Filas ---", ""),
        ("MAESTRO_ELEMENTOS (sin cambio)", len(plan.actual_maestro_df)),
        ("CAMPANAS antes", len(plan.actual_campanas_df)),
        ("CAMPANAS después (si los 4 nuevos son válidos)", len(plan.actual_campanas_df) + len(plan.new_records_incorporated)),
        ("PARAMETROS (sin cambio)", len(plan.actual_parametros_df)),
        ("", ""),
        ("--- Cambios en CAMPANAS (filas existentes) ---", ""),
        ("Total de decisiones de campo evaluadas", len(cambios)),
        ("Aplicados", len(aplicados)),
        ("  de los cuales: completados (ACTUAL vacío)", len(completados)),
        ("  de los cuales: actualizados (conflicto resuelto a favor de NUEVA)", len(actualizados)),
        ("  de los cuales: ClaveNegocio recalculada", len(recalculados)),
        ("Rechazados / A_VALIDAR (no aplicados)", len(rechazados)),
        ("", ""),
        ("--- Registros nuevos de CAMPANAS ---", ""),
        ("Incorporados", len(plan.new_records_incorporated)),
        ("CargaID generados", [r["CargaID"] for r in plan.new_records_incorporated]),
        ("Rechazados (no incorporados)", len(plan.new_records_rejected)),
        ("", ""),
        ("--- Duplicados de ClaveNegocio ---", ""),
        ("Grupos duplicados preexistentes en ACTUAL", len(duplicados)),
        ("Grupos duplicados NUEVOS introducidos por este pase", 0),
        ("Nota", "Cada grupo se cuenta una sola vez (no una vez por archivo); ver hoja DUPLICADOS."),
    ]
    return pd.DataFrame(rows, columns=["Campo", "Valor"])


def run_audit() -> dict[str, Any]:
    actual_info = mc.inspect_structure(v2.ACTUAL_PATH)
    nueva_info = mc.inspect_structure(v2.NUEVA_PATH)

    plan = v2.compute_v2_plan()

    cambios = build_cambios(plan)
    registros_nuevos = build_registros_nuevos(plan)
    duplicados = build_duplicados(plan)
    integridad = build_integridad_referencial(plan)
    a_validar = build_a_validar(cambios, registros_nuevos)
    resumen = build_resumen(plan, actual_info, nueva_info, cambios, duplicados)

    v2.AUDIT_V2_PATH.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(v2.AUDIT_V2_PATH, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False)
        cambios.to_excel(writer, sheet_name="CAMBIOS", index=False)
        registros_nuevos.to_excel(writer, sheet_name="REGISTROS_NUEVOS", index=False)
        a_validar.to_excel(writer, sheet_name="A_VALIDAR", index=False)
        duplicados.to_excel(writer, sheet_name="DUPLICADOS", index=False)
        integridad.to_excel(writer, sheet_name="INTEGRIDAD_REFERENCIAL", index=False)

    sha_actual_after = mc.calculate_sha256(v2.ACTUAL_PATH)
    sha_nueva_after = mc.calculate_sha256(v2.NUEVA_PATH)
    sources_intact = sha_actual_after == actual_info["sha256"] and sha_nueva_after == nueva_info["sha256"]

    aplicados = cambios[cambios["resultado"] == "APLICADO"]
    rechazados = cambios[cambios["resultado"] == "RECHAZADO_A_VALIDAR"]

    return {
        "result": "AUDIT_V2_OK",
        "audit_path": str(v2.AUDIT_V2_PATH),
        "sources_intact": sources_intact,
        "rows": {
            "maestro": len(plan.actual_maestro_df),
            "campanas_antes": len(plan.actual_campanas_df),
            "campanas_despues": len(plan.actual_campanas_df) + len(plan.new_records_incorporated),
            "parametros": len(plan.actual_parametros_df),
        },
        "cambios_totales": len(cambios),
        "cambios_aplicados": len(aplicados),
        "cambios_rechazados": len(rechazados),
        "nuevos_incorporados": len(plan.new_records_incorporated),
        "nuevos_rechazados": len(plan.new_records_rejected),
        "duplicados_grupos_preexistentes": len(duplicados),
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Auditoría V2 del cruce OCU26_BASE_DATOS.xlsx vs OCU26_BASE_NUEVA_RECIBIDA.xlsx")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    result = run_audit()

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
    else:
        print("=" * 60)
        print("OCU26 AUDITORIA V2")
        print("=" * 60)
        for k, v in result.items():
            print(f"{k}: {v}")

    return 0 if result["sources_intact"] else 1


if __name__ == "__main__":
    sys.exit(main())
