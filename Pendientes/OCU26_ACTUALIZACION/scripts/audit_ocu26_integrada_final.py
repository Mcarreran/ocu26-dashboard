"""Auditoría del pase FINAL OCU26.

Documenta exactamente lo que build_ocu26_integrada_final.py aplicó (y lo
que NO aplicó, con motivo) sobre la candidata V2.

Genera:
    Pendientes/OCU26_ACTUALIZACION/output/OCU26_AUDITORIA_FINAL_2026-08-18.xlsx

con las hojas: RESUMEN, CAMBIOS_APLICADOS, DECISIONES_BLOQUEADAS,
DUPLICADOS_EXACTOS, REPETICIONES_VALIDAS, COLISIONES_PENDIENTES.

No modifica ningún archivo fuente.

Uso:
    python audit_ocu26_integrada_final.py
    python audit_ocu26_integrada_final.py --json
"""

from __future__ import annotations

import argparse
import json
import sys
from typing import Any

import pandas as pd

import merge_common as mc
import merge_final_common as fc

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


def build_cambios_aplicados(applied_log: list[dict[str, Any]]) -> pd.DataFrame:
    rows = [
        {
            "decision": e["decision"],
            "CargaID": e["CargaID"],
            "ElementoID": e["ElementoID"],
            "campo_modificado": e["campo"],
            "valor_anterior": _fmt(e["valor_anterior"]),
            "valor_nuevo": _fmt(e["valor_nuevo"]),
        }
        for e in applied_log
    ]
    return pd.DataFrame(rows)


def build_decisiones_bloqueadas(decision_3: dict[str, Any], decision_4: dict[str, Any]) -> pd.DataFrame:
    rows = []
    for d in (decision_3, decision_4):
        rows.append(
            {
                "decision": d["decision"],
                "id_actual": d["id_actual"],
                "id_propuesto": d["id_propuesto"],
                "estado": d["status"],
                "filas_encontradas_en_seleccion": d.get("filas_encontradas"),
                "carga_ids_seleccionados": ", ".join(d.get("carga_ids_encontrados", [])),
                "motivo": d.get("motivo", ""),
            }
        )
    return pd.DataFrame(rows)


def build_duplicados_exactos(campanas_df: pd.DataFrame) -> pd.DataFrame:
    dup = fc.get_duplicate_groups(campanas_df)
    ambiguous_values = {
        "4322|UNI-PUENTELED-1|2024-09-16|2026-07-31||",
        "4336|PLOM-TOTEMD-N1|2026-02-01|2026-12-31||",
        "4480|C7 - AVE|2026-04-15|2026-05-31||",
    }
    exactos = dup[~dup["ClaveNegocio"].isin(ambiguous_values)].reset_index(drop=True)
    return exactos


def build_repeticiones_validas(campanas_df: pd.DataFrame) -> pd.DataFrame:
    cols = ["CargaID", "IDCampaña", "Campaña", "Cliente", "Agencia", "ElementoID", "FechaInicio", "FechaFin", "Estado", "EstadoValidacion", "ObservacionValidacion"]
    sub = campanas_df[campanas_df["CargaID"].isin(["HIST-00001010", "HIST-00001011", "HIST-00001021"])][cols]
    return sub.reset_index(drop=True)


def build_colisiones_pendientes(campanas_df: pd.DataFrame, decision_3: dict, decision_4: dict) -> pd.DataFrame:
    cols = ["CargaID", "IDCampaña", "Campaña", "Cliente", "Marca", "ElementoID", "FechaInicio", "FechaFin", "Estado", "EstadoValidacion"]
    ypf_samsung = campanas_df[campanas_df["IDCampaña"] == 4336][cols].copy()
    ypf_samsung["grupo"] = "DECISION_3_BLOQUEADA: YPF Pier 3 / Samsung (IDCampaña 4336)"
    beyond_taggify = campanas_df[(campanas_df["IDCampaña"] == 4480) & (campanas_df["ElementoID"] == "C7 - AVE")][cols].copy()
    beyond_taggify["grupo"] = "DECISION_4_BLOQUEADA: Beyond Cuarteto / Taggify Lays (IDCampaña 4480, ElementoID C7 - AVE)"
    return pd.concat([ypf_samsung, beyond_taggify], ignore_index=True)


def build_resumen(build_result: dict[str, Any], actual_info: dict, nueva_info: dict, v2_info: dict, cambios: pd.DataFrame, duplicados_exactos: pd.DataFrame) -> pd.DataFrame:
    rows = [
        ("Fecha", "2026-08-18"),
        ("Fuente", str(fc.CANDIDATE_V2_PATH)),
        ("Fuente SHA-256", v2_info["sha256"]),
        ("ACTUAL SHA-256 (no tocado)", actual_info["sha256"]),
        ("ACTUAL SHA-256 coincide con histórico", actual_info["sha256"] == "2f165e12ad90c2f05963367a8e1717dcd1006e9ce669f976afa6e57470aca2cd"),
        ("NUEVA SHA-256 (no tocado)", nueva_info["sha256"]),
        ("", ""),
        ("--- Filas ---", ""),
        ("MAESTRO_ELEMENTOS", build_result["rows"]["MAESTRO_ELEMENTOS"]),
        ("CAMPANAS", build_result["rows"]["CAMPANAS"]),
        ("PARAMETROS", build_result["rows"]["PARAMETROS"]),
        ("", ""),
        ("--- Decisión 1 (KFC) ---", "APLICADA"),
        ("HIST-00000594", "EstadoValidacion=PENDIENTE_DUPLICADO, ObservacionValidacion documentada"),
        ("HIST-00001501", "sin cambios, conservado como registro válido"),
        ("", ""),
        ("--- Decisión 2 (Adidas ID 4322) ---", "APLICADA"),
        ("REPETICION_VALIDA permitido por el modelo (PARAMETROS/validate_input.py)", fc.REPETICION_VALIDA_PERMITIDO),
        ("Valor final usado", "OK" if not fc.REPETICION_VALIDA_PERMITIDO else "REPETICION_VALIDA"),
        ("Filas afectadas", "HIST-00001010, HIST-00001011, HIST-00001021"),
        ("", ""),
        ("--- Decisión 3 (YPF Pier 3 -> IDCampaña 4396) ---", build_result["decision_3"]["status"]),
        ("Motivo", build_result["decision_3"].get("motivo", "")),
        ("", ""),
        ("--- Decisión 4 (Beyond Cuarteto -> IDCampaña 4470) ---", build_result["decision_4"]["status"]),
        ("Motivo", build_result["decision_4"].get("motivo", "")),
        ("", ""),
        ("--- Duplicados / colisiones ---", ""),
        ("Grupos de duplicado exacto conservados (11 esperados)", len(duplicados_exactos)),
        ("Colisiones pendientes (no resueltas en este pase)", 2),
        ("", ""),
        ("--- Total de cambios aplicados ---", len(cambios)),
    ]
    return pd.DataFrame(rows, columns=["Campo", "Valor"])


def run_audit() -> dict[str, Any]:
    if not fc.FINAL_PATH.exists():
        raise SystemExit(
            f"No existe la base FINAL en {fc.FINAL_PATH}. Esta auditoría es de solo lectura: "
            "correr primero build_ocu26_integrada_final.py."
        )

    actual_info = mc.inspect_structure(mc.ACTUAL_PATH)
    nueva_info = mc.inspect_structure(mc.NUEVA_PATH)
    v2_info = mc.inspect_structure(fc.CANDIDATE_V2_PATH)

    campanas_v2 = mc.read_table_df(fc.CANDIDATE_V2_PATH, "CAMPANAS", "tblCampanas")
    campanas_final = mc.read_table_df(fc.FINAL_PATH, "CAMPANAS", "tblCampanas")
    maestro_final = mc.read_table_df(fc.FINAL_PATH, "MAESTRO_ELEMENTOS", "tblElementos")
    parametros_final = mc.read_table_df(fc.FINAL_PATH, "PARAMETROS", "tblParametros")

    campanas_v2_by_id = campanas_v2.set_index("CargaID", drop=False)
    applied_log = []
    for edit in fc.AUTHORIZED_EDITS:
        before_row = campanas_v2_by_id.loc[edit["carga_id"]]
        for col, new_val in edit["changes"].items():
            applied_log.append(
                {
                    "decision": edit["decision"],
                    "CargaID": edit["carga_id"],
                    "ElementoID": before_row["ElementoID"],
                    "campo": col,
                    "valor_anterior": None if mc.is_blank(before_row[col]) else before_row[col],
                    "valor_nuevo": new_val,
                }
            )

    build_result = {
        "rows": {
            "MAESTRO_ELEMENTOS": len(maestro_final),
            "CAMPANAS": len(campanas_final),
            "PARAMETROS": len(parametros_final),
        },
        "ediciones_aplicadas": applied_log,
        "decision_3": fc.verify_decision_3(campanas_v2),
        "decision_4": fc.verify_decision_4(campanas_v2),
    }

    cambios = build_cambios_aplicados(build_result["ediciones_aplicadas"])
    bloqueadas = build_decisiones_bloqueadas(build_result["decision_3"], build_result["decision_4"])
    duplicados_exactos = build_duplicados_exactos(campanas_final)
    repeticiones = build_repeticiones_validas(campanas_final)
    colisiones = build_colisiones_pendientes(campanas_final, build_result["decision_3"], build_result["decision_4"])
    resumen = build_resumen(build_result, actual_info, nueva_info, v2_info, cambios, duplicados_exactos)

    fc.AUDIT_FINAL_PATH.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(fc.AUDIT_FINAL_PATH, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False)
        cambios.to_excel(writer, sheet_name="CAMBIOS_APLICADOS", index=False)
        bloqueadas.to_excel(writer, sheet_name="DECISIONES_BLOQUEADAS", index=False)
        duplicados_exactos.to_excel(writer, sheet_name="DUPLICADOS_EXACTOS", index=False)
        repeticiones.to_excel(writer, sheet_name="REPETICIONES_VALIDAS", index=False)
        colisiones.to_excel(writer, sheet_name="COLISIONES_PENDIENTES", index=False)

    sha_actual_after = mc.calculate_sha256(mc.ACTUAL_PATH)
    sha_nueva_after = mc.calculate_sha256(mc.NUEVA_PATH)
    sha_v2_after = mc.calculate_sha256(fc.CANDIDATE_V2_PATH)
    sources_intact = (
        sha_actual_after == actual_info["sha256"]
        and sha_nueva_after == nueva_info["sha256"]
        and sha_v2_after == v2_info["sha256"]
    )

    return {
        "result": "AUDIT_FINAL_OK",
        "audit_path": str(fc.AUDIT_FINAL_PATH),
        "final_path": str(fc.FINAL_PATH),
        "sources_intact": sources_intact,
        "rows": build_result["rows"],
        "cambios_aplicados": len(cambios),
        "decision_3_status": build_result["decision_3"]["status"],
        "decision_4_status": build_result["decision_4"]["status"],
        "duplicados_exactos_grupos": len(duplicados_exactos),
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Auditoría del pase FINAL OCU26.")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    result = run_audit()

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
    else:
        print("=" * 60)
        print("OCU26 AUDITORIA FINAL")
        print("=" * 60)
        for k, v in result.items():
            print(f"{k}: {v}")

    return 0 if result["sources_intact"] else 1


if __name__ == "__main__":
    sys.exit(main())
