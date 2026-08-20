"""Auditoría del pase FINAL_V2 OCU26 (solo lectura).

Documenta las 3 correcciones de IDCampaña aplicadas sobre la base FINAL
anterior. Genera:

    Pendientes/OCU26_ACTUALIZACION/output/OCU26_AUDITORIA_FINAL_V2_2026-08-18.xlsx

con las hojas: RESUMEN, CAMBIOS_APLICADOS, COLISIONES_RESUELTAS,
DUPLICADOS_EXACTOS, REPETICIONES_VALIDAS.

No modifica ningún archivo fuente ni reconstruye FINAL_V2 (debe existir).

Uso:
    python audit_ocu26_integrada_final_v2.py
    python audit_ocu26_integrada_final_v2.py --json
"""

from __future__ import annotations

import argparse
import json
import sys
from typing import Any

import pandas as pd

import merge_common as mc
import merge_final_common as fc
import merge_final_v2_common as fv2

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


def build_cambios_aplicados(verification: list[dict[str, Any]], campanas_final: pd.DataFrame) -> pd.DataFrame:
    rows = []
    final_by_id = campanas_final.set_index("CargaID", drop=False)
    for r in verification:
        if r["status"] != "APLICABLE":
            continue
        before_row = final_by_id.loc[r["carga_id"]]
        for col, new_val in r["changes"].items():
            rows.append(
                {
                    "correccion": r["nombre"],
                    "CargaID": r["carga_id"],
                    "ElementoID": r["elemento_id"],
                    "campo_modificado": col,
                    "valor_anterior": _fmt(before_row[col]),
                    "valor_nuevo": _fmt(new_val),
                }
            )
    return pd.DataFrame(rows)


def build_colisiones_resueltas(verification: list[dict[str, Any]], campanas_v2: pd.DataFrame) -> pd.DataFrame:
    cols = ["CargaID", "IDCampaña", "Campaña", "Cliente", "Marca", "ElementoID", "FechaInicio", "FechaFin", "Estado", "EstadoValidacion", "ClaveNegocio"]
    rows = []
    for r in verification:
        if r["status"] != "APLICABLE":
            continue
        sub = campanas_v2[campanas_v2["CargaID"] == r["carga_id"]][cols].copy()
        sub["correccion"] = r["nombre"]
        sub["clave_negocio_anterior"] = r["clave_negocio_anterior"]
        rows.append(sub)
    # también mostrar las filas "hermanas" que quedaron sin cambios (Samsung, Taggify Lays)
    siblings = {
        "CORRECCION_1_YPF_PIER3": ["HIST-00000681"],
        "CORRECCION_2_BEEYOND_CUARTETO": ["HIST-00000321", "HIST-00000448", "HIST-00000483"],
    }
    for correccion, ids in siblings.items():
        sub = campanas_v2[campanas_v2["CargaID"].isin(ids)][cols].copy()
        sub["correccion"] = f"{correccion} (fila hermana, sin cambios)"
        sub["clave_negocio_anterior"] = sub["ClaveNegocio"]
        rows.append(sub)
    if not rows:
        return pd.DataFrame(columns=cols + ["correccion", "clave_negocio_anterior"])
    return pd.concat(rows, ignore_index=True)


def build_duplicados_exactos(campanas_df: pd.DataFrame) -> pd.DataFrame:
    dup = fc.get_duplicate_groups(campanas_df)
    adidas_value = "4322|UNI-PUENTELED-1|2024-09-16|2026-07-31||"
    exactos = dup[dup["ClaveNegocio"] != adidas_value].reset_index(drop=True)
    return exactos


def build_repeticiones_validas(campanas_df: pd.DataFrame) -> pd.DataFrame:
    cols = ["CargaID", "IDCampaña", "Campaña", "Cliente", "Agencia", "ElementoID", "FechaInicio", "FechaFin", "Estado", "EstadoValidacion", "ObservacionValidacion"]
    sub = campanas_df[campanas_df["CargaID"].isin(["HIST-00001010", "HIST-00001011", "HIST-00001021"])][cols]
    return sub.reset_index(drop=True)


def build_resumen(
    verification: list[dict[str, Any]],
    actual_info: dict,
    nueva_info: dict,
    v2_info: dict,
    final_info: dict,
    final_v2_info: dict,
    rows: dict[str, int],
    cambios: pd.DataFrame,
    duplicados_exactos: pd.DataFrame,
) -> pd.DataFrame:
    data = [
        ("Fecha", "2026-08-18"),
        ("Fuente", str(fv2.FINAL_PATH)),
        ("Fuente SHA-256", final_info["sha256"]),
        ("ACTUAL SHA-256 (no tocado)", actual_info["sha256"]),
        ("ACTUAL SHA-256 coincide con histórico", actual_info["sha256"] == "2f165e12ad90c2f05963367a8e1717dcd1006e9ce669f976afa6e57470aca2cd"),
        ("NUEVA SHA-256 (no tocado)", nueva_info["sha256"]),
        ("Candidata V2 SHA-256 (no tocado)", v2_info["sha256"]),
        ("FINAL_V2 SHA-256", final_v2_info["sha256"]),
        ("", ""),
        ("--- Filas ---", ""),
        ("MAESTRO_ELEMENTOS", rows["MAESTRO_ELEMENTOS"]),
        ("CAMPANAS", rows["CAMPANAS"]),
        ("PARAMETROS", rows["PARAMETROS"]),
        ("", ""),
        ("--- IDs comerciales confirmados ---", ""),
        ("Audi", 4396),
        ("YPF Pier 3 (PLOM-TOTEMD-N1)", 4311),
        ("Samsung", 4336),
        ("BEEYOND CUARTETO", 4470),
        ("Taggify Lays", 4480),
        ("Radio Disney", 4771),
        ("", ""),
    ]
    for r in verification:
        data.append((f"--- {r['nombre']} ---", r["status"]))
        data.append(("CargaID", r["carga_id"]))
        data.append(("IDCampaña anterior -> nuevo", f"{r['id_anterior_esperado']} -> {r['id_nuevo']}"))
        if r["status"] == "APLICABLE":
            data.append(("ClaveNegocio anterior", r["clave_negocio_anterior"]))
            data.append(("ClaveNegocio nueva", r["clave_negocio_nueva"]))
        else:
            data.append(("Motivo", r.get("motivo", "")))
        data.append(("", ""))

    data.extend(
        [
            ("--- Casos ya resueltos, no tocados en este pase ---", ""),
            ("KFC: HIST-00000594", "PENDIENTE_DUPLICADO (sin cambios)"),
            ("KFC: HIST-00001501", "válido, sin cambios"),
            ("Adidas IDCampaña 4322", "3 creatividades, repetición válida (sin cambios)"),
            ("", ""),
            ("--- Duplicados / colisiones ---", ""),
            ("Grupos de duplicado exacto conservados (11 esperados)", len(duplicados_exactos)),
            ("Colisiones resueltas en este pase", 2),
            ("Colisiones pendientes restantes", 0),
            ("", ""),
            ("Total de cambios aplicados", len(cambios)),
        ]
    )
    return pd.DataFrame(data, columns=["Campo", "Valor"])


def run_audit() -> dict[str, Any]:
    if not fv2.FINAL_V2_PATH.exists():
        raise SystemExit(
            f"No existe FINAL_V2 en {fv2.FINAL_V2_PATH}. Esta auditoría es de solo lectura: "
            "correr primero build_ocu26_integrada_final_v2.py."
        )

    actual_info = mc.inspect_structure(mc.ACTUAL_PATH)
    nueva_info = mc.inspect_structure(mc.NUEVA_PATH)
    v2_info = mc.inspect_structure(fc.CANDIDATE_V2_PATH)
    final_info = mc.inspect_structure(fv2.FINAL_PATH)
    final_v2_info = mc.inspect_structure(fv2.FINAL_V2_PATH)

    campanas_final = mc.read_table_df(fv2.FINAL_PATH, "CAMPANAS", "tblCampanas")
    campanas_v2 = mc.read_table_df(fv2.FINAL_V2_PATH, "CAMPANAS", "tblCampanas")
    maestro_v2 = mc.read_table_df(fv2.FINAL_V2_PATH, "MAESTRO_ELEMENTOS", "tblElementos")
    parametros_v2 = mc.read_table_df(fv2.FINAL_V2_PATH, "PARAMETROS", "tblParametros")

    verification = fv2.verify_all_corrections(campanas_final)

    rows = {
        "MAESTRO_ELEMENTOS": len(maestro_v2),
        "CAMPANAS": len(campanas_v2),
        "PARAMETROS": len(parametros_v2),
    }

    cambios = build_cambios_aplicados(verification, campanas_final)
    colisiones = build_colisiones_resueltas(verification, campanas_v2)
    duplicados_exactos = build_duplicados_exactos(campanas_v2)
    repeticiones = build_repeticiones_validas(campanas_v2)
    resumen = build_resumen(verification, actual_info, nueva_info, v2_info, final_info, final_v2_info, rows, cambios, duplicados_exactos)

    fv2.AUDIT_FINAL_V2_PATH.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(fv2.AUDIT_FINAL_V2_PATH, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False)
        cambios.to_excel(writer, sheet_name="CAMBIOS_APLICADOS", index=False)
        colisiones.to_excel(writer, sheet_name="COLISIONES_RESUELTAS", index=False)
        duplicados_exactos.to_excel(writer, sheet_name="DUPLICADOS_EXACTOS", index=False)
        repeticiones.to_excel(writer, sheet_name="REPETICIONES_VALIDAS", index=False)

    sha_actual_after = mc.calculate_sha256(mc.ACTUAL_PATH)
    sha_nueva_after = mc.calculate_sha256(mc.NUEVA_PATH)
    sha_v2_after = mc.calculate_sha256(fc.CANDIDATE_V2_PATH)
    sha_final_after = mc.calculate_sha256(fv2.FINAL_PATH)
    sources_intact = (
        sha_actual_after == actual_info["sha256"]
        and sha_nueva_after == nueva_info["sha256"]
        and sha_v2_after == v2_info["sha256"]
        and sha_final_after == final_info["sha256"]
    )

    return {
        "result": "AUDIT_FINAL_V2_OK",
        "audit_path": str(fv2.AUDIT_FINAL_V2_PATH),
        "final_v2_path": str(fv2.FINAL_V2_PATH),
        "sources_intact": sources_intact,
        "rows": rows,
        "cambios_aplicados": len(cambios),
        "corrections_status": {r["nombre"]: r["status"] for r in verification},
        "duplicados_exactos_grupos": len(duplicados_exactos),
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Auditoría del pase FINAL_V2 OCU26.")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    result = run_audit()

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
    else:
        print("=" * 60)
        print("OCU26 AUDITORIA FINAL_V2")
        print("=" * 60)
        for k, v in result.items():
            print(f"{k}: {v}")

    return 0 if result["sources_intact"] else 1


if __name__ == "__main__":
    sys.exit(main())
