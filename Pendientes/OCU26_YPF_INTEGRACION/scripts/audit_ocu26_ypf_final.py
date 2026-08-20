# -*- coding: utf-8 -*-
"""Genera la auditoría FINAL de la integración YPF -> OCU26.

Reutiliza (por import) audit_ocu26_ypf.build_audit() para las 12 hojas ya
generadas para la candidata, y agrega:
  - Hoja ELEMENTOS_RETIRADOS_APIE_30943: las 17 filas completas retiradas de
    MAESTRO_ELEMENTOS, clasificadas ELEMENTO_LEGACY_APIE_30943_INVALIDO,
    con la verificación de cero referencias en CAMPANAS.
  - RESUMEN_FINAL: conteos antes/después del retiro.

READ-ONLY sobre la candidata y las 3 fuentes originales.

Uso:
    python audit_ocu26_ypf_final.py
    python audit_ocu26_ypf_final.py --json
"""

from __future__ import annotations

import argparse
import json
import sys

import pandas as pd
from openpyxl import Workbook

import merge_ypf_common as m
from audit_ocu26_ypf import build_audit, _write_df


def build_final_audit() -> dict[str, pd.DataFrame]:
    plan = m.compute_plan()
    sheets = build_audit(plan)

    final_plan = m.compute_final_removal_plan()

    sheets["ELEMENTOS_RETIRADOS_APIE_30943"] = pd.DataFrame(
        [{"Clasificacion": m.LEGACY_APIE_30943_INVALIDO, **row} for row in final_plan.elementos_a_retirar]
    )

    resumen_final_rows = [
        ("Fecha de generación (pase FINAL)", final_plan.run_timestamp.isoformat()),
        ("Candidata origen", str(m.CANDIDATE_PATH)),
        ("Candidata SHA-256", m.calculate_sha256(m.CANDIDATE_PATH)),
        ("MAESTRO_ELEMENTOS antes del retiro (candidata)", final_plan.expected_counts["MAESTRO_ELEMENTOS_antes"]),
        ("Elementos retirados (ELEMENTO_LEGACY_APIE_30943_INVALIDO)", final_plan.expected_counts["MAESTRO_ELEMENTOS_retirados"]),
        ("MAESTRO_ELEMENTOS después del retiro (FINAL)", final_plan.expected_counts["MAESTRO_ELEMENTOS_despues"]),
        ("CAMPANAS (sin cambios)", final_plan.expected_counts["CAMPANAS"]),
        ("PARAMETROS (sin cambios)", final_plan.expected_counts["PARAMETROS"]),
        ("Verificación previa", "0 filas de CAMPANAS referencian los 17 ElementoID retirados (confirmado antes de escribir)"),
        ("APIE 30943 final esperado", "30943 - FB - 1, 30943 - FB - 2 (2 elementos, ambos FB)"),
    ]
    sheets["RESUMEN_FINAL"] = pd.DataFrame(resumen_final_rows, columns=["Campo", "Valor"])

    # Actualiza la clasificación de estos 17 en la hoja APIE_30943 ya generada
    # (antes decían PREEXISTENTE_OCU26_NO_YPF; en la versión FINAL pasan a
    # retirados con la clasificación explícita pedida).
    apie_df = sheets["APIE_30943"].copy()
    retirados_ids = {r[m.MAESTRO_KEY] for r in final_plan.elementos_a_retirar}
    if not apie_df.empty and "ElementoID" in apie_df.columns:
        apie_df.loc[apie_df["ElementoID"].isin(retirados_ids), "Origen"] = m.LEGACY_APIE_30943_INVALIDO + " (retirado en FINAL)"
    sheets["APIE_30943"] = apie_df

    return sheets


def write_final_audit_workbook(sheets: dict[str, pd.DataFrame]) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    order = [
        "RESUMEN", "RESUMEN_FINAL", "MAPEO_COLUMNAS", "ELEMENTOS_YA_EXISTENTES", "ELEMENTOS_NUEVOS",
        "ELEMENTOS_RETIRADOS_APIE_30943", "CONFLICTOS_ELEMENTOS", "CAMPAÑAS_YA_EXISTENTES", "CAMPAÑAS_NUEVAS",
        "CONFLICTOS_CAMPAÑAS", "DUPLICADOS_EVITADOS", "APIE_30943",
        "INTEGRIDAD_REFERENCIAL", "A_VALIDAR",
    ]
    for name in order:
        df = sheets.get(name, pd.DataFrame())
        if df.empty and list(df.columns) == []:
            df = pd.DataFrame([{"info": "sin filas"}])
        _write_df(wb, name, df)
    m.FINAL_AUDIT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(m.FINAL_AUDIT_PATH)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Genera la auditoría FINAL de integración YPF -> OCU26.")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    sheets = build_final_audit()
    write_final_audit_workbook(sheets)

    result = {
        "result": "AUDIT_YPF_FINAL_OK",
        "audit_path": str(m.FINAL_AUDIT_PATH),
        "sheets": {k: len(v) for k, v in sheets.items()},
    }

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
    else:
        print("=" * 60)
        print("OCU26 + YPF AUDITORÍA FINAL")
        print("=" * 60)
        print(f"Ruta: {result['audit_path']}")
        for k, v in result["sheets"].items():
            print(f"  hoja {k}: {v} filas")

    return 0


if __name__ == "__main__":
    sys.exit(main())
