"""Motor del pase FINAL OCU26 (tercer pase, autorizado).

A diferencia de V1/V2 (merge_common.py / merge_v2_common.py, que son
motores de cruce genéricos entre dos archivos), este módulo NO cruza
archivos: aplica una lista fija y explícita de ediciones comerciales ya
confirmadas por el usuario sobre la candidata V2 (no sobre ACTUAL/NUEVA).

Cada edición autorizada declara:
  - `carga_id`: la fila exacta a tocar.
  - `expected_before`: el valor que DEBE tener esa celda en la candidata V2
    antes de escribir (si no coincide, se aborta el build entero — nunca se
    escribe "a ciegas").
  - `changes`: el valor final exacto, dado literalmente por el usuario.

Decisiones 3 y 4 (reasignación de IDCampaña) NO están en AUTORIZED_EDITS:
se verifican con `verify_decision_3` / `verify_decision_4` y, si la
verificación falla, quedan bloqueadas y documentadas — nunca se aplican
"a medias" ni se adivina un valor no autorizado.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import merge_common as mc  # noqa: E402
import merge_v2_common as v2  # noqa: E402

vi = mc.vi
is_blank = mc.is_blank

CANDIDATE_V2_PATH = v2.CANDIDATE_V2_PATH
FINAL_PATH = mc.UPDATE_ROOT / "output" / "OCU26_BASE_DATOS_INTEGRADA_FINAL_2026-08-18.xlsx"
AUDIT_FINAL_PATH = mc.UPDATE_ROOT / "output" / "OCU26_AUDITORIA_FINAL_2026-08-18.xlsx"

CAMPANAS_KEY = v2.CAMPANAS_KEY

KFC_OBSERVACION = (
    "Duplicado confirmado de HIST-00001501. Mismo IDCampaña 4525, ElementoID FPBROWN-FRONT-3, "
    "cliente KFC y FechaFin. Se conserva únicamente por trazabilidad histórica y no debe "
    "contabilizarse como activación válida."
)
ADIDAS_OBSERVACION = (
    "REPETICION_VALIDA: misma campaña ID 4322 con cambios sucesivos de creatividad sobre pantalla "
    "digital multislot. Las tres filas deben conservarse, pero los tableros deben contar la campaña "
    "por IDCampaña y no como tres campañas comerciales diferentes."
)

# EstadoValidacion permitido: intersección de lo que valida_input.py exige (Gate 1, más estricto)
# y lo documentado en PARAMETROS. REPETICION_VALIDA no pertenece a ninguno de los dos ->
# Decisión 2 usa la rama de contingencia explícitamente prevista por el usuario (conservar OK).
ESTADO_VALIDACION_PERMITIDOS = vi.ESTADO_VALIDACION_VALUES
REPETICION_VALIDA_PERMITIDO = "REPETICION_VALIDA" in ESTADO_VALIDACION_PERMITIDOS

AUTHORIZED_EDITS: list[dict[str, Any]] = [
    {
        "carga_id": "HIST-00000594",
        "decision": "DECISION_1_KFC",
        "expected_before": {"EstadoValidacion": "OK", "ObservacionValidacion": None},
        "changes": {"EstadoValidacion": "PENDIENTE_DUPLICADO", "ObservacionValidacion": KFC_OBSERVACION},
    },
    {
        "carga_id": "HIST-00001010",
        "decision": "DECISION_2_ADIDAS",
        "expected_before": {"EstadoValidacion": "PENDIENTE_COLISION", "ObservacionValidacion": None},
        "changes": {
            "EstadoValidacion": "REPETICION_VALIDA" if REPETICION_VALIDA_PERMITIDO else "OK",
            "ObservacionValidacion": ADIDAS_OBSERVACION,
        },
    },
    {
        "carga_id": "HIST-00001011",
        "decision": "DECISION_2_ADIDAS",
        "expected_before": {"EstadoValidacion": "PENDIENTE_COLISION", "ObservacionValidacion": None},
        "changes": {
            "EstadoValidacion": "REPETICION_VALIDA" if REPETICION_VALIDA_PERMITIDO else "OK",
            "ObservacionValidacion": ADIDAS_OBSERVACION,
        },
    },
    {
        "carga_id": "HIST-00001021",
        "decision": "DECISION_2_ADIDAS",
        "expected_before": {"EstadoValidacion": "PENDIENTE_COLISION", "ObservacionValidacion": None},
        "changes": {
            "EstadoValidacion": "REPETICION_VALIDA" if REPETICION_VALIDA_PERMITIDO else "OK",
            "ObservacionValidacion": ADIDAS_OBSERVACION,
        },
    },
]


def _norm(s: Any) -> str:
    if is_blank(s):
        return ""
    return " ".join(str(s).strip().split()).casefold()


def verify_decision_3(campanas_df: pd.DataFrame) -> dict[str, Any]:
    """DECISIÓN 3 — YPF Pier 3 (IDCampaña 4336 -> 4396). Solo verifica, no escribe."""
    result: dict[str, Any] = {"decision": "DECISION_3_YPF_PIER3", "id_actual": 4336, "id_propuesto": 4396}

    sel = campanas_df[
        (campanas_df["IDCampaña"] == 4336)
        & (campanas_df["ElementoID"] == "PLOM-TOTEMD-N1")
        & (campanas_df["Campaña"].apply(_norm) == _norm("YPF PIER 3"))
    ]
    result["filas_encontradas"] = len(sel)
    result["carga_ids_encontrados"] = sel["CargaID"].tolist()

    if len(sel) != 1:
        result["status"] = "BLOQUEADA"
        result["motivo"] = f"Se esperaba exactamente 1 fila con IDCampaña=4336, ElementoID=PLOM-TOTEMD-N1, Campaña~='YPF PIER 3'; se encontraron {len(sel)}."
        return result

    target_row = sel.iloc[0]
    conflict = campanas_df[campanas_df["IDCampaña"] == 4396]
    conflicting_names = sorted(set(conflict["Campaña"].dropna()))
    result["filas_existentes_en_id_propuesto"] = len(conflict)
    result["campanas_existentes_en_id_propuesto"] = conflicting_names

    if not conflict.empty and conflicting_names != []:
        # 4396 ya está en uso: solo sería "compatible" si TODAS las filas que ya lo usan
        # fueran también YPF Pier 3 (mismo cliente/campaña).
        if conflicting_names != ["YPF PIER 3"]:
            result["status"] = "BLOQUEADA"
            result["motivo"] = (
                f"IDCampaña 4396 ya pertenece a otra campaña establecida ({conflicting_names}, "
                f"{len(conflict)} fila(s): {conflict['CargaID'].tolist()}). No es compatible con YPF Pier 3; "
                "aplicar la reasignación fusionaría silenciosamente dos campañas distintas."
            )
            return result

    new_clave = v2.compute_clave_negocio(
        4396, target_row["ElementoID"], target_row["FechaInicio"], target_row["FechaFin"],
        target_row["FechaIndefinida"], target_row["HoraInicio"], target_row["HoraFin"],
    )
    existing_claves = set(campanas_df["ClaveNegocio"].dropna()) - {target_row["ClaveNegocio"]}
    result["nueva_clave_negocio"] = new_clave
    if new_clave in existing_claves:
        result["status"] = "BLOQUEADA"
        result["motivo"] = f"La ClaveNegocio recalculada ({new_clave!r}) colisionaría con otra fila existente."
        return result

    result["status"] = "APLICABLE"
    result["carga_id"] = target_row["CargaID"]
    result["samsung_carga_id"] = campanas_df.loc[
        (campanas_df["IDCampaña"] == 4336) & (campanas_df["CargaID"] != target_row["CargaID"]), "CargaID"
    ].tolist()
    return result


def verify_decision_4(campanas_df: pd.DataFrame) -> dict[str, Any]:
    """DECISIÓN 4 — Beyond Cuarteto (IDCampaña 4480 -> 4470). Solo verifica, no escribe."""
    result: dict[str, Any] = {"decision": "DECISION_4_BEYOND_CUARTETO", "id_actual": 4480, "id_propuesto": 4470}

    sel = campanas_df[
        (campanas_df["IDCampaña"] == 4480)
        & (campanas_df["ElementoID"].apply(_norm) == _norm("C7 - AVE"))
        & (campanas_df["Campaña"].apply(_norm) == _norm("BEYOND CUARTETO"))
    ]
    result["filas_encontradas"] = len(sel)
    result["carga_ids_encontrados"] = sel["CargaID"].tolist()

    all_4480 = campanas_df[campanas_df["IDCampaña"] == 4480][["CargaID", "Campaña", "ElementoID"]]
    result["filas_bajo_id_4480"] = all_4480.to_dict(orient="records")

    if len(sel) != 1:
        candidatos = campanas_df[
            (campanas_df["IDCampaña"] == 4480)
            & (campanas_df["ElementoID"].apply(_norm) == _norm("C7 - AVE"))
            & (campanas_df["Campaña"].apply(lambda x: not is_blank(x) and "cuartet" in _norm(x)))
        ]
        result["status"] = "BLOQUEADA"
        result["motivo"] = (
            "Se esperaba exactamente 1 fila con IDCampaña=4480, ElementoID~='C7 - AVE', "
            f"Campaña~='BEYOND CUARTETO' (tolerando mayúsculas/espacios); se encontraron {len(sel)}. "
            f"El valor real almacenado para el candidato más cercano es "
            f"{candidatos['Campaña'].tolist() if not candidatos.empty else 'ninguno'!r}, que difiere de "
            "'BEYOND CUARTETO' en algo más que mayúsculas/espacios (posible error de tipeo, pero no se "
            "aplica sin confirmación explícita: no se usa fuzzy matching para actualizar identificadores)."
        )
        result["candidato_mas_cercano"] = candidatos["CargaID"].tolist() if not candidatos.empty else []
        return result

    target_row = sel.iloc[0]
    conflict = campanas_df[campanas_df["IDCampaña"] == 4470]
    conflicting_names = sorted(set(conflict["Campaña"].dropna()))
    result["filas_existentes_en_id_propuesto"] = len(conflict)
    result["campanas_existentes_en_id_propuesto"] = conflicting_names

    incompatibles = conflict[~conflict["Campaña"].apply(lambda x: "cuartet" in _norm(x))]
    if not incompatibles.empty:
        result["status"] = "BLOQUEADA"
        result["motivo"] = (
            f"IDCampaña 4470 ya está compartido con {len(incompatibles)} fila(s) de otro cliente no "
            f"relacionado ({sorted(set(incompatibles['Campaña']))}, CargaID {incompatibles['CargaID'].tolist()}). "
            "No es un ID limpio y exclusivo para Beyond Cuarteto; reasignar fusionaría/oscurecería esa "
            "colisión preexistente."
        )
        result["conflicto_con"] = incompatibles[["CargaID", "IDCampaña", "Campaña", "ElementoID"]].to_dict(orient="records")
        return result

    new_clave = v2.compute_clave_negocio(
        4470, target_row["ElementoID"], target_row["FechaInicio"], target_row["FechaFin"],
        target_row["FechaIndefinida"], target_row["HoraInicio"], target_row["HoraFin"],
    )
    existing_claves = set(campanas_df["ClaveNegocio"].dropna()) - {target_row["ClaveNegocio"]}
    result["nueva_clave_negocio"] = new_clave
    if new_clave in existing_claves:
        result["status"] = "BLOQUEADA"
        result["motivo"] = f"La ClaveNegocio recalculada ({new_clave!r}) colisionaría con otra fila existente."
        return result

    result["status"] = "APLICABLE"
    result["carga_id"] = target_row["CargaID"]
    return result


def get_duplicate_groups(campanas_df: pd.DataFrame) -> pd.DataFrame:
    cn = campanas_df["ClaveNegocio"].dropna()
    counts = cn.value_counts()
    dup_values = sorted(counts[counts > 1].index)
    rows = []
    for val in dup_values:
        sub = campanas_df[campanas_df["ClaveNegocio"] == val]
        rows.append(
            {
                "ClaveNegocio": val,
                "ocurrencias": len(sub),
                "CargaID_involucrados": ", ".join(sub["CargaID"].tolist()),
                "EstadoValidacion_involucrados": ", ".join(sub["EstadoValidacion"].astype(str).tolist()),
            }
        )
    return pd.DataFrame(rows)
