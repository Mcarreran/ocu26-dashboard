"""Motor del pase FINAL_V2 OCU26 (cuarto pase, autorizado).

Fuente: OCU26_BASE_DATOS_INTEGRADA_FINAL_2026-08-18.xlsx (la base FINAL
anterior; no se sobrescribe). Aplica exclusivamente 3 correcciones de
IDCampaña ya confirmadas comercialmente, cada una identificada por su
CargaID exacto (no por búsqueda/fuzzy matching).

Cada corrección se verifica ANTES de escribir:
  1. El CargaID existe y es único.
  2. Campaña/ElementoID/IDCampaña-anterior coinciden con lo declarado
     (tolerando solo mayúsculas/espacios en Campaña/ElementoID).
  3. El IDCampaña propuesto es compatible (no perteneciente a una campaña
     distinta e incompatible).
  4. La ClaveNegocio recalculada no colisiona con ninguna otra fila.

Si cualquier control falla, esa corrección puntual queda BLOQUEADA y no se
escribe nada para ella (el resto de las correcciones, si son válidas,
proceden igual).
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import merge_common as mc  # noqa: E402
import merge_v2_common as v2  # noqa: E402
import merge_final_common as fc  # noqa: E402

vi = mc.vi
is_blank = mc.is_blank
compute_clave_negocio = v2.compute_clave_negocio

FINAL_PATH = fc.FINAL_PATH
FINAL_V2_PATH = mc.UPDATE_ROOT / "output" / "OCU26_BASE_DATOS_INTEGRADA_FINAL_V2_2026-08-18.xlsx"
AUDIT_FINAL_V2_PATH = mc.UPDATE_ROOT / "output" / "OCU26_AUDITORIA_FINAL_V2_2026-08-18.xlsx"

CAMPANAS_KEY = v2.CAMPANAS_KEY

CONFIRMED_IDS = {"AUDI": 4396, "YPF_PIER3_PLOM": 4311, "SAMSUNG": 4336, "BEEYOND_CUARTETO": 4470, "TAGGIFY_LAYS": 4480, "RADIO_DISNEY": 4771}


def _norm(s: Any) -> str:
    if is_blank(s):
        return ""
    return " ".join(str(s).strip().split()).casefold()


def _norm_elemento_id(s: Any) -> str:
    """Normalización para ElementoID: además de mayúsculas, ignora espacios
    alrededor del guion (el usuario escribe 'C7-AVE', la base tiene
    'C7 - AVE'; mismo código de inventario, no un ElementoID distinto)."""
    if is_blank(s):
        return ""
    return "".join(str(s).split()).casefold()


CORRECTIONS: list[dict[str, Any]] = [
    {
        "nombre": "CORRECCION_1_YPF_PIER3",
        "carga_id": "HIST-00000679",
        "expected_campana": "YPF Pier 3",
        "expected_elemento_id": "PLOM-TOTEMD-N1",
        "expected_id_anterior": 4336,
        "id_nuevo": 4311,
        "observacion": "IDCampaña corregido a 4311 por confirmación comercial. Corresponde exclusivamente a YPF Pier 3 en PLOM-TOTEMD-N1.",
    },
    {
        "nombre": "CORRECCION_2_BEEYOND_CUARTETO",
        "carga_id": "HIST-00000449",
        "expected_campana": "BEEYOND CUARTETO",
        "expected_elemento_id": "C7-AVE",
        "expected_id_anterior": 4480,
        "id_nuevo": 4470,
        "observacion": "IDCampaña corregido a 4470 por confirmación comercial. BEEYOND CUARTETO correctamente escrito con doble E.",
    },
    {
        "nombre": "CORRECCION_3_RADIO_DISNEY",
        "carga_id": "HIST-00000152",
        "expected_campana": "Radio Disney",
        "expected_elemento_id": None,  # no especificado por el usuario
        "expected_id_anterior": 4470,
        "id_nuevo": 4771,
        "observacion": "IDCampaña corregido de 4470 a 4771 por confirmación comercial.",
    },
]


def verify_correction(
    correction: dict[str, Any], campanas_df: pd.DataFrame, batch_carga_ids: frozenset[str] = frozenset()
) -> dict[str, Any]:
    result: dict[str, Any] = {
        "nombre": correction["nombre"],
        "carga_id": correction["carga_id"],
        "id_anterior_esperado": correction["expected_id_anterior"],
        "id_nuevo": correction["id_nuevo"],
    }

    sel = campanas_df[campanas_df["CargaID"] == correction["carga_id"]]
    result["filas_encontradas"] = len(sel)
    if len(sel) != 1:
        result["status"] = "BLOQUEADA"
        result["motivo"] = f"CargaID {correction['carga_id']!r} no identifica exactamente una fila (encontradas: {len(sel)})."
        return result

    row = sel.iloc[0]

    if _norm(row["Campaña"]) != _norm(correction["expected_campana"]):
        result["status"] = "BLOQUEADA"
        result["motivo"] = f"Campaña no coincide: esperado~={correction['expected_campana']!r}, encontrado={row['Campaña']!r}."
        return result

    if correction["expected_elemento_id"] is not None and _norm_elemento_id(row["ElementoID"]) != _norm_elemento_id(correction["expected_elemento_id"]):
        result["status"] = "BLOQUEADA"
        result["motivo"] = f"ElementoID no coincide: esperado~={correction['expected_elemento_id']!r}, encontrado={row['ElementoID']!r}."
        return result

    actual_id_anterior = row["IDCampaña"]
    if is_blank(actual_id_anterior) or int(actual_id_anterior) != correction["expected_id_anterior"]:
        result["status"] = "BLOQUEADA"
        result["motivo"] = f"IDCampaña anterior no coincide: esperado={correction['expected_id_anterior']}, encontrado={actual_id_anterior!r}."
        return result

    result["elemento_id"] = row["ElementoID"]
    result["campana_almacenada"] = row["Campaña"]
    result["clave_negocio_anterior"] = row["ClaveNegocio"]

    # --- Compatibilidad del ID propuesto: todas las filas ya existentes con
    # ese IDCampaña deben pertenecer al mismo cliente/campaña que la fila corregida.
    # Se excluyen las demás filas que son objeto de otra corrección del mismo lote
    # (su estado "antes" en este snapshot no representa el resultado final; cada
    # una se valida de forma independiente por su propia entrada en CORRECTIONS). ---
    id_nuevo = correction["id_nuevo"]
    exclude_ids = batch_carga_ids | {correction["carga_id"]}
    conflict = campanas_df[(campanas_df["IDCampaña"] == id_nuevo) & (~campanas_df["CargaID"].isin(exclude_ids))]
    result["filas_existentes_en_id_nuevo"] = len(conflict)
    if not conflict.empty:
        target_brand = _norm(row["Cliente"]) or _norm(row["Marca"])
        incompatible = conflict[
            conflict.apply(lambda r: _norm(r["Cliente"]) != target_brand and _norm(r["Marca"]) != target_brand, axis=1)
        ]
        if not incompatible.empty:
            result["status"] = "BLOQUEADA"
            result["motivo"] = (
                f"IDCampaña {id_nuevo} ya está compartido con {len(incompatible)} fila(s) de cliente/marca "
                f"distinto: {sorted(set(incompatible['Campaña']))} (CargaID {incompatible['CargaID'].tolist()})."
            )
            result["conflicto_con"] = incompatible[["CargaID", "IDCampaña", "Campaña", "ElementoID"]].to_dict(orient="records")
            return result

    # --- Recalcular ClaveNegocio y verificar colisión ---
    new_clave = compute_clave_negocio(
        id_nuevo, row["ElementoID"], row["FechaInicio"], row["FechaFin"], row["FechaIndefinida"], row["HoraInicio"], row["HoraFin"]
    )
    other_batch_old_claves = {
        campanas_df.loc[campanas_df["CargaID"] == cid, "ClaveNegocio"].iloc[0]
        for cid in batch_carga_ids
        if cid != correction["carga_id"] and (campanas_df["CargaID"] == cid).any()
    }
    existing_claves = set(campanas_df["ClaveNegocio"].dropna()) - {row["ClaveNegocio"]} - other_batch_old_claves
    result["clave_negocio_nueva"] = new_clave
    if new_clave in existing_claves:
        result["status"] = "BLOQUEADA"
        result["motivo"] = f"La ClaveNegocio recalculada ({new_clave!r}) colisionaría con otra fila existente."
        return result

    result["status"] = "APLICABLE"
    result["changes"] = {
        "IDCampaña": id_nuevo,
        "ClaveNegocio": new_clave,
        "EstadoValidacion": "OK",
        "ObservacionValidacion": correction["observacion"],
    }
    result["expected_before"] = {
        "IDCampaña": correction["expected_id_anterior"],
        "ClaveNegocio": row["ClaveNegocio"],
    }
    return result


def verify_all_corrections(campanas_df: pd.DataFrame) -> list[dict[str, Any]]:
    batch_ids = frozenset(c["carga_id"] for c in CORRECTIONS)
    return [verify_correction(c, campanas_df, batch_ids) for c in CORRECTIONS]
