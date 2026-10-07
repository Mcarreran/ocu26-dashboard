"""Tests del cierre de migracion historica (Etapa 2B.2).

Hermeticos: DataFrames sinteticos en memoria; no leen ni escriben la base.
"""

from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT / "scripts"))

import cierre_historico_2b2 as ci  # noqa: E402
import staging_import_historico as st  # noqa: E402
from validate_input import CAMPANAS_HEADERS  # noqa: E402

D = dt.date
CARGA_2B1 = dt.datetime(2026, 10, 7, 13, 29, 4)


def _camp(filas):
    """filas: (CargaID, OT, ElementoID, ini, fin, Estado, UsuarioCarga, FuenteCarga)."""
    out = []
    for cid, ot, eid, ini, fin, estado, usuario, fuente in filas:
        r = {h: None for h in CAMPANAS_HEADERS}
        r.update({"CargaID": cid, "IDCampaña": float(ot), "ElementoID": eid, "FechaInicio": pd.Timestamp(ini),
                  "FechaFin": pd.Timestamp(fin), "Estado": estado, "UsuarioCarga": usuario, "FuenteCarga": fuente,
                  "TipoCargaDeclarado": "Digital", "ClaveNegocio": st.clave_negocio(ot, eid, ini, fin)})
        out.append(r)
    return pd.DataFrame(out, columns=CAMPANAS_HEADERS)


def test_ts_de_origen():
    assert ci.ts_de_origen("REM-TS 3-V1 + REM-TS 3 -V2") == 3
    assert ci.ts_de_origen("REM-TS -5-V3") == 5
    assert ci.ts_de_origen("REM-TS 1-V1 + REM-TS 2-V1") is None
    assert ci.ts_de_origen("C1 - CAB") is None


def test_estado_extensiones_regla_y_reservada():
    camp = _camp([
        ("HIST-1", 1, "E", D(2026, 3, 1), D(2026, 12, 31), "Finalizada", "x", "Migración histórica"),
        ("HIST-2", 2, "E", D(2026, 3, 1), D(2026, 8, 31), "Activa", "x", "Migración histórica"),
        ("HIST-3", 3, "E", D(2026, 3, 1), D(2026, 12, 31), "Reservada", "x", "Migración histórica"),
        ("HIST-4", 4, "E", D(2026, 3, 1), D(2026, 10, 7), "Activa", "x", "Migración histórica"),
    ])
    cambios, control = ci.cambios_estado_extensiones(camp, ["HIST-1", "HIST-2", "HIST-3", "HIST-4"], CARGA_2B1)
    assert cambios[["CargaID", "Anterior", "Nuevo"]].values.tolist() == [
        ["HIST-1", "Finalizada", "Activa"], ["HIST-2", "Activa", "Finalizada"]]
    assert set(cambios["Columna"]) == {"Estado"}
    c = control.set_index("CargaID")
    assert c.loc["HIST-3", "EstadoNuevo"] == "Reservada" and "Reservada" in c.loc["HIST-3", "Accion"]
    assert c.loc["HIST-4", "EstadoNuevo"] == "Activa"  # FechaFin = fecha de carga -> Activa


def test_fuente_carga_canonica_solo_filas_propias():
    camp = _camp([
        ("HIST-1", 1, "E", D(2026, 3, 1), D(2026, 3, 31), "Activa", st.USUARIO_CARGA, "Migración histórica - OCUPACION_2026"),
        ("HIST-2", 2, "E", D(2026, 3, 1), D(2026, 3, 31), "Activa", "MIGRACION_YPF", "Otra"),
        ("HIST-3", 3, "E", D(2026, 3, 1), D(2026, 3, 31), "Activa", st.USUARIO_CARGA, "Migración histórica"),
    ])
    ch = ci.cambios_fuente_carga(camp)
    assert ch[["CargaID", "Columna", "Nuevo"]].values.tolist() == [["HIST-1", "FuenteCarga", "Migración histórica"]]


def test_conciliar_remeros():
    camp = _camp([
        ("HIST-10", 4600, "REM-DB-1", D(2026, 6, 1), D(2026, 6, 30), "Finalizada", st.USUARIO_CARGA, "x"),  # TS1 ok
        ("HIST-11", 4601, "REM-DB-3", D(2026, 7, 1), D(2026, 7, 31), "Finalizada", st.USUARIO_CARGA, "x"),  # TS3 numerico
    ])
    actuales = pd.DataFrame([
        {"CargaID": "HIST-10", "IDCampaña": 4600, "ElementoID": "REM-DB-1", "FechaInicio": D(2026, 6, 1),
         "FechaFin": D(2026, 6, 30), "TRZ_ElementoOrigen": "REM-TS 1-V1"},
        {"CargaID": "HIST-11", "IDCampaña": 4601, "ElementoID": "REM-DB-3", "FechaInicio": D(2026, 7, 1),
         "FechaFin": D(2026, 7, 31), "TRZ_ElementoOrigen": "REM-TS 3-V2"},
    ])
    deseadas = pd.DataFrame([
        {"IDCampaña": 4600, "ElementoID": "REM-DB-1", "FechaInicio": D(2026, 6, 1), "FechaFin": D(2026, 6, 30), "TRZ_ElementoOrigen": "REM-TS 1-V1"},
        {"IDCampaña": 4601, "ElementoID": "REM-DB-5", "FechaInicio": D(2026, 7, 1), "FechaFin": D(2026, 7, 31), "TRZ_ElementoOrigen": "REM-TS 3-V2"},
        {"IDCampaña": 4700, "ElementoID": "REM-DB-3", "FechaInicio": D(2026, 8, 1), "FechaFin": D(2026, 8, 31), "TRZ_ElementoOrigen": "REM-TS 2-V1"},
    ])
    medio = {f"REM-DB-{n}": "Digital" for n in (1, 3, 5, 6, 8, 10)}
    cambios, nuevas, errores = ci.conciliar_remeros(deseadas, actuales, camp, medio)
    assert errores == []
    assert cambios[["CargaID", "Columna", "Anterior", "Nuevo"]].values.tolist() == [
        ["HIST-11", "ElementoID", "REM-DB-3", "REM-DB-5"],
        ["HIST-11", "ClaveNegocio", "4601|REM-DB-3|2026-07-01|2026-07-31||", "4601|REM-DB-5|2026-07-01|2026-07-31||"]]
    assert nuevas["IDCampaña"].tolist() == [4700]


def test_conciliar_remeros_fila_sobrante_aborta():
    camp = _camp([("HIST-11", 4601, "REM-DB-3", D(2026, 7, 1), D(2026, 7, 31), "Finalizada", st.USUARIO_CARGA, "x")])
    actuales = pd.DataFrame([{"CargaID": "HIST-11", "IDCampaña": 4601, "ElementoID": "REM-DB-3", "FechaInicio": D(2026, 7, 1),
                              "FechaFin": D(2026, 7, 31), "TRZ_ElementoOrigen": "REM-TS 3-V2"}])
    deseadas = pd.DataFrame(columns=["IDCampaña", "ElementoID", "FechaInicio", "FechaFin", "TRZ_ElementoOrigen"])
    _, _, errores = ci.conciliar_remeros(deseadas, actuales, camp, {})
    assert errores and "requeririan borrado" in errores[0]
