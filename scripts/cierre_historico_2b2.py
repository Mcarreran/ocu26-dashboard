"""Cierre de la migracion historica OCU26 - Etapa 2B.2.

Parte del estado aplicado en la Etapa 2B.1 (base canonica con el SHA
esperado) y aplica SOLO los ajustes deterministicos confirmados:

1. Estado de las filas extendidas en 2B.1: FechaFin >= fecha de carga 2B.1 ->
   Activa, FechaFin < fecha de carga -> Finalizada. Reservada no se toca.
2. Remeros, regla posicional: REM-TS n = n-esimo soporte Remeros Digital del
   maestro. Las filas de 2B.1 cargadas con la regla numerica anterior se
   reasignan (ElementoID + ClaveNegocio, mismo CargaID); las asignaciones
   nuevas que resultan inequivocas se insertan.
3. FuenteCarga de las filas de esta migracion -> convencion canonica del
   vocabulario PARAMETROS ("Migración histórica"); la procedencia queda en
   UsuarioCarga = MIGRACION_OCUPACION_2026.

El staging se recalcula (solo lectura) contra el corte PRE-CARGA con las
reglas vigentes: para todo lo que no es Remeros debe reproducir exactamente
lo aplicado en 2B.1 (control de determinismo); si no, aborta.

Uso:
    python scripts/cierre_historico_2b2.py --fuente <OCUPACION_2026.xlsx>
        --corte <OCU26_BASE_DATOS_23SEP_PRE_CARGA.xlsx> --salida <dir> [--aplicar]
"""

from __future__ import annotations

import argparse
import datetime as dt
import re
import sys
from collections import Counter
from pathlib import Path
from typing import Any

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import staging_import_historico as st  # noqa: E402
from validate_input import CAMPANAS_HEADERS, resolve_input_path  # noqa: E402

SHA_BASE_POST_2B1 = "a1e207df01f2d45c61584f4a03aac5e8a6beaeb1b4b55b66381dbd0d64fb0c30"
SHA_CORTE_PRE_CARGA = st.EXPECTED_BASE_SHA256


def _clave(ot: Any, ini: Any, fin: Any) -> tuple[int, dt.date, dt.date]:
    return int(ot), st._a_fecha(ini), st._a_fecha(fin)


def ts_de_origen(origen: str) -> int | None:
    """'REM-TS 1-V1 + REM-TS 1-V3' -> 1 (un unico numero o None)."""
    nums = {int(n) for n in re.findall(r"REM-TS[\s-]*(\d+)", str(origen))}
    return nums.pop() if len(nums) == 1 else None


# ---------------------------------------------------------------------------
# Ajustes deterministicos
# ---------------------------------------------------------------------------

def cambios_estado_extensiones(camp: pd.DataFrame, cargas_extendidas: list[str],
                               fecha_carga: dt.datetime) -> tuple[pd.DataFrame, pd.DataFrame]:
    """-> (cambios, control CargaID|IDCampaña|ElementoID|EstadoAnterior|EstadoNuevo|FechaFin)."""
    pos = {st.texto(v): k for k, v in enumerate(camp["CargaID"])}
    cambios, control = [], []
    for cid in cargas_extendidas:
        k = pos[cid]
        r = camp.iloc[k]
        anterior = st.texto(r["Estado"])
        fin = st._a_fecha(r["FechaFin"])
        nuevo = anterior if anterior == "Reservada" else st.estado_por_fechas(fin, fecha_carga)
        if nuevo != anterior:
            cambios.append((k, cid, "Estado", anterior, nuevo))
        control.append({"CargaID": cid, "IDCampaña": st._a_ot(r["IDCampaña"]), "ElementoID": st.texto(r["ElementoID"]),
                        "EstadoAnterior": anterior, "EstadoNuevo": nuevo, "FechaFin": fin,
                        "Accion": "SIN CAMBIO (Reservada, no se modifica)" if anterior == "Reservada"
                        else ("CAMBIA" if nuevo != anterior else "SIN CAMBIO")})
    return pd.DataFrame(cambios, columns=st.CAMBIOS_COLUMNAS), pd.DataFrame(control)


def cambios_fuente_carga(camp: pd.DataFrame) -> pd.DataFrame:
    """Filas de esta migracion con FuenteCarga fuera de la convencion canonica."""
    filas = [(k, st.texto(r["CargaID"]), "FuenteCarga", st.texto(r["FuenteCarga"]), st.FUENTE_CARGA)
             for k, r in enumerate(camp.to_dict("records"))
             if st.texto(r["UsuarioCarga"]) == st.USUARIO_CARGA and st.texto(r["FuenteCarga"]) != st.FUENTE_CARGA]
    return pd.DataFrame(filas, columns=st.CAMBIOS_COLUMNAS)


def conciliar_remeros(deseadas: pd.DataFrame, actuales: pd.DataFrame, camp: pd.DataFrame,
                      medio_por_id: dict[str, str]) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """Concilia las asignaciones REM-TS deseadas (regla posicional) con las filas
    REM-TS cargadas en 2B.1.

    deseadas: IMPORTAR/INSERT del staging recalculado (IDCampaña, ElementoID,
    FechaInicio, FechaFin, TRZ_ElementoOrigen). actuales: filas 2B.1 (CargaID,
    IDCampaña, ElementoID, FechaInicio, FechaFin, TRZ_ElementoOrigen).
    Clave de conciliacion: (OT, inicio, fin, numero REM-TS).
    -> (cambios de reasignacion, deseadas sin fila actual = INSERT, errores)."""
    errores: list[str] = []
    pos = {st.texto(v): k for k, v in enumerate(camp["CargaID"])}
    act = {}
    for r in actuales.itertuples(index=False):
        key = _clave(r.IDCampaña, r.FechaInicio, r.FechaFin) + (ts_de_origen(r.TRZ_ElementoOrigen),)
        if key in act:
            errores.append(f"Filas 2B.1 REM duplicadas para {key}")
        act[key] = r
    cambios, nuevas_idx, usadas = [], [], set()
    for idx, d in deseadas.iterrows():
        key = _clave(d["IDCampaña"], d["FechaInicio"], d["FechaFin"]) + (ts_de_origen(d["TRZ_ElementoOrigen"]),)
        r = act.get(key)
        if r is None:
            nuevas_idx.append(idx)
            continue
        usadas.add(key)
        if r.ElementoID == d["ElementoID"]:
            continue
        k = pos[r.CargaID]
        fila = camp.iloc[k]
        if medio_por_id.get(d["ElementoID"]) != st.texto(fila["TipoCargaDeclarado"]):
            errores.append(f"{r.CargaID}: el medio del nuevo ElementoID no coincide con TipoCargaDeclarado")
        ot, ini, fin = key[:3]
        cambios += [(k, r.CargaID, "ElementoID", r.ElementoID, d["ElementoID"]),
                    (k, r.CargaID, "ClaveNegocio", st.texto(fila["ClaveNegocio"]),
                     st.clave_negocio(ot, d["ElementoID"], ini, fin))]
    sobrantes = set(act) - usadas
    if sobrantes:
        errores.append(f"Filas REM de 2B.1 sin asignacion deseada (requeririan borrado): {sorted(sobrantes)[:5]}")
    return pd.DataFrame(cambios, columns=st.CAMBIOS_COLUMNAS), deseadas.loc[nuevas_idx], errores


def _firma_ops(imp: pd.DataFrame) -> set[tuple]:
    return {(int(o), st.texto(e), st._a_fecha(a), st._a_fecha(b))
            for o, e, a, b in zip(imp["IDCampaña"], imp["ElementoID"], imp["FechaInicio"], imp["FechaFin"])}


def _es_rem(df: pd.DataFrame) -> pd.Series:
    return df["TRZ_ElementoOrigen"].astype(str).str.contains("REM-TS")


# ---------------------------------------------------------------------------
# Principal
# ---------------------------------------------------------------------------

def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--fuente", required=True)
    ap.add_argument("--corte", required=True, help="Base PRE-CARGA (solo lectura)")
    ap.add_argument("--salida", required=True, help="Directorio del staging (contiene el staging 2B.1)")
    ap.add_argument("--base", default=None)
    ap.add_argument("--sha-base", default=SHA_BASE_POST_2B1)
    ap.add_argument("--aplicar", action="store_true")
    args = ap.parse_args(argv)

    fuente, corte, salida = Path(args.fuente), Path(args.corte), Path(args.salida)
    base = resolve_input_path(args.base)
    sha_fuente, sha_corte, sha_base = (st.sha256_file(p) for p in (fuente, corte, base))
    if sha_base != args.sha_base or sha_corte != SHA_CORTE_PRE_CARGA:
        print(f"ERROR: SHA inesperado base={sha_base} corte={sha_corte}", file=sys.stderr)
        return 2
    xlsx_prev = salida / "OCU26_HISTORICO_STAGING.xlsx"
    prev = {h: pd.read_excel(xlsx_prev, sheet_name=h) for h in ("IMPORTAR", "EXTENSIONES_UPDATE", "REVISAR")}
    camp = pd.read_excel(base, sheet_name="CAMPANAS")
    maestro = pd.read_excel(base, sheet_name="MAESTRO_ELEMENTOS")
    parametros = pd.read_excel(base, sheet_name="PARAMETROS")
    medio_por_id = dict(zip(maestro["ElementoID"].map(st.texto), maestro["Medio"].map(st.texto)))
    errores: list[str] = []

    # 0) El staging 2B.1 describe exactamente lo que hay en la base.
    por_id = camp.set_index(camp["CargaID"].map(st.texto))
    for r in prev["IMPORTAR"].itertuples(index=False):
        f = por_id.loc[r.CargaID] if r.CargaID in por_id.index else None
        if f is None or _clave(f["IDCampaña"], f["FechaInicio"], f["FechaFin"]) != _clave(r.IDCampaña, r.FechaInicio, r.FechaFin) \
                or st.texto(f["ElementoID"]) != r.ElementoID:
            errores.append(f"Staging 2B.1 no coincide con la base para {r.CargaID}")
    for r in prev["EXTENSIONES_UPDATE"].itertuples(index=False):
        if st._a_fecha(por_id.loc[r.CargaID, "FechaFin"]) != st._a_fecha(r.FechaFin_Nueva):
            errores.append(f"Extension 2B.1 no reflejada en la base: {r.CargaID}")
    cargas_2b1 = camp.loc[camp["UsuarioCarga"].map(st.texto) == st.USUARIO_CARGA, "FechaHoraCarga"].dropna().unique()
    if len(cargas_2b1) != 1:
        errores.append(f"FechaHoraCarga de 2B.1 no unica: {cargas_2b1}")
    fecha_carga_2b1 = pd.Timestamp(cargas_2b1[0]).to_pydatetime()

    # 1) Staging recalculado contra el corte PRE-CARGA (reglas vigentes).
    fecha_carga = dt.datetime.now().replace(microsecond=0)
    res = st.construir_staging(fuente, corte, fecha_carga=fecha_carga)
    imp, ext = res["hojas"]["IMPORTAR"], res["hojas"]["EXTENSIONES_UPDATE"]

    # 2) Determinismo: fuera de Remeros debe reproducir lo aplicado en 2B.1.
    if _firma_ops(imp[~_es_rem(imp)]) != _firma_ops(prev["IMPORTAR"][~_es_rem(prev["IMPORTAR"])]):
        errores.append("INSERT fuera de Remeros distintos de 2B.1")
    if set(zip(ext["CargaID"], ext["FechaFin_Nueva"].map(st._a_fecha))) != \
            set(zip(prev["EXTENSIONES_UPDATE"]["CargaID"], prev["EXTENSIONES_UPDATE"]["FechaFin_Nueva"].map(st._a_fecha))):
        errores.append("UPDATE de extension distintos de 2B.1")
    rv_prev, rv = prev["REVISAR"], res["hojas"]["REVISAR"]
    firma_rv = lambda d: Counter(zip([st.texto(v) for v in d["TRZ_ElementoOrigen"]],  # noqa: E731
                                     [st._a_ot(v) for v in d["IDCampaña"]],
                                     [st.texto(v) for v in d["TRZ_MesesPresencia"]],
                                     [st.texto(v) for v in d["Motivo"]]))
    if firma_rv(rv_prev[~_es_rem(rv_prev)]) != firma_rv(rv[~_es_rem(rv)]):
        errores.append("REVISAR fuera de Remeros distinto de 2B.1")

    # 3) Remeros: reasignacion de filas 2B.1 + nuevas inserciones.
    cambios_rem, nuevas_rem, err_rem = conciliar_remeros(imp[_es_rem(imp)], prev["IMPORTAR"][_es_rem(prev["IMPORTAR"])],
                                                         camp, medio_por_id)
    errores += err_rem
    if (ext["ElementoID"].astype(str).str.startswith("REM-DB")).any():
        errores.append("Hay extensiones sobre REM-DB: no previstas en el cierre")

    # 4) Estados de extensiones y 5) FuenteCarga canonica.
    cambios_estado, control_estado = cambios_estado_extensiones(camp, list(prev["EXTENSIONES_UPDATE"]["CargaID"]), fecha_carga_2b1)
    cambios_fuente = cambios_fuente_carga(camp)
    cambios = pd.concat([cambios_rem, cambios_estado, cambios_fuente], ignore_index=True)

    # Filas nuevas: CargaID correlativo desde la base actual.
    nuevas = nuevas_rem.copy()
    primer = st.siguiente_carga_id(camp)
    nuevas["CargaID"] = [f"{st.PREFIJO_CARGA_ID}{primer + k:08d}" for k in range(len(nuevas))]
    nuevas["FuenteCarga"] = st.FUENTE_CARGA
    esperado = st.campanas_esperadas(camp, cambios, nuevas)
    tocadas = set(cambios["CargaID"]) | set(nuevas["CargaID"])
    errores += st.validar_estado_final(camp, esperado, maestro, parametros, tocadas)

    # Tablas de control
    pos_rem = {n + 1: e for n, e in enumerate(st.IndiceMaestro(maestro).rem_digital)}
    a = res["asignaciones"]
    a_rem = a[a["TRZ_ElementoOrigen"].astype(str).str.contains("REM-TS")].assign(
        TS=lambda d: d["TRZ_ElementoOrigen"].map(ts_de_origen))
    reasignadas = Counter(ts_de_origen(o) for o, cid in zip(prev["IMPORTAR"]["TRZ_ElementoOrigen"], prev["IMPORTAR"]["CargaID"])
                          if cid in set(cambios_rem["CargaID"]))
    insert_ts = Counter(nuevas_rem["TRZ_ElementoOrigen"].map(ts_de_origen))
    prev_rem = prev["IMPORTAR"][_es_rem(prev["IMPORTAR"])]
    tabla_rem = []
    for n in range(1, 7):
        g = a_rem[a_rem["TS"] == n]
        correctas = sum(1 for o, e in zip(prev_rem["TRZ_ElementoOrigen"], prev_rem["ElementoID"])
                        if ts_de_origen(o) == n and e == pos_rem.get(n))
        tabla_rem.append({
            "REM-TS historico": f"REM-TS {n}", "Posicion": n, "ElementoID actual": pos_rem.get(n, ""),
            "Asignaciones detectadas": len(g),
            "YA_EXISTE (cargadas en 2B.1 en el soporte correcto)": correctas,
            "INSERT": insert_ts.get(n, 0),
            "UPDATE (reasignacion ElementoID de filas 2B.1)": reasignadas.get(n, 0),
            "REVISAR": int((g["Estado_Staging"] == "REVISAR").sum()),
            "NO_IMPORTAR": int((g["Estado_Staging"] == "NO_IMPORTAR").sum()),
        })
    tabla_rem = pd.DataFrame(tabla_rem)
    resueltos_rem = int(_es_rem(rv_prev).sum() - _es_rem(rv).sum())
    mot = lambda d: Counter(re.sub(r":.*", "", m) for ms in d["Motivo"].astype(str) for m in ms.split(" | ") if m)  # noqa: E731
    mp, mn = mot(rv_prev), mot(rv)
    revisar = pd.DataFrame([{"Motivo": "TOTAL REVISAR (asignaciones)", "Antes (2B.1)": len(rv_prev), "Despues (2B.2)": len(rv)}] +
                           [{"Motivo": k, "Antes (2B.1)": mp.get(k, 0), "Despues (2B.2)": mn.get(k, 0)} for k in sorted(set(mp) | set(mn))])

    print("Errores previos:", errores or "ninguno")
    print(tabla_rem.to_string(index=False))
    print(control_estado["Accion"].value_counts().to_dict(), control_estado.groupby(["EstadoAnterior", "EstadoNuevo"]).size().to_dict())
    print(f"Cambios: reasignacion REM {cambios_rem['CargaID'].nunique()} filas, Estado {len(cambios_estado)}, "
          f"FuenteCarga {len(cambios_fuente)}; INSERT nuevas {len(nuevas)}")
    print(f"REVISAR 2B.1 {len(rv_prev)} -> {len(rv)} (resueltos Remeros {resueltos_rem})")

    antes = {"filas": len(camp), "ot_distintas": camp["IDCampaña"].map(st._a_ot).dropna().nunique(), "sha256": sha_base}
    despues, ctrl = None, None
    if errores:
        print("ABORTADO, base no modificada", file=sys.stderr)
    elif args.aplicar:
        ctrl = st.aplicar_operaciones(base, camp, cambios, nuevas, "etapa2b2")
        final = pd.read_excel(base, sheet_name="CAMPANAS")
        despues = {"filas": len(final), "ot_distintas": final["IDCampaña"].map(st._a_ot).dropna().nunique(),
                   "fecha_hora_carga_max": final["FechaHoraCarga"].max(), "sha256": st.sha256_file(base)}
        print("BASE DESPUES:", despues, ctrl["validate_input"], ctrl["partes_modificadas"])

    # Salidas: staging recalculado con los valores reales de la base.
    if not errores:
        final_ref = st.campanas_esperadas(camp, cambios, nuevas)
        propias = final_ref[final_ref["UsuarioCarga"].map(st.texto) == st.USUARIO_CARGA].set_index("ClaveNegocio")
        imp = imp.copy()
        for c in CAMPANAS_HEADERS:
            if c != "ClaveNegocio":
                imp[c] = imp["ClaveNegocio"].map(propias[c])
        for c in st.COLUMNAS_FECHA:
            imp[c] = imp[c].map(st._a_fecha)
        res["hojas"]["IMPORTAR"] = imp
        est = control_estado.set_index("CargaID")
        ext = ext.copy()
        ext.insert(ext.columns.get_loc("FechaFin_Nueva") + 1, "EstadoAnterior", ext["CargaID"].map(est["EstadoAnterior"]))
        ext.insert(ext.columns.get_loc("EstadoAnterior") + 1, "EstadoNuevo", ext["CargaID"].map(est["EstadoNuevo"]))
        res["hojas"]["EXTENSIONES_UPDATE"] = ext
        resumen = st.calcular_resumen(res)
        resumen["comparacion_revisar"] = revisar
        meta = [("Etapa", "2B.2 - cierre"), ("Fuente", str(fuente)), ("SHA-256 fuente", sha_fuente),
                ("Base (post 2B.1)", f"{base} {sha_base}"), ("Corte PRE-CARGA (lectura)", f"{corte} {sha_corte}"),
                ("Fecha de carga 2B.1", f"{fecha_carga_2b1:%Y-%m-%d %H:%M:%S}"),
                ("Fecha de carga 2B.2 (filas nuevas)", f"{fecha_carga:%Y-%m-%d %H:%M:%S}"),
                ("Aplicado", "SI" if ctrl else "NO")]
        st.escribir_excel(salida / "OCU26_HISTORICO_STAGING.xlsx", res, resumen, meta)
        st.escribir_csv(salida / "OCU26_HISTORICO_IMPORTAR.csv", imp)
        st.escribir_md(salida / "RESUMEN_IMPORTACION_HISTORICA.md", res, resumen, meta)
        reasig = cambios_rem[cambios_rem["Columna"] == "ElementoID"].rename(
            columns={"Anterior": "ElementoID_2B1", "Nuevo": "ElementoID_posicional"})[["CargaID", "ElementoID_2B1", "ElementoID_posicional"]]
        ops = pd.concat([cambios.assign(Operacion="UPDATE"),
                         nuevas[["CargaID", "IDCampaña", "ElementoID", "FechaInicio", "FechaFin", "ClaveNegocio", "Estado"]].assign(Operacion="INSERT")],
                        ignore_index=True)
        ops.to_csv(salida / "OCU26_HISTORICO_OPERACIONES_2B2.csv", index=False, encoding="utf-8-sig")
        l = ["# Reporte de cierre - Etapa 2B.2\n"]
        l.append(st._md_tabla(pd.DataFrame(meta, columns=["Clave", "Valor"])))
        l.append("\n## Base antes / despues\n")
        l.append(st._md_tabla(pd.DataFrame([antes] + ([despues] if despues else []))))
        l.append("\n## Remeros (regla posicional)\n")
        l.append(st._md_tabla(tabla_rem))
        l.append(f"\n### Filas 2B.1 reasignadas ({len(reasig)})\n")
        l.append(st._md_tabla(reasig))
        l.append(f"\n### Filas nuevas ({len(nuevas)})\n")
        l.append(st._md_tabla(nuevas[["CargaID", "IDCampaña", "ElementoID", "FechaInicio", "FechaFin", "Estado", "Campaña"]]))
        l.append("\n## Estado de las extensiones de 2B.1\n")
        l.append(st._md_tabla(control_estado.groupby(["EstadoAnterior", "EstadoNuevo", "Accion"]).size().reset_index(name="Filas")))
        l.append("\n" + st._md_tabla(control_estado))
        l.append(f"\n## FuenteCarga\n\n{len(cambios_fuente)} filas de MIGRACION_OCUPACION_2026 -> \"{st.FUENTE_CARGA}\" "
                 "(vocabulario PARAMETROS; procedencia en UsuarioCarga).\n")
        l.append("\n## REVISAR\n")
        l.append(st._md_tabla(revisar))
        if ctrl:
            l.append(f"\n## Controles de escritura\n\n- Partes modificadas: {', '.join(ctrl['partes_modificadas'])}\n"
                     f"- tblCampanas: {ctrl['tabla_ref']} (cubre filas: {ctrl['tabla_cubre_filas']})\n"
                     f"- validate_input: {ctrl['validate_input']} ({len(ctrl['validate_errors'])} errores)\n")
        (salida / "REPORTE_CIERRE_ETAPA2B2.md").write_text("".join(x if x.endswith("\n") else x + "\n" for x in l), encoding="utf-8")
    if st.sha256_file(fuente) != sha_fuente or st.sha256_file(corte) != sha_corte:
        print("ERROR: fuente o corte cambiaron", file=sys.stderr)
        return 3
    return 4 if errores else 0


if __name__ == "__main__":
    sys.exit(main())
