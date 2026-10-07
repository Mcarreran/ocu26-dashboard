"""Revision final del historico OCU26 - Etapa 2C (depuracion pre-SharePoint).

Parte de la base canonica post 2B.2 (SHA esperado) y re-evalua, solo con
evidencia, las asignaciones que quedaron en REVISAR al cierre 2B.2:

1. Recalcula el staging (reglas 2B vigentes) contra la base ACTUAL y
   reconcilia exactamente el universo REVISAR del staging 2B.2: cada
   asignacion sigue en REVISAR o pasa a IMPORTAR por una regla ya confirmada
   (periodo disjunto de una OT cuyo otro periodo ya esta cargado).
2. Reglas 2C, a nivel de celda de la fuente:
   - EXTENSION_FIN_DEMOSTRADA: la extension de 2B que solo quedo bloqueada por
     dudas ubicadas en meses que la fila existente ya cubre. Los meses nuevos
     tienen celdas propias de la OT con el mismo INICIO que la fila, FIN no
     decreciente bloque a bloque y presencia hasta el mes del FIN.
   - PERIODO_DISJUNTO_EXPLICITO: tramo nuevo de una OT + elemento ya cargada,
     sin contacto con sus filas, con el mismo par INICIO/FIN en todas las
     celdas (propias) de sus meses y presencia en todos ellos.
   - SIN_OT_CUBIERTA_POR_LA_MISMA_FILA: celda sin OT cuya fila muestra una
     unica OT con la misma pauta y las mismas fechas, ya cargada cubriendo
     esos meses. No genera operacion (no se reconstruye ninguna OT).
   - PRESENCIA_YA_CUBIERTA_EN_CAMPANAS (solo con --cerrar-presencia-cubierta,
     decision de negocio): CAMPANAS ya cubre todos los meses en que la grilla
     muestra la OT en el elemento; la fuente solo declara un INICIO anterior o
     un FIN posterior que la propia grilla no muestra. Se cierra sin cambios.
3. Todo lo demas queda en el archivo de pendientes humanos con motivo,
   evidencia, accion sugerida y responsable (EQUIPO / USUARIO).

Sin --aplicar no modifica la base (dry-run): escribe staging de revision,
operaciones propuestas, pendientes y resumen en --salida. Con --aplicar edita
quirurgicamente solo CAMPANAS (escritor de 2B: temporal, verificacion,
reemplazo atomico). Sobre la base canonica, antes de escribir, congela una
copia POST-2B.2 en --salida (SHA, filas y OT en un manifiesto).

Uso:
    python scripts/cierre_historico_2c.py --fuente <OCUPACION_2026.xlsx>
        --staging-2b2 <OCU26_HISTORICO_STAGING.xlsx> --salida <dir>
        [--base <xlsx>] [--sha-base <sha256>] [--cerrar-presencia-cubierta] [--aplicar]
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import shutil
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import staging_import_historico as st  # noqa: E402
from validate_input import resolve_input_path  # noqa: E402

SHA_BASE_POST_2B2 = "1c41aea8357a63a1e9f73e1184eb34421e30ff9129af15da4c04571d0633089c"
FILAS_POST_2B2, OT_POST_2B2, REVISAR_2B2 = 16365, 633, 458
NOMBRE_COPIA_CONGELADA = "OCU26_BASE_DATOS_POST_2B2_PRE_2C.xlsx"

R_DISJUNTO_2B = "NUEVO_PERIODO_DISJUNTO_DE_OT_YA_CARGADA"
R_EXTENSION = "EXTENSION_FIN_DEMOSTRADA"
R_DISJUNTO = "PERIODO_DISJUNTO_EXPLICITO"
R_SIN_OT = "SIN_OT_CUBIERTA_POR_LA_MISMA_FILA"
R_CUBIERTA = "PRESENCIA_YA_CUBIERTA_EN_CAMPANAS"

TIPOS_FECHA_VALIDOS = ("FECHA", "SERIAL")
UN_DIA = st.UN_DIA

# Motivo principal 2B.2 (prioridad del reporte de cierre) -> etiqueta legible.
MOTIVO_PRINCIPAL_2B2 = [
    ("ELEMENTO_NO_INEQUIVOCO_BAJA", "Elemento sin match (BAJA)"),
    ("ELEMENTO_NO_INEQUIVOCO_MEDIA", "Elemento ambiguo (MEDIA)"),
    ("ESTRUCTURA_FILA_SIN_CODIGO_ELEMENTO", "Fila sin código"),
    ("SIN_OT_NUMERICA", "Sin OT numérica"),
    ("SOLAPA_OTRA_CAMPANA_EN_ELEMENTO_ESTATICO", "Solapa otra campaña en elemento estático"),
    ("EXTENSION_CAMBIA_FECHAINICIO", "Extensión que cambiaría FechaInicio"),
    ("EXTENSION_AMBIGUA_VARIAS_FILAS_EXISTENTES", "Extensión ambigua (varias filas)"),
    ("FECHAS_NO_DETERMINABLES", "Fechas no determinables"),
    ("FECHAS_DUDOSAS", "Fechas dudosas"),
    ("OT_ELEMENTO_EXISTE_CON_OTRAS_FECHAS", "OT + elemento existen con otras fechas"),
]

ACCION_SUGERIDA = {
    "ELEMENTO_SIN_EQUIVALENCIA_EN_MAESTRO": "Confirmar si el soporte existe: si existe, alta en MAESTRO_ELEMENTOS (etapa aparte) y recarga; si no, descartar las campañas de esta fila.",
    "ELEMENTO_AMBIGUO": "Indicar a cuál de los candidatos corresponde el código histórico (o si es un soporte inexistente en el maestro).",
    "FILA_SIN_CODIGO_DE_ELEMENTO": "Identificar el soporte de la fila (no tiene código ni atributos) o descartarla.",
    "SIN_OT_NUMERICA": "Informar la OT numérica o confirmar que es uso no comercial (sin OT no se importa).",
    "SOLAPA_OTRA_CAMPANA_EN_ESTATICO": "Confirmar qué campaña ocupó el soporte estático en el período superpuesto (posible OT mal tipeada o reemplazo de pauta).",
    "EXTENSION_CAMBIARIA_FECHAINICIO": "Confirmar si la FechaInicio cargada debe adelantarse; 2C no modifica FechaInicio.",
    "EXTENSION_AMBIGUA_VARIAS_FILAS": "Indicar cuál de las filas existentes de la OT en el elemento se extiende.",
    "OT_ELEMENTO_EXISTE_CON_OTRAS_FECHAS": "Confirmar si es un tramo nuevo o una corrección de fechas de la fila ya cargada.",
    "FECHAS_SOLO_EN_CELDA_COMPARTIDA": "Informar INICIO/FIN de cada OT: la grilla solo trae fechas en celdas con varias OT.",
    "FECHA_INDETERMINADA_O_INCOMPLETA": "Informar INICIO y FIN: la grilla trae fechas vacías, INDET o no interpretables.",
    "PRESENCIA_YA_CUBIERTA_EN_CAMPANAS": "Sin cambio salvo indicación: CAMPANAS ya cubre todos los meses en que la grilla muestra la OT; confirmar si el INICIO/FIN declarado en la grilla es real.",
    "FECHAS_CONTRADICTORIAS_ENTRE_BLOQUES": "Confirmar el período vigente: los bloques mensuales declaran INICIO/FIN distintos para la misma OT.",
    "FIN_NO_CONFIRMADO_POR_LA_GRILLA": "Confirmar FechaFin real: la grilla deja de mostrar la OT antes del FIN declarado.",
    "INICIO_ANTERIOR_A_LA_PRESENCIA": "Confirmar FechaInicio real: la grilla muestra la OT recién meses después del INICIO declarado.",
    "PRESENCIA_FUERA_DEL_RANGO_DECLARADO": "Confirmar fechas: la OT figura en meses fuera de su propio INICIO/FIN.",
    "FECHAS_CONTRADICEN_LA_MISMA_OT": "Confirmar el período en este elemento: la grilla declara días en los que la misma OT no está cargada en ninguno de sus otros elementos.",
    "FECHAS_DUDOSAS_OTRAS": "Revisar fechas de la OT en la grilla.",
}
RESPONSABLE_USUARIO = {"ELEMENTO_SIN_EQUIVALENCIA_EN_MAESTRO", "ELEMENTO_AMBIGUO", "FILA_SIN_CODIGO_DE_ELEMENTO"}


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def motivos_base(motivo: Any) -> list[str]:
    """'A:x | B' -> ['A', 'B']."""
    return [re.sub(r":.*", "", m).strip() for m in st.texto(motivo).split(" | ") if m.strip()]


def motivo_principal_2b2(motivo: Any) -> str:
    ms = set(motivos_base(motivo))
    for clave, etiqueta in MOTIVO_PRINCIPAL_2B2:
        if clave in ms:
            return etiqueta
    return "Otro"


def firma(reg: Any) -> tuple:
    """Identidad estable de una asignacion entre corridas del staging (sin Motivo)."""
    g = reg.get if isinstance(reg, dict) else reg.__getitem__
    return (st.texto(g("TRZ_HojaOrigen")), st.texto(g("TRZ_ElementoOrigen")), st._a_ot(g("IDCampaña")),
            st.texto(g("TRZ_TipoGrupo")), st.texto(g("TRZ_MesesPresencia")), st.texto(g("TRZ_RefCeldas")),
            st._a_fecha(g("FechaInicio")), st._a_fecha(g("FechaFin")))


def ocurrencias_por_fila(ocurrencias: list[st.Ocurrencia]) -> dict[tuple[str, int], list[st.Ocurrencia]]:
    out: dict[tuple[str, int], list[st.Ocurrencia]] = defaultdict(list)
    for o in ocurrencias:
        out[(o.hoja, o.fila)].append(o)
    return out


def _filas_de(reg: dict) -> list[tuple[str, int]]:
    hojas = st.texto(reg["TRZ_HojaOrigen"]).split(" + ")
    filas = [int(f) for f in st.texto(reg["TRZ_FilasOrigen"]).split(", ") if f]
    return [(h, f) for h in hojas for f in filas]


def ocurrencias_del_registro(reg: dict, occ_fila: dict) -> list[st.Ocurrencia]:
    """Celdas de la fuente de la asignacion: su OT (o sin OT) en sus filas y codigos."""
    codigos = set(st.texto(reg["TRZ_ElementoOrigen"]).split(" + "))
    ot = st._a_ot(reg["IDCampaña"])
    meses = set(st.texto(reg["TRZ_MesesPresencia"]).split(", "))
    out = []
    for key in _filas_de(reg):
        for o in occ_fila.get(key, []):
            if o.codigo not in codigos and o.codigo:
                continue
            if reg["TRZ_TipoGrupo"] == "OT" or (reg["TRZ_TipoGrupo"] == "FILA_SIN_CODIGO" and ot is not None):
                if o.token is None or o.token.numero != ot:
                    continue
            elif o.token is not None or o.mes.strftime("%Y-%m") not in meses:
                continue
            out.append(o)
    return out


def _cubre(filas: list[st.FilaCampana], mes: dt.date) -> bool:
    return any(a <= st.fin_mes(mes) and mes <= b for a, b in st._intervalos(filas))


def _fmt_filas(filas: list[st.FilaCampana]) -> str:
    return "; ".join(f"{f.carga_id} {st.fmt_fecha(f.ini)}→{st.fmt_fecha(f.fin)}" for f in sorted(filas, key=lambda f: (f.ini or dt.date.min)))


def _fmt_celdas(occ: list[st.Ocurrencia]) -> str:
    """'2026-03..2026-05 03-01→06-30 (propia); 2026-06 03-01→07-31 (compartida)'."""
    tramos: list[list] = []
    for o in sorted(occ, key=lambda o: (o.mes, o.ref)):
        clave = (st.fmt_fecha(o.ini) or o.tipo_ini, st.fmt_fecha(o.fin) or o.tipo_fin, "compartida" if o.compartida else "propia")
        if tramos and tramos[-1][2] == clave and (tramos[-1][1] == o.mes or st.sumar_meses(tramos[-1][1], 1) == o.mes):
            tramos[-1][1] = o.mes
        elif not (tramos and tramos[-1][2] == clave and tramos[-1][1] == o.mes):
            tramos.append([o.mes, o.mes, clave])
    return "; ".join(f"{a:%Y-%m}{'' if a == b else f'..{b:%Y-%m}'} {c[0]}→{c[1]} ({c[2]})" for a, b, c in tramos)


# ---------------------------------------------------------------------------
# Reglas 2C (funciones puras sobre celdas de la fuente + filas de CAMPANAS)
# ---------------------------------------------------------------------------

def evaluar_extension_demostrada(motivo: str, ini: dt.date | None, fin: dt.date | None, occ: list[st.Ocurrencia],
                                 filas: list[st.FilaCampana], ventana: set[dt.date]) -> tuple[st.FilaCampana | None, str]:
    """Extension de FechaFin de UNA fila existente demostrada por las celdas propias.

    -> (fila a extender, evidencia) o (None, motivo por el que no aplica)."""
    ms = set(motivos_base(motivo))
    if "EXISTE_EN_CAMPANAS_COBERTURA_PARCIAL" not in ms or not ms <= {"FECHAS_DUDOSAS", "EXISTE_EN_CAMPANAS_COBERTURA_PARCIAL"}:
        return None, "NO_APLICA_MOTIVO"
    if not (ini and fin):
        return None, "SIN_PERIODO"
    fila, motivo_ext = st.evaluar_extension(filas, ini, fin)
    if fila is None:
        return None, motivo_ext
    if ini != fila.ini:
        return None, "INICIO_FUENTE_DISTINTO_DE_FILA"
    propias = [o for o in occ if not o.compartida]
    fines_mes: dict[dt.date, set[dt.date]] = defaultdict(set)
    for o in propias:
        if o.ini == fila.ini and o.fin and o.tipo_fin in TIPOS_FECHA_VALIDOS and o.fin >= o.ini:
            fines_mes[o.mes].add(o.fin)
    if any(len(v) > 1 for v in fines_mes.values()):
        return None, "FIN_DISTINTO_EN_UN_MISMO_MES"
    fines = [next(iter(fines_mes[m])) for m in sorted(fines_mes)]
    if not fines or any(a > b for a, b in zip(fines, fines[1:])):
        return None, "FIN_DECRECE_ENTRE_BLOQUES"
    if fines[-1] != fin:
        return None, "ULTIMO_FIN_DISTINTO_DEL_PERIODO"
    meses_ext = [m for m in st.meses_entre(fila.fin + UN_DIA, fin) if m in ventana]
    if not meses_ext:
        return None, "EXTENSION_FUERA_DE_LA_VENTANA"
    for m in meses_ext:
        cm = [o for o in occ if o.mes == m]
        if not cm:
            return None, f"MES_SIN_PRESENCIA_EN_EXTENSION:{m:%Y-%m}"
        if any(o.compartida for o in cm):
            return None, f"CELDA_COMPARTIDA_EN_EXTENSION:{m:%Y-%m}"
        for o in cm:
            if (o.tipo_ini not in TIPOS_FECHA_VALIDOS or o.tipo_fin not in TIPOS_FECHA_VALIDOS or o.ini != fila.ini
                    or o.fin < o.ini or o.fin < o.mes):
                return None, f"CELDA_INCONSISTENTE_EN_EXTENSION:{o.ref}"
    if any(o.mes > st.inicio_mes(fin) for o in occ):
        return None, "PRESENCIA_POSTERIOR_AL_FIN"
    mes_fin = st.inicio_mes(fin)
    if mes_fin in ventana and not any(o.mes == mes_fin and o.fin == fin and not o.compartida for o in occ):
        return None, "FIN_SIN_CELDA_PROPIA_EN_SU_MES"
    evid = (f"Fila {fila.carga_id} {fila.ini}→{fila.fin}; celdas propias con el mismo INICIO en "
            f"{meses_ext[0]:%Y-%m}..{meses_ext[-1]:%Y-%m}, FIN por bloque {' → '.join(str(f) for f in dict.fromkeys(fines))}")
    return fila, evid


def evaluar_periodo_disjunto(motivo: str, ini: dt.date | None, fin: dt.date | None, occ: list[st.Ocurrencia],
                             filas: list[st.FilaCampana], ventana: set[dt.date]) -> tuple[bool, str]:
    """Tramo nuevo, explicito y sin contacto con las filas de la misma OT + elemento."""
    ms = set(motivos_base(motivo))
    if "OT_ELEMENTO_EXISTE_CON_OTRAS_FECHAS" not in ms or not ms <= {"FECHAS_DUDOSAS", "OT_ELEMENTO_EXISTE_CON_OTRAS_FECHAS"}:
        return False, "NO_APLICA_MOTIVO"
    if not (ini and fin) or not filas:
        return False, "SIN_PERIODO_O_SIN_FILAS"
    for f in filas:
        f_fin = f.fin or dt.date.max - UN_DIA
        if f.ini and f.ini <= fin + UN_DIA and ini <= f_fin + UN_DIA:
            return False, f"TOCA_FILA_EXISTENTE:{f.carga_id}"
    meses_p = set(st.meses_entre(ini, fin))
    del_periodo = [o for o in occ if o.mes in meses_p]
    if any(o.compartida for o in del_periodo):
        return False, "CELDA_COMPARTIDA_EN_EL_PERIODO"
    if any((o.ini, o.fin) != (ini, fin) or o.tipo_ini not in TIPOS_FECHA_VALIDOS or o.tipo_fin not in TIPOS_FECHA_VALIDOS
           for o in del_periodo):
        return False, "PARES_DISTINTOS_EN_EL_PERIODO"
    faltan = sorted(m for m in meses_p if m in ventana and not any(o.mes == m for o in del_periodo))
    if faltan:
        return False, "MES_SIN_PRESENCIA:" + st.fmt_meses(faltan)
    if any((o.ini, o.fin) == (ini, fin) for o in occ if o.mes not in meses_p):
        return False, "PAR_DEL_PERIODO_FUERA_DE_SUS_MESES"
    return True, (f"Celdas propias {st.fmt_meses(sorted({o.mes for o in del_periodo}))} con {ini}→{fin}; "
                  f"sin contacto con {_fmt_filas(filas)}")


def evaluar_sin_ot_cubierta(occ_sin_ot: list[st.Ocurrencia], occ_filas: list[st.Ocurrencia],
                            filas_por_ot: dict[int, list[st.FilaCampana]]) -> tuple[int | None, str]:
    """Celda sin OT: la propia fila identifica una unica OT (misma pauta y fechas) ya cargada en esos meses."""
    pares = {(o.pauta.upper(), o.ini, o.fin) for o in occ_sin_ot}
    if len(pares) != 1:
        return None, "CELDAS_SIN_OT_HETEROGENEAS"
    pauta, ini, fin = next(iter(pares))
    if not (pauta and ini and fin):
        return None, "SIN_PAUTA_O_FECHAS"
    ots = {o.token.numero for o in occ_filas if o.token is not None and (o.pauta.upper(), o.ini, o.fin) == (pauta, ini, fin)}
    if len(ots) != 1:
        return None, f"OT_IDENTIFICABLE_NO_UNICA:{sorted(ots)}"
    ot = next(iter(ots))
    filas = filas_por_ot.get(ot, [])
    meses = sorted({o.mes for o in occ_sin_ot})
    if not filas or not all(_cubre(filas, m) for m in meses):
        return None, f"OT_{ot}_NO_CUBRE_LOS_MESES"
    return ot, (f"Misma fila: OT {ot} con pauta '{pauta}' y {ini}→{fin}; CAMPANAS {_fmt_filas(filas)} "
                f"cubre {st.fmt_meses(meses)}")


def contradice_misma_ot(desde: dt.date, hasta: dt.date, filas_otros: list[st.FilaCampana]) -> list[tuple[dt.date, dt.date]]:
    """Dias nuevos [desde, hasta] que la misma OT, cargada en OTROS elementos, no cubre.

    Sin filas en otros elementos no hay evidencia en contra -> []. Si la OT esta
    cargada en otros elementos y no cubre esos dias (p.ej. un hueco entre dos
    periodos), la fuente de este elemento contradice al resto -> no inequivoco."""
    if not filas_otros:
        return []
    return st.tramos_no_cubiertos(desde, hasta, st._intervalos(filas_otros))


def presencia_cubierta(occ: list[st.Ocurrencia], filas: list[st.FilaCampana]) -> bool:
    """Todos los meses en que la grilla muestra la OT estan cubiertos por sus filas de CAMPANAS."""
    meses = {o.mes for o in occ}
    return bool(filas) and bool(meses) and all(_cubre(filas, m) for m in meses)


def diagnostico_pendiente(reg: dict, occ: list[st.Ocurrencia], filas: list[st.FilaCampana],
                          ventana: set[dt.date], cubierta: bool) -> str:
    """Motivo de revision 2C (accionable) a partir de los motivos 2B y de las celdas."""
    ms = set(motivos_base(reg["Motivo"]))
    probs = st.texto(reg["TRZ_ProblemasFechas"])
    if "ELEMENTO_NO_INEQUIVOCO_BAJA" in ms:
        return "ELEMENTO_SIN_EQUIVALENCIA_EN_MAESTRO"
    if "ELEMENTO_NO_INEQUIVOCO_MEDIA" in ms:
        return "ELEMENTO_AMBIGUO"
    if "ESTRUCTURA_FILA_SIN_CODIGO_ELEMENTO" in ms:
        return "FILA_SIN_CODIGO_DE_ELEMENTO"
    if "SIN_OT_NUMERICA" in ms:
        return "SIN_OT_NUMERICA"
    if "SOLAPA_OTRA_CAMPANA_EN_ELEMENTO_ESTATICO" in ms:
        return "SOLAPA_OTRA_CAMPANA_EN_ESTATICO"
    if "EXTENSION_CAMBIA_FECHAINICIO" in ms:
        return "EXTENSION_CAMBIARIA_FECHAINICIO"
    if "EXTENSION_AMBIGUA_VARIAS_FILAS_EXISTENTES" in ms:
        return "EXTENSION_AMBIGUA_VARIAS_FILAS"
    if "FECHAS_SOLO_EN_CELDA_COMPARTIDA" in probs:
        return "FECHAS_SOLO_EN_CELDA_COMPARTIDA"
    if re.search(r"FECHA_INDETERMINADA|FECHA_INCOMPLETA|FECHA_NO_INTERPRETABLE|SIN_FECHAS_DETERMINABLES", probs):
        return "FECHA_INDETERMINADA_O_INCOMPLETA"
    if cubierta:
        return "PRESENCIA_YA_CUBIERTA_EN_CAMPANAS"
    if "FECHAS_DUDOSAS" not in ms:
        return "OT_ELEMENTO_EXISTE_CON_OTRAS_FECHAS"
    if re.search(r"RANGOS_DISTINTOS_SOLAPADOS|INICIO_POSTERIOR_A_FIN", probs):
        return "FECHAS_CONTRADICTORIAS_ENTRE_BLOQUES"
    ini, fin = st._a_fecha(reg["FechaInicio"]), st._a_fecha(reg["FechaFin"])
    pres = sorted({o.mes for o in occ})
    faltan = [m for m in re.findall(r"\d{4}-\d{2}", re.search(r"MESES_SIN_PRESENCIA:([^|]*)", probs).group(1))] \
        if "MESES_SIN_PRESENCIA" in probs else []
    if faltan and pres and fin and any(m > pres[-1].strftime("%Y-%m") for m in faltan):
        return "FIN_NO_CONFIRMADO_POR_LA_GRILLA"
    if faltan and pres and ini:
        return "INICIO_ANTERIOR_A_LA_PRESENCIA"
    if "PRESENCIA_FUERA_DE_RANGO" in probs:
        return "PRESENCIA_FUERA_DEL_RANGO_DECLARADO"
    if "FIN_EN_MES_SOLO_CELDA_COMPARTIDA" in probs:
        return "FECHAS_SOLO_EN_CELDA_COMPARTIDA"
    return "FECHAS_DUDOSAS_OTRAS"


def responsable(motivo_2c: str) -> str:
    return "USUARIO" if motivo_2c in RESPONSABLE_USUARIO else "EQUIPO"


def conflictos_estaticos(ops: list[tuple[str, int, dt.date, dt.date, Any]], idx_c: st.IndiceCampanas,
                         medio_por_id: dict[str, str]) -> set[Any]:
    """ops: (ElementoID, OT, desde, hasta, clave). Dias nuevos en estaticos que pisan otra OT
    (cargada o propuesta) -> claves en conflicto."""
    malas = set()
    for eid, ot, a, b, k in ops:
        if medio_por_id.get(eid) != "Estático":
            continue
        otros = {o for o, f in idx_c.por_elem.get(eid, []) if o != ot and f.ini and f.ini <= b and (f.fin or dt.date.max) >= a}
        otros |= {o2 for e2, o2, a2, b2, _ in ops if e2 == eid and o2 != ot and a2 <= b and a <= b2}
        if otros:
            malas.add(k)
    return malas


# ---------------------------------------------------------------------------
# Revision completa (sin efectos)
# ---------------------------------------------------------------------------

def revisar(res: dict[str, Any], previo_revisar: pd.DataFrame, cerrar_cubiertas: bool) -> dict[str, Any]:
    """Clasifica cada asignacion REVISAR 2B.2 -> resolucion 2C. Sin efectos sobre la base."""
    errores: list[str] = []
    a: pd.DataFrame = res["asignaciones"]
    camp: pd.DataFrame = res["campanas"]
    idx_c = st.IndiceCampanas(camp)
    campana_por_carga = dict(zip(camp["CargaID"].map(st.texto), camp["Campaña"].map(st.texto)))
    medio_por_id = {st.texto(e): st.texto(m) for e, m in zip(res["maestro"]["ElementoID"], res["maestro"]["Medio"])}
    occ_fila = ocurrencias_por_fila(res["ocurrencias"])

    # 1) Reconciliacion exacta con el universo REVISAR 2B.2.
    previo = {firma(r): r for r in previo_revisar.to_dict("records")}
    if len(previo) != len(previo_revisar):
        errores.append("Firmas no unicas en el REVISAR 2B.2")
    actuales = a[a["Estado_Staging"].isin(["REVISAR", "IMPORTAR"])]
    act = {firma(r): r for r in actuales.to_dict("records")}
    if len(act) != len(actuales):
        errores.append("Firmas no unicas en el staging recalculado")
    if set(act) != set(previo):
        errores.append(f"El staging recalculado no reconcilia con REVISAR 2B.2: "
                       f"{len(set(previo) - set(act))} ausentes, {len(set(act) - set(previo))} nuevos")

    filas_out = []
    for sig, p in previo.items():
        r = act.get(sig)
        if r is None:
            continue
        ot = st._a_ot(r["IDCampaña"])
        eid = st.texto(r["ElementoID"])
        ini, fin = st._a_fecha(r["FechaInicio"]), st._a_fecha(r["FechaFin"])
        ventana = set().union(*(res["ventanas"][h] for h in st.texto(r["TRZ_HojaOrigen"]).split(" + ")))
        occ = ocurrencias_del_registro(r, occ_fila)
        filas = idx_c.por_ot_elem.get((ot, eid), []) if ot is not None and eid else []
        out = {"StagingID_2B2": p["StagingID"], "Motivo_2B2": st.texto(p["Motivo"]), "MotivoPrincipal_2B2": motivo_principal_2b2(p["Motivo"]),
               "Motivo_recalculado": st.texto(r["Motivo"]), "Resolucion_2C": "REVISAR", "Regla_2C": "", "Operacion_2C": "",
               "Evidencia_2C": "", "Motivo_no_aplica": "", "CargaID_Extendida": "", "FechaFin_Anterior": None,
               "_reg": r, "_occ": occ, "_filas": filas}
        if r["Estado_Staging"] == "IMPORTAR":
            # Regla 2B confirmada, ahora aplicable porque el otro periodo ya esta cargado.
            out.update(Resolucion_2C="IMPORTAR", Regla_2C=R_DISJUNTO_2B, Operacion_2C=st.OP_INSERT,
                       Evidencia_2C=f"Regla 2B recalculada contra la base actual; otras filas de la OT en el elemento: {_fmt_filas(filas)}")
        elif r["TRZ_TipoGrupo"] == "SIN_OT" and eid:
            occ_filas = [o for key in _filas_de(r) for o in occ_fila.get(key, [])
                         if o.codigo in set(st.texto(r["TRZ_ElementoOrigen"]).split(" + "))]
            por_ot = {o.token.numero: idx_c.por_ot_elem.get((o.token.numero, eid), []) for o in occ_filas if o.token}
            ot_x, ev = evaluar_sin_ot_cubierta(occ, occ_filas, por_ot)
            if ot_x is not None:
                out.update(Resolucion_2C="CERRADO_SIN_CAMBIOS", Regla_2C=R_SIN_OT, Evidencia_2C=ev)
            else:
                out["Motivo_no_aplica"] = ev
        elif ot is not None and eid:
            fila, ev = evaluar_extension_demostrada(r["Motivo"], ini, fin, occ, filas, ventana)
            if fila is not None:
                out.update(Resolucion_2C="IMPORTAR", Regla_2C=R_EXTENSION, Operacion_2C=st.OP_UPDATE, Evidencia_2C=ev,
                           CargaID_Extendida=fila.carga_id, FechaFin_Anterior=fila.fin, _fila=fila)
            else:
                ok, ev2 = evaluar_periodo_disjunto(r["Motivo"], ini, fin, occ, filas, ventana)
                if ok:
                    out.update(Resolucion_2C="IMPORTAR", Regla_2C=R_DISJUNTO, Operacion_2C=st.OP_INSERT, Evidencia_2C=ev2)
                else:
                    out["Motivo_no_aplica"] = f"{ev} / {ev2}"
        filas_out.append(out)

    # 2) Coherencia con la misma OT en otros elementos y conflictos en estaticos.
    ops = []
    for k, o in enumerate(filas_out):
        if o["Resolucion_2C"] != "IMPORTAR":
            continue
        r = o["_reg"]
        desde = o["FechaFin_Anterior"] + UN_DIA if o["Operacion_2C"] == st.OP_UPDATE else st._a_fecha(r["FechaInicio"])
        ops.append((st.texto(r["ElementoID"]), st._a_ot(r["IDCampaña"]), desde, st._a_fecha(r["FechaFin"]), k))
    lote = defaultdict(set)
    for eid, ot, *_ in ops:
        lote[ot].add(eid)
    for eid, ot, desde, hasta, k in ops:
        otros = [f for (o2, e2), fs in idx_c.por_ot_elem.items() if o2 == ot and e2 not in lote[ot] for f in fs]
        falta = contradice_misma_ot(desde, hasta, otros)
        if falta:
            filas_out[k].update(Resolucion_2C="REVISAR", Regla_2C="", Operacion_2C="", Motivo_no_aplica=(
                f"CONTRADICE_LA_MISMA_OT_EN_OTROS_ELEMENTOS: {len({e for (o2, e) in idx_c.por_ot_elem if o2 == ot} - lote[ot])} elementos "
                f"no la tienen en {'; '.join(f'{a}→{b}' for a, b in falta)} ({filas_out[k]['Regla_2C']} {desde}→{hasta})"))
        elif otros:
            filas_out[k]["Evidencia_2C"] += f"; corroborado por la OT en otros elementos ({desde}→{hasta} cubierto)"
    ops = [op for op in ops if filas_out[op[4]]["Resolucion_2C"] == "IMPORTAR"]
    for k in conflictos_estaticos(ops, idx_c, medio_por_id):
        filas_out[k].update(Resolucion_2C="REVISAR", Regla_2C="", Operacion_2C="",
                            Motivo_no_aplica="SOLAPA_OTRA_CAMPANA_EN_ELEMENTO_ESTATICO (2C)")
    usos = Counter(o["CargaID_Extendida"] for o in filas_out if o["Operacion_2C"] == st.OP_UPDATE)
    for o in filas_out:
        if o["Operacion_2C"] == st.OP_UPDATE and usos[o["CargaID_Extendida"]] > 1:
            o.update(Resolucion_2C="REVISAR", Regla_2C="", Operacion_2C="", Motivo_no_aplica="VARIOS_PERIODOS_EXTIENDEN_LA_MISMA_FILA")

    # 3) Pendientes: diagnostico, responsable y (decision de negocio) presencia ya cubierta.
    for o in filas_out:
        if o["Resolucion_2C"] != "REVISAR":
            continue
        r, occ, filas = o["_reg"], o["_occ"], o["_filas"]
        ventana = set().union(*(res["ventanas"][h] for h in st.texto(r["TRZ_HojaOrigen"]).split(" + ")))
        cub = st.texto(r["ElementoID"]) != "" and r["TRZ_TipoGrupo"] == "OT" and presencia_cubierta(occ, filas)
        m2c = diagnostico_pendiente(r, occ, filas, ventana, cub)
        if o["Motivo_no_aplica"].startswith("CONTRADICE_LA_MISMA_OT"):
            m2c = "FECHAS_CONTRADICEN_LA_MISMA_OT"
        ini, fin = st._a_fecha(r["FechaInicio"]), st._a_fecha(r["FechaFin"])
        ot = st._a_ot(r["IDCampaña"])
        if ini and fin and ot is not None:
            iguales = [f.carga_id for (o2, e2), fs in idx_c.por_ot_elem.items() if o2 == ot and e2 != st.texto(r["ElementoID"])
                       for f in fs if (f.ini, f.fin) == (ini, fin)]
            if iguales:
                o["Nota_misma_ot"] = f"Mismas fechas que la OT en otros elementos de CAMPANAS ({len(iguales)} filas, p.ej. {iguales[0]})"
        m_solapa = re.search(r"SOLAPA_OTRA_CAMPANA_EN_ELEMENTO_ESTATICO:([\d,]+)", st.texto(r["Motivo"]))
        if m_solapa:
            eid = st.texto(r["ElementoID"])
            o["Nota_solapa"] = "Otra OT cargada en el elemento: " + "; ".join(
                f"{x} {f.carga_id} {f.ini}→{f.fin} '{campana_por_carga.get(f.carga_id, '')}'"
                for x in (int(v) for v in m_solapa.group(1).split(",")) for f in idx_c.por_ot_elem.get((x, eid), []))
        o["MotivoRevision"] = m2c
        o["Responsable"] = responsable(m2c)
        if m2c == "PRESENCIA_YA_CUBIERTA_EN_CAMPANAS" and cerrar_cubiertas:
            o.update(Resolucion_2C="CERRADO_SIN_CAMBIOS", Regla_2C=R_CUBIERTA,
                     Evidencia_2C=f"CAMPANAS {_fmt_filas(filas)} cubre {st.fmt_meses(sorted({x.mes for x in occ}))} (todos los meses con presencia)")
    return {"filas": filas_out, "errores": errores, "idx_c": idx_c, "medio_por_id": medio_por_id}


# ---------------------------------------------------------------------------
# Operaciones sobre CAMPANAS
# ---------------------------------------------------------------------------

def construir_operaciones(rev: dict[str, Any], res: dict[str, Any], fecha_carga: dt.datetime) -> tuple[pd.DataFrame, pd.DataFrame]:
    """-> (cambios UPDATE en formato del escritor 2B, filas INSERT con columnas CAMPANAS + TRZ)."""
    camp = res["campanas"]
    ins_regs, upd = [], []
    for o in rev["filas"]:
        if o["Resolucion_2C"] != "IMPORTAR":
            continue
        r = dict(o["_reg"])
        if o["Operacion_2C"] == st.OP_INSERT:
            r.update(Estado_Staging="IMPORTAR", Operacion=st.OP_INSERT, Motivo=o["Regla_2C"])
            ins_regs.append(r)
        else:
            f = o["_fila"]
            upd.append({"PosicionFila": f.pos, "CargaID": f.carga_id, "FechaFin_Anterior": f.fin,
                        "FechaFin_Nueva": st._a_fecha(r["FechaFin"]), "ClaveNegocio_Anterior": f.clave,
                        "ClaveNegocio_Nueva": st.clave_con_fin(f.clave, st._a_fecha(r["FechaFin"]))})
    upd_df = pd.DataFrame(upd, columns=["PosicionFila", "CargaID", "FechaFin_Anterior", "FechaFin_Nueva",
                                        "ClaveNegocio_Anterior", "ClaveNegocio_Nueva"])
    cambios = st.cambios_desde_extensiones(upd_df, camp, fecha_carga) if len(upd_df) else \
        pd.DataFrame(columns=st.CAMBIOS_COLUMNAS)
    base_cols = list(res["asignaciones"].columns)
    ins_df = pd.DataFrame(ins_regs, columns=base_cols)
    nuevas = st.preparar_importar(ins_df, rev["medio_por_id"], fecha_carga, st.siguiente_carga_id(camp)) if len(ins_df) else \
        st.preparar_importar(res["asignaciones"].iloc[0:0], rev["medio_por_id"], fecha_carga, 0)
    return cambios, nuevas


def validar_operaciones_2c(res: dict[str, Any], cambios: pd.DataFrame, nuevas: pd.DataFrame) -> list[str]:
    """Cruces en memoria antes de escribir. Cualquier error aborta."""
    camp, maestro = res["campanas"], res["maestro"]
    errores = []
    circ = dict(zip(maestro["ElementoID"].map(st.texto), maestro["CircuitoDashboard"].map(st.texto)))
    for r in nuevas.itertuples(index=False):
        if re.search("LONDON|CENCOMEDIA|YPF", st.texto(r.TRZ_AlcanceOrigen)) or circ.get(r.ElementoID) in st.CIRCUITOS_FUERA_DE_ALCANCE:
            errores.append(f"INSERT {r.CargaID}: alcance prohibido")
        if st.texto(r.ElementoID).upper().startswith(("PALS", "REM-TS")):
            errores.append(f"INSERT {r.CargaID}: ElementoID prohibido {r.ElementoID}")
    if set(nuevas["CargaID"]) & set(camp["CargaID"].map(st.texto)):
        errores.append("CargaID nuevos ya existentes")
    if set(cambios["Columna"]) - {"FechaFin", "ClaveNegocio", "Estado"}:
        errores.append("UPDATE sobre columnas no permitidas")
    esperado = st.campanas_esperadas(camp, cambios, nuevas)
    tocadas = set(cambios["CargaID"]) | set(nuevas["CargaID"])
    errores += st.validar_estado_final(camp, esperado, maestro, res["parametros"], tocadas)
    return errores


def congelar_base(base: Path, salida: Path) -> dict[str, Any]:
    """Copia congelada POST-2B.2 fuera del repo (idempotente: si existe, debe coincidir)."""
    destino = salida / NOMBRE_COPIA_CONGELADA
    if destino.exists():
        if st.sha256_file(destino) != SHA_BASE_POST_2B2:
            raise st.StagingError(f"La copia congelada {destino} no tiene el SHA POST-2B.2")
    else:
        shutil.copy2(base, destino)
    c = pd.read_excel(destino, sheet_name="CAMPANAS")
    man = {"archivo": destino.name, "origen": str(base), "sha256": st.sha256_file(destino),
           "filas_campanas": len(c), "ot_distintas": int(c["IDCampaña"].map(st._a_ot).dropna().nunique()),
           "fecha_hora_carga_max": str(c["FechaHoraCarga"].max()), "congelado": dt.datetime.now().replace(microsecond=0).isoformat()}
    if (man["sha256"], man["filas_campanas"], man["ot_distintas"]) != (SHA_BASE_POST_2B2, FILAS_POST_2B2, OT_POST_2B2):
        raise st.StagingError(f"Copia congelada inconsistente: {man}")
    man_path = salida / "MANIFEST_POST_2B2_PRE_2C.json"
    if not man_path.exists():
        man_path.write_text(json.dumps(man, ensure_ascii=False, indent=2), encoding="utf-8")
    return man


def controles_post_2c(base: Path, res: dict[str, Any], nuevas: pd.DataFrame, cambios: pd.DataFrame) -> list[tuple[str, str]]:
    """Validaciones obligatorias sobre la base ya escrita."""
    c = pd.read_excel(base, sheet_name="CAMPANAS")
    m = pd.read_excel(base, sheet_name="MAESTRO_ELEMENTOS")
    p = pd.read_excel(base, sheet_name="PARAMETROS")
    camp0 = res["campanas"]
    ok = lambda b: "OK" if b else "FALLA"  # noqa: E731
    m_ids = set(m["ElementoID"].map(st.texto))
    circ = dict(zip(m["ElementoID"].map(st.texto), m["CircuitoDashboard"].map(st.texto)))
    nuevas_b = c[c["CargaID"].isin(nuevas["CargaID"])]
    dup0 = int(camp0["ClaveNegocio"].map(st.texto).duplicated().sum())
    dup1 = int(c["ClaveNegocio"].map(st.texto).duplicated().sum())
    dups_nuevas = set(c.loc[c["ClaveNegocio"].map(st.texto).duplicated(keep=False), "CargaID"]) & (set(nuevas["CargaID"]) | set(cambios["CargaID"]))
    return [
        ("MAESTRO_ELEMENTOS identico", ok(st._norm_df(m).equals(st._norm_df(res["maestro"])))),
        ("PARAMETROS identico", ok(st._norm_df(p).equals(st._norm_df(res["parametros"])))),
        ("0 ElementoID huerfanos", ok(set(c["ElementoID"].map(st.texto)) <= m_ids)),
        ("0 CargaID duplicados", ok(not c["CargaID"].duplicated().any())),
        (f"Ninguna ClaveNegocio duplicada nueva (antes {dup0}, despues {dup1})", ok(dup1 <= dup0 and not dups_nuevas)),
        (f"Filas = antes {len(camp0)} + INSERT {len(nuevas)}", ok(len(c) == len(camp0) + len(nuevas))),
        ("IDCampaña nuevas numericas", ok(nuevas_b["IDCampaña"].map(lambda v: st._a_ot(v) is not None).all())),
        ("Ninguna fila nueva YPF / London / Cencomedia", ok(not nuevas_b["ElementoID"].map(lambda e: circ.get(st.texto(e), "")).isin(
            st.CIRCUITOS_FUERA_DE_ALCANCE).any() and not nuevas["TRZ_AlcanceOrigen"].astype(str).str.contains("LONDON|CENCOMEDIA|YPF").any())),
        ("Ninguna fila PALS", ok(not c["ElementoID"].map(st.texto).str.upper().str.startswith("PALS").any())),
        ("Ningun ElementoID REM-TS", ok(not c["ElementoID"].map(st.texto).str.upper().str.startswith("REM-TS").any())),
        ("Filas extendidas con la FechaFin nueva", ok(all(
            st._a_fecha(c.loc[c["CargaID"] == ch.CargaID, "FechaFin"].iloc[0]) == st._a_fecha(ch.Nuevo)
            for ch in cambios[cambios["Columna"] == "FechaFin"].itertuples()))),
    ]


# ---------------------------------------------------------------------------
# Salidas
# ---------------------------------------------------------------------------

COLUMNAS_PENDIENTES = [
    "HojaOrigen", "Código histórico", "ElementoID candidato", "IDCampaña / OT", "Campaña / Pauta",
    "FechaInicio fuente", "FechaFin fuente", "FechaInicio actualmente cargada", "FechaFin actualmente cargada",
    "MotivoRevision", "Evidencia", "Candidatos", "Acción sugerida", "Observación", "Responsable",
    "Motivo principal 2B.2", "Celdas fuente", "StagingID 2B.2",
]


def tabla_revision(rev: dict[str, Any]) -> pd.DataFrame:
    """Una fila por asignacion REVISAR 2B.2 con su resolucion 2C (trazabilidad completa)."""
    out = []
    for o in rev["filas"]:
        r, occ, filas = o["_reg"], o["_occ"], o["_filas"]
        conf = st.texto(r["TRZ_ConfianzaMatch"])
        cand = st.texto(r["TRZ_CandidatosElemento"])
        m2c = o.get("MotivoRevision", "")
        no_aplica = " / ".join(x for x in o["Motivo_no_aplica"].split(" / ") if x and x != "NO_APLICA_MOTIVO")
        evid = o["Evidencia_2C"] if o["Resolucion_2C"] != "REVISAR" else "; ".join(x for x in (
            f"OT en celda: {st.texto(r['TRZ_OT_Raw'])}" if st.texto(r["TRZ_OT_Raw"]) and (
                st.texto(r["TRZ_CeldaCompartida"]) == "SI" or r["TRZ_TipoGrupo"] != "OT") else "",
            f"Grilla: {_fmt_celdas(occ)}" if occ else "",
            f"CAMPANAS misma OT+elemento: {_fmt_filas(filas)}" if filas else "",
            f"Presencia: {st.texto(r['TRZ_MesesPresencia'])}",
            f"Problemas: {st.texto(r['TRZ_ProblemasFechas'])}" if st.texto(r["TRZ_ProblemasFechas"]) else "",
            o.get("Nota_solapa", ""),
            f"Regla 2C no aplicable: {no_aplica}" if no_aplica else "",
            o.get("Nota_misma_ot", ""),
        ) if x)
        out.append({
            "StagingID 2B.2": o["StagingID_2B2"], "Resolucion_2C": o["Resolucion_2C"], "Regla_2C": o["Regla_2C"],
            "Operacion_2C": o["Operacion_2C"], "CargaID extendida": o["CargaID_Extendida"],
            "HojaOrigen": st.texto(r["TRZ_HojaOrigen"]), "Código histórico": st.texto(r["TRZ_ElementoOrigen"]) or "(fila sin código)",
            "ElementoID candidato": st.texto(r["ElementoID"]) or (cand if conf == "MEDIA" else ""),
            "IDCampaña / OT": st._a_ot(r["IDCampaña"]) if st._a_ot(r["IDCampaña"]) is not None else (st.texto(r["TRZ_OT_Raw"]) or "(sin OT)"),
            "Campaña / Pauta": st.texto(r["TRZ_Pauta_Origen"]),
            "FechaInicio fuente": st._a_fecha(r["FechaInicio"]), "FechaFin fuente": st._a_fecha(r["FechaFin"]),
            "FechaInicio actualmente cargada": "; ".join(st.fmt_fecha(f.ini) for f in sorted(filas, key=lambda f: f.ini or dt.date.min)),
            "FechaFin actualmente cargada": "; ".join(st.fmt_fecha(f.fin) for f in sorted(filas, key=lambda f: f.ini or dt.date.min)),
            "MotivoRevision": m2c, "Evidencia": evid[:1500],
            "Candidatos": cand if conf != "ALTA" else "",
            "Acción sugerida": ACCION_SUGERIDA.get(m2c, ""),
            "Observación": " | ".join(x for x in (f"Motivo 2B.2: {o['Motivo_2B2']}", st.texto(r["TRZ_Obs_Origen"]),
                                                  "OT bonificada (B)" if st.texto(r["TRZ_OT_Bonificada"]) == "SI" else "") if x),
            "Responsable": o.get("Responsable", "") if o["Resolucion_2C"] == "REVISAR" else "",
            "Motivo principal 2B.2": o["MotivoPrincipal_2B2"], "Celdas fuente": st.texto(r["TRZ_RefCeldas"])[:500],
        })
    return pd.DataFrame(out)


def _escribir_xlsx(path: Path, hojas: list[tuple[str, pd.DataFrame]]) -> None:
    import openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for nombre, df in hojas:
        ws = wb.create_sheet(nombre)
        st._escribir_df(ws, df, 1)
        ws.freeze_panes = "A2"
        if len(df.columns):
            ws.auto_filter.ref = f"A1:{st.col_letra(len(df.columns) - 1)}{len(df) + 1}"
        for j, col in enumerate(df.columns):
            ws.column_dimensions[st.col_letra(j)].width = min(60, max(12, len(str(col)) + 2))
    wb.save(path)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--fuente", required=True)
    ap.add_argument("--staging-2b2", required=True, help="OCU26_HISTORICO_STAGING.xlsx del cierre 2B.2 (lectura)")
    ap.add_argument("--salida", required=True)
    ap.add_argument("--base", default=None)
    ap.add_argument("--sha-base", default=SHA_BASE_POST_2B2)
    ap.add_argument("--cerrar-presencia-cubierta", action="store_true",
                    help="Decision de negocio: cerrar sin cambios lo que CAMPANAS ya cubre en todos los meses con presencia")
    ap.add_argument("--aplicar", action="store_true")
    args = ap.parse_args(argv)

    fuente, salida, stg = Path(args.fuente), Path(args.salida), Path(args.staging_2b2)
    base = resolve_input_path(args.base)
    salida.mkdir(parents=True, exist_ok=True)
    sha_fuente, sha_base, sha_stg = st.sha256_file(fuente), st.sha256_file(base), st.sha256_file(stg)
    print(f"BASE {base} sha256={sha_base}")
    if sha_base != args.sha_base:
        print(f"ERROR: SHA de la base {sha_base} != esperado {args.sha_base}. Detenido.", file=sys.stderr)
        return 2
    canonica = sha_base == SHA_BASE_POST_2B2 and base.resolve() == resolve_input_path(None).resolve()

    fecha_carga = dt.datetime.now().replace(microsecond=0)
    previo = pd.read_excel(stg, sheet_name="REVISAR")
    res = st.construir_staging(fuente, base, fecha_carga=fecha_carga)
    camp = res["campanas"]
    antes = {"filas": len(camp), "ot_distintas": int(camp["IDCampaña"].map(st._a_ot).dropna().nunique()), "sha256": sha_base,
             "fecha_hora_carga_max": str(camp["FechaHoraCarga"].max())}
    print("BASE ANTES:", antes)
    errores = []
    if len(previo) != REVISAR_2B2:
        errores.append(f"REVISAR 2B.2 = {len(previo)} (se esperaban {REVISAR_2B2})")

    rev = revisar(res, previo, args.cerrar_presencia_cubierta)
    errores += rev["errores"]
    cambios, nuevas = construir_operaciones(rev, res, fecha_carga)
    errores += validar_operaciones_2c(res, cambios, nuevas)

    tabla = tabla_revision(rev)
    pend = tabla[tabla["Resolucion_2C"] == "REVISAR"]
    print(tabla.groupby(["Resolucion_2C", "Regla_2C"]).size().to_string())
    print(f"Operaciones: INSERT {len(nuevas)}, UPDATE filas {cambios['CargaID'].nunique()} ({len(cambios)} celdas)")
    print(f"Pendientes: {len(pend)} asignaciones, {pend['IDCampaña / OT'].map(st._a_ot).dropna().nunique()} OT; "
          f"{pend['Responsable'].value_counts().to_dict()}")
    print("Errores previos:", errores or "ninguno")

    despues, ctrl, controles, man = None, None, [], None
    if not errores and args.aplicar:
        if canonica:
            man = congelar_base(base, salida)
            print("Copia congelada:", man)
        try:
            ctrl = st.aplicar_operaciones(base, camp, cambios, nuevas, "etapa2c")
        except st.StagingError as exc:
            print(f"ERROR: {exc}", file=sys.stderr)
            return 4
        final = pd.read_excel(base, sheet_name="CAMPANAS")
        despues = {"filas": len(final), "ot_distintas": int(final["IDCampaña"].map(st._a_ot).dropna().nunique()),
                   "sha256": st.sha256_file(base), "fecha_hora_carga_max": str(final["FechaHoraCarga"].max())}
        controles = controles_post_2c(base, res, nuevas, cambios)
        controles.append(("validate_input", f"{ctrl['validate_input']} ({len(ctrl['validate_errors'])} errores)"))
        controles.append(("tblCampanas cubre todas las filas", "OK" if ctrl["tabla_cubre_filas"] else "FALLA"))
        print("BASE DESPUES:", despues)
        for k, v in controles:
            print(f"  [{v}] {k}")

    if errores:
        print("ABORTADO, base no modificada", file=sys.stderr)
        return 4

    # Salidas (siempre, tambien en dry-run).
    ops = pd.concat([
        nuevas[["CargaID", "IDCampaña", "ElementoID", "FechaInicio", "FechaFin", "ClaveNegocio", "Estado", "Campaña",
                "Observaciones", "TRZ_HojaOrigen", "TRZ_ElementoOrigen", "TRZ_RefCeldas"]].assign(Operacion="INSERT"),
        cambios.assign(Operacion="UPDATE"),
    ], ignore_index=True)
    regla_por_ref = {(st.texto(o["_reg"]["TRZ_RefCeldas"]), st._a_ot(o["_reg"]["IDCampaña"]), st.texto(o["_reg"]["ElementoID"])): o["Regla_2C"]
                     for o in rev["filas"] if o["Resolucion_2C"] == "IMPORTAR"}
    ops.insert(1, "Regla_2C", [regla_por_ref.get((st.texto(rf), st._a_ot(ot), st.texto(e)), "") if op == "INSERT" else R_EXTENSION
                               for rf, ot, e, op in zip(ops["TRZ_RefCeldas"], ops["IDCampaña"], ops["ElementoID"], ops["Operacion"])])
    nombre_ops = "OCU26_OPERACIONES_APLICADAS_2C.csv" if ctrl else "OCU26_OPERACIONES_PROPUESTAS_2C.csv"
    ops.to_csv(salida / nombre_ops, index=False, encoding="utf-8-sig")
    resumen_motivos = (pend.groupby(["MotivoRevision", "Responsable"]).agg(Asignaciones=("MotivoRevision", "size"),
                       OT=("IDCampaña / OT", lambda s: s.map(st._a_ot).dropna().nunique())).reset_index()
                       .sort_values("Asignaciones", ascending=False))
    _escribir_xlsx(salida / "OCU26_PENDIENTES_PRE_SHAREPOINT.xlsx",
                   [("PENDIENTES", pend[COLUMNAS_PENDIENTES].sort_values(["Responsable", "MotivoRevision", "HojaOrigen", "IDCampaña / OT"],
                                                                         key=lambda s: s.astype(str))),
                    ("RESUMEN", resumen_motivos)])
    _escribir_xlsx(salida / "OCU26_REVISION_2C.xlsx",
                   [("REVISION_458", tabla.drop(columns=["Responsable"]).assign(Responsable=tabla["Responsable"])),
                    ("OPERACIONES", ops)])

    def md_tabla(df: pd.DataFrame) -> str:
        return st._md_tabla(df)

    resol = tabla.groupby(["Resolucion_2C", "Regla_2C"]).size().reset_index(name="Asignaciones")
    l = ["# Etapa 2C - Revision final del historico (pre-SharePoint)\n",
         md_tabla(pd.DataFrame([("Fuente", f"{fuente} {sha_fuente}"), ("Staging 2B.2", f"{stg} {sha_stg}"),
                                ("Base", f"{base} {sha_base}"), ("Fecha de carga 2C", f"{fecha_carga:%Y-%m-%d %H:%M:%S}"),
                                ("Cerrar presencia ya cubierta (decision)", "SI" if args.cerrar_presencia_cubierta else "NO"),
                                ("Aplicado", "SI" if ctrl else "NO (dry-run)")], columns=["Clave", "Valor"])),
         "\n## Base antes / despues\n", md_tabla(pd.DataFrame([antes] + ([despues] if despues else []))),
         "\n## Resolucion de los REVISAR 2B.2\n", md_tabla(resol),
         "\n### Por motivo principal 2B.2\n",
         md_tabla(tabla.pivot_table(index="Motivo principal 2B.2", columns="Resolucion_2C", values="StagingID 2B.2",
                                    aggfunc="count", fill_value=0).reset_index()),
         f"\n## Operaciones ({len(nuevas)} INSERT, {cambios['CargaID'].nunique()} UPDATE)\n",
         md_tabla(ops[["Operacion", "Regla_2C", "CargaID", "IDCampaña", "ElementoID", "FechaInicio", "FechaFin", "Columna", "Anterior", "Nuevo"]]
                  if "Columna" in ops else ops),
         f"\n## Pendientes ({len(pend)} asignaciones, {pend['IDCampaña / OT'].map(st._a_ot).dropna().nunique()} OT)\n",
         md_tabla(resumen_motivos),
         "\n" + md_tabla(pend["Responsable"].value_counts().rename_axis("Responsable").reset_index(name="Asignaciones"))]
    if controles:
        l += ["\n## Controles\n", md_tabla(pd.DataFrame(controles, columns=["Control", "Resultado"])),
              f"\n- Partes modificadas del xlsx: {', '.join(ctrl['partes_modificadas'])}\n- tblCampanas: {ctrl['tabla_ref']}\n"]
    if man:
        l += ["\n## Copia congelada POST-2B.2\n", md_tabla(pd.DataFrame(list(man.items()), columns=["Clave", "Valor"]))]
    (salida / "RESUMEN_ETAPA2C.md").write_text("".join(x if x.endswith("\n") else x + "\n" for x in l), encoding="utf-8")
    print(f"Escritos en {salida}")
    if st.sha256_file(fuente) != sha_fuente or st.sha256_file(stg) != sha_stg:
        print("ERROR: la fuente o el staging 2B.2 cambiaron", file=sys.stderr)
        return 3
    return 5 if any(v.startswith("FALLA") for _, v in controles) else 0


if __name__ == "__main__":
    sys.exit(main())
