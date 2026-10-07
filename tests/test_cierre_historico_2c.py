"""Tests de la revision final del historico (Etapa 2C).

Hermeticos: celdas sinteticas (Ocurrencia) y filas de CAMPANAS en memoria; no
leen ni escriben la base, la fuente ni output/.
"""

from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT / "scripts"))

import cierre_historico_2c as c2  # noqa: E402
import staging_import_historico as st  # noqa: E402

D = dt.date
VENTANA = {st.sumar_meses(D(2025, 10, 1), k) for k in range(18)}  # 2025-10 .. 2027-03
DUDAS_PARCIAL = "FECHAS_DUDOSAS | EXISTE_EN_CAMPANAS_COBERTURA_PARCIAL"
OTRAS_FECHAS = "FECHAS_DUDOSAS | OT_ELEMENTO_EXISTE_CON_OTRAS_FECHAS"


def _occ(mes, ini, fin, ot=4000, compartida=False, pauta="P", tipo_ini="FECHA", tipo_fin="FECHA", fila=3):
    return st.Ocurrencia(
        hoja="PLED", fila=fila, ref=f"PLED!{mes:%m}{fila}", mes=mes, codigo="X-V1", attrs={}, ot_raw=str(ot),
        token=None if ot is None else st.OTToken(ot), n_tokens=2 if compartida else 1, texto_no_ot="", pauta=pauta,
        ini=ini, fin=fin, tipo_ini=tipo_ini, tipo_fin=tipo_fin, obs="", alcance="EN_ALCANCE",
    )


def _fila(ini, fin, cid="HIST-1", ot=4000, eid="X"):
    return st.FilaCampana(carga_id=cid, ini=ini, fin=fin, indefinida=False, clave=st.clave_negocio(ot, eid, ini, fin))


# ---------------------------------------------------------------------------
# EXTENSION_FIN_DEMOSTRADA
# ---------------------------------------------------------------------------

def _serie_2826():
    """Bloques 2025 con otro INICIO (duda ya cubierta por la fila) y extension limpia en 2027."""
    occ = [_occ(D(2025, 11, 1), D(2025, 1, 26), D(2025, 12, 31)), _occ(D(2025, 12, 1), D(2025, 1, 26), D(2025, 12, 31))]
    occ += [_occ(st.sumar_meses(D(2026, 1, 1), k), D(2025, 1, 7), D(2026, 12, 31)) for k in range(8)]
    occ += [_occ(st.sumar_meses(D(2026, 9, 1), k), D(2025, 1, 7), D(2027, 12, 31)) for k in range(7)]
    return occ


def test_extension_demostrada_dudas_solo_en_meses_cubiertos():
    fila = _fila(D(2025, 1, 7), D(2026, 12, 31))
    f, evid = c2.evaluar_extension_demostrada(DUDAS_PARCIAL, D(2025, 1, 7), D(2027, 12, 31), _serie_2826(), [fila], VENTANA)
    assert f is fila
    assert "2027-01..2027-03" in evid


def test_extension_rechaza_fin_que_decrece():
    """4365: junio dice FIN 07-31 y julio 07-13 -> contradiccion."""
    occ = [_occ(D(2026, m, 1), D(2026, 3, 1), D(2026, 6, 30)) for m in (3, 4, 5)]
    occ += [_occ(D(2026, 6, 1), D(2026, 3, 1), D(2026, 7, 31)), _occ(D(2026, 7, 1), D(2026, 3, 1), D(2026, 7, 13))]
    f, motivo = c2.evaluar_extension_demostrada(DUDAS_PARCIAL, D(2026, 3, 1), D(2026, 7, 31), occ,
                                                [_fila(D(2026, 3, 1), D(2026, 6, 30))], VENTANA)
    assert f is None and motivo == "FIN_DECRECE_ENTRE_BLOQUES"


def test_extension_rechaza_inicio_distinto_en_meses_nuevos():
    """2578: los bloques de la extension traen otro INICIO (renovacion solapada)."""
    occ = [_occ(D(2026, m, 1), D(2025, 6, 4), D(2026, 4, 30)) for m in (1, 2, 3, 4)]
    occ += [_occ(D(2026, m, 1), D(2026, 4, 5), D(2026, 7, 31)) for m in (5, 6, 7)]
    f, motivo = c2.evaluar_extension_demostrada(DUDAS_PARCIAL, D(2025, 6, 4), D(2026, 7, 31), occ,
                                                [_fila(D(2025, 6, 4), D(2026, 4, 30))], VENTANA)
    assert f is None and motivo in ("ULTIMO_FIN_DISTINTO_DEL_PERIODO", "CELDA_INCONSISTENTE_EN_EXTENSION:PLED!053")


def test_extension_rechaza_mes_nuevo_sin_presencia_o_compartido():
    base = [_occ(D(2026, 5, 1), D(2026, 5, 4), D(2026, 6, 30)), _occ(D(2026, 6, 1), D(2026, 5, 4), D(2026, 8, 31))]
    fila = [_fila(D(2026, 5, 4), D(2026, 6, 30))]
    sin_agosto = base + [_occ(D(2026, 7, 1), D(2026, 5, 4), D(2026, 8, 31))]
    f, motivo = c2.evaluar_extension_demostrada(DUDAS_PARCIAL, D(2026, 5, 4), D(2026, 8, 31), sin_agosto, fila, VENTANA)
    assert f is None and motivo == "MES_SIN_PRESENCIA_EN_EXTENSION:2026-08"
    compartida = sin_agosto + [_occ(D(2026, 8, 1), D(2026, 5, 4), D(2026, 8, 31), compartida=True)]
    f, motivo = c2.evaluar_extension_demostrada(DUDAS_PARCIAL, D(2026, 5, 4), D(2026, 8, 31), compartida, fila, VENTANA)
    assert f is None and motivo == "CELDA_COMPARTIDA_EN_EXTENSION:2026-08"
    completa = sin_agosto + [_occ(D(2026, 8, 1), D(2026, 5, 4), D(2026, 8, 31))]
    f, _ = c2.evaluar_extension_demostrada(DUDAS_PARCIAL, D(2026, 5, 4), D(2026, 8, 31), completa, fila, VENTANA)
    assert f is not None


def test_extension_no_aplica_a_otros_motivos_ni_cambia_inicio():
    occ = [_occ(D(2026, m, 1), D(2026, 3, 1), D(2026, 5, 31)) for m in (3, 4, 5)]
    f, motivo = c2.evaluar_extension_demostrada("EXISTE_EN_CAMPANAS_COBERTURA_PARCIAL | EXTENSION_CAMBIA_FECHAINICIO",
                                                D(2026, 3, 1), D(2026, 5, 31), occ, [_fila(D(2026, 3, 18), D(2026, 5, 31))], VENTANA)
    assert f is None and motivo == "NO_APLICA_MOTIVO"
    f, motivo = c2.evaluar_extension_demostrada(DUDAS_PARCIAL, D(2026, 3, 1), D(2026, 6, 30), occ,
                                                [_fila(D(2026, 3, 18), D(2026, 5, 31))], VENTANA)
    assert f is None and motivo == "EXTENSION_CAMBIA_FECHAINICIO"


# ---------------------------------------------------------------------------
# PERIODO_DISJUNTO_EXPLICITO
# ---------------------------------------------------------------------------

def _serie_4533():
    """Abril anticipa el primer periodo (ya cargado); julio-agosto es un tramo nuevo explicito."""
    return [_occ(D(2026, 4, 1), D(2026, 5, 4), D(2026, 6, 2)), _occ(D(2026, 5, 1), D(2026, 5, 4), D(2026, 6, 2)),
            _occ(D(2026, 7, 1), D(2026, 7, 15), D(2026, 8, 12)), _occ(D(2026, 8, 1), D(2026, 7, 15), D(2026, 8, 12))]


def test_periodo_disjunto_explicito():
    ok, evid = c2.evaluar_periodo_disjunto(OTRAS_FECHAS, D(2026, 7, 15), D(2026, 8, 12), _serie_4533(),
                                           [_fila(D(2026, 5, 4), D(2026, 6, 2))], VENTANA)
    assert ok and "2026-07, 2026-08" in evid


def test_periodo_disjunto_rechaza_contacto_pares_distintos_y_meses_faltantes():
    filas = [_fila(D(2026, 5, 4), D(2026, 7, 14))]  # contiguo: seria una extension, no un tramo nuevo
    ok, motivo = c2.evaluar_periodo_disjunto(OTRAS_FECHAS, D(2026, 7, 15), D(2026, 8, 12), _serie_4533(), filas, VENTANA)
    assert not ok and motivo.startswith("TOCA_FILA_EXISTENTE")
    filas = [_fila(D(2026, 5, 4), D(2026, 6, 2))]
    occ = _serie_4533()[:3] + [_occ(D(2026, 8, 1), D(2026, 7, 15), D(2026, 8, 20))]
    ok, motivo = c2.evaluar_periodo_disjunto(OTRAS_FECHAS, D(2026, 7, 15), D(2026, 8, 20), occ, filas, VENTANA)
    assert not ok and motivo == "PARES_DISTINTOS_EN_EL_PERIODO"
    ok, motivo = c2.evaluar_periodo_disjunto(OTRAS_FECHAS, D(2026, 7, 15), D(2026, 8, 12), _serie_4533()[:3], filas, VENTANA)
    assert not ok and motivo == "MES_SIN_PRESENCIA:2026-08"
    ok, motivo = c2.evaluar_periodo_disjunto(OTRAS_FECHAS, D(2026, 7, 15), D(2026, 8, 12), _serie_4533(), [], VENTANA)
    assert not ok


# ---------------------------------------------------------------------------
# SIN_OT_CUBIERTA_POR_LA_MISMA_FILA
# ---------------------------------------------------------------------------

def test_sin_ot_cubierta_por_unica_ot_de_la_fila():
    sin_ot = [_occ(D(2026, 3, 1), D(2026, 2, 1), D(2026, 12, 31), ot=None, pauta="YPF PIER 3")]
    fila = sin_ot + [_occ(D(2026, m, 1), D(2026, 2, 1), D(2026, 12, 31), ot=4311, pauta="YPF PIER 3") for m in (2, 4, 5)]
    ot, evid = c2.evaluar_sin_ot_cubierta(sin_ot, fila, {4311: [_fila(D(2026, 2, 1), D(2026, 12, 31), ot=4311)]})
    assert ot == 4311 and "2026-03" in evid


def test_sin_ot_no_se_atribuye_si_la_fila_muestra_dos_ot_o_no_hay_cobertura():
    sin_ot = [_occ(D(2026, 2, 1), D(2024, 1, 1), D(2026, 12, 31), ot=None, pauta="SAMSUNG")]
    fila = sin_ot + [_occ(D(2026, 3, 1), D(2024, 1, 1), D(2026, 12, 31), ot=3862, pauta="SAMSUNG", compartida=True),
                     _occ(D(2026, 3, 1), D(2024, 1, 1), D(2026, 12, 31), ot=4383, pauta="SAMSUNG", compartida=True)]
    filas = {4383: [_fila(D(2026, 1, 28), D(2026, 12, 31), ot=4383)], 3862: [_fila(D(2024, 1, 1), D(2026, 1, 27), ot=3862)]}
    ot, motivo = c2.evaluar_sin_ot_cubierta(sin_ot, fila, filas)
    assert ot is None and motivo.startswith("OT_IDENTIFICABLE_NO_UNICA")
    fila = sin_ot + [_occ(D(2026, 4, 1), D(2024, 1, 1), D(2026, 12, 31), ot=4383, pauta="SAMSUNG")]
    ot, motivo = c2.evaluar_sin_ot_cubierta(sin_ot, fila, {4383: [_fila(D(2026, 3, 1), D(2026, 12, 31), ot=4383)]})
    assert ot is None and motivo == "OT_4383_NO_CUBRE_LOS_MESES"


# ---------------------------------------------------------------------------
# Coherencia con la misma OT, presencia cubierta, diagnostico
# ---------------------------------------------------------------------------

def test_contradice_misma_ot_en_otros_elementos():
    """4591: los demas elementos tienen un hueco 06-22..07-20 entre dos periodos."""
    otros = [_fila(D(2026, 5, 21), D(2026, 6, 21), "H1"), _fila(D(2026, 7, 21), D(2026, 8, 21), "H2")]
    assert c2.contradice_misma_ot(D(2026, 6, 22), D(2026, 7, 31), otros) == [(D(2026, 6, 22), D(2026, 7, 20))]
    assert c2.contradice_misma_ot(D(2026, 7, 21), D(2026, 8, 21), otros) == []
    assert c2.contradice_misma_ot(D(2026, 6, 22), D(2026, 7, 31), []) == []  # sin evidencia en contra


def test_presencia_cubierta_y_diagnostico():
    occ = [_occ(D(2026, m, 1), D(2024, 1, 1), D(2026, 12, 31)) for m in (3, 4, 5)]
    filas = [_fila(D(2026, 3, 1), D(2026, 12, 31))]
    assert c2.presencia_cubierta(occ, filas)
    assert not c2.presencia_cubierta(occ, [_fila(D(2026, 4, 1), D(2026, 12, 31))])
    reg = {"Motivo": DUDAS_PARCIAL, "TRZ_ProblemasFechas": "MESES_SIN_PRESENCIA:2025-10, 2026-01, 2026-02",
           "FechaInicio": D(2024, 1, 1), "FechaFin": D(2026, 12, 31)}
    assert c2.diagnostico_pendiente(reg, occ, filas, VENTANA, True) == "PRESENCIA_YA_CUBIERTA_EN_CAMPANAS"
    assert c2.diagnostico_pendiente(reg, occ, [], VENTANA, False) == "INICIO_ANTERIOR_A_LA_PRESENCIA"
    reg["TRZ_ProblemasFechas"] = "MESES_SIN_PRESENCIA:2026-06, 2026-07"
    assert c2.diagnostico_pendiente(reg, occ, [], VENTANA, False) == "FIN_NO_CONFIRMADO_POR_LA_GRILLA"
    for motivo, esperado in (("ELEMENTO_NO_INEQUIVOCO_BAJA | FECHAS_DUDOSAS", "ELEMENTO_SIN_EQUIVALENCIA_EN_MAESTRO"),
                             ("EXISTE_EN_CAMPANAS_COBERTURA_PARCIAL | EXTENSION_CAMBIA_FECHAINICIO", "EXTENSION_CAMBIARIA_FECHAINICIO")):
        assert c2.diagnostico_pendiente({**reg, "Motivo": motivo}, occ, filas, VENTANA, True) == esperado


def test_responsable_y_motivo_principal():
    assert c2.responsable("ELEMENTO_AMBIGUO") == "USUARIO"
    assert c2.responsable("FILA_SIN_CODIGO_DE_ELEMENTO") == "USUARIO"
    assert c2.responsable("FECHAS_SOLO_EN_CELDA_COMPARTIDA") == "EQUIPO"
    assert c2.motivo_principal_2b2("FECHAS_DUDOSAS | EXISTE_EN_CAMPANAS_COBERTURA_PARCIAL") == "Fechas dudosas"
    assert c2.motivo_principal_2b2("ELEMENTO_NO_INEQUIVOCO_BAJA | FECHAS_NO_DETERMINABLES") == "Elemento sin match (BAJA)"
    assert c2.motivo_principal_2b2("EXISTE_EN_CAMPANAS_COBERTURA_PARCIAL | EXTENSION_CAMBIA_FECHAINICIO") == \
        "Extensión que cambiaría FechaInicio"


def test_firma_estable_entre_staging_leido_y_recalculado():
    import pandas as pd
    recalculado = {"TRZ_HojaOrigen": "PLED", "TRZ_ElementoOrigen": "PCAB-3600seg-V9", "IDCampaña": 4803,
                   "TRZ_TipoGrupo": "OT", "TRZ_MesesPresencia": "2026-08", "TRZ_RefCeldas": "PLED!BD91",
                   "FechaInicio": D(2026, 8, 12), "FechaFin": D(2026, 9, 12)}
    leido = {**recalculado, "IDCampaña": 4803.0, "FechaInicio": pd.Timestamp("2026-08-12"), "FechaFin": pd.Timestamp("2026-09-12")}
    assert c2.firma(recalculado) == c2.firma(leido)
    assert c2.firma({**leido, "IDCampaña": float("nan"), "FechaInicio": pd.NaT})[2] is None
