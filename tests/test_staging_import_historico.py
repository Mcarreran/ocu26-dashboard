"""Tests del staging de migracion historica (Etapa 2B).

Hermeticos: filas sinteticas en memoria con la estructura de
OCUPACION_2026 (bloques mensuales OT | PAUTA | INICIO | FIN | OBS). No leen
ni escriben ningun Excel real, HTML productivo ni output/.
"""

from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

import pandas as pd
import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT / "scripts"))

import staging_import_historico as st  # noqa: E402
from validate_input import CAMPANAS_HEADERS  # noqa: E402

D = dt.date


# ---------------------------------------------------------------------------
# Helpers sinteticos
# ---------------------------------------------------------------------------

def _hoja(filas: list[list], meses=("NOV", "DIC", "ENE", "FEB"), inicio=D(2025, 11, 1), attrs=None, mes_faltante=None):
    """Arma filas tipo OCUPACION: fila de meses, encabezado y datos.

    filas: [codigo, clas, {mes_idx: (ot, pauta, ini, fin, obs)}]
    """
    attrs = attrs or ["Ciudad", "Shop", "Clas", "Ubic", "Disp", "Código", "Descripción"]
    n_attr = len(attrs)
    fila_meses = [None] * n_attr
    header = list(attrs)
    for k, et in enumerate(meses):
        m = st.sumar_meses(inicio, k)
        fila_meses += [None if k == mes_faltante else dt.datetime(m.year, m.month, 1), None, None, None, None]
        header += [et, "PAUTA", "INICIO", "FIN", "OBS"]
    rows = [tuple(fila_meses), tuple(header)]
    for codigo, clas, ubic, desc, bloques in filas:
        r = ["CIUDAD", "INDOOR", clas, ubic, "TD", codigo, desc] + [None] * (5 * len(meses))
        for k, vals in bloques.items():
            for j, v in enumerate(vals):
                r[n_attr + 5 * k + j] = v
        rows.append(tuple(r))
    return rows


def _occ(mes, ini, fin, n_tokens=1, tipo_ini="FECHA", tipo_fin="FECHA", ot=4000):
    return st.Ocurrencia(
        hoja="CENCO D", fila=3, ref="CENCO D!K3", mes=mes, codigo="X", attrs={}, ot_raw=str(ot),
        token=st.OTToken(ot), n_tokens=n_tokens, texto_no_ot="", pauta="P", ini=ini, fin=fin,
        tipo_ini=tipo_ini if ini or tipo_ini != "FECHA" else "VACIA",
        tipo_fin=tipo_fin if fin or tipo_fin != "FECHA" else "VACIA", obs="", alcance="EN_ALCANCE",
    )


VENTANA = [st.sumar_meses(D(2025, 10, 1), k) for k in range(18)]  # 2025-10 .. 2027-03


def _maestro(rows):
    cols = ["ElementoID", "CircuitoDashboard", "Subcircuito", "Ubicacion", "Descripcion", "Medio"]
    return pd.DataFrame(rows, columns=cols)


MAESTRO = _maestro([
    ["UNI-TRIEDRO-1A-2", "Shoppings Digital", "CENCOSUD", "UNICENTER", "Triedro Digital - Octogono Chico", "Digital"],
    ["UNI-TOTEMD-1G-5", "Shoppings Digital", "CENCOSUD", "UNICENTER", "Totem Digital Garganta Reloj - Nivel 1", "Digital"],
    ["UNI-TOTEMD-2G-4", "Shoppings Digital", "CENCOSUD", "UNICENTER", "Totem Digital Garganta Reloj - Nivel 2", "Digital"],
    ["UNI-LONAER3x5-K1", "Shoppings Estático", "CENCOSUD", "UNICENTER", "Pasillo PACO", "Estático"],
    ["EZEPAP002 - 1", "AA2000", "EZEIZA", "EZEIZA", "MUPIS", "Estático"],
    ["EZE-TS-1", "AA2000", "TRIPSTORE", "EZEIZA", "Totem Simple - Ezeiza", "Digital"],
    ["C1 - CAB", "Pantalla Led", "PLED", "CABILDO", "PCAB-3600seg", "Digital"],
    ["C5 - OLA", "Pantalla Led", "PLED", "OLAZABAL", "PTRI-3600seg", "Digital"],
    ["REM-DB-1", "Shoppings Digital", "REMEROS", "REMEROS", "TV Led", "Digital"],
    ["REM-DB-3", "Shoppings Digital", "REMEROS", "REMEROS", "TV Led", "Digital"],
    ["REM-CH-1D", "Shoppings Estático", "REMEROS", "REMEROS", "Chupete", "Estático"],
    ["Pantalla 4", "London Supply", "LS", "CALAFATE", "Pantalla", "Digital"],
])


# ---------------------------------------------------------------------------
# Extraccion
# ---------------------------------------------------------------------------

class TestParseOT:
    def test_numero_simple_y_float(self):
        assert [t.numero for t in st.parse_ot_cell(4284).tokens] == [4284]
        assert [t.numero for t in st.parse_ot_cell(4284.0).tokens] == [4284]

    def test_prefijo_proveedor_no_es_parte_del_id(self):
        c = st.parse_ot_cell("PUBLI 4381")
        assert c.tokens == [st.OTToken(4381, "PUBLI", False)]

    def test_varias_ot_en_celda(self):
        c = st.parse_ot_cell("2186/4284")
        assert [t.numero for t in c.tokens] == [2186, 4284]

    def test_bonificada_b(self):
        c = st.parse_ot_cell("PUBLISHING 4554 / B 4726")
        assert c.tokens[0] == st.OTToken(4554, "PUBLISHING", False)
        assert c.tokens[1].numero == 4726 and c.tokens[1].bonificada
        c2 = st.parse_ot_cell("PUBLI 4433 / PUBLI B 4510")
        assert c2.tokens[1] == st.OTToken(4510, "PUBLI", True)

    def test_texto_sin_ot(self):
        c = st.parse_ot_cell("DECO NAV")
        assert c.tokens == [] and c.texto_no_ot == "DECO NAV"
        c = st.parse_ot_cell("DECO NAV / 4500")
        assert [t.numero for t in c.tokens] == [4500] and c.texto_no_ot == "DECO NAV"


class TestParseFecha:
    def test_tipos(self):
        assert st.parse_fecha(dt.datetime(2026, 3, 1)) == (D(2026, 3, 1), "FECHA")
        assert st.parse_fecha(46266) == (D(2026, 9, 1), "SERIAL")
        assert st.parse_fecha("INDET") == (None, "INDET")
        assert st.parse_fecha("indef") == (None, "INDET")
        assert st.parse_fecha("CONSULTAR") == (None, "INVALIDA")
        assert st.parse_fecha(None) == (None, "VACIA")
        assert st.parse_fecha("15/04/2026") == (D(2026, 4, 15), "FECHA")


class TestEstructura:
    def test_mes_faltante_se_infiere_por_secuencia(self):
        rows = _hoja([], mes_faltante=1)
        est = st.detectar_estructura(rows, "T")
        assert [b.mes for b in est.bloques] == [D(2025, 11, 1), D(2025, 12, 1), D(2026, 1, 1), D(2026, 2, 1)]
        assert est.attr_cols["Codigo"] == 5

    def test_etiqueta_inconsistente_falla(self):
        rows = _hoja([], meses=("NOV", "ENE", "ENE", "FEB"))
        with pytest.raises(st.StagingError):
            st.detectar_estructura(rows, "T")

    def test_extraer_ocurrencias_y_alcance(self):
        rows = _hoja([
            ["UN-TRIEDRO-1A-2", "CENCOSUD", "UNICENTER", "Triedro", {2: ("2186/4284", "MOTO", dt.datetime(2026, 1, 1), dt.datetime(2026, 1, 31), None)}],
            ["", "CENCOSUD", "", "", {3: (4164, "RIO GALLEGOS", None, None, None)}],
        ])
        occ = st.extraer_hoja(rows, "CENCO D")
        assert [(o.token.numero, o.n_tokens, o.ref) for o in occ[:2]] == [(2186, 2, "CENCO D!R3"), (4284, 2, "CENCO D!R3")]
        assert occ[2].alcance == "FILA_SIN_CODIGO"

    def test_tripstore_compartida_separa_london(self):
        assert st.alcance_fila("TRIPSTORE Y LS D", {"Clas": "TRIPSTORE"}) == "EN_ALCANCE"
        assert st.alcance_fila("TRIPSTORE Y LS D", {"Clas": "LONDON SUPPLY"}) == "LONDON/LS"
        assert st.alcance_fila("TRIPSTORE Y LS D", {"Clas": "OTRA"}) == "LONDON/LS"
        assert st.alcance_fila("CENCO F", {"Clas": "CENCOMEDIA"}) == "CENCOMEDIA"


# ---------------------------------------------------------------------------
# Reconstruccion de fechas
# ---------------------------------------------------------------------------

class TestFechas:
    def test_bloques_contiguos_se_consolidan(self):
        occ = [_occ(D(2026, 3, 1), D(2026, 3, 1), D(2026, 3, 31)), _occ(D(2026, 4, 1), D(2026, 4, 1), D(2026, 4, 30))]
        r = st.reconstruir_periodos(occ, VENTANA)
        assert [(p.ini, p.fin, p.regla, p.dudas) for p in r.periodos] == [
            (D(2026, 3, 1), D(2026, 4, 30), "BLOQUES_CONTIGUOS_CONSOLIDADOS", [])]

    def test_rango_repetido_mes_a_mes(self):
        occ = [_occ(D(2026, m, 1), D(2026, 3, 1), D(2026, 5, 31)) for m in (3, 4, 5)]
        r = st.reconstruir_periodos(occ, VENTANA)
        assert [(p.ini, p.fin, p.regla) for p in r.periodos] == [(D(2026, 3, 1), D(2026, 5, 31), "RANGO_UNICO")]
        assert not r.problemas and not r.periodos[0].dudas

    def test_hueco_real_no_se_une(self):
        occ = [_occ(D(2026, 11, 1), D(2026, 11, 1), D(2026, 12, 31)), _occ(D(2026, 12, 1), D(2026, 11, 1), D(2026, 12, 31)),
               _occ(D(2027, 2, 1), D(2027, 2, 1), D(2027, 2, 28))]
        r = st.reconstruir_periodos(occ, VENTANA)
        assert r.disjuntos
        assert [(p.ini, p.fin) for p in r.periodos] == [(D(2026, 11, 1), D(2026, 12, 31)), (D(2027, 2, 1), D(2027, 2, 28))]

    def test_extension_progresiva_valida(self):
        occ = [_occ(D(2026, 3, 1), D(2026, 3, 1), D(2026, 6, 30)), _occ(D(2026, 4, 1), D(2026, 3, 1), D(2026, 6, 30)),
               _occ(D(2026, 6, 1), D(2026, 3, 1), D(2026, 7, 31)), _occ(D(2026, 5, 1), D(2026, 3, 1), D(2026, 6, 30)),
               _occ(D(2026, 7, 1), D(2026, 3, 1), D(2026, 7, 31))]
        r = st.reconstruir_periodos(occ, VENTANA)
        p = r.periodos[0]
        assert (p.ini, p.fin, p.regla, p.dudas) == (D(2026, 3, 1), D(2026, 7, 31), "EXTENSION_FIN_PROGRESIVA", [])

    def test_fin_que_retrocede_es_duda(self):
        occ = [_occ(D(2026, 3, 1), D(2026, 3, 1), D(2026, 4, 30)), _occ(D(2026, 4, 1), D(2026, 3, 1), D(2026, 4, 15))]
        p = st.reconstruir_periodos(occ, VENTANA).periodos[0]
        assert p.regla == "RANGOS_SOLAPADOS" and "RANGOS_DISTINTOS_SOLAPADOS" in p.dudas

    def test_mes_del_rango_sin_presencia_es_duda(self):
        # INICIO arrastrado de una campana anterior: rango 2025-01 a 2026-12, OT solo desde 2026-08.
        occ = [_occ(D(2026, m, 1), D(2025, 1, 20), D(2026, 12, 31)) for m in range(8, 13)]
        p = st.reconstruir_periodos(occ, VENTANA).periodos[0]
        assert any(d.startswith("MESES_SIN_PRESENCIA") for d in p.dudas)

    def test_celda_compartida_no_aporta_fechas(self):
        occ = [_occ(D(2026, 7, 1), D(2026, 3, 1), D(2026, 12, 31), n_tokens=2)]
        r = st.reconstruir_periodos(occ, VENTANA)
        assert r.periodos == [] and "FECHAS_SOLO_EN_CELDA_COMPARTIDA" in r.problemas

    def test_celda_compartida_cuenta_como_presencia(self):
        occ = [_occ(D(2026, 6, 1), D(2026, 6, 1), D(2026, 7, 31)), _occ(D(2026, 7, 1), D(2026, 1, 1), D(2026, 12, 31), n_tokens=2)]
        r = st.reconstruir_periodos(occ, VENTANA)
        assert [(p.ini, p.fin, p.dudas) for p in r.periodos] == [(D(2026, 6, 1), D(2026, 7, 31), [])]

    def test_indeterminada_y_no_interpretable(self):
        r = st.reconstruir_periodos([_occ(D(2026, 1, 1), D(2025, 2, 15), None, tipo_fin="INDET")], VENTANA)
        assert "FECHA_INDETERMINADA" in r.problemas and not r.periodos
        r = st.reconstruir_periodos([_occ(D(2026, 1, 1), D(2026, 1, 1), None, tipo_fin="INVALIDA")], VENTANA)
        assert "FECHA_NO_INTERPRETABLE" in r.problemas

    def test_bloque_sin_fechas_solo_aporta_presencia(self):
        occ = [_occ(D(2026, 3, 1), D(2026, 3, 1), D(2026, 4, 30)), _occ(D(2026, 4, 1), None, None)]
        r = st.reconstruir_periodos(occ, VENTANA)
        assert not r.problemas and r.periodos[0].dudas == []


# ---------------------------------------------------------------------------
# Crosswalk
# ---------------------------------------------------------------------------

def _xw(hoja, codigo, ubic, desc, clas="CENCOSUD"):
    o = _occ(D(2026, 3, 1), D(2026, 3, 1), D(2026, 3, 31))
    o.hoja, o.codigo, o.attrs = hoja, codigo, {"Ubic": ubic, "Descripcion": desc, "Clas": clas}
    o.alcance = st.alcance_fila(hoja, o.attrs)
    return o


class TestCrosswalk:
    def _run(self, occ):
        return st.construir_crosswalk(occ, st.IndiceMaestro(MAESTRO)).set_index("ElementoOrigen")

    def test_metodos_alta(self):
        xw = self._run([
            _xw("AEROPUERTOS F", "EZEPAP002 - 1", "EZEIZA", "MUPIS"),
            _xw("CENCO D", "UN-TRIEDRO-1A-2", "UNICENTER", "Triedro Digital - Octogono Chico"),
            _xw("TRIPSTORE Y LS D", "EZE-TS-1-V2", "EZEIZA", "Totem Simple - Ezeiza", clas="TRIPSTORE"),
            _xw("PLED", "PCAB-3600seg-V3", "CAB", "PLED"),
            _xw("PLED", "PTRI-3600seg-V09", "TRI", "PLED"),
            _xw("CENCO D", "UNI-DOBLES-1G-5-V2", "UNICENTER", "Totem Digital Garganta Reloj - Nivel 1"),
        ])
        esperado = {
            "EZEPAP002 - 1": ("EZEPAP002 - 1", "EXACTO"),
            "UN-TRIEDRO-1A-2": ("UNI-TRIEDRO-1A-2", "PREFIJO_UN_UNI"),
            "EZE-TS-1-V2": ("EZE-TS-1", "SUFIJO_SLOT"),
            "PCAB-3600seg-V3": ("C1 - CAB", "PLED_DESCRIPCION_MAESTRO"),
            "PTRI-3600seg-V09": ("C5 - OLA", "PLED_DESCRIPCION_MAESTRO"),
            "UNI-DOBLES-1G-5-V2": ("UNI-TOTEMD-1G-5", "DESCRIPCION_Y_POSICION"),
        }
        for cod, (eid, metodo) in esperado.items():
            assert (xw.loc[cod, "ElementoID_OCU26"], xw.loc[cod, "MetodoMatch"], xw.loc[cod, "Confianza"]) == (eid, metodo, "ALTA"), cod

    def test_slot_con_descripcion_autocompletada_hereda_de_hermanos(self):
        xw = self._run([
            _xw("CENCO D", "UNI-DOBLES-1G-5-V1", "UNICENTER", "Totem Digital Garganta Reloj - Nivel 1"),
            _xw("CENCO D", "UNI-DOBLES-1G-5-V12", "UNICENTER", "Totem Digital Garganta Reloj - Nivel 2"),  # autocompletado
            _xw("CENCO D", "UNI-DOBLES-1G-5-V13", "UNICENTER", "Totem Digital Garganta Reloj - Nivel 9"),
        ])
        for cod in ("UNI-DOBLES-1G-5-V12", "UNI-DOBLES-1G-5-V13"):
            assert (xw.loc[cod, "ElementoID_OCU26"], xw.loc[cod, "Confianza"], xw.loc[cod, "MetodoMatch"]) == (
                "UNI-TOTEMD-1G-5", "ALTA", "HERENCIA_SLOTS_HERMANOS"), cod

    def test_sin_hermano_alta_no_hereda(self):
        xw = self._run([_xw("REM-PIL", "REM-TS 1-V1", "REMEROS", "TV Led / x", clas="REMEROS"),
                        _xw("REM-PIL", "REM-TS 1-V2", "REMEROS", "TV Led / x", clas="REMEROS")])
        assert set(xw["Confianza"]) == {"MEDIA"} and set(xw["ElementoID_OCU26"]) == {""}

    def test_un_uni_no_es_replace_global(self):
        # UN-LONAER3x5-K2 -> UNI-LONAER3x5-K2 no existe: no se inventa ni se toma K1.
        xw = self._run([_xw("CENCO F", "UN-LONAER3x5-K2", "UNICENTER", "Pasillo PACO")])
        r = xw.loc["UN-LONAER3x5-K2"]
        assert r["ElementoID_OCU26"] == "" and r["Confianza"] != "ALTA"
        assert "no existe en maestro" in r["Observacion"]

    def test_pled_sin_descripcion_en_maestro_no_matchea(self):
        xw = self._run([_xw("PLED", "PALS-3600seg-V1", "ALS", "PLED")])
        assert xw.loc["PALS-3600seg-V1", "Confianza"] == "BAJA"
        assert xw.loc["PALS-3600seg-V1", "ElementoID_OCU26"] == ""

    def test_familia_por_descripcion_es_ambigua(self):
        occ = [_xw("REM-PIL", "REM-CH-1D", "REM", "Chupete", clas="REMEROS"),
               _xw("REM-PIL", "REM-TS 1-V1", "REM", "TV Led / lado derecho", clas="REMEROS")]
        xw = self._run(occ)
        r = xw.loc["REM-TS 1-V1"]
        assert r["Confianza"] == "MEDIA" and r["ElementoID_OCU26"] == ""
        assert r["Candidatos"] == "REM-DB-1, REM-DB-3"

    def test_medio_contradictorio_baja_a_media(self):
        xw = self._run([_xw("CENCO F", "UNI-TRIEDRO-1A-2", "UNICENTER", "Triedro Digital - Octogono Chico")])
        assert xw.loc["UNI-TRIEDRO-1A-2", "Confianza"] == "MEDIA"

    def test_london_nunca_se_mapea(self):
        xw = st.construir_crosswalk([_xw("TRIPSTORE Y LS D", "Pantalla 4 - V3", "CALAFATE", "Pantalla", clas="LONDON SUPPLY")],
                                    st.IndiceMaestro(MAESTRO))
        assert xw.empty
        assert "Pantalla 4" not in st.IndiceMaestro(MAESTRO).por_id


# ---------------------------------------------------------------------------
# Deduplicacion
# ---------------------------------------------------------------------------

def _fc(ini, fin, cid="HIST-1", indef=False):
    return st.FilaCampana(cid, ini, fin, indef)


class TestDedup:
    def test_resultados(self):
        cmp = st.comparar_con_campanas
        pres = [D(2026, 3, 1)]
        assert cmp([], D(2026, 3, 1), D(2026, 3, 31), pres).resultado == "NO_EXISTE"
        assert cmp([_fc(D(2026, 3, 1), D(2026, 3, 31))], D(2026, 3, 1), D(2026, 3, 31), pres).resultado == "EXACTO"
        assert cmp([_fc(D(2026, 1, 1), D(2026, 12, 31))], D(2026, 3, 1), D(2026, 3, 31), pres).resultado == "CUBIERTO"
        # dos filas contiguas cubren el periodo
        assert cmp([_fc(D(2026, 3, 1), D(2026, 3, 15)), _fc(D(2026, 3, 16), D(2026, 4, 30))],
                   D(2026, 3, 1), D(2026, 3, 31), pres).resultado == "CUBIERTO"
        r = cmp([_fc(D(2026, 3, 1), D(2026, 6, 30))], D(2026, 3, 1), D(2026, 12, 31), pres)
        assert (r.resultado, r.no_cubierto) == ("PARCIAL", [(D(2026, 7, 1), D(2026, 12, 31))])
        assert cmp([_fc(D(2026, 5, 21), D(2026, 6, 21))], D(2026, 7, 21), D(2026, 8, 21), pres).resultado == "OTRAS_FECHAS"
        assert cmp([_fc(D(2026, 1, 1), None, indef=True)], D(2026, 3, 1), D(2026, 3, 31), pres).resultado == "CUBIERTO"

    def test_sin_fechas_usa_meses_de_presencia(self):
        r = st.comparar_con_campanas([_fc(D(2026, 1, 1), D(2026, 7, 12))], None, None, [D(2025, 12, 1), D(2026, 1, 1), D(2026, 6, 1)])
        assert r.resultado == "CUBIERTO"

    def test_periodo_disjunto_de_ot_cargada_se_promueve(self):
        regs = [{"Estado_Staging": "YA_EXISTE", "_motivos": ["YA_EXISTE_EXACTO"]},
                {"Estado_Staging": "REVISAR", "_motivos": [st.MOTIVO_SOLO_OTRAS_FECHAS]},
                {"Estado_Staging": "REVISAR", "_motivos": [st.MOTIVO_SOLO_OTRAS_FECHAS, "FECHAS_DUDOSAS"]}]
        st.promover_periodos_disjuntos(regs)
        assert [r["Estado_Staging"] for r in regs] == ["YA_EXISTE", "IMPORTAR", "REVISAR"]

    def test_otras_fechas_sin_periodo_cargado_queda_en_revisar(self):
        regs = [{"Estado_Staging": "REVISAR", "_motivos": [st.MOTIVO_SOLO_OTRAS_FECHAS]}]
        st.promover_periodos_disjuntos(regs)
        assert regs[0]["Estado_Staging"] == "REVISAR"

    def test_clave_negocio_formato_historico(self):
        assert st.clave_negocio(4572, "C1 - CAB", D(2026, 5, 15), D(2026, 6, 16)) == "4572|C1 - CAB|2026-05-15|2026-06-16||"


# ---------------------------------------------------------------------------
# Clasificacion de punta a punta (exclusiones incluidas)
# ---------------------------------------------------------------------------

def _campanas(rows):
    df = pd.DataFrame(columns=CAMPANAS_HEADERS)
    for k, (ot, eid, ini, fin) in enumerate(rows):
        df.loc[k] = {"CargaID": f"HIST-{k}", "IDCampaña": float(ot), "ElementoID": eid,
                     "FechaInicio": pd.Timestamp(ini), "FechaFin": pd.Timestamp(fin), "FechaIndefinida": "No"}
    return df


class TestClasificacion:
    def _clasificar(self, occ, campanas_rows=()):
        idx_m = st.IndiceMaestro(MAESTRO)
        idx_c = st.IndiceCampanas(_campanas(list(campanas_rows)))
        xw = st.construir_crosswalk(occ, idx_m)
        ventanas = {h: VENTANA for h in st.HOJAS_PROCESAR}
        return st.consolidar(occ, xw, idx_m, idx_c, ventanas)

    def _o(self, hoja, codigo, ubic, desc, mes, ini, fin, ot, clas="CENCOSUD", token=True, bonif=False, pauta="P"):
        o = _xw(hoja, codigo, ubic, desc, clas)
        o.mes, o.ini, o.fin, o.pauta = mes, ini, fin, pauta
        o.tipo_ini = "FECHA" if ini else "VACIA"
        o.tipo_fin = "FECHA" if fin else "VACIA"
        if token:
            o.token, o.ot_raw = st.OTToken(ot, bonificada=bonif), str(ot)
        else:
            o.token, o.n_tokens, o.ot_raw, o.texto_no_ot = None, 0, str(ot), str(ot)
        return o

    def test_estados(self):
        uni = ("CENCO D", "UN-TRIEDRO-1A-2", "UNICENTER", "Triedro Digital - Octogono Chico")
        occ = [
            self._o(*uni, D(2026, 8, 1), D(2026, 8, 1), D(2026, 8, 31), 4800),            # nueva -> IMPORTAR
            self._o(*uni, D(2026, 3, 1), D(2026, 3, 1), D(2026, 3, 31), 4300),            # ya cargada
            self._o(*uni, D(2025, 11, 1), D(2025, 11, 1), D(2025, 11, 30), 4100),         # fuera de 2026
            self._o(*uni, D(2026, 9, 1), D(2026, 9, 1), D(2026, 9, 30), 4810, bonif=True),  # B -> REVISAR
            self._o("CENCO F", "UNI-ASC -1", "UNICENTER", "Ascensor", D(2026, 8, 1), D(2026, 8, 1), D(2026, 8, 31), 4729),
            self._o("TRIPSTORE Y LS D", "Pantalla 4 - V3", "CALAFATE", "Pantalla", D(2026, 8, 1), D(2026, 8, 1),
                    D(2026, 8, 31), 4900, clas="LONDON SUPPLY"),
            self._o("CENCO F", "UNI-TRIEDRO-1A-2", "UNICENTER", "x", D(2026, 12, 1), None, None, "DECO NAV", token=False),
        ]
        df = self._clasificar(occ, [(4300, "UNI-TRIEDRO-1A-2", D(2026, 3, 1), D(2026, 3, 31))])
        estado = {(None if pd.isna(r["IDCampaña"]) else int(r["IDCampaña"]), r["TRZ_ElementoOrigen"]): (r["Estado_Staging"], r["Motivo"])
                  for _, r in df.iterrows()}
        assert estado[(4800, "UN-TRIEDRO-1A-2")] == ("IMPORTAR", "NUEVA_ASIGNACION")
        assert estado[(4300, "UN-TRIEDRO-1A-2")] == ("YA_EXISTE", "YA_EXISTE_EXACTO")
        assert estado[(4100, "UN-TRIEDRO-1A-2")] == ("NO_IMPORTAR", "FUERA_DE_2026")
        assert estado[(4810, "UN-TRIEDRO-1A-2")][0] == "REVISAR" and "OT_BONIFICADA_PREFIJO_B" in estado[(4810, "UN-TRIEDRO-1A-2")][1]
        assert estado[(4729, "UNI-ASC -1")][0] == "REVISAR" and "ELEMENTO_NO_INEQUIVOCO" in estado[(4729, "UNI-ASC -1")][1]
        assert estado[(4900, "Pantalla 4 - V3")] == ("NO_IMPORTAR", "FUERA_DE_ALCANCE_LONDON/LS")
        assert estado[(None, "UNI-TRIEDRO-1A-2")] == ("NO_IMPORTAR", "NO_COMERCIAL_SIN_OT")
        assert set(df["Estado_Staging"]) <= set(st.ESTADOS)

    def test_importar_columnas_destino(self):
        uni = ("CENCO D", "UN-TRIEDRO-1A-2", "UNICENTER", "Triedro Digital - Octogono Chico")
        df = self._clasificar([self._o(*uni, D(2026, 8, 1), D(2026, 8, 1), D(2026, 8, 31), 4800, pauta="PLAYSTATION")])
        df.insert(0, "StagingID", "STG2B-000001")
        df["TRZ_Medio_OCU26"] = "Digital"
        df["TRZ_CircuitoDashboard_OCU26"] = "Shoppings Digital"
        imp = st.preparar_importar(df, {"UNI-TRIEDRO-1A-2": "Digital"})
        assert list(imp.columns[: len(CAMPANAS_HEADERS)]) == CAMPANAS_HEADERS
        r = imp.iloc[0]
        assert (r["IDCampaña"], r["ElementoID"], r["FechaInicio"], r["FechaFin"]) == (4800, "UNI-TRIEDRO-1A-2", D(2026, 8, 1), D(2026, 8, 31))
        assert r["ClaveNegocio"] == "4800|UNI-TRIEDRO-1A-2|2026-08-01|2026-08-31||"
        assert r["TipoCargaDeclarado"] == "Digital" and r["FechaIndefinida"] == "No" and r["Campaña"] == "PLAYSTATION"
        for vacia in ("CargaID", "Cliente", "Marca", "Agencia", "Proveedor", "EstadoValidacion", "Estado"):
            assert pd.isna(r[vacia]), vacia
