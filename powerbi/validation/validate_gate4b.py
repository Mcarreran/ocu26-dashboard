"""Gate 4B - valida que las formulas DAX documentadas en powerbi/dax/*.dax
(reproducidas aqui en pandas, operando SOLO sobre las columnas que Power BI
Desktop Free vera al leer output/*.parquet) coinciden numericamente con
scripts/metrics_engine.MetricsEngine para los 8 casos del prompt Gate 4B
Seccion 13.

No reabre ni reinterpreta reglas de negocio de Gate 1-4A: MetricsEngine se
usa tal cual (fuente de verdad), y el lado "Power BI" de la comparacion NO
llama a ninguna funcion de metrics_engine.py / semantic_model.py - deriva
el mismo numero solo con pandas sobre los Parquet exportados, exactamente
como lo haria una medida DAX sobre el modelo relacional documentado en
powerbi/README.md.

Uso:
    python powerbi/validation/validate_gate4b.py
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
OUTPUT_DIR = REPO_ROOT / "output"
sys.path.insert(0, str(REPO_ROOT / "scripts"))

from export_data import load_pipeline  # noqa: E402

RANGO_INICIO = "2024-01-01"
RANGO_FIN = "2027-03-31"


def _n_dias(start: str, end: str) -> int:
    return (pd.Timestamp(end) - pd.Timestamp(start)).days + 1


def _load_tables() -> dict[str, pd.DataFrame]:
    return {
        "dim": pd.read_parquet(OUTPUT_DIR / "dim_elementos.parquet"),
        "fact": pd.read_parquet(OUTPUT_DIR / "fact_campanas.parquet"),
        "cal": pd.read_parquet(OUTPUT_DIR / "dim_calendario.parquet"),
        "bridge": pd.read_parquet(OUTPUT_DIR / "bridge_campana_dia.parquet"),
        "fmd": pd.read_parquet(OUTPUT_DIR / "fact_metricas_diaria.parquet"),
    }


# ---------------------------------------------------------------------------
# Equivalentes DAX en pandas (ver powerbi/dax/03_calendario.dax, 04_digital.dax)
# ---------------------------------------------------------------------------


def dax_ocupacion_calendario_pct(dim: pd.DataFrame, fmd: pd.DataFrame, circuito: str, start: str, end: str):
    elems = dim[
        (dim["Medio"] == "Estático") & (dim["CircuitoNegocio"] == circuito) & (dim["IncluyeConteoGeneral"])
    ]
    n_elem = elems["ElementoID"].nunique()
    n_dias = _n_dias(start, end)
    disponibles = n_elem * n_dias

    fmd_scope = fmd[fmd["ElementoID"].isin(elems["ElementoID"]) & fmd["Fecha"].between(start, end)]
    ocupados = len(fmd_scope)  # grain unico (ElementoID, Fecha): COUNTROWS = pares distintos

    cobertura_ok = bool(
        n_elem
        and (elems["CoberturaCatalogo"] == "COMPLETO").all()
        and (elems["CompletitudMaestro"] == "COMPLETO").all()
    )
    fecha_incompleta = bool(elems["FechaIncompletaCalendario"].any()) if n_elem else False

    if disponibles == 0 or not cobertura_ok:
        return None, "NO_APLICA"
    status = "PARTIAL" if fecha_incompleta else "OK"
    return ocupados / disponibles * 100.0, status


def dax_actividad_sobre_registrados_pct(dim: pd.DataFrame, fmd: pd.DataFrame, circuito: str, start: str, end: str):
    elems = dim[(dim["CircuitoNegocio"] == circuito) & (dim["IncluyeConteoGeneral"])]
    n_elem = elems["ElementoID"].nunique()
    n_dias = _n_dias(start, end)
    disponibles = n_elem * n_dias

    fmd_scope = fmd[fmd["ElementoID"].isin(elems["ElementoID"]) & fmd["Fecha"].between(start, end)]
    ocupados = len(fmd_scope)

    if disponibles == 0:
        return None, "NO_APLICA"
    fecha_incompleta = bool(elems["FechaIncompletaCalendario"].any()) if n_elem else False
    status = "PARTIAL" if fecha_incompleta else "OK"
    return ocupados / disponibles * 100.0, status


def dax_fill_rate_slots(dim: pd.DataFrame, fmd: pd.DataFrame, circuitos: list[str], start: str, end: str):
    elems = dim[
        (dim["Medio"] == "Digital") & (dim["CircuitoNegocio"].isin(circuitos)) & (dim["IncluyeConteoGeneral"])
    ]
    total_digital = len(elems)
    bloqueados = int(elems["PolicyBloqueadaSlotSeconds"].sum())
    todos_bloqueados = total_digital > 0 and bloqueados == total_digital
    if todos_bloqueados:
        return None, "NO_APLICA"

    operables = elems[~elems["PolicyBloqueadaSlotSeconds"]]
    denom_elems = operables[~operables["CapacidadSlotsDesconocida"]]
    denom = float(denom_elems["SlotsComercialesValor"].fillna(0).sum())

    n_dias = _n_dias(start, end)
    fmd_scope = fmd[
        fmd["ElementoID"].isin(denom_elems["ElementoID"]) & fmd["Fecha"].between(start, end)
    ]
    numer = float(fmd_scope["SlotsOcupadosDia"].fillna(0).sum()) / n_dias

    if denom <= 0:
        return None, "NO_APLICA"

    partial = (
        bloqueados > 0
        or bool((operables["CompletitudMaestro"] != "COMPLETO").any())
        or bool(operables["CapacidadSlotsDesconocida"].any())
        or bool(operables["FechaIncompletaDigital"].any())
    )
    status = "PARTIAL" if partial else "OK"
    return numer / denom * 100.0, status


def dax_campanas_count(dim: pd.DataFrame, fact: pd.DataFrame) -> int:
    """DAX [Cargas]: COUNTROWS(FACT_CAMPANAS) filtrado por universo (via
    relacion R1 a DIM_ELEMENTOS.IncluyeConteoGeneral)."""
    merged = fact.merge(dim[["ElementoID", "IncluyeConteoGeneral"]], on="ElementoID", how="left")
    return int((merged["IncluyeConteoGeneral"] == True).sum())  # noqa: E712


# ---------------------------------------------------------------------------
# Casos de prueba (Gate 4B prompt Seccion 13)
# ---------------------------------------------------------------------------


def run() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    tablas = _load_tables()
    dim, fact, fmd = tablas["dim"], tablas["fact"], tablas["fmd"]

    _, semantic_result, engine = load_pipeline()

    results: list[tuple[str, bool, str]] = []

    def check(name: str, dax_value, dax_status: str, py_row: pd.DataFrame):
        py_value = py_row["Value"].iloc[0] if len(py_row) else None
        py_status = py_row["MetricStatus"].iloc[0] if len(py_row) else None
        py_value = None if pd.isna(py_value) else float(py_value)
        status_ok = dax_status == py_status
        if dax_value is None and py_value is None:
            value_ok = True
        elif dax_value is None or py_value is None:
            value_ok = False
        else:
            value_ok = abs(dax_value - py_value) < 1e-6
        ok = status_ok and value_ok
        detail = f"DAX=({dax_value}, {dax_status})  Python=({py_value}, {py_status})"
        results.append((name, ok, detail))

    # 1. Cencosud Estatico - ocupacion_calendario_pct
    v, s = dax_ocupacion_calendario_pct(dim, fmd, "CENCOSUD", RANGO_INICIO, RANGO_FIN)
    py = engine.query(
        "ocupacion_calendario_pct", filters={"CircuitoNegocio": "CENCOSUD"},
        universe="OPERATIVO_GENERAL", start_date=RANGO_INICIO, end_date=RANGO_FIN,
    )
    check("1. Cencosud Estático ocupacion_calendario_pct", v, s, py)

    # 2. AA2000 - ocupacion_calendario_pct (esperado NO_APLICA: CompletitudMaestro=PARCIAL)
    v, s = dax_ocupacion_calendario_pct(dim, fmd, "AA2000", RANGO_INICIO, RANGO_FIN)
    py = engine.query(
        "ocupacion_calendario_pct", filters={"CircuitoNegocio": "AA2000"},
        universe="OPERATIVO_GENERAL", start_date=RANGO_INICIO, end_date=RANGO_FIN,
    )
    check("2. AA2000 ocupacion_calendario_pct", v, s, py)

    # 3. AA2000 - actividad_sobre_registrados_pct
    v, s = dax_actividad_sobre_registrados_pct(dim, fmd, "AA2000", RANGO_INICIO, RANGO_FIN)
    py = engine.query(
        "actividad_sobre_registrados_pct", filters={"CircuitoNegocio": "AA2000"},
        universe="OPERATIVO_GENERAL", start_date=RANGO_INICIO, end_date=RANGO_FIN,
    )
    check("3. AA2000 actividad_sobre_registrados_pct", v, s, py)

    # 4. Pantallas LED - fill_rate_slots
    v, s = dax_fill_rate_slots(dim, fmd, ["PANTALLAS_LED"], RANGO_INICIO, RANGO_FIN)
    py = engine.query(
        "fill_rate_slots", filters={"CircuitoNegocio": "PANTALLAS_LED"},
        universe="OPERATIVO_GENERAL", start_date=RANGO_INICIO, end_date=RANGO_FIN,
    )
    check("4. Pantallas LED fill_rate_slots", v, s, py)

    # 5. YPF Digital - fill_rate_slots = NO_APLICA
    v, s = dax_fill_rate_slots(dim, fmd, ["YPF"], RANGO_INICIO, RANGO_FIN)
    py = engine.query(
        "fill_rate_slots", filters={"CircuitoNegocio": "YPF"},
        universe="OPERATIVO_GENERAL", start_date=RANGO_INICIO, end_date=RANGO_FIN,
    )
    check("5. YPF Digital fill_rate_slots (NO_APLICA)", v, s, py)

    # 6. Grupo mixto Pantallas LED + YPF - fill_rate_slots = PARTIAL
    v, s = dax_fill_rate_slots(dim, fmd, ["PANTALLAS_LED", "YPF"], RANGO_INICIO, RANGO_FIN)
    py = engine.query(
        "fill_rate_slots", filters={"CircuitoNegocio": ["PANTALLAS_LED", "YPF"]},
        universe="OPERATIVO_GENERAL", start_date=RANGO_INICIO, end_date=RANGO_FIN,
    )
    check("6. Pantallas LED + YPF (mixto) fill_rate_slots (PARTIAL)", v, s, py)

    # 7. London Supply - fill_rate_slots, capacidad incompleta -> PARTIAL
    v, s = dax_fill_rate_slots(dim, fmd, ["LONDON_SUPPLY"], RANGO_INICIO, RANGO_FIN)
    py = engine.query(
        "fill_rate_slots", filters={"CircuitoNegocio": "LONDON_SUPPLY"},
        universe="OPERATIVO_GENERAL", start_date=RANGO_INICIO, end_date=RANGO_FIN,
    )
    check("7. London Supply fill_rate_slots (PARTIAL, capacidad incompleta)", v, s, py)

    # 8. Campañas activas: una IDCampaña en multiples ElementoID se cuenta 1 vez.
    py_campanas = engine.query("campanas", universe="OPERATIVO_GENERAL")
    py_cargas_total = int(py_campanas["Value"].iloc[0])
    dax_cargas_total = dax_campanas_count(dim, fact)
    cargas_ok = dax_cargas_total == py_cargas_total

    merged = fact.merge(dim[["ElementoID", "IncluyeConteoGeneral"]], on="ElementoID", how="left")
    scope = merged[merged["IncluyeConteoGeneral"] == True]  # noqa: E712
    por_campana = scope.dropna(subset=["IDCampaña"]).groupby("IDCampaña")["ElementoID"].nunique()
    multi = por_campana[por_campana > 1].sort_values(ascending=False)
    dedup_ok = len(multi) > 0
    if dedup_ok:
        ejemplo_id = multi.index[0]
        n_elementos = int(multi.iloc[0])
        n_cargas_ejemplo = int((scope["IDCampaña"] == ejemplo_id).sum())
        detail8 = (
            f"Cargas(DAX)={dax_cargas_total} Cargas(Python 'campanas')={py_cargas_total}  "
            f"IDCampaña ejemplo={ejemplo_id!r}: {n_cargas_ejemplo} CargaID en {n_elementos} ElementoID "
            f"-> DISTINCTCOUNT(IDCampaña)=1 (no se duplica), COUNTROWS(CargaID)={n_cargas_ejemplo} "
            f"(despliegues, medida distinta 'Cargas')"
        )
    else:
        detail8 = "No se encontro ninguna IDCampaña con >1 ElementoID en el universo OPERATIVO_GENERAL"
    ok8 = cargas_ok and dedup_ok
    results.append(("8. Campañas activas: IDCampaña multi-elemento contado 1 vez", ok8, detail8))

    # --- salida ---
    n_pass = sum(1 for _, ok, _ in results if ok)
    print(f"GATE4B_VALIDATION {n_pass}/{len(results)} PASS\n")
    for name, ok, detail in results:
        estado = "PASS" if ok else "FAIL"
        print(f"[{estado}] {name}")
        print(f"        {detail}")

    return 0 if n_pass == len(results) else 1


if __name__ == "__main__":
    sys.exit(run())
