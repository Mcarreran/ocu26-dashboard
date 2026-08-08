"""Motor generico de metricas del pipeline OCU26 - Gate 3B.

Se ejecuta DESPUES de scripts/semantic_model.py. Unica fuente de datos: el
dict devuelto por semantic_model.build_semantic_model() (que a su vez viene
de transform_data()). No reabre el Excel, no reimplementa Gate 1/Gate 2 ni
la resolucion de metadata de Gate 3B (semantic_model.py).

Expone una API generica:

    engine = MetricsEngine(semantic_result)
    engine.query(
        metric="elementos_registrados",
        group_by=["CircuitoNegocio", "Ciudad"],
        filters={"Medio": "Digital"},
        universe="OPERATIVO_GENERAL",
    )

Ninguna funcion aqui esta escrita por circuito (nada de calcular_ypf(),
calcular_cencosud(), etc.): toda diferencia de comportamiento entre
circuitos proviene de las dimensiones/flags ya resueltas por
semantic_model.py (CircuitoNegocio, PortfolioTier, CoberturaCatalogo,
CertezaDato, SlotsComerciales, etc.), nunca de un `if circuito == ...` aqui.

Resultado tidy: cada consulta devuelve un DataFrame con las columnas de
group_by pedidas + Metric, Value, Numerator, Denominator, MetricStatus,
Warnings. MetricStatus in {OK, PARTIAL, NO_APLICA, REQUIERE_CONFIRMACION}.
El motor nunca devuelve 0 cuando la respuesta real es "no sabemos": usa
NO_APLICA/REQUIERE_CONFIRMACION en Value (pd.NA) en esos casos.
"""

from __future__ import annotations

from typing import Any, Callable, Iterable, Iterator

import pandas as pd

from semantic_model import filter_universe

METRIC_STATUS_OK = "OK"
METRIC_STATUS_PARTIAL = "PARTIAL"
METRIC_STATUS_NO_APLICA = "NO_APLICA"
METRIC_STATUS_REQUIERE_CONFIRMACION = "REQUIERE_CONFIRMACION"

# Severidad para combinar varias razones simultaneas (Gate3B.1): la peor gana,
# los warnings de todas las razones aplicables se concatenan.
_STATUS_RANK = {
    METRIC_STATUS_OK: 0,
    METRIC_STATUS_PARTIAL: 1,
    METRIC_STATUS_REQUIERE_CONFIRMACION: 2,
    METRIC_STATUS_NO_APLICA: 3,
}

_FECHA_INCOMPLETA_WARNING = "Existen campanas con fechas incompletas excluidas del calculo temporal."


def _combine_status(reasons: list[tuple[str, str] | None]) -> tuple[str, str]:
    """Combina razones (status, warning) tomando la mas severa; entradas None
    se ignoran. Sin razones aplicables -> (OK, "")."""
    applicable = [r for r in reasons if r is not None]
    if not applicable:
        return METRIC_STATUS_OK, ""
    worst_status = max((s for s, _ in applicable), key=lambda s: _STATUS_RANK[s])
    warning = "; ".join(w for _, w in applicable if w)
    return worst_status, warning


class MetricsEngineError(Exception):
    """Error bloqueante del motor de metricas."""


class UnknownMetricError(MetricsEngineError):
    pass


class UnknownDimensionError(MetricsEngineError):
    pass


class UnknownUniverseError(MetricsEngineError):
    pass


class UnsupportedBusinessCaseError(MetricsEngineError):
    """Se encontro un caso de negocio para el que todavia no existe una
    formula aprobada (spot != 10s, exclusividad). El motor nunca inventa la
    matematica: prefiere fallar explicitamente a devolver un numero incorrecto."""


# ---------------------------------------------------------------------------
# Catalogo de metricas (Gate3B prompt Sec.41): documenta aplicabilidad para
# que ninguna vista futura tenga que hardcodear estos controles.
# ---------------------------------------------------------------------------

METRIC_CATALOG: dict[str, dict[str, Any]] = {
    "elementos_registrados": {
        "tipo": "conteo", "requiere_periodo": False, "requiere_medio": None,
        "requiere_cobertura_completa": False,
        "descripcion": "Cantidad de ElementoID en el maestro para el filtro/universo dado.",
    },
    "elementos_confirmados": {
        "tipo": "conteo", "requiere_periodo": False, "requiere_medio": None,
        "requiere_cobertura_completa": False,
        "descripcion": "Subconjunto de elementos_registrados con CertezaDato=CONFIRMADO (regla por circuito).",
    },
    "elementos_con_actividad": {
        "tipo": "conteo", "requiere_periodo": False, "requiere_medio": None,
        "requiere_cobertura_completa": False,
        "descripcion": "ElementoID con >=1 campana (global, o dentro de [start_date,end_date] si se especifica).",
    },
    "campanas": {
        "tipo": "conteo", "requiere_periodo": False, "requiere_medio": None,
        "requiere_cobertura_completa": False,
        "descripcion": "Filas de CAMPANAS en el filtro/universo/periodo dado.",
    },
    "campanas_activas": {
        "tipo": "conteo", "requiere_periodo": False, "requiere_medio": None,
        "requiere_cobertura_completa": False, "descripcion": "campanas con Estado=Activa.",
    },
    "campanas_reservadas": {
        "tipo": "conteo", "requiere_periodo": False, "requiere_medio": None,
        "requiere_cobertura_completa": False, "descripcion": "campanas con Estado=Reservada.",
    },
    "clientes_activos": {
        "tipo": "conteo_distinct", "requiere_periodo": False, "requiere_medio": None,
        "requiere_cobertura_completa": False, "descripcion": "Cliente distintos entre las campanas en scope.",
    },
    "marcas_activas": {
        "tipo": "conteo_distinct", "requiere_periodo": False, "requiere_medio": None,
        "requiere_cobertura_completa": False, "descripcion": "Marca distintas entre las campanas en scope.",
    },
    "dias_periodo": {
        "tipo": "escalar", "requiere_periodo": True, "requiere_medio": "Estático",
        "requiere_cobertura_completa": False, "descripcion": "Dias calendario inclusivos de [start_date,end_date].",
    },
    "elemento_dias_disponibles": {
        "tipo": "suma", "requiere_periodo": True, "requiere_medio": "Estático",
        "requiere_cobertura_completa": False,
        "descripcion": "elementos_registrados (Estatico) x dias_periodo.",
    },
    "elemento_dias_ocupados": {
        "tipo": "suma", "requiere_periodo": True, "requiere_medio": "Estático",
        "requiere_cobertura_completa": False,
        "descripcion": "Union de dias calendario ocupados por ElementoID (nunca duplica un dia solapado).",
    },
    "ocupacion_calendario_pct": {
        "tipo": "ratio", "requiere_periodo": True, "requiere_medio": "Estático",
        "requiere_cobertura_completa": True,
        "descripcion": "elemento_dias_ocupados/elemento_dias_disponibles. NO_APLICA si CoberturaCatalogo!=COMPLETO "
                        "o si CompletitudMaestro!=COMPLETO (universo conocido pero maestro incompleto, ej. AA2000).",
    },
    "actividad_sobre_registrados_pct": {
        "tipo": "ratio", "requiere_periodo": True, "requiere_medio": None,
        "requiere_cobertura_completa": False,
        "descripcion": "Mismo ratio calendario que ocupacion_calendario_pct (dias con campana / dias registrados) "
                        "pero sin exigir CoberturaCatalogo/CompletitudMaestro completos y sin filtrar por Medio: "
                        "funciona para Estatico y para Digital (ej. YPF). Nunca se confunde con ocupacion total.",
    },
    "slots_comerciales": {
        "tipo": "suma", "requiere_periodo": True, "requiere_medio": "Digital",
        "requiere_cobertura_completa": False,
        "descripcion": "Suma de SlotsComerciales (capacidad efectiva, no legacy) de los elementos en scope.",
    },
    "slots_ocupados": {
        "tipo": "suma", "requiere_periodo": True, "requiere_medio": "Digital",
        "requiere_cobertura_completa": False,
        "descripcion": "Promedio diario de campanas concurrentes por ElementoID (CargaID, no SalidasVendidas), sumado "
                        "en el grupo. NO_APLICA para circuitos en metric_policies (ej. YPF: Brand Plus no administra el CMS).",
    },
    "slots_disponibles": {
        "tipo": "suma", "requiere_periodo": True, "requiere_medio": "Digital",
        "requiere_cobertura_completa": False,
        "descripcion": "slots_comerciales - slots_ocupados (solo elementos con capacidad conocida). "
                        "NO_APLICA para circuitos en metric_policies (ej. YPF).",
    },
    "fill_rate_slots": {
        "tipo": "ratio", "requiere_periodo": True, "requiere_medio": "Digital",
        "requiere_cobertura_completa": False,
        "descripcion": "slots_ocupados/slots_comerciales. PARTIAL si se excluyen elementos con capacidad "
                        "REQUIERE_CONFIRMACION o con CompletitudMaestro!=COMPLETO. NO_APLICA para circuitos en metric_policies (ej. YPF).",
    },
    "segundos_comerciales": {
        "tipo": "suma", "requiere_periodo": True, "requiere_medio": "Digital",
        "requiere_cobertura_completa": False, "descripcion": "Suma de SegundosComerciales (default 72.000) de los elementos en scope.",
    },
    "segundos_vendidos": {
        "tipo": "suma", "requiere_periodo": True, "requiere_medio": "Digital",
        "requiere_cobertura_completa": False,
        "descripcion": "Promedio diario de SalidasVendidas*1800 por ElementoID, sumado en el grupo. "
                        "REQUIERE_CONFIRMACION si hay SalidasVendidas vacio en el periodo (nunca se asume 0). "
                        "Lanza UnsupportedBusinessCaseError si hay spots!=10s o exclusividad. "
                        "NO_APLICA para circuitos en metric_policies (ej. YPF).",
    },
    "segundos_disponibles": {
        "tipo": "suma", "requiere_periodo": True, "requiere_medio": "Digital",
        "requiere_cobertura_completa": False,
        "descripcion": "segundos_comerciales - segundos_vendidos. NO_APLICA para circuitos en metric_policies (ej. YPF).",
    },
    "fill_rate_segundos": {
        "tipo": "ratio", "requiere_periodo": True, "requiere_medio": "Digital",
        "requiere_cobertura_completa": False,
        "descripcion": "segundos_vendidos/segundos_comerciales. NO_APLICA para circuitos en metric_policies (ej. YPF).",
    },
}

_COUNT_CAMPANAS_ESTADO = {
    "campanas": None,
    "campanas_activas": "Activa",
    "campanas_reservadas": "Reservada",
}
_DISTINCT_CAMPANAS_FIELD = {
    "clientes_activos": "Cliente",
    "marcas_activas": "Marca",
}
_STATIC_METRICS = {
    "dias_periodo",
    "elemento_dias_disponibles",
    "elemento_dias_ocupados",
    "ocupacion_calendario_pct",
}
_DIGITAL_METRICS = {
    "slots_comerciales",
    "slots_ocupados",
    "slots_disponibles",
    "fill_rate_slots",
    "segundos_comerciales",
    "segundos_vendidos",
    "segundos_disponibles",
    "fill_rate_segundos",
}

_TIDY_TAIL_COLUMNS = ["Metric", "Value", "Numerator", "Denominator", "MetricStatus", "Warnings"]


def _apply_filters(df: pd.DataFrame, filters: dict[str, Any]) -> pd.DataFrame:
    if not filters:
        return df
    mask = pd.Series(True, index=df.index)
    for field, value in filters.items():
        if field not in df.columns:
            raise UnknownDimensionError(
                f"Filtro sobre dimension desconocida: {field!r}. Columnas validas: {sorted(df.columns)}"
            )
        if isinstance(value, (list, tuple, set)):
            mask &= df[field].isin(list(value))
        else:
            mask &= df[field] == value
    return df[mask]


def _iter_groups(df: pd.DataFrame, dims: list[str]) -> Iterator[tuple[dict[str, Any], pd.DataFrame]]:
    if not dims:
        yield {}, df
        return
    missing = [d for d in dims if d not in df.columns]
    if missing:
        raise UnknownDimensionError(
            f"group_by con dimension(es) desconocida(s): {missing}. Columnas validas: {sorted(df.columns)}"
        )
    if df.empty:
        return
    for key, sub in df.groupby(dims, dropna=False):
        key_t = key if isinstance(key, tuple) else (key,)
        yield dict(zip(dims, key_t)), sub


def _finalize_tidy(rows: list[dict[str, Any]], dims: Iterable[str], metric_name: str) -> pd.DataFrame:
    dims = list(dims)
    cols = dims + _TIDY_TAIL_COLUMNS
    if not rows:
        return pd.DataFrame(columns=cols)
    df = pd.DataFrame(rows)
    df["Metric"] = metric_name
    if "Numerator" not in df.columns:
        df["Numerator"] = df["Value"]
    if "Denominator" not in df.columns:
        df["Denominator"] = pd.NA
    if "MetricStatus" not in df.columns:
        df["MetricStatus"] = METRIC_STATUS_OK
    if "Warnings" not in df.columns:
        df["Warnings"] = ""
    return df[cols]


def _explode_days(scope: pd.DataFrame, extra_cols: list[str] | None = None) -> pd.DataFrame:
    """Expande cada fila [_eff_start, _eff_end] a una fila por dia calendario.
    scope debe tener columnas ElementoID, _eff_start, _eff_end (+ extra_cols)."""
    cols = ["ElementoID"] + (extra_cols or [])
    if scope.empty:
        return pd.DataFrame(columns=cols + ["_day"])
    days_series = [
        pd.date_range(s, e, freq="D") for s, e in zip(scope["_eff_start"], scope["_eff_end"])
    ]
    exploded = scope[cols].assign(_days=days_series).explode("_days")
    return exploded.rename(columns={"_days": "_day"})


class MetricsEngine:
    """Motor generico resolver(metric, group_by, filters, universe, periodo).

    Construido sobre el resultado de semantic_model.build_semantic_model();
    no lee el Excel ni recalcula metadata semantica.
    """

    def __init__(self, semantic_result: dict[str, Any]):
        self.maestro: pd.DataFrame = semantic_result["maestro"]
        self.campanas: pd.DataFrame = semantic_result["campanas"]
        self.config: dict[str, Any] = semantic_result["config"]

    # -- scoping helpers ----------------------------------------------------

    def _scoped_maestro(self, universe: str, filters: dict[str, Any], medio: str | None = None) -> pd.DataFrame:
        if universe not in self.config["universes"]:
            raise UnknownUniverseError(f"Universo desconocido: {universe!r}. Validos: {sorted(self.config['universes'])}")
        df = filter_universe(self.maestro, universe, self.config)
        if medio is not None:
            df = df[df["Medio"] == medio]
        return _apply_filters(df, filters)

    def _scoped_campanas(
        self,
        universe: str,
        filters: dict[str, Any],
        start_date: str | None = None,
        end_date: str | None = None,
    ) -> pd.DataFrame:
        if universe not in self.config["universes"]:
            raise UnknownUniverseError(f"Universo desconocido: {universe!r}. Validos: {sorted(self.config['universes'])}")
        df = filter_universe(self.campanas, universe, self.config)
        df = _apply_filters(df, filters)
        if start_date is not None or end_date is not None:
            if start_date is None or end_date is None:
                raise MetricsEngineError("start_date y end_date deben especificarse juntos")
            element_ids = df["ElementoID"].dropna().unique().tolist()
            overlap = self._campanas_overlap(element_ids, start_date, end_date)
            df = df[df["CargaID"].isin(overlap["CargaID"])]
        return df

    def _campanas_overlap(
        self,
        element_ids: list[Any],
        start_date: str,
        end_date: str,
        exclude_no_aplica: bool = False,
    ) -> pd.DataFrame:
        """CAMPANAS filtrado a element_ids, sin Cancelado, con solapamiento
        inclusivo contra [start_date, end_date]; agrega _eff_start/_eff_end.
        FechaIndefinida=Si + FechaFin vacia se trata como vigente hasta
        end_date (no se persiste una fecha final)."""
        start_ts, end_ts = pd.Timestamp(start_date), pd.Timestamp(end_date)
        if end_ts < start_ts:
            raise MetricsEngineError("end_date debe ser >= start_date")

        scope = self.campanas[self.campanas["ElementoID"].isin(element_ids)].copy()
        scope = scope[scope["Estado"] != "Cancelado"]
        if exclude_no_aplica:
            scope = scope[scope["ModalidadPauta"] != "No aplica"]
        if scope.empty:
            return scope.assign(_eff_start=pd.NaT, _eff_end=pd.NaT)

        fecha_inicio = scope["FechaInicio"]
        fecha_fin = scope["FechaFin"]
        indefinida_mask = (scope["FechaIndefinida"] == "Si") & fecha_fin.isna()
        effective_fin = fecha_fin.where(~indefinida_mask, end_ts)

        eff_start = fecha_inicio.clip(lower=start_ts)
        eff_end = effective_fin.clip(upper=end_ts)
        overlap_mask = fecha_inicio.notna() & effective_fin.notna() & (eff_start <= eff_end)

        scope = scope.loc[overlap_mask].copy()
        scope["_eff_start"] = eff_start.loc[overlap_mask]
        scope["_eff_end"] = eff_end.loc[overlap_mask]
        return scope

    def _incomplete_date_elements(self, element_ids: list[Any], exclude_no_aplica: bool = False) -> set[Any]:
        """ElementoID con campanas (Estado != Cancelado) cuya fecha no permite
        asignarlas de forma confiable a ningun periodo: FechaInicio vacia, o
        FechaFin vacia sin FechaIndefinida=Si (Gate3B.1 Sec.4). No se inventa
        fecha ni se bloquea el motor entero: solo se reporta el ElementoID
        afectado para que el llamador marque PARTIAL con warning."""
        scope = self.campanas[self.campanas["ElementoID"].isin(element_ids)]
        scope = scope[scope["Estado"] != "Cancelado"]
        if exclude_no_aplica:
            scope = scope[scope["ModalidadPauta"] != "No aplica"]
        if scope.empty:
            return set()
        indefinida_ok = (scope["FechaIndefinida"] == "Si") & scope["FechaFin"].isna()
        incomplete_mask = scope["FechaInicio"].isna() | (scope["FechaFin"].isna() & ~indefinida_ok)
        return set(scope.loc[incomplete_mask, "ElementoID"].unique())

    def _element_day_stats(
        self, scope_df: pd.DataFrame, group_by: list[str], start_date: str, end_date: str
    ) -> tuple[pd.DataFrame, int]:
        """Base compartida de ocupacion estatica/actividad calendario: union de
        dias calendario inclusivos ocupados por ElementoID (nunca duplica un
        dia solapado) mas el flag de fechas incompletas. Reutilizada por
        _compute_static_metric (Medio=Estatico) y _compute_actividad_sobre_registrados
        (cualquier Medio)."""
        start_ts, end_ts = pd.Timestamp(start_date), pd.Timestamp(end_date)
        if end_ts < start_ts:
            raise MetricsEngineError("end_date debe ser >= start_date")
        n_dias = (end_ts - start_ts).days + 1

        base_cols = {"ElementoID", "CoberturaCatalogo", "CompletitudMaestro"}
        needed_cols = list(base_cols) + [d for d in group_by if d not in base_cols]
        per_element = scope_df[needed_cols].drop_duplicates(subset=["ElementoID"]).copy()
        per_element["dias_disponibles"] = n_dias

        if per_element.empty:
            per_element["dias_ocupados"] = pd.Series(dtype="int64")
            per_element["_fecha_incompleta"] = pd.Series(dtype="bool")
            return per_element, n_dias

        element_ids = per_element["ElementoID"].tolist()
        overlap = self._campanas_overlap(element_ids, start_date, end_date)
        exploded = _explode_days(overlap)
        occupied = exploded.groupby("ElementoID")["_day"].nunique() if len(exploded) else pd.Series(dtype="int64")
        per_element["dias_ocupados"] = per_element["ElementoID"].map(occupied).fillna(0).astype(int)

        incomplete_ids = self._incomplete_date_elements(element_ids)
        per_element["_fecha_incompleta"] = per_element["ElementoID"].isin(incomplete_ids)

        return per_element, n_dias

    # -- API publica ----------------------------------------------------

    def query(
        self,
        metric: str,
        group_by: Iterable[str] = (),
        filters: dict[str, Any] | None = None,
        universe: str = "OPERATIVO_GENERAL",
        start_date: str | None = None,
        end_date: str | None = None,
    ) -> pd.DataFrame:
        filters = dict(filters or {})
        group_by = list(group_by)

        if metric == "elementos_registrados":
            return self._count_elements(group_by, filters, universe, None, metric)
        if metric == "elementos_confirmados":
            return self._count_elements(
                group_by, filters, universe, lambda df: df["CertezaDato"] == "CONFIRMADO", metric
            )
        if metric == "elementos_con_actividad":
            return self._count_elementos_con_actividad(group_by, filters, universe, start_date, end_date)
        if metric in _COUNT_CAMPANAS_ESTADO:
            return self._count_campanas(
                group_by, filters, universe, start_date, end_date, _COUNT_CAMPANAS_ESTADO[metric], metric
            )
        if metric in _DISTINCT_CAMPANAS_FIELD:
            return self._distinct_count(
                _DISTINCT_CAMPANAS_FIELD[metric], group_by, filters, universe, start_date, end_date, metric
            )
        if metric == "actividad_sobre_registrados_pct":
            return self._compute_actividad_sobre_registrados(group_by, filters, universe, start_date, end_date)
        if metric in _STATIC_METRICS:
            return self._compute_static_metric(metric, group_by, filters, universe, start_date, end_date)
        if metric in _DIGITAL_METRICS:
            return self._compute_digital_metric(metric, group_by, filters, universe, start_date, end_date)

        raise UnknownMetricError(f"Metrica desconocida: {metric!r}. Ver metrics_engine.METRIC_CATALOG.")

    # -- familia: conteos sobre MAESTRO_ELEMENTOS ----------------------------

    def _count_elements(
        self,
        group_by: list[str],
        filters: dict[str, Any],
        universe: str,
        mask_fn: Callable[[pd.DataFrame], pd.Series] | None,
        metric_name: str,
    ) -> pd.DataFrame:
        df = self._scoped_maestro(universe, filters)
        if mask_fn is not None:
            df = df[mask_fn(df)]
        rows = [{**key, "Value": sub["ElementoID"].nunique()} for key, sub in _iter_groups(df, group_by)]
        return _finalize_tidy(rows, group_by, metric_name)

    def _count_elementos_con_actividad(
        self,
        group_by: list[str],
        filters: dict[str, Any],
        universe: str,
        start_date: str | None,
        end_date: str | None,
    ) -> pd.DataFrame:
        df = self._scoped_maestro(universe, filters)
        if start_date is not None or end_date is not None:
            if start_date is None or end_date is None:
                raise MetricsEngineError("start_date y end_date deben especificarse juntos")
            active_ids = set(
                self._campanas_overlap(df["ElementoID"].tolist(), start_date, end_date)["ElementoID"].unique()
            )
            df = df[df["ElementoID"].isin(active_ids)]
        else:
            df = df[df["TieneActividadComercial"]]
        rows = [{**key, "Value": sub["ElementoID"].nunique()} for key, sub in _iter_groups(df, group_by)]
        return _finalize_tidy(rows, group_by, "elementos_con_actividad")

    # -- familia: conteos sobre CAMPANAS ----------------------------

    def _count_campanas(
        self,
        group_by: list[str],
        filters: dict[str, Any],
        universe: str,
        start_date: str | None,
        end_date: str | None,
        estado: str | None,
        metric_name: str,
    ) -> pd.DataFrame:
        df = self._scoped_campanas(universe, filters, start_date, end_date)
        if estado is not None:
            df = df[df["Estado"] == estado]
        rows = [{**key, "Value": len(sub)} for key, sub in _iter_groups(df, group_by)]
        return _finalize_tidy(rows, group_by, metric_name)

    def _distinct_count(
        self,
        field: str,
        group_by: list[str],
        filters: dict[str, Any],
        universe: str,
        start_date: str | None,
        end_date: str | None,
        metric_name: str,
    ) -> pd.DataFrame:
        df = self._scoped_campanas(universe, filters, start_date, end_date)
        df = df[df[field].notna() & (df[field].astype(str).str.strip() != "")]
        rows = [{**key, "Value": sub[field].nunique()} for key, sub in _iter_groups(df, group_by)]
        return _finalize_tidy(rows, group_by, metric_name)

    # -- familia: ocupacion estatica ----------------------------

    def _compute_static_metric(
        self,
        metric: str,
        group_by: list[str],
        filters: dict[str, Any],
        universe: str,
        start_date: str | None,
        end_date: str | None,
    ) -> pd.DataFrame:
        if start_date is None or end_date is None:
            raise MetricsEngineError(f"La metrica '{metric}' requiere start_date y end_date")

        scope_df = self._scoped_maestro(universe, filters, medio="Estático")
        missing = [c for c in group_by if c not in scope_df.columns]
        if missing:
            raise UnknownDimensionError(f"group_by con dimension(es) desconocida(s): {missing}")

        per_element, n_dias = self._element_day_stats(scope_df, group_by, start_date, end_date)

        rows = []
        for key, sub in _iter_groups(per_element, group_by):
            disponibles = int(sub["dias_disponibles"].sum())
            ocupados = int(sub["dias_ocupados"].sum())
            cobertura_y_completitud_ok = bool(
                len(sub)
                and (sub["CoberturaCatalogo"] == "COMPLETO").all()
                and (sub["CompletitudMaestro"] == "COMPLETO").all()
            )
            fecha_reason = (
                (METRIC_STATUS_PARTIAL, _FECHA_INCOMPLETA_WARNING)
                if len(sub) and sub["_fecha_incompleta"].any()
                else None
            )

            if metric == "dias_periodo":
                rows.append({**key, "Value": n_dias})
            elif metric == "elemento_dias_disponibles":
                rows.append({**key, "Value": disponibles})
            elif metric == "elemento_dias_ocupados":
                status, warning = _combine_status([fecha_reason])
                rows.append({**key, "Value": ocupados, "MetricStatus": status, "Warnings": warning})
            elif metric == "ocupacion_calendario_pct":
                if disponibles == 0:
                    rows.append({
                        **key, "Value": pd.NA, "Numerator": ocupados, "Denominator": disponibles,
                        "MetricStatus": METRIC_STATUS_NO_APLICA, "Warnings": "Sin elementos en el universo/filtro",
                    })
                elif not cobertura_y_completitud_ok:
                    status, warning = _combine_status([
                        (METRIC_STATUS_NO_APLICA,
                         "CoberturaCatalogo y/o CompletitudMaestro != COMPLETO para uno o mas elementos del grupo "
                         "(universo comercial conocido pero maestro incompleto, o universo desconocido): "
                         "ocupacion total no aplica, usar actividad_sobre_registrados_pct"),
                        fecha_reason,
                    ])
                    rows.append({
                        **key, "Value": pd.NA, "Numerator": ocupados, "Denominator": disponibles,
                        "MetricStatus": status, "Warnings": warning,
                    })
                else:
                    status, warning = _combine_status([fecha_reason])
                    rows.append({
                        **key, "Value": ocupados / disponibles * 100.0, "Numerator": ocupados, "Denominator": disponibles,
                        "MetricStatus": status, "Warnings": warning,
                    })

        return _finalize_tidy(rows, group_by, metric)

    def _compute_actividad_sobre_registrados(
        self,
        group_by: list[str],
        filters: dict[str, Any],
        universe: str,
        start_date: str | None,
        end_date: str | None,
    ) -> pd.DataFrame:
        """Actividad calendario sobre elementos REGISTRADOS (dias con campana /
        dias registrados), sin exigir CoberturaCatalogo/CompletitudMaestro
        completos y sin filtrar por Medio: funciona tanto para circuitos
        estaticos con universo completo como para circuitos con cobertura
        desconocida o CMS de terceros (ej. YPF Digital, Gate3B.1 Sec.2-3)."""
        if start_date is None or end_date is None:
            raise MetricsEngineError("La metrica 'actividad_sobre_registrados_pct' requiere start_date y end_date")

        scope_df = self._scoped_maestro(universe, filters)
        missing = [c for c in group_by if c not in scope_df.columns]
        if missing:
            raise UnknownDimensionError(f"group_by con dimension(es) desconocida(s): {missing}")

        per_element, _n_dias = self._element_day_stats(scope_df, group_by, start_date, end_date)

        rows = []
        for key, sub in _iter_groups(per_element, group_by):
            disponibles = int(sub["dias_disponibles"].sum())
            ocupados = int(sub["dias_ocupados"].sum())
            if disponibles == 0:
                rows.append({
                    **key, "Value": pd.NA, "Numerator": ocupados, "Denominator": disponibles,
                    "MetricStatus": METRIC_STATUS_NO_APLICA, "Warnings": "Sin elementos en el universo/filtro",
                })
                continue
            fecha_reason = (
                (METRIC_STATUS_PARTIAL, _FECHA_INCOMPLETA_WARNING)
                if len(sub) and sub["_fecha_incompleta"].any()
                else None
            )
            status, warning = _combine_status([fecha_reason])
            rows.append({
                **key, "Value": ocupados / disponibles * 100.0, "Numerator": ocupados, "Denominator": disponibles,
                "MetricStatus": status, "Warnings": warning,
            })

        return _finalize_tidy(rows, group_by, "actividad_sobre_registrados_pct")

    # -- familia: capacidad digital ----------------------------

    def _digital_period_activity(
        self, element_ids: list[Any], start_date: str, end_date: str
    ) -> tuple[pd.Series, pd.Series, set[Any], set[Any]]:
        """Devuelve (slots_ocupados_promedio_diario, segundos_vendidos_promedio_diario,
        elementos_con_fecha_incompleta, elementos_con_salidas_indeterminadas) por
        ElementoID. Lanza UnsupportedBusinessCaseError ante exclusividad o spots
        de duracion no soportada (Gate3A Sec.24-26): nunca inventa la formula.
        SalidasVendidas vacia NUNCA se trata como 0 (Gate3B.1 Sec.5): se excluye
        de la suma de segundos y el ElementoID se reporta para que el llamador
        decida el MetricStatus (nunca se estima silenciosamente)."""
        cfg = self.config["digital_capacity"]
        spot_base = cfg["spot_duracion_base_seg"]
        seg_por_salida = cfg["spot_segundos_por_salida"]

        incomplete_ids = self._incomplete_date_elements(element_ids, exclude_no_aplica=True)

        scope = self._campanas_overlap(element_ids, start_date, end_date, exclude_no_aplica=True)
        if scope.empty:
            return pd.Series(dtype="float64"), pd.Series(dtype="float64"), incomplete_ids, set()

        exclusividad_mask = scope["TipoExclusividad"].notna() | scope["ModalidadPauta"].astype(str).str.startswith(
            "Exclusividad"
        )
        if exclusividad_mask.any():
            offenders = sorted(scope.loc[exclusividad_mask, "ElementoID"].unique().tolist())[:10]
            raise UnsupportedBusinessCaseError(
                f"{int(exclusividad_mask.sum())} fila(s) de CAMPANAS con exclusividad (TipoExclusividad/"
                f"ModalidadPauta) dentro del periodo consultado. Gate 3B no calcula matematica de "
                f"exclusividad todavia (Gate3A Sec.26). ElementoID afectados (muestra): {offenders}"
            )

        spot_invalid_mask = scope["DuracionSpotSeg"].notna() & (scope["DuracionSpotSeg"] != spot_base)
        if spot_invalid_mask.any():
            offenders = (
                scope.loc[spot_invalid_mask, ["ElementoID", "DuracionSpotSeg"]]
                .drop_duplicates()
                .head(10)
                .to_dict("records")
            )
            raise UnsupportedBusinessCaseError(
                f"Se encontraron spot(s) con DuracionSpotSeg != {spot_base}s dentro del periodo consultado. "
                f"Gate 3B no calcula el escalado de segundos para duraciones distintas a {spot_base}s "
                f"todavia (Gate3A Sec.25). Filas afectadas (muestra): {offenders}"
            )

        salidas_indeterminadas_ids = set(scope.loc[scope["SalidasVendidas"].isna(), "ElementoID"].unique())

        exploded = _explode_days(scope, extra_cols=["SalidasVendidas"])
        if exploded.empty:
            return pd.Series(dtype="float64"), pd.Series(dtype="float64"), incomplete_ids, salidas_indeterminadas_ids

        daily = exploded.groupby(["ElementoID", "_day"]).agg(
            slots=("ElementoID", "size"),
            # SalidasVendidas vacia se excluye de la suma (dropna), nunca se
            # trata como 0: el llamador ve el ElementoID en salidas_indeterminadas_ids.
            segundos=("SalidasVendidas", lambda s: float((s.dropna() * seg_por_salida).sum())),
        )

        n_dias = (pd.Timestamp(end_date) - pd.Timestamp(start_date)).days + 1
        per_element = daily.groupby("ElementoID").agg(total_slots=("slots", "sum"), total_segundos=("segundos", "sum"))
        return (
            per_element["total_slots"] / n_dias,
            per_element["total_segundos"] / n_dias,
            incomplete_ids,
            salidas_indeterminadas_ids,
        )

    def _compute_digital_metric(
        self,
        metric: str,
        group_by: list[str],
        filters: dict[str, Any],
        universe: str,
        start_date: str | None,
        end_date: str | None,
    ) -> pd.DataFrame:
        if start_date is None or end_date is None:
            raise MetricsEngineError(f"La metrica '{metric}' requiere start_date y end_date")

        scope_df = self._scoped_maestro(universe, filters, medio="Digital")
        missing = [c for c in group_by if c not in scope_df.columns]
        if missing:
            raise UnknownDimensionError(f"group_by con dimension(es) desconocida(s): {missing}")

        sentinel = self.config["digital_capacity"]["requiere_confirmacion_sentinel"]
        policy_cfg = self.config["metric_policies"]
        policy_circuitos = set(policy_cfg["slot_seconds_no_aplica_circuitos"])
        policy_metrics = set(policy_cfg["slot_seconds_no_aplica_metrics"])
        policy_warning = policy_cfg["slot_seconds_no_aplica_warning"]

        base_cols = {"ElementoID", "SlotsComerciales", "SegundosComerciales", "CompletitudMaestro", "CircuitoNegocio"}
        cols = list(base_cols) + [d for d in group_by if d not in base_cols]
        per_element = scope_df[cols].drop_duplicates(subset=["ElementoID"]).copy()
        per_element["_policy_bloqueada"] = per_element["CircuitoNegocio"].isin(policy_circuitos)

        # Gate3B.1.1: la capacidad pura (slots_comerciales/segundos_comerciales)
        # no depende de campanas -> nunca se inspecciona actividad para estas
        # dos metricas. Para el resto, los ElementoID bloqueados por politica
        # (ej. YPF) se excluyen ANTES de mirar sus campanas: una campana con
        # exclusividad o DuracionSpotSeg!=10 en un elemento ya excluido nunca
        # debe poder romper el calculo del resto del grupo.
        is_pure_capacity_metric = metric in ("slots_comerciales", "segundos_comerciales")
        policy_applies_to_metric = metric in policy_metrics
        if is_pure_capacity_metric or per_element.empty:
            slots_s, segundos_s = pd.Series(dtype="float64"), pd.Series(dtype="float64")
            incomplete_ids, salidas_indeterminadas_ids = set(), set()
        else:
            activity_ids = (
                per_element.loc[~per_element["_policy_bloqueada"], "ElementoID"].tolist()
                if policy_applies_to_metric
                else per_element["ElementoID"].tolist()
            )
            if activity_ids:
                slots_s, segundos_s, incomplete_ids, salidas_indeterminadas_ids = self._digital_period_activity(
                    activity_ids, start_date, end_date
                )
            else:
                slots_s, segundos_s = pd.Series(dtype="float64"), pd.Series(dtype="float64")
                incomplete_ids, salidas_indeterminadas_ids = set(), set()

        per_element["slots_ocupados"] = per_element["ElementoID"].map(slots_s).fillna(0.0)
        per_element["segundos_vendidos"] = per_element["ElementoID"].map(segundos_s).fillna(0.0)
        per_element["_capacidad_desconocida"] = per_element["SlotsComerciales"] == sentinel
        per_element["_slots_conocidos"] = pd.to_numeric(
            per_element["SlotsComerciales"].where(~per_element["_capacidad_desconocida"]), errors="coerce"
        )
        per_element["_segundos_comerciales"] = pd.to_numeric(per_element["SegundosComerciales"], errors="coerce")
        per_element["_completitud_parcial"] = per_element["CompletitudMaestro"] != "COMPLETO"
        per_element["_fecha_incompleta"] = per_element["ElementoID"].isin(incomplete_ids)
        per_element["_salidas_indeterminadas"] = per_element["ElementoID"].isin(salidas_indeterminadas_ids)

        completitud_reason = (
            METRIC_STATUS_PARTIAL,
            "Resultado calculado sobre elementos registrados; el maestro del circuito esta incompleto.",
        )
        fecha_reason_tmpl = (METRIC_STATUS_PARTIAL, _FECHA_INCOMPLETA_WARNING)
        salidas_reason_tmpl = (
            METRIC_STATUS_REQUIERE_CONFIRMACION,
            "Existen campanas con SalidasVendidas vacio en el periodo; no se asume 0 segundos vendidos, no se estima.",
        )

        rows = []
        for key, sub in _iter_groups(per_element, group_by):
            policy_applies = metric in policy_metrics
            blocked_mask = sub["_policy_bloqueada"] if policy_applies else pd.Series(False, index=sub.index)
            n_blocked = int(blocked_mask.sum())

            if policy_applies and len(sub) and n_blocked == len(sub):
                rows.append({
                    **key, "Value": pd.NA, "Numerator": pd.NA, "Denominator": pd.NA,
                    "MetricStatus": METRIC_STATUS_NO_APLICA, "Warnings": policy_warning,
                })
                continue

            usable = sub[~blocked_mask] if policy_applies else sub
            policy_reason = (
                (METRIC_STATUS_PARTIAL, f"{n_blocked} elemento(s) excluido(s) por politica de circuito: {policy_warning}")
                if policy_applies and n_blocked
                else None
            )
            completitud_present = completitud_reason if len(usable) and usable["_completitud_parcial"].any() else None
            fecha_reason = fecha_reason_tmpl if len(usable) and usable["_fecha_incompleta"].any() else None
            salidas_reason = salidas_reason_tmpl if len(usable) and usable["_salidas_indeterminadas"].any() else None

            n_confirm = int(usable["_capacidad_desconocida"].sum())
            known_capacity = usable[~usable["_capacidad_desconocida"]]
            slots_conocidos_sum = float(known_capacity["_slots_conocidos"].fillna(0).sum())
            slots_ocupados_conocidos_sum = float(known_capacity["slots_ocupados"].sum())
            slots_ocupados_total_sum = float(usable["slots_ocupados"].sum())
            segundos_comerciales_sum = float(usable["_segundos_comerciales"].fillna(0).sum())
            segundos_vendidos_sum = float(usable["segundos_vendidos"].sum())
            capacidad_reason = (
                (METRIC_STATUS_PARTIAL, f"{n_confirm} elemento(s) con SlotsComerciales=REQUIERE_CONFIRMACION excluido(s) del calculo")
                if n_confirm
                else None
            )

            if metric == "slots_comerciales":
                status, warning = _combine_status([policy_reason, completitud_present, capacidad_reason])
                rows.append({**key, "Value": slots_conocidos_sum, "Numerator": slots_conocidos_sum, "MetricStatus": status, "Warnings": warning})

            elif metric == "slots_ocupados":
                status, warning = _combine_status([policy_reason, completitud_present, fecha_reason])
                rows.append({**key, "Value": slots_ocupados_total_sum, "Numerator": slots_ocupados_total_sum, "MetricStatus": status, "Warnings": warning})

            elif metric == "slots_disponibles":
                status, warning = _combine_status([policy_reason, completitud_present, capacidad_reason, fecha_reason])
                value = slots_conocidos_sum - slots_ocupados_conocidos_sum
                rows.append({**key, "Value": value, "Numerator": value, "MetricStatus": status, "Warnings": warning})

            elif metric == "fill_rate_slots":
                if slots_conocidos_sum <= 0:
                    status, warning = _combine_status([
                        (METRIC_STATUS_NO_APLICA, "Sin capacidad de slots conocida en el grupo"),
                        policy_reason, completitud_present, fecha_reason,
                    ])
                    rows.append({
                        **key, "Value": pd.NA, "Numerator": slots_ocupados_conocidos_sum, "Denominator": slots_conocidos_sum,
                        "MetricStatus": status, "Warnings": warning,
                    })
                else:
                    status, warning = _combine_status([policy_reason, completitud_present, capacidad_reason, fecha_reason])
                    rows.append({
                        **key, "Value": slots_ocupados_conocidos_sum / slots_conocidos_sum * 100.0,
                        "Numerator": slots_ocupados_conocidos_sum, "Denominator": slots_conocidos_sum,
                        "MetricStatus": status, "Warnings": warning,
                    })

            elif metric == "segundos_comerciales":
                status, warning = _combine_status([completitud_present])
                rows.append({**key, "Value": segundos_comerciales_sum, "Numerator": segundos_comerciales_sum, "MetricStatus": status, "Warnings": warning})

            elif metric == "segundos_vendidos":
                status, warning = _combine_status([policy_reason, completitud_present, fecha_reason, salidas_reason])
                if status == METRIC_STATUS_REQUIERE_CONFIRMACION:
                    rows.append({**key, "Value": pd.NA, "Numerator": pd.NA, "MetricStatus": status, "Warnings": warning})
                else:
                    rows.append({**key, "Value": segundos_vendidos_sum, "Numerator": segundos_vendidos_sum, "MetricStatus": status, "Warnings": warning})

            elif metric == "segundos_disponibles":
                status, warning = _combine_status([policy_reason, completitud_present, fecha_reason, salidas_reason])
                if status == METRIC_STATUS_REQUIERE_CONFIRMACION:
                    rows.append({**key, "Value": pd.NA, "Numerator": pd.NA, "MetricStatus": status, "Warnings": warning})
                else:
                    value = segundos_comerciales_sum - segundos_vendidos_sum
                    rows.append({**key, "Value": value, "Numerator": value, "MetricStatus": status, "Warnings": warning})

            elif metric == "fill_rate_segundos":
                if segundos_comerciales_sum <= 0:
                    status, warning = _combine_status([
                        (METRIC_STATUS_NO_APLICA, "Sin capacidad de segundos conocida en el grupo"),
                        policy_reason, completitud_present, fecha_reason, salidas_reason,
                    ])
                    rows.append({
                        **key, "Value": pd.NA, "Numerator": segundos_vendidos_sum, "Denominator": segundos_comerciales_sum,
                        "MetricStatus": status, "Warnings": warning,
                    })
                else:
                    status, warning = _combine_status([policy_reason, completitud_present, fecha_reason, salidas_reason])
                    if status == METRIC_STATUS_REQUIERE_CONFIRMACION:
                        rows.append({
                            **key, "Value": pd.NA, "Numerator": pd.NA, "Denominator": segundos_comerciales_sum,
                            "MetricStatus": status, "Warnings": warning,
                        })
                    else:
                        rows.append({
                            **key, "Value": segundos_vendidos_sum / segundos_comerciales_sum * 100.0,
                            "Numerator": segundos_vendidos_sum, "Denominator": segundos_comerciales_sum,
                            "MetricStatus": status, "Warnings": warning,
                        })

        return _finalize_tidy(rows, group_by, metric)
