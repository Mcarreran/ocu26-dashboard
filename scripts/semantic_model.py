"""Capa semantica central del pipeline OCU26 - Gate 3B.

Se ejecuta DESPUES de scripts/transform_data.py (Gate 2) y ANTES del motor de
metricas (scripts/metrics_engine.py). Unica fuente de datos: transform_data().
No vuelve a leer el Excel ni reimplementa validaciones/normalizaciones ya
resueltas en Gate 1/Gate 2.

Responsabilidad de este modulo:
    1. Cargar y validar config/business_semantics.json (metadata de negocio).
    2. Resolver, por fila de MAESTRO_ELEMENTOS, la jerarquia determinista:
       override por ElementoID > CircuitoDashboard+Subcircuito >
       CircuitoDashboard > regla semantica/generica > default seguro.
    3. Enriquecer MAESTRO_ELEMENTOS con dimensiones semanticas nuevas sin
       tocar ni una sola columna original (CircuitoDashboard, Subcircuito,
       Ubicacion, Medio, TipoCatalogo, TipoInventario, etc. quedan intactas).
    4. Unir a CAMPANAS las dimensiones semanticas de circuito para filtrar
       comodamente sin joins manuales en cada consulta.

100% en memoria, 100% read-only sobre el Excel (no lo toca en absoluto: solo
consume el dict ya devuelto por transform_data()).

Uso:
    from transform_data import transform_data
    from semantic_model import build_semantic_model

    result = build_semantic_model(transform_data())
    maestro_enriquecido = result["maestro"]
"""

from __future__ import annotations

import json
from collections import Counter
from pathlib import Path
from typing import Any

import pandas as pd

DEFAULT_CONFIG_PATH = Path(__file__).resolve().parent.parent / "config" / "business_semantics.json"

REQUIRED_SCHEMA_VERSION = 1

_BUNDLE_FIELDS = [
    "circuito_negocio",
    "portfolio_tier",
    "incluye_performance_core",
    "incluye_conteo_general",
    "visible_por_defecto",
    "cobertura_catalogo",
    "completitud_maestro",
    "modo_disponibilidad",
    "certeza_dato_regla",
]

_BOOLEAN_FIELDS = {"incluye_performance_core", "incluye_conteo_general", "visible_por_defecto"}

ALLOWED_PORTFOLIO_TIERS = {"CORE", "COMPLEMENTARIO", "LEGACY", "NO_CLASIFICADO"}
ALLOWED_COBERTURA_CATALOGO = {"COMPLETO", "DESCONOCIDO"}
ALLOWED_COMPLETITUD_MAESTRO = {"COMPLETO", "PARCIAL", "NO_APLICA"}
ALLOWED_MODO_DISPONIBILIDAD = {"CALCULABLE", "CONSULTA", "MIXTO"}

_TOP_LEVEL_REQUIRED_KEYS = [
    "schema_version",
    "resolution_order",
    "defaults",
    "elemento_overrides",
    "circuito_subcircuito_rules",
    "circuito_dashboard_rules",
    "generic_rules",
    "certeza_dato",
    "sitio_negocio",
    "formato_negocio",
    "digital_capacity",
    "universes",
    "metric_policies",
]

_KNOWN_RESOLUTION_LEVELS = {
    "elemento_overrides",
    "circuito_subcircuito_rules",
    "circuito_dashboard_rules",
    "generic_rules",
    "defaults",
}

_CAPACITY_OVERRIDE_FIELDS = {"slots_comerciales", "segundos_comerciales"}

SEMANTIC_DIMENSION_COLUMNS = [
    "CircuitoNegocio",
    "SitioNegocio",
    "FormatoNegocio",
    "CoberturaCatalogo",
    "CompletitudMaestro",
    "CertezaDato",
    "ModoDisponibilidad",
    "PortfolioTier",
    "IncluyePerformanceCore",
    "IncluyeConteoGeneral",
    "VisiblePorDefecto",
    "TieneActividadComercial",
    "CantidadCampanasHistoricas",
    "FechaPrimeraCampana",
    "FechaUltimaCampana",
    "SlotsComerciales",
    "SegundosComerciales",
]


class SemanticModelError(Exception):
    """Error bloqueante al enriquecer el modelo semantico (invariante rota)."""


class SemanticConfigError(SemanticModelError):
    """config/business_semantics.json invalido o inconsistente."""


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    if value is pd.NA:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


# ---------------------------------------------------------------------------
# Carga y validacion de configuracion
# ---------------------------------------------------------------------------


def load_config(path: str | Path = DEFAULT_CONFIG_PATH) -> dict[str, Any]:
    """Carga config/business_semantics.json y la valida. Lanza SemanticConfigError
    si el archivo no existe, no es JSON valido o no pasa validate_config()."""
    path = Path(path)
    if not path.exists():
        raise SemanticConfigError(f"No existe el archivo de configuracion semantica: {path}")
    try:
        with open(path, encoding="utf-8") as fh:
            config = json.load(fh)
    except json.JSONDecodeError as exc:
        raise SemanticConfigError(f"{path}: JSON invalido: {exc}") from exc

    validate_config(config)
    return config


def validate_config(config: dict[str, Any]) -> None:
    """Valida schema_version, tipos, valores permitidos, reglas duplicadas y
    prioridades ambiguas. Lanza SemanticConfigError con TODOS los problemas
    encontrados (no se detiene en el primero) para facilitar la correccion."""
    errors: list[str] = []

    if config.get("schema_version") != REQUIRED_SCHEMA_VERSION:
        errors.append(
            f"schema_version debe ser {REQUIRED_SCHEMA_VERSION}, encontrado: {config.get('schema_version')!r}"
        )

    for key in _TOP_LEVEL_REQUIRED_KEYS:
        if key not in config:
            errors.append(f"Falta la clave requerida de nivel superior: '{key}'")

    if errors:
        # Sin estas claves no se puede seguir validando de forma segura.
        raise SemanticConfigError("Configuracion semantica invalida:\n- " + "\n- ".join(errors))

    _validate_resolution_order(config["resolution_order"], errors)
    _validate_defaults(config["defaults"], errors)
    _validate_rule_list(config["circuito_subcircuito_rules"], "circuito_subcircuito_rules", errors, keyed=True)
    _validate_rule_list(config["circuito_dashboard_rules"], "circuito_dashboard_rules", errors, keyed=False)
    _validate_generic_rules(config["generic_rules"], errors)
    _validate_elemento_overrides(config["elemento_overrides"], errors)
    _validate_ambiguity(config, errors)
    _validate_certeza_dato(config, errors)
    _validate_digital_capacity(config, errors)
    _validate_sitio_negocio(config["sitio_negocio"], errors)
    _validate_formato_negocio(config["formato_negocio"], errors)
    _validate_universes(config["universes"], errors)
    _validate_metric_policies(config, errors)

    if errors:
        raise SemanticConfigError("Configuracion semantica invalida:\n- " + "\n- ".join(errors))


def _validate_resolution_order(order: Any, errors: list[str]) -> None:
    """resolution_order debe ser exactamente una permutacion de los 5 niveles
    conocidos, sin duplicados ni niveles desconocidos, con 'defaults' al final
    (es el fallback incondicional: cualquier nivel despues de 'defaults' seria
    inalcanzable)."""
    if not isinstance(order, list):
        errors.append("resolution_order: debe ser una lista")
        return
    if len(set(order)) != len(order):
        errors.append(f"resolution_order: contiene niveles duplicados: {order}")
    unknown = [lvl for lvl in order if lvl not in _KNOWN_RESOLUTION_LEVELS]
    if unknown:
        errors.append(f"resolution_order: nivel(es) desconocido(s) {unknown}; validos: {sorted(_KNOWN_RESOLUTION_LEVELS)}")
    missing = _KNOWN_RESOLUTION_LEVELS - set(order)
    if missing:
        errors.append(f"resolution_order: falta(n) nivel(es) obligatorio(s): {sorted(missing)}")
    if order and order[-1] != "defaults":
        errors.append("resolution_order: 'defaults' debe ser el ultimo nivel (fallback incondicional)")


def _validate_bundle_fields(bundle: dict[str, Any], path: str, errors: list[str]) -> None:
    for field in _BUNDLE_FIELDS:
        if field not in bundle:
            errors.append(f"{path}: falta el campo requerido '{field}'")
            continue
        value = bundle[field]
        if field in _BOOLEAN_FIELDS and not isinstance(value, bool):
            errors.append(f"{path}.{field}: debe ser booleano, encontrado {value!r}")
        if field == "portfolio_tier" and value not in ALLOWED_PORTFOLIO_TIERS:
            errors.append(f"{path}.{field}: valor no permitido {value!r} (permitidos: {sorted(ALLOWED_PORTFOLIO_TIERS)})")
        if field == "cobertura_catalogo" and value not in ALLOWED_COBERTURA_CATALOGO:
            errors.append(f"{path}.{field}: valor no permitido {value!r} (permitidos: {sorted(ALLOWED_COBERTURA_CATALOGO)})")
        if field == "completitud_maestro" and value not in ALLOWED_COMPLETITUD_MAESTRO:
            errors.append(f"{path}.{field}: valor no permitido {value!r} (permitidos: {sorted(ALLOWED_COMPLETITUD_MAESTRO)})")
        if field == "modo_disponibilidad" and value not in ALLOWED_MODO_DISPONIBILIDAD:
            errors.append(f"{path}.{field}: valor no permitido {value!r} (permitidos: {sorted(ALLOWED_MODO_DISPONIBILIDAD)})")
        if field == "circuito_negocio" and not isinstance(value, str):
            errors.append(f"{path}.{field}: debe ser texto, encontrado {value!r}")


def _validate_defaults(defaults: dict[str, Any], errors: list[str]) -> None:
    _validate_bundle_fields(defaults, "defaults", errors)


def _validate_rule_list(rules: Any, key: str, errors: list[str], keyed: bool) -> None:
    if not isinstance(rules, list):
        errors.append(f"{key}: debe ser una lista")
        return
    for i, rule in enumerate(rules):
        path = f"{key}[{i}]"
        if not isinstance(rule, dict):
            errors.append(f"{path}: debe ser un objeto")
            continue
        cd = rule.get("circuito_dashboard")
        if not isinstance(cd, list) or not cd or not all(isinstance(v, str) for v in cd):
            errors.append(f"{path}.circuito_dashboard: debe ser una lista no vacia de strings")
        if keyed:
            sc = rule.get("subcircuito")
            if not isinstance(sc, list) or not sc or not all(isinstance(v, str) for v in sc):
                errors.append(f"{path}.subcircuito: debe ser una lista no vacia de strings")
        _validate_bundle_fields(rule, path, errors)


def _validate_generic_rules(rules: Any, errors: list[str]) -> None:
    if not isinstance(rules, list):
        errors.append("generic_rules: debe ser una lista")
        return
    seen_ids: set[str] = set()
    for i, rule in enumerate(rules):
        path = f"generic_rules[{i}]"
        if not isinstance(rule, dict):
            errors.append(f"{path}: debe ser un objeto")
            continue
        rule_id = rule.get("id")
        if not isinstance(rule_id, str) or not rule_id:
            errors.append(f"{path}.id: requerido")
        elif rule_id in seen_ids:
            errors.append(f"{path}.id: id duplicado '{rule_id}'")
        else:
            seen_ids.add(rule_id)
        when = rule.get("when")
        if not isinstance(when, dict) or "field" not in when or ("equals" not in when and "in" not in when):
            errors.append(f"{path}.when: debe tener 'field' y ('equals' o 'in')")
        _validate_bundle_fields(rule, path, errors)


def _validate_elemento_overrides(overrides: Any, errors: list[str]) -> None:
    if not isinstance(overrides, list):
        errors.append("elemento_overrides: debe ser una lista")
        return
    seen: set[str] = set()
    for i, rule in enumerate(overrides):
        path = f"elemento_overrides[{i}]"
        if not isinstance(rule, dict):
            errors.append(f"{path}: debe ser un objeto")
            continue
        eid = rule.get("elemento_id")
        if not isinstance(eid, str) or not eid:
            errors.append(f"{path}.elemento_id: requerido (string no vacio)")
        elif eid in seen:
            errors.append(f"{path}.elemento_id: ElementoID duplicado en overrides: '{eid}'")
        else:
            seen.add(eid)
        _validate_bundle_fields(rule, path, errors)
        _validate_capacity_override_fields(rule, path, errors)


def _validate_capacity_override_fields(rule: dict[str, Any], path: str, errors: list[str]) -> None:
    """slots_comerciales/segundos_comerciales son OPCIONALES dentro de un
    elemento_override: permiten corregir la capacidad comercial de un
    ElementoID puntual sin tocar CapacidadSlotsReel/SegundosDia originales
    ni el perfil de FormatoNegocio."""
    if "slots_comerciales" in rule:
        value = rule["slots_comerciales"]
        if not isinstance(value, int) or isinstance(value, bool) or value <= 0:
            errors.append(f"{path}.slots_comerciales: debe ser un entero positivo, encontrado {value!r}")
    if "segundos_comerciales" in rule:
        value = rule["segundos_comerciales"]
        if not isinstance(value, (int, float)) or isinstance(value, bool) or value <= 0:
            errors.append(f"{path}.segundos_comerciales: debe ser un numero positivo, encontrado {value!r}")


def _validate_ambiguity(config: dict[str, Any], errors: list[str]) -> None:
    """Detecta reglas ambiguas dentro del mismo nivel jerarquico: dos reglas
    de circuito_subcircuito_rules que puedan matchear la misma combinacion
    (CircuitoDashboard, Subcircuito), o dos de circuito_dashboard_rules que
    compartan un mismo CircuitoDashboard."""
    seen_pairs: dict[tuple[str, str], int] = {}
    for i, rule in enumerate(config.get("circuito_subcircuito_rules", [])):
        cds = rule.get("circuito_dashboard") or []
        scs = rule.get("subcircuito") or []
        for cd in cds:
            for sc in scs:
                pair = (cd, sc)
                if pair in seen_pairs:
                    errors.append(
                        f"circuito_subcircuito_rules: combinacion ambigua {pair} en las reglas "
                        f"[{seen_pairs[pair]}] y [{i}]"
                    )
                else:
                    seen_pairs[pair] = i

    seen_cd: dict[str, int] = {}
    for i, rule in enumerate(config.get("circuito_dashboard_rules", [])):
        for cd in rule.get("circuito_dashboard") or []:
            if cd in seen_cd:
                errors.append(
                    f"circuito_dashboard_rules: CircuitoDashboard ambiguo '{cd}' en las reglas "
                    f"[{seen_cd[cd]}] y [{i}]"
                )
            else:
                seen_cd[cd] = i


def _validate_certeza_dato(config: dict[str, Any], errors: list[str]) -> None:
    cd = config.get("certeza_dato")
    if not isinstance(cd, dict):
        errors.append("certeza_dato: debe ser un objeto")
        return
    outputs = cd.get("rule_outputs")
    if not isinstance(outputs, dict) or not outputs:
        errors.append("certeza_dato.rule_outputs: debe ser un objeto no vacio")
        outputs = {}
    if "technical_override_value" not in cd:
        errors.append("certeza_dato.technical_override_value: requerido")

    known_ids = set(outputs) | {"campana_confirma"}
    used_ids: set[str] = set()
    for section in ("circuito_subcircuito_rules", "circuito_dashboard_rules", "generic_rules"):
        for rule in config.get(section, []):
            if isinstance(rule, dict) and "certeza_dato_regla" in rule:
                used_ids.add(rule["certeza_dato_regla"])
    defaults = config.get("defaults", {})
    if isinstance(defaults, dict) and "certeza_dato_regla" in defaults:
        used_ids.add(defaults["certeza_dato_regla"])
    for rule in config.get("elemento_overrides", []):
        if isinstance(rule, dict) and "certeza_dato_regla" in rule:
            used_ids.add(rule["certeza_dato_regla"])

    for rid in used_ids:
        if rid not in known_ids:
            errors.append(
                f"certeza_dato: certeza_dato_regla '{rid}' usado en una regla de circuito no esta "
                f"definido en certeza_dato.rule_outputs (ni es 'campana_confirma')"
            )
    if "campana_confirma" in used_ids:
        for required in ("campana_confirma_con_actividad", "campana_confirma_sin_actividad"):
            if required not in outputs:
                errors.append(f"certeza_dato.rule_outputs: falta '{required}' (requerido por la regla 'campana_confirma')")


def _validate_digital_capacity(config: dict[str, Any], errors: list[str]) -> None:
    dc = config.get("digital_capacity")
    if not isinstance(dc, dict):
        errors.append("digital_capacity: debe ser un objeto")
        return
    for field in ("default_segundos_comerciales", "spot_duracion_base_seg", "spot_segundos_por_salida"):
        value = dc.get(field)
        if not isinstance(value, (int, float)) or isinstance(value, bool) or value <= 0:
            errors.append(f"digital_capacity.{field}: debe ser un numero positivo, encontrado {value!r}")
    profiles = dc.get("slots_profiles")
    if not isinstance(profiles, dict):
        errors.append("digital_capacity.slots_profiles: debe ser un objeto")
    else:
        for formato, slots in profiles.items():
            if not isinstance(slots, int) or isinstance(slots, bool) or slots <= 0:
                errors.append(f"digital_capacity.slots_profiles.{formato}: debe ser un entero positivo, encontrado {slots!r}")
    if not isinstance(dc.get("requiere_confirmacion_sentinel"), str):
        errors.append("digital_capacity.requiere_confirmacion_sentinel: requerido (string)")
    if not isinstance(dc.get("use_legacy_source_if_positive_for_circuitos"), list):
        errors.append("digital_capacity.use_legacy_source_if_positive_for_circuitos: debe ser una lista")


def _validate_sitio_negocio(sn: Any, errors: list[str]) -> None:
    if not isinstance(sn, dict):
        errors.append("sitio_negocio: debe ser un objeto")
        return
    if not isinstance(sn.get("default_field"), str):
        errors.append("sitio_negocio.default_field: requerido (string)")
    if not isinstance(sn.get("circuito_field_overrides"), dict):
        errors.append("sitio_negocio.circuito_field_overrides: debe ser un objeto")
    if not isinstance(sn.get("ubicacion_lookup"), dict):
        errors.append("sitio_negocio.ubicacion_lookup: debe ser un objeto")


def _validate_formato_negocio(fn: Any, errors: list[str]) -> None:
    if not isinstance(fn, dict):
        errors.append("formato_negocio: debe ser un objeto")
        return
    if not isinstance(fn.get("ypf_elemento_id_token_map"), dict):
        errors.append("formato_negocio.ypf_elemento_id_token_map: debe ser un objeto")
    if not isinstance(fn.get("circuito_dashboard_formato_fijo"), dict):
        errors.append("formato_negocio.circuito_dashboard_formato_fijo: debe ser un objeto")
    rules = fn.get("descripcion_keyword_rules")
    if not isinstance(rules, list):
        errors.append("formato_negocio.descripcion_keyword_rules: debe ser una lista")
    else:
        for i, rule in enumerate(rules):
            if not isinstance(rule, dict) or "contains" not in rule or "formato" not in rule:
                errors.append(f"formato_negocio.descripcion_keyword_rules[{i}]: requiere 'contains' y 'formato'")
    if not isinstance(fn.get("default"), str):
        errors.append("formato_negocio.default: requerido (string)")


_KNOWN_UNIVERSE_FLAGS = {"IncluyeConteoGeneral", "IncluyePerformanceCore", "VisiblePorDefecto", None}


def _validate_universes(universes: Any, errors: list[str]) -> None:
    if not isinstance(universes, dict) or not universes:
        errors.append("universes: debe ser un objeto no vacio")
        return
    for name, spec in universes.items():
        if name.startswith("_"):
            continue
        if not isinstance(spec, dict):
            errors.append(f"universes.{name}: debe ser un objeto")
            continue
        flag = spec.get("require_flag", "__missing__")
        if flag == "__missing__":
            errors.append(f"universes.{name}: falta 'require_flag' (puede ser null)")
        elif flag not in _KNOWN_UNIVERSE_FLAGS:
            errors.append(f"universes.{name}.require_flag: valor no reconocido {flag!r}")


def _validate_metric_policies(config: dict[str, Any], errors: list[str]) -> None:
    """Politicas declarativas de aplicabilidad (ej. YPF sin ocupacion de CMS
    conocida). No se valida contra la lista de metricas de metrics_engine.py
    a proposito: semantic_model.py no debe depender de metrics_engine.py."""
    mp = config.get("metric_policies")
    if not isinstance(mp, dict):
        errors.append("metric_policies: debe ser un objeto")
        return
    circuitos = mp.get("slot_seconds_no_aplica_circuitos")
    if not isinstance(circuitos, list) or not all(isinstance(v, str) for v in circuitos):
        errors.append("metric_policies.slot_seconds_no_aplica_circuitos: debe ser una lista de strings")
    metrics = mp.get("slot_seconds_no_aplica_metrics")
    if not isinstance(metrics, list) or not metrics or not all(isinstance(v, str) for v in metrics):
        errors.append("metric_policies.slot_seconds_no_aplica_metrics: debe ser una lista no vacia de strings")
    if not isinstance(mp.get("slot_seconds_no_aplica_warning"), str) or not mp.get("slot_seconds_no_aplica_warning"):
        errors.append("metric_policies.slot_seconds_no_aplica_warning: requerido (string no vacio)")


# ---------------------------------------------------------------------------
# Resolucion jerarquica por fila
# ---------------------------------------------------------------------------


def _bundle_from_rule(rule: dict[str, Any]) -> dict[str, Any]:
    return {field: rule[field] for field in _BUNDLE_FIELDS}


def _match_circuito_subcircuito(row: pd.Series, rules: list[dict[str, Any]]) -> dict[str, Any] | None:
    cd, sc = row["CircuitoDashboard"], row["Subcircuito"]
    for rule in rules:
        if cd in rule["circuito_dashboard"] and sc in rule["subcircuito"]:
            return rule
    return None


def _match_circuito_dashboard(row: pd.Series, rules: list[dict[str, Any]]) -> dict[str, Any] | None:
    cd = row["CircuitoDashboard"]
    for rule in rules:
        if cd in rule["circuito_dashboard"]:
            return rule
    return None


def _match_generic(row: pd.Series, rules: list[dict[str, Any]]) -> dict[str, Any] | None:
    for rule in rules:
        when = rule["when"]
        field_value = row.get(when["field"])
        if "equals" in when:
            if field_value == when["equals"]:
                return rule
        elif "in" in when:
            if field_value in when["in"]:
                return rule
    return None


def _try_elemento_override(
    row: pd.Series, config: dict[str, Any], elemento_override_map: dict[str, dict[str, Any]]
) -> tuple[dict[str, Any], str] | None:
    eid = row["ElementoID"]
    if not _is_blank(eid) and eid in elemento_override_map:
        return _bundle_from_rule(elemento_override_map[eid]), "elemento_override"
    return None


def _try_circuito_subcircuito(
    row: pd.Series, config: dict[str, Any], elemento_override_map: dict[str, dict[str, Any]]
) -> tuple[dict[str, Any], str] | None:
    rule = _match_circuito_subcircuito(row, config["circuito_subcircuito_rules"])
    return (_bundle_from_rule(rule), "circuito_subcircuito") if rule is not None else None


def _try_circuito_dashboard(
    row: pd.Series, config: dict[str, Any], elemento_override_map: dict[str, dict[str, Any]]
) -> tuple[dict[str, Any], str] | None:
    rule = _match_circuito_dashboard(row, config["circuito_dashboard_rules"])
    return (_bundle_from_rule(rule), "circuito_dashboard") if rule is not None else None


def _try_generic(
    row: pd.Series, config: dict[str, Any], elemento_override_map: dict[str, dict[str, Any]]
) -> tuple[dict[str, Any], str] | None:
    rule = _match_generic(row, config["generic_rules"])
    return (_bundle_from_rule(rule), "generic") if rule is not None else None


def _try_defaults(
    row: pd.Series, config: dict[str, Any], elemento_override_map: dict[str, dict[str, Any]]
) -> tuple[dict[str, Any], str] | None:
    return dict(config["defaults"]), "default"


# Un handler por nivel de resolution_order (Gate3B.1 Sec.6): el orden real de
# evaluacion lo decide config["resolution_order"], no el orden de este dict.
_LEVEL_HANDLERS = {
    "elemento_overrides": _try_elemento_override,
    "circuito_subcircuito_rules": _try_circuito_subcircuito,
    "circuito_dashboard_rules": _try_circuito_dashboard,
    "generic_rules": _try_generic,
    "defaults": _try_defaults,
}


def _resolve_circuit_bundle(
    row: pd.Series,
    elemento_override_map: dict[str, dict[str, Any]],
    config: dict[str, Any],
    warnings: list[str],
) -> tuple[dict[str, Any], str]:
    """Resuelve la jerarquia siguiendo literalmente config["resolution_order"]
    (validado por validate_config: permutacion de los 5 niveles conocidos,
    'defaults' al final). Devuelve (bundle_completo, nivel_que_matcheo) para
    trazabilidad/tests."""
    eid = row["ElementoID"]

    for level in config["resolution_order"]:
        result = _LEVEL_HANDLERS[level](row, config, elemento_override_map)
        if result is None:
            continue
        bundle, matched_level = result
        if matched_level == "default" and not _is_blank(eid):
            warnings.append(
                f"ElementoID={eid!r} (CircuitoDashboard={row['CircuitoDashboard']!r}, "
                f"Subcircuito={row['Subcircuito']!r}) no coincide con ninguna regla semantica; "
                f"usando default seguro (CircuitoNegocio={config['defaults']['circuito_negocio']}, "
                f"IncluyePerformanceCore=False)"
            )
        return bundle, matched_level

    # Solo alcanzable si resolution_order no incluye "defaults" (rechazado
    # por validate_config, pero se protege por si se llama con config sin
    # validar).
    raise SemanticModelError("resolution_order no incluye 'defaults': no se pudo resolver la fila")


def _resolve_formato_negocio(row: pd.Series, circuito_negocio: str, config: dict[str, Any], warnings: list[str]) -> str:
    fmt_cfg = config["formato_negocio"]

    if circuito_negocio == "YPF":
        eid = "" if _is_blank(row["ElementoID"]) else str(row["ElementoID"])
        parts = eid.split(" - ")
        token = parts[1].strip() if len(parts) >= 2 else None
        mapped = fmt_cfg["ypf_elemento_id_token_map"].get(token) if token else None
        if mapped:
            return mapped
        warnings.append(
            f"YPF ElementoID={row['ElementoID']!r}: token de formato no reconocido ({token!r}); "
            f"FormatoNegocio={fmt_cfg['default']}"
        )
        return fmt_cfg["default"]

    fixed = fmt_cfg["circuito_dashboard_formato_fijo"].get(row["CircuitoDashboard"])
    if fixed:
        return fixed

    desc = "" if _is_blank(row["Descripcion"]) else str(row["Descripcion"])
    desc_lower = desc.lower()
    for rule in fmt_cfg["descripcion_keyword_rules"]:
        if rule["contains"].lower() in desc_lower:
            return rule["formato"]

    return fmt_cfg["default"]


def _resolve_sitio_negocio(row: pd.Series, circuito_negocio: str, config: dict[str, Any]) -> str | None:
    sn_cfg = config["sitio_negocio"]
    field = sn_cfg["circuito_field_overrides"].get(circuito_negocio, sn_cfg["default_field"])
    raw = row.get(field)
    if _is_blank(raw):
        return None
    raw_str = str(raw)
    if field == "Ubicacion":
        return sn_cfg["ubicacion_lookup"].get(raw_str, raw_str)
    return raw_str


def _resolve_certeza_dato(
    row: pd.Series, certeza_dato_regla: str, tiene_actividad: bool, config: dict[str, Any]
) -> str:
    cd_cfg = config["certeza_dato"]
    outputs = cd_cfg["rule_outputs"]

    if certeza_dato_regla == "campana_confirma":
        key = "campana_confirma_con_actividad" if tiene_actividad else "campana_confirma_sin_actividad"
        base = outputs[key]
    else:
        base = outputs[certeza_dato_regla]

    if row["Medio"] == "Digital":
        legacy = row["CapacidadSlotsReel"]
        if not pd.isna(legacy) and int(legacy) == 0:
            return cd_cfg["technical_override_value"]

    return base


def _resolve_digital_capacity(
    row: pd.Series,
    circuito_negocio: str,
    formato_negocio: str,
    config: dict[str, Any],
    warnings: list[str],
    elemento_override_map: dict[str, dict[str, Any]],
) -> tuple[Any, Any]:
    """Devuelve (SlotsComerciales, SegundosComerciales). Solo aplica a Medio=Digital;
    para elementos estaticos ambos son pd.NA (no aplica capacidad de reel).

    Precedencia de slots (Gate3B.1 Sec.7): override ElementoID > perfil
    FormatoNegocio > fallback legacy permitido > REQUIERE_CONFIRMACION.
    Precedencia de segundos: override ElementoID > default comercial.
    CapacidadSlotsReel/SegundosDia originales nunca se leen ni se modifican
    fuera de este fallback de lectura."""
    if row["Medio"] != "Digital":
        return pd.NA, pd.NA

    dc_cfg = config["digital_capacity"]
    default_segundos = dc_cfg["default_segundos_comerciales"]
    sentinel = dc_cfg["requiere_confirmacion_sentinel"]

    eid = row["ElementoID"]
    override = elemento_override_map.get(eid) if not _is_blank(eid) else None

    if override is not None and "slots_comerciales" in override:
        slots = override["slots_comerciales"]
    else:
        profile_slots = dc_cfg["slots_profiles"].get(formato_negocio)
        if profile_slots is not None:
            slots = profile_slots
        else:
            legacy = row["CapacidadSlotsReel"]
            legacy_val = None if pd.isna(legacy) else int(legacy)
            if legacy_val and legacy_val > 0:
                if circuito_negocio not in dc_cfg["use_legacy_source_if_positive_for_circuitos"]:
                    warnings.append(
                        f"ElementoID={eid!r} ({circuito_negocio}/{formato_negocio}): sin perfil "
                        f"comercial de slots confirmado; se usa capacidad fuente legacy ({legacy_val}) como fallback"
                    )
                slots = legacy_val
            else:
                warnings.append(
                    f"ElementoID={eid!r} ({circuito_negocio}/{formato_negocio}): capacidad fuente no "
                    f"cargada (CapacidadSlotsReel={legacy_val}); SlotsComerciales={sentinel}"
                )
                slots = sentinel

    if override is not None and "segundos_comerciales" in override:
        segundos = override["segundos_comerciales"]
    else:
        segundos = default_segundos

    return slots, segundos


# ---------------------------------------------------------------------------
# Construccion del modelo semantico completo
# ---------------------------------------------------------------------------


def build_semantic_model(transform_result: dict[str, Any], config: dict[str, Any] | None = None) -> dict[str, Any]:
    """Enriquece la salida de transform_data() con la capa semantica de Gate 3B.

    No reabre el Excel ni reimplementa nada de Gate 1/Gate 2: opera
    exclusivamente sobre los DataFrames ya devueltos por transform_data().

    Devuelve dict con: maestro (enriquecido), campanas (enriquecido con join
    de dimensiones de circuito), parametros (passthrough), config, warnings,
    stats.
    """
    if config is None:
        config = load_config()
    else:
        validate_config(config)

    maestro_raw: pd.DataFrame = transform_result["maestro"]
    campanas_raw: pd.DataFrame = transform_result["campanas"]
    parametros_raw: pd.DataFrame = transform_result["parametros"]

    maestro = maestro_raw.copy(deep=True)
    campanas = campanas_raw.copy(deep=True)

    warnings: list[str] = list(transform_result.get("warnings", []))

    elemento_override_map = {o["elemento_id"]: o for o in config["elemento_overrides"]}

    bundles: list[dict[str, Any]] = []
    matched_levels: list[str] = []
    for _, row in maestro.iterrows():
        bundle, level = _resolve_circuit_bundle(row, elemento_override_map, config, warnings)
        bundles.append(bundle)
        matched_levels.append(level)

    bundle_df = pd.DataFrame(bundles, index=maestro.index)
    maestro["CircuitoNegocio"] = bundle_df["circuito_negocio"]
    maestro["PortfolioTier"] = bundle_df["portfolio_tier"]
    maestro["IncluyePerformanceCore"] = bundle_df["incluye_performance_core"].astype(bool)
    maestro["IncluyeConteoGeneral"] = bundle_df["incluye_conteo_general"].astype(bool)
    maestro["VisiblePorDefecto"] = bundle_df["visible_por_defecto"].astype(bool)
    maestro["CoberturaCatalogo"] = bundle_df["cobertura_catalogo"]
    maestro["CompletitudMaestro"] = bundle_df["completitud_maestro"]
    maestro["ModoDisponibilidad"] = bundle_df["modo_disponibilidad"]
    certeza_dato_regla = bundle_df["certeza_dato_regla"]

    # --- Actividad comercial (independiente de CertezaDato/CoberturaCatalogo) ---
    if len(campanas_raw):
        activity = campanas_raw.groupby("ElementoID").agg(
            CantidadCampanasHistoricas=("CargaID", "count"),
            FechaPrimeraCampana=("FechaInicio", "min"),
            FechaUltimaCampana=("FechaInicio", "max"),
        )
    else:
        activity = pd.DataFrame(
            columns=["CantidadCampanasHistoricas", "FechaPrimeraCampana", "FechaUltimaCampana"]
        )
    maestro = maestro.merge(activity, left_on="ElementoID", right_index=True, how="left")
    maestro["CantidadCampanasHistoricas"] = maestro["CantidadCampanasHistoricas"].fillna(0).astype("Int64")
    maestro["TieneActividadComercial"] = maestro["CantidadCampanasHistoricas"] > 0

    # --- FormatoNegocio (depende de CircuitoNegocio ya resuelto) ---
    maestro["FormatoNegocio"] = [
        _resolve_formato_negocio(row, row["CircuitoNegocio"], config, warnings) for _, row in maestro.iterrows()
    ]

    # --- SitioNegocio (depende de CircuitoNegocio ya resuelto) ---
    maestro["SitioNegocio"] = [
        _resolve_sitio_negocio(row, row["CircuitoNegocio"], config) for _, row in maestro.iterrows()
    ]

    # --- CertezaDato (depende de certeza_dato_regla + TieneActividadComercial) ---
    maestro["CertezaDato"] = [
        _resolve_certeza_dato(row, regla, row["TieneActividadComercial"], config)
        for (_, row), regla in zip(maestro.iterrows(), certeza_dato_regla)
    ]

    # --- Capacidad digital comercial efectiva (separada de la legacy) ---
    capacidades = [
        _resolve_digital_capacity(
            row, row["CircuitoNegocio"], row["FormatoNegocio"], config, warnings, elemento_override_map
        )
        for _, row in maestro.iterrows()
    ]
    maestro["SlotsComerciales"] = [c[0] for c in capacidades]
    maestro["SegundosComerciales"] = [c[1] for c in capacidades]

    # --- Invariantes de integridad ---
    if len(maestro) != len(maestro_raw):
        raise SemanticModelError("MAESTRO_ELEMENTOS: la cantidad de filas cambio durante el enriquecimiento semantico")

    original_columns = list(maestro_raw.columns)
    if not maestro[original_columns].equals(maestro_raw):
        raise SemanticModelError(
            "MAESTRO_ELEMENTOS: una o mas columnas originales fueron modificadas durante el enriquecimiento "
            "semantico (deben quedar intactas)"
        )

    ids_raw = maestro_raw.loc[maestro_raw["ElementoID"].notna(), "ElementoID"]
    ids_new = maestro.loc[maestro["ElementoID"].notna(), "ElementoID"]
    if set(ids_raw) != set(ids_new):
        raise SemanticModelError("MAESTRO_ELEMENTOS: se perdieron o alteraron ElementoID durante el enriquecimiento")
    if ids_raw.duplicated().sum() != ids_new.duplicated().sum():
        raise SemanticModelError("MAESTRO_ELEMENTOS: la cantidad de ElementoID duplicados cambio durante el enriquecimiento")

    # --- Join semantico sobre CAMPANAS (solo lectura, sin recalcular Gate1/2) ---
    # Se unen tambien columnas originales de MAESTRO_ELEMENTOS utiles para
    # cruces libres (Medio, Ciudad, CircuitoDashboard, etc.) para permitir
    # group_by/filters combinando dimensiones de ambas tablas sin joins
    # manuales por consulta. "Proveedor" y "Observaciones" existen en ambas
    # tablas con significado propio: se preserva el valor de CAMPANAS (no se
    # sobrescribe) y no se unen desde MAESTRO_ELEMENTOS.
    join_cols = [
        "ElementoID",
        "Medio",
        "Ciudad",
        "CircuitoDashboard",
        "Subcircuito",
        "Ubicacion",
        "TipoInventario",
        "TipoCatalogo",
    ] + SEMANTIC_DIMENSION_COLUMNS
    campanas = campanas.merge(maestro[join_cols], on="ElementoID", how="left")
    if len(campanas) != len(campanas_raw):
        raise SemanticModelError("CAMPANAS: la cantidad de filas cambio durante el join semantico")

    stats = {
        "rows": {"maestro": len(maestro), "campanas": len(campanas)},
        "circuito_negocio_counts": {k: int(v) for k, v in maestro["CircuitoNegocio"].value_counts().items()},
        "matched_level_counts": dict(Counter(matched_levels)),
        "no_clasificado_count": int((maestro["CircuitoNegocio"] == config["defaults"]["circuito_negocio"]).sum()),
    }

    return {
        "maestro": maestro,
        "campanas": campanas,
        "parametros": parametros_raw.copy(deep=True),
        "config": config,
        "warnings": warnings,
        "stats": stats,
    }


def filter_universe(df: pd.DataFrame, universe: str, config: dict[str, Any]) -> pd.DataFrame:
    """Filtra un DataFrame ya enriquecido (maestro o campanas con join semantico)
    segun el universo de reporte pedido. No duplica logica por circuito: cada
    universo es un filtro sobre flags ya resueltos por fila."""
    universes = config["universes"]
    if universe not in universes:
        raise SemanticModelError(f"Universo desconocido: {universe!r}. Universos validos: {sorted(universes)}")
    flag = universes[universe]["require_flag"]
    if flag is None:
        return df
    if flag not in df.columns:
        raise SemanticModelError(f"El universo {universe!r} requiere la columna {flag!r}, no presente en el DataFrame")
    return df[df[flag] == True]  # noqa: E712
