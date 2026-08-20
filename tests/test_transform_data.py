"""Pruebas para scripts/transform_data.py (Gate 2 del pipeline OCU26).

Ninguna prueba modifica input/OCU26_BASE_DATOS.xlsx. Los workbooks inválidos
o de tamaño reducido se generan en tmp_path, nunca sobre el archivo
productivo. Reutiliza los helpers de construcción de workbooks de
test_validate_input.py (no los redefine) sin modificar ese archivo.
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pandas as pd
import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT / "scripts"))
sys.path.insert(0, str(REPO_ROOT / "tests"))

import transform_data as td  # noqa: E402
import validate_input as vi  # noqa: E402
from test_validate_input import (  # noqa: E402
    _base_campana_row,
    _base_maestro_row,
    _write_workbook,
)

PRODUCTION_FILE = REPO_ROOT / "input" / "OCU26_BASE_DATOS.xlsx"

MAESTRO_IDX = {h: i for i, h in enumerate(vi.MAESTRO_HEADERS)}
CAMPANAS_IDX = {h: i for i, h in enumerate(vi.CAMPANAS_HEADERS)}


def _maestro_row(elemento_id: str = "ELEM-0001", **overrides):
    row = _base_maestro_row(elemento_id)
    for field, value in overrides.items():
        row[MAESTRO_IDX[field]] = value
    return row


def _campana_row(carga_id: str, elemento_id: str = "ELEM-0001", **overrides):
    row = _base_campana_row(carga_id, elemento_id=elemento_id)
    for field, value in overrides.items():
        row[CAMPANAS_IDX[field]] = value
    return row


def _build_single_element_workbook(
    tmp_path: Path, filename: str, elemento_id: str = "ELEM-0001", medio: str = "Digital", **maestro_overrides
) -> Path:
    """Construye un workbook mínimo válido con un solo elemento, coherente
    entre MAESTRO_ELEMENTOS.Medio y CAMPANAS.TipoCargaDeclarado."""
    maestro_overrides.setdefault("Medio", medio)
    maestro_row = _maestro_row(elemento_id, **maestro_overrides)
    campana_row = _campana_row("HIST-0001", elemento_id=elemento_id, TipoCargaDeclarado=medio)
    path = tmp_path / filename
    _write_workbook(path, maestro_rows=[maestro_row], campanas_rows=[campana_row])
    return path


# ---------------------------------------------------------------------------
# 1-6: archivo productivo real
# ---------------------------------------------------------------------------


def test_production_file_transforms_successfully():
    assert PRODUCTION_FILE.exists(), "input/OCU26_BASE_DATOS.xlsx no existe"
    sha_before = vi.calculate_sha256(PRODUCTION_FILE)

    result = td.transform_data(PRODUCTION_FILE)

    assert result["validation"]["result"] in ("VALID", "VALID_WITH_WARNINGS")
    assert len(result["maestro"]) == result["stats"]["rows"]["maestro"]
    assert len(result["campanas"]) == result["stats"]["rows"]["campanas"]
    assert len(result["parametros"]) == result["stats"]["rows"]["parametros"]

    maestro_raw = td.read_excel_table(PRODUCTION_FILE, "MAESTRO_ELEMENTOS", "tblElementos")
    campanas_raw = td.read_excel_table(PRODUCTION_FILE, "CAMPANAS", "tblCampanas")
    parametros_raw = td.read_excel_table(PRODUCTION_FILE, "PARAMETROS", "tblParametros")

    assert len(result["maestro"]) == len(maestro_raw)
    assert len(result["campanas"]) == len(campanas_raw)
    assert len(result["parametros"]) == len(parametros_raw)

    sha_after = vi.calculate_sha256(PRODUCTION_FILE)
    assert sha_after == sha_before


def test_production_file_via_cli_does_not_modify_input():
    sha_before = vi.calculate_sha256(PRODUCTION_FILE)
    proc = subprocess.run(
        [sys.executable, str(REPO_ROOT / "scripts" / "transform_data.py")],
        cwd=REPO_ROOT,
        capture_output=True,
    )
    assert proc.returncode == 0
    assert vi.calculate_sha256(PRODUCTION_FILE) == sha_before


# ---------------------------------------------------------------------------
# 7-11: regla TipoInventario
# ---------------------------------------------------------------------------


def test_tipo_inventario_digital_sin_keywords(tmp_path):
    path = _build_single_element_workbook(tmp_path, "digital.xlsx", medio="Digital")
    result = td.transform_data(path)
    assert result["maestro"].loc[0, "TipoInventario"] == "Digital"


def test_tipo_inventario_estatico_sin_keywords(tmp_path):
    path = _build_single_element_workbook(tmp_path, "estatico.xlsx", medio="Estático")
    result = td.transform_data(path)
    assert result["maestro"].loc[0, "TipoInventario"] == "Físico estático"


@pytest.mark.parametrize("keyword", ["CARRO", "STOPPER", "FLOORGRAPHIC", "CUBRE", "ALARMA"])
def test_tipo_inventario_keyword_produces_flexible_grafico(tmp_path, keyword):
    path = _build_single_element_workbook(
        tmp_path, f"kw_{keyword}.xlsx", medio="Digital", Ubicacion=f"Zona con {keyword} instalado"
    )
    result = td.transform_data(path)
    assert result["maestro"].loc[0, "TipoInventario"] == "Flexible gráfico"


def test_tipo_inventario_keyword_case_insensitive(tmp_path):
    path = _build_single_element_workbook(tmp_path, "lower.xlsx", medio="Digital", Ubicacion="tiene un carro ahi")
    result = td.transform_data(path)
    assert result["maestro"].loc[0, "TipoInventario"] == "Flexible gráfico"


@pytest.mark.parametrize("field", ["Ubicacion", "Descripcion", "Subcircuito", "CircuitoDashboard"])
def test_tipo_inventario_keyword_detected_in_any_search_field(tmp_path, field):
    path = _build_single_element_workbook(tmp_path, f"field_{field}.xlsx", medio="Digital", **{field: "STOPPER"})
    result = td.transform_data(path)
    assert result["maestro"].loc[0, "TipoInventario"] == "Flexible gráfico"


def test_derive_tipo_inventario_blank_elemento_id_returns_empty():
    row = {
        "ElementoID": None,
        "Medio": "Digital",
        "Ubicacion": "",
        "Descripcion": "",
        "Subcircuito": "",
        "CircuitoDashboard": "",
    }
    assert td.derive_tipo_inventario(row) == ""


# ---------------------------------------------------------------------------
# 12-14: regla AplicaCantidad
# ---------------------------------------------------------------------------


def test_aplica_cantidad_flexible_grafico_es_si(tmp_path):
    path = _build_single_element_workbook(tmp_path, "aplica_flex.xlsx", medio="Digital", Descripcion="Tiene ALARMA")
    result = td.transform_data(path)
    assert result["maestro"].loc[0, "TipoInventario"] == "Flexible gráfico"
    assert result["maestro"].loc[0, "AplicaCantidad"] == "SI"


def test_aplica_cantidad_digital_es_no(tmp_path):
    path = _build_single_element_workbook(tmp_path, "aplica_digital.xlsx", medio="Digital")
    result = td.transform_data(path)
    assert result["maestro"].loc[0, "AplicaCantidad"] == "NO"


def test_aplica_cantidad_fisico_estatico_es_no(tmp_path):
    path = _build_single_element_workbook(tmp_path, "aplica_estatico.xlsx", medio="Estático")
    result = td.transform_data(path)
    assert result["maestro"].loc[0, "AplicaCantidad"] == "NO"


# ---------------------------------------------------------------------------
# 15-18: política fuente vacía / correcta / contradictoria
# ---------------------------------------------------------------------------


def test_tipo_inventario_fuente_vacia_se_completa(tmp_path):
    path = _build_single_element_workbook(tmp_path, "blank_tipo.xlsx", medio="Digital", TipoInventario=None)
    result = td.transform_data(path)
    assert result["maestro"].loc[0, "TipoInventario"] == "Digital"


def test_tipo_inventario_fuente_correcta_se_acepta(tmp_path):
    path = _build_single_element_workbook(tmp_path, "correct_tipo.xlsx", medio="Digital", TipoInventario="Digital")
    result = td.transform_data(path)
    assert result["maestro"].loc[0, "TipoInventario"] == "Digital"


def test_tipo_inventario_fuente_contradictoria_lanza_transform_error(tmp_path):
    path = _build_single_element_workbook(
        tmp_path, "wrong_tipo.xlsx", medio="Digital", TipoInventario="Físico estático"
    )
    with pytest.raises(td.TransformError, match="TipoInventario"):
        td.transform_data(path)


def test_aplica_cantidad_fuente_contradictoria_lanza_transform_error(tmp_path):
    path = _build_single_element_workbook(tmp_path, "wrong_aplica.xlsx", medio="Digital", AplicaCantidad="SI")
    with pytest.raises(td.TransformError, match="AplicaCantidad"):
        td.transform_data(path)


# ---------------------------------------------------------------------------
# 19-20: normalización numérica
# ---------------------------------------------------------------------------


def test_capacidad_slots_reel_numero_y_texto_dan_mismo_valor(tmp_path):
    path_num = _build_single_element_workbook(tmp_path, "cap_num.xlsx", CapacidadSlotsReel=20)
    path_text = _build_single_element_workbook(tmp_path, "cap_text.xlsx", CapacidadSlotsReel="20")

    result_num = td.transform_data(path_num)
    result_text = td.transform_data(path_text)

    assert result_num["maestro"].loc[0, "CapacidadSlotsReel"] == 20
    assert result_text["maestro"].loc[0, "CapacidadSlotsReel"] == 20
    assert str(result_num["maestro"]["CapacidadSlotsReel"].dtype) == "Int64"
    assert str(result_text["maestro"]["CapacidadSlotsReel"].dtype) == "Int64"


def test_segundos_dia_numero_y_texto_dan_mismo_valor(tmp_path):
    path_num = _build_single_element_workbook(tmp_path, "seg_num.xlsx", SegundosDia=100800)
    path_text = _build_single_element_workbook(tmp_path, "seg_text.xlsx", SegundosDia="100800")

    result_num = td.transform_data(path_num)
    result_text = td.transform_data(path_text)

    assert result_num["maestro"].loc[0, "SegundosDia"] == 100800
    assert result_text["maestro"].loc[0, "SegundosDia"] == 100800
    assert str(result_num["maestro"]["SegundosDia"].dtype) == "Int64"
    assert str(result_text["maestro"]["SegundosDia"].dtype) == "Int64"


# ---------------------------------------------------------------------------
# 21-23: equivalencia de campos no objetivo
# ---------------------------------------------------------------------------


def test_maestro_non_target_columns_unchanged(tmp_path):
    path = _build_single_element_workbook(
        tmp_path, "non_target.xlsx", medio="Digital", Ciudad="Rosario", Observaciones="Nota original"
    )
    raw = td.read_excel_table(path, "MAESTRO_ELEMENTOS", "tblElementos")
    result = td.transform_data(path)
    maestro = result["maestro"]

    for column in ("Ciudad", "Observaciones", "RevisionMaestro", "Proveedor", "Nivel", "Descripcion"):
        assert maestro.loc[0, column] == raw.loc[0, column] or (
            td._is_blank(maestro.loc[0, column]) and td._is_blank(raw.loc[0, column])
        )


def test_campanas_equivalent_value_by_value(tmp_path):
    path = _build_single_element_workbook(tmp_path, "campanas_passthrough.xlsx", medio="Digital")
    raw = td.read_excel_table(path, "CAMPANAS", "tblCampanas")
    result = td.transform_data(path)
    assert raw.equals(result["campanas"])


def test_parametros_equivalent_value_by_value(tmp_path):
    path = _build_single_element_workbook(tmp_path, "parametros_passthrough.xlsx", medio="Digital")
    raw = td.read_excel_table(path, "PARAMETROS", "tblParametros")
    result = td.transform_data(path)
    assert raw.equals(result["parametros"])


# ---------------------------------------------------------------------------
# 24: input INVALID aborta
# ---------------------------------------------------------------------------


def test_invalid_input_aborts_transformation(tmp_path):
    path = tmp_path / "dup_elemento_id.xlsx"
    _write_workbook(
        path,
        maestro_rows=[_base_maestro_row("ELEM-0001"), _base_maestro_row("ELEM-0001")],
        campanas_rows=[_base_campana_row("HIST-0001")],
    )
    sha_before = vi.calculate_sha256(path)

    with pytest.raises(td.TransformError):
        td.transform_data(path)

    assert vi.calculate_sha256(path) == sha_before


def test_invalid_input_via_cli_returns_exit_code_1(tmp_path):
    path = tmp_path / "dup_elemento_id_cli.xlsx"
    _write_workbook(
        path,
        maestro_rows=[_base_maestro_row("ELEM-0001"), _base_maestro_row("ELEM-0001")],
        campanas_rows=[_base_campana_row("HIST-0001")],
    )
    proc = subprocess.run(
        [sys.executable, str(REPO_ROOT / "scripts" / "transform_data.py"), "--file", str(path)],
        capture_output=True,
    )
    assert proc.returncode == 1


# ---------------------------------------------------------------------------
# CLI adicional: --json, lectura por rango real de tabla
# ---------------------------------------------------------------------------


def test_cli_json_output_is_valid_json(tmp_path):
    path = _build_single_element_workbook(tmp_path, "json_valid.xlsx", medio="Digital")
    proc = subprocess.run(
        [sys.executable, str(REPO_ROOT / "scripts" / "transform_data.py"), "--file", str(path), "--json"],
        capture_output=True,
    )
    assert proc.returncode == 0
    data = json.loads(proc.stdout.decode("utf-8"))
    assert data["result"] == "TRANSFORM_OK"
    assert "stats" in data
    assert data["stats"]["rows"]["maestro"] == 1


def test_read_excel_table_uses_real_table_range_not_hardcoded_size(tmp_path):
    path = _build_single_element_workbook(tmp_path, "small_range.xlsx", medio="Digital")
    maestro = td.read_excel_table(path, "MAESTRO_ELEMENTOS", "tblElementos")
    campanas = td.read_excel_table(path, "CAMPANAS", "tblCampanas")
    assert len(maestro) == 1
    assert len(campanas) == 1


# ---------------------------------------------------------------------------
# Fix Gate 2 (integración YPF, 2026-08-18): la validación de passthrough de
# CAMPANAS.FilaOrigen usaba list(...) != list(...), que compara NaN con NaN
# vía `==` de Python y siempre da False -> falso positivo de "cambio" en
# columnas con blancos legítimos (FilaOrigen es opcional; ElementoID/CargaID
# no lo son y ya están bloqueados por Gate 1 si vinieran vacíos, por eso no
# comparten el mismo riesgo y no se tocaron). Ahora usa
# Series.reset_index(drop=True).equals(...), NaN-seguro. Estas pruebas fijan
# el comportamiento exacto de esa comparación y verifican los dos archivos
# reales involucrados en la promoción YPF.
# ---------------------------------------------------------------------------

YPF_BACKUP_FILE = REPO_ROOT / "Pendientes" / "OCU26_YPF_INTEGRACION" / "backup" / "OCU26_BASE_DATOS_PRE_YPF_2026-08-18.xlsx"
YPF_FINAL_FILE = REPO_ROOT / "Pendientes" / "OCU26_YPF_INTEGRACION" / "output" / "OCU26_BASE_DATOS_CON_YPF_FINAL_2026-08-18.xlsx"


def test_filaorigen_nan_en_mismas_posiciones_no_bloquea(tmp_path):
    """1. Passthrough real con NaN en la misma posición: no debe lanzar
    TransformError (regresión directa del bug list(...) != list(...))."""
    path = tmp_path / "filaorigen_nan_passthrough.xlsx"
    _write_workbook(
        path,
        maestro_rows=[_base_maestro_row("ELEM-0001")],
        campanas_rows=[
            _campana_row("HIST-0001", FilaOrigen=1),
            _campana_row("HIST-0002", FilaOrigen=None),
            _campana_row("HIST-0003", FilaOrigen=3),
        ],
    )
    result = td.transform_data(path)  # no debe lanzar TransformError
    filaorigen = result["campanas"]["FilaOrigen"]
    assert len(filaorigen) == 3
    assert filaorigen.isna().sum() == 1


def test_filaorigen_equals_semantics_valor_distinto_bloquea():
    """2. Un valor distinto en la misma posición se detecta como cambio."""
    a = pd.Series([1, 2, 3])
    b = pd.Series([1, 99, 3])
    assert not a.reset_index(drop=True).equals(b.reset_index(drop=True))


def test_filaorigen_equals_semantics_nan_en_mismas_posiciones_no_bloquea():
    """(complemento de 1, a nivel de la primitiva) NaN en la misma posición
    en ambas series no se considera un cambio."""
    a = pd.Series([1, None, 3])
    b = pd.Series([1, None, 3])
    assert a.reset_index(drop=True).equals(b.reset_index(drop=True))


def test_filaorigen_equals_semantics_nan_vs_valor_real_bloquea():
    """3. NaN en una serie frente a un valor real en la otra, misma
    posición, se detecta como cambio."""
    a = pd.Series([1, None, 3])
    b = pd.Series([1, 2, 3])
    assert not a.reset_index(drop=True).equals(b.reset_index(drop=True))


def test_filaorigen_equals_semantics_cambio_de_orden_bloquea():
    """4. Mismos valores, orden distinto: se detecta como cambio."""
    a = pd.Series([1, 2, 3])
    b = pd.Series([3, 2, 1])
    assert not a.reset_index(drop=True).equals(b.reset_index(drop=True))


def test_filaorigen_equals_semantics_diferencia_de_filas_bloquea():
    """5. Distinta cantidad de filas: se detecta como cambio."""
    a = pd.Series([1, 2, 3])
    b = pd.Series([1, 2, 3, 4])
    assert not a.reset_index(drop=True).equals(b.reset_index(drop=True))


def test_pre_ypf_backup_sigue_pasando_el_gate():
    """6. El input previo a la promoción YPF (respaldo, 0 FilaOrigen vacíos)
    sigue pasando Gate 2 sin cambio de comportamiento tras el fix."""
    if not YPF_BACKUP_FILE.exists():
        pytest.skip("Respaldo PRE_YPF no disponible en este entorno")
    sha_before = vi.calculate_sha256(YPF_BACKUP_FILE)
    result = td.transform_data(YPF_BACKUP_FILE)
    assert result["validation"]["result"] in ("VALID", "VALID_WITH_WARNINGS")
    assert vi.calculate_sha256(YPF_BACKUP_FILE) == sha_before


def test_base_final_ypf_pasa_el_gate_sin_alterar_filaorigen():
    """7. La base FINAL con YPF integrado (13.616 filas nuevas con
    FilaOrigen vacío) pasa Gate 2, y el passthrough de FilaOrigen es
    exactamente igual al crudo leído del Excel (ninguna fila perdida,
    reordenada ni modificada)."""
    if not YPF_FINAL_FILE.exists():
        pytest.skip("Base FINAL YPF no disponible en este entorno")
    sha_before = vi.calculate_sha256(YPF_FINAL_FILE)
    raw = td.read_excel_table(YPF_FINAL_FILE, "CAMPANAS", "tblCampanas")
    result = td.transform_data(YPF_FINAL_FILE)
    assert result["validation"]["result"] in ("VALID", "VALID_WITH_WARNINGS")
    assert result["campanas"]["FilaOrigen"].reset_index(drop=True).equals(raw["FilaOrigen"].reset_index(drop=True))
    assert result["campanas"]["FilaOrigen"].isna().sum() >= 13616
    assert vi.calculate_sha256(YPF_FINAL_FILE) == sha_before
