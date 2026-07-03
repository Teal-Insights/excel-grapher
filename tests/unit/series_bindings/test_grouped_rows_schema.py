"""Schema and graph-validation tests for grouped-row matrix geometry (1.5.0)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    load_series_bindings,
    validate_series_bindings,
)
from excel_grapher.series_bindings.schema import (
    SeriesBindingsSchemaError,
    validate_bindings_document,
)
from excel_grapher.series_bindings.versions import (
    IMPLEMENTED_BIND_KINDS,
    SUPPORTED_SCHEMA_VERSIONS,
)
from tests.fixtures.series_bindings.grouped_matrix_helpers import (
    MATRIX_GROUPED_ROWS_BINDINGS,
    grouped_matrix_bindings_document,
    write_grouped_matrix_workbook,
)

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "series_bindings"


def test_schema_version_1_5_0_supported() -> None:
    assert "1.5.0" in SUPPORTED_SCHEMA_VERSIONS
    assert "value_map" in IMPLEMENTED_BIND_KINDS


def test_schema_accepts_grouped_rows_fixture() -> None:
    bindings = load_series_bindings(MATRIX_GROUPED_ROWS_BINDINGS)
    series = bindings["series"][0]
    assert series["exclude_rows"] == [2, "5:6"]
    scenario_bind = series["structure"]["dimensions"][0]["bind"]
    assert scenario_bind["include"] == [2, 6]
    assert scenario_bind["fill"] is True


def test_schema_rejects_skip_and_include_together() -> None:
    doc = grouped_matrix_bindings_document()
    bind = doc["series"][0]["structure"]["dimensions"][0]["bind"]
    bind["skip"] = [3]
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_rejects_bad_row_spec_string() -> None:
    doc = grouped_matrix_bindings_document()
    doc["series"][0]["exclude_rows"] = ["5-6"]
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_rejects_bad_missing_value() -> None:
    doc = grouped_matrix_bindings_document()
    bind = doc["series"][0]["structure"]["dimensions"][0]["bind"]
    bind["missing"] = "skip"
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_accepts_missing_null_as_yaml_null() -> None:
    """`missing: null` in YAML parses to None; the schema accepts both spellings."""
    doc = grouped_matrix_bindings_document()
    bind = doc["series"][0]["structure"]["dimensions"][0]["bind"]
    bind["missing"] = None
    validate_bindings_document(doc)
    bind["missing"] = "null"
    validate_bindings_document(doc)


def test_schema_accepts_value_map_bind() -> None:
    doc = grouped_matrix_bindings_document()
    doc["series"][0]["structure"]["dimensions"][0]["bind"] = {
        "kind": "value_map",
        "values": {"Paris": "3:4", "Moderate": [7, 8]},
    }
    validate_bindings_document(doc)


def test_schema_rejects_value_map_without_values() -> None:
    doc = grouped_matrix_bindings_document()
    doc["series"][0]["structure"]["dimensions"][0]["bind"] = {"kind": "value_map"}
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def _validated_graph_and_bindings(tmp_path: Path, doc: dict[str, Any]) -> Any:
    from excel_grapher.series_bindings.ranges import expand_data_range

    wb_path = tmp_path / "grouped_inputs.xlsx"
    write_grouped_matrix_workbook(wb_path)
    targets = expand_data_range("Inputs!C2:D8")
    graph = create_dependency_graph(wb_path, targets, load_values=True)
    bindings = validate_bindings_document(doc)
    return wb_path, graph, bindings


def test_validate_grouped_rows_fixture_ok(tmp_path: Path) -> None:
    wb_path, graph, bindings = _validated_graph_and_bindings(
        tmp_path, grouped_matrix_bindings_document()
    )
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is True, report["issues"]
    codes = {issue["code"] for issue in report["issues"]}
    assert "bind_not_implemented" not in codes


def test_validate_value_map_mixed_axis_errors(tmp_path: Path) -> None:
    doc = grouped_matrix_bindings_document()
    doc["series"][0]["structure"]["dimensions"][0]["bind"] = {
        "kind": "value_map",
        "values": {"Paris": "3:4", "Moderate": "C:D"},
    }
    wb_path, graph, bindings = _validated_graph_and_bindings(tmp_path, doc)
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is False
    codes = {issue["code"] for issue in report["issues"]}
    assert "invalid_bind_geometry" in codes


def test_validate_value_map_overlapping_specs_errors(tmp_path: Path) -> None:
    doc = grouped_matrix_bindings_document()
    doc["series"][0]["structure"]["dimensions"][0]["bind"] = {
        "kind": "value_map",
        "values": {"Paris": "3:5", "Moderate": [5, 6]},
    }
    wb_path, graph, bindings = _validated_graph_and_bindings(tmp_path, doc)
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is False
    codes = {issue["code"] for issue in report["issues"]}
    assert "invalid_bind_geometry" in codes
