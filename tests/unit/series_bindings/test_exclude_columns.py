"""Schema and resolve tests for series-level `exclude_columns` (schema 1.12.0)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    expand_data_range,
    resolve_series_binding,
    validate_series_bindings,
)
from excel_grapher.series_bindings.schema import (
    SeriesBindingsSchemaError,
    validate_bindings_document,
)
from excel_grapher.series_bindings.versions import SUPPORTED_SCHEMA_VERSIONS


def _matrix_doc(
    *,
    exclude_columns: list[Any] | None = None,
    exclude_rows: list[Any] | None = None,
    data_range: str = "Demo!B2:D3",
) -> dict[str, Any]:
    series: dict[str, Any] = {
        "id": "demo_matrix",
        "sheet": "Demo",
        "data_range": data_range,
        "layout": "matrix",
        "output": {"compute": {"name": "compute_demo_matrix"}},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell"},
            },
            "dimensions": [
                {
                    "id": "SCENARIO",
                    "concept": "SCENARIO",
                    "dtype": "string",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "row_label", "label_column": "A", "read": "string"},
                },
                {
                    "id": "TIME_PERIOD",
                    "concept": "TIME_PERIOD",
                    "dtype": "int",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                },
            ],
        },
        "key": ["SCENARIO", "TIME_PERIOD"],
    }
    if exclude_columns is not None:
        series["exclude_columns"] = exclude_columns
    if exclude_rows is not None:
        series["exclude_rows"] = exclude_rows
    return {"schema_version": "1.12.0", "series": [series]}


def _write_mcve_workbook(path: Path) -> None:
    """Sheet Demo: headers 2028/2050/2029 in B1:D1; Paris/London rows in A2:A3."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Demo")
    for col, year in enumerate([2028, 2050, 2029], start=1):
        ws.write(0, col, year)
    ws.write("A2", "Paris")
    ws.write("A3", "London")
    for row, values in enumerate([[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]], start=1):
        for col, value in enumerate(values, start=1):
            ws.write_number(row, col, value)
    # Block-wide SUM so milestone column C stays a graph leaf when not excluded.
    ws.write_formula("E4", "=SUM(B2:D3)")
    wb.close()


def _resolve_doc(tmp_path: Path, doc: dict[str, Any]) -> dict[str, Any]:
    wb_path = tmp_path / "exclude_columns.xlsx"
    _write_mcve_workbook(wb_path)
    targets = expand_data_range("Demo!B2:D3")
    graph = create_dependency_graph(wb_path, targets, load_values=True)
    bindings = validate_bindings_document(doc)
    return resolve_series_binding(graph, wb_path, bindings["series"][0])


def test_schema_version_1_12_0_supported() -> None:
    assert "1.12.0" in SUPPORTED_SCHEMA_VERSIONS


def test_schema_accepts_exclude_columns_single_and_range() -> None:
    doc = _matrix_doc(exclude_columns=["C", "B:B", "AR:BQ"])
    bindings = validate_bindings_document(doc)
    assert bindings["series"][0]["exclude_columns"] == ["C", "B:B", "AR:BQ"]


def test_schema_rejects_bad_column_spec() -> None:
    doc = _matrix_doc(exclude_columns=["5"])
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_validate_rejects_invalid_exclude_columns_geometry(tmp_path: Path) -> None:
    """Smoke validation expands specs even when schema pattern is bypassed."""
    wb_path = tmp_path / "exclude_columns.xlsx"
    _write_mcve_workbook(wb_path)
    targets = expand_data_range("Demo!B2:D3")
    graph = create_dependency_graph(wb_path, targets, load_values=True)
    doc = _matrix_doc(exclude_columns=["C"])
    bindings = validate_bindings_document(doc)
    bindings["series"][0]["exclude_columns"] = ["ZZZZ"]
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is False
    codes = {issue["code"] for issue in report["issues"]}
    assert "invalid_bind_geometry" in codes


def test_exclude_columns_drops_single_column(tmp_path: Path) -> None:
    """MCVE: data_range B2:D2 minus column C keeps B2 and D2 only."""
    doc = _matrix_doc(exclude_columns=["C"], data_range="Demo!B2:D2")
    resolved = _resolve_doc(tmp_path, doc)
    assert resolved["ok"] is True, resolved["issues"]
    addresses = {leaf["address"] for leaf in resolved["leaves"]}
    assert addresses == {"Demo!B2", "Demo!D2"}
    periods = {leaf["key"]["TIME_PERIOD"] for leaf in resolved["leaves"]}
    assert periods == {2028, 2029}


def test_exclude_columns_drops_column_range(tmp_path: Path) -> None:
    doc = _matrix_doc(exclude_columns=["C:D"])
    resolved = _resolve_doc(tmp_path, doc)
    assert resolved["ok"] is True, resolved["issues"]
    addresses = {leaf["address"] for leaf in resolved["leaves"]}
    assert addresses == {"Demo!B2", "Demo!B3"}


def test_exclude_columns_composes_with_exclude_rows(tmp_path: Path) -> None:
    doc = _matrix_doc(exclude_columns=["C"], exclude_rows=[3])
    resolved = _resolve_doc(tmp_path, doc)
    assert resolved["ok"] is True, resolved["issues"]
    addresses = {leaf["address"] for leaf in resolved["leaves"]}
    assert addresses == {"Demo!B2", "Demo!D2"}
    assert all(leaf["key"]["SCENARIO"] == "Paris" for leaf in resolved["leaves"])
