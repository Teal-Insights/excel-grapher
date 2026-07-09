"""Tests for schema version and feature support metadata."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    IMPLEMENTED_LAYOUTS,
    load_series_bindings,
    validate_series_bindings,
)
from excel_grapher.series_bindings.versions import SUPPORTED_SCHEMA_VERSIONS
from tests.paths import SERIES_BINDINGS_FIXTURES as FIXTURES


def test_supported_schema_versions() -> None:
    expected = frozenset({"1.0.0", "1.1.0", "1.2.0", "1.3.0", "1.4.0", "1.5.0", "1.6.0", "1.7.0"})
    assert expected == SUPPORTED_SCHEMA_VERSIONS


def test_validate_explicit_matrix_no_implementation_warnings(tmp_path: Path) -> None:
    from tests.fixtures.series_bindings.matrix_helpers import write_matrix_explicit_workbook

    wb_path = tmp_path / "matrix_inputs.xlsx"
    write_matrix_explicit_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        ["Inputs!B3", "Inputs!C3", "Inputs!D5"],
        load_values=True,
    )
    bindings = load_series_bindings(FIXTURES / "matrix_explicit_1_4_0.yaml")
    report = validate_series_bindings(graph, bindings)
    codes = {i["code"] for i in report["issues"]}
    assert "unknown_layout" not in codes
    assert "unknown_bind_kind" not in codes
    assert "matrix" in IMPLEMENTED_LAYOUTS


def test_validate_country_block_matrix_fixture(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Inputs")
    for col, year in enumerate([1, 2], start=5):
        ws.write(0, col, year)
    ws.write_number(2, 5, 1.0)
    ws.write_number(3, 5, 2.0)
    wb.close()

    graph = create_dependency_graph(wb_path, ["Inputs!F3", "Inputs!F4"], load_values=True)
    bindings = load_series_bindings(FIXTURES / "matrix_country_block_1_1_0.yaml")
    report = validate_series_bindings(graph, bindings)
    codes = {i["code"] for i in report["issues"]}
    assert "unknown_layout" not in codes
    assert "unknown_bind_kind" not in codes
    assert "matrix" in IMPLEMENTED_LAYOUTS
