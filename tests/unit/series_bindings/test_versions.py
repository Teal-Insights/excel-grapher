"""Tests for schema version and feature support metadata."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    IMPLEMENTED_LAYOUTS,
    PLANNED_BIND_KINDS,
    load_series_bindings,
    validate_series_bindings,
)
from excel_grapher.series_bindings.versions import SUPPORTED_SCHEMA_VERSIONS

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "series_bindings"


def test_supported_schema_versions() -> None:
    assert frozenset({"1.0.0", "1.1.0"}) == SUPPORTED_SCHEMA_VERSIONS


def test_validate_1_1_0_matrix_emits_implementation_warnings(tmp_path: Path) -> None:
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
    assert "layout_not_implemented" in codes
    assert "bind_not_implemented" in codes
    assert "matrix" not in IMPLEMENTED_LAYOUTS
    assert "row_hierarchy" in PLANNED_BIND_KINDS
