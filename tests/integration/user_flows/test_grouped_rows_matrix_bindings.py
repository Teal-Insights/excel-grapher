"""Integration tests for grouped-row matrix bindings (Discrete Risks band MCVE)."""

from __future__ import annotations

from pathlib import Path

from excel_grapher.series_bindings import (
    load_series_bindings,
    validate_bindings_workbook,
)
from excel_grapher.series_bindings.workflow import setter_names
from tests.fixtures.series_bindings.grouped_matrix_helpers import (
    MATRIX_GROUPED_ROWS_BINDINGS,
    write_grouped_matrix_workbook,
)


def test_grouped_rows_matrix_bindings_validate(tmp_path: Path) -> None:
    workbook = tmp_path / "grouped_inputs.xlsx"
    write_grouped_matrix_workbook(workbook)
    bindings = load_series_bindings(MATRIX_GROUPED_ROWS_BINDINGS)
    result = validate_bindings_workbook(workbook, MATRIX_GROUPED_ROWS_BINDINGS)
    assert result["report"]["ok"] is True, result["report"]["issues"]
    assert not any(issue["level"] == "error" for issue in result["report"]["issues"])
    assert len(result["input_series"]) == len(setter_names(bindings))


def test_grouped_rows_matrix_bindings_catalog(tmp_path: Path) -> None:
    workbook = tmp_path / "grouped_inputs.xlsx"
    write_grouped_matrix_workbook(workbook)
    result = validate_bindings_workbook(workbook, MATRIX_GROUPED_ROWS_BINDINGS)
    assert result["setters"] == ["set_discrete_risks"]
    assert result["computes"] == []
