"""Integration tests for explicit matrix layout series bindings."""

from __future__ import annotations

from pathlib import Path

from excel_grapher.series_bindings import (
    load_series_bindings,
    run_binding_checks,
    validate_bindings_workbook,
)
from excel_grapher.series_bindings.workflow import setter_names
from tests.fixtures.series_bindings.matrix_helpers import (
    MATRIX_EXPLICIT_BINDINGS,
    write_matrix_explicit_workbook,
)


def test_matrix_explicit_bindings_validate(tmp_path: Path) -> None:
    workbook = tmp_path / "matrix_inputs.xlsx"
    write_matrix_explicit_workbook(workbook)
    bindings = load_series_bindings(MATRIX_EXPLICIT_BINDINGS)
    result = validate_bindings_workbook(workbook, MATRIX_EXPLICIT_BINDINGS)
    assert result["report"]["ok"] is True
    assert not any(issue["level"] == "error" for issue in result["report"]["issues"])
    assert len(result["input_series"]) == len(setter_names(bindings))


def test_matrix_explicit_bindings_run_checks(tmp_path: Path) -> None:
    workbook = tmp_path / "matrix_inputs.xlsx"
    write_matrix_explicit_workbook(workbook)
    result = run_binding_checks(
        workbook,
        MATRIX_EXPLICIT_BINDINGS,
        module_dir=tmp_path / "bindings_module",
        package_name="bindings_module",
    )
    assert result["setters"] == ["set_macro_matrix"]
    assert result["computes"] == []
