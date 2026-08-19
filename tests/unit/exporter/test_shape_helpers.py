"""Tests for per-shape helper emission in CodeGenerator."""

from __future__ import annotations

from pathlib import Path

import fastpyxl

from excel_grapher import CodeGenerator, FormulaEvaluator, create_dependency_graph
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator


def _autofill_workbook(path: Path) -> None:
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for row in range(1, 6):
        ws.cell(row, 1, row)  # A
        ws.cell(row, 2, f"=A{row}*2")  # B
        ws.cell(row, 3, f"=SUM(A1:A{row})")  # C unique shapes by range extent
    wb.save(path)
    wb.close()


def test_codegen_emits_shared_shape_helper_for_autofill(tmp_path: Path) -> None:
    path = tmp_path / "shapes.xlsx"
    _autofill_workbook(path)
    graph = create_dependency_graph(
        path,
        ["Sheet1!B1:B5"],
        load_values=True,
        warm_formula_shapes=True,
    )
    code = CodeGenerator(graph).generate(["Sheet1!B1", "Sheet1!B5"])
    assert "def _shape_0(" in code
    assert "return _shape_0(ctx," in code
    assert code.count("def _shape_") == 1
    # Cell wrappers stay thin; the arithmetic lives in the helper.
    assert "xl_mul" in code or "*" in code


def test_codegen_without_shapes_does_not_emit_helpers(tmp_path: Path) -> None:
    path = tmp_path / "noshapes.xlsx"
    _autofill_workbook(path)
    graph = create_dependency_graph(path, ["Sheet1!B1:B5"], load_values=True)
    code = CodeGenerator(graph).generate(["Sheet1!B1", "Sheet1!B5"])
    assert "def _shape_" not in code


def test_shape_helpers_match_evaluator(tmp_path: Path) -> None:
    path = tmp_path / "parity.xlsx"
    _autofill_workbook(path)
    graph = create_dependency_graph(
        path,
        ["Sheet1!B1:B5"],
        load_values=True,
        warm_formula_shapes=True,
    )
    result = assert_codegen_matches_evaluator(graph, ["Sheet1!B1", "Sheet1!B2", "Sheet1!B5"])
    assert result.evaluator_results["Sheet1!B1"] == 2.0
    assert result.generated_results["Sheet1!B1"] == 2.0
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate("Sheet1!B5") == 10.0
