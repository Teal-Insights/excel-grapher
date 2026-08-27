"""Tests for per-shape helper emission in CodeGenerator."""

from __future__ import annotations

from pathlib import Path

import fastpyxl

from excel_grapher import CodeGenerator, FormulaEvaluator, create_dependency_graph
from excel_grapher.core.formula_shape import fingerprint_formula_shape
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator


def _graph(path: Path, targets: list[str]):
    return create_dependency_graph(
        path,
        targets,
        load_values=True,
        warm_formula_shapes=True,
    )


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
    # Interned params stay relative; call sites pass host-resolved A1.
    assert "C[-1]" not in code
    assert "R[" not in code
    assert "return _shape_0(ctx, 'Sheet1!A1')" in code
    assert "return _shape_0(ctx, 'Sheet1!A5')" in code
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


def test_ref_inspecting_skeletons_are_not_helper_eligible() -> None:
    """OFFSET/INDEX/ROW emit paths inspect concrete ref ASTs, not address holes."""
    cases = (
        "=OFFSET(Sheet1!A1,1,0)",
        "=INDEX(Sheet1!A1:A3,1)",
        "=ROW(Sheet1!A1)",
        "=COLUMN(Sheet1!A1)",
        "=COLUMNS(Sheet1!A1:B1)",
        "=ROW(OFFSET(Sheet1!A1,1,0))",
    )
    for formula in cases:
        skeleton = fingerprint_formula_shape(formula).skeleton
        assert CodeGenerator._shape_helper_eligible(skeleton) is False, formula


def test_sum_range_helpers_match_evaluator(tmp_path: Path) -> None:
    path = tmp_path / "sum_ranges.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for row in range(1, 4):
        ws.cell(row, 1, row)
        ws.cell(row, 2, row * 10)
    ws["C1"] = "=SUM(A1:A2)"
    ws["C2"] = "=SUM(B1:B2)"
    wb.save(path)
    wb.close()

    graph = _graph(path, ["Sheet1!C1", "Sheet1!C2"])
    code = CodeGenerator(graph).generate(["Sheet1!C1", "Sheet1!C2"])
    assert "def _shape_" in code
    result = assert_codegen_matches_evaluator(graph, ["Sheet1!C1", "Sheet1!C2"])
    assert result.evaluator_results["Sheet1!C1"] == 3
    assert result.generated_results["Sheet1!C1"] == 3


def test_formula_to_formula_helpers_match_evaluator(tmp_path: Path) -> None:
    path = tmp_path / "formula_refs.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = 10
    ws["A2"] = 20
    ws["B1"] = "=A1+1"
    ws["B2"] = "=A2+1"
    ws["C1"] = "=B1*2"
    ws["C2"] = "=B2*2"
    wb.save(path)
    wb.close()

    graph = _graph(path, ["Sheet1!C1", "Sheet1!C2"])
    code = CodeGenerator(graph).generate(["Sheet1!C1", "Sheet1!C2"])
    assert "def _shape_" in code
    result = assert_codegen_matches_evaluator(graph, ["Sheet1!C1", "Sheet1!C2"])
    assert result.evaluator_results["Sheet1!C1"] == 22
    assert result.generated_results["Sheet1!C1"] == 22


def test_offset_autofill_matches_evaluator_without_helpers(tmp_path: Path) -> None:
    path = tmp_path / "offset.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = 10
    ws["A2"] = 20
    ws["A3"] = 30
    ws["B1"] = "=OFFSET(A1,1,0)"
    ws["B2"] = "=OFFSET(A2,1,0)"
    wb.save(path)
    wb.close()

    graph = _graph(path, ["Sheet1!B1", "Sheet1!B2"])
    code = CodeGenerator(graph).generate(["Sheet1!B1", "Sheet1!B2"])
    assert "def _shape_" not in code
    result = assert_codegen_matches_evaluator(graph, ["Sheet1!B1", "Sheet1!B2"])
    assert result.evaluator_results["Sheet1!B1"] == 20
    assert result.evaluator_results["Sheet1!B2"] == 30


def test_index_autofill_matches_evaluator_without_helpers(tmp_path: Path) -> None:
    path = tmp_path / "index.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = 10
    ws["A2"] = 20
    ws["A3"] = 30
    ws["C1"] = 100
    ws["C2"] = 200
    ws["C3"] = 300
    ws["B1"] = "=INDEX(A1:A3,1)"
    ws["B2"] = "=INDEX(C1:C3,1)"
    wb.save(path)
    wb.close()

    graph = _graph(path, ["Sheet1!B1", "Sheet1!B2"])
    code = CodeGenerator(graph).generate(["Sheet1!B1", "Sheet1!B2"])
    assert "def _shape_" not in code
    result = assert_codegen_matches_evaluator(graph, ["Sheet1!B1", "Sheet1!B2"])
    assert result.evaluator_results["Sheet1!B1"] == 10
    assert result.evaluator_results["Sheet1!B2"] == 100


def test_one_by_one_range_operators_match_evaluator(tmp_path: Path) -> None:
    path = tmp_path / "scalar_ranges.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = 3
    ws["B1"] = 5
    ws["C1"] = "=A1:A1*2"
    ws["C2"] = "=B1:B1*2"
    wb.save(path)
    wb.close()

    graph = _graph(path, ["Sheet1!C1", "Sheet1!C2"])
    code = CodeGenerator(graph).generate(["Sheet1!C1", "Sheet1!C2"])
    assert "def _shape_" not in code
    result = assert_codegen_matches_evaluator(graph, ["Sheet1!C1", "Sheet1!C2"])
    assert result.evaluator_results["Sheet1!C1"] == 6
    assert result.generated_results["Sheet1!C1"] == 6
