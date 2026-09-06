"""Element guards for array `IF` match Excel's array evaluation (slow, run-if-available).

A guard on edge `target -> dep` claims `dep` can only change `target`'s value
when the guard holds (influence, not whether Excel lists `dep` as a precedent).
This checks that claim against the real engine: perturb one element of the
value range, recalculate through Excel, and require the target to move
exactly when the graph's element guard is satisfied.

Skips cleanly when Excel automation (xlwings / WSL+COM) is unavailable.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import fastpyxl
import pytest
import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.guard import And, CellRef, Compare, GuardExpr, Literal, Not, Or
from tests.utils.modify_and_recalculate import (
    ExcelRecalculationError,
    modify_and_recalculate_workbook,
)

# A2 = 0 and A5 = -1 make rows 2 and 5 inactive; rows 1, 3, 4 are active.
_A_VALUES = {1: 3.0, 2: 0.0, 3: 7.0, 4: 1.0, 5: -1.0}
_B_VALUES = {row: 10.0 * row for row in _A_VALUES}
_C_VALUES = {row: 100.0 * row for row in _A_VALUES}


def _write_workbook(path: Path, formula: str) -> Path:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    for row, value in _A_VALUES.items():
        ws.write_number(row - 1, 0, value)
        ws.write_number(row - 1, 1, _B_VALUES[row])
        ws.write_number(row - 1, 2, _C_VALUES[row])
    ws.write_array_formula("E1:E1", formula, None, 0)
    wb.close()
    return path


def _recalculate(source: Path, output: Path, modifications: dict[str, float]) -> float:
    try:
        modify_and_recalculate_workbook(source, output, modifications)
    except (ExcelRecalculationError, RuntimeError, ImportError) as exc:
        pytest.skip(f"Excel recalculation not available: {exc}")
    wb = fastpyxl.load_workbook(output, data_only=True)
    try:
        value = wb["Sheet1"]["E1"].value
    finally:
        wb.close()
    assert isinstance(value, (int, float))
    return float(value)


def _guard_holds(guard: GuardExpr | None, values: dict[str, float]) -> bool:
    """Evaluate a scalar edge guard over cell values (`None` = unconditional)."""
    if guard is None:
        return True
    if isinstance(guard, Not):
        return not _guard_holds(guard.operand, values)
    if isinstance(guard, And):
        return all(_guard_holds(operand, values) for operand in guard.operands)
    if isinstance(guard, Or):
        return any(_guard_holds(operand, values) for operand in guard.operands)
    if isinstance(guard, Compare):
        left = _guard_operand(guard.left, values)
        right = _guard_operand(guard.right, values)
        return {
            "=": left == right,
            "<>": left != right,
            ">": left > right,
            "<": left < right,
            ">=": left >= right,
            "<=": left <= right,
        }[guard.op]
    raise AssertionError(f"Unexpected guard form on an edge: {guard!r}")


def _guard_operand(expr: GuardExpr, values: dict[str, float]) -> Any:
    if isinstance(expr, Literal):
        return expr.value
    if isinstance(expr, CellRef):
        return values[expr.key]
    raise AssertionError(f"Unexpected guard operand: {expr!r}")


@pytest.mark.slow
@pytest.mark.parametrize(
    ("formula", "guarded_column"),
    [
        ("=SUM(IF(A1:A5>0,B1:B5,0))", "B"),
        ("=SUM(IF(A1:A5>0,B1:B5,C1:C5))", "B"),
        ("=SUM(IF(A1:A5>0,B1:B5,C1:C5))", "C"),
    ],
)
def test_element_guards_predict_which_cells_excel_reads(
    tmp_path: Path, formula: str, guarded_column: str
) -> None:
    source = _write_workbook(tmp_path / "array_if.xlsx", formula)
    graph = create_dependency_graph(source, ["Sheet1!E1"], load_values=False)
    cell_values = {f"Sheet1!A{row}": value for row, value in _A_VALUES.items()}

    baseline = _recalculate(source, tmp_path / "baseline.xlsx", {})

    for row in _A_VALUES:
        dep = f"Sheet1!{guarded_column}{row}"
        guard_holds = _guard_holds(graph.get_edge_guard("Sheet1!E1", dep), cell_values)
        recalculated = _recalculate(
            source, tmp_path / f"perturbed_{guarded_column}{row}.xlsx", {dep: 12345.0}
        )
        assert (recalculated != baseline) == guard_holds, (
            f"{dep}: guard says active={guard_holds}, but Excel "
            f"{'changed' if recalculated != baseline else 'did not change'} E1"
        )
