"""CHOOSE branch guards stay sound when the index can be fractional (integration).

Excel truncates a non-integer `CHOOSE` index (`CHOOSE(1.5, a, b)` is `a`), so a
branch guard of `index = i` would prune live edges and hide real cycles.
"""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Any
from typing import Literal as TypingLiteral

import xlsxwriter

from excel_grapher import DynamicRefConfig, create_dependency_graph
from excel_grapher.core.cell_types import Between, RealBetween
from excel_grapher.grapher.guard import And, CellRef, Compare, GuardExpr, Literal, Not, Or


def _guard_holds(guard: GuardExpr | None, values: dict[str, float]) -> bool:
    """Evaluate a scalar edge guard over numeric cell values (`None` = unconditional)."""
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


def _make_choose_cycle_workbook(path: Path) -> None:
    """D1 = CHOOSE(A1, B1, C1); B1 = D1 + 1 closes a cycle through branch 1."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 1.5)  # A1
    ws.write_formula(0, 1, "=D1+1", None, 0)  # B1
    ws.write_number(0, 2, 20)  # C1
    ws.write_formula(0, 3, "=CHOOSE(A1,B1,C1)", None, 0)  # D1
    wb.close()


def test_choose_fractional_index_keeps_truncated_branch_live(tmp_path: Path) -> None:
    path = tmp_path / "choose_fractional.xlsx"
    _make_choose_cycle_workbook(path)

    graph = create_dependency_graph(path, ["Sheet1!D1"], load_values=False)

    b1_guard = graph.get_edge_guard("Sheet1!D1", "Sheet1!B1")
    c1_guard = graph.get_edge_guard("Sheet1!D1", "Sheet1!C1")
    for index, (b1_live, c1_live) in {
        1.0: (True, False),
        1.5: (True, False),
        2.0: (False, True),
        2.99: (False, True),
    }.items():
        values = {"Sheet1!A1": index}
        assert _guard_holds(b1_guard, values) is b1_live, index
        assert _guard_holds(c1_guard, values) is c1_live, index


def test_choose_fractional_index_domain_reports_cycle(tmp_path: Path) -> None:
    path = tmp_path / "choose_fractional_cycle.xlsx"
    _make_choose_cycle_workbook(path)
    constraints = {"Sheet1!A1": Annotated[float, RealBetween(1.2, 1.8)]}

    graph = create_dependency_graph(
        path,
        ["Sheet1!D1"],
        load_values=False,
        dynamic_refs=DynamicRefConfig.from_constraints(constraints),
    )
    report = graph.cycle_report()

    assert report.has_may_cycles is True
    assert any({"Sheet1!B1", "Sheet1!D1"} <= scc for scc in report.may_cycles)


def test_choose_integer_index_domain_keeps_equality_pruning(tmp_path: Path) -> None:
    path = tmp_path / "choose_integer_cycle.xlsx"
    _make_choose_cycle_workbook(path)

    for domain in (Annotated[int, Between(2, 3)], TypingLiteral[2, 3]):
        constraints = {"Sheet1!A1": domain}
        graph = create_dependency_graph(
            path,
            ["Sheet1!D1"],
            load_values=False,
            dynamic_refs=DynamicRefConfig.from_constraints(constraints),
        )

        assert graph.get_edge_guard("Sheet1!D1", "Sheet1!B1") == Compare(
            left=CellRef(key="Sheet1!A1"), op="=", right=Literal(value=1)
        )
        report = graph.cycle_report()
        assert report.has_may_cycles is False, domain
