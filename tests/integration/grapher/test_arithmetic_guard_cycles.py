"""Debt-schedule guards with arithmetic operands resolve false may-cycles (#1029)."""

from __future__ import annotations

from pathlib import Path
from typing import Annotated

import xlsxwriter

from excel_grapher import DynamicRefConfig, create_dependency_graph
from excel_grapher.core.cell_types import Between
from excel_grapher.grapher.guard import And, Arith, CellRef, Compare


def _make_schedule_workbook(path: Path) -> None:
    """P1 pays tranche year A2 in year A1 and vice versa; G1 is grace, H1 maturity."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 2020)  # A1
    ws.write_number(1, 0, 2021)  # A2
    ws.write_number(0, 6, 1)  # G1
    ws.write_number(0, 7, 10)  # H1
    ws.write_formula(0, 15, "=IF(AND(A1>A2+G1,A1<=A2+H1),P2/(H1-G1),0)", None, 0)  # P1
    ws.write_formula(1, 15, "=IF(AND(A2>A1+G1,A2<=A1+H1),P1/(H1-G1),0)", None, 0)  # P2
    wb.close()


def _graph(path: Path, grace_min: int):
    constraints = {"Sheet1!G1": Annotated[int, Between(grace_min, 50)]}
    return create_dependency_graph(
        path,
        ["Sheet1!P1"],
        load_values=False,
        dynamic_refs=DynamicRefConfig.from_constraints(constraints),
    )


def test_schedule_branch_edge_carries_arithmetic_guard(tmp_path: Path) -> None:
    path = tmp_path / "schedule.xlsx"
    _make_schedule_workbook(path)
    graph = _graph(path, 0)

    a1, a2, g1, h1 = (CellRef(f"Sheet1!{c}") for c in ("A1", "A2", "G1", "H1"))
    assert graph.get_edge_guard("Sheet1!P1", "Sheet1!P2") == And(
        (
            Compare(left=a1, op=">", right=Arith(a2, "+", g1)),
            Compare(left=a1, op="<=", right=Arith(a2, "+", h1)),
        )
    )


def test_schedule_cycle_is_infeasible_only_for_non_negative_grace(tmp_path: Path) -> None:
    path = tmp_path / "schedule.xlsx"
    _make_schedule_workbook(path)

    assert _graph(path, 0).cycle_report().has_may_cycles is False
    assert _graph(path, -5).cycle_report().has_may_cycles is True
