"""Grapher guard-extraction gaps."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.guard import And, CellRef, Compare, Literal, Not
from tests.integration.user_flows.utils import (
    WorkbookFactory,
    build_workbook_factory,
    write_single_row,
)


@pytest.fixture
def workbook_factory(tmp_path: Path) -> WorkbookFactory:
    return build_workbook_factory(tmp_path, prefix="nested_guard_gap")


def test_nested_conditional_in_cell_does_not_consolidate_and_guard_on_b1(
    workbook_factory: WorkbookFactory,
) -> None:
    """Nested ``IF`` should AND-combine ``B1`` guards; today only ``C1`` gets ``NOT(B1=1)``."""
    path = workbook_factory(
        lambda ws, _wb: write_single_row(
            ws, ("Nested conditional in a cell", 0, 10, "=IF(NOT(B1=1),IF(B1=0,C1,1),0)")
        )
    )
    graph = create_dependency_graph(path, ["Sheet1!D1"], load_values=True)

    c1_guard = graph.get_edge_guard("Sheet1!D1", "Sheet1!C1")
    b1_guard = graph.get_edge_guard("Sheet1!D1", "Sheet1!B1")

    assert c1_guard == Not(
        operand=Compare(left=CellRef(key="Sheet1!B1"), op="=", right=Literal(value=1))
    )
    assert b1_guard is None
    assert c1_guard != And(
        operands=(
            Not(operand=Compare(left=CellRef(key="Sheet1!B1"), op="=", right=Literal(value=1))),
            Compare(left=CellRef(key="Sheet1!B1"), op="=", right=Literal(value=0)),
        )
    )
