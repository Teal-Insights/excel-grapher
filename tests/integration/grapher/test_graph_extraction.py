"""
Tests that ``grapher`` correctly and reliably extracts cells, dependencies,
and guards for different cases, including nested conditionals and cycles.

Test cases (roadmap; only the first is implemented here):
[
    ("Formula with no dependencies","=1+1"),
    ("Linear dependency","=B1+1"),
    ("Conditional branches","=IF(B1=1,C1,D1)"),
    ("Nested conditional in a cell","=IF(NOT(B1=1),IF(B1=0,C1,1),0)"),
    ("Nested conditional across cells","=IF(B1=1,C1,D1)"),
    ("Will cycle","=C1+1","=B1+1"),
    ("Won't cycle","=IF(B1=0,1,D1)","=IF(NOT(B1=0),2,C1)"),
    ("May cycle","=IF(B1=0,1,D1)","=IF(B1=1,2,C1)"),
]
"""

from __future__ import annotations

import pytest

from pathlib import Path

import xlsxwriter

from excel_grapher.grapher import DependencyGraph, create_dependency_graph
from excel_grapher.grapher.node import NodeKey, NodeView


def _make_single_row_workbook(
    path: Path,
    row_cells: tuple[int | float | str, ...]
) -> None:
    """Create a one-sheet workbook whose first row is populated for tests."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    for col, cell in enumerate(row_cells):
        if isinstance(cell, str) and cell.startswith("="):
            ws.write_formula(0, col, cell, None, 0)
        elif isinstance(cell, str):
            ws.write_string(0, col, cell)
        elif isinstance(cell, (float, int)):
            ws.write_number(0, col, cell)
    wb.close()


@pytest.mark.xfail(
        reason="This test is expected to fail until business logic for identifying"
         "leaf nodes is successfully updated from 'not a formula' to 'no dependencies'."
        )
def test_formula_with_no_dependencies_is_extracted_as_single_formula_node(tmp_path: Path) -> None:
    workbook: Path = tmp_path / "book.xlsx"
    _make_single_row_workbook(
        workbook, ("Formula with no dependencies", "=1+1")
    )
    graph: DependencyGraph = create_dependency_graph(workbook, ["Sheet1!B1"], load_values=True)
    assert len(graph._nodes) == 1

    node: NodeView | None = graph.get_node("Sheet1!B1")
    assert node is not None
    # Any cell with a leading "=" is a formula node, even if it has no dep edges.
    assert node.sheet == "Sheet1"
    assert node.column == "B"
    assert node.row == 1
    assert node.formula == "=1+1"
    assert node.normalized_formula == "=1+1"
    assert node.value == 2
    # IMPORTANT! DO NOT MODIFY THIS ASSERTION!
    # Formula with no dependencies is a leaf node!
    assert node.is_leaf is True
    assert dict(node.metadata) == {}

    dependencies: frozenset[NodeKey] = graph.get_dependencies("Sheet1!B1")
    assert dependencies == frozenset()
    dependents: frozenset[NodeKey] = graph.get_dependents("Sheet1!B1")
    assert dependents == frozenset()
    assert not graph._guards
    assert not graph._edge_extra
    assert not graph._hooks
    assert graph.leaf_classification is None
