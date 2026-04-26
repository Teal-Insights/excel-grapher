"""
Tests that ``grapher`` correctly and reliably extracts cells, dependencies,
and guards for different cases, including nested conditionals and cycles.

Test cases (roadmap; only the first is implemented here):
[
    ("Formula with no dependencies","=1+1"),
    ("Linear dependency",1,"=B1+1"),
    ("Conditional branches",1,10,20,"=IF(B1=1,C1,D1)"),
    ("Nested conditional in a cell",0,10,"=IF(NOT(B1=1),IF(B1=0,C1,1),0)"),
    ("Nested conditional across cells",1,1,"=IF(B1=0,B1,2)","=IF(A1=1,B1,C1)"),
    ("Will cycle","=C1+1","=B1+1"),
    ("Won't cycle","=IF(B1=0,1,D1)","=IF(NOT(B1=0),2,C1)"),
    ("May cycle","=IF(B1=0,1,D1)","=IF(B1=1,2,C1)"),
]
"""

from __future__ import annotations

from contextlib import suppress
import pytest
from typing import Callable, Generator
from pathlib import Path

import xlsxwriter

from excel_grapher.grapher import DependencyGraph, create_dependency_graph
from excel_grapher.grapher.node import NodeKey, NodeView


@pytest.fixture
def workbook_path_factory(
    tmp_path: Path,
) -> Generator[Callable[[tuple[int | float | str, ...]], Path], None, None]:
    created_paths: list[Path] = []

    def _create_workbook_path(row_cells: tuple[int | float | str, ...]) -> Path:
        path = tmp_path / f"book_{len(created_paths)}.xlsx"
        wb = xlsxwriter.Workbook(path)
        ws = wb.add_worksheet("Sheet1")
        _populate_single_row(ws, row_cells)
        wb.close()
        created_paths.append(path)
        return path

    try:
        yield _create_workbook_path
    finally:
        for path in created_paths:
            with suppress(OSError):
                path.unlink()


def _populate_single_row(
    ws: xlsxwriter.Worksheet,
    row_cells: tuple[int | float | str, ...]
):
    for col, cell in enumerate(row_cells):
        if isinstance(cell, str) and cell.startswith("="):
            # Write formulas with cached value 0 for purposes of this test
            ws.write_formula(0, col, cell, None, 0)
        elif isinstance(cell, str):
            ws.write_string(0, col, cell)
        elif isinstance(cell, (float, int)):
            ws.write_number(0, col, cell)


@pytest.mark.xfail(
        reason="This test is expected to fail until business logic for identifying"
         "leaf nodes is successfully updated from 'not a formula' to 'no dependencies'."
        )
def test_formula_with_no_dependencies_is_extracted_as_single_formula_leaf_node(
    workbook_path_factory: Callable[[tuple[int | float | str, ...]], Path]
) -> None:
    path = workbook_path_factory(("Formula with no dependencies", "=1+1"))
    graph: DependencyGraph = create_dependency_graph(path, ["Sheet1!B1"], load_values=True)
    assert len(graph._nodes) == 1

    node: NodeView | None = graph.get_node("Sheet1!B1")
    assert node is not None
    assert node.sheet == "Sheet1"
    assert node.column == "B"
    assert node.row == 1
    assert node.formula == "=1+1"
    assert node.normalized_formula == "=1+1"
    assert node.value == 0
    # Formula with no dependencies must be a leaf node
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


def test_linear_dependency_is_extracted_as_two_nodes_with_one_edge(
    workbook_path_factory: Callable[[tuple[int | float | str, ...]], Path]
) -> None:
    path = workbook_path_factory(("Linear dependency", 1, "=B1+1"))
    graph: DependencyGraph = create_dependency_graph(path, ["Sheet1!C1"], load_values=True)
    assert len(graph._nodes) == 2

    target_node: NodeView | None = graph.get_node("Sheet1!C1")
    assert target_node is not None
    assert target_node.sheet == "Sheet1"
    assert target_node.column == "C"
    assert target_node.row == 1
    assert target_node.formula == "=B1+1"
    assert target_node.normalized_formula == "=Sheet1!B1+1"
    assert target_node.value == 0
    assert target_node.is_leaf is False
    assert dict(target_node.metadata) == {}

    leaf_node: NodeView | None = graph.get_node("Sheet1!B1")
    assert leaf_node is not None
    assert leaf_node.sheet == "Sheet1"
    assert leaf_node.column == "B"
    assert leaf_node.row == 1
    assert leaf_node.formula is None
    assert leaf_node.normalized_formula is None
    assert leaf_node.value == 1
    assert leaf_node.is_leaf is True
    assert dict(leaf_node.metadata) == {}

    dependencies: frozenset[NodeKey] = graph.get_dependencies("Sheet1!C1")
    assert dependencies == frozenset(["Sheet1!B1"])
    dependents: frozenset[NodeKey] = graph.get_dependents("Sheet1!C1")
    assert dependents == frozenset()
    assert not graph._guards
    assert not graph._edge_extra
    assert not graph._hooks
    assert graph.leaf_classification is None