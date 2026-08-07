"""Tests that excluded edges are not merged into compressed TACO patterns."""

from __future__ import annotations

from pathlib import Path

import pytest
import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.guard import CellRef as GuardCellRef
from excel_grapher.grapher.guard import Compare, Literal
from excel_grapher.grapher.node import Node
from excel_grapher.grapher.range_compression import PatternKind, build_taco_index

from .parity_helpers import assert_taco_parity


def _make_node(key: str, formula: str | None, *, is_leaf: bool = False) -> Node:
    sheet, rest = key.split("!", 1)
    col = "".join(c for c in rest if c.isalpha())
    row = int("".join(c for c in rest if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=None,
        is_leaf=is_leaf,
    )


def test_guarded_edge_stays_single() -> None:
    graph = DependencyGraph()
    for row in range(3, 6):
        graph.add_node(_make_node(f"Sheet1!B{row}", formula=None, is_leaf=True))
        graph.add_node(_make_node(f"Sheet1!D{row}", formula=f"=IF(A1,B{row},0)"))
        graph.add_edge(
            f"Sheet1!D{row}",
            f"Sheet1!B{row}",
            guard=Compare(GuardCellRef("Sheet1!A1"), ">", Literal(0)),
        )
    index = build_taco_index(graph)
    assert index.compressed_edges == []
    assert len(index.single_edges) > 0
    assert_taco_parity(graph, index)


def test_non_autofill_row_span_not_one_rr_group(tmp_path: Path) -> None:
    """A row formula referencing multiple columns must not become one compressed group."""
    path = tmp_path / "wide_row.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Demo")
    for col in range(3, 8):
        ws.write_number(2, col, 1.0)
    ws.write_formula(2, 8, "=D3+H3")
    wb.close()

    graph = create_dependency_graph(path, ["Demo!I3"], load_values=False)
    index = build_taco_index(graph)
    assert not any(e.meta.kind == PatternKind.rr for e in index.compressed_edges)
    assert_taco_parity(graph, index)


@pytest.mark.skip(reason="dynamic OFFSET workbooks require constraint config in CI")
def test_dynamic_offset_not_compressed(_tmp_path: Path) -> None:
    pytest.skip("dynamic OFFSET workbooks require constraint config in CI")


def test_static_range_one_off_sum_not_compressed() -> None:
    graph = DependencyGraph()
    for row in range(3, 6):
        graph.add_node(_make_node(f"Sheet1!E{row}", formula=None, is_leaf=True))
    graph.add_node(_make_node("Sheet1!F3", formula="=SUM(E3:E5)"))
    dr = DependencyCause.static_range
    for row in range(3, 6):
        graph.add_edge(
            "Sheet1!F3",
            f"Sheet1!E{row}",
            provenance=EdgeProvenance(causes=dr),
        )
    index = build_taco_index(graph)
    assert index.compressed_edges == []
    assert_taco_parity(graph, index)
