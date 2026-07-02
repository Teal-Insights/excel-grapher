"""Tests for TACO RR pattern compression along rows (fill-right)."""

from __future__ import annotations

from pathlib import Path

import fastpyxl.utils.cell
import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node
from excel_grapher.grapher.range_compression import (
    PatternKind,
    RangeRef,
    build_taco_index,
    materialize_dependents,
    materialize_precedents,
)
from excel_grapher.grapher.range_compression.patterns import rr_materialize_precedent

from .parity_helpers import assert_taco_parity


def _make_node(
    key: str,
    formula: str | None,
    *,
    is_leaf: bool = False,
) -> Node:
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


def test_rr_row_span_materialize_precedent() -> None:
    dep = RangeRef.row_span("Sheet1", 9, "F", "J")
    prec = RangeRef.row_span("Sheet1", 9, "D", "H")
    assert rr_materialize_precedent(dep, prec, "Sheet1!H9") == "Sheet1!F9"


def test_rr_row_fill_right_manual_graph() -> None:
    graph = DependencyGraph()
    dr = DependencyCause.direct_ref
    for col_i in range(4, 11):  # D through J precedents/leaves
        prec_col = fastpyxl.utils.cell.get_column_letter(col_i)
        if col_i <= 8:
            graph.add_node(_make_node(f"Sheet1!{prec_col}9", formula=None, is_leaf=True))
    for col_i in range(6, 11):  # F through J dependents
        dep_col = fastpyxl.utils.cell.get_column_letter(col_i)
        prec_col = fastpyxl.utils.cell.get_column_letter(col_i - 2)
        graph.add_node(_make_node(f"Sheet1!{dep_col}9", formula=f"={prec_col}9"))
        graph.add_edge(
            f"Sheet1!{dep_col}9",
            f"Sheet1!{prec_col}9",
            provenance=EdgeProvenance(causes=frozenset({dr})),
        )

    index = build_taco_index(graph)
    rr_edges = [e for e in index.compressed_edges if e.meta.kind == PatternKind.rr]
    assert len(rr_edges) == 1
    edge = rr_edges[0]
    assert edge.dependent == RangeRef.row_span("Sheet1", 9, "F", "J")
    assert edge.precedent == RangeRef.row_span("Sheet1", 9, "D", "H")
    assert_taco_parity(graph, index)


def test_rr_row_fill_right_workbook_parity(tmp_path: Path) -> None:
    path = tmp_path / "rr_row.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Data")
    row = 9
    first_col, last_col = 6, 10  # F through J
    for col_i in range(first_col - 2, last_col):
        ws.write_number(row - 1, col_i - 1, float(col_i))
    for col_i in range(first_col, last_col + 1):
        prec_col = fastpyxl.utils.cell.get_column_letter(col_i - 2)
        ws.write_formula(row - 1, col_i - 1, f"={prec_col}{row}")
    wb.close()

    first = fastpyxl.utils.cell.get_column_letter(first_col)
    last = fastpyxl.utils.cell.get_column_letter(last_col)
    graph = create_dependency_graph(path, [f"Data!{first}{row}:{last}{row}"], load_values=False)
    index = build_taco_index(graph)
    rr_edges = [e for e in index.compressed_edges if e.meta.kind == PatternKind.rr]
    assert len(rr_edges) == 1
    assert rr_edges[0].dependent == RangeRef.row_span("Data", row, first, last)
    assert_taco_parity(graph, index)


def test_rr_row_find_precedents_and_dependents() -> None:
    graph = DependencyGraph()
    dr = DependencyCause.direct_ref
    for col_i in range(4, 11):
        prec_col = fastpyxl.utils.cell.get_column_letter(col_i)
        if col_i <= 8:
            graph.add_node(_make_node(f"Sheet1!{prec_col}9", formula=None, is_leaf=True))
    for col_i in range(6, 11):
        dep_col = fastpyxl.utils.cell.get_column_letter(col_i)
        prec_col = fastpyxl.utils.cell.get_column_letter(col_i - 2)
        graph.add_node(_make_node(f"Sheet1!{dep_col}9", formula=f"={prec_col}9"))
        graph.add_edge(
            f"Sheet1!{dep_col}9",
            f"Sheet1!{prec_col}9",
            provenance=EdgeProvenance(causes=frozenset({dr})),
        )

    index = build_taco_index(graph)
    assert index.find_precedents("Sheet1!H9") == [RangeRef.row_span("Sheet1", 9, "D", "H")]
    assert index.find_dependents("Sheet1!F9") == [RangeRef.row_span("Sheet1", 9, "F", "J")]
    assert materialize_precedents(index, "Sheet1!H9") == {"Sheet1!F9"}
    assert materialize_dependents(index, "Sheet1!F9") == {"Sheet1!H9"}
