"""Tests for TACO RR (relative-relative) pattern compression."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node
from excel_grapher.grapher.range_compression import (
    PatternKind,
    RangeRef,
    TacoIndex,
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
    if sheet.startswith("'"):
        sheet = sheet[1:-1]
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


def test_rr_pattern_materialize_precedent() -> None:
    dep = RangeRef(sheet="Sheet1", min_col="D", min_row=3, max_col="D", max_row=7)
    prec = RangeRef(sheet="Sheet1", min_col="B", min_row=3, max_col="B", max_row=7)
    assert rr_materialize_precedent(dep, prec, "Sheet1!D5") == "Sheet1!B5"


def test_rr_manual_graph_parity() -> None:
    graph = DependencyGraph()
    for row in range(3, 8):
        graph.add_node(
            _make_node(f"Sheet1!B{row}", formula=None, is_leaf=True),
        )
        graph.add_node(
            _make_node(f"Sheet1!C{row}", formula=None, is_leaf=True),
        )
        graph.add_node(_make_node(f"Sheet1!D{row}", formula=f"=B{row}*C{row}"))
        dr = DependencyCause.direct_ref
        graph.add_edge(
            f"Sheet1!D{row}",
            f"Sheet1!B{row}",
            provenance=EdgeProvenance(causes=frozenset({dr})),
        )
        graph.add_edge(
            f"Sheet1!D{row}",
            f"Sheet1!C{row}",
            provenance=EdgeProvenance(causes=frozenset({dr})),
        )

    index = build_taco_index(graph)
    assert isinstance(index, TacoIndex)
    rr_edges = [e for e in index.compressed_edges if e.meta.kind == PatternKind.rr]
    assert len(rr_edges) == 2
    assert_taco_parity(graph, index)


def test_rr_synthetic_workbook_parity(tmp_path: Path) -> None:
    path = tmp_path / "rr_column.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Data")
    first, last = 3, 12
    for row in range(first, last + 1):
        ws.write_number(row - 1, 1, float(row))  # col B
        ws.write_number(row - 1, 2, float(row * 10))  # col C
        ws.write_formula(row - 1, 3, f"=B{row}*C{row}")
    wb.close()

    graph = create_dependency_graph(path, [f"Data!D{first}:D{last}"], load_values=False)
    index = build_taco_index(graph)
    assert_taco_parity(graph, index)
    assert len(index.compressed_edges) >= 2
    cell_edges = sum(
        len(graph.get_dependencies(k))
        for k in graph
        if (node := graph.get_node(k)) is not None and not node.is_leaf
    )
    assert len(index.compressed_edges) < cell_edges


def test_build_taco_index_does_not_mutate_graph() -> None:
    graph = DependencyGraph()
    graph.add_node(_make_node("Sheet1!B3", formula=None, is_leaf=True))
    graph.add_node(_make_node("Sheet1!D3", formula="=B3"))
    dr = DependencyCause.direct_ref
    graph.add_edge(
        "Sheet1!D3",
        "Sheet1!B3",
        provenance=EdgeProvenance(causes=frozenset({dr})),
    )
    snapshot = len(graph), frozenset(graph), graph.get_dependencies("Sheet1!D3")
    build_taco_index(graph)
    assert (len(graph), frozenset(graph), graph.get_dependencies("Sheet1!D3")) == snapshot


def test_find_dependents_returns_range_ref() -> None:
    graph = DependencyGraph()
    for row in range(3, 6):
        graph.add_node(_make_node(f"Sheet1!B{row}", formula=None, is_leaf=True))
        graph.add_node(_make_node(f"Sheet1!D{row}", formula=f"=B{row}"))
        graph.add_edge(
            f"Sheet1!D{row}",
            f"Sheet1!B{row}",
            provenance=EdgeProvenance(causes=frozenset({DependencyCause.direct_ref})),
        )
    index = build_taco_index(graph)
    deps = index.find_dependents("Sheet1!B4")
    assert len(deps) == 1
    assert deps[0] == RangeRef(sheet="Sheet1", min_col="D", min_row=3, max_col="D", max_row=5)
    assert materialize_dependents(index, "Sheet1!B4") == {"Sheet1!D4"}


def test_find_precedents_returns_range_ref() -> None:
    graph = DependencyGraph()
    for row in range(3, 6):
        graph.add_node(_make_node(f"Sheet1!B{row}", formula=None, is_leaf=True))
        graph.add_node(_make_node(f"Sheet1!D{row}", formula=f"=B{row}"))
        graph.add_edge(
            f"Sheet1!D{row}",
            f"Sheet1!B{row}",
            provenance=EdgeProvenance(causes=frozenset({DependencyCause.direct_ref})),
        )
    index = build_taco_index(graph)
    precs = index.find_precedents("Sheet1!D4")
    assert len(precs) == 1
    assert precs[0] == RangeRef(sheet="Sheet1", min_col="B", min_row=3, max_col="B", max_row=5)
    assert materialize_precedents(index, "Sheet1!D4") == {"Sheet1!B4"}
