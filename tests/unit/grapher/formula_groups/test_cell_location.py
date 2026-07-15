"""Unit tests for locating cells within graph nodes."""

from __future__ import annotations

import pytest

from excel_grapher.core.address_keys import parse_node_key
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import (
    CellLocation,
    Node,
    NodeKind,
    locate_cell,
    make_union_node,
)


def _cell(sheet: str, col: str, row: int) -> Node:
    return Node(sheet, col, row, None, None, 0, True)


def _row_span(start: str = "D", end: str = "Y", row: int = 63) -> Node:
    return Node(
        sheet=None,
        column=None,
        row=None,
        formula=None,
        normalized_formula=None,
        value=None,
        is_leaf=True,
        address=parse_node_key(f"Sheet1!{start}{row}:{end}{row}"),
    )


def test_locate_cell_as_cell_node() -> None:
    g = DependencyGraph()
    g.add_node(_cell("Sheet1", "A", 1))

    loc = locate_cell(g, "Sheet1!A1")
    assert loc == CellLocation(
        cell_key="Sheet1!A1",
        kind=NodeKind.cell,
        node_key="Sheet1!A1",
        column="A",
    )


def test_locate_cell_inside_one_row_union_node() -> None:
    g = DependencyGraph()
    g.add_node(_row_span())

    loc = locate_cell(g, "Sheet1!E63")
    assert loc == CellLocation(
        cell_key="Sheet1!E63",
        kind=NodeKind.union,
        node_key="Sheet1!D63:Y63",
        column="E",
    )
    assert locate_cell(g, "Sheet1!$E$63") == loc


def test_locate_cell_inside_sparse_union_node() -> None:
    g = DependencyGraph()
    union = make_union_node(["Sheet1!D63", "Sheet1!F63", "Sheet1!Y63"])
    g.add_node(union)

    loc = locate_cell(g, "Sheet1!F63")
    assert loc is not None
    assert loc.kind is NodeKind.union
    assert loc.node_key == "Sheet1!D63,F63,Y63"


def test_locate_cell_missing() -> None:
    g = DependencyGraph()
    g.add_node(_row_span())
    assert locate_cell(g, "Sheet1!A1") is None
    assert locate_cell(g, "Sheet1!C63") is None


def test_locate_cell_rejects_range_key() -> None:
    g = DependencyGraph()
    g.add_node(_row_span())
    with pytest.raises(ValueError, match="single-cell|single cell|cell key"):
        locate_cell(g, "Sheet1!D63:Y63")


def test_locate_cell_rejects_duplicate_occupancy_cell_and_union() -> None:
    g = DependencyGraph()
    g.add_node(_row_span())
    g._nodes["Sheet1!E63"] = _cell("Sheet1", "E", 63)
    g._edges.setdefault("Sheet1!E63", set())
    g._reverse_edges.setdefault("Sheet1!E63", set())

    with pytest.raises(ValueError, match="unique|occupancy|duplicate"):
        locate_cell(g, "Sheet1!E63")


def test_add_node_rejects_overlapping_union_nodes() -> None:
    g = DependencyGraph()
    g.add_node(_row_span("D", "M"))
    with pytest.raises(ValueError, match="occupancy|already|owned|conflict"):
        g.add_node(_row_span("E", "Y"))
