"""Unit tests for cell location within row nodes."""

from __future__ import annotations

import pytest

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import (
    CellLocation,
    Node,
    NodeKind,
    find_row_nodes_covering,
    locate_cell,
    make_row_node,
    row_node_covers_cell,
)


def _cell(sheet: str, col: str, row: int) -> Node:
    return Node(sheet, col, row, None, None, 0, True)


def test_row_node_covers_cell_inside_span() -> None:
    row = make_row_node("Sheet1", 63, "D", "Y")
    assert row_node_covers_cell(row, "Sheet1!D63")
    assert row_node_covers_cell(row, "Sheet1!E63")
    assert row_node_covers_cell(row, "Sheet1!Y63")
    assert row_node_covers_cell(row, "Sheet1!$E$63")


def test_row_node_covers_cell_outside_span() -> None:
    row = make_row_node("Sheet1", 63, "D", "Y")
    assert not row_node_covers_cell(row, "Sheet1!C63")
    assert not row_node_covers_cell(row, "Sheet1!Z63")
    assert not row_node_covers_cell(row, "Sheet1!E64")
    assert not row_node_covers_cell(row, "Other!E63")


def test_row_node_covers_cell_rejects_non_row() -> None:
    cell = _cell("Sheet1", "E", 63)
    with pytest.raises(ValueError, match="row"):
        row_node_covers_cell(cell, "Sheet1!E63")


def test_find_row_nodes_covering() -> None:
    g = DependencyGraph()
    g.add_node(make_row_node("Sheet1", 63, "D", "Y"))
    g.add_node(make_row_node("Sheet1", 64, "D", "Y"))
    g.add_node(_cell("Sheet1", "A", 1))

    assert find_row_nodes_covering(g, "Sheet1!E63") == ["Sheet1!D63:Y63"]
    assert find_row_nodes_covering(g, "Sheet1!E64") == ["Sheet1!D64:Y64"]
    assert find_row_nodes_covering(g, "Sheet1!A1") == []
    assert find_row_nodes_covering(g, "Sheet1!C63") == []


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


def test_locate_cell_inside_row_node() -> None:
    g = DependencyGraph()
    g.add_node(make_row_node("Sheet1", 63, "D", "Y"))

    loc = locate_cell(g, "Sheet1!E63")
    assert loc == CellLocation(
        cell_key="Sheet1!E63",
        kind=NodeKind.row,
        node_key="Sheet1!D63:Y63",
        column="E",
    )
    assert locate_cell(g, "Sheet1!$E$63") == loc


def test_locate_cell_missing() -> None:
    g = DependencyGraph()
    g.add_node(make_row_node("Sheet1", 63, "D", "Y"))
    assert locate_cell(g, "Sheet1!A1") is None
    assert locate_cell(g, "Sheet1!C63") is None


def test_locate_cell_rejects_row_key() -> None:
    g = DependencyGraph()
    g.add_node(make_row_node("Sheet1", 63, "D", "Y"))
    with pytest.raises(ValueError, match="single cell|cell key"):
        locate_cell(g, "Sheet1!D63:Y63")


def test_locate_cell_rejects_duplicate_occupancy_cell_and_row() -> None:
    g = DependencyGraph()
    # Bypass add_node checks by inserting directly if enforcement exists;
    # otherwise construct a violating graph via internal dict.
    g.add_node(make_row_node("Sheet1", 63, "D", "Y"))
    g._nodes["Sheet1!E63"] = _cell("Sheet1", "E", 63)
    g._edges.setdefault("Sheet1!E63", set())
    g._reverse_edges.setdefault("Sheet1!E63", set())

    with pytest.raises(ValueError, match="unique|occupancy|duplicate"):
        locate_cell(g, "Sheet1!E63")


def test_locate_cell_rejects_overlapping_row_nodes() -> None:
    g = DependencyGraph()
    g.add_node(make_row_node("Sheet1", 63, "D", "M"))
    g._nodes["Sheet1!E63:Y63"] = make_row_node("Sheet1", 63, "E", "Y")
    g._edges.setdefault("Sheet1!E63:Y63", set())
    g._reverse_edges.setdefault("Sheet1!E63:Y63", set())

    with pytest.raises(ValueError, match="unique|occupancy|overlap|duplicate"):
        locate_cell(g, "Sheet1!E63")
