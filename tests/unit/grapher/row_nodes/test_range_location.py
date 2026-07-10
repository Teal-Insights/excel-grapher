"""Unit tests for one-row subrange location within row nodes."""

from __future__ import annotations

import pytest

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import (
    Node,
    NodeKind,
    RangeLocation,
    find_row_nodes_covering,
    locate_range,
    make_row_node,
    row_node_covers_range,
)


def _cell(sheet: str, col: str, row: int) -> Node:
    return Node(sheet, col, row, None, None, 0, True)


def test_row_node_covers_range_subspan() -> None:
    row = make_row_node("Sheet1", 63, "D", "Y")
    assert row_node_covers_range(row, "Sheet1!E63:G63")
    assert row_node_covers_range(row, "Sheet1!D63:Y63")  # exact
    assert row_node_covers_range(row, "Sheet1!Y63:E63")  # inverted -> canonicalized
    assert row_node_covers_range(row, "Sheet1!D63:Sheet1!F63")
    assert row_node_covers_range(row, "Sheet1!$E$63:$G$63")


def test_row_node_covers_range_rejects_partial_or_outside() -> None:
    row = make_row_node("Sheet1", 63, "D", "Y")
    assert not row_node_covers_range(row, "Sheet1!C63:E63")  # starts outside
    assert not row_node_covers_range(row, "Sheet1!X63:Z63")  # ends outside
    assert not row_node_covers_range(row, "Sheet1!E64:G64")  # wrong row
    assert not row_node_covers_range(row, "Other!E63:G63")


def test_row_node_covers_range_rejects_multi_row() -> None:
    row = make_row_node("Sheet1", 63, "D", "Y")
    with pytest.raises(ValueError, match="one-row|same row"):
        row_node_covers_range(row, "Sheet1!D63:Y64")


def test_find_row_nodes_covering_accepts_subrange() -> None:
    g = DependencyGraph()
    g.add_node(make_row_node("Sheet1", 63, "D", "Y"))
    g.add_node(make_row_node("Sheet1", 64, "D", "Y"))

    assert find_row_nodes_covering(g, "Sheet1!E63:G63") == ["Sheet1!D63:Y63"]
    assert find_row_nodes_covering(g, "Sheet1!G63:E63") == ["Sheet1!D63:Y63"]
    assert find_row_nodes_covering(g, "Sheet1!E64:G64") == ["Sheet1!D64:Y64"]
    assert find_row_nodes_covering(g, "Sheet1!C63:E63") == []


def test_locate_range_exact_row_node() -> None:
    g = DependencyGraph()
    g.add_node(make_row_node("Sheet1", 63, "D", "Y"))

    loc = locate_range(g, "Sheet1!Y63:D63")
    assert loc == RangeLocation(
        range_key="Sheet1!D63:Y63",
        kind=NodeKind.row,
        node_key="Sheet1!D63:Y63",
        min_col="D",
        max_col="Y",
        row=63,
    )


def test_locate_range_subspan_inside_row_node() -> None:
    g = DependencyGraph()
    g.add_node(make_row_node("Sheet1", 63, "D", "Y"))

    loc = locate_range(g, "Sheet1!E63:G63")
    assert loc == RangeLocation(
        range_key="Sheet1!E63:G63",
        kind=NodeKind.row,
        node_key="Sheet1!D63:Y63",
        min_col="E",
        max_col="G",
        row=63,
    )


def test_locate_range_missing() -> None:
    g = DependencyGraph()
    g.add_node(make_row_node("Sheet1", 63, "D", "Y"))
    assert locate_range(g, "Sheet1!C63:E63") is None


def test_locate_range_rejects_cell_key() -> None:
    g = DependencyGraph()
    g.add_node(make_row_node("Sheet1", 63, "D", "Y"))
    with pytest.raises(ValueError, match="one-row|range"):
        locate_range(g, "Sheet1!E63")


def test_locate_range_rejects_overlapping_owners() -> None:
    g = DependencyGraph()
    g.add_node(make_row_node("Sheet1", 63, "D", "M"))
    g._nodes["Sheet1!E63:Y63"] = make_row_node("Sheet1", 63, "E", "Y")
    g._edges.setdefault("Sheet1!E63:Y63", set())
    g._reverse_edges.setdefault("Sheet1!E63:Y63", set())

    with pytest.raises(ValueError, match="unique|occupancy|overlap|duplicate"):
        locate_range(g, "Sheet1!E63:G63")
