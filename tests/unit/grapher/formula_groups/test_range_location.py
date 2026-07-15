"""Unit tests for locating one-row ranges within graph nodes."""

from __future__ import annotations

import pytest

from excel_grapher.core.address_keys import parse_node_key
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import (
    Node,
    NodeKind,
    RangeLocation,
    locate_range,
)


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


def test_locate_range_exact_one_row_union_node() -> None:
    g = DependencyGraph()
    g.add_node(_row_span())

    loc = locate_range(g, "Sheet1!Y63:D63")
    assert loc == RangeLocation(
        range_key="Sheet1!D63:Y63",
        kind=NodeKind.union,
        node_key="Sheet1!D63:Y63",
        min_col="D",
        max_col="Y",
        row=63,
    )


def test_locate_range_subspan_inside_one_row_union_node() -> None:
    g = DependencyGraph()
    g.add_node(_row_span())

    loc = locate_range(g, "Sheet1!E63:G63")
    assert loc == RangeLocation(
        range_key="Sheet1!E63:G63",
        kind=NodeKind.union,
        node_key="Sheet1!D63:Y63",
        min_col="E",
        max_col="G",
        row=63,
    )


def test_locate_range_missing() -> None:
    g = DependencyGraph()
    g.add_node(_row_span())
    assert locate_range(g, "Sheet1!C63:E63") is None


def test_locate_range_rejects_cell_or_collapsed_one_by_one_key() -> None:
    g = DependencyGraph()
    g.add_node(_row_span())
    for key in ("Sheet1!E63", "Sheet1!E63:E63"):
        with pytest.raises(ValueError, match="one-row|range"):
            locate_range(g, key)


def test_locate_range_rejects_overlapping_owners() -> None:
    g = DependencyGraph()
    g.add_node(_row_span("D", "M"))
    g._nodes["Sheet1!E63:Y63"] = _row_span("E", "Y")
    g._edges.setdefault("Sheet1!E63:Y63", set())
    g._reverse_edges.setdefault("Sheet1!E63:Y63", set())

    with pytest.raises(ValueError, match="unique|occupancy|overlap|duplicate"):
        locate_range(g, "Sheet1!E63:G63")
