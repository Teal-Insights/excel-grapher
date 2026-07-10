"""Option B row-node fixtures: shared template, no member cell nodes."""

from __future__ import annotations

from dataclasses import dataclass

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import (
    Node,
    NodeKind,
    locate_cell,
    make_row_node,
    row_member_keys,
)

# Tiny stripe used by eval/codegen parity later (Sprint 3+).
OPTION_B_SHEET = "Sheet1"
OPTION_B_ROW = 63
OPTION_B_MIN_COL = "D"
OPTION_B_MAX_COL = "E"
OPTION_B_TEMPLATE = "=Sheet1!D35*2"
OPTION_B_VARYING_REF_SLOTS: tuple[int, ...] = (0,)
OPTION_B_ROW_KEY = "Sheet1!D63:E63"


@dataclass(frozen=True, slots=True)
class OptionBStripeFixture:
    """Hand-built Option B graph plus its cell-only twin."""

    option_b: DependencyGraph
    cell_only: DependencyGraph
    row_key: str
    member_keys: tuple[str, ...]
    template: str
    varying_ref_slots: tuple[int, ...]


def _leaf(sheet: str, col: str, row: int, value: object) -> Node:
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=None,
        normalized_formula=None,
        value=value,
        is_leaf=True,
    )


def _formula_cell(sheet: str, col: str, row: int, formula: str) -> Node:
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=None,
        is_leaf=False,
        is_target=True,
    )


def assert_unique_occupancy_for_row(graph: DependencyGraph, row_key: str) -> None:
    """Assert no `kind=cell` node occupies a member address of `row_key`."""
    row = graph.get_node(row_key)
    if row is None or row.kind is not NodeKind.row:
        raise AssertionError(f"expected row node at {row_key!r}")
    for member in row_member_keys(row):
        node = graph.get_node(member)
        if node is not None and node.kind is NodeKind.cell:
            raise AssertionError(
                f"unique occupancy violated: cell node {member!r} inside row span {row_key!r}"
            )
        loc = locate_cell(graph, member)
        if loc is None or loc.kind is not NodeKind.row or loc.node_key != row.key:
            raise AssertionError(
                f"expected locate_cell({member!r}) to resolve to row {row_key!r}, got {loc!r}"
            )


def build_option_b_product_graph() -> DependencyGraph:
    """Build Option B graph: row template `=Sheet1!D35*2` over `D63:E63`, no member cells.

    Shared precedents `D35` / `E35` are leaf value nodes. The row node owns the
    template and `varying_ref_slots=(0,)`.
    """
    g = DependencyGraph()
    d35 = _leaf(OPTION_B_SHEET, "D", 35, 3)
    e35 = _leaf(OPTION_B_SHEET, "E", 35, 5)
    row = make_row_node(
        OPTION_B_SHEET,
        OPTION_B_ROW,
        OPTION_B_MIN_COL,
        OPTION_B_MAX_COL,
        formula=OPTION_B_TEMPLATE,
        normalized_formula=OPTION_B_TEMPLATE,
        varying_ref_slots=OPTION_B_VARYING_REF_SLOTS,
        is_leaf=False,
        is_target=True,
    )
    g.add_node(row)
    g.add_node(d35)
    g.add_node(e35)
    g.add_edge(row.key, d35.key)
    g.add_edge(row.key, e35.key)
    return g


def build_cell_only_product_twin() -> DependencyGraph:
    """Cell-only twin of `build_option_b_product_graph` (separate D63 / E63 nodes)."""
    g = DependencyGraph()
    d35 = _leaf(OPTION_B_SHEET, "D", 35, 3)
    e35 = _leaf(OPTION_B_SHEET, "E", 35, 5)
    d63 = _formula_cell(OPTION_B_SHEET, "D", 63, "=Sheet1!D35*2")
    e63 = _formula_cell(OPTION_B_SHEET, "E", 63, "=Sheet1!E35*2")
    g.add_node(d35)
    g.add_node(e35)
    g.add_node(d63)
    g.add_node(e63)
    g.add_edge(d63.key, d35.key)
    g.add_edge(e63.key, e35.key)
    return g


def build_option_b_stripe_fixture() -> OptionBStripeFixture:
    """Return Option B + cell-only twin graphs for the product stripe."""
    option_b = build_option_b_product_graph()
    cell_only = build_cell_only_product_twin()
    row = option_b.get_node(OPTION_B_ROW_KEY)
    assert row is not None
    members = tuple(row_member_keys(row))
    return OptionBStripeFixture(
        option_b=option_b,
        cell_only=cell_only,
        row_key=OPTION_B_ROW_KEY,
        member_keys=members,
        template=OPTION_B_TEMPLATE,
        varying_ref_slots=OPTION_B_VARYING_REF_SLOTS,
    )
