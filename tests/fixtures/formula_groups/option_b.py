"""Hand-built Option B formula-group fixtures (Issue 2 sprint 2).

Graphs contain a multi-cell group node plus leaf precedents. Member cells are
owned by the group (no member cell nodes). Cell-only twins mirror the same
public formulas as discrete cells for later eval/codegen parity.
"""

from __future__ import annotations

from dataclasses import dataclass

from excel_grapher.core.formula_ast import (
    AddressHoleNode,
    AddressLeafKind,
    AstNode,
    CellRefNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
)
from excel_grapher.grapher.formula_groups import shape_fingerprint
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node, make_cell_node, make_union_node, member_keys


def index_match_skeleton() -> AstNode:
    """INDEX/MATCH template with baked ranges and one `CELL` hole (lookup)."""
    return FunctionCallNode(
        name="INDEX",
        args=[
            RangeNode(start="Sheet1!D40", end="Sheet1!AJ50"),
            FunctionCallNode(
                name="MATCH",
                args=[
                    NumberNode(value=1.0),
                    RangeNode(start="Sheet1!AJ40", end="Sheet1!AJ50"),
                    NumberNode(value=0.0),
                ],
            ),
            FunctionCallNode(
                name="MATCH",
                args=[
                    AddressHoleNode(kind=AddressLeafKind.cell, slot=0),
                    RangeNode(start="Sheet1!D35", end="Sheet1!Y35"),
                    NumberNode(value=0.0),
                ],
            ),
        ],
    )


def index_match_fingerprint() -> str:
    """Fingerprint for `index_match_skeleton`."""
    return shape_fingerprint(index_match_skeleton())


def specialized_index_match_formula(lookup_cell: str) -> str:
    """Normalized formula string matching the specialized INDEX/MATCH AST."""
    return (
        f"=INDEX(Sheet1!D40:AJ50,MATCH(1,Sheet1!AJ40:AJ50,0),MATCH({lookup_cell},Sheet1!D35:Y35,0))"
    )


def _add_shared_precedents(graph: DependencyGraph) -> None:
    """Add leaf cells referenced by the INDEX/MATCH template ranges."""
    # Minimal leaves so occupancy/edges are realistic without filling whole blocks.
    for key, value in (
        ("Sheet1!D40", 10.0),
        ("Sheet1!AJ40", 1.0),
        ("Sheet1!AJ50", 1.0),
        ("Sheet1!D35", "D"),
        ("Sheet1!E35", "E"),
        ("Sheet1!F35", "F"),
        ("Sheet1!Y35", "Y"),
        ("Sheet2!Z9", "Z"),
    ):
        sheet, cell = key.split("!")
        col = "".join(ch for ch in cell if ch.isalpha())
        row = int("".join(ch for ch in cell if ch.isdigit()))
        graph.add_node(make_cell_node(sheet, col, row, value=value, is_leaf=True))


@dataclass(frozen=True)
class OptionBFixture:
    """Option B graph plus the owning group node key."""

    graph: DependencyGraph
    group_key: str
    members: tuple[str, ...]


def build_row_stripe_option_b() -> OptionBFixture:
    """Contiguous one-row stripe `Sheet1!D63:F63` with INDEX/MATCH template."""
    members = ("Sheet1!D63", "Sheet1!E63", "Sheet1!F63")
    skeleton = index_match_skeleton()
    fp = shape_fingerprint(skeleton)
    bindings = {
        "Sheet1!D63": (CellRefNode(address="Sheet1!D35"),),
        "Sheet1!E63": (CellRefNode(address="Sheet1!E35"),),
        "Sheet1!F63": (CellRefNode(address="Sheet1!F35"),),
    }
    group = make_union_node(
        members,
        is_leaf=False,
        shape_fingerprint=fp,
        skeleton=skeleton,
        member_bindings=bindings,
    )
    g = DependencyGraph()
    g.sheet_order = ["Sheet1", "Sheet2"]
    _add_shared_precedents(g)
    g.add_node(group)
    for dep in (
        "Sheet1!D40",
        "Sheet1!AJ40",
        "Sheet1!AJ50",
        "Sheet1!D35",
        "Sheet1!E35",
        "Sheet1!F35",
    ):
        g.add_edge(group.key, dep)
    for m in members:
        assert g.get_node(m) is None
        assert g.cell_owner(m) == group.key
    return OptionBFixture(graph=g, group_key=group.key, members=members)


def build_cross_sheet_union_option_b() -> OptionBFixture:
    """Non-contiguous cross-sheet union with the same INDEX/MATCH template."""
    members = ("Sheet1!D63", "Sheet2!B10")
    skeleton = index_match_skeleton()
    fp = shape_fingerprint(skeleton)
    bindings = {
        "Sheet1!D63": (CellRefNode(address="Sheet1!D35"),),
        "Sheet2!B10": (CellRefNode(address="Sheet2!Z9"),),
    }
    group = make_union_node(
        members,
        is_leaf=False,
        shape_fingerprint=fp,
        skeleton=skeleton,
        member_bindings=bindings,
    )
    g = DependencyGraph()
    g.sheet_order = ["Sheet1", "Sheet2"]
    _add_shared_precedents(g)
    g.add_node(group)
    for dep in ("Sheet1!D40", "Sheet1!AJ40", "Sheet1!AJ50", "Sheet1!D35", "Sheet2!Z9"):
        g.add_edge(group.key, dep)
    for m in members:
        assert g.get_node(m) is None
        assert g.cell_owner(m) == group.key
    return OptionBFixture(graph=g, group_key=group.key, members=members)


def build_row_stripe_cell_only_twin() -> DependencyGraph:
    """Cell-only twin of `build_row_stripe_option_b` (no multi-cell node)."""
    g = DependencyGraph()
    g.sheet_order = ["Sheet1", "Sheet2"]
    _add_shared_precedents(g)
    for member, lookup in (
        ("Sheet1!D63", "Sheet1!D35"),
        ("Sheet1!E63", "Sheet1!E35"),
        ("Sheet1!F63", "Sheet1!F35"),
    ):
        sheet, cell = member.split("!")
        col = "".join(ch for ch in cell if ch.isalpha())
        row = int("".join(ch for ch in cell if ch.isdigit()))
        formula = specialized_index_match_formula(lookup)
        node = make_cell_node(
            sheet,
            col,
            row,
            formula=formula,
            normalized_formula=formula,
            is_leaf=False,
        )
        g.add_node(node)
        for dep in ("Sheet1!D40", "Sheet1!AJ40", "Sheet1!AJ50", lookup):
            g.add_edge(member, dep)
    return g


def build_cross_sheet_cell_only_twin() -> DependencyGraph:
    """Cell-only twin of `build_cross_sheet_union_option_b`."""
    g = DependencyGraph()
    g.sheet_order = ["Sheet1", "Sheet2"]
    _add_shared_precedents(g)
    for member, lookup in (
        ("Sheet1!D63", "Sheet1!D35"),
        ("Sheet2!B10", "Sheet2!Z9"),
    ):
        sheet, cell = member.split("!")
        col = "".join(ch for ch in cell if ch.isalpha())
        row = int("".join(ch for ch in cell if ch.isdigit()))
        formula = specialized_index_match_formula(lookup)
        g.add_node(
            make_cell_node(
                sheet,
                col,
                row,
                formula=formula,
                normalized_formula=formula,
                is_leaf=False,
            )
        )
        for dep in ("Sheet1!D40", "Sheet1!AJ40", "Sheet1!AJ50", lookup):
            g.add_edge(member, dep)
    return g


def assert_option_b_occupancy(group: Node) -> None:
    """Assert group owns its members and has a validated template."""
    assert group.skeleton is not None
    assert group.member_bindings is not None
    assert group.shape_fingerprint is not None
    owned = set(member_keys(group))
    assert owned == set(group.member_bindings)
