"""Hand-built formula-group fixtures (Issue 2).

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
    BinaryOpNode,
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
    """Add leaf cells so INDEX/MATCH over the template ranges can resolve."""
    import fastpyxl.utils.cell

    def _leaf(sheet: str, col: str, row: int, value: object) -> None:
        key = f"{sheet}!{col}{row}"
        if graph.get_node(key) is None:
            graph.add_node(make_cell_node(sheet, col, row, value=value, is_leaf=True))

    # Header row scanned by MATCH(..., D35:Y35, 0)
    d_idx = fastpyxl.utils.cell.column_index_from_string("D")
    y_idx = fastpyxl.utils.cell.column_index_from_string("Y")
    for col_i in range(d_idx, y_idx + 1):
        col = fastpyxl.utils.cell.get_column_letter(col_i)
        _leaf("Sheet1", col, 35, col)

    # Row-match column: AJ40 = 1 (first match); other AJ rows distinct
    for row in range(40, 51):
        _leaf("Sheet1", "AJ", row, 1.0 if row == 40 else 0.0)

    # INDEX result row (row 1 of D40:AJ50) for D/E/F columns used by fixtures
    _leaf("Sheet1", "D", 40, 10.0)
    _leaf("Sheet1", "E", 40, 11.0)
    _leaf("Sheet1", "F", 40, 12.0)

    # Cross-sheet lookup value used by Sheet2!B10 bindings
    _leaf("Sheet2", "Z", 9, "D")


@dataclass(frozen=True)
class FormulaGroupFixture:
    """Formula-group graph plus the owning group node key.

    Args:
        graph: Dependency graph containing the group (and optional dependents).
        group_key: Multi-cell key of the formula-group node.
        members: Member cell addresses owned by the group.
        dependent: Optional non-member formula that references the group
            (e.g. ``Sheet1!B1`` when only the dependent is a codegen target).
    """

    graph: DependencyGraph
    group_key: str
    members: tuple[str, ...]
    dependent: str | None = None


def _wire_group_precedent_edges(graph: DependencyGraph, group_key: str) -> None:
    """Connect the group to every leaf currently in the graph (fixture precedents)."""
    for key in list(graph.keys()):
        node = graph.get_node(key)
        if node is None or not node.is_leaf:
            continue
        if key == group_key:
            continue
        graph.add_edge(group_key, key)


def build_row_stripe_group() -> FormulaGroupFixture:
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
    _wire_group_precedent_edges(g, group.key)
    for m in members:
        assert g.get_node(m) is None
        assert g.cell_owner(m) == group.key
    return FormulaGroupFixture(graph=g, group_key=group.key, members=members)


def build_cross_sheet_union_group() -> FormulaGroupFixture:
    """Non-contiguous cross-sheet union with the same INDEX/MATCH template."""
    members = ("Sheet1!D63", "Sheet2!B10")
    skeleton = index_match_skeleton()
    fp = shape_fingerprint(skeleton)
    bindings = {
        "Sheet1!D63": (CellRefNode(address="Sheet1!D35"),),
        "Sheet2!B10": (CellRefNode(address="Sheet2!Z9"),),  # value "D" -> col D -> 10.0
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
    _wire_group_precedent_edges(g, group.key)
    for m in members:
        assert g.get_node(m) is None
        assert g.cell_owner(m) == group.key
    return FormulaGroupFixture(graph=g, group_key=group.key, members=members)


def build_row_stripe_cell_only_twin() -> DependencyGraph:
    """Cell-only twin of `build_row_stripe_group` (no multi-cell node)."""
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
        # Mirror the group fixture: every shared leaf is a precedent (export closure).
        for key in list(g.keys()):
            leaf = g.get_node(key)
            if leaf is not None and leaf.is_leaf:
                g.add_edge(member, key)
    return g


def build_cross_sheet_cell_only_twin() -> DependencyGraph:
    """Cell-only twin of `build_cross_sheet_union_group`."""
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
        for key in list(g.keys()):
            leaf = g.get_node(key)
            if leaf is not None and leaf.is_leaf:
                g.add_edge(member, key)
    return g


def assert_formula_group_occupancy(group: Node) -> None:
    """Assert group owns its members and has a validated template."""
    assert group.skeleton is not None
    assert group.member_bindings is not None
    assert group.shape_fingerprint is not None
    owned = set(member_keys(group))
    assert owned == set(group.member_bindings)


def build_div_zero_group() -> FormulaGroupFixture:
    """Formula group whose specialized body is `1 / <cell>` (error-channel fixture)."""
    members = ("Sheet1!A1", "Sheet1!B1")
    skeleton = BinaryOpNode(
        op="/",
        left=NumberNode(value=1.0),
        right=AddressHoleNode(kind=AddressLeafKind.cell, slot=0),
    )
    fp = shape_fingerprint(skeleton)
    bindings = {
        "Sheet1!A1": (CellRefNode(address="Sheet1!Z1"),),
        "Sheet1!B1": (CellRefNode(address="Sheet1!Z2"),),
    }
    group = make_union_node(
        members,
        is_leaf=False,
        shape_fingerprint=fp,
        skeleton=skeleton,
        member_bindings=bindings,
    )
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(make_cell_node("Sheet1", "Z", 1, value=0.0, is_leaf=True))
    g.add_node(make_cell_node("Sheet1", "Z", 2, value=2.0, is_leaf=True))
    g.add_node(group)
    g.add_edge(group.key, "Sheet1!Z1")
    g.add_edge(group.key, "Sheet1!Z2")
    for m in members:
        assert g.get_node(m) is None
        assert g.cell_owner(m) == group.key
    return FormulaGroupFixture(graph=g, group_key=group.key, members=members)


def build_div_zero_cell_only_twin() -> DependencyGraph:
    """Cell-only twin of `build_div_zero_group`."""
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(make_cell_node("Sheet1", "Z", 1, value=0.0, is_leaf=True))
    g.add_node(make_cell_node("Sheet1", "Z", 2, value=2.0, is_leaf=True))
    for member, denom, formula in (
        ("Sheet1!A1", "Sheet1!Z1", "=1/Sheet1!Z1"),
        ("Sheet1!B1", "Sheet1!Z2", "=1/Sheet1!Z2"),
    ):
        sheet, cell = member.split("!")
        col = "".join(ch for ch in cell if ch.isalpha())
        row = int("".join(ch for ch in cell if ch.isdigit()))
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
        g.add_edge(member, denom)
    return g


def row_self_skeleton() -> AstNode:
    """Shared template: bare `ROW()` (member-context row number)."""
    return FunctionCallNode(name="ROW", args=[])


def column_self_skeleton() -> AstNode:
    """Shared template: bare `COLUMN()` (member-context column number)."""
    return FunctionCallNode(name="COLUMN", args=[])


def build_row_self_group() -> FormulaGroupFixture:
    """Column stripe `Sheet1!B10:B12` whose body is bare `ROW()`."""
    members = ("Sheet1!B10", "Sheet1!B11", "Sheet1!B12")
    skeleton = row_self_skeleton()
    bindings = {m: () for m in members}
    group = make_union_node(
        members,
        is_leaf=True,
        shape_fingerprint=shape_fingerprint(skeleton),
        skeleton=skeleton,
        member_bindings=bindings,
    )
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(group)
    for m in members:
        assert g.get_node(m) is None
        assert g.cell_owner(m) == group.key
    return FormulaGroupFixture(graph=g, group_key=group.key, members=members)


def build_column_self_group() -> FormulaGroupFixture:
    """Row stripe `Sheet1!D5:F5` whose body is bare `COLUMN()`."""
    members = ("Sheet1!D5", "Sheet1!E5", "Sheet1!F5")
    skeleton = column_self_skeleton()
    bindings = {m: () for m in members}
    group = make_union_node(
        members,
        is_leaf=True,
        shape_fingerprint=shape_fingerprint(skeleton),
        skeleton=skeleton,
        member_bindings=bindings,
    )
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(group)
    for m in members:
        assert g.get_node(m) is None
        assert g.cell_owner(m) == group.key
    return FormulaGroupFixture(graph=g, group_key=group.key, members=members)


def build_sum_over_constant_group() -> FormulaGroupFixture:
    """Minimal LIC-DSF repro: ``SUM`` over a group whose members are not targets.

    After coalesce, ``A1:A2`` are group members (no cell nodes). Only ``B1`` is a
    generate target. Exported ``xl_range`` must still resolve ``A1`` and ``A2`` via
    member wrappers (otherwise ``KeyError: Cell … not found in graph``).
    """
    members = ("Sheet1!A1", "Sheet1!A2")
    skeleton = NumberNode(value=10.0)
    bindings = {m: () for m in members}
    group = make_union_node(
        members,
        is_leaf=True,
        shape_fingerprint=shape_fingerprint(skeleton),
        skeleton=skeleton,
        member_bindings=bindings,
    )
    dependent = "Sheet1!B1"
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(group)
    g.add_node(
        make_cell_node(
            "Sheet1",
            "B",
            1,
            formula="=SUM(Sheet1!A1:A2)",
            normalized_formula="=SUM(Sheet1!A1:A2)",
            is_leaf=False,
        )
    )
    g.add_edge(dependent, group.key)
    for m in members:
        assert g.get_node(m) is None
        assert g.cell_owner(m) == group.key
    return FormulaGroupFixture(
        graph=g,
        group_key=group.key,
        members=members,
        dependent=dependent,
    )


def build_chart_data_threshold_sum_group() -> FormulaGroupFixture:
    """Chart Data ``D74 = SUM(D67:E67)`` over coalesced IF-threshold members.

    Mirrors ``lic_dsf_2025_08_12_formula_groups``::

        cell_chart_data_d74 → xl_sum(xl_range(..., D67:N67))

    where ``D67`` is a formula-group member (``=IF(D61>D66,1,0)``), not a generate
    target. Without member wrappers, export raises
    ``KeyError: Cell 'Chart Data'!D67 not found in graph``.
    """
    members = ("Sheet1!D67", "Sheet1!E67")
    skeleton = FunctionCallNode(
        name="IF",
        args=[
            BinaryOpNode(
                op=">",
                left=CellRefNode(address="Sheet1!D61"),
                right=CellRefNode(address="Sheet1!D66"),
            ),
            NumberNode(value=1.0),
            NumberNode(value=0.0),
        ],
    )
    bindings = {m: () for m in members}
    group = make_union_node(
        members,
        is_leaf=False,
        shape_fingerprint=shape_fingerprint(skeleton),
        skeleton=skeleton,
        member_bindings=bindings,
    )
    dependent = "Sheet1!D74"
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(make_cell_node("Sheet1", "D", 61, value=2.0, is_leaf=True))
    g.add_node(make_cell_node("Sheet1", "D", 66, value=1.0, is_leaf=True))
    g.add_node(group)
    g.add_edge(group.key, "Sheet1!D61")
    g.add_edge(group.key, "Sheet1!D66")
    g.add_node(
        make_cell_node(
            "Sheet1",
            "D",
            74,
            formula="=SUM(D67:E67)",
            normalized_formula="=SUM(Sheet1!D67:E67)",
            is_leaf=False,
        )
    )
    g.add_edge(dependent, group.key)
    for m in members:
        assert g.get_node(m) is None
        assert g.cell_owner(m) == group.key
    return FormulaGroupFixture(
        graph=g,
        group_key=group.key,
        members=members,
        dependent=dependent,
    )


def build_offset_address_hole_group() -> FormulaGroupFixture:
    """LIC-DSF OFFSET translation cells: ``OFFSET(hole, 0, 0)`` in a shared helper.

    Without hole-aware OFFSET emission, group helpers return ``xl_raise(#REF!)``.
    """
    members = ("Sheet1!A12", "Sheet1!A13")
    skeleton = FunctionCallNode(
        name="OFFSET",
        args=[
            AddressHoleNode(kind=AddressLeafKind.cell, slot=0),
            NumberNode(value=0.0),
            NumberNode(value=0.0),
        ],
    )
    bindings = {
        "Sheet1!A12": (CellRefNode(address="Sheet1!Z10"),),
        "Sheet1!A13": (CellRefNode(address="Sheet1!Z11"),),
    }
    group = make_union_node(
        members,
        is_leaf=False,
        shape_fingerprint=shape_fingerprint(skeleton),
        skeleton=skeleton,
        member_bindings=bindings,
    )
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(make_cell_node("Sheet1", "Z", 10, value="alpha", is_leaf=True))
    g.add_node(make_cell_node("Sheet1", "Z", 11, value="beta", is_leaf=True))
    g.add_node(group)
    g.add_edge(group.key, "Sheet1!Z10")
    g.add_edge(group.key, "Sheet1!Z11")
    for m in members:
        assert g.get_node(m) is None
        assert g.cell_owner(m) == group.key
    return FormulaGroupFixture(graph=g, group_key=group.key, members=members)
