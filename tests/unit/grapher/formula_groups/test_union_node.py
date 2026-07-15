"""Sprint 2 unit tests for address-centric Node and union factories."""

from __future__ import annotations

import pytest

from excel_grapher.core.address_keys import CellKey, NodeShape, RangeKey, UnionKey, parse_node_key
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import (
    Node,
    NodeKind,
    locate_cell,
    make_cell_node,
    make_union_node,
    member_keys,
    node_to_view,
)


def test_cell_node_legacy_ctor_sets_canonical_address() -> None:
    node = Node("Sheet1", "D", 63, None, None, 1, True)
    assert isinstance(node.address, CellKey)
    assert node.address == "Sheet1!D63"
    assert node.key == "Sheet1!D63"
    assert node.shape is NodeShape.cell
    assert node.kind is NodeKind.cell
    assert node.sheet == "Sheet1"
    assert node.column == "D"
    assert node.row == 63


def test_make_cell_node() -> None:
    node = make_cell_node("Sheet1", "E", 63, value=1.0, is_leaf=True)
    assert node.address == "Sheet1!E63"
    assert node.value == 1.0
    assert node.shape is NodeShape.cell


def test_make_union_node_from_scattered_members() -> None:
    node = make_union_node(
        ["Sheet1!E5", "Sheet1!A1", "Sheet1!C1", "Sheet1!B1", "Sheet1!D1"],
        is_leaf=False,
    )
    assert isinstance(node.address, UnionKey)
    assert node.address == "Sheet1!A1:D1,E5"
    assert node.key == "Sheet1!A1:D1,E5"
    assert node.shape is NodeShape.union
    assert node.kind is NodeKind.union
    assert node.value is None
    assert node.formula is None
    assert node.normalized_formula is None
    assert set(member_keys(node)) == {
        "Sheet1!A1",
        "Sheet1!B1",
        "Sheet1!C1",
        "Sheet1!D1",
        "Sheet1!E5",
    }


def test_make_union_node_filled_block_is_range() -> None:
    from fastpyxl.utils.cell import get_column_letter

    cells = [
        f"Sheet1!{get_column_letter(c)}{r}"
        for r in range(4, 19)
        for c in range(5, 10)  # E..I
    ]
    node = make_union_node(cells)
    assert isinstance(node.address, RangeKey)
    assert node.address == "Sheet1!E4:I18"
    assert node.shape is NodeShape.range
    assert node.kind is NodeKind.union
    assert node.value is None


def test_make_union_node_cross_sheet() -> None:
    node = make_union_node(["Sheet2!B2", "Sheet1!A1"])
    assert isinstance(node.address, UnionKey)
    assert node.address == "Sheet1!A1,Sheet2!B2"
    assert node.sheet is None  # multi-sheet
    assert member_keys(node) == ["Sheet1!A1", "Sheet2!B2"]


def test_make_union_node_empty_rejected() -> None:
    with pytest.raises(ValueError, match="empty"):
        make_union_node([])


def test_make_union_node_single_cell_prefers_cell_node() -> None:
    node = make_union_node(["Sheet1!E63"], value=None)
    assert isinstance(node.address, CellKey)
    assert node.kind is NodeKind.cell
    assert node.address == "Sheet1!E63"


def test_make_union_node_rejects_formula_for_multi_cell() -> None:
    with pytest.raises(ValueError, match="Multi-cell nodes cannot have formula"):
        make_union_node(["Sheet1!D63", "Sheet1!Y63"], formula="=SUM(D63:Y63)")


def test_node_from_one_row_address_kwarg_uses_union_kind() -> None:
    node = Node(
        sheet=None,
        column=None,
        row=None,
        formula=None,
        normalized_formula=None,
        value=None,
        is_leaf=True,
        address=parse_node_key("Sheet1!D63:Y63"),
    )
    assert isinstance(node.address, RangeKey)
    assert node.address == "Sheet1!D63:Y63"
    assert node.shape is NodeShape.row
    assert node.kind is NodeKind.union
    assert node.column is None
    assert node.row == 63
    assert node.min_col == "D"
    assert node.max_col == "Y"


def test_node_from_address_kwarg() -> None:
    node = Node(
        sheet=None,
        column=None,
        row=None,
        formula=None,
        normalized_formula=None,
        value=None,
        is_leaf=True,
        address=UnionKey("Sheet1!A1:D1,E5"),
    )
    assert isinstance(node.address, UnionKey)
    assert node.key == "Sheet1!A1:D1,E5"
    assert node.kind is NodeKind.union


def test_node_to_view_preserves_union_address() -> None:
    node = make_union_node(["Sheet1!A1", "Sheet1!E5"], metadata={"tag": "u"})
    view = node_to_view(node)
    assert view.address == node.address
    assert view.key == node.key
    assert view.kind is NodeKind.union
    assert view.metadata["tag"] == "u"
    assert view.shape is NodeShape.union


def test_locate_cell_finds_union_owner() -> None:
    g = DependencyGraph()
    g.add_node(make_union_node(["Sheet1!A1", "Sheet1!B1", "Sheet1!E5"]))
    loc = locate_cell(g, "Sheet1!E5")
    assert loc is not None
    assert loc.cell_key == "Sheet1!E5"
    assert loc.kind is NodeKind.union
    assert loc.node_key == "Sheet1!A1:B1,E5"
    assert loc.column == "E"


def test_locate_cell_prefers_exact_cell_node() -> None:
    g = DependencyGraph()
    g.add_node(make_cell_node("Sheet1", "E", 5, value=1))
    loc = locate_cell(g, "Sheet1!E5")
    assert loc is not None
    assert loc.kind is NodeKind.cell
    assert loc.node_key == "Sheet1!E5"
