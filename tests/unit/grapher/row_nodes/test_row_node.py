"""Unit tests for row Node / NodeView construction (issue #374 sprint 2)."""

from __future__ import annotations

import pytest

from excel_grapher.grapher.node import (
    Node,
    NodeKind,
    make_row_node,
    node_to_view,
    row_member_keys,
)


def test_cell_node_backcompat_defaults_kind_and_bounds() -> None:
    node = Node("Sheet1", "D", 63, None, None, 1, True)
    assert node.kind is NodeKind.cell
    assert node.column == "D"
    assert node.row == 63
    assert node.min_col == "D"
    assert node.max_col == "D"
    assert node.min_row == 63
    assert node.max_row == 63
    assert node.key == "Sheet1!D63"
    assert node.address == "D63"
    assert node.column_index == 4


def test_make_row_node_sets_kind_and_extent() -> None:
    node = make_row_node("Sheet1", 63, "D", "Y")
    assert node.kind is NodeKind.row
    assert node.sheet == "Sheet1"
    assert node.row == 63
    assert node.column is None
    assert node.min_col == "D"
    assert node.max_col == "Y"
    assert node.min_row == 63
    assert node.max_row == 63
    assert node.key == "Sheet1!D63:Y63"
    assert node.address == "D63:Y63"
    assert node.is_leaf is True
    assert node.formula is None


def test_make_row_node_orders_inverted_columns() -> None:
    node = make_row_node("Sheet1", 63, "Y", "D")
    assert node.min_col == "D"
    assert node.max_col == "Y"
    assert node.key == "Sheet1!D63:Y63"


def test_make_row_node_quoted_sheet() -> None:
    node = make_row_node("My Sheet", 1, "A", "Z")
    assert node.key == "'My Sheet'!A1:Z1"


def test_row_node_rejects_multi_row_extent() -> None:
    with pytest.raises(ValueError, match="one-row|same row|min_row"):
        Node(
            sheet="Sheet1",
            column=None,
            row=40,
            formula=None,
            normalized_formula=None,
            value=None,
            is_leaf=True,
            kind=NodeKind.row,
            min_col="D",
            min_row=40,
            max_col="AJ",
            max_row=50,
        )


def test_row_members_differ_only_by_column() -> None:
    from fastpyxl.utils.cell import coordinate_from_string

    from excel_grapher.core.address_keys import parse_address

    node = make_row_node("Sheet1", 63, "D", "F")
    keys = row_member_keys(node)
    assert keys == ["Sheet1!D63", "Sheet1!E63", "Sheet1!F63"]

    parsed = [parse_address(k) for k in keys]
    sheets = {sheet for sheet, _ in parsed}
    coords = [coordinate_from_string(coord) for _, coord in parsed]
    rows = {int(row) for _, row in coords}
    cols = [col for col, _ in coords]
    assert sheets == {"Sheet1"}
    assert rows == {63}
    assert cols == ["D", "E", "F"]


def test_row_member_keys_rejects_cell_node() -> None:
    cell = Node("Sheet1", "D", 63, None, None, 1, True)
    with pytest.raises(ValueError, match="row"):
        row_member_keys(cell)


def test_node_to_view_preserves_row_fields() -> None:
    node = make_row_node(
        "Sheet1",
        63,
        "D",
        "Y",
        formula="=D$35",
        normalized_formula="=D$35",
        varying_ref_slots=(0,),
        value=None,
        is_leaf=False,
        metadata={"tag": "row"},
    )
    view = node_to_view(node)
    assert view.kind is NodeKind.row
    assert view.key == "Sheet1!D63:Y63"
    assert view.min_col == "D"
    assert view.max_col == "Y"
    assert view.min_row == 63
    assert view.max_row == 63
    assert view.column is None
    assert view.row == 63
    assert view.varying_ref_slots == (0,)
    assert view.metadata["tag"] == "row"
    assert view.address == "D63:Y63"


def test_one_by_one_row_node_distinct_from_cell() -> None:
    row = make_row_node("Sheet1", 63, "D", "D")
    cell = Node("Sheet1", "D", 63, None, None, None, True)
    assert row.kind is NodeKind.row
    assert cell.kind is NodeKind.cell
    assert row.key == "Sheet1!D63:D63"
    assert cell.key == "Sheet1!D63"
    assert row.key != cell.key
