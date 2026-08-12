"""Unit tests for CellKey / RangeKey / UnionKey address parsing."""

from __future__ import annotations

import pytest

from excel_grapher.core.address_keys import (
    CellKey,
    NodeShape,
    RangeKey,
    UnionKey,
    parse_node_key,
)


@pytest.mark.parametrize(
    ("key_cls", "text"),
    [
        (CellKey, "Sheet1!E63"),
        (RangeKey, "Sheet1!D63:Y63"),
        (UnionKey, "Sheet1!A1:D1,E5"),
    ],
)
def test_address_keys_are_slotted(key_cls: type[str], text: str) -> None:
    key = key_cls(text)
    assert not hasattr(key, "__dict__")
    with pytest.raises(AttributeError):
        key.dynamic_attr = 1  # type: ignore[attr-defined]
    assert key == text
    assert hash(key) == hash(text)
    assert {key: "ok"}[text] == "ok"


def test_parse_cell_key() -> None:
    key = parse_node_key("Sheet1!E63")
    assert isinstance(key, CellKey)
    assert key == "Sheet1!E63"
    assert key.shape is NodeShape.cell
    assert key.sheet == "Sheet1"
    assert key.column == "E"
    assert key.row == 63


def test_parse_cell_key_strips_dollars_and_unnecessary_quotes() -> None:
    key = parse_node_key("'Sheet1'!$E$63")
    assert isinstance(key, CellKey)
    assert key == "Sheet1!E63"


def test_parse_row_range_orders_columns() -> None:
    key = parse_node_key("Sheet1!Y63:D63")
    assert isinstance(key, RangeKey)
    assert key == "Sheet1!D63:Y63"
    assert key.shape is NodeShape.row
    assert key.sheet == "Sheet1"
    assert key.min_col == "D"
    assert key.max_col == "Y"
    assert key.min_row == 63
    assert key.max_row == 63
    assert key.row == 63
    assert key.column is None


def test_parse_filled_block_is_range_shape() -> None:
    key = parse_node_key("Sheet1!E4:I18")
    assert isinstance(key, RangeKey)
    assert key == "Sheet1!E4:I18"
    assert key.shape is NodeShape.range
    assert key.min_col == "E"
    assert key.max_col == "I"
    assert key.min_row == 4
    assert key.max_row == 18
    assert key.row is None
    assert key.column is None


def test_parse_single_column_range_shape() -> None:
    key = parse_node_key("Sheet1!E4:E18")
    assert isinstance(key, RangeKey)
    assert key.shape is NodeShape.column
    assert key.column == "E"
    assert key.row is None


def test_parse_one_by_one_range_collapses_to_cell() -> None:
    key = parse_node_key("Sheet1!D63:D63")
    assert isinstance(key, CellKey)
    assert key == "Sheet1!D63"


def test_parse_both_ends_sheet_qualified_range() -> None:
    key = parse_node_key("Sheet1!D63:Sheet1!Y63")
    assert isinstance(key, RangeKey)
    assert key == "Sheet1!D63:Y63"


def test_parse_cross_sheet_range_rejected() -> None:
    with pytest.raises(ValueError, match="sheet"):
        parse_node_key("Sheet1!D63:Sheet2!Y63")


def test_quoted_sheet_range_round_trip() -> None:
    key = parse_node_key("'My Sheet'!A1:Z1")
    assert isinstance(key, RangeKey)
    assert key == "'My Sheet'!A1:Z1"
    assert key.sheet == "My Sheet"
    assert key.shape is NodeShape.row


def test_parse_union_same_sheet() -> None:
    key = parse_node_key("Sheet1!E4:I18,H23:K36,D9:H19")
    assert isinstance(key, UnionKey)
    # Sorted by (min_row, min_col, max_row, max_col)
    assert key == "Sheet1!E4:I18,D9:H19,H23:K36"
    assert key.shape is NodeShape.union
    assert len(key.members) == 3
    assert all(isinstance(m, RangeKey) for m in key.members)
    assert key.members[0] == "Sheet1!E4:I18"


def test_parse_union_collapses_duplicate_and_one_member() -> None:
    key = parse_node_key("Sheet1!A1,Sheet1!A1")
    assert isinstance(key, CellKey)
    assert key == "Sheet1!A1"


def test_parse_empty_union_rejected() -> None:
    with pytest.raises(ValueError, match="empty|union"):
        parse_node_key("Sheet1!")


def test_parse_cross_sheet_union() -> None:
    key = parse_node_key("Sheet1!A1,Sheet2!B2")
    assert isinstance(key, UnionKey)
    assert key == "Sheet1!A1,Sheet2!B2"
    assert key.members[0] == "Sheet1!A1"
    assert key.members[1] == "Sheet2!B2"


def test_parse_union_strips_spaces_around_commas() -> None:
    key = parse_node_key("Sheet1!A1:D1, E5")
    assert key == "Sheet1!A1:D1,E5"


def test_parse_node_key_idempotent() -> None:
    key = parse_node_key("Sheet1!Y63:D63")
    assert parse_node_key(key) is key or parse_node_key(key) == key
