"""Sprint 1 unit tests for CellKey / RangeKey / UnionKey and cover algorithm."""

from __future__ import annotations

import pytest

from excel_grapher.core.address_keys import (
    CellKey,
    NodeShape,
    RangeKey,
    UnionKey,
    members_to_node_key,
    parse_node_key,
)


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


def test_members_to_node_key_order_independent_union() -> None:
    members = [
        "Sheet1!E5",
        "Sheet1!A1",
        "Sheet1!C1",
        "Sheet1!B1",
        "Sheet1!D1",
    ]
    shuffled = list(reversed(members))
    key_a = members_to_node_key(members)
    key_b = members_to_node_key(shuffled)
    assert key_a == key_b == "Sheet1!A1:D1,E5"
    assert isinstance(key_a, UnionKey)
    assert key_a.shape is NodeShape.union


def test_members_to_node_key_filled_block_single_range() -> None:
    cells: list[str] = []
    for row in range(4, 19):
        for col in ("E", "F", "G", "H", "I"):
            cells.append(f"Sheet1!{col}{row}")
    key = members_to_node_key(cells)
    assert isinstance(key, RangeKey)
    assert key == "Sheet1!E4:I18"
    assert key.shape is NodeShape.range


def test_members_to_node_key_single_cell() -> None:
    key = members_to_node_key(["Sheet1!E63"])
    assert isinstance(key, CellKey)
    assert key == "Sheet1!E63"


def test_members_to_node_key_row_stripe() -> None:
    # non-contiguous columns on one row -> horizontal runs, not one D:Y span
    key = members_to_node_key(["Sheet1!D63", "Sheet1!E63", "Sheet1!F63", "Sheet1!Y63"])
    assert key == "Sheet1!D63:F63,Y63"
    # contiguous D..Y
    from fastpyxl.utils.cell import get_column_letter

    start = 4  # D
    end = 25  # Y
    contig = [f"Sheet1!{get_column_letter(i)}63" for i in range(start, end + 1)]
    key2 = members_to_node_key(contig)
    assert isinstance(key2, RangeKey)
    assert key2 == "Sheet1!D63:Y63"
    assert key2.shape is NodeShape.row


def test_members_to_node_key_dedupes_and_rejects_empty() -> None:
    key = members_to_node_key(["Sheet1!A1", "Sheet1!A1", "Sheet1!B1"])
    assert key == "Sheet1!A1:B1"
    with pytest.raises(ValueError, match="empty"):
        members_to_node_key([])


def test_members_to_node_key_cross_sheet() -> None:
    key = members_to_node_key(["Sheet2!B2", "Sheet1!A1"])
    assert isinstance(key, UnionKey)
    assert key == "Sheet1!A1,Sheet2!B2"


def test_members_to_node_key_quoted_sheet() -> None:
    key = members_to_node_key(["'My Sheet'!A1", "'My Sheet'!B1"])
    assert key == "'My Sheet'!A1:B1"


def test_members_to_node_key_rejects_non_cell_members() -> None:
    with pytest.raises(ValueError, match="cell"):
        members_to_node_key(["Sheet1!A1:B2"])
