"""Unit tests for one-row node key helpers (issue #374 sprint 1)."""

from __future__ import annotations

import pytest

from excel_grapher.core.address_keys import (
    format_row_key,
    normalize_row_key,
    parse_row_key,
)


def test_format_parse_row_key_round_trip() -> None:
    key = format_row_key("Sheet1", "D", 63, "Y")
    assert key == "Sheet1!D63:Y63"

    parsed = parse_row_key(key)
    assert parsed.sheet == "Sheet1"
    assert parsed.row == 63
    assert parsed.min_col == "D"
    assert parsed.max_col == "Y"
    assert normalize_row_key(key) == key


def test_format_row_key_orders_inverted_columns() -> None:
    assert format_row_key("Sheet1", "Y", 63, "D") == "Sheet1!D63:Y63"


def test_normalize_row_key_orders_inverted_columns() -> None:
    assert normalize_row_key("Sheet1!Y63:D63") == "Sheet1!D63:Y63"


def test_row_key_rejects_multi_row() -> None:
    with pytest.raises(ValueError, match="one-row|single row|same row"):
        parse_row_key("Sheet1!D40:AJ50")
    with pytest.raises(ValueError, match="one-row|single row|same row"):
        normalize_row_key("Sheet1!D40:AJ50")


def test_quoted_sheet_row_key_round_trip() -> None:
    key = format_row_key("My Sheet", "A", 1, "Z")
    assert key == "'My Sheet'!A1:Z1"

    parsed = parse_row_key(key)
    assert parsed.sheet == "My Sheet"
    assert parsed.row == 1
    assert parsed.min_col == "A"
    assert parsed.max_col == "Z"
    assert normalize_row_key(key) == key


def test_normalize_row_key_strips_unnecessary_sheet_quotes() -> None:
    assert normalize_row_key("'Sheet1'!D63:Y63") == "Sheet1!D63:Y63"


def test_parse_row_key_accepts_both_ends_sheet_qualified() -> None:
    parsed = parse_row_key("Sheet1!D63:Sheet1!Y63")
    assert parsed.sheet == "Sheet1"
    assert parsed.min_col == "D"
    assert parsed.max_col == "Y"
    assert parsed.row == 63
    assert normalize_row_key("Sheet1!D63:Sheet1!Y63") == "Sheet1!D63:Y63"


def test_parse_row_key_rejects_cross_sheet_range() -> None:
    with pytest.raises(ValueError, match="sheet"):
        parse_row_key("Sheet1!D63:Sheet2!Y63")


def test_parse_row_key_rejects_cell_only_key() -> None:
    with pytest.raises(ValueError, match="row|range|:"):
        parse_row_key("Sheet1!D63")


def test_parse_row_key_allows_one_by_one_span() -> None:
    parsed = parse_row_key("Sheet1!D63:D63")
    assert parsed.min_col == "D"
    assert parsed.max_col == "D"
    assert parsed.row == 63
    assert normalize_row_key("Sheet1!D63:D63") == "Sheet1!D63:D63"


def test_parse_row_key_strips_absolute_markers() -> None:
    assert normalize_row_key("Sheet1!$D$63:$Y$63") == "Sheet1!D63:Y63"


def test_quoted_sheet_with_apostrophe_row_key() -> None:
    key = format_row_key("O'Neil", "A", 1, "C")
    assert key == "'O''Neil'!A1:C1"
    parsed = parse_row_key(key)
    assert parsed.sheet == "O'Neil"
    assert normalize_row_key(key) == key
