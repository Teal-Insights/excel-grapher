from __future__ import annotations

import pytest

from excel_grapher.core.address_keys import (
    make_node_key_sort_key,
    normalize_range_key,
    sort_node_keys,
)


def test_sort_node_keys_respects_workbook_sheet_order_then_row_then_column() -> None:
    sheet_order = ["Inputs", "Calc", "Summary"]
    keys = [
        "Calc!B2",
        "Inputs!C1",
        "Inputs!A1",
        "Calc!A2",
        "Summary!A1",
        "Calc!A1",
    ]

    assert sort_node_keys(keys, sheet_order=sheet_order) == [
        "Inputs!A1",
        "Inputs!C1",
        "Calc!A1",
        "Calc!A2",
        "Calc!B2",
        "Summary!A1",
    ]


def test_sort_node_keys_handles_quoted_sheet_names_in_workbook_order() -> None:
    sheet_order = ["Main", "Data Set", "O'Neil"]
    keys = [
        "'Data Set'!A1",
        "'O''Neil'!A1",
        "Main!A1",
    ]

    assert sort_node_keys(keys, sheet_order=sheet_order) == [
        "Main!A1",
        "'Data Set'!A1",
        "'O''Neil'!A1",
    ]


def test_make_node_key_sort_key_places_unknown_sheets_after_known() -> None:
    key_fn = make_node_key_sort_key(sheet_order=["Known"])
    keys = ["Known!A1", "Other!A1", "Another!A1"]

    assert sorted(keys, key=key_fn) == ["Known!A1", "Another!A1", "Other!A1"]


def test_normalize_range_key_local_range_uses_current_sheet() -> None:
    assert normalize_range_key("$A$1:B3", current_sheet="Sheet1") == "Sheet1!A1:Sheet1!B3"


def test_normalize_range_key_accepts_sheet_once_form() -> None:
    assert normalize_range_key("Sheet1!A1:B3") == "Sheet1!A1:Sheet1!B3"


def test_normalize_range_key_preserves_quoted_sheet_names() -> None:
    normalized = normalize_range_key("'My Sheet'!$A$1:'My Sheet'!B3")
    assert normalized == "'My Sheet'!A1:'My Sheet'!B3"


def test_normalize_range_key_rejects_cross_sheet_ranges() -> None:
    with pytest.raises(ValueError):
        normalize_range_key("Sheet1!A1:Sheet2!B3")
