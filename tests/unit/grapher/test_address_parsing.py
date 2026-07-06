"""Tests for consolidated sheet-qualified address parsing."""

from __future__ import annotations

from excel_grapher.core.address_keys import format_key, parse_address
from excel_grapher.grapher.blank_ranges import address_in_blank_ranges, parse_blank_range_spec
from excel_grapher.grapher.builder import _workbook_sorted_sheet_a1_pairs


def test_format_key_round_trip_handles_apostrophe_sheet_names() -> None:
    sheet = "O'Neil"
    a1 = "B2"
    key = format_key(sheet, a1)
    assert key == "'O''Neil'!B2"
    assert parse_address(key) == (sheet, a1)


def test_workbook_sorted_sheet_a1_pairs_handles_apostrophe_sheet_names() -> None:
    pairs = [("O'Neil", "A1"), ("Main", "B2")]
    assert _workbook_sorted_sheet_a1_pairs(pairs, sheet_order=["Main", "O'Neil"]) == [
        ("Main", "B2"),
        ("O'Neil", "A1"),
    ]


def test_parse_blank_range_spec_handles_apostrophe_sheet_names() -> None:
    assert parse_blank_range_spec("'It''s Data'!A1:C2") == ("It's Data", 1, 1, 2, 3)


def test_address_in_blank_ranges_handles_apostrophe_sheet_names() -> None:
    rects = (parse_blank_range_spec("'It''s Data'!A1:B2"),)
    assert address_in_blank_ranges("'It''s Data'!A1", rects)
    assert not address_in_blank_ranges("'It''s Data'!C3", rects)
