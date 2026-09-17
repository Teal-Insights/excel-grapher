from __future__ import annotations

import pytest

from excel_grapher.grapher.blank_ranges import (
    cell_in_blank_ranges,
    normalize_blank_range_specs,
    parse_blank_range_spec,
)


def test_parse_blank_range_spec_single_cell() -> None:
    assert parse_blank_range_spec("Sheet1!B2") == ("Sheet1", 2, 2, 2, 2)


def test_parse_blank_range_spec_rectangle() -> None:
    assert parse_blank_range_spec("Sheet1!B2:D4") == ("Sheet1", 2, 2, 4, 4)


def test_parse_blank_range_spec_quoted_sheet() -> None:
    assert parse_blank_range_spec("'My Sheet'!A1:C2") == ("My Sheet", 1, 1, 2, 3)


def test_normalize_blank_range_specs_rejects_str() -> None:
    with pytest.raises(TypeError):
        normalize_blank_range_specs("Sheet1!A1")


def test_cell_in_blank_ranges() -> None:
    rects = normalize_blank_range_specs(["S!A2:B3"])
    assert cell_in_blank_ranges("S", 2, 1, rects)
    assert cell_in_blank_ranges("S", 3, 2, rects)
    assert not cell_in_blank_ranges("S", 1, 1, rects)
    assert not cell_in_blank_ranges("T", 2, 1, rects)
