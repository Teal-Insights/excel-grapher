"""Unit tests for LIC-DSF chart target helpers (no workbook required)."""

from __future__ import annotations

import pytest

from tests.evaluator.lic_dsf_chart_targets import (
    cells_in_range,
    chart_parity_shortlist_keys,
    parse_range_spec,
)


def test_parse_range_spec_quoted_sheet_with_space() -> None:
    sheet, a1 = parse_range_spec("'Chart Data'!D10:D17")
    assert sheet == "Chart Data"
    assert a1 == "D10:D17"


def test_parse_range_spec_unquoted_sheet() -> None:
    sheet, a1 = parse_range_spec("Sheet1!A1")
    assert sheet == "Sheet1"
    assert a1 == "A1"


def test_parse_range_spec_rejects_missing_bang() -> None:
    with pytest.raises(ValueError, match="!"):
        parse_range_spec("A1")


def test_cells_in_range_single_cell() -> None:
    assert cells_in_range("Chart Data", "B5") == ["'Chart Data'!B5"]


def test_cells_in_range_small_rectangle() -> None:
    keys = cells_in_range("S", "A1:B2")
    assert set(keys) == {"S!A1", "S!A2", "S!B1", "S!B2"}


def test_chart_parity_shortlist_keys_are_normalized() -> None:
    keys = chart_parity_shortlist_keys()
    assert len(keys) == 2
    assert all(k.startswith("'Chart Data'!") for k in keys)
    assert keys[0].endswith("U63")
    assert keys[1].endswith("U66")
