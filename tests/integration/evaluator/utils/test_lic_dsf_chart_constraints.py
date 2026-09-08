"""Unit tests for LIC-DSF chart-shortlist DynamicRefConfig ranges (no workbook I/O)."""

from __future__ import annotations

from typing import get_args

import pytest

from excel_grapher import DynamicRefError
from tests.integration.evaluator.utils.lic_dsf_chart_constraints import (
    CHART_SHORTLIST_CONSTRAINT_RANGES,
    _literal_for_value,
    expand_chart_shortlist_constraint_keys,
)
from tests.integration.evaluator.utils.lic_dsf_chart_targets import parse_range_spec


def test_chart_shortlist_constraint_ranges_are_sheet_qualified() -> None:
    for spec in CHART_SHORTLIST_CONSTRAINT_RANGES:
        sheet, a1 = parse_range_spec(spec)
        assert sheet
        assert a1


def test_chart_shortlist_constraint_keys_cover_discovered_leaves() -> None:
    keys = expand_chart_shortlist_constraint_keys()
    assert len(keys) == 1654
    assert len(set(keys)) == 1654
    expected = {
        "'Input 1 - Basics'!C7",
        "START!K10",
        "lookup!C4",
        "lookup!G73",
        "lookup!BB4",
        "Trigger!AB5",
        "Classification!A16",
        "Country_Information!A1",
        "data_GDPUSD!A1",
        "data_REM_FINAL!AA7",
        "'BLEND floating calculations WB'!K10",
    }
    assert expected <= set(keys)


def test_literal_for_value_freezes_template_leaves() -> None:
    assert get_args(_literal_for_value(None)) == (None,)
    assert get_args(_literal_for_value("")) == (None,)
    assert get_args(_literal_for_value("Ghana")) == ("Ghana",)
    assert get_args(_literal_for_value(652)) == (652,)
    with pytest.raises(DynamicRefError, match="formula"):
        _literal_for_value("=A1")
