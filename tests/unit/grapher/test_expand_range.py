"""`expand_range` fails closed when a rectangle exceeds `max_cells`."""

from __future__ import annotations

import pytest

from excel_grapher.grapher.parser import expand_range
from excel_grapher.grapher.target_expansion import expand_targets_to_roots
from excel_grapher.series_bindings import expand_data_range


def test_expand_range_at_max_cells_enumerates_every_cell() -> None:
    got = expand_range(
        sheet="Sheet1",
        start_col="A",
        start_row=1,
        end_col="A",
        end_row=5000,
        max_cells=5000,
    )
    assert len(got) == 5000
    assert got[0] == ("Sheet1", "A1")
    assert got[-1] == ("Sheet1", "A5000")


def test_expand_range_over_max_cells_raises() -> None:
    with pytest.raises(ValueError, match=r"5001 cells, exceeding max_cells=5000") as exc_info:
        expand_range(
            sheet="Sheet1",
            start_col="A",
            start_row=1,
            end_col="A",
            end_row=5001,
            max_cells=5000,
        )
    assert "Sheet1!A1:A5001" in str(exc_info.value)


def test_expand_range_quoted_sheet_in_error() -> None:
    with pytest.raises(ValueError, match=r"'My Sheet'!A1:C2") as exc_info:
        expand_range(
            sheet="My Sheet",
            start_col="A",
            start_row=1,
            end_col="C",
            end_row=2,
            max_cells=4,
        )
    assert "6 cells" in str(exc_info.value)


def test_expand_targets_to_roots_over_max_range_cells_raises() -> None:
    with pytest.raises(ValueError, match="exceeding max_cells=2"):
        expand_targets_to_roots(
            ["Sheet1!A1:A3"],
            sheetnames=["Sheet1"],
            named_ranges={},
            named_range_ranges={},
            max_range_cells=2,
        )


def test_expand_data_range_over_default_cap_raises() -> None:
    with pytest.raises(ValueError, match="exceeding max_cells=5000"):
        expand_data_range("Sheet1!A1:A5001")
