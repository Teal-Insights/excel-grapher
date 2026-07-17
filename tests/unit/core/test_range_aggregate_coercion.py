"""Excel range-argument filtering for SUM/AVERAGE/MIN/MAX/STDEV (#419).

Range/array args keep only numbers (skip blanks, text, booleans); errors
propagate. Literal scalar args still coerce (`SUM(1, "2", TRUE)` is 4).
"""

from __future__ import annotations

import math
from typing import cast

import pytest

from excel_grapher.core import CellValue, XlError
from excel_grapher.core.grid import Range
from excel_grapher.core.math_funcs import (
    average_cells,
    max_cells,
    min_cells,
    stdev_cells,
    sum_cells,
)


def _mcve_range() -> Range:
    """Synthetic range matching issue #419: numbers, \"\", blank, text, bool."""
    values: dict[str, CellValue] = {
        "S!A1": -0.2,
        "S!A2": -0.3,
        "S!A3": "",  # empty text from a guard formula
        "S!A4": None,  # truly blank
        "S!A5": "n.a.",
        "S!A6": True,
    }
    return Range("S", 1, 1, 6, 1, lambda a: values[a])


def test_sum_range_ignores_blanks_text_and_booleans() -> None:
    assert sum_cells(cast(CellValue, _mcve_range())) == -0.5


def test_average_range_ignores_blanks_text_and_booleans() -> None:
    assert average_cells(cast(CellValue, _mcve_range())) == -0.25


def test_min_range_ignores_blanks_text_and_booleans() -> None:
    assert min_cells(cast(CellValue, _mcve_range())) == -0.3


def test_max_range_ignores_blanks_text_and_booleans() -> None:
    assert max_cells(cast(CellValue, _mcve_range())) == -0.2


def test_stdev_range_ignores_blanks_text_and_booleans() -> None:
    result = stdev_cells(cast(CellValue, _mcve_range()))
    assert isinstance(result, float)
    assert result == pytest.approx(math.sqrt(0.005))


def test_average_range_with_only_blanks_and_empty_text_does_not_halve() -> None:
    """Silent failure shape: empties must not count as zeros in the divisor."""
    values: dict[str, CellValue] = {
        "S!A1": -0.2,
        "S!A2": -0.3,
        "S!A3": "",
        "S!A4": None,
    }
    rng = Range("S", 1, 1, 4, 1, lambda a: values[a])
    assert average_cells(cast(CellValue, rng)) == -0.25


def test_nested_list_range_ignores_non_numerics() -> None:
    assert sum_cells([[-0.2], [-0.3], [""], [None], ["n.a."], [True]]) == -0.5
    assert average_cells([[-0.2], [-0.3], [""], [None]]) == -0.25


def test_literal_scalars_still_coerce() -> None:
    assert sum_cells(1, "2", True) == 4.0
    assert average_cells(1, "3", False) == 4.0 / 3.0
    assert min_cells(1, "2", True) == 1.0
    assert max_cells(1, "2", True) == 2.0


def test_mixed_literal_and_range_args() -> None:
    rng = Range("S", 1, 1, 2, 1, lambda a: {"S!A1": 10.0, "S!A2": "skip"}[a])
    assert sum_cells(1, cast(CellValue, rng), True) == 12.0


def test_average_of_no_numbers_is_div_error() -> None:
    rng = Range("S", 1, 1, 2, 1, lambda a: {"S!A1": "", "S!A2": None}[a])
    assert average_cells(cast(CellValue, rng)) == XlError.DIV
    assert average_cells() == XlError.DIV


def test_stdev_of_fewer_than_two_numbers_is_div_error() -> None:
    rng = Range("S", 1, 1, 2, 1, lambda a: {"S!A1": 1.0, "S!A2": ""}[a])
    assert stdev_cells(cast(CellValue, rng)) == XlError.DIV


def test_sum_min_max_of_no_numbers_is_zero() -> None:
    rng = Range("S", 1, 1, 2, 1, lambda a: {"S!A1": "x", "S!A2": True}[a])
    assert sum_cells(cast(CellValue, rng)) == 0.0
    assert min_cells(cast(CellValue, rng)) == 0.0
    assert max_cells(cast(CellValue, rng)) == 0.0


def test_range_error_propagates() -> None:
    calls: list[str] = []

    def resolve(address: str) -> CellValue:
        calls.append(address)
        return {"S!A1": 1.0, "S!A2": XlError.DIV, "S!A3": 99.0}[address]

    rng = Range("S", 1, 1, 3, 1, resolve)
    assert sum_cells(cast(CellValue, rng)) == XlError.DIV
    assert average_cells(cast(CellValue, rng)) == XlError.DIV
    assert "S!A3" not in calls
