"""Aggregates ignore non-numeric text like Excel (#420 follow-up).

Arithmetic still raises `#VALUE!` on empty text via `to_number`; SUM/AVERAGE/
MIN/MAX/STDEV skip text cells (including guard empty text) and keep
propagating embedded errors.
"""

from __future__ import annotations

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
from excel_grapher.core.operators import xl_add
from excel_grapher.core.sumproduct import sumproduct_cells


@pytest.mark.parametrize(
    ("func", "args", "expected"),
    [
        (sum_cells, (1, "", 2), 3.0),
        (sum_cells, (1, " ", "abc", 2), 3.0),
        (sum_cells, ("", "  "), 0.0),
        (average_cells, (2, "", 6), 4.0),
        (average_cells, ("x", 10), 10.0),
        (min_cells, (3, "", 1, "nope"), 1.0),
        (max_cells, (3, "", 1, "nope"), 3.0),
        (stdev_cells, (1, "", 3), pytest.approx(2.0**0.5)),
    ],
)
def test_aggregates_skip_non_numeric_text(func, args, expected) -> None:
    assert func(*args) == expected


def test_sum_still_propagates_embedded_errors() -> None:
    assert sum_cells(1, XlError.DIV, 2) == XlError.DIV


def test_sum_skips_empty_text_in_lazy_range() -> None:
    values = {"S!A1": 1, "S!A2": "", "S!A3": 2}

    def resolve(address: str) -> CellValue:
        return values[address]

    rng = Range("S", 1, 1, 3, 1, resolve)
    assert sum_cells(cast(CellValue, rng)) == 3.0


def test_arithmetic_still_rejects_empty_text() -> None:
    """Companion invariant: operators stay strict after aggregate skip fix."""
    assert xl_add(1, "") == XlError.VALUE


def test_sumproduct_treats_text_as_zero() -> None:
    """SUMPRODUCT coerces non-numeric text to 0 (Excel), not `#VALUE!`."""
    assert sumproduct_cells([1, ""], [1, 1]) == 1.0
    assert sumproduct_cells([1, "x"], [2, 3]) == 2.0
