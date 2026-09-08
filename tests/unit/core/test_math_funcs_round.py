"""Excel ROUND uses ties-away-from-zero, not Python bankers rounding."""

from __future__ import annotations

import pytest

from excel_grapher.core.math_funcs import round_number
from excel_grapher.core.types import CellValue, XlError


@pytest.mark.parametrize(
    ("number", "num_digits", "expected"),
    [
        (1.25, 1, 1.3),
        (-1.25, 1, -1.3),
        (2.5, 0, 3.0),
        (-2.5, 0, -3.0),
        (125, -1, 130.0),
        (-125, -1, -130.0),
        (2.567, 2, 2.57),
        (1.24, 1, 1.2),
        (1.26, 1, 1.3),
        (0.5, 0, 1.0),
        (-0.5, 0, -1.0),
        (0, 0, 0.0),
        (True, 0, 1.0),
        ("1.25", 1, 1.3),
    ],
)
def test_round_number_excel_half_away_from_zero(
    number: CellValue, num_digits: CellValue, expected: float
) -> None:
    assert round_number(number, num_digits) == expected


def test_round_number_text_returns_value_error() -> None:
    assert round_number("not a number", 0) == XlError.VALUE
    assert round_number(1.25, "digits") == XlError.VALUE


def test_round_number_propagates_xl_error() -> None:
    assert round_number(XlError.DIV, 0) == XlError.DIV
    assert round_number(1.25, XlError.NA) == XlError.NA
