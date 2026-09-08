"""Unit tests for ``xl_round`` runtime semantics."""

from __future__ import annotations

import pytest

from excel_grapher.core.types import CellValue, XlError
from excel_grapher.runtime.math import xl_round


@pytest.mark.parametrize(
    ("number", "num_digits", "expected"),
    [
        (1.25, 1, 1.3),
        (-1.25, 1, -1.3),
        (2.5, 0, 3.0),
        (125, -1, 130.0),
    ],
)
def test_xl_round_excel_half_away_from_zero(
    number: CellValue, num_digits: CellValue, expected: float
) -> None:
    assert xl_round(number, num_digits) == expected


def test_xl_round_text_returns_value_error() -> None:
    assert xl_round("not a number", 0) == XlError.VALUE


def test_xl_round_propagates_xl_error() -> None:
    assert xl_round(XlError.DIV, 0) == XlError.DIV
    assert xl_round(1.25, XlError.NA) == XlError.NA
