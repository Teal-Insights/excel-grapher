"""Unit tests for ``xl_abs`` export runtime semantics."""

from __future__ import annotations

import pytest

from excel_grapher.core.types import CellValue, XlError
from excel_grapher.runtime.math import xl_abs


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        (-3.5, 3.5),
        (3.5, 3.5),
        (0, 0.0),
        (True, 1.0),
        (False, 0.0),
        ("-2", 2.0),
        ("2.5", 2.5),
    ],
)
def test_xl_abs_numeric(raw: CellValue, expected: float) -> None:
    assert xl_abs(raw) == expected


def test_xl_abs_text_returns_value_error() -> None:
    assert xl_abs("not a number") == XlError.VALUE


def test_xl_abs_propagates_xl_error() -> None:
    assert xl_abs(XlError.DIV) == XlError.DIV
    assert xl_abs(XlError.NA) == XlError.NA


def test_xl_abs_wrong_arity_returns_value_error() -> None:
    assert xl_abs() == XlError.VALUE
    assert xl_abs(1, 2) == XlError.VALUE
