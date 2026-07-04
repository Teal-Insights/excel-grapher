"""Unit tests for ``xl_exp`` export runtime semantics."""

from __future__ import annotations

import math

import pytest

from excel_grapher.core.types import CellValue, XlError
from excel_grapher.runtime.math import xl_exp


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        (0, 1.0),
        (1, math.e),
        (-1, math.exp(-1)),
        (True, math.e),
        (False, 1.0),
        ("1", math.e),
        ("-2", math.exp(-2)),
    ],
)
def test_xl_exp_numeric(raw: CellValue, expected: float) -> None:
    assert xl_exp(raw) == pytest.approx(expected)


def test_xl_exp_text_returns_value_error() -> None:
    assert xl_exp("not a number") == XlError.VALUE


def test_xl_exp_propagates_xl_error() -> None:
    assert xl_exp(XlError.DIV) == XlError.DIV
    assert xl_exp(XlError.NA) == XlError.NA


def test_xl_exp_wrong_arity_returns_value_error() -> None:
    assert xl_exp() == XlError.VALUE
    assert xl_exp(1, 2) == XlError.VALUE


def test_xl_exp_overflow_returns_num_error() -> None:
    assert xl_exp(1000) == XlError.NUM
