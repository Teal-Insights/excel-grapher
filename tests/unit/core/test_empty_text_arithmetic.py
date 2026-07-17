"""Empty text vs blank cell coercion in scalar arithmetic (#420).

Excel coerces blank cells to 0 in arithmetic, but empty-string / whitespace-only
text never coerces and raises `#VALUE!`.
"""

from __future__ import annotations

import pytest

from excel_grapher.core.coercions import to_number, try_coerce_string_to_float
from excel_grapher.core.operators import (
    xl_add,
    xl_div,
    xl_mul,
    xl_neg,
    xl_percent,
    xl_pow,
    xl_sub,
)
from excel_grapher.core.types import XlError


@pytest.mark.parametrize("text", ["", " ", "  ", "\t", "\n"])
def test_try_coerce_string_to_float_rejects_empty_and_whitespace(text: str) -> None:
    assert try_coerce_string_to_float(text) is None


@pytest.mark.parametrize("text", ["", " ", "  ", "\t"])
def test_to_number_rejects_empty_and_whitespace_text(text: str) -> None:
    assert to_number(text) == XlError.VALUE


def test_to_number_still_coerces_blank_none_to_zero() -> None:
    assert to_number(None) == 0.0


@pytest.mark.parametrize(
    ("op", "left", "right"),
    [
        (xl_add, 1, ""),
        (xl_add, "", 1),
        (xl_sub, 1, " "),
        (xl_mul, 1, "\t"),
        (xl_div, 1, ""),
        (xl_pow, 1, ""),
        (xl_add, 1, "  "),
    ],
)
def test_scalar_arithmetic_rejects_empty_text(op, left, right) -> None:
    assert op(left, right) == XlError.VALUE


@pytest.mark.parametrize(
    ("op", "left", "right", "expected"),
    [
        (xl_add, 1, None, 1.0),
        (xl_div, 1, None, XlError.DIV),
        (xl_mul, None, 5, 0.0),
        (xl_sub, 3, None, 3.0),
    ],
)
def test_scalar_arithmetic_blank_none_still_coerces_to_zero(op, left, right, expected) -> None:
    assert op(left, right) == expected


def test_unary_ops_reject_empty_text() -> None:
    assert xl_neg("") == XlError.VALUE
    assert xl_neg("  ") == XlError.VALUE
    assert xl_percent("") == XlError.VALUE


def test_unary_ops_blank_none_still_coerces_to_zero() -> None:
    assert xl_neg(None) == 0.0
    assert xl_percent(None) == 0.0
