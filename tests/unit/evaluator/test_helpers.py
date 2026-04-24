import numpy as np
import pytest

from excel_grapher.evaluator.helpers import (
    to_bool,
    to_number,
    to_string,
    xl_eq,
    xl_ge,
    xl_gt,
    xl_le,
    xl_lt,
    xl_ne,
)
from excel_grapher.evaluator.types import CellValue, XlError


def test_to_number_basic_types() -> None:
    assert to_number(None) == 0.0
    assert to_number(True) == 1.0
    assert to_number(False) == 0.0
    assert to_number(3) == 3.0
    assert to_number(3.5) == 3.5
    assert to_number("  ") == 0.0
    assert to_number("2") == 2.0
    assert to_number("2.5") == 2.5
    assert to_number("abc") == XlError.VALUE
    assert to_number(XlError.DIV) == XlError.DIV


def test_to_string_basic_types() -> None:
    assert to_string(None) == ""
    assert to_string(True) == "TRUE"
    assert to_string(False) == "FALSE"
    assert to_string(2.0) == "2"
    assert to_string(2.25) == "2.25"
    assert to_string("abc") == "abc"
    assert to_string(XlError.NA) == "#N/A"


def test_to_bool_basic_types() -> None:
    assert to_bool(None) is False
    assert to_bool(True) is True
    assert to_bool(False) is False
    assert to_bool(0) is False
    assert to_bool(0.0) is False
    assert to_bool(2) is True
    assert to_bool(-1.0) is True
    assert to_bool("") is False
    assert to_bool("true") is True
    assert to_bool("FALSE") is False
    assert to_bool("nope") == XlError.VALUE
    assert to_bool(XlError.REF) == XlError.REF


def test_helpers_accept_numpy_scalars() -> None:
    assert to_number(np.float64(3.5)) == 3.5


@pytest.mark.parametrize(
    ("left", "right", "expected"),
    [
        # --- Numeric-string coercion (matches FormulaEvaluator behavior) ---
        ("0", 0, True),
        ("  2.0 ", 2, True),
        ("", 0, True),  # empty string coerces to 0.0 via to_number
        (None, 0, True),  # None coerces to 0.0 via to_number
        # --- Non-numeric strings fall back to case-insensitive string comparison ---
        ("TRUE", True, True),  # to_number("TRUE") fails -> compare to_string values
        ("FALSE", False, True),
        ("abc", 0, False),  # compare "abc" vs "0" as strings
        ("AbC", "aBc", True),
        # --- Error propagation ---
        (XlError.NA, 0, XlError.NA),
        (0, XlError.REF, XlError.REF),
    ],
)
def test_xl_eq_semantics(left: CellValue, right: CellValue, expected: bool | XlError) -> None:
    assert xl_eq(left, right) == expected


@pytest.mark.parametrize(
    ("left", "right", "expected"),
    [
        # Numeric-string coercion
        ("0", 0, False),
        ("-1", 0, True),
        ("", 0, False),
        (None, 0, False),
        # Fallback-to-string
        ("abc", 0, False),  # "abc" < "0" is False (string compare)
        ("0", "abc", True),  # "0" < "abc" is True (string compare)
        ("AbC", "b", True),
        # Errors
        (XlError.VALUE, 0, XlError.VALUE),
        (0, XlError.DIV, XlError.DIV),
    ],
)
def test_xl_lt_semantics(left: CellValue, right: CellValue, expected: bool | XlError) -> None:
    assert xl_lt(left, right) == expected


@pytest.mark.parametrize(
    ("left", "right", "expected"),
    [
        # Numeric-string coercion
        ("0", 0, False),
        ("1", 0, True),
        ("", 0, False),
        (None, 0, False),
        # Fallback-to-string
        ("abc", 0, True),  # "abc" > "0" is True (string compare)
        ("0", "abc", False),
        ("b", "AbC", True),
        # Errors
        (XlError.NAME, 0, XlError.NAME),
        (0, XlError.NUM, XlError.NUM),
    ],
)
def test_xl_gt_semantics(left: CellValue, right: CellValue, expected: bool | XlError) -> None:
    assert xl_gt(left, right) == expected


@pytest.mark.parametrize(
    ("left", "right", "expected"),
    [
        ("0", 0, True),  # numeric coercion
        ("", 0, True),
        ("0", 1, True),
        ("abc", 0, False),  # "abc" <= "0" is False (string compare)
        (XlError.NA, 0, XlError.NA),
    ],
)
def test_xl_le_semantics(left: CellValue, right: CellValue, expected: bool | XlError) -> None:
    assert xl_le(left, right) == expected


@pytest.mark.parametrize(
    ("left", "right", "expected"),
    [
        ("0", 0, True),  # numeric coercion
        ("", 0, True),
        ("1", 0, True),
        ("0", 1, False),
        ("abc", 0, True),  # "abc" >= "0" is True (string compare)
        (XlError.REF, 0, XlError.REF),
    ],
)
def test_xl_ge_semantics(left: CellValue, right: CellValue, expected: bool | XlError) -> None:
    assert xl_ge(left, right) == expected


@pytest.mark.parametrize(
    ("left", "right", "expected"),
    [
        ("0", 0, False),
        ("", 0, False),
        ("AbC", "aBc", False),
        ("abc", 0, True),
        (XlError.NA, 0, XlError.NA),
    ],
)
def test_xl_ne_semantics(left: CellValue, right: CellValue, expected: bool | XlError) -> None:
    assert xl_ne(left, right) == expected
