"""Element-wise array semantics for shared Excel operators."""

# ruff: noqa: E402
from __future__ import annotations

import pytest

np = pytest.importorskip("numpy")

from excel_grapher.core.operators import xl_concat, xl_eq, xl_mul, xl_pow
from excel_grapher.core.types import XlError
from tests.unit.core.operators_test_helpers import array_tolist


def test_xl_pow_elementwise_array_and_scalar() -> None:
    left = np.array([[2.0, 3.0], [4.0, 5.0]], dtype=object)
    result = xl_pow(left, 2)
    rows = array_tolist(result)
    assert len(rows) == 2 and len(rows[0]) == 2
    assert rows == [[4.0, 9.0], [16.0, 25.0]]


def test_xl_pow_broadcasts_scalar_base_across_array_exponent() -> None:
    right = np.array([[2.0, 3.0]], dtype=object)
    assert array_tolist(xl_pow(3.0, right)) == [[9.0, 27.0]]


def test_xl_pow_shape_mismatch_returns_value() -> None:
    left = np.array([[1.0, 2.0]], dtype=object)
    right = np.array([[1.0, 2.0, 3.0]], dtype=object)
    assert xl_pow(left, right) == XlError.VALUE


def test_xl_pow_propagates_top_level_errors() -> None:
    left = np.array([[2.0]], dtype=object)
    assert xl_pow(XlError.DIV, left) == XlError.DIV
    assert xl_pow(left, XlError.NA) == XlError.NA


def test_xl_pow_invalid_power_returns_num() -> None:
    assert xl_pow(-1.0, 0.5) == XlError.NUM


def test_xl_eq_elementwise_string_equality() -> None:
    left = np.array([["Software", "Hardware"], ["Software", "Other"]], dtype=object)
    assert array_tolist(xl_eq(left, "Software")) == [[True, False], [True, False]]


def test_xl_eq_array_compare_fail_fast_on_first_cell_error() -> None:
    left = np.array([[1.0, XlError.NA]], dtype=object)
    assert xl_eq(left, 1.0) == XlError.NA


def test_xl_mul_array_arithmetic_fail_fast_on_first_cell_error() -> None:
    left = np.array([[True, XlError.DIV]], dtype=object)
    assert xl_mul(left, 1) == XlError.DIV


def test_xl_concat_elementwise_arrays() -> None:
    left = np.array([["a", "b"], ["c", "d"]], dtype=object)
    right = np.array([["1", "2"], ["3", "4"]], dtype=object)
    assert array_tolist(xl_concat(left, right)) == [["a1", "b2"], ["c3", "d4"]]


def test_xl_concat_broadcasts_scalar_suffix() -> None:
    left = np.array([["x", "y"]], dtype=object)
    assert array_tolist(xl_concat(left, "!")) == [["x!", "y!"]]


def test_xl_concat_broadcasts_scalar_prefix() -> None:
    right = np.array([[1.0, 2.0]], dtype=object)
    assert array_tolist(xl_concat("v", right)) == [["v1", "v2"]]


def test_xl_concat_shape_mismatch_returns_value() -> None:
    left = np.array([["a"]], dtype=object)
    right = np.array([["b", "c"]], dtype=object)
    assert xl_concat(left, right) == XlError.VALUE


def test_xl_concat_propagates_top_level_errors() -> None:
    left = np.array([["a"]], dtype=object)
    assert xl_concat(XlError.REF, left) == XlError.REF
    assert xl_concat(left, XlError.VALUE) == XlError.VALUE


@pytest.mark.parametrize(
    ("left", "right", "expected"),
    [
        (2.0, 3.0, 8.0),
        ("a", "b", "ab"),
    ],
)
def test_xl_pow_and_concat_scalar_paths_unchanged(left, right, expected) -> None:
    if isinstance(expected, str):
        assert xl_concat(left, right) == expected
    else:
        assert xl_pow(left, right) == expected
