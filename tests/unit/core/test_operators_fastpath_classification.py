"""Operand cell-type classification drives fast-path tier selection.

`try_fastpath_compare_array` used to answer every tier's yes/no question with
its own full Python scan of both operands. The classification is derived once
per operand instead; these tests pin the semantics that gating depends on —
error precedence, `str` subclasses, and coercions that only the per-cell
fallbacks can do.
"""

# ruff: noqa: E402
from __future__ import annotations

import pytest

np = pytest.importorskip("numpy")

from excel_grapher.core.operator_thresholds import MIN_OPERATOR_FASTPATH_CELLS
from excel_grapher.core.operators import xl_eq, xl_gt
from excel_grapher.core.sumproduct import sumproduct_cells
from excel_grapher.core.types import XlError
from tests.unit.core.operators_test_helpers import array_tolist, assert_compare_matches_reference

LARGE = MIN_OPERATOR_FASTPATH_CELLS * 2


def _string_column(size: int, value: str = "Software") -> np.ndarray:
    return np.array([[value] for _ in range(size)], dtype=object)


def test_error_inside_an_otherwise_all_string_column_wins_over_string_compare() -> None:
    """`XlError` is a `str` subclass; it must fail fast, not casefold-compare."""
    left = _string_column(LARGE)
    left[LARGE // 2, 0] = XlError.NA
    right = _string_column(LARGE)
    assert xl_eq(left, right) == XlError.NA


def test_error_inside_a_numeric_string_column_wins_over_numeric_compare() -> None:
    left = np.array([[str(i)] for i in range(LARGE)], dtype=object)
    left[7, 0] = XlError.DIV
    right = np.array([[float(i)] for i in range(LARGE)], dtype=object)
    assert xl_eq(left, right) == XlError.DIV


def test_first_error_in_c_order_still_wins_across_both_operands() -> None:
    left = _string_column(LARGE)
    right = _string_column(LARGE)
    left[9, 0] = XlError.REF
    right[4, 0] = XlError.VALUE
    assert xl_eq(left, right) == XlError.VALUE


def test_numpy_string_cells_still_compare_case_insensitively() -> None:
    """`np.str_` cells are `str` subclasses and must keep Excel compare semantics."""
    left = np.array([[np.str_("Software")] for _ in range(LARGE)], dtype=object)
    right = _string_column(LARGE, "SOFTWARE")
    assert array_tolist(xl_eq(left, right)) == [[True]] * LARGE


def test_numeric_string_column_with_one_blank_cell_matches_reference() -> None:
    """A blank string cannot coerce, so the whole operand falls back per cell."""
    left = np.array([[str(i + 1)] for i in range(LARGE)], dtype=object)
    left[3, 0] = ""
    right = np.array([[float(i + 1)] for i in range(LARGE)], dtype=object)
    assert_compare_matches_reference("=", left, right)


def test_unicode_whitespace_numeric_strings_match_reference() -> None:
    left = np.array([[f" {i + 1} "] for i in range(LARGE)], dtype=object)
    right = np.array([[float(i + 1)] for i in range(LARGE)], dtype=object)
    assert_compare_matches_reference("<=", left, right)


def test_bool_cells_compare_as_numbers_against_a_threshold() -> None:
    left = np.array([[i % 2 == 0] for i in range(LARGE)], dtype=object)
    assert_compare_matches_reference(">", left, 0.5)


def test_none_cells_still_coerce_to_zero_not_nan() -> None:
    left = np.array([[None] for _ in range(LARGE)], dtype=object)
    assert array_tolist(xl_gt(left, -1.0)) == [[True]] * LARGE


def test_sumproduct_still_treats_plain_text_as_zero_at_scale() -> None:
    values = [1.0] * LARGE
    weights: list[object] = [2.0] * LARGE
    weights[5] = "x"
    assert sumproduct_cells(values, weights) == pytest.approx(2.0 * (LARGE - 1))


def test_sumproduct_propagates_embedded_errors_at_scale() -> None:
    values: list[object] = [1.0] * LARGE
    values[2] = XlError.NA
    assert sumproduct_cells(values, [1.0] * LARGE) == XlError.NA
