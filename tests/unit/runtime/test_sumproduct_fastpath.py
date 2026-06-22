"""Vectorized ``xl_sumproduct`` fast path."""

from __future__ import annotations

import numpy as np

from excel_grapher.core.operators import xl_eq, xl_mul
from excel_grapher.core.operators_bench import category_column, numeric_column
from excel_grapher.core.operators_reference import reference_sumproduct_arrays
from excel_grapher.core.sumproduct import xl_sumproduct
from excel_grapher.core.types import CellValue, XlError

LARGE_SHAPE = (2_000, 1)


def _assert_matches_reference(*args: CellValue) -> None:
    arrays = [
        arg if isinstance(arg, np.ndarray) else np.array([[arg]], dtype=object) for arg in args
    ]
    expected = reference_sumproduct_arrays(arrays)
    actual = xl_sumproduct(*args)
    assert actual == expected


def test_sumproduct_fastpath_matches_reference_on_numeric_arrays() -> None:
    left = numeric_column(LARGE_SHAPE, seed=51)
    right = numeric_column(LARGE_SHAPE, seed=52)
    _assert_matches_reference(left, right)


def test_sumproduct_fastpath_matches_reference_on_criteria_chain() -> None:
    categories = category_column(LARGE_SHAPE, seed=61)
    values = numeric_column(LARGE_SHAPE, seed=62)
    criteria = xl_mul(xl_eq(categories, "Software"), values)
    assert isinstance(criteria, np.ndarray)
    _assert_matches_reference(criteria)


def test_sumproduct_fastpath_matches_reference_with_bool_mask() -> None:
    mask = np.array([[True, False], [True, True]], dtype=object)
    values = np.array([[10.0, 20.0], [30.0, 40.0]], dtype=object)
    _assert_matches_reference(mask, values)


def test_sumproduct_fastpath_falls_back_on_embedded_error() -> None:
    left = np.array([[1.0, XlError.NA]], dtype=object)
    right = np.array([[2.0, 2.0]], dtype=object)
    assert xl_sumproduct(left, right) == XlError.NA


def test_sumproduct_fastpath_shape_mismatch_returns_value() -> None:
    left = np.array([[1.0, 2.0]], dtype=object)
    right = np.array([[1.0, 2.0, 3.0]], dtype=object)
    assert xl_sumproduct(left, right) == XlError.VALUE


def test_sumproduct_empty_args_returns_zero() -> None:
    assert xl_sumproduct() == 0.0


def test_sumproduct_scalar_args() -> None:
    assert xl_sumproduct(2.0, 3.0) == 6.0
