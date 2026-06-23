"""Tests for top-level array result helpers."""

from __future__ import annotations

import numpy as np

from excel_grapher.core.array_results import array_values_equal, is_array_result
from excel_grapher.core.types import XlError


def test_is_array_result_true_for_multicell_ndarray() -> None:
    value = np.array([[True], [False]], dtype=object)
    assert is_array_result(value)


def test_is_array_result_false_for_scalar_and_1x1() -> None:
    assert not is_array_result(True)
    assert not is_array_result(np.array([[1.0]], dtype=object))


def test_array_values_equal_matches_bool_columns() -> None:
    left = np.array([[True], [False], [True]], dtype=object)
    right = np.array([[True], [False], [True]], dtype=object)
    assert array_values_equal(left, right)


def test_array_values_equal_rejects_shape_mismatch() -> None:
    left = np.array([[True], [False]], dtype=object)
    right = np.array([[True, False]], dtype=object)
    assert not array_values_equal(left, right)


def test_array_values_equal_compares_xlerror_cells() -> None:
    left = np.array([[XlError.VALUE, 1]], dtype=object)
    right = np.array([[XlError.VALUE, 1]], dtype=object)
    assert array_values_equal(left, right)
