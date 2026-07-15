"""Shared helpers for operator semantics and fast-path parity tests."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any, cast

import pytest

np = pytest.importorskip("numpy")

from excel_grapher.core.operators import (
    xl_add,
    xl_concat,
    xl_div,
    xl_eq,
    xl_ge,
    xl_gt,
    xl_le,
    xl_lt,
    xl_mul,
    xl_ne,
    xl_pow,
    xl_sub,
)
from excel_grapher.core.operators_reference import (
    broadcast_pair,
    reference_arithmetic_array,
    reference_compare_array,
    reference_concat_array,
    reference_sumproduct_arrays,
)
from excel_grapher.core.sumproduct import sumproduct_cells as xl_sumproduct
from excel_grapher.core.types import CellValue, FormulaValue, XlError

COMPARE_OPS = ("=", "<>", "<", ">", "<=", ">=")

COMPARE_DISPATCH: dict[str, Callable[[FormulaValue, FormulaValue], FormulaValue]] = {
    "=": xl_eq,
    "<>": xl_ne,
    "<": xl_lt,
    ">": xl_gt,
    "<=": xl_le,
    ">=": xl_ge,
}

ARITHMETIC_DISPATCH: dict[str, Callable[[FormulaValue, FormulaValue], FormulaValue]] = {
    "+": xl_add,
    "-": xl_sub,
    "*": xl_mul,
    "/": xl_div,
    "^": xl_pow,
}


def _to_nested_rows(value: object) -> object:
    if isinstance(value, np.ndarray):
        return cast(Any, value).tolist()
    return value


def array_tolist(value: object) -> list[list[object]]:
    """Normalize operator array results (ndarray or nested list) for assertions."""
    if isinstance(value, np.ndarray):
        return cast(Any, value).tolist()
    if isinstance(value, list):
        return cast("list[list[object]]", value)
    raise AssertionError(f"expected array result, got {type(value)!r}")


def as_ndarray(value: object) -> np.ndarray | list[list[object]]:
    return array_tolist(value)


def assert_cellvalue_equal(actual: object, expected: object) -> None:
    assert _to_nested_rows(actual) == _to_nested_rows(expected)


def reference_compare(op: str, left: CellValue, right: CellValue) -> object:
    pair = broadcast_pair(left, right)
    if isinstance(pair, XlError):
        return pair
    return reference_compare_array(op, pair[0], pair[1])


def reference_arithmetic(op: str, left: CellValue, right: CellValue) -> object:
    pair = broadcast_pair(left, right)
    if isinstance(pair, XlError):
        return pair
    return reference_arithmetic_array(op, pair[0], pair[1])


def reference_concat(left: CellValue, right: CellValue) -> object:
    pair = broadcast_pair(left, right)
    if isinstance(pair, XlError):
        return pair
    return reference_concat_array(pair[0], pair[1])


def _assert_public_matches_reference(
    left: CellValue,
    right: CellValue,
    *,
    expected: object,
    actual: object,
) -> None:
    if isinstance(expected, XlError):
        assert actual == expected
        return
    assert isinstance(actual, (np.ndarray, list))
    assert_cellvalue_equal(actual, expected)


def assert_compare_matches_reference(op: str, left: CellValue, right: CellValue) -> None:
    pair = broadcast_pair(left, right)
    assert not isinstance(pair, XlError)
    expected = reference_compare_array(op, pair[0], pair[1])
    actual = COMPARE_DISPATCH[op](left, right)
    _assert_public_matches_reference(left, right, expected=expected, actual=actual)


def assert_arithmetic_matches_reference(op: str, left: CellValue, right: CellValue) -> None:
    pair = broadcast_pair(left, right)
    assert not isinstance(pair, XlError)
    expected = reference_arithmetic_array(op, pair[0], pair[1])
    actual = ARITHMETIC_DISPATCH[op](left, right)
    _assert_public_matches_reference(left, right, expected=expected, actual=actual)


def assert_concat_matches_reference(left: CellValue, right: CellValue) -> None:
    pair = broadcast_pair(left, right)
    assert not isinstance(pair, XlError)
    expected = reference_concat_array(pair[0], pair[1])
    actual = xl_concat(left, right)
    assert isinstance(actual, (np.ndarray, list))
    assert_cellvalue_equal(actual, expected)


def assert_sumproduct_matches_reference(*args: CellValue) -> None:
    arrays = [
        arg if isinstance(arg, np.ndarray) else np.array([[arg]], dtype=object) for arg in args
    ]
    expected = reference_sumproduct_arrays(arrays)
    assert xl_sumproduct(*args) == expected
