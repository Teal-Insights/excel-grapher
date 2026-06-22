"""Vectorized fast paths for Excel binary operators (numeric arithmetic)."""

from __future__ import annotations

import numpy as np

from .coercions import to_number
from .types import XlError

_NUMERIC_CELL_TYPES = (int, float, np.integer, np.floating)


def _try_asarray_float64(arr: np.ndarray) -> np.ndarray | None:
    """Return a float64 view when every cell is a plain int or float."""
    if arr.dtype in (np.float64, np.int64, np.float32, np.int32):
        return np.asarray(arr, dtype=np.float64)
    if arr.dtype != object:
        return None
    flat = arr.ravel()
    if flat.size == 0:
        return np.empty(arr.shape, dtype=np.float64)
    for value in flat:
        if not isinstance(value, _NUMERIC_CELL_TYPES):
            return None
    return np.asarray(arr, dtype=np.float64)


def batch_coerce_to_float64(arr: np.ndarray) -> np.ndarray | None:
    """Coerce an object ndarray to float64, or return None when any cell fails."""
    direct = _try_asarray_float64(arr)
    if direct is not None:
        return direct

    flat = arr.ravel()
    out = np.empty(flat.size, dtype=np.float64)
    for index, value in enumerate(flat):
        if isinstance(value, XlError):
            return None
        if value is None:
            out[index] = 0.0
            continue
        if isinstance(value, bool):
            out[index] = 1.0 if value else 0.0
            continue
        if isinstance(value, (int, float, np.integer, np.floating)):
            out[index] = float(value)
            continue
        if isinstance(value, str):
            number = to_number(value)
            if isinstance(number, XlError):
                return None
            out[index] = float(number)
            continue
        return None
    return out.reshape(arr.shape)


def _first_zero_index(values: np.ndarray) -> int | None:
    """Return the C-order flat index of the first zero, if any."""
    flat = values.ravel()
    matches = np.flatnonzero(flat == 0.0)
    if matches.size == 0:
        return None
    return int(matches[0])


def _fastpath_pow(left: np.ndarray, right: np.ndarray) -> np.ndarray | XlError:
    """Element-wise power on pre-coerced float arrays (C-order fail-fast)."""
    flat_left = left.ravel()
    flat_right = right.ravel()
    out = np.empty(flat_left.size, dtype=np.float64)
    for index in range(flat_left.size):
        base = float(flat_left[index])
        exponent = float(flat_right[index])
        try:
            value = base**exponent
        except (ValueError, OverflowError):
            return XlError.NUM
        if isinstance(value, complex):
            return XlError.NUM
        out[index] = value
    return out.reshape(left.shape).astype(object)


def _float_result_as_object(values: np.ndarray) -> np.ndarray:
    return values.astype(object)


def try_fastpath_arithmetic_array(
    op: str,
    arr_left: np.ndarray,
    arr_right: np.ndarray,
) -> np.ndarray | XlError | None:
    """Apply a vectorized numeric fast path, or return None to use the reference loop."""
    left = batch_coerce_to_float64(arr_left)
    if left is None:
        return None
    right = batch_coerce_to_float64(arr_right)
    if right is None:
        return None

    if op == "+":
        return _float_result_as_object(left + right)
    if op == "-":
        return _float_result_as_object(left - right)
    if op == "*":
        return _float_result_as_object(left * right)
    if op == "/":
        if _first_zero_index(right) is not None:
            return XlError.DIV
        return _float_result_as_object(left / right)
    if op == "^":
        return _fastpath_pow(left, right)
    raise ValueError(f"Unknown arithmetic operator: {op}")
