"""Reference (per-cell loop) array implementations for Excel binary operators.

These functions preserve the Sprint 0 semantics contract used as the golden
reference when adding vectorized fast paths in later sprints.
"""

from __future__ import annotations

import numpy as np

from .coercions import excel_casefold, to_number, to_string
from .types import CellValue, XlError


def broadcast_pair(
    left: CellValue,
    right: CellValue,
) -> tuple[np.ndarray, np.ndarray] | XlError:
    """Broadcast scalar/array operands to matching object ndarrays."""
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right
    if isinstance(left, np.ndarray) and isinstance(right, np.ndarray):
        if left.shape != right.shape:
            return XlError.VALUE
        return left, right
    if isinstance(left, np.ndarray):
        return left, np.full(left.shape, right, dtype=object)
    if isinstance(right, np.ndarray):
        return np.full(right.shape, left, dtype=object), right
    raise TypeError("expected at least one ndarray operand")


def compare_scalars(op: str, left: CellValue, right: CellValue) -> bool | XlError:
    """Compare two scalar cell values using Excel coercion rules."""
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right

    def _cmp_str(a: str, b: str) -> bool:
        if op == "=":
            return a == b
        if op == "<>":
            return a != b
        if op == "<":
            return a < b
        if op == ">":
            return a > b
        if op == "<=":
            return a <= b
        if op == ">=":
            return a >= b
        raise ValueError(f"Unknown comparison operator: {op}")

    def _cmp_float(a: float, b: float) -> bool:
        if op == "=":
            return a == b
        if op == "<>":
            return a != b
        if op == "<":
            return a < b
        if op == ">":
            return a > b
        if op == "<=":
            return a <= b
        if op == ">=":
            return a >= b
        raise ValueError(f"Unknown comparison operator: {op}")

    if isinstance(left, str) and isinstance(right, str):
        return _cmp_str(excel_casefold(left), excel_casefold(right))

    ln = to_number(left)
    rn = to_number(right)
    if isinstance(ln, XlError) or isinstance(rn, XlError):
        return _cmp_str(excel_casefold(to_string(left)), excel_casefold(to_string(right)))

    return _cmp_float(float(ln), float(rn))


def concat_scalars(left: CellValue, right: CellValue) -> str:
    return to_string(left) + to_string(right)


def apply_arithmetic(op: str, ln: float, rn: float) -> float | XlError:
    """Apply an Excel arithmetic operator to two coerced numbers."""
    if op == "+":
        return ln + rn
    if op == "-":
        return ln - rn
    if op == "*":
        return ln * rn
    if op == "/":
        if rn == 0:
            return XlError.DIV
        return ln / rn
    if op == "^":
        try:
            value = ln**rn
        except (ValueError, OverflowError):
            return XlError.NUM
        if isinstance(value, complex):
            return XlError.NUM
        return value
    raise ValueError(f"Unknown arithmetic operator: {op}")


def reference_compare_array(
    op: str,
    arr_left: np.ndarray,
    arr_right: np.ndarray,
) -> np.ndarray | XlError:
    """Element-wise comparison over broadcast object ndarrays (C-order, fail-fast)."""
    result = np.empty(arr_left.shape, dtype=object)
    for indices in np.ndindex(arr_left.shape):
        cell = compare_scalars(op, arr_left[indices], arr_right[indices])
        if isinstance(cell, XlError):
            return cell
        result[indices] = cell
    return result


def reference_arithmetic_array(
    op: str,
    arr_left: np.ndarray,
    arr_right: np.ndarray,
) -> np.ndarray | XlError:
    """Element-wise arithmetic over broadcast object ndarrays (C-order, fail-fast)."""
    result = np.empty(arr_left.shape, dtype=object)
    for indices in np.ndindex(arr_left.shape):
        ln = to_number(arr_left[indices])
        rn = to_number(arr_right[indices])
        if isinstance(ln, XlError):
            return ln
        if isinstance(rn, XlError):
            return rn
        if op == "+":
            result[indices] = ln + rn
        elif op == "-":
            result[indices] = ln - rn
        elif op == "*":
            result[indices] = ln * rn
        elif op == "/":
            if rn == 0:
                return XlError.DIV
            result[indices] = ln / rn
        elif op == "^":
            try:
                value = ln**rn
            except (ValueError, OverflowError):
                return XlError.NUM
            if isinstance(value, complex):
                return XlError.NUM
            result[indices] = value
        else:
            raise ValueError(f"Unknown arithmetic operator: {op}")
    return result


def reference_concat_array(
    arr_left: np.ndarray,
    arr_right: np.ndarray,
) -> np.ndarray:
    """Element-wise string concatenation over broadcast object ndarrays."""
    result = np.empty(arr_left.shape, dtype=object)
    for indices in np.ndindex(arr_left.shape):
        result[indices] = concat_scalars(arr_left[indices], arr_right[indices])
    return result


def reference_sumproduct_arrays(arrays: list[np.ndarray]) -> float | XlError:
    """Element-wise product reduction (C-order, fail-fast on ``to_number`` errors)."""
    if not arrays:
        return 0.0
    shape = arrays[0].shape
    result = 0.0
    for indices in np.ndindex(shape):
        product = 1.0
        for arr in arrays:
            number = to_number(arr[indices])
            if isinstance(number, XlError):
                return number
            product *= number
        result += product
    return result
