"""Reference (per-cell loop) array implementations for Excel binary operators.

These functions preserve the Sprint 0 semantics contract used as the golden
reference when adding vectorized fast paths in later sprints.

Array helpers lazily import NumPy so the default (no-`fast`) install can load
this module without the accelerator. Scalar helpers stay NumPy-free and are
shared with `operator_maps` / export.
"""

from __future__ import annotations

from typing import Any

from .coercions import excel_casefold, to_number, to_string
from .types import CellValue, FormulaValue, XlError


def broadcast_pair(
    left: CellValue,
    right: CellValue,
) -> tuple[Any, Any] | XlError:
    """Broadcast scalar/array operands to matching object ndarrays."""
    import numpy as np

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


def compare_scalars(op: str, left: FormulaValue, right: FormulaValue) -> bool | XlError:
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

    # Exact empty text compares as 0 (Excel); whitespace-only does not coerce.
    if isinstance(left, str) and left == "":
        left = 0.0
    if isinstance(right, str) and right == "":
        right = 0.0

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
    arr_left: Any,
    arr_right: Any,
) -> Any | XlError:
    """Element-wise comparison over broadcast object ndarrays (C-order, fail-fast)."""
    import numpy as np

    result = np.empty(arr_left.shape, dtype=object)
    for indices in np.ndindex(arr_left.shape):
        cell = compare_scalars(op, arr_left[indices], arr_right[indices])
        if isinstance(cell, XlError):
            return cell
        result[indices] = cell
    return result


def reference_arithmetic_array(
    op: str,
    arr_left: Any,
    arr_right: Any,
) -> Any | XlError:
    """Element-wise arithmetic over broadcast object ndarrays (C-order, fail-fast)."""
    import numpy as np

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
    arr_left: Any,
    arr_right: Any,
) -> Any:
    """Element-wise string concatenation over broadcast object ndarrays."""
    import numpy as np

    result = np.empty(arr_left.shape, dtype=object)
    for indices in np.ndindex(arr_left.shape):
        result[indices] = concat_scalars(arr_left[indices], arr_right[indices])
    return result


def reference_sumproduct_arrays(arrays: list[Any]) -> float | XlError:
    """Element-wise product reduction (C-order, fail-fast on ``XlError``).

    Non-numeric text is treated as 0 (Excel SUMPRODUCT). `XlError` is a
    `StrEnum`, so error sentinels are checked before the text branch.
    """
    import numpy as np

    if not arrays:
        return 0.0
    shape = arrays[0].shape
    result = 0.0
    for indices in np.ndindex(shape):
        product = 1.0
        for arr in arrays:
            cell = arr[indices]
            if isinstance(cell, XlError):
                return cell
            if isinstance(cell, str):
                number = 0.0
            else:
                number = to_number(cell)
                if isinstance(number, XlError):
                    return number
            product *= number
        result += product
    return result
