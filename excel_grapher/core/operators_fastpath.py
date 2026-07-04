"""Vectorized fast paths for Excel binary operators and SUMPRODUCT reduction.

``+``, ``-``, ``*``, and ``/`` use NumPy element-wise ops on coerced float64 arrays.
The ``^`` exponent operator is not vectorized: it runs a C-order Python loop so
fail-fast ``#NUM!`` semantics match the reference implementation.
"""

from __future__ import annotations

import operator
from dataclasses import dataclass

import numpy as np

from .coercions import excel_casefold, to_number, to_string, try_coerce_string_to_float
from .types import XlError

MIN_OPERATOR_FASTPATH_CELLS = 64
_NUMERIC_CELL_TYPES = (int, float, np.integer, np.floating)
_CASEFOLD_UFUNC = np.frompyfunc(excel_casefold, 1, 1)
_TO_STRING_UFUNC = np.frompyfunc(to_string, 1, 1)


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


def _try_batch_coerce_numeric_strings(arr: np.ndarray) -> np.ndarray | None:
    """Coerce an all-string object ndarray to float64 without per-cell ``to_number`` calls."""
    if arr.dtype != object:
        return None

    flat = arr.ravel()
    if flat.size == 0:
        return np.empty(arr.shape, dtype=np.float64)
    if not isinstance(flat[0], str):
        return None

    out = np.empty(flat.size, dtype=np.float64)
    for index, value in enumerate(flat):
        if not isinstance(value, str):
            return None
        number = try_coerce_string_to_float(value)
        if number is None:
            return None
        out[index] = number
    return out.reshape(arr.shape)


def batch_coerce_to_float64(arr: np.ndarray) -> np.ndarray | None:
    """Coerce an object ndarray to float64, or return None when any cell fails."""
    direct = _try_asarray_float64(arr)
    if direct is not None:
        return direct

    numeric_strings = _try_batch_coerce_numeric_strings(arr)
    if numeric_strings is not None:
        return numeric_strings

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
    """Element-wise power on pre-coerced float arrays (C-order fail-fast).

    Uses Python ``**`` per cell rather than ``np.power`` so complex results and
    exceptions align with ``operators_reference``.
    """
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


def _bool_result_as_object(values: np.ndarray) -> np.ndarray:
    return values.astype(object)


def _float_result_as_object(values: np.ndarray) -> np.ndarray:
    return values.astype(object)


def _first_paired_error_c_order(
    arr_left: np.ndarray,
    arr_right: np.ndarray,
) -> XlError | None:
    """Return the first embedded ``XlError`` in C-order (left operand wins per cell)."""
    for left_value, right_value in zip(arr_left.ravel(), arr_right.ravel(), strict=True):
        if isinstance(left_value, XlError):
            return left_value
        if isinstance(right_value, XlError):
            return right_value
    return None


@dataclass(frozen=True)
class _StringArrayMeta:
    """Summary of a homogeneous string ndarray gathered in one C-order scan."""

    is_constant: bool
    constant_value: str
    ascii_only: bool


def _string_array_meta(arr: np.ndarray) -> _StringArrayMeta | None:
    """Classify a string-only ndarray, or return None when any cell is not ``str``."""
    if arr.dtype.kind in "SU" or arr.dtype == object:
        flat = arr.ravel()
    else:
        return None

    if flat.size == 0:
        return _StringArrayMeta(is_constant=True, constant_value="", ascii_only=True)

    first = flat[0]
    if not isinstance(first, str):
        return None

    if flat.size == 1 or (flat[-1] is first and flat[flat.size // 2] is first):
        return _StringArrayMeta(
            is_constant=True,
            constant_value=first,
            ascii_only=first.isascii(),
        )

    is_constant = True
    ascii_only = first.isascii()
    for value in flat[1:]:
        if not isinstance(value, str):
            return None
        if value != first:
            is_constant = False
        if not value.isascii():
            ascii_only = False

    return _StringArrayMeta(
        is_constant=is_constant,
        constant_value=first,
        ascii_only=ascii_only,
    )


def _as_unicode_strings(arr: np.ndarray) -> np.ndarray:
    if arr.dtype.kind in "SU":
        return arr
    return np.asarray(arr, dtype=np.str_)


def _fold_string_column(arr: np.ndarray, meta: _StringArrayMeta) -> np.ndarray:
    unicode_arr = _as_unicode_strings(arr)
    if meta.ascii_only:
        return np.char.lower(unicode_arr)
    return np.asarray(_CASEFOLD_UFUNC(unicode_arr), dtype=np.str_)


def _fold_scalar_string(value: str, *, ascii_only: bool) -> str:
    folded = excel_casefold(value)
    return folded.lower() if ascii_only else folded


def _compare_folded_strings(
    op: str,
    left_folded: np.ndarray,
    right_folded: np.ndarray,
) -> np.ndarray:
    if op == "=":
        return _bool_result_as_object(np.char.equal(left_folded, right_folded))
    if op == "<>":
        return _bool_result_as_object(np.char.not_equal(left_folded, right_folded))
    if op == "<":
        return _bool_result_as_object(np.char.less(left_folded, right_folded))
    if op == ">":
        return _bool_result_as_object(np.char.greater(left_folded, right_folded))
    if op == "<=":
        return _bool_result_as_object(np.char.less_equal(left_folded, right_folded))
    if op == ">=":
        return _bool_result_as_object(np.char.greater_equal(left_folded, right_folded))
    raise ValueError(f"Unknown comparison operator: {op}")


def _try_string_compare_fastpath(
    op: str,
    arr_left: np.ndarray,
    arr_right: np.ndarray,
) -> np.ndarray | None:
    """Vectorized casefolded string compare when both sides are plain strings."""
    left_meta = _string_array_meta(arr_left)
    if left_meta is None:
        return None
    right_meta = _string_array_meta(arr_right)
    if right_meta is None:
        return None

    if right_meta.is_constant and not left_meta.is_constant:
        left_folded = _fold_string_column(arr_left, left_meta)
        right_folded = _fold_scalar_string(
            right_meta.constant_value,
            ascii_only=right_meta.ascii_only,
        )
        return _compare_folded_strings(op, left_folded, np.asarray(right_folded))
    if left_meta.is_constant and not right_meta.is_constant:
        right_folded = _fold_string_column(arr_right, right_meta)
        left_folded = _fold_scalar_string(
            left_meta.constant_value,
            ascii_only=left_meta.ascii_only,
        )
        return _compare_folded_strings(op, np.asarray(left_folded), right_folded)

    left_folded = _fold_string_column(arr_left, left_meta)
    right_folded = _fold_string_column(arr_right, right_meta)
    return _compare_folded_strings(op, left_folded, right_folded)


_COMPARE_UFUNCS = {
    "=": operator.eq,
    "<>": operator.ne,
    "<": operator.lt,
    ">": operator.gt,
    "<=": operator.le,
    ">=": operator.ge,
}


def _numpy_compare(op: str, left: np.ndarray, right: np.ndarray) -> np.ndarray:
    compare = _COMPARE_UFUNCS[op]
    return _bool_result_as_object(compare(left, right))


def try_fastpath_compare_array(
    op: str,
    arr_left: np.ndarray,
    arr_right: np.ndarray,
) -> np.ndarray | XlError | None:
    """Apply a vectorized compare fast path, or return None to use the reference loop.

    Dispatch tiers (arrays with at least ``MIN_OPERATOR_FASTPATH_CELLS`` elements):

    1. Fail-fast scan for embedded ``XlError`` values (C-order, left wins per cell).
    2. String path when both sides are plain ``str`` cells (scalar-broadcast aware).
    3. Numeric path when both sides batch-coerce to float64 (including all-string
       numeric text columns via ``_try_batch_coerce_numeric_strings``).

    Smaller arrays and mixed-type cells fall through to the per-cell reference loop.
    """
    if arr_left.size < MIN_OPERATOR_FASTPATH_CELLS:
        return None

    error = _first_paired_error_c_order(arr_left, arr_right)
    if error is not None:
        return error

    string_result = _try_string_compare_fastpath(op, arr_left, arr_right)
    if string_result is not None:
        return string_result

    left_numeric = batch_coerce_to_float64(arr_left)
    right_numeric = batch_coerce_to_float64(arr_right)
    if left_numeric is not None and right_numeric is not None:
        return _numpy_compare(op, left_numeric, right_numeric)

    return None


def try_fastpath_arithmetic_array(
    op: str,
    arr_left: np.ndarray,
    arr_right: np.ndarray,
) -> np.ndarray | XlError | None:
    """Apply a vectorized numeric fast path, or return None to use the reference loop.

    ``+``, ``-``, ``*``, and ``/`` are fully vectorized. ``^`` delegates to a
    per-cell Python loop (not ``np.power``) to preserve reference error semantics.
    """
    if arr_left.size < MIN_OPERATOR_FASTPATH_CELLS:
        return None

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


def _detect_concat_column_kind(arr: np.ndarray) -> str | None:
    """Return a homogeneous column kind tag, or None when cells are mixed."""
    flat = arr.ravel()
    if flat.size == 0:
        return "string"

    kind: str | None = None
    for value in flat:
        if isinstance(value, str):
            current = "string"
        elif isinstance(value, bool):
            current = "bool"
        elif isinstance(value, (int, np.integer)):
            current = "integer"
        elif isinstance(value, (float, np.floating)):
            current = "float"
        elif value is None:
            current = "none"
        elif isinstance(value, XlError):
            current = "xl_error"
        else:
            return None
        if kind is None:
            kind = current
        elif kind != current:
            return None
    return kind


def _format_concat_column(arr: np.ndarray, kind: str) -> np.ndarray:
    """Format a homogeneous object column with Excel ``to_string`` rules."""
    if kind == "string":
        return _as_unicode_strings(arr)
    if kind == "integer":
        return np.asarray(np.asarray(arr, dtype=np.int64), dtype=np.str_)
    if kind == "float":
        return np.asarray(_TO_STRING_UFUNC(arr.ravel()), dtype=np.str_).reshape(arr.shape)
    if kind == "bool":
        flat = arr.ravel()
        formatted = np.empty(flat.size, dtype=np.str_)
        for index, value in enumerate(flat):
            formatted[index] = "TRUE" if value else "FALSE"
        return formatted.reshape(arr.shape)
    if kind == "none":
        return np.full(arr.shape, "", dtype=np.str_)
    if kind == "xl_error":
        flat = arr.ravel()
        formatted = np.empty(flat.size, dtype=np.str_)
        for index, value in enumerate(flat):
            formatted[index] = value.value
        return formatted.reshape(arr.shape)
    raise ValueError(f"Unknown concat column kind: {kind}")


def _concat_formatted_columns(left: np.ndarray, right: np.ndarray) -> np.ndarray:
    result = np.char.add(left, right)
    return result.astype(object)


def try_fastpath_concat_array(
    arr_left: np.ndarray,
    arr_right: np.ndarray,
) -> np.ndarray | None:
    """Apply a vectorized concat fast path, or return None to use the reference loop.

    Handles homogeneous string, integer, float, bool, null, and ``XlError`` columns by
    batch-formatting each side with ``to_string`` rules, then ``np.char.add``.
    """
    if arr_left.size < MIN_OPERATOR_FASTPATH_CELLS:
        return None

    left_kind = _detect_concat_column_kind(arr_left)
    if left_kind is None:
        return None
    right_kind = _detect_concat_column_kind(arr_right)
    if right_kind is None:
        return None

    left_formatted = _format_concat_column(arr_left, left_kind)
    right_formatted = _format_concat_column(arr_right, right_kind)

    if left_kind == "string":
        left_meta = _string_array_meta(arr_left)
        if left_meta is not None and left_meta.is_constant and right_kind != "string":
            return _concat_formatted_columns(
                np.asarray(left_meta.constant_value),
                right_formatted,
            )
    if right_kind == "string":
        right_meta = _string_array_meta(arr_right)
        if right_meta is not None and right_meta.is_constant and left_kind != "string":
            return _concat_formatted_columns(
                left_formatted,
                np.asarray(right_meta.constant_value),
            )

    return _concat_formatted_columns(left_formatted, right_formatted)


def try_fastpath_sumproduct(arrays: list[np.ndarray]) -> float | None:
    """Sum element-wise products when every array batch-coerces to float64."""
    if not arrays:
        return 0.0
    if arrays[0].size < MIN_OPERATOR_FASTPATH_CELLS:
        return None
    coerced: list[np.ndarray] = []
    for arr in arrays:
        batch = batch_coerce_to_float64(arr)
        if batch is None:
            return None
        coerced.append(batch)
    product = coerced[0]
    for arr in coerced[1:]:
        product = product * arr
    return float(np.sum(product))
