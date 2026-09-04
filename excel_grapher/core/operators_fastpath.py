"""Vectorized fast paths for Excel binary operators and SUMPRODUCT reduction.

``+``, ``-``, ``*``, and ``/`` use NumPy element-wise ops on coerced float64 arrays.
The ``^`` exponent operator is not vectorized: it runs a C-order Python loop so
per-element ``#NUM!`` embedding matches the reference implementation.
"""

from __future__ import annotations

import operator
from dataclasses import dataclass

import numpy as np

from .coercions import excel_casefold, to_number, to_string, try_coerce_string_to_float
from .operator_thresholds import MIN_OPERATOR_FASTPATH_CELLS
from .types import XlError

_NUMERIC_CELL_TYPES = (int, float, np.integer, np.floating)
_DIRECT_FLOAT_CELL_TYPES = frozenset({int, float, bool})
_CASEFOLD_UFUNC = np.frompyfunc(excel_casefold, 1, 1)
_TO_STRING_UFUNC = np.frompyfunc(to_string, 1, 1)


def _object_cell_types(arr: np.ndarray) -> frozenset[type] | None:
    """Return the distinct cell types of an object ndarray, `None` for typed dtypes.

    Each dispatch tier needs one yes/no fact about an operand (does it hold an
    error, is it all text, is it all numeric). One C-level pass answers all of
    them, replacing a full Python scan per tier. `None` means the operand was
    not classified, so callers keep their own per-cell checks.
    """
    if arr.dtype != object:
        return None
    return frozenset(map(type, arr.ravel()))


def _may_hold_error(cell_types: frozenset[type] | None) -> bool:
    """Whether an operand can contain an embedded ``XlError``."""
    return cell_types is None or any(issubclass(kind, XlError) for kind in cell_types)


def _is_number_type(kind: type) -> bool:
    """True for int/float (including NumPy), excluding `bool`."""
    if issubclass(kind, bool):
        return False
    return issubclass(kind, (int, float))


def _all_numbers_or_blank(cell_types: frozenset[type] | None) -> bool:
    """True when every cell is a number or blank (`None`).

    Comparison does not coerce text or logicals to numbers (Excel type-rank).
    """
    if cell_types is None:
        return False
    return all(kind is type(None) or _is_number_type(kind) for kind in cell_types)


def _as_compare_float64(arr: np.ndarray) -> np.ndarray | None:
    """Map number/blank cells to float64 (`None` → `0`) without cross-type coerce."""
    direct = _try_asarray_float64(arr)
    if direct is not None:
        return direct
    flat = arr.ravel()
    out = np.empty(flat.size, dtype=np.float64)
    for index, value in enumerate(flat):
        if value is None:
            out[index] = 0.0
            continue
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            return None
        out[index] = float(value)
    return out.reshape(arr.shape)


def _all_plain_strings(cell_types: frozenset[type] | None) -> bool:
    """Whether every cell is text; ``XlError`` is a ``str`` subclass but not text."""
    if cell_types is None:
        return True
    return all(issubclass(kind, str) and not issubclass(kind, XlError) for kind in cell_types)


def _holds_plain_text(arr: np.ndarray, cell_types: frozenset[type] | None) -> bool:
    """Whether any cell is text, falling back to a scan for unclassified operands."""
    if cell_types is None:
        return any(
            isinstance(value, str) and not isinstance(value, XlError) for value in arr.ravel()
        )
    return any(issubclass(kind, str) and not issubclass(kind, XlError) for kind in cell_types)


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


def _try_map_float_strings(arr: np.ndarray) -> np.ndarray | None:
    """Parse an all-text operand with one C-level ``float`` map.

    ``float`` already ignores surrounding whitespace, so this accepts exactly
    what `try_coerce_string_to_float` accepts. Text it rejects (blank cells, ISO
    dates) returns `None` so the per-cell tiers can apply the remaining Excel
    coercions.
    """
    flat = arr.ravel()
    try:
        values = np.fromiter(map(float, flat), dtype=np.float64, count=flat.size)
    except (TypeError, ValueError):
        return None
    return values.reshape(arr.shape)


def batch_coerce_to_float64(
    arr: np.ndarray,
    cell_types: frozenset[type] | None = None,
) -> np.ndarray | None:
    """Coerce an object ndarray to float64, or return None when any cell fails.

    Args:
        arr: Object (or already numeric) ndarray of Excel cell values.
        cell_types: Cell types from `_object_cell_types` when the caller has
            already classified `arr`; omit to classify here.

    Returns:
        A float64 array, or `None` when any cell cannot coerce.
    """
    if cell_types is None:
        cell_types = _object_cell_types(arr)
    if cell_types is not None:
        if cell_types <= _DIRECT_FLOAT_CELL_TYPES:
            return np.asarray(arr, dtype=np.float64)
        if _all_plain_strings(cell_types):
            mapped = _try_map_float_strings(arr)
            if mapped is not None:
                return mapped

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


def _fastpath_pow(left: np.ndarray, right: np.ndarray) -> np.ndarray:
    """Element-wise power on pre-coerced float arrays.

    Uses Python ``**`` per cell rather than ``np.power`` so complex results and
    exceptions align with ``operators_reference``. Invalid cells embed ``#NUM!``.
    """
    flat_left = left.ravel()
    flat_right = right.ravel()
    out = np.empty(flat_left.size, dtype=object)
    for index in range(flat_left.size):
        base = float(flat_left[index])
        exponent = float(flat_right[index])
        try:
            value = base**exponent
        except (ValueError, OverflowError):
            out[index] = XlError.NUM
            continue
        if isinstance(value, complex):
            out[index] = XlError.NUM
            continue
        out[index] = value
    return out.reshape(left.shape)


def _fastpath_divide(left: np.ndarray, right: np.ndarray) -> np.ndarray:
    """Element-wise division embedding ``#DIV/0!`` where the divisor is zero."""
    zero_mask = right == 0.0
    out = np.empty(left.shape, dtype=object)
    with np.errstate(divide="ignore", invalid="ignore"):
        quotients = left / right
    out[~zero_mask] = quotients[~zero_mask]
    if bool(np.any(zero_mask)):
        out[zero_mask] = XlError.DIV
    return out


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

    Both operands are classified once with `_object_cell_types`; each tier below
    is entered only when that classification says it can apply.

    Dispatch tiers (arrays with at least ``MIN_OPERATOR_FASTPATH_CELLS`` elements):

    1. Fail-fast scan for embedded ``XlError`` values (C-order, left wins per cell).
    2. String path when both sides are plain ``str`` cells (scalar-broadcast aware).
    3. Numeric path when both sides are numbers or blanks (no text/logical coerce).

    Smaller arrays and mixed-type cells fall through to the per-cell reference loop.
    """
    if arr_left.size < MIN_OPERATOR_FASTPATH_CELLS:
        return None

    left_types = _object_cell_types(arr_left)
    right_types = _object_cell_types(arr_right)

    if _may_hold_error(left_types) or _may_hold_error(right_types):
        error = _first_paired_error_c_order(arr_left, arr_right)
        if error is not None:
            return error

    if _all_plain_strings(left_types) and _all_plain_strings(right_types):
        string_result = _try_string_compare_fastpath(op, arr_left, arr_right)
        if string_result is not None:
            return string_result

    if _all_numbers_or_blank(left_types) and _all_numbers_or_blank(right_types):
        left_numeric = _as_compare_float64(arr_left)
        right_numeric = _as_compare_float64(arr_right)
        if left_numeric is None or right_numeric is None:
            return None
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
        return _fastpath_divide(left, right)
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
    """Sum element-wise products when every array batch-coerces to float64.

    Arrays containing plain text fall back to the reference path so SUMPRODUCT
    can treat text as 0 (Excel). `batch_coerce_to_float64` would otherwise
    parse numeric strings, which diverges from range-text semantics.
    """
    if not arrays:
        return 0.0
    if arrays[0].size < MIN_OPERATOR_FASTPATH_CELLS:
        return None
    coerced: list[np.ndarray] = []
    for arr in arrays:
        # `XlError` is a `StrEnum`; plain text must not use numeric coerce.
        cell_types = _object_cell_types(arr)
        if _holds_plain_text(arr, cell_types):
            return None
        batch = batch_coerce_to_float64(arr, cell_types)
        if batch is None:
            return None
        coerced.append(batch)
    product = coerced[0]
    for arr in coerced[1:]:
        product = product * arr
    return float(np.sum(product))
