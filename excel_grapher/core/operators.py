"""Excel-style scalar and array operators (representation-agnostic).

Array operands try vectorized fast paths in ``operators_fastpath`` first when the
broadcast array has at least ``MIN_OPERATOR_FASTPATH_CELLS`` (64) elements; smaller
arrays and failed guards use per-cell reference loops in ``operators_reference``.
Exponent (``^``) stays on a per-cell loop inside the fast path for parity.

Lazy ``Range`` / nested-list operands use shared ``operator_maps`` (same loops as
export ``xl_map_*``). Large grids materialize once at this boundary and try
``operators_fastpath``; on a miss the same arrays feed ``reference_*_array``
rather than walking the ``Range`` a second time.
"""

from __future__ import annotations

from typing import cast

import numpy as np

from .coercions import to_number
from .grid import Grid, Range
from .operator_maps import map_arithmetic, map_compare, map_concat, map_unary
from .operators_fastpath import (
    MIN_OPERATOR_FASTPATH_CELLS,
    try_fastpath_arithmetic_array,
    try_fastpath_compare_array,
    try_fastpath_concat_array,
)
from .operators_reference import (
    apply_arithmetic,
    broadcast_pair,
    compare_scalars,
    concat_scalars,
    reference_arithmetic_array,
    reference_compare_array,
    reference_concat_array,
)
from .types import CellValue, XlError


def _broadcast_pair(
    left: CellValue,
    right: CellValue,
) -> tuple[np.ndarray, np.ndarray] | XlError:
    return broadcast_pair(left, right)


def _is_range_or_list(value: object) -> bool:
    return isinstance(value, (Range, list, tuple))


def _materialize_grid(grid: Grid) -> np.ndarray:
    return np.array(
        [[grid.at(row0, col0) for col0 in range(grid.ncols)] for row0 in range(grid.nrows)],
        dtype=object,
    )


def _apply_large_grid_binary(
    op: str | None,
    left: object,
    right: object,
    *,
    kind: str,
) -> CellValue | None:
    """Materialize large grids once, then run fastpath or reference loops.

    Returns `None` when operands are below ``MIN_OPERATOR_FASTPATH_CELLS`` (or
    cannot form a grid pair) so the caller can use shared ``map_*`` loops.
    When this path materializes, a fastpath miss reuses the same arrays via
    ``reference_*_array`` — it does not fall through to a second Range walk.
    """
    left_grid = Grid.wrap(left)
    right_grid = Grid.wrap(right)
    if left_grid is None and right_grid is None:
        return None
    if left_grid is not None and right_grid is not None:
        if (left_grid.nrows, left_grid.ncols) != (right_grid.nrows, right_grid.ncols):
            return None
        if left_grid.size < MIN_OPERATOR_FASTPATH_CELLS:
            return None
        arr_left = _materialize_grid(left_grid)
        arr_right = _materialize_grid(right_grid)
    elif left_grid is not None:
        if left_grid.size < MIN_OPERATOR_FASTPATH_CELLS:
            return None
        arr_left = _materialize_grid(left_grid)
        arr_right = np.full(arr_left.shape, right, dtype=object)
    else:
        assert right_grid is not None
        if right_grid.size < MIN_OPERATOR_FASTPATH_CELLS:
            return None
        arr_right = _materialize_grid(right_grid)
        arr_left = np.full(arr_right.shape, left, dtype=object)

    if kind == "compare":
        assert op is not None
        fast = try_fastpath_compare_array(op, arr_left, arr_right)
        if fast is not None:
            return fast
        return reference_compare_array(op, arr_left, arr_right)
    if kind == "concat":
        fast = try_fastpath_concat_array(arr_left, arr_right)
        if fast is not None:
            return fast
        return reference_concat_array(arr_left, arr_right)
    assert op is not None
    fast = try_fastpath_arithmetic_array(op, arr_left, arr_right)
    if fast is not None:
        return fast
    return reference_arithmetic_array(op, arr_left, arr_right)


def _compare_scalars(op: str, left: CellValue, right: CellValue) -> bool | XlError:
    return compare_scalars(op, left, right)


def _xl_compare(op: str, left: CellValue, right: CellValue) -> CellValue:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right

    if _is_range_or_list(left) or _is_range_or_list(right):
        large = _apply_large_grid_binary(op, left, right, kind="compare")
        if large is not None:
            return large
        return cast(CellValue, map_compare(op, left, right))

    if isinstance(left, np.ndarray) or isinstance(right, np.ndarray):
        pair = _broadcast_pair(left, right)
        if isinstance(pair, XlError):
            return pair
        arr_left, arr_right = pair
        fast = try_fastpath_compare_array(op, arr_left, arr_right)
        if fast is not None:
            return fast
        return reference_compare_array(op, arr_left, arr_right)

    return _compare_scalars(op, left, right)


def _xl_arithmetic(
    op: str,
    left: CellValue,
    right: CellValue,
) -> CellValue:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right

    if _is_range_or_list(left) or _is_range_or_list(right):
        large = _apply_large_grid_binary(op, left, right, kind="arithmetic")
        if large is not None:
            return large
        return cast(CellValue, map_arithmetic(op, left, right))

    if isinstance(left, np.ndarray) or isinstance(right, np.ndarray):
        pair = _broadcast_pair(left, right)
        if isinstance(pair, XlError):
            return pair
        arr_left, arr_right = pair
        fast = try_fastpath_arithmetic_array(op, arr_left, arr_right)
        if fast is not None:
            return fast
        return reference_arithmetic_array(op, arr_left, arr_right)

    ln = to_number(left)
    rn = to_number(right)
    if isinstance(ln, XlError):
        return ln
    if isinstance(rn, XlError):
        return rn
    return apply_arithmetic(op, ln, rn)


def _concat_scalars(left: CellValue, right: CellValue) -> str:
    return concat_scalars(left, right)


def _xl_concat(left: CellValue, right: CellValue) -> CellValue:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right

    if _is_range_or_list(left) or _is_range_or_list(right):
        large = _apply_large_grid_binary(None, left, right, kind="concat")
        if large is not None:
            return large
        return cast(CellValue, map_concat(left, right))

    if isinstance(left, np.ndarray) or isinstance(right, np.ndarray):
        pair = _broadcast_pair(left, right)
        if isinstance(pair, XlError):
            return pair
        arr_left, arr_right = pair
        fast = try_fastpath_concat_array(arr_left, arr_right)
        if fast is not None:
            return fast
        return reference_concat_array(arr_left, arr_right)

    return _concat_scalars(left, right)


def xl_concat(left: CellValue, right: CellValue) -> CellValue:
    return _xl_concat(left, right)


def xl_eq(left: CellValue, right: CellValue) -> CellValue:
    return _xl_compare("=", left, right)


def xl_ne(left: CellValue, right: CellValue) -> CellValue:
    return _xl_compare("<>", left, right)


def xl_lt(left: CellValue, right: CellValue) -> CellValue:
    return _xl_compare("<", left, right)


def xl_gt(left: CellValue, right: CellValue) -> CellValue:
    return _xl_compare(">", left, right)


def xl_le(left: CellValue, right: CellValue) -> CellValue:
    return _xl_compare("<=", left, right)


def xl_ge(left: CellValue, right: CellValue) -> CellValue:
    return _xl_compare(">=", left, right)


def xl_iferror(value: CellValue, value_if_error: CellValue) -> CellValue:
    if isinstance(value, XlError):
        return value_if_error
    return value


def xl_div(left: CellValue, right: CellValue) -> CellValue:
    return _xl_arithmetic("/", left, right)


def xl_add(left: CellValue, right: CellValue) -> CellValue:
    return _xl_arithmetic("+", left, right)


def xl_sub(left: CellValue, right: CellValue) -> CellValue:
    return _xl_arithmetic("-", left, right)


def xl_mul(left: CellValue, right: CellValue) -> CellValue:
    return _xl_arithmetic("*", left, right)


def xl_pow(left: CellValue, right: CellValue) -> CellValue:
    return _xl_arithmetic("^", left, right)


def xl_neg(value: CellValue) -> CellValue:
    return cast(CellValue, map_unary("-", value))


def xl_pos(value: CellValue) -> CellValue:
    return cast(CellValue, map_unary("+", value))


def xl_percent(value: CellValue) -> CellValue:
    """Excel postfix percent operator (%): divide a numeric value by 100."""
    return cast(CellValue, map_unary("%", value))
