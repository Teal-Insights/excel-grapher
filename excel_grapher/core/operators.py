"""Excel-style scalar and array operators (representation-agnostic).

Array operands try vectorized fast paths in ``operators_fastpath`` first when the
broadcast array has at least ``MIN_OPERATOR_FASTPATH_CELLS`` (64) elements; smaller
arrays and failed guards use per-cell reference loops in ``operators_reference``.
Exponent (``^``) stays on a per-cell loop inside the fast path for parity.

Lazy ``Range`` / nested-list operands use shared ``operator_maps`` (same loops as
export ``xl_map_*``). Large grids materialize once at this boundary and try
``operators_fastpath``; on a miss the same arrays feed ``reference_*_array``
rather than walking the ``Range`` a second time.

When the optional ``fast`` extra (NumPy) is not installed, large-grid
materialization is skipped and callers fall through to ``operator_maps``.
"""

from __future__ import annotations

from typing import Any, cast

from .coercions import to_number
from .grid import Grid
from .grid.grid import _as_nested_rows_from_ndarray
from .operator_maps import map_arithmetic, map_compare, map_concat, map_unary
from .operator_thresholds import MIN_OPERATOR_FASTPATH_CELLS
from .operators_reference import (
    apply_arithmetic,
    compare_scalars,
    concat_scalars,
    reference_arithmetic_array,
    reference_compare_array,
    reference_concat_array,
)
from .types import CellValue, FormulaValue, XlError

try:
    import numpy as np
except ImportError:  # pragma: no cover - exercised when the `fast` extra is absent
    np = None  # type: ignore[assignment]

if np is not None:
    from .operators_fastpath import (
        try_fastpath_arithmetic_array,
        try_fastpath_compare_array,
        try_fastpath_concat_array,
    )
else:  # pragma: no cover - exercised when the `fast` extra is absent
    from .operators_fastpath_stub import (
        try_fastpath_arithmetic_array,
        try_fastpath_compare_array,
        try_fastpath_concat_array,
    )


def _is_grid_operand(value: object) -> bool:
    return Grid.wrap(value) is not None


def _as_cell_result(value: object) -> FormulaValue:
    """Normalize operator results to ``FormulaValue`` (nested lists, not ndarrays)."""
    rows = _as_nested_rows_from_ndarray(value)
    if rows is not None:
        return rows
    return cast(FormulaValue, value)


def _materialize_grid(grid: Grid) -> Any:
    assert np is not None
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
) -> FormulaValue | None:
    """Materialize large grids once, then run fastpath or reference loops.

    Returns `None` when operands are below ``MIN_OPERATOR_FASTPATH_CELLS`` (or
    cannot form a grid pair) so the caller can use shared ``map_*`` loops.
    When this path materializes, a fastpath miss reuses the same arrays via
    ``reference_*_array`` — it does not fall through to a second Range walk.

    Without NumPy (`fast` extra), always returns `None` so callers use
    ``operator_maps``.
    """
    if np is None:
        return None

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
            return _as_cell_result(fast)
        return _as_cell_result(reference_compare_array(op, arr_left, arr_right))
    if kind == "concat":
        fast = try_fastpath_concat_array(arr_left, arr_right)
        if fast is not None:
            return _as_cell_result(fast)
        return _as_cell_result(reference_concat_array(arr_left, arr_right))
    assert op is not None
    fast = try_fastpath_arithmetic_array(op, arr_left, arr_right)
    if fast is not None:
        return _as_cell_result(fast)
    return _as_cell_result(reference_arithmetic_array(op, arr_left, arr_right))


def _compare_scalars(op: str, left: FormulaValue, right: FormulaValue) -> bool | XlError:
    return compare_scalars(op, left, right)


def _xl_compare(op: str, left: FormulaValue, right: FormulaValue) -> FormulaValue:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right

    if _is_grid_operand(left) or _is_grid_operand(right):
        large = _apply_large_grid_binary(op, left, right, kind="compare")
        if large is not None:
            return large
        return cast(FormulaValue, map_compare(op, left, right))

    return _compare_scalars(op, left, right)


def _xl_arithmetic(
    op: str,
    left: FormulaValue,
    right: FormulaValue,
) -> FormulaValue:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right

    if _is_grid_operand(left) or _is_grid_operand(right):
        large = _apply_large_grid_binary(op, left, right, kind="arithmetic")
        if large is not None:
            return large
        return cast(FormulaValue, map_arithmetic(op, left, right))

    ln = to_number(left)
    rn = to_number(right)
    if isinstance(ln, XlError):
        return ln
    if isinstance(rn, XlError):
        return rn
    return apply_arithmetic(op, ln, rn)


def _concat_scalars(left: CellValue, right: CellValue) -> str:
    return concat_scalars(left, right)


def _xl_concat(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right

    if _is_grid_operand(left) or _is_grid_operand(right):
        large = _apply_large_grid_binary(None, left, right, kind="concat")
        if large is not None:
            return large
        return cast(FormulaValue, map_concat(left, right))

    return _concat_scalars(cast(CellValue, left), cast(CellValue, right))


def xl_concat(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_concat(left, right)


def xl_eq(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_compare("=", left, right)


def xl_ne(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_compare("<>", left, right)


def xl_lt(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_compare("<", left, right)


def xl_gt(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_compare(">", left, right)


def xl_le(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_compare("<=", left, right)


def xl_ge(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_compare(">=", left, right)


def xl_iferror(value: CellValue, value_if_error: CellValue) -> CellValue:
    if isinstance(value, XlError):
        return value_if_error
    return value


def xl_div(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_arithmetic("/", left, right)


def xl_add(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_arithmetic("+", left, right)


def xl_sub(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_arithmetic("-", left, right)


def xl_mul(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_arithmetic("*", left, right)


def xl_pow(left: FormulaValue, right: FormulaValue) -> FormulaValue:
    return _xl_arithmetic("^", left, right)


def xl_neg(value: FormulaValue) -> FormulaValue:
    return cast(FormulaValue, map_unary("-", value))


def xl_pos(value: FormulaValue) -> FormulaValue:
    return cast(FormulaValue, map_unary("+", value))


def xl_percent(value: FormulaValue) -> FormulaValue:
    """Excel postfix percent operator (%): divide a numeric value by 100."""
    return cast(FormulaValue, map_unary("%", value))
