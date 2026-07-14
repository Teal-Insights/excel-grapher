"""Shared element-wise operator maps over lazy `Grid` values.

Implementations return `XlError` sentinels. Evaluator operators call these
directly; `export_runtime.operators` wrappers raise `XlErrorException`.

Evaluator and export value vocabularies differ slightly (export `CellValue`
also includes nested lists); shared maps accept opaque objects and narrow at
use sites — same pattern as `lookup_funcs`.
"""

from __future__ import annotations

from typing import cast

from excel_grapher.core.coercions import to_number
from excel_grapher.core.grid import Grid, Scalar
from excel_grapher.core.operators_reference import (
    apply_arithmetic,
    compare_scalars,
    concat_scalars,
)
from excel_grapher.core.types import CellValue, XlError

__all__ = [
    "map_arithmetic",
    "map_compare",
    "map_concat",
    "map_unary",
]


def _broadcast_grids(left: object, right: object) -> tuple[Grid, Grid] | None | XlError:
    """Wrap array operands as aligned grids; `None` when both operands are scalar.

    Named distinctly from ``operators._broadcast_pair`` so flattened export
    embeddings do not collide on the private helper name.
    """
    left_grid = Grid.wrap(left)
    right_grid = Grid.wrap(right)
    if left_grid is None and right_grid is None:
        return None
    if left_grid is not None and right_grid is not None:
        if (left_grid.nrows, left_grid.ncols) != (right_grid.nrows, right_grid.ncols):
            return XlError.VALUE
        return left_grid, right_grid
    if left_grid is not None:
        scalar_right = Grid.wrap([[right] * left_grid.ncols for _ in range(left_grid.nrows)])
        assert scalar_right is not None
        return left_grid, scalar_right
    assert right_grid is not None
    scalar_left = Grid.wrap([[left] * right_grid.ncols for _ in range(right_grid.nrows)])
    assert scalar_left is not None
    return scalar_left, right_grid


def _apply_arithmetic_cell(op: str, left: Scalar, right: Scalar) -> CellValue:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right
    ln = to_number(cast(CellValue, left))
    rn = to_number(cast(CellValue, right))
    if isinstance(ln, XlError):
        return ln
    if isinstance(rn, XlError):
        return rn
    return apply_arithmetic(op, ln, rn)


def map_arithmetic(op: str, left: object, right: object) -> object:
    """Element-wise arithmetic over scalar or broadcast array operands."""
    pair = _broadcast_grids(left, right)
    if isinstance(pair, XlError):
        return pair
    if pair is None:
        return _apply_arithmetic_cell(op, cast(Scalar, left), cast(Scalar, right))

    arr_left, arr_right = pair
    result: list[list[CellValue]] = []
    for row0 in range(arr_left.nrows):
        out_row: list[CellValue] = []
        for col0 in range(arr_left.ncols):
            cell = _apply_arithmetic_cell(op, arr_left.at(row0, col0), arr_right.at(row0, col0))
            if isinstance(cell, XlError):
                return cell
            out_row.append(cell)
        result.append(out_row)
    return result


def map_compare(op: str, left: object, right: object) -> object:
    """Element-wise comparison over scalar or broadcast array operands."""
    pair = _broadcast_grids(left, right)
    if isinstance(pair, XlError):
        return pair
    if pair is None:
        return compare_scalars(op, cast(CellValue, left), cast(CellValue, right))

    arr_left, arr_right = pair
    result: list[list[CellValue]] = []
    for row0 in range(arr_left.nrows):
        out_row: list[CellValue] = []
        for col0 in range(arr_left.ncols):
            cell = compare_scalars(
                op,
                cast(CellValue, arr_left.at(row0, col0)),
                cast(CellValue, arr_right.at(row0, col0)),
            )
            if isinstance(cell, XlError):
                return cell
            out_row.append(cell)
        result.append(out_row)
    return result


def map_concat(left: object, right: object) -> object:
    """Element-wise string concatenation over scalar or broadcast array operands."""
    pair = _broadcast_grids(left, right)
    if isinstance(pair, XlError):
        return pair
    if pair is None:
        if isinstance(left, XlError):
            return left
        if isinstance(right, XlError):
            return right
        return concat_scalars(cast(CellValue, left), cast(CellValue, right))

    arr_left, arr_right = pair
    result: list[list[CellValue]] = []
    for row0 in range(arr_left.nrows):
        out_row: list[CellValue] = []
        for col0 in range(arr_left.ncols):
            lv = arr_left.at(row0, col0)
            rv = arr_right.at(row0, col0)
            if isinstance(lv, XlError):
                return lv
            if isinstance(rv, XlError):
                return rv
            out_row.append(concat_scalars(cast(CellValue, lv), cast(CellValue, rv)))
        result.append(out_row)
    return result


def map_unary(op: str, value: object) -> object:
    """Apply a unary Excel operator over scalar or array operands."""
    grid = Grid.wrap(value)
    if grid is None:
        if isinstance(value, XlError):
            return value
        number = to_number(cast(CellValue, value))
        if isinstance(number, XlError):
            return number
        if op == "-":
            return -number
        if op == "+":
            return +number
        if op == "%":
            return number / 100.0
        raise ValueError(f"Unknown unary operator: {op}")

    result: list[list[CellValue]] = []
    for row0 in range(grid.nrows):
        out_row: list[CellValue] = []
        for col0 in range(grid.ncols):
            cell = grid.at(row0, col0)
            if isinstance(cell, XlError):
                return cell
            number = to_number(cast(CellValue, cell))
            if isinstance(number, XlError):
                return number
            if op == "-":
                out_row.append(-number)
            elif op == "+":
                out_row.append(+number)
            elif op == "%":
                out_row.append(number / 100.0)
            else:
                raise ValueError(f"Unknown unary operator: {op}")
        result.append(out_row)
    return result
