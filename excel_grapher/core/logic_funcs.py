"""Shared logical worksheet function implementations."""

from __future__ import annotations

from collections.abc import Iterator
from typing import cast

from .coercions import to_bool
from .grid import Grid
from .types import CellValue, XlError

__all__ = ["logical_and", "logical_if", "logical_not", "logical_or"]


def _iter_logical_cells(arg: CellValue) -> Iterator[CellValue]:
    """Yield scalar cells from a range/array argument in row-major order."""
    grid = Grid.wrap(arg)
    if grid is not None:
        for cell in grid.iter_raw():
            yield cast(CellValue, cell)
        return
    yield arg


def _consume_logical_cells(
    args: tuple[CellValue, ...],
    *,
    short_circuit_true: bool,
) -> bool | XlError:
    """Walk logical cells across arguments with Excel AND/OR semantics."""
    found_logical = False
    for arg in args:
        for cell in _iter_logical_cells(arg):
            if cell is None:
                continue
            if isinstance(cell, XlError):
                return cell
            found_logical = True
            b = to_bool(cell)
            if isinstance(b, XlError):
                return b
            if short_circuit_true:
                if b:
                    return True
            elif not b:
                return False
    if not found_logical:
        return XlError.VALUE
    return not short_circuit_true


def logical_and(*args: CellValue) -> bool | XlError:
    """Return logical AND across scalar and range arguments."""
    return _consume_logical_cells(args, short_circuit_true=False)


def logical_or(*args: CellValue) -> bool | XlError:
    """Return logical OR across scalar and range arguments."""
    return _consume_logical_cells(args, short_circuit_true=True)


def logical_not(arg: CellValue) -> bool | XlError:
    """Return logical NOT of an argument."""
    b = to_bool(arg)
    if isinstance(b, XlError):
        return b
    return not b


def logical_if(
    cond: object,
    then_value: object,
    else_value: object = False,
) -> object:
    """Return `then_value` or `else_value` under Excel `IF` semantics.

    A scalar condition picks one branch. When any operand is a range or nested
    array, selection is element-wise: scalars broadcast, and mixed array shapes
    return `#VALUE!`. An omitted else is `FALSE`.
    """
    cond_grid = Grid.wrap(cond)
    then_grid = Grid.wrap(then_value)
    else_grid = Grid.wrap(else_value)
    if cond_grid is None and then_grid is None and else_grid is None:
        flag = to_bool(cast(CellValue, cond))
        if isinstance(flag, XlError):
            return flag
        return then_value if flag else else_value

    grids = [grid for grid in (cond_grid, then_grid, else_grid) if grid is not None]
    nrows, ncols = grids[0].nrows, grids[0].ncols
    if any(grid.nrows != nrows or grid.ncols != ncols for grid in grids[1:]):
        return XlError.VALUE

    def _at(grid: Grid | None, scalar: object, row: int, col: int) -> object:
        if grid is None:
            return scalar
        return grid.at(row, col)

    result: list[list[CellValue]] = []
    for row in range(nrows):
        out_row: list[CellValue] = []
        for col in range(ncols):
            flag = to_bool(cast(CellValue, _at(cond_grid, cond, row, col)))
            if isinstance(flag, XlError):
                out_row.append(flag)
            elif flag:
                out_row.append(cast(CellValue, _at(then_grid, then_value, row, col)))
            else:
                out_row.append(cast(CellValue, _at(else_grid, else_value, row, col)))
        result.append(out_row)
    return result
