"""SUMPRODUCT over aligned Grid / nested-list / scalar arguments.

Shared Grid traversal for evaluator and export. Optional NumPy acceleration
lives in `runtime.math.xl_sumproduct` (evaluator only) so exported runtimes
stay NumPy-free.
"""

from __future__ import annotations

from typing import cast

from .coercions import to_number
from .grid import Grid
from .types import CellValue, XlError


def _as_grid(arg: CellValue) -> Grid:
    grid = Grid.wrap(arg)
    if grid is not None:
        return grid
    scalar = Grid.wrap([[arg]])
    assert scalar is not None
    return scalar


def sumproduct_cells(*args: CellValue) -> float | XlError:
    """Return the sum of element-wise products across aligned array arguments."""
    if len(args) == 0:
        return 0.0
    grids = [_as_grid(arg) for arg in args]
    shape = (grids[0].nrows, grids[0].ncols)
    for grid in grids[1:]:
        if (grid.nrows, grid.ncols) != shape:
            return XlError.VALUE

    result = 0.0
    for index0 in range(grids[0].size):
        product = 1.0
        for grid in grids:
            cell = cast(CellValue, grid.at_flat(index0))
            # Excel SUMPRODUCT treats non-numeric text as 0 (not `#VALUE!`).
            # `XlError` is a `StrEnum` — check errors before the `str` branch.
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
