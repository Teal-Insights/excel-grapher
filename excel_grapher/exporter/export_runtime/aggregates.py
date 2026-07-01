"""Range reductions over lazy ranges for exported code."""

from __future__ import annotations

from excel_grapher.core import XlError, to_number

from .values import CellValue, Grid

__all__ = ["xl_sumproduct"]


def xl_sumproduct(*args: CellValue) -> float | XlError:
    if len(args) == 0:
        return 0.0
    grids: list[Grid] = []
    for arg in args:
        grid = Grid.wrap(arg)
        if grid is None:
            scalar_grid = Grid.wrap([[arg]])
            assert scalar_grid is not None
            grid = scalar_grid
        grids.append(grid)
    shape = (grids[0].nrows, grids[0].ncols)
    for grid in grids[1:]:
        if (grid.nrows, grid.ncols) != shape:
            return XlError.VALUE

    result = 0.0
    for index0 in range(grids[0].size):
        product = 1.0
        for grid in grids:
            number = to_number(grid.at_flat(index0))
            if isinstance(number, XlError):
                return number
            product *= number
        result += product
    return result
