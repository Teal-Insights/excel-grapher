from __future__ import annotations

import numpy as np

from excel_grapher.core import CellValue, XlError, flatten
from excel_grapher.core.grid import Grid
from excel_grapher.core.math_funcs import (
    abs_number,
    average_cells,
    averageif_cells,
    countif_cells,
    exp_number,
    large_kth,
    max_cells,
    min_cells,
    normdist_value,
    npv_cells,
    rank_number,
    round_number,
    rounddown_number,
    stdev_cells,
    sum_cells,
)
from excel_grapher.core.operators_fastpath import (
    MIN_OPERATOR_FASTPATH_CELLS,
    try_fastpath_sumproduct,
)
from excel_grapher.core.operators_reference import reference_sumproduct_arrays
from excel_grapher.core.sumproduct import sumproduct_cells

__all__ = [
    "xl_abs",
    "xl_exp",
    "xl_average",
    "xl_averageif",
    "xl_count",
    "xl_counta",
    "xl_countif",
    "xl_large",
    "xl_max",
    "xl_min",
    "xl_normdist",
    "xl_npv",
    "xl_rank",
    "xl_round",
    "xl_rounddown",
    "xl_stdev",
    "xl_sum",
    "xl_sumproduct",
]


def xl_sum(*args: CellValue) -> float | XlError:
    return sum_cells(*args)


def xl_average(*args: CellValue) -> float | XlError:
    return average_cells(*args)


def xl_min(*args: CellValue) -> float | XlError:
    return min_cells(*args)


def xl_max(*args: CellValue) -> float | XlError:
    return max_cells(*args)


def xl_count(*args: CellValue) -> int:
    count = 0
    for v in flatten(*args):
        if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
            count += 1
    return count


def xl_counta(*args: CellValue) -> int:
    count = 0
    for v in flatten(*args):
        if v is not None and v != "":
            count += 1
    return count


def xl_round(number: CellValue, num_digits: CellValue) -> float | XlError:
    return round_number(number, num_digits)


def xl_rounddown(number: CellValue, num_digits: CellValue) -> float | XlError:
    return rounddown_number(number, num_digits)


def xl_npv(rate: CellValue, *values: CellValue) -> float | XlError:
    return npv_cells(rate, *values)


def xl_stdev(*args: CellValue) -> float | XlError:
    return stdev_cells(*args)


def xl_countif(range_values: CellValue, criteria: CellValue) -> int | XlError:
    return countif_cells(range_values, criteria)


def xl_averageif(
    range_values: CellValue,
    criteria: CellValue,
    average_range: CellValue | None = None,
) -> float | XlError:
    return averageif_cells(range_values, criteria, average_range)


def xl_large(array: CellValue, k: CellValue) -> float | XlError:
    return large_kth(array, k)


def xl_rank(number: CellValue, ref: CellValue, order: CellValue = 0) -> int | XlError:
    return rank_number(number, ref, order)


def xl_normdist(
    x: CellValue,
    mean: CellValue,
    standard_dev: CellValue,
    cumulative: CellValue,
) -> float | XlError:
    return normdist_value(x, mean, standard_dev, cumulative)


def xl_abs(*args: CellValue) -> float | XlError:
    return abs_number(*args)


def xl_exp(*args: CellValue) -> float | XlError:
    return exp_number(*args)


def xl_sumproduct(*args: CellValue) -> float | XlError:
    """SUMPRODUCT with optional NumPy acceleration for large fully-consumed grids."""
    if len(args) == 0:
        return 0.0

    grids: list[Grid] = []
    for arg in args:
        grid = Grid.wrap(arg)
        if grid is None:
            scalar = Grid.wrap([[arg]])
            assert scalar is not None
            grid = scalar
        grids.append(grid)

    shape = (grids[0].nrows, grids[0].ncols)
    for grid in grids[1:]:
        if (grid.nrows, grid.ncols) != shape:
            return XlError.VALUE

    if grids[0].size >= MIN_OPERATOR_FASTPATH_CELLS:
        arrays = [
            np.array(
                [[grid.at(row0, col0) for col0 in range(grid.ncols)] for row0 in range(grid.nrows)],
                dtype=object,
            )
            for grid in grids
        ]
        fast = try_fastpath_sumproduct(arrays)
        if fast is not None:
            return fast
        return reference_sumproduct_arrays(arrays)

    return sumproduct_cells(*args)
