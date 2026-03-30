from __future__ import annotations

import numpy as np

from ..export_runtime.lookup import (
    xl__xlfn_xlookup as _rt_xl__xlfn_xlookup,
)
from ..export_runtime.lookup import (
    xl_hlookup as _rt_xl_hlookup,
)
from ..export_runtime.lookup import (
    xl_index as _rt_xl_index,
)
from ..export_runtime.lookup import (
    xl_lookup as _rt_xl_lookup,
)
from ..export_runtime.lookup import (
    xl_match as _rt_xl_match,
)
from ..export_runtime.lookup import (
    xl_vlookup as _rt_xl_vlookup,
)
from ..types import CellValue, XlError
from . import register


@register("INDEX")
def xl_index(array: np.ndarray, row_num: CellValue, col_num: CellValue = None) -> CellValue:
    """Return a value from a specific position in an array."""
    return _rt_xl_index(array, row_num, col_num)


@register("MATCH")
def xl_match(
    lookup_value: CellValue,
    lookup_array: CellValue,
    match_type: CellValue = 1,
) -> int | XlError:
    """Find the position of a value in a range."""
    return _rt_xl_match(lookup_value, lookup_array, match_type)


@register("LOOKUP")
def xl_lookup(
    lookup_value: CellValue,
    lookup_vector_or_array: np.ndarray,
    result_vector: np.ndarray | None = None,
) -> CellValue:
    return _rt_xl_lookup(lookup_value, lookup_vector_or_array, result_vector)


@register("VLOOKUP")
def xl_vlookup(
    lookup_value: CellValue,
    table_array: np.ndarray,
    col_index_num: CellValue,
    range_lookup: CellValue = True,
) -> CellValue:
    """Search for a value in the first column and return a value in the same row from another column."""
    return _rt_xl_vlookup(lookup_value, table_array, col_index_num, range_lookup)


@register("HLOOKUP")
def xl_hlookup(
    lookup_value: CellValue,
    table_array: np.ndarray,
    row_index_num: CellValue,
    range_lookup: CellValue = True,
) -> CellValue:
    """Search for a value in the first row and return a value in the same column from another row."""
    return _rt_xl_hlookup(lookup_value, table_array, row_index_num, range_lookup)


@register("_XLFN.XLOOKUP")
@register("XLOOKUP")
def xl__xlfn_xlookup(
    lookup_value: CellValue,
    lookup_array: np.ndarray,
    return_array: np.ndarray,
    if_not_found: CellValue = None,
    match_mode: CellValue = 0,
    search_mode: CellValue = 1,
) -> CellValue:
    return _rt_xl__xlfn_xlookup(
        lookup_value,
        lookup_array,
        return_array,
        if_not_found,
        match_mode,
        search_mode,
    )
