"""Export-owned runtime primitives for generated Python code.

The lazy range value is named `Range`. It stores rectangular worksheet geometry
and a resolver callable with the shape `resolver(address: str) -> CellValue`.
Range consumers in this package accept lazy ranges (and nested lists) instead
of numpy object arrays.

Excel errors raise `XlErrorException` in exported code; error-consuming
functions (`IFERROR`, `IS*`) receive lazily-evaluated thunks.
"""

from .aggregates import xl_sumproduct
from .error_funcs import xl_iferror, xl_ifna, xl_isblank, xl_iserror, xl_isna
from .errors import XlErrorException, xl_raise
from .lookup import xl_hlookup, xl_index, xl_lookup, xl_match, xl_vlookup, xl_xlookup
from .offset import xl_offset, xl_range, xl_range_rows
from .operators import (
    xl_compare,
    xl_is_array,
    xl_map_arithmetic,
    xl_map_compare,
    xl_map_concat,
    xl_map_unary,
    xl_number,
    xl_pow_numbers,
)
from .ranges import Range
from .values import ExcelRange, Grid, flatten

__all__ = [
    "ExcelRange",
    "Grid",
    "Range",
    "XlErrorException",
    "flatten",
    "xl_compare",
    "xl_hlookup",
    "xl_iferror",
    "xl_ifna",
    "xl_index",
    "xl_is_array",
    "xl_isblank",
    "xl_iserror",
    "xl_isna",
    "xl_lookup",
    "xl_map_arithmetic",
    "xl_map_compare",
    "xl_map_concat",
    "xl_map_unary",
    "xl_match",
    "xl_number",
    "xl_offset",
    "xl_pow_numbers",
    "xl_raise",
    "xl_range",
    "xl_range_rows",
    "xl_sumproduct",
    "xl_vlookup",
    "xl_xlookup",
]
