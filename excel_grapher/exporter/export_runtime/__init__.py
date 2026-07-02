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
    xl_add,
    xl_concat,
    xl_div,
    xl_eq,
    xl_ge,
    xl_gt,
    xl_le,
    xl_lt,
    xl_mul,
    xl_ne,
    xl_neg,
    xl_percent,
    xl_pos,
    xl_pow,
    xl_sub,
)
from .ranges import Range
from .values import ExcelRange, Grid, flatten

__all__ = [
    "ExcelRange",
    "Grid",
    "Range",
    "XlErrorException",
    "flatten",
    "xl_add",
    "xl_concat",
    "xl_div",
    "xl_eq",
    "xl_ge",
    "xl_gt",
    "xl_hlookup",
    "xl_iferror",
    "xl_ifna",
    "xl_index",
    "xl_isblank",
    "xl_iserror",
    "xl_isna",
    "xl_le",
    "xl_lookup",
    "xl_lt",
    "xl_match",
    "xl_mul",
    "xl_ne",
    "xl_neg",
    "xl_offset",
    "xl_percent",
    "xl_pos",
    "xl_pow",
    "xl_raise",
    "xl_range",
    "xl_range_rows",
    "xl_sub",
    "xl_sumproduct",
    "xl_vlookup",
    "xl_xlookup",
]
