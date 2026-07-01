"""Export-owned runtime primitives for generated Python code.

The lazy range value is named `Range`. It stores rectangular worksheet geometry
and a resolver callable with the shape `resolver(address: str) -> CellValue`.
Range consumers in this package accept lazy ranges (and nested lists) instead
of numpy object arrays.
"""

from .aggregates import xl_sumproduct
from .errors import XlErrorException
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
    "xl_index",
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
    "xl_range",
    "xl_range_rows",
    "xl_sub",
    "xl_sumproduct",
    "xl_vlookup",
    "xl_xlookup",
]
