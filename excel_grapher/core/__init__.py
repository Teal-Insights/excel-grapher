"""
Shared Excel semantics (types, coercions, operators).

Representation-agnostic types and logic used by both the evaluator runtime
and the standalone export runtime.
"""

from .coercions import (
    excel_casefold,
    flatten,
    get_error,
    numeric_values,
    to_bool,
    to_int,
    to_native,
    to_number,
    to_string,
)
from .operators import (
    _xl_compare,
    xl_add,
    xl_concat,
    xl_div,
    xl_eq,
    xl_ge,
    xl_gt,
    xl_iferror,
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
from .types import CellValue, ExcelRange, XlError

__all__ = [
    "CellValue",
    "ExcelRange",
    "XlError",
    "excel_casefold",
    "flatten",
    "get_error",
    "numeric_values",
    "to_bool",
    "to_int",
    "to_native",
    "to_number",
    "to_string",
    "_xl_compare",
    "xl_add",
    "xl_concat",
    "xl_div",
    "xl_eq",
    "xl_ge",
    "xl_gt",
    "xl_iferror",
    "xl_le",
    "xl_lt",
    "xl_mul",
    "xl_ne",
    "xl_neg",
    "xl_percent",
    "xl_pos",
    "xl_pow",
    "xl_sub",
]
