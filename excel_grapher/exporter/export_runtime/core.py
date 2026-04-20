"""Re-export shared types and coercions from excel_grapher.core for export runtime consumers."""

from __future__ import annotations

from excel_grapher.core import (
    CellValue,
    ExcelRange,
    XlError,
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
from excel_grapher.core.coercions import _format_general_number

__all__ = [
    "CellValue",
    "ExcelRange",
    "XlError",
    "_format_general_number",
    "excel_casefold",
    "flatten",
    "get_error",
    "numeric_values",
    "to_bool",
    "to_int",
    "to_native",
    "to_number",
    "to_string",
]
