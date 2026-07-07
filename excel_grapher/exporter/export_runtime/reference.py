"""Raise-only boundary wrappers for reference worksheet functions."""

from __future__ import annotations

from excel_grapher.core import CellValue
from excel_grapher.core.reference_funcs import (
    address_string,
    column_number,
    columns_count,
    row_number,
)

from .errors import raise_if_sentinel_int, raise_if_sentinel_str

__all__ = ["xl_address", "xl_column", "xl_columns", "xl_row"]


def xl_address(
    row_num: CellValue,
    column_num: CellValue,
    abs_num: CellValue = 1,
    a1: CellValue = True,
    sheet_text: CellValue = None,
) -> str:
    """Build an A1-style address string, raising on Excel errors."""
    return raise_if_sentinel_str(address_string(row_num, column_num, abs_num, a1, sheet_text))


def xl_row(ref: CellValue) -> int:
    """Return the row number of a reference, raising on Excel errors."""
    return raise_if_sentinel_int(row_number(ref))


def xl_column(ref: CellValue) -> int:
    """Return the column number of a reference, raising on Excel errors."""
    return raise_if_sentinel_int(column_number(ref))


def xl_columns(ref: CellValue) -> int:
    """Return the column count of a reference, raising on Excel errors."""
    return raise_if_sentinel_int(columns_count(ref))
