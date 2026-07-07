from __future__ import annotations

from excel_grapher.core import CellValue, XlError
from excel_grapher.core.reference_funcs import (
    address_string,
    column_number,
    columns_count,
    row_number,
)

__all__ = ["xl_address", "xl_column", "xl_columns", "xl_row"]


def xl_address(
    row_num: CellValue,
    column_num: CellValue,
    abs_num: CellValue = 1,
    a1: CellValue = True,
    sheet_text: CellValue = None,
) -> str | XlError:
    return address_string(row_num, column_num, abs_num, a1, sheet_text)


def xl_row(ref: CellValue) -> int | XlError:
    return row_number(ref)


def xl_column(ref: CellValue) -> int | XlError:
    return column_number(ref)


def xl_columns(ref: CellValue) -> int | XlError:
    return columns_count(ref)
