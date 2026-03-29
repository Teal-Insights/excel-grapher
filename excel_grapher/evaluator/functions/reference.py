from __future__ import annotations

import re

import fastpyxl.utils.cell

from ..helpers import to_bool, to_number, to_string
from ..types import CellValue, XlError
from . import register


def _quote_sheet_name(sheet: str) -> str:
    # Excel quotes sheet names with spaces or special characters using single quotes.
    if re.fullmatch(r"[A-Za-z0-9_]+", sheet):
        return sheet
    escaped = sheet.replace("'", "''")
    return f"'{escaped}'"


@register("ADDRESS")
def xl_address(
    row_num: CellValue,
    column_num: CellValue,
    abs_num: CellValue = 1,
    a1: CellValue = True,
    sheet_text: CellValue = None,
) -> str | XlError:
    """Create a cell address as text (A1 style only)."""
    rn = to_number(row_num)
    if isinstance(rn, XlError):
        return rn
    cn = to_number(column_num)
    if isinstance(cn, XlError):
        return cn
    an = to_number(abs_num)
    if isinstance(an, XlError):
        return an
    a1_mode = to_bool(a1)
    if isinstance(a1_mode, XlError):
        return a1_mode
    if not a1_mode:
        return XlError.VALUE

    r = int(rn)
    c = int(cn)
    if r < 1 or c < 1:
        return XlError.VALUE

    abs_flag = int(an)
    if abs_flag not in (1, 2, 3, 4):
        return XlError.VALUE

    col_letters = fastpyxl.utils.cell.get_column_letter(c)

    col_abs = abs_flag in (1, 2)
    row_abs = abs_flag in (1, 3)
    col_prefix = "$" if col_abs else ""
    row_prefix = "$" if row_abs else ""

    addr = f"{col_prefix}{col_letters}{row_prefix}{r}"

    if sheet_text is None or sheet_text == "":
        return addr

    sheet = to_string(sheet_text)
    return f"{_quote_sheet_name(sheet)}!{addr}"
