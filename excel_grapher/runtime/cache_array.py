"""Array-aware range evaluation helpers for exported runtime (issue #284)."""

from __future__ import annotations

import fastpyxl.utils.cell

from excel_grapher.core import CellValue, ExcelRange, XlError
from excel_grapher.core.address_keys import normalize_key
from excel_grapher.core.array_results import read_spill_scalar, scalar_for_range_member

from .cache import EvalContext, _parse_range_address, xl_cell

__all__ = [
    "xl_cell_in_range",
    "xl_range",
]


def xl_cell_in_range(ctx: EvalContext, address: str) -> CellValue:
    """Evaluate one range member, projecting array anchors to spilled scalars."""
    norm = normalize_key(address)
    if ctx.resolver(norm) is not None or norm in ctx.inputs or norm in ctx.cache:
        value = xl_cell(ctx, norm)
        return scalar_for_range_member(norm, value)
    spilled = read_spill_scalar(norm, ctx.cache)
    if spilled is not None:
        return spilled
    raise KeyError(f"Cell {address} not found in graph")


def xl_range(ctx: EvalContext, address: str) -> CellValue:
    """Evaluate a sheet-qualified range and return a 2D numpy array of values."""
    parsed = _parse_range_address(address)
    if isinstance(parsed, XlError):
        return parsed
    sheet, start_cell, end_cell = parsed
    try:
        start_col, start_row = fastpyxl.utils.cell.coordinate_from_string(start_cell)
        end_col, end_row = fastpyxl.utils.cell.coordinate_from_string(end_cell)
        start_col_idx = fastpyxl.utils.cell.column_index_from_string(start_col)
        end_col_idx = fastpyxl.utils.cell.column_index_from_string(end_col)
    except ValueError:
        return XlError.VALUE

    if start_row > end_row:
        start_row, end_row = end_row, start_row
    if start_col_idx > end_col_idx:
        start_col_idx, end_col_idx = end_col_idx, start_col_idx

    rng = ExcelRange(sheet, start_row, start_col_idx, end_row, end_col_idx)
    return rng.resolve(lambda addr: xl_cell_in_range(ctx, addr))
