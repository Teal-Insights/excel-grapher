"""TACO pattern operations (v1: RR and Single)."""

from __future__ import annotations

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import format_cell_key, parse_address

from .types import PatternKind, PatternMeta, RangeRef


def is_rr_ref(*, is_absolute_col: bool, is_absolute_row: bool) -> bool:
    """Return True when a reference is relative in both dimensions (RR)."""
    return not is_absolute_col and not is_absolute_row


def rr_materialize_precedent(dependent: RangeRef, precedent: RangeRef, dep_key: str) -> str:
    """Map one dependent cell key to its RR precedent cell key."""
    sheet, coord = parse_address(dep_key)
    dep_col, dep_row = fastpyxl.utils.cell.coordinate_from_string(coord)
    dep_col_i = fastpyxl.utils.cell.column_index_from_string(dep_col)
    rel_row = dep_row - dependent.min_row
    prec_row = precedent.min_row + rel_row
    rel_col = dep_col_i - fastpyxl.utils.cell.column_index_from_string(dependent.min_col)
    prec_col_i = fastpyxl.utils.cell.column_index_from_string(precedent.min_col) + rel_col
    prec_col = fastpyxl.utils.cell.get_column_letter(prec_col_i)
    return format_cell_key(sheet, prec_col, prec_row)


def rr_materialize_dependent(precedent: RangeRef, dependent: RangeRef, prec_key: str) -> str:
    """Map one precedent cell key to its RR dependent cell key."""
    sheet, coord = parse_address(prec_key)
    prec_col, prec_row = fastpyxl.utils.cell.coordinate_from_string(coord)
    prec_col_i = fastpyxl.utils.cell.column_index_from_string(prec_col)
    rel_row = prec_row - precedent.min_row
    dep_row = dependent.min_row + rel_row
    rel_col = prec_col_i - fastpyxl.utils.cell.column_index_from_string(precedent.min_col)
    dep_col_i = fastpyxl.utils.cell.column_index_from_string(dependent.min_col) + rel_col
    dep_col = fastpyxl.utils.cell.get_column_letter(dep_col_i)
    return format_cell_key(sheet, dep_col, dep_row)


def validate_rr_edge(precedent: RangeRef, dependent: RangeRef, meta: PatternMeta) -> bool:
    """Return True when `meta` describes a consistent RR relationship."""
    if meta.kind != PatternKind.rr:
        return False
    dep_rows = dependent.max_row - dependent.min_row
    prec_rows = precedent.max_row - precedent.min_row
    if dep_rows != prec_rows:
        return False
    dep_cols = fastpyxl.utils.cell.column_index_from_string(
        dependent.max_col
    ) - fastpyxl.utils.cell.column_index_from_string(dependent.min_col)
    prec_cols = fastpyxl.utils.cell.column_index_from_string(
        precedent.max_col
    ) - fastpyxl.utils.cell.column_index_from_string(precedent.min_col)
    return dep_cols == prec_cols
