"""TACO pattern operations and materialization."""

from __future__ import annotations

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import format_cell_key, parse_address

from .types import Orientation, PatternKind, PatternMeta, RangeRef


def is_rr_ref(*, is_absolute_col: bool, is_absolute_row: bool) -> bool:
    """Return True when a reference is relative in both dimensions (RR)."""
    return not is_absolute_col and not is_absolute_row


def is_rr_chain_ref(
    *,
    dep_col: str,
    dep_row: int,
    prec_col: str,
    prec_row: int,
    is_absolute_col: bool,
    is_absolute_row: bool,
    orientation: Orientation = Orientation.column,
) -> bool:
    """Return True when a cell ref is an RR-chain step along the run axis.

    Column runs: precedent is the cell directly above in the same column.
    Row runs: precedent is the cell directly to the left on the same row.
    """
    if not is_rr_ref(is_absolute_col=is_absolute_col, is_absolute_row=is_absolute_row):
        return False
    if orientation is Orientation.column:
        return prec_col == dep_col and prec_row == dep_row - 1
    dep_col_i = fastpyxl.utils.cell.column_index_from_string(dep_col)
    prec_col_i = fastpyxl.utils.cell.column_index_from_string(prec_col)
    return prec_row == dep_row and prec_col_i == dep_col_i - 1


def rr_materialize_precedent(dependent: RangeRef, precedent: RangeRef, dep_key: str) -> str:
    """Map one dependent cell key to its RR precedent cell key."""
    _, coord = parse_address(dep_key)
    dep_col, dep_row = fastpyxl.utils.cell.coordinate_from_string(coord)
    dep_col_i = fastpyxl.utils.cell.column_index_from_string(dep_col)
    rel_row = dep_row - dependent.min_row
    prec_row = precedent.min_row + rel_row
    rel_col = dep_col_i - fastpyxl.utils.cell.column_index_from_string(dependent.min_col)
    prec_col_i = fastpyxl.utils.cell.column_index_from_string(precedent.min_col) + rel_col
    prec_col = fastpyxl.utils.cell.get_column_letter(prec_col_i)
    return format_cell_key(precedent.sheet, prec_col, prec_row)


def rr_materialize_dependent(precedent: RangeRef, dependent: RangeRef, prec_key: str) -> str:
    """Map one precedent cell key to its RR dependent cell key."""
    _, coord = parse_address(prec_key)
    prec_col, prec_row = fastpyxl.utils.cell.coordinate_from_string(coord)
    prec_col_i = fastpyxl.utils.cell.column_index_from_string(prec_col)
    rel_row = prec_row - precedent.min_row
    dep_row = dependent.min_row + rel_row
    rel_col = prec_col_i - fastpyxl.utils.cell.column_index_from_string(precedent.min_col)
    dep_col_i = fastpyxl.utils.cell.column_index_from_string(dependent.min_col) + rel_col
    dep_col = fastpyxl.utils.cell.get_column_letter(dep_col_i)
    return format_cell_key(dependent.sheet, dep_col, dep_row)


def materialize_precedents_for_edge(
    edge_precedent: RangeRef,
    edge_dependent: RangeRef,
    meta: PatternMeta,
    dep_key: str,
) -> set[str]:
    """Materialize precedent cell keys for one dependent cell."""
    if meta.kind in (PatternKind.rr, PatternKind.rr_chain):
        return {rr_materialize_precedent(edge_dependent, edge_precedent, dep_key)}
    if meta.kind == PatternKind.rf:
        if meta.orientation is Orientation.row:
            return _rf_materialize_precedents_row(meta, dep_key, precedent=edge_precedent)
        return _rf_materialize_precedents(meta, dep_key, precedent=edge_precedent)
    if meta.kind == PatternKind.fr:
        if meta.orientation is Orientation.row:
            return _fr_materialize_precedents_row(meta, dep_key, precedent=edge_precedent)
        return _fr_materialize_precedents(meta, dep_key, precedent=edge_precedent)
    if meta.kind == PatternKind.ff:
        return set(edge_precedent.cell_keys())
    return set(edge_precedent.cell_keys())


def materialize_dependents_for_edge(
    edge_precedent: RangeRef,
    edge_dependent: RangeRef,
    meta: PatternMeta,
    prec_key: str,
) -> set[str]:
    """Materialize dependent cell keys for one precedent cell."""
    if meta.kind in (PatternKind.rr, PatternKind.rr_chain):
        return {rr_materialize_dependent(edge_precedent, edge_dependent, prec_key)}
    if meta.kind == PatternKind.rf:
        if meta.orientation is Orientation.row:
            return _rf_materialize_dependents_row(edge_dependent, meta, prec_key)
        return _rf_materialize_dependents(edge_dependent, meta, prec_key)
    if meta.kind == PatternKind.fr:
        if meta.orientation is Orientation.row:
            return _fr_materialize_dependents_row(edge_dependent, meta, prec_key)
        return _fr_materialize_dependents(edge_dependent, meta, prec_key)
    if meta.kind == PatternKind.ff:
        return set(edge_dependent.cell_keys())
    return set(edge_dependent.cell_keys())


def _dep_row_col(dep_key: str) -> tuple[str, str, int]:
    sheet, coord = parse_address(dep_key)
    col, row = fastpyxl.utils.cell.coordinate_from_string(coord)
    return sheet, col, row


def _rf_materialize_precedents(meta: PatternMeta, dep_key: str, *, precedent: RangeRef) -> set[str]:
    _, col, row = _dep_row_col(dep_key)
    sheet = precedent.sheet
    assert meta.fixed_tail_col is not None and meta.fixed_tail_row is not None
    tail_col = meta.fixed_tail_col
    tail_row = meta.fixed_tail_row
    head_row = row
    if head_row > tail_row:
        return set()
    out: set[str] = set()
    for r in range(head_row, tail_row + 1):
        out.add(format_cell_key(sheet, tail_col, r))
    return out


def _rf_materialize_dependents(dependent: RangeRef, meta: PatternMeta, prec_key: str) -> set[str]:
    _, col, row = _dep_row_col(prec_key)
    assert meta.fixed_tail_col is not None and meta.fixed_tail_row is not None
    if col != meta.fixed_tail_col or row > meta.fixed_tail_row:
        return set()
    out: set[str] = set()
    for dep_key in dependent.cell_keys():
        _, _, dep_row = _dep_row_col(dep_key)
        if dep_row <= row <= meta.fixed_tail_row:
            out.add(dep_key)
    return out


def _fr_materialize_precedents(meta: PatternMeta, dep_key: str, *, precedent: RangeRef) -> set[str]:
    _, col, row = _dep_row_col(dep_key)
    sheet = precedent.sheet
    assert meta.fixed_head_col is not None and meta.fixed_head_row is not None
    head_col = meta.fixed_head_col
    head_row = meta.fixed_head_row
    tail_row = row
    if tail_row < head_row:
        return set()
    out: set[str] = set()
    for r in range(head_row, tail_row + 1):
        out.add(format_cell_key(sheet, head_col, r))
    return out


def _fr_materialize_dependents(dependent: RangeRef, meta: PatternMeta, prec_key: str) -> set[str]:
    _, col, row = _dep_row_col(prec_key)
    assert meta.fixed_head_col is not None and meta.fixed_head_row is not None
    if col != meta.fixed_head_col or row < meta.fixed_head_row:
        return set()
    out: set[str] = set()
    for dep_key in dependent.cell_keys():
        _, _, dep_row = _dep_row_col(dep_key)
        if dep_row >= row:
            out.add(dep_key)
    return out


def _rf_materialize_precedents_row(
    meta: PatternMeta, dep_key: str, *, precedent: RangeRef
) -> set[str]:
    _, dep_col, _ = _dep_row_col(dep_key)
    dep_col_i = fastpyxl.utils.cell.column_index_from_string(dep_col)
    head_col_i = dep_col_i + meta.col_offset
    assert meta.fixed_tail_col is not None and meta.fixed_tail_row is not None
    tail_col_i = fastpyxl.utils.cell.column_index_from_string(meta.fixed_tail_col)
    if head_col_i > tail_col_i:
        return set()
    sheet = precedent.sheet
    row = meta.fixed_tail_row
    return {
        format_cell_key(sheet, fastpyxl.utils.cell.get_column_letter(col_i), row)
        for col_i in range(head_col_i, tail_col_i + 1)
    }


def _rf_materialize_dependents_row(
    dependent: RangeRef, meta: PatternMeta, prec_key: str
) -> set[str]:
    _, prec_col, prec_row = _dep_row_col(prec_key)
    assert meta.fixed_tail_col is not None and meta.fixed_tail_row is not None
    prec_col_i = fastpyxl.utils.cell.column_index_from_string(prec_col)
    tail_col_i = fastpyxl.utils.cell.column_index_from_string(meta.fixed_tail_col)
    if prec_row != meta.fixed_tail_row or prec_col_i > tail_col_i:
        return set()
    out: set[str] = set()
    for dep_key in dependent.cell_keys():
        _, dep_col, dep_row = _dep_row_col(dep_key)
        if dep_row != meta.fixed_tail_row:
            continue
        dep_col_i = fastpyxl.utils.cell.column_index_from_string(dep_col)
        head_col_i = dep_col_i + meta.col_offset
        if head_col_i <= prec_col_i <= tail_col_i:
            out.add(dep_key)
    return out


def _fr_materialize_precedents_row(
    meta: PatternMeta, dep_key: str, *, precedent: RangeRef
) -> set[str]:
    _, dep_col, _ = _dep_row_col(dep_key)
    assert meta.fixed_head_col is not None and meta.fixed_head_row is not None
    head_col_i = fastpyxl.utils.cell.column_index_from_string(meta.fixed_head_col)
    tail_col_i = fastpyxl.utils.cell.column_index_from_string(dep_col)
    if tail_col_i < head_col_i:
        return set()
    sheet = precedent.sheet
    row = meta.fixed_head_row
    return {
        format_cell_key(sheet, fastpyxl.utils.cell.get_column_letter(col_i), row)
        for col_i in range(head_col_i, tail_col_i + 1)
    }


def _fr_materialize_dependents_row(
    dependent: RangeRef, meta: PatternMeta, prec_key: str
) -> set[str]:
    _, prec_col, prec_row = _dep_row_col(prec_key)
    assert meta.fixed_head_col is not None and meta.fixed_head_row is not None
    head_col_i = fastpyxl.utils.cell.column_index_from_string(meta.fixed_head_col)
    prec_col_i = fastpyxl.utils.cell.column_index_from_string(prec_col)
    if prec_col_i < head_col_i or prec_row != meta.fixed_head_row:
        return set()
    out: set[str] = set()
    for dep_key in dependent.cell_keys():
        _, dep_col, dep_row = _dep_row_col(dep_key)
        if dep_row != meta.fixed_head_row:
            continue
        dep_col_i = fastpyxl.utils.cell.column_index_from_string(dep_col)
        if dep_col_i >= prec_col_i:
            out.add(dep_key)
    return out
