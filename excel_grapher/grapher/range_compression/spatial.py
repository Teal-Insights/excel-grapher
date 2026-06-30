"""Spatial overlap index for TACO range queries."""

from __future__ import annotations

import fastpyxl.utils.cell

from .types import RangeRef


class RangeSpatialIndex:
    """Sheet-local index mapping ranges to edge indices for overlap queries."""

    def __init__(self) -> None:
        self._entries: dict[str, list[tuple[RangeRef, int]]] = {}

    def add(self, ref: RangeRef, edge_index: int) -> None:
        """Register a range for later overlap or point queries."""
        self._entries.setdefault(ref.sheet, []).append((ref, edge_index))

    def query_point(self, sheet: str, column: str, row: int) -> list[int]:
        """Return edge indices whose registered range contains the cell."""
        col_i = fastpyxl.utils.cell.column_index_from_string(column)
        out: list[int] = []
        for ref, edge_index in self._entries.get(sheet, []):
            c_lo = fastpyxl.utils.cell.column_index_from_string(ref.min_col)
            c_hi = fastpyxl.utils.cell.column_index_from_string(ref.max_col)
            if c_lo <= col_i <= c_hi and ref.min_row <= row <= ref.max_row:
                out.append(edge_index)
        return out

    def query_overlap(self, query: RangeRef) -> list[int]:
        """Return edge indices whose registered range overlaps `query`."""
        out: list[int] = []
        q_c1 = fastpyxl.utils.cell.column_index_from_string(query.min_col)
        q_c2 = fastpyxl.utils.cell.column_index_from_string(query.max_col)
        for ref, edge_index in self._entries.get(query.sheet, []):
            c_lo = fastpyxl.utils.cell.column_index_from_string(ref.min_col)
            c_hi = fastpyxl.utils.cell.column_index_from_string(ref.max_col)
            cols_overlap = c_lo <= q_c2 and q_c1 <= c_hi
            rows_overlap = ref.min_row <= query.max_row and query.min_row <= ref.max_row
            if cols_overlap and rows_overlap:
                out.append(edge_index)
        return out
