"""TACO compressed dependency index."""

from __future__ import annotations

from dataclasses import dataclass, field

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import parse_address
from excel_grapher.grapher.node import NodeKey

from .types import CompressedEdge, RangeRef, SingleEdge


@dataclass
class TacoIndex:
    """Parallel index of compressed range-pattern dependency edges."""

    compressed_edges: list[CompressedEdge] = field(default_factory=list)
    single_edges: list[SingleEdge] = field(default_factory=list)

    def find_dependents(self, query: NodeKey | RangeRef) -> list[RangeRef]:
        """Return dependent ranges that depend on `query`."""
        if isinstance(query, RangeRef):
            return self._find_dependents_range(query)
        out: list[RangeRef] = []
        seen: set[tuple[str, str, int, str, int]] = set()
        for edge in self.compressed_edges:
            if edge.precedent.contains(query):
                sig = _range_sig(edge.dependent)
                if sig not in seen:
                    seen.add(sig)
                    out.append(edge.dependent)
        for single in self.single_edges:
            if single.precedent == query:
                ref = RangeRef.single_cell(*_split_key(single.dependent))
                sig = _range_sig(ref)
                if sig not in seen:
                    seen.add(sig)
                    out.append(ref)
        return out

    def find_precedents(self, query: NodeKey | RangeRef) -> list[RangeRef]:
        """Return precedent ranges that `query` depends on."""
        if isinstance(query, RangeRef):
            return self._find_precedents_range(query)
        out: list[RangeRef] = []
        seen: set[tuple[str, str, int, str, int]] = set()
        for edge in self.compressed_edges:
            if edge.dependent.contains(query):
                sig = _range_sig(edge.precedent)
                if sig not in seen:
                    seen.add(sig)
                    out.append(edge.precedent)
        for single in self.single_edges:
            if single.dependent == query:
                ref = RangeRef.single_cell(*_split_key(single.precedent))
                sig = _range_sig(ref)
                if sig not in seen:
                    seen.add(sig)
                    out.append(ref)
        return out

    def _find_dependents_range(self, query: RangeRef) -> list[RangeRef]:
        out: list[RangeRef] = []
        seen: set[tuple[str, str, int, str, int]] = set()
        for edge in self.compressed_edges:
            if _ranges_overlap(edge.precedent, query):
                sig = _range_sig(edge.dependent)
                if sig not in seen:
                    seen.add(sig)
                    out.append(edge.dependent)
        return out

    def _find_precedents_range(self, query: RangeRef) -> list[RangeRef]:
        out: list[RangeRef] = []
        seen: set[tuple[str, str, int, str, int]] = set()
        for edge in self.compressed_edges:
            if _ranges_overlap(edge.dependent, query):
                sig = _range_sig(edge.precedent)
                if sig not in seen:
                    seen.add(sig)
                    out.append(edge.precedent)
        return out


def _range_sig(ref: RangeRef) -> tuple[str, str, int, str, int]:
    return (ref.sheet, ref.min_col, ref.min_row, ref.max_col, ref.max_row)


def _split_key(key: NodeKey) -> tuple[str, str, int]:
    sheet, coord = parse_address(key)
    col, row = fastpyxl.utils.cell.coordinate_from_string(coord)
    return sheet, col, row


def _ranges_overlap(a: RangeRef, b: RangeRef) -> bool:
    if a.sheet != b.sheet:
        return False

    a_c1 = fastpyxl.utils.cell.column_index_from_string(a.min_col)
    a_c2 = fastpyxl.utils.cell.column_index_from_string(a.max_col)
    b_c1 = fastpyxl.utils.cell.column_index_from_string(b.min_col)
    b_c2 = fastpyxl.utils.cell.column_index_from_string(b.max_col)
    cols_overlap = a_c1 <= b_c2 and b_c1 <= a_c2
    rows_overlap = a.min_row <= b.max_row and b.min_row <= a.max_row
    return cols_overlap and rows_overlap
