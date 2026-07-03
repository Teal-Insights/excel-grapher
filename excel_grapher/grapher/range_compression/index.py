"""TACO compressed dependency index."""

from __future__ import annotations

from dataclasses import dataclass, field

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import parse_address
from excel_grapher.grapher.node import NodeKey

from .spatial import RangeSpatialIndex
from .types import CompressedEdge, RangeRef, SingleEdge


@dataclass
class TacoIndex:
    """Parallel index of compressed range-pattern dependency edges."""

    compressed_edges: list[CompressedEdge] = field(default_factory=list)
    single_edges: list[SingleEdge] = field(default_factory=list)
    _prec_spatial: RangeSpatialIndex = field(default_factory=RangeSpatialIndex, repr=False)
    _dep_spatial: RangeSpatialIndex = field(default_factory=RangeSpatialIndex, repr=False)
    _single_prec: dict[NodeKey, list[NodeKey]] = field(default_factory=dict, repr=False)
    _single_dep: dict[NodeKey, list[NodeKey]] = field(default_factory=dict, repr=False)

    def rebuild_spatial_indices(self) -> None:
        """Rebuild spatial lookup structures after edges are added."""
        self._prec_spatial = RangeSpatialIndex()
        self._dep_spatial = RangeSpatialIndex()
        self._single_prec = {}
        self._single_dep = {}
        for i, edge in enumerate(self.compressed_edges):
            self._prec_spatial.add(edge.precedent, i)
            self._dep_spatial.add(edge.dependent, i)
        for single in self.single_edges:
            self._single_prec.setdefault(single.precedent, []).append(single.dependent)
            self._single_dep.setdefault(single.dependent, []).append(single.precedent)

    def find_dependents(self, query: NodeKey | RangeRef) -> list[RangeRef]:
        """Return dependent ranges that depend on `query`."""
        if isinstance(query, RangeRef):
            return self._find_dependents_range(query)
        sheet, col, row = _split_key(query)
        out: list[RangeRef] = []
        seen: set[tuple[str, str, int, str, int]] = set()
        for edge_index in self._prec_spatial.query_point(sheet, col, row):
            edge = self.compressed_edges[edge_index]
            sig = _range_sig(edge.dependent)
            if sig not in seen:
                seen.add(sig)
                out.append(edge.dependent)
        for dep_key in self._single_prec.get(query, []):
            ref = RangeRef.single_cell(*_split_key(dep_key))
            sig = _range_sig(ref)
            if sig not in seen:
                seen.add(sig)
                out.append(ref)
        return out

    def find_precedents(self, query: NodeKey | RangeRef) -> list[RangeRef]:
        """Return precedent ranges that `query` depends on."""
        if isinstance(query, RangeRef):
            return self._find_precedents_range(query)
        sheet, col, row = _split_key(query)
        out: list[RangeRef] = []
        seen: set[tuple[str, str, int, str, int]] = set()
        for edge_index in self._dep_spatial.query_point(sheet, col, row):
            edge = self.compressed_edges[edge_index]
            sig = _range_sig(edge.precedent)
            if sig not in seen:
                seen.add(sig)
                out.append(edge.precedent)
        for prec_key in self._single_dep.get(query, []):
            ref = RangeRef.single_cell(*_split_key(prec_key))
            sig = _range_sig(ref)
            if sig not in seen:
                seen.add(sig)
                out.append(ref)
        return out

    def _find_dependents_range(self, query: RangeRef) -> list[RangeRef]:
        out: list[RangeRef] = []
        seen: set[tuple[str, str, int, str, int]] = set()
        for edge_index in self._prec_spatial.query_overlap(query):
            edge = self.compressed_edges[edge_index]
            sig = _range_sig(edge.dependent)
            if sig not in seen:
                seen.add(sig)
                out.append(edge.dependent)
        return out

    def _find_precedents_range(self, query: RangeRef) -> list[RangeRef]:
        out: list[RangeRef] = []
        seen: set[tuple[str, str, int, str, int]] = set()
        for edge_index in self._dep_spatial.query_overlap(query):
            edge = self.compressed_edges[edge_index]
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
