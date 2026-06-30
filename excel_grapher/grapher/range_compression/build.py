"""Build a TACO index from an existing dependency graph."""

from __future__ import annotations

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import parse_address
from excel_grapher.grapher.dependency_provenance import DependencyCause
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import NodeKey

from .grouping import column_adjacent_groups
from .index import TacoIndex
from .patterns import is_rr_ref
from .ref_parser import abs_ref_to_key, parse_cell_refs_with_abs
from .types import CompressedEdge, PatternKind, PatternMeta, RangeRef, SingleEdge

_EXCLUDED_CAUSES = frozenset(
    {
        DependencyCause.static_range,
        DependencyCause.dynamic_offset,
        DependencyCause.dynamic_indirect,
    }
)


def build_taco_index(graph: DependencyGraph) -> TacoIndex:
    """Build a TACO compressed index from `graph` without mutating it."""
    index = TacoIndex()
    covered: set[tuple[NodeKey, NodeKey]] = set()

    for group in column_adjacent_groups(graph):
        _compress_rr_group(graph, index, group, covered)

    for dep in graph:
        for prec in graph.get_dependencies(dep):
            pair = (dep, prec)
            if pair in covered:
                continue
            index.single_edges.append(SingleEdge(precedent=prec, dependent=dep))
            covered.add(pair)

    return index


def _is_compressible_pair(graph: DependencyGraph, dependent: NodeKey, precedent: NodeKey) -> bool:
    if graph.get_edge_guard(dependent, precedent) is not None:
        return False
    attrs = graph.get_edge_attrs(dependent, precedent)
    prov = attrs.provenance
    return prov is None or not (prov.causes & _EXCLUDED_CAUSES)


def _compress_rr_group(
    graph: DependencyGraph,
    index: TacoIndex,
    group: list[NodeKey],
    covered: set[tuple[NodeKey, NodeKey]],
) -> None:
    first = graph.get_node(group[0])
    if first is None or not first.formula:
        return
    ref_count = len(parse_cell_refs_with_abs(first.formula, default_sheet=first.sheet))
    if ref_count == 0:
        return

    for ref_idx in range(ref_count):
        stream = _collect_rr_stream(graph, group, ref_idx)
        if stream is None:
            continue
        dep_range, prec_range, meta = stream
        edge = CompressedEdge(precedent=prec_range, dependent=dep_range, meta=meta)
        index.compressed_edges.append(edge)
        for dep_key in dep_range.cell_keys():
            prec_key = _rr_pair_key(dep_range, prec_range, dep_key)
            covered.add((dep_key, prec_key))


def _collect_rr_stream(
    graph: DependencyGraph,
    group: list[NodeKey],
    ref_idx: int,
) -> tuple[RangeRef, RangeRef, PatternMeta] | None:
    dep_sheet: str | None = None
    dep_col: str | None = None
    prec_col: str | None = None
    first_row: int | None = None
    last_row: int | None = None
    col_offset: int | None = None
    row_offset: int | None = None

    for dep_key in group:
        node = graph.get_node(dep_key)
        if node is None or not node.formula:
            return None
        refs = parse_cell_refs_with_abs(node.formula, default_sheet=node.sheet)
        if ref_idx >= len(refs):
            return None
        ref = refs[ref_idx]
        if not is_rr_ref(is_absolute_col=ref.is_absolute_col, is_absolute_row=ref.is_absolute_row):
            return None
        prec_key = abs_ref_to_key(ref, default_sheet=node.sheet)
        if prec_key not in graph.get_dependencies(dep_key):
            return None
        if not _is_compressible_pair(graph, dep_key, prec_key):
            return None

        dep_c = node.column
        dep_r = node.row
        _, prec_coord = parse_address(prec_key)
        prec_c, prec_r = fastpyxl.utils.cell.coordinate_from_string(prec_coord)

        dep_col_i = fastpyxl.utils.cell.column_index_from_string(dep_c)
        prec_col_i = fastpyxl.utils.cell.column_index_from_string(prec_c)
        this_col_offset = prec_col_i - dep_col_i
        this_row_offset = prec_r - dep_r

        if dep_sheet is None:
            dep_sheet = node.sheet
            dep_col = dep_c
            prec_col = prec_c
            first_row = dep_r
            last_row = dep_r
            col_offset = this_col_offset
            row_offset = this_row_offset
        else:
            if node.sheet != dep_sheet or dep_c != dep_col or prec_c != prec_col:
                return None
            if this_col_offset != col_offset or this_row_offset != row_offset:
                return None
            assert last_row is not None
            if dep_r != last_row + 1:
                return None
            last_row = dep_r

    assert dep_sheet is not None and dep_col is not None and prec_col is not None
    assert first_row is not None and last_row is not None
    assert col_offset is not None and row_offset is not None

    dep_range = RangeRef.column_span(dep_sheet, dep_col, first_row, last_row)
    prec_range = RangeRef.column_span(
        dep_sheet, prec_col, first_row + row_offset, last_row + row_offset
    )
    meta = PatternMeta(kind=PatternKind.rr, col_offset=col_offset, row_offset=row_offset)
    return dep_range, prec_range, meta


def _rr_pair_key(dep_range: RangeRef, prec_range: RangeRef, dep_key: NodeKey) -> NodeKey:
    from .patterns import rr_materialize_precedent

    return rr_materialize_precedent(dep_range, prec_range, dep_key)
