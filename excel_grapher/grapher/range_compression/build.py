"""Build a TACO index from an existing dependency graph."""

from __future__ import annotations

from collections.abc import Iterable, Sequence
from pathlib import Path
from typing import TYPE_CHECKING

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import parse_address
from excel_grapher.grapher.dependency_provenance import DependencyCause
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import NodeKey

from .boundaries import (
    precedent_may_compress,
    range_ref_dependents_may_compress,
    range_ref_precedents_may_compress,
)
from .config import TacoBuildConfig
from .grouping import Orientation, adjacent_groups
from .index import TacoIndex
from .patterns import is_rr_chain_ref, is_rr_ref, materialize_precedents_for_edge
from .ref_parser import (
    AbsCellRef,
    AbsRangeRef,
    FormulaRef,
    abs_ref_to_key,
    classify_range_pattern,
    parse_ref_streams,
    range_ref_to_keys,
)
from .types import CompressedEdge, PatternKind, PatternMeta, RangeRef, SingleEdge

if TYPE_CHECKING:
    from excel_grapher.series_bindings.types import WorkbookSeriesBindings

_EXCLUDED_CAUSES = frozenset(
    {
        DependencyCause.static_range,
        DependencyCause.dynamic_offset,
        DependencyCause.dynamic_indirect,
    }
)


def build_taco_index(
    graph: DependencyGraph,
    config: TacoBuildConfig | None = None,
) -> TacoIndex:
    """Build a TACO compressed index from `graph` without mutating it.

    Uses column- and row-adjacent grouping (column-first) so fill-down and
    fill-right autofill runs both compress when pattern rules match.
    """
    cfg = config or TacoBuildConfig()
    index = TacoIndex()
    covered: set[tuple[NodeKey, NodeKey]] = set()

    for group in adjacent_groups(graph, config=cfg):
        orientation = _infer_group_orientation(graph, group)
        _compress_group(graph, index, group, covered, cfg, orientation)

    for dep in graph:
        for prec in graph.get_dependencies(dep):
            pair = (dep, prec)
            if pair in covered:
                continue
            index.single_edges.append(SingleEdge(precedent=prec, dependent=dep))
            covered.add(pair)

    index.rebuild_spatial_indices()
    return index


def build_codegen_taco_index(
    graph: DependencyGraph,
    *,
    input_ranges: Sequence[str] | None = None,
    series_bindings: WorkbookSeriesBindings | None = None,
    bindings_workbook: Path | str | None = None,
    export_addresses: Iterable[str] | None = None,
    internal_only: bool = True,
    attach_to_graph: bool = False,
) -> TacoIndex:
    """Build a codegen-boundary TACO index for export planning.

    Targets, declared ``input_ranges``, graph input leaves, and series-binding
    setter leaves stay at cell granularity on the dependent side.
    """
    config = TacoBuildConfig.for_codegen_export(
        graph,
        input_ranges=input_ranges,
        series_bindings=series_bindings,
        bindings_workbook=bindings_workbook,
        export_addresses=export_addresses,
        internal_only=internal_only,
    )
    index = build_taco_index(graph, config)
    if attach_to_graph:
        graph.codegen_taco_index = index
    return index


def _edge_is_excluded_from_pattern_inference(
    graph: DependencyGraph, dependent: NodeKey, precedent: NodeKey
) -> bool:
    if graph.get_edge_guard(dependent, precedent) is not None:
        return True
    attrs = graph.get_edge_attrs(dependent, precedent)
    prov = attrs.provenance
    if prov is None:
        return False
    return bool(prov.causes & (_EXCLUDED_CAUSES - {DependencyCause.static_range}))


def _infer_group_orientation(graph: DependencyGraph, group: list[NodeKey]) -> Orientation:
    """Infer whether a grouped run advances down a column or across a row."""
    cols: set[str] = set()
    rows: set[int] = set()
    for key in group:
        node = graph.get_node(key)
        if node is None:
            continue
        cols.add(node.column)
        rows.add(node.row)
    if len(cols) == 1:
        return Orientation.column
    if len(rows) == 1:
        return Orientation.row
    raise ValueError(f"ambiguous TACO group orientation: {group!r}")


def _compress_group(
    graph: DependencyGraph,
    index: TacoIndex,
    group: list[NodeKey],
    covered: set[tuple[NodeKey, NodeKey]],
    config: TacoBuildConfig,
    orientation: Orientation,
) -> None:
    first = graph.get_node(group[0])
    if first is None or not first.formula:
        return
    streams = parse_ref_streams(first.formula, default_sheet=first.sheet)
    if not streams:
        return

    for ref_idx in range(len(streams)):
        if isinstance(streams[ref_idx], AbsCellRef):
            stream = _collect_cell_stream(graph, group, ref_idx, config, orientation)
        else:
            stream = _collect_range_stream(graph, group, ref_idx, config, orientation)
        if stream is None:
            continue
        dep_range, prec_range, meta = stream
        if not range_ref_dependents_may_compress(graph, dep_range, config):
            continue
        if not range_ref_precedents_may_compress(graph, prec_range, config):
            continue
        edge = CompressedEdge(precedent=prec_range, dependent=dep_range, meta=meta)
        index.compressed_edges.append(edge)
        for dep_key in dep_range.cell_keys():
            for prec_key in materialize_precedents_for_edge(prec_range, dep_range, meta, dep_key):
                if _edge_is_excluded_from_pattern_inference(graph, dep_key, prec_key):
                    continue
                covered.add((dep_key, prec_key))


def _collect_cell_stream(
    graph: DependencyGraph,
    group: list[NodeKey],
    ref_idx: int,
    config: TacoBuildConfig,
    orientation: Orientation,
) -> tuple[RangeRef, RangeRef, PatternMeta] | None:
    dep_sheet: str | None = None
    prec_sheet: str | None = None
    dep_fixed: str | int | None = None
    prec_fixed: str | int | None = None
    first_advance: int | None = None
    last_advance: int | None = None
    col_offset: int | None = None
    row_offset: int | None = None
    chain = False

    for dep_key in group:
        node = graph.get_node(dep_key)
        if node is None or not node.formula:
            return None
        streams = parse_ref_streams(node.formula, default_sheet=node.sheet)
        if ref_idx >= len(streams):
            return None
        stream_item = streams[ref_idx]
        if not isinstance(stream_item, AbsCellRef):
            return None
        ref = stream_item
        if not is_rr_ref(is_absolute_col=ref.is_absolute_col, is_absolute_row=ref.is_absolute_row):
            return None
        prec_key = abs_ref_to_key(ref, default_sheet=node.sheet)
        if prec_key not in graph.get_dependencies(dep_key):
            return None
        if not precedent_may_compress(graph, prec_key, config):
            return None
        if _edge_is_excluded_from_pattern_inference(graph, dep_key, prec_key):
            return None
        if not _stream_expected_deps(graph, node, streams, ref_idx, {prec_key}):
            return None

        dep_c = node.column
        dep_r = node.row
        dep_col_i = fastpyxl.utils.cell.column_index_from_string(dep_c)
        prec_sheet_name, prec_coord = parse_address(prec_key)
        prec_c, prec_r = fastpyxl.utils.cell.coordinate_from_string(prec_coord)
        prec_col_i = fastpyxl.utils.cell.column_index_from_string(prec_c)
        this_chain = is_rr_chain_ref(
            dep_col=dep_c,
            dep_row=dep_r,
            prec_col=prec_c,
            prec_row=prec_r,
            is_absolute_col=ref.is_absolute_col,
            is_absolute_row=ref.is_absolute_row,
            orientation=orientation,
        )
        if this_chain and prec_sheet_name != node.sheet:
            this_chain = False

        this_col_offset = prec_col_i - dep_col_i
        this_row_offset = prec_r - dep_r

        if orientation is Orientation.column:
            advance = dep_r
            this_dep_fixed = dep_c
            this_prec_fixed = prec_c
        else:
            advance = dep_col_i
            this_dep_fixed = dep_r
            this_prec_fixed = prec_r

        if dep_sheet is None:
            dep_sheet = node.sheet
            prec_sheet = prec_sheet_name
            dep_fixed = this_dep_fixed
            prec_fixed = this_prec_fixed
            first_advance = advance
            last_advance = advance
            col_offset = this_col_offset
            row_offset = this_row_offset
            chain = this_chain
        else:
            if node.sheet != dep_sheet or this_dep_fixed != dep_fixed:
                return None
            if prec_sheet_name != prec_sheet or this_prec_fixed != prec_fixed:
                return None
            if this_col_offset != col_offset or this_row_offset != row_offset:
                return None
            if this_chain != chain:
                return None
            assert last_advance is not None
            if advance != last_advance + 1:
                return None
            last_advance = advance

    assert dep_sheet is not None and prec_sheet is not None
    assert dep_fixed is not None and prec_fixed is not None
    assert first_advance is not None and last_advance is not None
    assert col_offset is not None and row_offset is not None

    if orientation is Orientation.column:
        assert isinstance(dep_fixed, str)
        assert isinstance(prec_fixed, str)
        dep_range = RangeRef.column_span(dep_sheet, dep_fixed, first_advance, last_advance)
        prec_range = RangeRef.column_span(
            prec_sheet,
            prec_fixed,
            first_advance + row_offset,
            last_advance + row_offset,
        )
    else:
        assert isinstance(dep_fixed, int)
        assert isinstance(prec_fixed, int)
        first_col = fastpyxl.utils.cell.get_column_letter(first_advance)
        last_col = fastpyxl.utils.cell.get_column_letter(last_advance)
        prec_first_col = fastpyxl.utils.cell.get_column_letter(first_advance + col_offset)
        prec_last_col = fastpyxl.utils.cell.get_column_letter(last_advance + col_offset)
        dep_range = RangeRef.row_span(dep_sheet, dep_fixed, first_col, last_col)
        prec_range = RangeRef.row_span(
            prec_sheet,
            prec_fixed + row_offset,
            prec_first_col,
            prec_last_col,
        )

    kind = PatternKind.rr_chain if chain else PatternKind.rr
    meta = PatternMeta(
        kind=kind,
        col_offset=col_offset,
        row_offset=row_offset,
        orientation=orientation,
    )
    return dep_range, prec_range, meta


def _collect_range_stream(
    graph: DependencyGraph,
    group: list[NodeKey],
    ref_idx: int,
    config: TacoBuildConfig,
    orientation: Orientation,
) -> tuple[RangeRef, RangeRef, PatternMeta] | None:
    dep_sheet: str | None = None
    dep_fixed: str | int | None = None
    first_advance: int | None = None
    last_advance: int | None = None
    pattern: str | None = None
    prec_range: RangeRef | None = None
    meta: PatternMeta | None = None
    col_offset: int | None = None

    for dep_key in group:
        node = graph.get_node(dep_key)
        if node is None or not node.formula:
            return None
        streams = parse_ref_streams(node.formula, default_sheet=node.sheet)
        if ref_idx >= len(streams):
            return None
        stream_item = streams[ref_idx]
        if not isinstance(stream_item, AbsRangeRef):
            return None
        ref = stream_item
        this_pattern = classify_range_pattern(ref)
        if this_pattern not in {"RF", "FR", "FF"}:
            return None
        expected = set(range_ref_to_keys(ref, default_sheet=node.sheet))
        if not expected:
            return None
        if any(not precedent_may_compress(graph, prec, config) for prec in expected):
            return None
        if any(
            _edge_is_excluded_from_pattern_inference(graph, dep_key, prec)
            for prec in expected
            if prec not in graph.get_dependencies(dep_key)
        ):
            return None
        if not _stream_expected_deps(graph, node, streams, ref_idx, expected):
            return None

        sheet, sc, sr, ec, er = _bounding_box(ref, default_sheet=node.sheet)
        cell_prec = RangeRef.rectangle(sheet, sc, sr, ec, er)

        if orientation is Orientation.column:
            advance = node.row
            this_dep_fixed = node.column
        else:
            advance = node.column_index
            this_dep_fixed = node.row

        if dep_sheet is None:
            dep_sheet = node.sheet
            dep_fixed = this_dep_fixed
            first_advance = advance
            last_advance = advance
            pattern = this_pattern
            prec_range = cell_prec
            meta = _meta_for_range_pattern(this_pattern, ref, orientation, node.column, node.row)
            if orientation is Orientation.row and this_pattern == "RF":
                start_col_i = fastpyxl.utils.cell.column_index_from_string(ref.start_col)
                col_offset = start_col_i - node.column_index
        else:
            if node.sheet != dep_sheet or this_dep_fixed != dep_fixed or this_pattern != pattern:
                return None
            assert last_advance is not None and prec_range is not None and meta is not None
            if advance != last_advance + 1:
                return None
            if orientation is Orientation.row and this_pattern == "RF":
                start_col_i = fastpyxl.utils.cell.column_index_from_string(ref.start_col)
                this_col_offset = start_col_i - node.column_index
                if col_offset is not None and this_col_offset != col_offset:
                    return None
            last_advance = advance
            prec_range = _union_ranges(prec_range, cell_prec)

    assert dep_sheet is not None and dep_fixed is not None and prec_range is not None
    assert first_advance is not None and last_advance is not None and meta is not None

    if orientation is Orientation.column:
        assert isinstance(dep_fixed, str)
        dep_range = RangeRef.column_span(dep_sheet, dep_fixed, first_advance, last_advance)
    else:
        assert isinstance(dep_fixed, int)
        first_col = fastpyxl.utils.cell.get_column_letter(first_advance)
        last_col = fastpyxl.utils.cell.get_column_letter(last_advance)
        dep_range = RangeRef.row_span(dep_sheet, dep_fixed, first_col, last_col)
    return dep_range, prec_range, meta


def _meta_for_range_pattern(
    pattern: str,
    ref: AbsRangeRef,
    orientation: Orientation,
    dep_col: str,
    _dep_row: int,
) -> PatternMeta:
    if pattern == "RF":
        col_offset = 0
        if orientation is Orientation.row:
            start_col_i = fastpyxl.utils.cell.column_index_from_string(ref.start_col)
            dep_col_i = fastpyxl.utils.cell.column_index_from_string(dep_col)
            col_offset = start_col_i - dep_col_i
        return PatternMeta(
            kind=PatternKind.rf,
            fixed_tail_col=ref.end_col,
            fixed_tail_row=ref.end_row,
            orientation=orientation,
            col_offset=col_offset,
        )
    if pattern == "FR":
        return PatternMeta(
            kind=PatternKind.fr,
            fixed_head_col=ref.start_col,
            fixed_head_row=ref.start_row,
            orientation=orientation,
        )
    return PatternMeta(kind=PatternKind.ff, orientation=orientation)


def _bounding_box(ref: AbsRangeRef, *, default_sheet: str) -> tuple[str, str, int, str, int]:
    sheet = ref.sheet if ref.sheet is not None else default_sheet
    c1 = fastpyxl.utils.cell.column_index_from_string(ref.start_col)
    c2 = fastpyxl.utils.cell.column_index_from_string(ref.end_col)
    r1, r2 = (
        (ref.start_row, ref.end_row)
        if ref.start_row <= ref.end_row
        else (ref.end_row, ref.start_row)
    )
    clo, chi = (c1, c2) if c1 <= c2 else (c2, c1)
    return (
        sheet,
        fastpyxl.utils.cell.get_column_letter(clo),
        r1,
        fastpyxl.utils.cell.get_column_letter(chi),
        r2,
    )


def _union_ranges(a: RangeRef, b: RangeRef) -> RangeRef:
    if a.sheet != b.sheet:
        raise ValueError(f"cannot union ranges on different sheets: {a!r} vs {b!r}")
    a_c1 = fastpyxl.utils.cell.column_index_from_string(a.min_col)
    a_c2 = fastpyxl.utils.cell.column_index_from_string(a.max_col)
    b_c1 = fastpyxl.utils.cell.column_index_from_string(b.min_col)
    b_c2 = fastpyxl.utils.cell.column_index_from_string(b.max_col)
    return RangeRef.rectangle(
        a.sheet,
        fastpyxl.utils.cell.get_column_letter(min(a_c1, b_c1)),
        min(a.min_row, b.min_row),
        fastpyxl.utils.cell.get_column_letter(max(a_c2, b_c2)),
        max(a.max_row, b.max_row),
    )


def _stream_expected_deps(
    graph: DependencyGraph,
    node,
    streams: list[FormulaRef],
    ref_idx: int,
    expected: set[NodeKey],
) -> bool:
    deps = set(graph.get_dependencies(node.key))
    if len(streams) == 1:
        return expected == deps
    all_expected: set[NodeKey] = set()
    for stream in streams:
        if isinstance(stream, AbsCellRef):
            all_expected.add(abs_ref_to_key(stream, default_sheet=node.sheet))
        else:
            all_expected.update(range_ref_to_keys(stream, default_sheet=node.sheet))
    if all_expected != deps:
        return False
    return expected <= deps
