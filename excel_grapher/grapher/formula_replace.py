"""Topology-aware formula replace for `DependencyGraph`.

Projection `set_node_formula` parses into `formula_ast` and leaves edges
alone. `replace_node_formula` is the separate durable-edit API: it re-extracts
outgoing edges, guards, and provenance, materializes newly referenced cells when
a workbook is provided, and fails closed when that context is required but
missing.
"""

from __future__ import annotations

import re
from collections.abc import Mapping
from pathlib import Path
from typing import TYPE_CHECKING

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import format_key, normalize_key, parse_address
from excel_grapher.core.formula_ast import (
    parse_formula_text,
    parse_preserving_axes_optional,
)

from .dependency_provenance import DependencyCause, EdgeProvenance
from .guard import GuardExpr
from .node import Node, NodeKey, copy_node
from .parser import (
    DEFAULT_MAX_RANGE_CELLS,
    FormulaNormalizer,
    _find_function_calls_with_spans,
    _split_function_args,
    expand_range,
    expand_range_ref,
    mask_ref_only_function_calls,
    mask_spans,
    parse_dynamic_range_refs_with_spans,
    parse_range_refs_with_spans,
    parse_standalone_cell_refs_with_spans,
)

if TYPE_CHECKING:
    from .dynamic_refs import DynamicRefConfig
    from .graph import DependencyGraph

_NAME_TOKEN_RE = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\b(?!\s*!)")
_CELL_LIKE_TOKEN_RE = re.compile(r"^[A-Za-z]{1,3}\d+$")
_BOOL_TOKENS = frozenset({"TRUE", "FALSE"})
_DYNAMIC_REF_FNS = frozenset({"OFFSET", "INDIRECT", "INDEX"})

__all__ = ["WorkbookContextRequiredError", "replace_node_formula"]


class WorkbookContextRequiredError(ValueError):
    """Raised when `replace_node_formula` needs workbook context that was omitted.

    Named ranges missing from the graph maps, dynamic OFFSET/INDIRECT/INDEX
    that need the book, and newly referenced cells absent from the subgraph all
    require a workbook. The graph is left unchanged.
    """


def replace_node_formula(
    graph: DependencyGraph,
    key: NodeKey,
    formula: str | None,
    normalized_formula: str | None,
    *,
    workbook: str | Path | None = None,
    dynamic_refs: DynamicRefConfig | None = None,
    use_cached_dynamic_refs: bool = False,
    load_values: bool = True,
    capture_dependency_provenance: bool = True,
    max_depth: int = 50,
    expand_ranges: bool = True,
    max_range_cells: int = DEFAULT_MAX_RANGE_CELLS,
) -> None:
    """Replace a node's formula and rewire outgoing topology from extraction.

    See `DependencyGraph.replace_node_formula`.
    """
    nk = normalize_key(key)
    node = graph._nodes.get(nk)
    if node is None:
        raise KeyError(f"Cell {key} not found in graph")

    if formula is None:
        _apply_formula_text(node, None, None)
        _replace_outgoing(graph, nk, ())
        node.is_leaf = True
        graph._invalidate_formula_shapes()
        return

    if workbook is not None:
        _replace_with_workbook(
            graph,
            nk,
            node,
            formula,
            normalized_formula,
            workbook=workbook,
            dynamic_refs=dynamic_refs,
            use_cached_dynamic_refs=use_cached_dynamic_refs,
            load_values=load_values,
            capture_dependency_provenance=capture_dependency_provenance,
            max_depth=max_depth,
            expand_ranges=expand_ranges,
            max_range_cells=max_range_cells,
        )
        return

    edges = _graph_only_edges(
        graph,
        node,
        nk,
        formula,
        max_range_cells=max_range_cells,
        expand_ranges=expand_ranges,
    )
    _apply_formula_text(
        node,
        formula,
        normalized_formula,
        named_ranges=graph.named_ranges,
        named_range_ranges=graph.named_range_ranges,
    )
    _replace_outgoing(graph, nk, edges)
    node.is_leaf = not bool(graph._edges.get(nk))
    graph._invalidate_formula_shapes()


def _apply_formula_text(
    node: Node,
    formula: str | None,
    normalized_formula: str | None,
    *,
    named_ranges: dict[str, tuple[str, str]] | None = None,
    named_range_ranges: dict[str, tuple[str, str, str]] | None = None,
) -> None:
    node.formula = formula
    ast = None
    if formula is not None:
        ast = parse_preserving_axes_optional(
            formula,
            anchor=node.address or node.key,
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
        )
    if ast is None:
        ast = parse_formula_text(
            normalized_formula,
            anchor=node.address or node.key,
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
        )
    node.formula_ast = ast
    if ast is not None:
        node._unparseable_formula = None
    else:
        node._unparseable_formula = normalized_formula


def _replace_outgoing(
    graph: DependencyGraph,
    host: NodeKey,
    edges: tuple[tuple[NodeKey, GuardExpr | None, EdgeProvenance | None], ...],
) -> None:
    for dep in list(graph._edges.get(host, set())):
        graph._remove_edge(host, dep)
    for dep_key, guard, provenance in edges:
        graph.add_edge(host, dep_key, guard=guard, provenance=provenance)


def _replace_with_workbook(
    graph: DependencyGraph,
    nk: NodeKey,
    node: Node,
    formula: str,
    normalized_formula: str | None,
    *,
    workbook: str | Path,
    dynamic_refs: DynamicRefConfig | None,
    use_cached_dynamic_refs: bool,
    load_values: bool,
    capture_dependency_provenance: bool,
    max_depth: int,
    expand_ranges: bool,
    max_range_cells: int,
) -> None:
    from excel_grapher.grapher.builder import create_dependency_graph

    override = formula if formula.startswith("=") else f"={formula}"
    extracted = create_dependency_graph(
        workbook,
        [nk],
        formula_overrides={nk: override},
        dynamic_refs=dynamic_refs,
        use_cached_dynamic_refs=use_cached_dynamic_refs,
        load_values=load_values,
        capture_dependency_provenance=capture_dependency_provenance,
        max_depth=max_depth,
        expand_ranges=expand_ranges,
        max_range_cells=max_range_cells,
    )
    extracted_host = extracted.get_node(nk)
    if extracted_host is None:
        raise WorkbookContextRequiredError(f"workbook extraction did not produce host cell {nk}")

    new_deps = extracted.get_dependencies(nk)
    missing_after = [dep for dep in sorted(new_deps) if dep not in extracted]
    if missing_after:
        raise WorkbookContextRequiredError(
            f"{nk} formula refs missing subgraph cells: {', '.join(missing_after)}"
        )

    _apply_formula_text(
        node,
        formula,
        normalized_formula,
        named_ranges=extracted.named_ranges or graph.named_ranges,
        named_range_ranges=extracted.named_range_ranges or graph.named_range_ranges,
    )
    _merge_extracted_subgraph(graph, nk, extracted)
    node.is_leaf = not bool(graph._edges.get(nk))
    graph._invalidate_formula_shapes()


def _merge_extracted_subgraph(
    graph: DependencyGraph,
    host: NodeKey,
    extracted: DependencyGraph,
) -> None:
    existing = set(graph._nodes)
    new_keys = [key for key in extracted if key not in existing]

    for dep in list(graph._edges.get(host, set())):
        graph._remove_edge(host, dep)

    for key in new_keys:
        graph.add_node(copy_node(extracted._nodes[key]))

    for dep in extracted.get_dependencies(host):
        attrs = extracted.get_edge_attrs(host, dep)
        graph.add_edge(host, dep, guard=attrs.guard, provenance=attrs.provenance)

    for key in new_keys:
        for dep in extracted.get_dependencies(key):
            attrs = extracted.get_edge_attrs(key, dep)
            graph.add_edge(key, dep, guard=attrs.guard, provenance=attrs.provenance)

    if extracted.named_ranges:
        merged = dict(graph.named_ranges or {})
        merged.update(extracted.named_ranges)
        graph.named_ranges = merged
    if extracted.named_range_ranges:
        merged_ranges = dict(graph.named_range_ranges or {})
        merged_ranges.update(extracted.named_range_ranges)
        graph.named_range_ranges = merged_ranges
    if extracted.sheet_bounds:
        merged_bounds = dict(graph.sheet_bounds or {})
        merged_bounds.update(extracted.sheet_bounds)
        graph.sheet_bounds = merged_bounds
    if graph.sheet_order is None and extracted.sheet_order:
        graph.sheet_order = list(extracted.sheet_order)


def _graph_only_edges(
    graph: DependencyGraph,
    node: Node,
    nk: NodeKey,
    formula: str,
    *,
    max_range_cells: int,
    expand_ranges: bool,
) -> tuple[tuple[NodeKey, GuardExpr | None, EdgeProvenance | None], ...]:
    sheet = node.sheet or parse_address(nk)[0]
    a1 = f"{node.column}{node.row}"
    named_ranges = dict(graph.named_ranges or {})
    named_range_ranges = dict(graph.named_range_ranges or {})
    normalizer = FormulaNormalizer(named_ranges, named_range_ranges)
    text = formula if formula.startswith("=") else f"={formula}"
    normalized = normalizer.normalize(text, sheet)

    dynamic_reason = _dynamic_ref_requires_workbook(
        normalized,
        sheet,
        a1,
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
        normalizer=normalizer,
    )
    if dynamic_reason is not None:
        raise WorkbookContextRequiredError(dynamic_reason)

    unknown = _unknown_defined_name(normalized, named_ranges, named_range_ranges)
    if unknown is not None:
        raise WorkbookContextRequiredError(
            f"{nk} formula references defined name {unknown!r} that is not "
            "on the graph; pass workbook= to resolve named ranges"
        )

    deps, whole_ref_reason = _static_dep_keys(
        normalized,
        current_sheet=sheet,
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
        sheet_bounds=graph.sheet_bounds,
        max_range_cells=max_range_cells,
        expand_ranges=expand_ranges,
    )
    if whole_ref_reason is not None:
        raise WorkbookContextRequiredError(whole_ref_reason)

    missing = sorted(dep for dep in deps if dep not in graph._nodes)
    if missing:
        raise WorkbookContextRequiredError(
            f"{nk} formula refs missing subgraph cells: {', '.join(missing)}; "
            "pass workbook= to materialize newly referenced cells"
        )

    edges: list[tuple[NodeKey, GuardExpr | None, EdgeProvenance | None]] = []
    for dep_key, cause in deps.items():
        edges.append((dep_key, None, EdgeProvenance(causes=cause)))
    return tuple(edges)


def _dynamic_ref_requires_workbook(
    formula: str,
    current_sheet: str,
    current_a1: str,
    *,
    named_ranges: dict[str, tuple[str, str]],
    named_range_ranges: dict[str, tuple[str, str, str]],
    normalizer: FormulaNormalizer,
) -> str | None:
    calls = _find_function_calls_with_spans(formula, _DYNAMIC_REF_FNS)
    if not calls:
        return None
    try:
        static_calls = parse_dynamic_range_refs_with_spans(
            formula,
            current_sheet=current_sheet,
            current_cell_a1=current_a1,
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
            normalizer=normalizer,
            value_resolver=None,
        )
        static_ok = {span for _s, _e, span, _a in static_calls}
    except ValueError:
        static_ok = set()

    needing: list[str] = []
    for fn_name, inner, span in calls:
        if fn_name == "INDEX" and not _index_row_col_non_literal(inner):
            continue
        if span in static_ok:
            continue
        needing.append(fn_name)
    if not needing:
        return None
    names = "/".join(sorted(set(needing)))
    return (
        f"formula contains {names} that require dynamic resolution; "
        "pass workbook= (and use_cached_dynamic_refs=True or dynamic_refs=...) "
        "to re-extract"
    )


def _index_row_col_non_literal(inner: str) -> bool:
    idx_args = _split_function_args(inner)
    if idx_args is None or len(idx_args) < 2:
        return False
    for j, idx_arg in enumerate(idx_args):
        if j == 0:
            continue
        try:
            float(idx_arg.strip())
        except ValueError:
            return True
    return False


def _unknown_defined_name(
    formula: str,
    named_ranges: Mapping[str, tuple[str, str]],
    named_range_ranges: Mapping[str, tuple[str, str, str]],
) -> str | None:
    for match in _NAME_TOKEN_RE.finditer(formula):
        token = match.group(1)
        rest = formula[match.end() :].lstrip()
        if rest.startswith("("):
            continue
        if token.upper() in _BOOL_TOKENS:
            continue
        if _CELL_LIKE_TOKEN_RE.fullmatch(token):
            continue
        if token in named_ranges or token in named_range_ranges:
            continue
        return token
    return None


def _static_dep_keys(
    formula: str,
    *,
    current_sheet: str,
    named_ranges: dict[str, tuple[str, str]],
    named_range_ranges: dict[str, tuple[str, str, str]],
    sheet_bounds: dict[str, tuple[int, int]] | None,
    max_range_cells: int,
    expand_ranges: bool,
) -> tuple[dict[NodeKey, DependencyCause], str | None]:
    deps: dict[NodeKey, DependencyCause] = {}
    masked = formula
    dyn_spans: list[tuple[int, int]] = []
    try:
        static_dyn = parse_dynamic_range_refs_with_spans(
            formula,
            current_sheet=current_sheet,
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
            value_resolver=None,
        )
    except ValueError:
        static_dyn = []
    for start, end, span, arg_refs in static_dyn:
        dyn_spans.append(span)
        sheet = start.sheet if start.sheet is not None else current_sheet
        cause = DependencyCause.dynamic_offset
        try:
            for dep_sheet, dep_a1 in expand_range_ref(
                start=start,
                end=end,
                default_sheet=sheet,
                max_cells=max_range_cells,
                sheet_bounds=sheet_bounds,
            ):
                deps[format_key(dep_sheet, dep_a1)] = cause
        except ValueError as exc:
            return {}, f"dynamic range expansion failed: {exc}"
        for ref in arg_refs:
            arg_sheet = ref.sheet if ref.sheet is not None else current_sheet
            deps[format_key(arg_sheet, f"{ref.column}{ref.row}")] = cause
    masked = mask_spans(masked, dyn_spans)
    masked = mask_ref_only_function_calls(masked)

    if expand_ranges:
        for start, end, _span in parse_range_refs_with_spans(masked):
            sheet = start.sheet if start.sheet is not None else current_sheet
            if start.range_kind in {"whole_column", "whole_row"} and (
                not sheet_bounds or sheet not in sheet_bounds
            ):
                return {}, (f"whole-column/row ref on {sheet!r} requires sheet_bounds or workbook=")
            try:
                for dep_sheet, dep_a1 in expand_range_ref(
                    start=start,
                    end=end,
                    default_sheet=sheet,
                    max_cells=max_range_cells,
                    sheet_bounds=sheet_bounds,
                ):
                    deps[format_key(dep_sheet, dep_a1)] = DependencyCause.static_range
            except ValueError as exc:
                return {}, f"range expansion failed: {exc}"

    for ref, _span in parse_standalone_cell_refs_with_spans(masked):
        sh = ref.sheet if ref.sheet is not None else current_sheet
        key = format_key(sh, f"{ref.column}{ref.row}")
        deps.setdefault(key, DependencyCause.direct_ref)

    for match in _NAME_TOKEN_RE.finditer(masked):
        token = match.group(1)
        rest = masked[match.end() :].lstrip()
        if rest.startswith("(") or token.upper() in _BOOL_TOKENS:
            continue
        if _CELL_LIKE_TOKEN_RE.fullmatch(token):
            continue
        resolved = named_ranges.get(token)
        if resolved is not None:
            deps.setdefault(format_key(resolved[0], resolved[1]), DependencyCause.direct_ref)
            continue
        resolved_range = named_range_ranges.get(token)
        if resolved_range is not None and expand_ranges:
            rsh, start_a1, end_a1 = resolved_range
            start_col, start_row = fastpyxl.utils.cell.coordinate_from_string(start_a1)
            end_col, end_row = fastpyxl.utils.cell.coordinate_from_string(end_a1)
            for dep_sheet, dep_a1 in expand_range(
                sheet=rsh,
                start_col=start_col,
                start_row=int(start_row),
                end_col=end_col,
                end_row=int(end_row),
                max_cells=max_range_cells,
            ):
                deps.setdefault(format_key(dep_sheet, dep_a1), DependencyCause.static_range)

    return deps, None
