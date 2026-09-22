"""Fail-closed formula/edge consistency checks for `DependencyGraph`.

Projection primitives (`set_node_formula`, `remove_node`) do not rewire
edges. Codegen and `to_networkx` trust the stored edge set, so a projected
graph whose formulas and edges disagree can emit quietly wrong artifacts.

This module only *detects* disagreement. It never rewrites the graph.

Unlabeled edges count as static formula-ref claims unless the host formula
contains `OFFSET`, `INDIRECT`, or dynamic `INDEX`. Range interiors of those
calls are not treated as static extraction edges. Resolve failures, ranges
over `max_cells`, and whole-column/row refs without `sheet_bounds` are
reported as issues rather than skipped.
"""

from __future__ import annotations

from collections.abc import Callable, Iterator
from dataclasses import dataclass
from enum import StrEnum
from typing import TYPE_CHECKING

from fastpyxl.utils.cell import column_index_from_string, coordinate_from_string

from excel_grapher.core.address_keys import (
    format_key,
    format_range_key,
    normalize_key,
    parse_address,
)
from excel_grapher.core.cell_types import CellTypeEnv
from excel_grapher.core.excel_function_names import normalize_excel_function_name
from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    CellRefNode,
    FormulaParseError,
    FormulaStyle,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    UnaryOpNode,
    WholeColumnNode,
    WholeRowNode,
    render_formula,
    resolve_cell_ref,
    resolve_whole_column_ref,
    resolve_whole_row_ref,
)
from excel_grapher.core.range_shorthand import (
    expand_whole_column_span_deps,
    expand_whole_row_span_deps,
)
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.dynamic_refs import DynamicRefLimits, feasible_choose_option_indices
from excel_grapher.grapher.node import NodeKey
from excel_grapher.grapher.parser import DEFAULT_MAX_RANGE_CELLS, expand_range

if TYPE_CHECKING:
    from excel_grapher.grapher.graph import DependencyGraph
    from excel_grapher.grapher.node import Node

__all__ = [
    "GraphConsistencyError",
    "GraphConsistencyIssue",
    "GraphConsistencyKind",
    "collect_graph_consistency_issues",
    "validate_graph_consistency",
]

_FORMULA_REF_CAUSES = DependencyCause.direct_ref | DependencyCause.static_range
_DYNAMIC_CAUSES = (
    DependencyCause.dynamic_offset
    | DependencyCause.dynamic_indirect
    | DependencyCause.dynamic_index
)


class GraphConsistencyKind(StrEnum):
    """Kind of formula/edge/flag disagreement."""

    missing_formula_edge = "missing_formula_edge"
    extra_formula_edge = "extra_formula_edge"
    edge_from_missing_node = "edge_from_missing_node"
    edge_to_missing_node = "edge_to_missing_node"
    leaf_flag_mismatch = "leaf_flag_mismatch"
    formula_flag_mismatch = "formula_flag_mismatch"
    unresolved_formula_ref = "unresolved_formula_ref"
    range_too_large = "range_too_large"
    unbounded_whole_ref = "unbounded_whole_ref"


@dataclass(frozen=True, slots=True)
class GraphConsistencyIssue:
    """One structured consistency failure.

    Attributes:
        kind: Disagreement class.
        node: Host node for flag issues, when applicable.
        from_key: Edge source, when applicable.
        to_key: Edge target, when applicable.
        message: Human-readable description.
    """

    kind: GraphConsistencyKind
    node: NodeKey | None = None
    from_key: NodeKey | None = None
    to_key: NodeKey | None = None
    message: str = ""


class GraphConsistencyError(ValueError):
    """Raised when `validate_consistency` finds formula/edge disagreement."""

    def __init__(self, issues: tuple[GraphConsistencyIssue, ...]) -> None:
        self.issues = issues
        summary = "; ".join(issue.message for issue in issues) or "unknown disagreement"
        super().__init__(f"graph consistency check failed: {summary}")


def collect_graph_consistency_issues(
    graph: DependencyGraph,
) -> tuple[GraphConsistencyIssue, ...]:
    """Return structured disagreements without mutating `graph`."""
    issues: list[GraphConsistencyIssue] = []
    seen: set[tuple[str, str | None, str | None, str | None]] = set()

    def add(issue: GraphConsistencyIssue) -> None:
        key = (issue.kind.value, issue.node, issue.from_key, issue.to_key)
        if key in seen:
            return
        seen.add(key)
        issues.append(issue)

    nodes = graph._nodes
    sources: set[NodeKey] = set(nodes)
    if graph._staging:
        sources.update(graph._edges)
        sources.update(graph._reverse_edges)
        for dests in graph._reverse_edges.values():
            sources.update(dests)
    for src in sources:
        src_missing = src not in nodes
        for dst in graph._iter_dep_keys(src):
            if src_missing:
                add(
                    GraphConsistencyIssue(
                        kind=GraphConsistencyKind.edge_from_missing_node,
                        from_key=src,
                        to_key=dst,
                        message=f"edge from missing node {src} to {dst}",
                    )
                )
            if dst not in nodes:
                add(
                    GraphConsistencyIssue(
                        kind=GraphConsistencyKind.edge_to_missing_node,
                        from_key=src,
                        to_key=dst,
                        message=f"edge from {src} to missing node {dst}",
                    )
                )

    reverse_dests: set[NodeKey] = set(nodes)
    if graph._staging:
        reverse_dests.update(graph._reverse_edges)
    for dst in reverse_dests:
        for src in graph._iter_dependent_keys(dst):
            if graph._has_stored_edge(src, dst):
                continue
            if src not in nodes:
                add(
                    GraphConsistencyIssue(
                        kind=GraphConsistencyKind.edge_from_missing_node,
                        from_key=src,
                        to_key=dst,
                        message=f"edge from missing node {src} to {dst}",
                    )
                )
            if dst not in nodes:
                add(
                    GraphConsistencyIssue(
                        kind=GraphConsistencyKind.edge_to_missing_node,
                        from_key=src,
                        to_key=dst,
                        message=f"edge from {src} to missing node {dst}",
                    )
                )

    for key, node in nodes.items():
        deps = graph.get_dependencies(key)
        has_outgoing = bool(deps)
        if node.is_leaf == has_outgoing:
            add(
                GraphConsistencyIssue(
                    kind=GraphConsistencyKind.leaf_flag_mismatch,
                    node=key,
                    from_key=key,
                    message=(
                        f"{key} is_leaf={node.is_leaf} but outgoing edges "
                        f"{'exist' if has_outgoing else 'are absent'}"
                    ),
                )
            )
        if not node.has_formula and has_outgoing:
            add(
                GraphConsistencyIssue(
                    kind=GraphConsistencyKind.formula_flag_mismatch,
                    node=key,
                    from_key=key,
                    message=f"{key} has dependency edges but no formula",
                )
            )
        _add_formula_edge_issues(graph, node, key, deps, add)

    issues.sort(
        key=lambda issue: (
            issue.kind.value,
            issue.from_key or issue.node or "",
            issue.to_key or "",
            issue.message,
        )
    )
    return tuple(issues)


def validate_graph_consistency(graph: DependencyGraph) -> None:
    """Raise `GraphConsistencyError` when `graph` is internally inconsistent."""
    issues = collect_graph_consistency_issues(graph)
    if issues:
        raise GraphConsistencyError(issues)


def _claims_formula_ref(provenance: EdgeProvenance | None, *, host_has_dynamic: bool) -> bool:
    """True when an outgoing edge should match a static formula CellRef/range.

    Unlabeled edges (`provenance` missing or empty) count as formula-ref
    claims unless the host formula contains OFFSET/INDIRECT/dynamic INDEX.
    Dynamic-only provenance is never a static formula-ref claim.
    """
    if provenance is not None:
        causes = provenance.causes
        if causes & _FORMULA_REF_CAUSES:
            return True
        if causes & _DYNAMIC_CAUSES:
            return False
    return not host_has_dynamic


def _add_formula_edge_issues(
    graph: DependencyGraph,
    node: Node,
    key: NodeKey,
    deps: frozenset[NodeKey],
    add: Callable[[GraphConsistencyIssue], None],
) -> None:
    if node.formula_ast is None:
        return
    host_has_dynamic = _formula_has_dynamic_call(node.formula_ast)
    expected = _expected_formula_ref_keys(
        graph,
        node,
        key,
        add,
        cell_type_env=graph.cell_type_env,
        choose_limits=_choose_limits(graph),
    )
    for dest in sorted(expected - deps):
        add(
            GraphConsistencyIssue(
                kind=GraphConsistencyKind.missing_formula_edge,
                node=key,
                from_key=key,
                to_key=dest,
                message=f"{key} formula refs {dest} but has no outgoing edge",
            )
        )
    for dest in sorted(deps):
        if dest in expected:
            continue
        provenance = graph.get_edge_attrs(key, dest).provenance
        if not _claims_formula_ref(provenance, host_has_dynamic=host_has_dynamic):
            continue
        add(
            GraphConsistencyIssue(
                kind=GraphConsistencyKind.extra_formula_edge,
                node=key,
                from_key=key,
                to_key=dest,
                message=f"{key} has formula-ref edge to {dest} not present in formula",
            )
        )


def _choose_limits(graph: DependencyGraph) -> DynamicRefLimits:
    raw = getattr(graph, "dynamic_ref_limits", None)
    if raw is None or len(raw) != 3:
        return DynamicRefLimits()
    return DynamicRefLimits(max_branches=raw[0], max_cells=raw[1], max_depth=raw[2])


def _feasible_choose_options(
    index: AstNode,
    option_count: int,
    *,
    anchor: str,
    cell_type_env: CellTypeEnv | None,
    limits: DynamicRefLimits,
) -> list[int] | None:
    """Options extraction would keep for this `CHOOSE` index.

    `None` means the index domain is unknown, matching extraction's
    over-approximation. A render failure also returns `None` so consistency
    does not require fewer edges than a graph that kept every alternative.
    """
    try:
        rendered = render_formula(index, anchor=anchor or None, style=FormulaStyle.A1_ABSOLUTE)
    except (ValueError, FormulaParseError):
        return None
    sheet = ""
    row: int | None = None
    col: int | None = None
    if anchor:
        try:
            sheet, a1 = parse_address(anchor)
            col_letter, row_i = coordinate_from_string(a1)
            row = int(row_i)
            col = column_index_from_string(col_letter)
        except ValueError:
            sheet = ""
            row = None
            col = None
    return feasible_choose_option_indices(
        rendered,
        option_count,
        current_sheet=sheet,
        cell_type_env=cell_type_env,
        limits=limits,
        current_row=row,
        current_col=col,
    )


def _expected_formula_ref_keys(
    graph: DependencyGraph,
    node: Node,
    key: NodeKey,
    add: Callable[[GraphConsistencyIssue], None],
    *,
    cell_type_env: CellTypeEnv | None = None,
    choose_limits: DynamicRefLimits | None = None,
) -> set[NodeKey]:
    ast = node.formula_ast
    if ast is None:
        return set()
    anchor = node.address or key
    expected: set[NodeKey] = set()
    for leaf in _iter_static_address_leaves(
        ast,
        anchor=str(anchor),
        cell_type_env=cell_type_env,
        choose_limits=choose_limits or DynamicRefLimits(),
    ):
        expected.update(
            _keys_for_address_leaf(graph, leaf, anchor=str(anchor), host_key=key, add=add)
        )
    return expected


def _iter_static_address_leaves(
    node: AstNode,
    *,
    dynamic_mask: bool = False,
    anchor: str = "",
    cell_type_env: CellTypeEnv | None = None,
    choose_limits: DynamicRefLimits | None = None,
) -> Iterator[CellRefNode | RangeNode | WholeColumnNode | WholeRowNode]:
    """Yield address leaves extraction would record as static formula refs.

    CellRefs inside OFFSET/INDIRECT/dynamic INDEX still count (argument
    refs). Range and whole-column/row leaves inside those calls are masked,
    matching builder extraction. `CHOOSE` alternatives outside the index
    domain are omitted when that domain is known, matching extraction.
    `CHOOSE` nested inside a dynamic-ref call is not narrowed: argument
    analysis still records every written reference.
    """
    limits = choose_limits or DynamicRefLimits()

    def walk(
        child: AstNode, *, masked: bool
    ) -> Iterator[CellRefNode | RangeNode | WholeColumnNode | WholeRowNode]:
        return _iter_static_address_leaves(
            child,
            dynamic_mask=masked,
            anchor=anchor,
            cell_type_env=cell_type_env,
            choose_limits=limits,
        )

    match node:
        case CellRefNode():
            yield node
        case RangeNode() | WholeColumnNode() | WholeRowNode() if dynamic_mask:
            return
        case RangeNode() | WholeColumnNode() | WholeRowNode():
            yield node
        case BinaryOpNode(left=left, right=right):
            yield from walk(left, masked=dynamic_mask)
            yield from walk(right, masked=dynamic_mask)
        case UnaryOpNode(operand=operand):
            yield from walk(operand, masked=dynamic_mask)
        case FunctionCallNode(name=name, args=args) if (
            not dynamic_mask and normalize_excel_function_name(name) == "CHOOSE" and len(args) >= 2
        ):
            yield from walk(args[0], masked=False)
            live = _feasible_choose_options(
                args[0],
                len(args) - 1,
                anchor=anchor,
                cell_type_env=cell_type_env,
                limits=limits,
            )
            indexes = range(1, len(args)) if live is None else live
            for index in indexes:
                yield from walk(args[index], masked=False)
        case FunctionCallNode(args=args):
            nested = dynamic_mask or _masks_static_ranges(node)
            for arg in args:
                yield from walk(arg, masked=nested)
        case _:
            return


def _formula_has_dynamic_call(node: AstNode) -> bool:
    match node:
        case FunctionCallNode() if _masks_static_ranges(node):
            return True
        case FunctionCallNode(args=args):
            return any(_formula_has_dynamic_call(arg) for arg in args)
        case BinaryOpNode(left=left, right=right):
            return _formula_has_dynamic_call(left) or _formula_has_dynamic_call(right)
        case UnaryOpNode(operand=operand):
            return _formula_has_dynamic_call(operand)
        case _:
            return False


def _masks_static_ranges(node: FunctionCallNode) -> bool:
    name = normalize_excel_function_name(node.name)
    if name in {"OFFSET", "INDIRECT"}:
        return True
    if name == "INDEX":
        return _index_row_col_non_literal(node)
    return False


def _index_row_col_non_literal(node: FunctionCallNode) -> bool:
    if len(node.args) < 2:
        return False
    return any(not _is_numeric_literal(arg) for arg in node.args[1:])


def _is_numeric_literal(node: AstNode) -> bool:
    match node:
        case NumberNode():
            return True
        case UnaryOpNode(op=op, operand=operand) if op in {"+", "-"}:
            return _is_numeric_literal(operand)
        case _:
            return False


def _keys_for_address_leaf(
    graph: DependencyGraph,
    leaf: CellRefNode | RangeNode | WholeColumnNode | WholeRowNode,
    *,
    anchor: str,
    host_key: NodeKey,
    add: Callable[[GraphConsistencyIssue], None],
) -> set[NodeKey]:
    try:
        match leaf:
            case CellRefNode():
                return {normalize_key(resolve_cell_ref(leaf, anchor))}
            case RangeNode():
                return _keys_for_range(graph, leaf, anchor=anchor, host_key=host_key, add=add)
            case WholeColumnNode():
                return _keys_for_whole_column(
                    graph, leaf, anchor=anchor, host_key=host_key, add=add
                )
            case WholeRowNode():
                return _keys_for_whole_row(graph, leaf, anchor=anchor, host_key=host_key, add=add)
    except ValueError as exc:
        add(
            GraphConsistencyIssue(
                kind=GraphConsistencyKind.unresolved_formula_ref,
                node=host_key,
                from_key=host_key,
                message=f"{host_key} formula ref could not be resolved: {exc}",
            )
        )
        return set()
    return set()


def _keys_for_range(
    graph: DependencyGraph,
    leaf: RangeNode,
    *,
    anchor: str,
    host_key: NodeKey,
    add: Callable[[GraphConsistencyIssue], None],
) -> set[NodeKey]:
    start = normalize_key(resolve_cell_ref(leaf.start_ref, anchor))
    end = normalize_key(resolve_cell_ref(leaf.end_ref, anchor))
    start_sheet, start_a1 = parse_address(start)
    end_sheet, end_a1 = parse_address(end)
    if start_sheet != end_sheet:
        add(
            GraphConsistencyIssue(
                kind=GraphConsistencyKind.unresolved_formula_ref,
                node=host_key,
                from_key=host_key,
                message=(f"{host_key} range spans sheets {start_sheet!r} and {end_sheet!r}"),
            )
        )
        return set()
    start_col, start_row = coordinate_from_string(start_a1)
    end_col, end_row = coordinate_from_string(end_a1)
    try:
        pairs = expand_range(
            sheet=start_sheet,
            start_col=start_col,
            start_row=int(start_row),
            end_col=end_col,
            end_row=int(end_row),
            max_cells=DEFAULT_MAX_RANGE_CELLS,
        )
    except ValueError as exc:
        add(
            GraphConsistencyIssue(
                kind=GraphConsistencyKind.range_too_large,
                node=host_key,
                from_key=host_key,
                message=f"{host_key} range exceeds expansion budget: {exc}",
            )
        )
        return set()
    return {format_key(sheet, a1) for sheet, a1 in pairs}


def _keys_for_whole_column(
    graph: DependencyGraph,
    leaf: WholeColumnNode,
    *,
    anchor: str,
    host_key: NodeKey,
    add: Callable[[GraphConsistencyIssue], None],
) -> set[NodeKey]:
    sheet, start_letter, end_letter = resolve_whole_column_ref(leaf, anchor)
    bounds = graph.sheet_bounds
    if not bounds or sheet not in bounds:
        add(
            GraphConsistencyIssue(
                kind=GraphConsistencyKind.unbounded_whole_ref,
                node=host_key,
                from_key=host_key,
                message=(
                    f"{host_key} whole-column ref "
                    f"{format_range_key(sheet, start_letter, end_letter)} "
                    "requires sheet_bounds"
                ),
            )
        )
        return set()
    pairs = expand_whole_column_span_deps(sheet, start_letter, end_letter, bounds)
    return {format_key(dep_sheet, a1) for dep_sheet, a1 in pairs}


def _keys_for_whole_row(
    graph: DependencyGraph,
    leaf: WholeRowNode,
    *,
    anchor: str,
    host_key: NodeKey,
    add: Callable[[GraphConsistencyIssue], None],
) -> set[NodeKey]:
    sheet, start_row, end_row = resolve_whole_row_ref(leaf, anchor)
    bounds = graph.sheet_bounds
    if not bounds or sheet not in bounds:
        add(
            GraphConsistencyIssue(
                kind=GraphConsistencyKind.unbounded_whole_ref,
                node=host_key,
                from_key=host_key,
                message=(
                    f"{host_key} whole-row ref "
                    f"{format_range_key(sheet, str(start_row), str(end_row))} "
                    "requires sheet_bounds"
                ),
            )
        )
        return set()
    pairs = expand_whole_row_span_deps(sheet, start_row, end_row, bounds)
    return {format_key(dep_sheet, a1) for dep_sheet, a1 in pairs}
