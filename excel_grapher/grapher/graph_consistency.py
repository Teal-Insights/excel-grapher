"""Fail-closed formula/edge consistency checks for `DependencyGraph`.

Projection primitives (`set_node_formula`, `remove_node`) do not rewire
edges. Codegen and `to_networkx` trust the stored edge set, so a projected
graph whose formulas and edges disagree can emit quietly wrong artifacts.

This module only *detects* disagreement. It never rewrites the graph.
"""

from __future__ import annotations

from collections.abc import Callable, Iterator
from dataclasses import dataclass
from enum import StrEnum
from typing import TYPE_CHECKING

from fastpyxl.utils.cell import column_index_from_string, coordinate_from_string

from excel_grapher.core.address_keys import format_key, normalize_key, parse_address
from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    CellRefNode,
    FunctionCallNode,
    RangeNode,
    UnaryOpNode,
    WholeColumnNode,
    WholeRowNode,
    resolve_cell_ref,
    resolve_whole_column_ref,
    resolve_whole_row_ref,
)
from excel_grapher.core.range_shorthand import (
    expand_whole_column_deps,
    expand_whole_row_deps,
)
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
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


class GraphConsistencyKind(StrEnum):
    """Kind of formula/edge/flag disagreement."""

    missing_formula_edge = "missing_formula_edge"
    extra_formula_edge = "extra_formula_edge"
    edge_from_missing_node = "edge_from_missing_node"
    edge_to_missing_node = "edge_to_missing_node"
    leaf_flag_mismatch = "leaf_flag_mismatch"
    formula_flag_mismatch = "formula_flag_mismatch"


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
    for src, dests in graph._edges.items():
        src_missing = src not in nodes
        for dst in dests:
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

    for dst, srcs in graph._reverse_edges.items():
        for src in srcs:
            if dst in graph._edges.get(src, ()):
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


def _claims_formula_ref(provenance: EdgeProvenance | None) -> bool:
    """True when an edge is labeled as a formula CellRef or static range."""
    if provenance is None:
        return False
    return bool(provenance.causes & _FORMULA_REF_CAUSES)


def _add_formula_edge_issues(
    graph: DependencyGraph,
    node: Node,
    key: NodeKey,
    deps: frozenset[NodeKey],
    add: Callable[[GraphConsistencyIssue], None],
) -> None:
    if node.formula_ast is None:
        return
    expected = _expected_formula_ref_keys(graph, node, key)
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
        if not _claims_formula_ref(provenance):
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


def _expected_formula_ref_keys(
    graph: DependencyGraph,
    node: Node,
    key: NodeKey,
) -> set[NodeKey]:
    ast = node.formula_ast
    if ast is None:
        return set()
    anchor = node.address or key
    expected: set[NodeKey] = set()
    for leaf in _iter_address_leaves(ast):
        expected.update(_keys_for_address_leaf(graph, leaf, anchor=str(anchor)))
    return expected


def _iter_address_leaves(
    node: AstNode,
) -> Iterator[CellRefNode | RangeNode | WholeColumnNode | WholeRowNode]:
    match node:
        case CellRefNode() | RangeNode() | WholeColumnNode() | WholeRowNode():
            yield node
        case BinaryOpNode(left=left, right=right):
            yield from _iter_address_leaves(left)
            yield from _iter_address_leaves(right)
        case UnaryOpNode(operand=operand):
            yield from _iter_address_leaves(operand)
        case FunctionCallNode(args=args):
            for arg in args:
                yield from _iter_address_leaves(arg)
        case _:
            return


def _keys_for_address_leaf(
    graph: DependencyGraph,
    leaf: CellRefNode | RangeNode | WholeColumnNode | WholeRowNode,
    *,
    anchor: str,
) -> set[NodeKey]:
    try:
        match leaf:
            case CellRefNode():
                return {normalize_key(resolve_cell_ref(leaf, anchor))}
            case RangeNode():
                return _keys_for_range(graph, leaf, anchor=anchor)
            case WholeColumnNode():
                return _keys_for_whole_column(graph, leaf, anchor=anchor)
            case WholeRowNode():
                return _keys_for_whole_row(graph, leaf, anchor=anchor)
    except ValueError:
        return set()
    return set()


def _keys_for_range(
    graph: DependencyGraph,
    leaf: RangeNode,
    *,
    anchor: str,
) -> set[NodeKey]:
    start = normalize_key(resolve_cell_ref(leaf.start_ref, anchor))
    end = normalize_key(resolve_cell_ref(leaf.end_ref, anchor))
    start_sheet, start_a1 = parse_address(start)
    end_sheet, end_a1 = parse_address(end)
    if start_sheet != end_sheet:
        return {start, end}
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
    except ValueError:
        return _graph_nodes_in_rectangle(
            graph,
            sheet=start_sheet,
            start_col=start_col,
            start_row=int(start_row),
            end_col=end_col,
            end_row=int(end_row),
        )
    return {format_key(sheet, a1) for sheet, a1 in pairs}


def _graph_nodes_in_rectangle(
    graph: DependencyGraph,
    *,
    sheet: str,
    start_col: str,
    start_row: int,
    end_col: str,
    end_row: int,
) -> set[NodeKey]:
    c1 = column_index_from_string(start_col)
    c2 = column_index_from_string(end_col)
    clo, chi = (c1, c2) if c1 <= c2 else (c2, c1)
    rlo, rhi = (start_row, end_row) if start_row <= end_row else (end_row, start_row)
    out: set[NodeKey] = set()
    for key, node in graph._nodes.items():
        if node.sheet != sheet or node.column is None or node.row is None:
            continue
        col_i = column_index_from_string(node.column)
        if clo <= col_i <= chi and rlo <= int(node.row) <= rhi:
            out.add(key)
    return out


def _keys_for_whole_column(
    graph: DependencyGraph,
    leaf: WholeColumnNode,
    *,
    anchor: str,
) -> set[NodeKey]:
    sheet, letter = resolve_whole_column_ref(leaf, anchor)
    bounds = graph.sheet_bounds
    if bounds and sheet in bounds:
        pairs = expand_whole_column_deps(sheet, letter, bounds)
        return {format_key(dep_sheet, a1) for dep_sheet, a1 in pairs}
    return {
        key for key, node in graph._nodes.items() if node.sheet == sheet and node.column == letter
    }


def _keys_for_whole_row(
    graph: DependencyGraph,
    leaf: WholeRowNode,
    *,
    anchor: str,
) -> set[NodeKey]:
    sheet, row = resolve_whole_row_ref(leaf, anchor)
    bounds = graph.sheet_bounds
    if bounds and sheet in bounds:
        pairs = expand_whole_row_deps(sheet, row, bounds)
        return {format_key(dep_sheet, a1) for dep_sheet, a1 in pairs}
    return {key for key, node in graph._nodes.items() if node.sheet == sheet and node.row == row}
