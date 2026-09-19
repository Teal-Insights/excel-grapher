"""Sheet-level quotient of a cell dependency graph.

One node per worksheet. Cell-level edges between the same pair of sheets
collapse into a single directed edge whose `edge_count` is the number of
underlying cell edges. Intra-sheet edges are omitted unless requested as
self-loops. Edge direction matches `DependencyGraph`: `A -> B` means sheet
`A` depends on sheet `B`.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass

from .graph import GraphReadView


@dataclass(frozen=True, slots=True)
class SheetGraphNode:
    """One worksheet in a sheet-level quotient of a dependency graph."""

    name: str
    cell_count: int
    formula_count: int
    leaf_count: int


@dataclass(frozen=True, slots=True)
class SheetGraphEdge:
    """Consolidated directed dependence from `source` onto `target`.

    `source -> target` means cells on `source` depend on cells on `target`.
    `edge_count` is the number of merged cell-level edges; `guarded_count` is
    how many of those carried a guard.
    """

    source: str
    target: str
    edge_count: int
    guarded_count: int


@dataclass(frozen=True, slots=True)
class SheetGraph:
    """Sheet-level quotient: one node per sheet, cross-sheet edges merged."""

    nodes: tuple[SheetGraphNode, ...]
    edges: tuple[SheetGraphEdge, ...]


def _sheet_sort_key(sheet_order: list[str] | None) -> tuple[dict[str, int], int]:
    """Return `(rank, fallback)` so unknown sheets sort after `sheet_order`."""
    if not sheet_order:
        return {}, 0
    rank = {name: idx for idx, name in enumerate(sheet_order)}
    return rank, len(rank)


def _ordered_sheet_names(names: set[str], sheet_order: list[str] | None) -> list[str]:
    """Order sheet names by workbook `sheet_order`, then by name."""
    rank, fallback = _sheet_sort_key(sheet_order)
    return sorted(names, key=lambda name: (rank.get(name, fallback), name))


def to_sheet_graph(
    graph: GraphReadView,
    *,
    include_self_loops: bool = False,
) -> SheetGraph:
    """Collapse a cell graph to one node per sheet with consolidated edges.

    Sheets that appear as graph nodes are included even when they have no
    cross-sheet edges. Dangling edge endpoints that do not resolve to a graph
    node are skipped, matching `to_graphviz`.

    Args:
        graph: Cell-level dependency graph or read view.
        include_self_loops: If `True`, keep intra-sheet cell edges as a single
            self-loop per sheet. If `False` (default), drop them.

    Returns:
        Immutable sheet graph. Node order follows `graph.sheet_order` when
        present; otherwise names sort lexicographically. Edges follow the
        same sheet order on `(source, target)`.
    """
    cell_count: dict[str, int] = defaultdict(int)
    formula_count: dict[str, int] = defaultdict(int)
    leaf_count: dict[str, int] = defaultdict(int)
    sheets_by_key: dict[str, str] = {}

    for key in graph:
        node = graph.get_node(key)
        if node is None or not node.sheet:
            continue
        sheet = node.sheet
        sheets_by_key[key] = sheet
        cell_count[sheet] += 1
        if node.has_formula:
            formula_count[sheet] += 1
        if node.is_leaf:
            leaf_count[sheet] += 1

    pair_count: dict[tuple[str, str], int] = defaultdict(int)
    pair_guarded: dict[tuple[str, str], int] = defaultdict(int)

    for source_key, source_sheet in sheets_by_key.items():
        for dep in graph.get_dependencies(source_key):
            resolved = graph.resolve_endpoint(dep)
            if resolved is None:
                continue
            target_sheet = sheets_by_key.get(resolved)
            if target_sheet is None:
                continue
            if source_sheet == target_sheet and not include_self_loops:
                continue
            pair = (source_sheet, target_sheet)
            pair_count[pair] += 1
            if graph.is_guarded(source_key, dep):
                pair_guarded[pair] += 1

    ordered_names = _ordered_sheet_names(set(cell_count), graph.sheet_order)
    nodes = tuple(
        SheetGraphNode(
            name=name,
            cell_count=cell_count[name],
            formula_count=formula_count[name],
            leaf_count=leaf_count[name],
        )
        for name in ordered_names
    )
    rank, fallback = _sheet_sort_key(graph.sheet_order)
    edges = tuple(
        SheetGraphEdge(
            source=source,
            target=target,
            edge_count=pair_count[(source, target)],
            guarded_count=pair_guarded[(source, target)],
        )
        for source, target in sorted(
            pair_count,
            key=lambda pair: (
                rank.get(pair[0], fallback),
                pair[0],
                rank.get(pair[1], fallback),
                pair[1],
            ),
        )
    )
    return SheetGraph(nodes=nodes, edges=edges)


__all__ = [
    "SheetGraph",
    "SheetGraphEdge",
    "SheetGraphNode",
    "to_sheet_graph",
]
