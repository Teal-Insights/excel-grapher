"""Series-level quotient of a cell dependency graph.

One node per bound series. Cell-level edges between the same pair of series
collapse into a single directed edge whose `edge_count` is the number of
underlying cell edges. Intra-series edges are omitted unless requested as
self-loops.

Unlike statement-level and cell-level visualization, this quotient does not
compute a schedule ranking. Contracting cells into series can invent cycles
even when the cell graph is a DAG; those cycles are kept.
"""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Sequence
from dataclasses import dataclass
from typing import Any

from excel_grapher.core.address_keys import as_canonical, parse_address
from excel_grapher.semantic_model import CatalogOccupancy as SeriesGraphOccupancy

from .graph import DependencyGraph, GraphReadView

MIXED_SHEET = "mixed"


@dataclass(frozen=True, slots=True)
class SeriesGraphNode:
    """One bound series in a series-level quotient of a dependency graph."""

    series_id: str
    direction: str
    layout: str
    cell_count: int
    formula_count: int
    leaf_count: int
    sheet: str


@dataclass(frozen=True, slots=True)
class SeriesGraphEdge:
    """Consolidated directed dependence from `source` onto `target`.

    `source -> target` means cells in `source` depend on cells in `target`.
    `edge_count` is the number of merged cell-level edges; `guarded_count` is
    how many of those carried a guard.
    """

    source: str
    target: str
    edge_count: int
    guarded_count: int


@dataclass(frozen=True, slots=True)
class SeriesGraph:
    """Series-level quotient: one node per series, cross-series edges merged."""

    nodes: tuple[SeriesGraphNode, ...]
    edges: tuple[SeriesGraphEdge, ...]


def _series_sheet(addresses: Sequence[str]) -> str:
    """Return the shared worksheet name, or `mixed` when sheets disagree."""
    sheets: set[str] = set()
    for address in addresses:
        sheets.add(parse_address(address)[0])
        if len(sheets) > 1:
            return MIXED_SHEET
    if len(sheets) != 1:
        return MIXED_SHEET
    return next(iter(sheets))


def _series_sort_key(order: Sequence[str]) -> tuple[dict[str, int], int]:
    """Return `(rank, fallback)` so unknown series sort after `order`."""
    rank = {name: idx for idx, name in enumerate(order)}
    return rank, len(rank)


def _ordered_series_ids(names: set[str], order: Sequence[str]) -> list[str]:
    """Order series ids by occupancy `order`, then by name."""
    rank, fallback = _series_sort_key(order)
    return sorted(names, key=lambda name: (rank.get(name, fallback), name))


def _series_meta(occupancy: SeriesGraphOccupancy, series_id: str) -> tuple[str, str]:
    """Return `(direction, layout)` for `series_id`."""
    series: Any = occupancy.get(series_id)
    direction = getattr(series, "direction", None)
    layout = getattr(series, "layout", None)
    return (
        str(direction) if direction else "internal",
        str(layout) if layout else "series",
    )


def to_series_graph(
    graph: DependencyGraph | GraphReadView,
    occupancy: SeriesGraphOccupancy,
    *,
    include_self_loops: bool = False,
) -> SeriesGraph:
    """Collapse a cell graph to one node per bound series with consolidated edges.

    Series that own graph nodes are included even when they have no
    cross-series edges. Unbound cells and dangling endpoints that do not
    resolve to a graph node are skipped, matching `to_sheet_graph`.

    Args:
        graph: Cell-level dependency graph or read view.
        occupancy: Cell-to-series mapping. A `SeriesCatalog` satisfies
            `CatalogOccupancy` (`SeriesGraphOccupancy`).
        include_self_loops: If `True`, keep intra-series cell edges as a single
            self-loop per series. If `False` (default), drop them.

    Returns:
        Immutable series graph. Node order follows `occupancy.order` when
        present; otherwise ids sort lexicographically. Edges follow the same
        order on `(source, target)`.
    """
    cell_count: dict[str, int] = defaultdict(int)
    formula_count: dict[str, int] = defaultdict(int)
    leaf_count: dict[str, int] = defaultdict(int)
    series_by_key: dict[str, str] = {}
    owned_keys: dict[str, list[str]] = defaultdict(list)
    occupancy_order = tuple(getattr(occupancy, "order", ()))

    for key in graph:
        node = graph.get_node(key)
        if node is None:
            continue
        series_id = occupancy.series_id_for(as_canonical(key))
        if series_id is None:
            continue
        series_by_key[key] = series_id
        owned_keys[series_id].append(key)
        cell_count[series_id] += 1
        if node.has_formula:
            formula_count[series_id] += 1
        if node.is_leaf:
            leaf_count[series_id] += 1

    pair_count: dict[tuple[str, str], int] = defaultdict(int)
    pair_guarded: dict[tuple[str, str], int] = defaultdict(int)

    for source_key, source_series in series_by_key.items():
        for dep in graph.get_dependencies(source_key):
            resolved = graph.resolve_endpoint(dep)
            if resolved is None:
                continue
            target_series = series_by_key.get(resolved)
            if target_series is None:
                continue
            if source_series == target_series and not include_self_loops:
                continue
            pair = (source_series, target_series)
            pair_count[pair] += 1
            if graph.is_guarded(source_key, dep):
                pair_guarded[pair] += 1

    ordered_ids = _ordered_series_ids(set(cell_count), occupancy_order)
    nodes: list[SeriesGraphNode] = []
    for series_id in ordered_ids:
        direction, layout = _series_meta(occupancy, series_id)
        nodes.append(
            SeriesGraphNode(
                series_id=series_id,
                direction=direction,
                layout=layout,
                cell_count=cell_count[series_id],
                formula_count=formula_count[series_id],
                leaf_count=leaf_count[series_id],
                sheet=_series_sheet(owned_keys[series_id]),
            )
        )
    rank, fallback = _series_sort_key(occupancy_order)
    edges = tuple(
        SeriesGraphEdge(
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
    return SeriesGraph(nodes=tuple(nodes), edges=edges)


__all__ = [
    "MIXED_SHEET",
    "SeriesGraph",
    "SeriesGraphEdge",
    "SeriesGraphNode",
    "SeriesGraphOccupancy",
    "to_series_graph",
]
