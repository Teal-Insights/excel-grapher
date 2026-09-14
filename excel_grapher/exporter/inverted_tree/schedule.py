"""Statement-graph scheduling: condensation and distance-zero legality.

A bound series is a statement. Excel's graph is over instances. Contracting
statements invents cycles that do not exist at cell grain (#603). The legality
test is: condense, drop lexicographically positive-distance edges, and require
the distance-zero residual to be a DAG per outer-key partition
(Allen–Kennedy / Lustre causality).
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import TYPE_CHECKING

from excel_grapher.exporter.inverted_tree.catalog import (
    SeriesCatalog,
    schedule_axis_coord,
    schedule_partition,
)
from excel_grapher.exporter.inverted_tree.deps import DependenceEdge, collect_series_edges
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.series_bindings.types import Scalar

if TYPE_CHECKING:
    from collections.abc import Mapping

    from excel_grapher.exporter.inverted_tree.deps import SeriesDeps
    from excel_grapher.grapher.graph import DependencyGraph


def scan_function_name(scc: tuple[str, ...]) -> str:
    """Return the internals helper name for a recurrence group, after its first member."""
    return "scan_" + scc[0]


def scc_external_params(
    scc: tuple[str, ...],
    deps: Mapping[str, SeriesDeps],
    catalog_order: Sequence[str],
) -> tuple[str, ...]:
    """Return first-level params of `scc` that live outside the component."""
    members = set(scc)
    ids: set[str] = set()
    for series_id in scc:
        info = deps.get(series_id)
        if info is None:
            continue
        for param_id in info.param_ids:
            if param_id not in members:
                ids.add(param_id)
    return tuple(sid for sid in catalog_order if sid in ids)


def tarjan_series_sccs(
    series_ids: Sequence[str],
    deps: Mapping[str, SeriesDeps],
) -> list[tuple[str, ...]]:
    """Return series SCCs in topological order (dependencies first).

    Members of each SCC follow `series_ids` order (bindings order). This step
    never raises: the condensation of any directed graph is a DAG.
    """
    selected = set(series_ids)
    adj: dict[str, list[str]] = {sid: [] for sid in series_ids}
    for sid in series_ids:
        info = deps.get(sid)
        if info is None:
            continue
        for param_id in info.param_ids:
            if param_id in selected and param_id != sid:
                adj[sid].append(param_id)

    index = 0
    stack: list[str] = []
    on_stack: set[str] = set()
    indices: dict[str, int] = {}
    lowlink: dict[str, int] = {}
    sccs_rev: list[set[str]] = []

    def strongconnect(v: str) -> None:
        nonlocal index
        indices[v] = index
        lowlink[v] = index
        index += 1
        stack.append(v)
        on_stack.add(v)
        for w in adj[v]:
            if w not in indices:
                strongconnect(w)
                lowlink[v] = min(lowlink[v], lowlink[w])
            elif w in on_stack:
                lowlink[v] = min(lowlink[v], indices[w])
        if lowlink[v] == indices[v]:
            component: set[str] = set()
            while True:
                w = stack.pop()
                on_stack.remove(w)
                component.add(w)
                if w == v:
                    break
            sccs_rev.append(component)

    for sid in series_ids:
        if sid not in indices:
            strongconnect(sid)

    ordered: list[tuple[str, ...]] = []
    for component in sccs_rev:
        ordered.append(tuple(sid for sid in series_ids if sid in component))
    return ordered


def collect_dependence_edges(
    catalog: SeriesCatalog,
    graph: DependencyGraph | None,
    series_ids: Sequence[str],
    *,
    edges: Sequence[DependenceEdge] | None = None,
) -> tuple[DependenceEdge, ...]:
    """Collect instance-level cell-ref edges among `series_ids`.

    `whole` and `dynamic` accesses have no fixed distance and are omitted so
    the residual test stays over concrete instance reads. Pass `edges` to
    filter an already-walked catalog instead of re-visiting formula ASTs.
    """
    wanted = set(series_ids)
    if edges is None:
        if graph is None:
            raise TypeError("collect_dependence_edges requires edges or graph")
        collected: list[DependenceEdge] = []
        for series_id in series_ids:
            collected.extend(
                collect_series_edges(catalog.get(series_id), catalog=catalog, graph=graph)
            )
        edges = collected
    return tuple(
        edge
        for edge in edges
        if edge.consumer_id in wanted
        and edge.producer_id in wanted
        and edge.access not in {"whole", "dynamic"}
    )


def _zero_distance_edges(
    scc: tuple[str, ...],
    edges: Sequence[DependenceEdge],
) -> list[DependenceEdge]:
    members = set(scc)
    zero: list[DependenceEdge] = []
    for edge in edges:
        if edge.consumer_id not in members or edge.producer_id not in members:
            continue
        if edge.access == "cross_partition":
            continue
        if edge.distance != 0:
            continue
        zero.append(edge)
    return zero


def _residual_by_index_partition(
    scc: tuple[str, ...],
    edges: Sequence[DependenceEdge],
    catalog: SeriesCatalog,
    *,
    include_guarded: bool,
) -> dict[tuple[int, tuple[Scalar, ...]], dict[str, list[str]]]:
    """Group distance-zero residuals by `(TIME_PERIOD, outer partition)`.

    Opposite orientations in different `ISSUANCE_YEAR` (or `INSTRUMENT` /
    `HOLDER`) blocks must not share one series-id quotient (#762).
    """
    by_key: dict[tuple[int, tuple[Scalar, ...]], dict[str, list[str]]] = {}
    for edge in _zero_distance_edges(scc, edges):
        if edge.guarded and not include_guarded:
            continue
        index = schedule_axis_coord(edge.consumer_cell, catalog)
        part = schedule_partition(edge.consumer_cell, catalog)
        residual = by_key.setdefault((index, part), _empty_residual(scc))
        _add_residual_edge(residual, edge)
    return by_key


def _empty_residual(scc: tuple[str, ...]) -> dict[str, list[str]]:
    return {sid: [] for sid in scc}


def _add_residual_edge(
    residual: dict[str, list[str]],
    edge: DependenceEdge,
) -> None:
    residual[edge.consumer_id].append(edge.producer_id)


def _topo_order(
    scc: tuple[str, ...],
    residual: dict[str, list[str]],
) -> tuple[str, ...] | None:
    remaining = set(scc)
    ordered: list[str] = []
    while remaining:
        ready = [
            sid
            for sid in scc
            if sid in remaining
            and all(pred not in remaining for pred in residual[sid] if pred in remaining)
        ]
        if not ready:
            return None
        for sid in ready:
            remaining.remove(sid)
            ordered.append(sid)
    return tuple(ordered)


def _first_cyclic_pair(residual: dict[str, list[str]]) -> tuple[str, str] | None:
    """Return one residual edge that participates in a cycle."""
    visiting: set[str] = set()
    visited: set[str] = set()

    def dfs(node: str) -> tuple[str, str] | None:
        visiting.add(node)
        for pred in residual.get(node, []):
            if pred in visiting:
                return node, pred
            if pred not in visited:
                found = dfs(pred)
                if found is not None:
                    return found
        visiting.remove(node)
        visited.add(node)
        return None

    for start in residual:
        if start not in visited:
            pair = dfs(start)
            if pair is not None:
                return pair
    return None


def _residual_cycle_message(
    scc: tuple[str, ...],
    index: int,
    residual: dict[str, list[str]],
    edges: Sequence[DependenceEdge],
    catalog: SeriesCatalog,
    partition: tuple[Scalar, ...] = (),
) -> str:
    """Name the two statements and the index point of a residual cycle."""
    pair = _first_cyclic_pair(residual)
    prefix = f"distance-zero residual of zipper series {list(scc)!r} is cyclic at index {index}"
    if pair is None:
        return prefix
    consumer_id, producer_id = pair

    def _match(*, require_unguarded: bool) -> DependenceEdge | None:
        return next(
            (
                edge
                for edge in _zero_distance_edges(scc, edges)
                if schedule_axis_coord(edge.consumer_cell, catalog) == index
                and schedule_partition(edge.consumer_cell, catalog) == partition
                and edge.consumer_id == consumer_id
                and edge.producer_id == producer_id
                and (not require_unguarded or not edge.guarded)
            ),
            None,
        )

    match = _match(require_unguarded=True) or _match(require_unguarded=False)
    if match is not None:
        return (
            f"{prefix} ({consumer_id} {match.consumer_cell} reads "
            f"{producer_id} {match.producer_cell})"
        )
    return f"{prefix} ({consumer_id} reads {producer_id})"


def _scc_partitions(scc: tuple[str, ...], catalog: SeriesCatalog) -> tuple[tuple[Scalar, ...], ...]:
    """Return outer-key blocks in catalog appearance order."""
    seen: list[tuple[Scalar, ...]] = []
    seen_set: set[tuple[Scalar, ...]] = set()
    for series_id in scc:
        for address in catalog.get(series_id).cells:
            part = schedule_partition(address, catalog)
            if part not in seen_set:
                seen_set.add(part)
                seen.append(part)
    return tuple(seen)


def _first_partition_cycle(
    residual: dict[tuple[Scalar, ...], list[tuple[Scalar, ...]]],
) -> tuple[tuple[Scalar, ...], tuple[Scalar, ...]] | None:
    """Return one partition edge that participates in a cycle."""
    visiting: set[tuple[Scalar, ...]] = set()
    visited: set[tuple[Scalar, ...]] = set()

    def dfs(node: tuple[Scalar, ...]) -> tuple[tuple[Scalar, ...], tuple[Scalar, ...]] | None:
        visiting.add(node)
        for pred in residual.get(node, []):
            if pred in visiting:
                return node, pred
            if pred not in visited:
                found = dfs(pred)
                if found is not None:
                    return found
        visiting.remove(node)
        visited.add(node)
        return None

    for start in residual:
        if start not in visited:
            pair = dfs(start)
            if pair is not None:
                return pair
    return None


def _cross_partition_cycle_message(
    scc: tuple[str, ...],
    pair: tuple[tuple[Scalar, ...], tuple[Scalar, ...]],
    edges: Sequence[DependenceEdge],
    catalog: SeriesCatalog,
) -> str:
    """Name both cells of a mutual same-index cross-partition cycle."""
    consumer_part, producer_part = pair
    matches = [
        edge
        for edge in edges
        if edge.access == "cross_partition"
        and schedule_partition(edge.consumer_cell, catalog) == consumer_part
        and schedule_partition(edge.producer_cell, catalog) == producer_part
    ]
    reverse = [
        edge
        for edge in edges
        if edge.access == "cross_partition"
        and schedule_partition(edge.consumer_cell, catalog) == producer_part
        and schedule_partition(edge.producer_cell, catalog) == consumer_part
    ]
    prefix = f"distance-zero residual of zipper series {list(scc)!r} is cyclic"
    if matches and reverse:
        first, second = matches[0], reverse[0]
        return (
            f"{prefix} ({first.consumer_id} {first.consumer_cell} reads "
            f"{first.producer_id} {first.producer_cell}, "
            f"{second.consumer_id} {second.consumer_cell} reads "
            f"{second.producer_id} {second.producer_cell})"
        )
    if matches:
        edge = matches[0]
        return (
            f"{prefix} ({edge.consumer_id} {edge.consumer_cell} reads "
            f"{edge.producer_id} {edge.producer_cell})"
        )
    return prefix


def _unconditional_cell_graph_is_cyclic(
    scc: tuple[str, ...],
    edges: Sequence[DependenceEdge],
) -> bool:
    """Return True when unguarded intra-SCC cell edges contain a cycle."""
    members = set(scc)
    residual: dict[str, list[str]] = {}
    for edge in edges:
        if edge.guarded:
            continue
        if edge.consumer_id not in members or edge.producer_id not in members:
            continue
        residual.setdefault(edge.consumer_cell, []).append(edge.producer_cell)
    visiting: set[str] = set()
    visited: set[str] = set()

    def dfs(node: str) -> bool:
        visiting.add(node)
        for pred in residual.get(node, ()):
            if pred in visiting:
                return True
            if pred not in visited and dfs(pred):
                return True
        visiting.remove(node)
        visited.add(node)
        return False

    return any(node not in visited and dfs(node) for node in residual)


def _assert_cross_partition_legal(
    scc: tuple[str, ...],
    edges: Sequence[DependenceEdge],
    catalog: SeriesCatalog,
) -> None:
    """Fail closed when unguarded cross-partition cell edges form a cycle.

    A cycle in the contracted partition graph is not enough to reject: an
    acyclic cell graph is legal under demand-driven evaluation. A real
    circular reference at cell grain raises.

    Raises:
        InvertedTreeExportError: Unguarded cell edges form a cycle.
    """
    partitions = _scc_partitions(scc, catalog)
    if len(partitions) < 2:
        return
    members = set(scc)
    cross = [
        edge
        for edge in edges
        if edge.access == "cross_partition"
        and edge.consumer_id in members
        and edge.producer_id in members
    ]
    if not cross:
        return
    residual: dict[tuple[Scalar, ...], list[tuple[Scalar, ...]]] = {part: [] for part in partitions}
    for edge in cross:
        consumer_part = schedule_partition(edge.consumer_cell, catalog)
        producer_part = schedule_partition(edge.producer_cell, catalog)
        if consumer_part == producer_part:
            continue
        if consumer_part not in residual or producer_part not in residual:
            return
        residual[consumer_part].append(producer_part)
    pair = _first_partition_cycle(residual)
    if pair is not None and _unconditional_cell_graph_is_cyclic(scc, edges):
        raise InvertedTreeExportError(_cross_partition_cycle_message(scc, pair, cross, catalog))


def assert_distance_zero_legal(
    scc: tuple[str, ...],
    edges: Sequence[DependenceEdge],
    catalog: SeriesCatalog,
) -> None:
    """Fail closed when some partition has an unconditional same-index cycle.

    Residual edges are a DAG per outer-key partition (`ISSUANCE_YEAR`, and
    `INSTRUMENT` / `HOLDER` when present) at one `TIME_PERIOD`. Opposite
    orientations across vintages are not a must-cycle (#762). A cycle with
    no guarded edges raises at plan time. A cycle through a guarded edge
    is a may-cycle and is decided at runtime. Mutual same-index reads
    across outer keys are a cell-grain circular and also fail closed.

    Raises:
        InvertedTreeExportError: Some partition's residual has an
            unconditional cycle at a schedule index, or unguarded
            cross-partition cell edges form a circular reference.
    """
    for (index, part), residual in _residual_by_index_partition(
        scc, edges, catalog, include_guarded=False
    ).items():
        if _topo_order(scc, residual) is None:
            raise InvertedTreeExportError(
                _residual_cycle_message(scc, index, residual, edges, catalog, part)
            )
    _assert_cross_partition_legal(scc, edges, catalog)


def build_scc_map(
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    graph: DependencyGraph | None = None,
    *,
    edges: Sequence[DependenceEdge] | None = None,
) -> dict[str, tuple[str, ...]]:
    """Map each formula series to its SCC (bindings order).

    Multi-series SCCs fail closed only when some (schedule index, outer
    partition) has an unconditional same-index must-cycle. Opposite
    residual orientations in different vintages are legal (#762).
    May-cycles through guarded edges do not raise here; runtime demand
    decides them.

    Pass `edges` when the catalog has already been walked.
    """
    ids = [series.series_id for series in catalog.formula_series()]
    mapping: dict[str, tuple[str, ...]] = {}
    for scc in tarjan_series_sccs(ids, deps):
        if len(scc) > 1:
            scc_edges = collect_dependence_edges(catalog, graph, scc, edges=edges)
            assert_distance_zero_legal(scc, scc_edges, catalog)
        for sid in scc:
            mapping[sid] = scc
    return mapping
