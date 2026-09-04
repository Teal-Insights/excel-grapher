"""Statement-graph scheduling: condensation, distance, residual legality.

A bound series is a statement. Excel's graph is over instances. Contracting
statements invents cycles that do not exist at cell grain (#603). The legality
test is: condense, drop lexicographically positive-distance edges, and require
the distance-zero residual to be a DAG (Allen–Kennedy / Lustre causality).
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING, Literal

from excel_grapher.exporter.inverted_tree.catalog import (
    SeriesCatalog,
    schedule_axis_coord,
    schedule_partition,
)
from excel_grapher.exporter.inverted_tree.deps import (
    DependenceEdge,
    collect_series_edges,
    requires_demand_driven,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.series_bindings.types import Scalar

if TYPE_CHECKING:
    from collections.abc import Mapping

    from excel_grapher.exporter.inverted_tree.deps import SeriesDeps
    from excel_grapher.grapher.graph import DependencyGraph


@dataclass(frozen=True, slots=True)
class IndexSet:
    """Symbolic set of catalog indices.

    Range, strided slice, affine image, and residue predicates normalize to an
    arithmetic progression when they are one. Genuinely irregular gathers keep
    a literal tuple, which is the only form whose emitted size grows with the
    number of members.
    """

    _start: int | None
    _stop: int
    _step: int
    _items: tuple[int, ...]

    def is_empty(self) -> bool:
        """True when the set contains no indices."""
        if self._start is None:
            return not self._items
        return len(range(self._start, self._stop, self._step)) == 0

    def is_progression(self) -> bool:
        """True when this set is empty, a singleton, or an arithmetic progression."""
        return self._start is not None

    def materialize(self) -> tuple[int, ...]:
        """Return the indices in increasing order."""
        if self._start is None:
            return self._items
        return tuple(range(self._start, self._stop, self._step))

    def to_source(self) -> str:
        """Return a Python expression for `take(..., <this>)`."""
        items = self.materialize()
        if len(items) <= 1:
            if not items:
                return "()"
            return f"({items[0]},)"
        if self._start is not None:
            if self._step == 1:
                return f"range({self._start}, {self._stop})"
            return f"range({self._start}, {self._stop}, {self._step})"
        return f"({', '.join(str(i) for i in items)})"

    def union(self, other: IndexSet) -> IndexSet:
        """Return the set-theoretic union, recompressed if it is a progression."""
        if self.is_empty():
            return other
        if other.is_empty():
            return self
        if (
            self._start is not None
            and other._start is not None
            and self._step == other._step
            and (self._start - other._start) % self._step == 0
            and self._stop >= other._start
            and other._stop >= self._start
        ):
            return IndexSet.interval(
                min(self._start, other._start),
                max(self._stop, other._stop),
                self._step,
            )
        return IndexSet.from_indices((*self.materialize(), *other.materialize()))

    def map_affine(self, coeff: int, offset: int) -> IndexSet:
        """Return `{coeff * i + offset | i in self}`."""
        if self.is_empty():
            return self
        if coeff == 0:
            return IndexSet.interval(offset, offset + 1)
        if self._start is not None:
            return IndexSet.interval(
                coeff * self._start + offset,
                coeff * self._stop + offset,
                coeff * self._step,
            )
        return IndexSet.from_indices(coeff * i + offset for i in self._items)

    def filter_residue(self, modulus: int, residue: int) -> IndexSet:
        """Return `{i in self | i % modulus == residue}`."""
        if modulus <= 0:
            raise ValueError("IndexSet residue modulus must be positive")
        residue %= modulus
        if self._start is not None and self._step == 1:
            first = self._start + (residue - self._start) % modulus
            if first >= self._stop:
                return IndexSet.empty()
            return IndexSet.interval(first, self._stop, modulus)
        return IndexSet.from_indices(i for i in self.materialize() if i % modulus == residue)

    def filter(self, pred: Callable[[int], bool]) -> IndexSet:
        """Return `{i in self | pred(i)}`, recompressed when the result is a slice."""
        return IndexSet.from_indices(i for i in self.materialize() if pred(i))

    def closure_under(self, distances: Sequence[int]) -> IndexSet:
        """Close this set under `i -> i - d` for each positive `d` in `distances`.

        Unit lag fills `[0, max]` in one step. A single stride on one residue
        class walks that progression down. Mixed or irregular distances fall
        back to a finite fixed-point on the materialized members.
        """
        lags = tuple(sorted({d for d in distances if d > 0}))
        if not lags or self.is_empty():
            return self
        if len(lags) == 1 and self._start is not None:
            distance = lags[0]
            count = len(range(self._start, self._stop, self._step))
            if count <= 1 or self._step % distance == 0:
                start = self._start
                while start - distance >= 0:
                    start -= distance
                last = self._stop - self._step
                return IndexSet.interval(start, last + distance, distance)
        if 1 in lags:
            return IndexSet.interval(0, self._max() + 1)
        needed = set(self.materialize())
        stack = list(needed)
        while stack:
            index = stack.pop()
            for distance in lags:
                pred = index - distance
                if pred >= 0 and pred not in needed:
                    needed.add(pred)
                    stack.append(pred)
        return IndexSet.from_indices(needed)

    def closure_under_edges(
        self,
        edges: Sequence[DependenceEdge],
        *,
        producer_id: str | None = None,
    ) -> IndexSet:
        """Close under positive distances of `edges` that produce `producer_id`."""
        distances = [
            edge.distance
            for edge in edges
            if edge.distance > 0 and (producer_id is None or edge.producer_id == producer_id)
        ]
        return self.closure_under(distances)

    def positions_in(self, universe: IndexSet) -> IndexSet:
        """Return this set's positions inside `universe`'s materialized order.

        Raises:
            InvertedTreeExportError: An index is missing from `universe`.
        """
        if self.is_empty():
            return self
        if (
            self._start is not None
            and universe._start is not None
            and self._step % universe._step == 0
            and (self._start - universe._start) % universe._step == 0
            and self._start >= universe._start
            and (self._stop - self._step) <= (universe._stop - universe._step)
        ):
            start_pos = (self._start - universe._start) // universe._step
            step_pos = self._step // universe._step
            count = len(range(self._start, self._stop, self._step))
            return IndexSet.interval(start_pos, start_pos + count * step_pos, step_pos)
        pos = {value: index for index, value in enumerate(universe.materialize())}
        try:
            return IndexSet.from_indices(pos[index] for index in self.materialize())
        except KeyError as exc:
            raise InvertedTreeExportError(f"index {exc.args[0]} is not in the universe") from exc

    def _max(self) -> int:
        if self.is_empty():
            raise ValueError("empty IndexSet has no maximum")
        if self._start is None:
            return self._items[-1]
        return self._stop - self._step

    @classmethod
    def empty(cls) -> IndexSet:
        """Return the empty set."""
        return cls(_start=0, _stop=0, _step=1, _items=())

    @classmethod
    def interval(cls, start: int, stop: int, step: int = 1) -> IndexSet:
        """Return the arithmetic progression `range(start, stop, step)`.

        A negative `step` is stored as the equivalent increasing progression.
        """
        if step == 0:
            raise ValueError("IndexSet step must be nonzero")
        if step < 0:
            return cls.from_indices(range(start, stop, step))
        items = range(start, stop, step)
        if not items:
            return cls.empty()
        first = items[0]
        last = items[-1]
        return cls(_start=first, _stop=last + step, _step=step, _items=())

    @classmethod
    def from_indices(cls, indices: Iterable[int]) -> IndexSet:
        """Compress `indices` to a progression, or keep a sorted gather."""
        items = tuple(sorted(set(indices)))
        if not items:
            return cls.empty()
        if len(items) == 1:
            return cls.interval(items[0], items[0] + 1)
        step = items[1] - items[0]
        if step > 0 and items == tuple(range(items[0], items[-1] + step, step)):
            return cls.interval(items[0], items[-1] + step, step)
        return cls(_start=None, _stop=0, _step=1, _items=items)

    @classmethod
    def affine(cls, base: IndexSet, coeff: int, offset: int) -> IndexSet:
        """Return the affine image of `base` under `i -> coeff * i + offset`."""
        return base.map_affine(coeff, offset)


def indices_to_source(indices: Sequence[int]) -> str:
    """Return a Python expression for an ordered index sequence.

    Unlike `IndexSet.to_source`, this preserves decreasing progressions so
    `take` can realign an anti-monotone affine map.
    """
    items = tuple(indices)
    if len(items) <= 1:
        if not items:
            return "()"
        return f"({items[0]},)"
    step = items[1] - items[0]
    if step != 0:
        stop = items[-1] + step
        if items == tuple(range(items[0], stop, step)):
            if step == 1:
                return f"range({items[0]}, {stop})"
            return f"range({items[0]}, {stop}, {step})"
    return f"({', '.join(str(i) for i in items)})"


def scan_function_name(scc: tuple[str, ...]) -> str:
    """Return the internals helper name for a fused or demand-driven SCC."""
    return "scan_" + "_".join(scc)


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


@dataclass(frozen=True, slots=True)
class FusedRegion:
    """One residual-order / access-class span on the union schedule."""

    start: int
    stop: int
    body_order: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class FusedPlan:
    """Union-domain loop for a fusible SCC (rung 2, and rung-1 as a singleton).

    `regions` is the source of truth for residual order along the schedule.
    `partitions` is the ordered outer-key nest (`()` when the SCC is 1-D).
    When partitions are not isomorphic, `unroll` is True and
    `partition_regions` holds one region tuple per outer key.
    """

    scc: tuple[str, ...]
    schedule: tuple[int, ...]
    domain: dict[str, tuple[int, int]]
    regions: tuple[FusedRegion, ...]
    direction: Literal["forward", "reversed"] = "forward"
    partitions: tuple[tuple[Scalar, ...], ...] = ()
    unroll: bool = False
    partition_regions: tuple[tuple[FusedRegion, ...], ...] = ()

    @property
    def coord_to_t(self) -> dict[int, int]:
        """Map each schedule-axis coordinate to its union index `t`."""
        return {coord: index for index, coord in enumerate(self.schedule)}

    @property
    def is_nested(self) -> bool:
        """True when emission wraps the fused body in an outer partition loop."""
        return len(self.partitions) > 1


def _contiguous_span(locals_: Sequence[int]) -> tuple[int, int] | None:
    """Return `[start, stop)` when `locals_` is a hole-free interval."""
    if not locals_:
        return None
    start, stop = min(locals_), max(locals_) + 1
    if len(locals_) != stop - start or set(locals_) != set(range(start, stop)):
        return None
    return start, stop


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


def _contiguous_domain(
    series_id: str,
    catalog: SeriesCatalog,
    coord_to_t: dict[int, int],
    partitions: Sequence[tuple[Scalar, ...]],
) -> tuple[int, int] | None:
    """Return `[start, stop)` in union-index space, or None if a domain has holes.

    Contiguity is checked per outer-key block (#638). A late-start `adj`
    that is `[1, 2]` in France and `[1, 2]` in Kenya fuses; the flattened
    `[1, 2, 4, 5]` does not have to be a single interval. Every partition
    must carry the same span so one inner loop can serve all blocks.
    """
    series = catalog.get(series_id)
    nested = len(partitions) > 1
    if not nested:
        locals_: list[int] = []
        for address in series.cells:
            coord = schedule_axis_coord(address, catalog)
            if coord not in coord_to_t:
                return None
            locals_.append(coord_to_t[coord])
        return _contiguous_span(locals_)
    by_part: dict[tuple[Scalar, ...], list[int]] = {part: [] for part in partitions}
    for address in series.cells:
        part = schedule_partition(address, catalog)
        coord = schedule_axis_coord(address, catalog)
        if coord not in coord_to_t or part not in by_part:
            return None
        by_part[part].append(coord_to_t[coord])
    spans: list[tuple[int, int]] = []
    for part in partitions:
        span = _contiguous_span(by_part[part])
        if span is None:
            return None
        spans.append(span)
    if len(set(spans)) != 1:
        return None
    return spans[0]


def _statement_at_union(
    catalog: SeriesCatalog,
    series_id: str,
    union_t: int,
    index: int,
) -> str:
    """Return the statement covering union index `union_t`."""
    series = catalog.get(series_id)
    if len(series.statements) <= 1:
        return series_id
    for stmt in series.statements:
        for cell in stmt.cells:
            if schedule_axis_coord(cell, catalog) == index:
                return stmt.statement_id
    return series_id


def _residual_order_at_index(
    members: tuple[str, ...],
    index_edges: Sequence[DependenceEdge],
) -> tuple[str, ...] | None:
    """Return the distance-zero topo among `members` at one schedule index."""
    if not members:
        return ()
    residual = _empty_residual(members)
    active = set(members)
    for edge in index_edges:
        if edge.distance != 0:
            continue
        if edge.consumer_id not in active or edge.producer_id not in active:
            continue
        _add_residual_edge(residual, edge)
    return _topo_order(members, residual)


def _access_signature(
    members: tuple[str, ...],
    index_edges: Sequence[DependenceEdge],
) -> tuple[tuple[str, str, str], ...]:
    """Return intra-member access classes at `index`, sorted for grouping."""
    active = set(members)
    sig = [
        (edge.consumer_id, edge.producer_id, edge.access)
        for edge in index_edges
        if edge.consumer_id in active and edge.producer_id in active
    ]
    sig.sort()
    return tuple(sig)


def _index_region_key(
    scc: tuple[str, ...],
    *,
    catalog: SeriesCatalog,
    domain: Mapping[str, tuple[int, int]],
    index_edges: Sequence[DependenceEdge],
    union_t: int,
    index: int,
) -> tuple[tuple[str, ...], tuple[tuple[str, str], ...], tuple[tuple[str, str, str], ...]] | None:
    """Return `(body_order, shape_sig, access_sig)` for one union index."""
    active = tuple(sid for sid in scc if domain[sid][0] <= union_t < domain[sid][1])
    if not active:
        return None
    order = _residual_order_at_index(active, index_edges)
    if order is None:
        return None
    shape_sig = tuple((sid, _statement_at_union(catalog, sid, union_t, index)) for sid in active)
    return order, shape_sig, _access_signature(active, index_edges)


def _bucket_edges_by_consumer_coord(
    edges: Sequence[DependenceEdge],
    catalog: SeriesCatalog,
    *,
    partition: tuple[Scalar, ...] | None = None,
) -> dict[int, list[DependenceEdge]]:
    """Group same-partition intra-SCC edges by the consumer's axis coordinate."""
    buckets: dict[int, list[DependenceEdge]] = {}
    for edge in edges:
        if edge.access == "cross_partition":
            continue
        if partition is not None and schedule_partition(edge.consumer_cell, catalog) != partition:
            continue
        coord = schedule_axis_coord(edge.consumer_cell, catalog)
        buckets.setdefault(coord, []).append(edge)
    return buckets


def _fuse_regions(
    scc: tuple[str, ...],
    *,
    catalog: SeriesCatalog,
    domain: Mapping[str, tuple[int, int]],
    edges: Sequence[DependenceEdge],
    coords: Sequence[int],
    partition: tuple[Scalar, ...] | None = None,
) -> tuple[FusedRegion, ...] | None:
    """Group contiguous union indices that share residual order and access."""
    by_coord = _bucket_edges_by_consumer_coord(edges, catalog, partition=partition)
    regions: list[FusedRegion] = []
    run_start = 0
    run_key: (
        tuple[tuple[str, ...], tuple[tuple[str, str], ...], tuple[tuple[str, str, str], ...]] | None
    ) = None
    for union_t, index in enumerate(coords):
        key = _index_region_key(
            scc,
            catalog=catalog,
            domain=domain,
            index_edges=by_coord.get(index, ()),
            union_t=union_t,
            index=index,
        )
        if key is None:
            return None
        if run_key is None:
            run_start = union_t
            run_key = key
            continue
        if key != run_key:
            regions.append(FusedRegion(start=run_start, stop=union_t, body_order=run_key[0]))
            run_start = union_t
            run_key = key
    if run_key is None:
        return None
    regions.append(FusedRegion(start=run_start, stop=len(coords), body_order=run_key[0]))
    return tuple(regions)


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


def _assert_cross_partition_legal(
    scc: tuple[str, ...],
    edges: Sequence[DependenceEdge],
    catalog: SeriesCatalog,
    partitions: Sequence[tuple[Scalar, ...]],
) -> bool:
    """Fail closed on a partition cycle; return False for a forward outer read.

    A `cross_partition` edge is legal only when it points at an already
    completed outer iteration under `partitions` order. Mutual same-index
    reads are a real circular reference and raise.

    Raises:
        InvertedTreeExportError: Two partitions read each other at the same
            index.
    """
    members = set(scc)
    cross = [
        edge
        for edge in edges
        if edge.access == "cross_partition"
        and edge.consumer_id in members
        and edge.producer_id in members
    ]
    if not cross:
        return True
    residual: dict[tuple[Scalar, ...], list[tuple[Scalar, ...]]] = {part: [] for part in partitions}
    for edge in cross:
        consumer_part = schedule_partition(edge.consumer_cell, catalog)
        producer_part = schedule_partition(edge.producer_cell, catalog)
        if consumer_part == producer_part:
            continue
        if consumer_part not in residual or producer_part not in residual:
            return False
        residual[consumer_part].append(producer_part)
    pair = _first_partition_cycle(residual)
    if pair is not None:
        raise InvertedTreeExportError(_cross_partition_cycle_message(scc, pair, cross, catalog))
    rank = {part: index for index, part in enumerate(partitions)}
    for edge in cross:
        consumer_part = schedule_partition(edge.consumer_cell, catalog)
        producer_part = schedule_partition(edge.producer_cell, catalog)
        if rank[producer_part] > rank[consumer_part]:
            return False
    return True


def plan_fused_scc(
    scc: tuple[str, ...],
    *,
    catalog: SeriesCatalog,
    graph: DependencyGraph | None = None,
    edges: Sequence[DependenceEdge] | None = None,
) -> FusedPlan | None:
    """Return a fused loop plan, or None when the SCC must stay on rung 3.

    Requires uniform loop direction (all intra-SCC nonzero same-partition
    distances positive for a forward loop, or all negative for a reversed
    loop) and a contiguous domain per statement on the union schedule.
    Contiguity is per outer-key block when the key is a nest (#638).
    Residual order, formula shape, and access class may change along the
    schedule; each distinct span becomes a `FusedRegion`. A singleton SCC
    with positive-distance self-lags is the rung-1 scan: peel the first
    `max(D)` members and index the growing buffer.

    Pass `edges` when the catalog has already been walked.

    Raises:
        InvertedTreeExportError: Some index's residual is a real same-index
            cycle, including a mutual cross-partition read.
    """
    if not scc:
        return None
    members = set(scc)
    edges = collect_dependence_edges(catalog, graph, scc, edges=edges)
    intra = [edge for edge in edges if edge.consumer_id in members and edge.producer_id in members]
    same_part = [edge for edge in intra if edge.access != "cross_partition"]
    if (
        len(scc) < 2
        and not any(edge.distance != 0 for edge in same_part)
        and not any(edge.access == "cross_partition" for edge in intra)
    ):
        return None
    nonzero = [edge.distance for edge in same_part if edge.distance != 0]
    has_pos = any(d > 0 for d in nonzero)
    has_neg = any(d < 0 for d in nonzero)
    if has_pos and has_neg:
        return None
    direction: Literal["forward", "reversed"] = "reversed" if has_neg else "forward"
    assert_distance_zero_legal(scc, edges, catalog)
    if has_residual_may_cycle(scc, edges, catalog):
        return None
    partitions = _scc_partitions(scc, catalog)
    nested = len(partitions) > 1
    if nested and not _assert_cross_partition_legal(scc, intra, catalog, partitions):
        return None
    coords = sorted(
        {schedule_axis_coord(addr, catalog) for sid in scc for addr in catalog.get(sid).cells},
        reverse=(direction == "reversed"),
    )
    if not coords:
        return None
    coord_to_t = {coord: index for index, coord in enumerate(coords)}
    domain: dict[str, tuple[int, int]] = {}
    for sid in scc:
        span = _contiguous_domain(sid, catalog, coord_to_t, partitions if nested else ())
        if span is None:
            return None
        domain[sid] = span
    representative = partitions[0] if nested else None
    regions = _fuse_regions(
        scc,
        catalog=catalog,
        domain=domain,
        edges=edges,
        coords=coords,
        partition=representative,
    )
    if regions is None:
        return None
    unroll = False
    partition_regions: tuple[tuple[FusedRegion, ...], ...] = ()
    if nested:
        per_part: list[tuple[FusedRegion, ...]] = []
        for part in partitions:
            part_regions = _fuse_regions(
                scc,
                catalog=catalog,
                domain=domain,
                edges=edges,
                coords=coords,
                partition=part,
            )
            if part_regions is None:
                return None
            per_part.append(part_regions)
        has_cross = any(edge.access == "cross_partition" for edge in intra)
        if has_cross or any(part_regions != regions for part_regions in per_part):
            unroll = True
            partition_regions = tuple(per_part)
    return FusedPlan(
        scc=scc,
        schedule=tuple(coords),
        domain=domain,
        regions=regions,
        direction=direction,
        partitions=partitions if nested else (),
        unroll=unroll,
        partition_regions=partition_regions,
    )


Rung = Literal[0, 1, 2, 3]


@dataclass(frozen=True, slots=True)
class SccPlan:
    """Evaluation rung for one SCC, plus the fused plan when the SCC is fusible.

    Rungs:
        0: first-level helper (`emit_helper_body`), including reversed unit scans
        1: singleton fused forward scan (`emit_rung2_scc`)
        2: multi-series fused loop (`emit_rung2_scc`)
        3: demand-driven instance evaluation (`emit_rung3_scc`)
    """

    rung: Rung
    plan: FusedPlan | None = None


def plan_scc(
    scc: tuple[str, ...],
    *,
    catalog: SeriesCatalog,
    graph: DependencyGraph | None = None,
    edges: Sequence[DependenceEdge] | None = None,
) -> SccPlan:
    """Select the emit rung for `scc`.

    Fused classification is `plan_fused_scc`. A singleton that is not a
    forward fused scan uses `requires_demand_driven` to choose rung 3 vs 0.
    Tests should assert `plan_scc(...).rung` rather than grepping emitted
    source for `eval_instance`. Pass `edges` when the catalog has already
    been walked.
    """
    fused = plan_fused_scc(scc, catalog=catalog, graph=graph, edges=edges)
    if len(scc) > 1:
        if fused is not None:
            return SccPlan(rung=2, plan=fused)
        return SccPlan(rung=3)
    if fused is not None and fused.direction == "forward":
        return SccPlan(rung=1, plan=fused)
    if requires_demand_driven(catalog.get(scc[0]), catalog=catalog, graph=graph, edges=edges):
        return SccPlan(rung=3)
    return SccPlan(rung=0)


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
) -> str:
    """Name the two statements and the index point of a residual cycle."""
    pair = _first_cyclic_pair(residual)
    prefix = f"distance-zero residual of zipper series {list(scc)!r} is cyclic at index {index}"
    if pair is None:
        return prefix
    consumer_id, producer_id = pair
    match = next(
        (
            edge
            for edge in _zero_distance_edges(scc, edges)
            if schedule_axis_coord(edge.consumer_cell, catalog) == index
            and edge.consumer_id == consumer_id
            and edge.producer_id == producer_id
            and not edge.guarded
        ),
        None,
    ) or next(
        (
            edge
            for edge in _zero_distance_edges(scc, edges)
            if schedule_axis_coord(edge.consumer_cell, catalog) == index
            and edge.consumer_id == consumer_id
            and edge.producer_id == producer_id
        ),
        None,
    )
    if match is not None:
        return (
            f"{prefix} ({consumer_id} {match.consumer_cell} reads "
            f"{producer_id} {match.producer_cell})"
        )
    return f"{prefix} ({consumer_id} reads {producer_id})"


def assert_distance_zero_legal(
    scc: tuple[str, ...],
    edges: Sequence[DependenceEdge],
    catalog: SeriesCatalog,
) -> None:
    """Fail closed when some schedule index has an unconditional same-index cycle.

    A cycle with no guarded edges is a must-cycle and raises at plan time.
    A cycle passing through at least one guarded edge is a may-cycle and
    is decided at runtime (demoted to rung 3).

    Raises:
        InvertedTreeExportError: Some index's residual has an unconditional cycle.
    """
    by_index: dict[int, dict[str, list[str]]] = {}
    for edge in _zero_distance_edges(scc, edges):
        if edge.guarded:
            continue
        index = schedule_axis_coord(edge.consumer_cell, catalog)
        residual = by_index.setdefault(index, _empty_residual(scc))
        _add_residual_edge(residual, edge)
    for index, residual in by_index.items():
        if _topo_order(scc, residual) is None:
            raise InvertedTreeExportError(
                _residual_cycle_message(scc, index, residual, edges, catalog)
            )


def has_residual_may_cycle(
    scc: tuple[str, ...],
    edges: Sequence[DependenceEdge],
    catalog: SeriesCatalog,
) -> bool:
    """Return True if any schedule index has a residual cycle using guarded edges."""
    by_index: dict[int, dict[str, list[str]]] = {}
    for edge in _zero_distance_edges(scc, edges):
        index = schedule_axis_coord(edge.consumer_cell, catalog)
        residual = by_index.setdefault(index, _empty_residual(scc))
        _add_residual_edge(residual, edge)
    return any(_topo_order(scc, residual) is None for residual in by_index.values())


def residual_body_order(
    scc: tuple[str, ...],
    edges: Sequence[DependenceEdge],
    catalog: SeriesCatalog,
) -> tuple[str, ...] | None:
    """Return one in-loop statement order, or None if index points disagree.

    Raises:
        InvertedTreeExportError: A real same-index circular reference exists
            at some schedule index.
    """
    assert_distance_zero_legal(scc, edges, catalog)
    union = _empty_residual(scc)
    for edge in _zero_distance_edges(scc, edges):
        _add_residual_edge(union, edge)
    return _topo_order(scc, union)


def build_scc_map(
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    graph: DependencyGraph | None = None,
    *,
    edges: Sequence[DependenceEdge] | None = None,
) -> dict[str, tuple[str, ...]]:
    """Map each formula series to its SCC (bindings order).

    Multi-series SCCs fail closed only when some schedule index has an
    unconditional same-index must-cycle. May-cycles through guarded edges
    do not raise here; they demote to rung 3 in plan_fused_scc.

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
