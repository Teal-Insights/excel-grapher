"""Statement-graph scheduling: condensation, distance, residual legality.

A bound series is a statement. Excel's graph is over instances. Contracting
statements invents cycles that do not exist at cell grain (#603). The legality
test is: condense, drop lexicographically positive-distance edges, and require
the distance-zero residual to be a DAG (Allen–Kennedy / Lustre causality).
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING

from excel_grapher.core.address_keys import parse_cell_coords
from excel_grapher.exporter.inverted_tree.deps import (
    DependenceEdge,
    collect_series_edges,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError

if TYPE_CHECKING:
    from collections.abc import Mapping

    from excel_grapher.exporter.inverted_tree.catalog import SeriesCatalog
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
            ValueError: An index is missing from `universe`.
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
            raise ValueError(f"index {exc.args[0]} is not in the universe") from exc

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


def schedule_coord(address: str) -> int:
    """Return the layout schedule coordinate of `address` (1-based column)."""
    _sheet, _row, col = parse_cell_coords(address)
    return col


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
    graph: DependencyGraph,
    series_ids: Sequence[str],
) -> tuple[DependenceEdge, ...]:
    """Collect instance-level cell-ref edges among `series_ids`.

    `whole` and `dynamic` accesses have no fixed distance and are omitted so
    the residual test stays over concrete instance reads.
    """
    wanted = set(series_ids)
    edges: list[DependenceEdge] = []
    for series_id in series_ids:
        for edge in collect_series_edges(catalog.get(series_id), catalog=catalog, graph=graph):
            if edge.producer_id not in wanted:
                continue
            if edge.access in {"whole", "dynamic"}:
                continue
            edges.append(edge)
    return tuple(edges)


@dataclass(frozen=True, slots=True)
class FusedRegion:
    """One residual-order / access-class span on the union schedule."""

    start: int
    stop: int
    body_order: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class FusedPlan:
    """Union-domain loop for a fusible multi-series SCC (rung 2).

    `regions` is the source of truth. `body_order` and `peel_stop` describe
    the last region: its residual order and the union index where it starts.
    """

    scc: tuple[str, ...]
    schedule: tuple[int, ...]
    domain: dict[str, tuple[int, int]]
    regions: tuple[FusedRegion, ...]

    @property
    def body_order(self) -> tuple[str, ...]:
        """In-loop statement order of the last region."""
        return self.regions[-1].body_order

    @property
    def peel_stop(self) -> int:
        """Union index where the last region begins."""
        return self.regions[-1].start


def _contiguous_domain(
    series_id: str,
    catalog: SeriesCatalog,
    coord_to_t: dict[int, int],
) -> tuple[int, int] | None:
    """Return `[start, stop)` in union-index space, or None if the domain has holes."""
    locals_: list[int] = []
    for address in catalog.get(series_id).cells:
        coord = schedule_coord(address)
        if coord not in coord_to_t:
            return None
        locals_.append(coord_to_t[coord])
    if not locals_:
        return None
    start, stop = locals_[0], locals_[-1] + 1
    if locals_ != list(range(start, stop)):
        return None
    return start, stop


def _statement_at_union(
    catalog: SeriesCatalog,
    series_id: str,
    union_t: int,
    domain_start: int,
) -> str:
    """Return the statement covering union index `union_t`."""
    host_index = union_t - domain_start
    series = catalog.get(series_id)
    for stmt in series.statements:
        if stmt.start <= host_index < stmt.stop:
            return stmt.statement_id
    return series_id


def _residual_order_at_column(
    members: tuple[str, ...],
    edges: Sequence[DependenceEdge],
    column: int,
) -> tuple[str, ...] | None:
    """Return the distance-zero topo among `members` in one schedule column."""
    if not members:
        return ()
    residual = _empty_residual(members)
    for edge in _zero_distance_edges(members, edges):
        if schedule_coord(edge.consumer_cell) != column:
            continue
        _add_residual_edge(residual, edge, members, column=column)
    return _topo_order(members, residual)


def _access_signature(
    members: tuple[str, ...],
    edges: Sequence[DependenceEdge],
    column: int,
) -> tuple[tuple[str, str, str], ...]:
    """Return intra-member access classes at `column`, sorted for grouping."""
    active = set(members)
    sig = [
        (edge.consumer_id, edge.producer_id, edge.access)
        for edge in edges
        if edge.consumer_id in active
        and edge.producer_id in active
        and schedule_coord(edge.consumer_cell) == column
    ]
    sig.sort()
    return tuple(sig)


def _column_region_key(
    scc: tuple[str, ...],
    *,
    catalog: SeriesCatalog,
    domain: Mapping[str, tuple[int, int]],
    edges: Sequence[DependenceEdge],
    union_t: int,
    column: int,
) -> tuple[tuple[str, ...], tuple[tuple[str, str], ...], tuple[tuple[str, str, str], ...]] | None:
    """Return `(body_order, shape_sig, access_sig)` for one union index."""
    active = tuple(sid for sid in scc if domain[sid][0] <= union_t < domain[sid][1])
    if not active:
        return None
    order = _residual_order_at_column(active, edges, column)
    if order is None:
        return None
    shape_sig = tuple(
        (sid, _statement_at_union(catalog, sid, union_t, domain[sid][0])) for sid in active
    )
    return order, shape_sig, _access_signature(active, edges, column)


def _fuse_regions(
    scc: tuple[str, ...],
    *,
    catalog: SeriesCatalog,
    domain: Mapping[str, tuple[int, int]],
    edges: Sequence[DependenceEdge],
    coords: Sequence[int],
) -> tuple[FusedRegion, ...] | None:
    """Group contiguous union indices that share residual order and access."""
    regions: list[FusedRegion] = []
    run_start = 0
    run_key: (
        tuple[tuple[str, ...], tuple[tuple[str, str], ...], tuple[tuple[str, str, str], ...]] | None
    ) = None
    for union_t, column in enumerate(coords):
        key = _column_region_key(
            scc,
            catalog=catalog,
            domain=domain,
            edges=edges,
            union_t=union_t,
            column=column,
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


def plan_fused_scc(
    scc: tuple[str, ...],
    *,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
) -> FusedPlan | None:
    """Return a fused loop plan, or None when the SCC must stay on rung 3.

    Requires no look-ahead edges and a contiguous domain per statement on the
    union schedule. Residual order, formula shape, and access class may change
    along the schedule; each distinct span becomes a `FusedRegion`.

    Raises:
        InvertedTreeExportError: Some column's residual is a real same-index
            cycle.
    """
    if len(scc) < 2:
        return None
    members = set(scc)
    edges = collect_dependence_edges(catalog, graph, scc)
    intra = [edge for edge in edges if edge.consumer_id in members and edge.producer_id in members]
    if any(edge.distance < 0 for edge in intra):
        return None
    assert_distance_zero_legal(scc, edges)
    coords = sorted({schedule_coord(addr) for sid in scc for addr in catalog.get(sid).cells})
    if not coords:
        return None
    coord_to_t = {coord: index for index, coord in enumerate(coords)}
    domain: dict[str, tuple[int, int]] = {}
    for sid in scc:
        span = _contiguous_domain(sid, catalog, coord_to_t)
        if span is None:
            return None
        domain[sid] = span
    regions = _fuse_regions(scc, catalog=catalog, domain=domain, edges=edges, coords=coords)
    if regions is None:
        return None
    return FusedPlan(
        scc=scc,
        schedule=tuple(coords),
        domain=domain,
        regions=regions,
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
        if edge.distance != 0:
            continue
        zero.append(edge)
    return zero


def _empty_residual(scc: tuple[str, ...]) -> dict[str, list[str]]:
    return {sid: [] for sid in scc}


def _add_residual_edge(
    residual: dict[str, list[str]],
    edge: DependenceEdge,
    scc: tuple[str, ...],
    *,
    column: int,
) -> None:
    if edge.consumer_id == edge.producer_id:
        raise InvertedTreeExportError(
            f"distance-zero residual of zipper series {list(scc)!r} is cyclic "
            f"at column {column} ({edge.consumer_id} {edge.consumer_cell} reads "
            f"{edge.producer_id} {edge.producer_cell})"
        )
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
    column: int,
    residual: dict[str, list[str]],
    edges: Sequence[DependenceEdge],
) -> str:
    """Name the two statements and the index point of a residual cycle."""
    pair = _first_cyclic_pair(residual)
    prefix = f"distance-zero residual of zipper series {list(scc)!r} is cyclic at column {column}"
    if pair is None:
        return prefix
    consumer_id, producer_id = pair
    match = next(
        (
            edge
            for edge in _zero_distance_edges(scc, edges)
            if schedule_coord(edge.consumer_cell) == column
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
) -> None:
    """Fail closed when some schedule column has a same-index cycle.

    Distance-zero edges from different columns are not contracted. A cycle in
    the union of those edges is a regime flip, not a circular reference.

    Raises:
        InvertedTreeExportError: Some column's residual still has a cycle.
    """
    by_column: dict[int, dict[str, list[str]]] = {}
    for edge in _zero_distance_edges(scc, edges):
        column = schedule_coord(edge.consumer_cell)
        residual = by_column.setdefault(column, _empty_residual(scc))
        _add_residual_edge(residual, edge, scc, column=column)
    for column, residual in by_column.items():
        if _topo_order(scc, residual) is None:
            raise InvertedTreeExportError(_residual_cycle_message(scc, column, residual, edges))


def residual_body_order(
    scc: tuple[str, ...],
    edges: Sequence[DependenceEdge],
) -> tuple[str, ...] | None:
    """Return one in-loop statement order, or None if columns disagree.

    Raises:
        InvertedTreeExportError: A real same-index circular reference exists
            in some schedule column.
    """
    assert_distance_zero_legal(scc, edges)
    union = _empty_residual(scc)
    for edge in _zero_distance_edges(scc, edges):
        _add_residual_edge(union, edge, scc, column=schedule_coord(edge.consumer_cell))
    return _topo_order(scc, union)


def build_scc_map(
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    graph: DependencyGraph,
) -> dict[str, tuple[str, ...]]:
    """Map each formula series to its SCC (bindings order).

    Multi-series SCCs fail closed only when some schedule column has a
    same-index cycle. A residual order that changes across columns is legal
    and is fused region-locally when each span has a residual DAG.
    """
    ids = [series.series_id for series in catalog.formula_series()]
    mapping: dict[str, tuple[str, ...]] = {}
    for scc in tarjan_series_sccs(ids, deps):
        if len(scc) > 1:
            edges = collect_dependence_edges(catalog, graph, scc)
            assert_distance_zero_legal(scc, edges)
        for sid in scc:
            mapping[sid] = scc
    return mapping
