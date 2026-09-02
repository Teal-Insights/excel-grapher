"""Statement-graph scheduling: condensation, distance, residual legality.

A bound series is a statement. Excel's graph is over instances. Contracting
statements invents cycles that do not exist at cell grain (#603). The legality
test is: condense, drop lexicographically positive-distance edges, and require
the distance-zero residual to be a DAG (Allen–Kennedy / Lustre causality).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from excel_grapher.core.address_keys import normalize_key as normalize_address
from excel_grapher.core.address_keys import parse_cell_coords
from excel_grapher.core.excel_function_names import normalize_excel_function_name
from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    CellRefNode,
    FunctionCallNode,
    RangeNode,
    UnaryOpNode,
    resolve_cell_ref,
)
from excel_grapher.exporter.inverted_tree.deps import node_formula_ast
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError

if TYPE_CHECKING:
    from collections.abc import Iterable, Mapping, Sequence

    from excel_grapher.exporter.inverted_tree.catalog import SeriesCatalog
    from excel_grapher.exporter.inverted_tree.deps import SeriesDeps
    from excel_grapher.grapher.graph import DependencyGraph


@dataclass(frozen=True, slots=True)
class DependenceEdge:
    """One instance-level read, annotated with schedule distance.

    `distance` is `coord(consumer) - coord(producer)` in the layout schedule
    (column index for `layout: series` TIME_PERIOD rows). Positive means the
    producer is an earlier period (a `pre` / lag).
    """

    consumer_id: str
    producer_id: str
    consumer_cell: str
    producer_cell: str
    distance: int


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


def _iter_cell_refs(node: AstNode, host_cell: str) -> Iterable[str]:
    match node:
        case CellRefNode():
            yield resolve_cell_ref(node, host_cell)
        case RangeNode():
            return
        case BinaryOpNode():
            yield from _iter_cell_refs(node.left, host_cell)
            yield from _iter_cell_refs(node.right, host_cell)
        case UnaryOpNode():
            yield from _iter_cell_refs(node.operand, host_cell)
        case FunctionCallNode():
            name = normalize_excel_function_name(node.name)
            if name in {"OFFSET", "INDEX", "MATCH"}:
                for arg in node.args[1:]:
                    yield from _iter_cell_refs(arg, host_cell)
                return
            for arg in node.args:
                yield from _iter_cell_refs(arg, host_cell)
        case _:
            return


def collect_dependence_edges(
    catalog: SeriesCatalog,
    graph: DependencyGraph,
    series_ids: Sequence[str],
) -> tuple[DependenceEdge, ...]:
    """Collect instance-level cell-ref edges among `series_ids`."""
    wanted = set(series_ids)
    edges: list[DependenceEdge] = []
    for series_id in series_ids:
        series = catalog.get(series_id)
        for address in series.cells:
            ast = node_formula_ast(graph, address)
            for dep in _iter_cell_refs(ast, address):
                producer = catalog.series_id_for(dep)
                if producer is None or producer not in wanted:
                    continue
                consumer_cell = normalize_address(address)
                producer_cell = normalize_address(dep)
                edges.append(
                    DependenceEdge(
                        consumer_id=series_id,
                        producer_id=producer,
                        consumer_cell=consumer_cell,
                        producer_cell=producer_cell,
                        distance=schedule_coord(consumer_cell) - schedule_coord(producer_cell),
                    )
                )
    return tuple(edges)


@dataclass(frozen=True, slots=True)
class FusedPlan:
    """Union-domain loop for a fusible multi-series SCC (rung 2)."""

    scc: tuple[str, ...]
    schedule: tuple[int, ...]
    domain: dict[str, tuple[int, int]]
    body_order: tuple[str, ...]
    peel_stop: int


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


def plan_fused_scc(
    scc: tuple[str, ...],
    *,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
) -> FusedPlan | None:
    """Return a fused loop plan, or None when the SCC must stay on rung 3.

    Requires an acyclic distance-zero residual, no look-ahead edges, and a
    contiguous domain per statement on the union schedule.

    Raises:
        InvertedTreeExportError: The residual is a real same-index cycle.
    """
    if len(scc) < 2:
        return None
    members = set(scc)
    edges = collect_dependence_edges(catalog, graph, scc)
    intra = [edge for edge in edges if edge.consumer_id in members and edge.producer_id in members]
    if any(edge.distance < 0 for edge in intra):
        return None
    body_order = residual_body_order(scc, edges)
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
    main_active = {
        sid for sid, (start, stop) in domain.items() if start <= (len(coords) - 1) < stop
    }
    peel_stop = len(coords) - 1
    while peel_stop > 0:
        prev = peel_stop - 1
        prev_active = {sid for sid, (start, stop) in domain.items() if start <= prev < stop}
        if prev_active != main_active:
            break
        peel_stop = prev
    return FusedPlan(
        scc=scc,
        schedule=tuple(coords),
        domain=domain,
        body_order=body_order,
        peel_stop=peel_stop,
    )


def residual_body_order(
    scc: tuple[str, ...],
    edges: Sequence[DependenceEdge],
) -> tuple[str, ...]:
    """Return in-loop statement order after dropping positive-distance edges.

    Raises:
        InvertedTreeExportError: The distance-zero residual still has a cycle
            (a real same-index circular reference).
    """
    members = set(scc)
    residual: dict[str, list[str]] = {sid: [] for sid in scc}
    for edge in edges:
        if edge.consumer_id not in members or edge.producer_id not in members:
            continue
        if edge.distance > 0:
            continue
        if edge.consumer_id == edge.producer_id:
            raise InvertedTreeExportError(
                f"distance-zero residual of zipper series {list(scc)!r} is cyclic "
                f"({edge.consumer_cell} reads {edge.producer_cell})"
            )
        residual[edge.consumer_id].append(edge.producer_id)

    remaining = set(scc)
    ordered: list[str] = []
    while remaining:
        ready = [
            sid
            for sid in scc
            if sid in remaining
            and all(pred not in remaining for pred in residual[sid] if pred in members)
        ]
        if not ready:
            raise InvertedTreeExportError(
                f"distance-zero residual of zipper series {list(scc)!r} is cyclic"
            )
        for sid in ready:
            remaining.remove(sid)
            ordered.append(sid)
    return tuple(ordered)


def build_scc_map(
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    graph: DependencyGraph,
) -> dict[str, tuple[str, ...]]:
    """Map each formula series to its SCC (bindings order).

    Multi-series SCCs must have an acyclic distance-zero residual. Same-index
    circular refs fail closed.
    """
    ids = [series.series_id for series in catalog.formula_series()]
    mapping: dict[str, tuple[str, ...]] = {}
    for scc in tarjan_series_sccs(ids, deps):
        if len(scc) > 1:
            edges = collect_dependence_edges(catalog, graph, scc)
            residual_body_order(scc, edges)
        for sid in scc:
            mapping[sid] = scc
    return mapping
