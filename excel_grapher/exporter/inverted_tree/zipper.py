"""Lag-only series SCCs (period-lag zippers) for inverted-tree emit."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import TYPE_CHECKING

from excel_grapher.core.address_keys import normalize_key as normalize_address
from excel_grapher.core.address_keys import parse_cell_coords
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError

if TYPE_CHECKING:
    from collections.abc import Iterable, Mapping, Sequence

    from excel_grapher.exporter.inverted_tree.catalog import SeriesCatalog
    from excel_grapher.exporter.inverted_tree.deps import SeriesDeps
    from excel_grapher.grapher.graph import DependencyGraph


def scan_function_name(scc: tuple[str, ...]) -> str:
    """Return the internals helper name for a zipper SCC."""
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

    Members of each SCC follow `series_ids` order (bindings order).
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


def cells_have_cycle(
    addresses: Iterable[str],
    graph: DependencyGraph,
) -> bool:
    """Return True when `addresses` contain a directed cycle in `graph`."""
    cells = {normalize_address(addr) for addr in addresses}
    color: dict[str, int] = {cell: 0 for cell in cells}

    def dfs(node: str) -> bool:
        color[node] = 1
        graph_node = graph.get_node(node)
        deps = graph.get_dependencies(node) if graph_node is not None else ()
        for dep in deps:
            nxt = normalize_address(dep)
            if nxt not in cells:
                continue
            if color[nxt] == 1:
                return True
            if color[nxt] == 0 and dfs(nxt):
                return True
        color[node] = 2
        return False

    return any(color[cell] == 0 and dfs(cell) for cell in cells)


def assert_scc_cell_dag(
    scc: tuple[str, ...],
    *,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
) -> None:
    """Fail closed when a series SCC contains a cell-level cycle."""
    addresses = [addr for sid in scc for addr in catalog.get(sid).cells]
    if cells_have_cycle(addresses, graph):
        raise InvertedTreeExportError(f"cell cycle among zipper series {list(scc)!r}")


def build_scc_map(
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    graph: DependencyGraph,
) -> dict[str, tuple[str, ...]]:
    """Map each formula series to its SCC (bindings order).

    Multi-series SCCs must be cell DAGs (lag zippers). Same-year circular
    refs fail closed.
    """
    ids = [series.series_id for series in catalog.formula_series()]
    mapping: dict[str, tuple[str, ...]] = {}
    for scc in tarjan_series_sccs(ids, deps):
        if len(scc) > 1:
            assert_scc_cell_dag(scc, catalog=catalog, graph=graph)
        for sid in scc:
            mapping[sid] = scc
    return mapping


def topo_cells(
    addresses: Sequence[str],
    graph: DependencyGraph,
) -> list[str]:
    """Return `addresses` in cell-DAG order (predecessors first)."""
    ordered = [normalize_address(addr) for addr in addresses]
    cellset = set(ordered)
    succ: dict[str, list[str]] = {cell: [] for cell in ordered}
    indeg = {cell: 0 for cell in ordered}
    for cell in ordered:
        node = graph.get_node(cell)
        if node is None:
            continue
        for dep in graph.get_dependencies(cell):
            pred = normalize_address(dep)
            if pred not in cellset:
                continue
            succ[pred].append(cell)
            indeg[cell] += 1
    queue = [cell for cell in ordered if indeg[cell] == 0]
    result: list[str] = []
    while queue:
        cell = queue.pop(0)
        result.append(cell)
        for nxt in succ[cell]:
            indeg[nxt] -= 1
            if indeg[nxt] == 0:
                queue.append(nxt)
    if len(result) != len(ordered):
        raise InvertedTreeExportError("cell cycle while scheduling a zipper period")
    return result


@dataclass(frozen=True, slots=True)
class ZipperPlan:
    """Seed periods plus a homogeneous year-loop template for a lag SCC."""

    scc: tuple[str, ...]
    seed_cells: tuple[str, ...]
    loop_cells: tuple[str, ...]
    loop_n: int
    series_start: dict[str, int]


def period_column(address: str) -> int:
    """Return the 1-based column of a sheet-qualified cell (TIME_PERIOD axis)."""
    _sheet, _row, col = parse_cell_coords(address)
    return col


def plan_zipper(
    scc: tuple[str, ...],
    *,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
) -> ZipperPlan:
    """Plan a seed prefix plus a repeating within-year schedule.

    Period is the cell column. The repeating suffix must include every SCC
    series; otherwise emit would unroll Ω(years).
    """
    by_period: dict[int, list[str]] = defaultdict(list)
    sheets: set[str] = set()
    for sid in scc:
        for address in catalog.get(sid).cells:
            sheet, _row, col = parse_cell_coords(address)
            sheets.add(sheet)
            by_period[col].append(address)
    if len(sheets) != 1:
        raise InvertedTreeExportError(f"zipper series {list(scc)!r} span multiple sheets")
    if not by_period:
        raise InvertedTreeExportError(f"zipper series {list(scc)!r} have no cells")
    periods = sorted(by_period)
    period_cells = [topo_cells(by_period[period], graph) for period in periods]
    sigs = [tuple(catalog.series_id_for(cell) for cell in cells) for cells in period_cells]
    repeat = sigs[-1]
    start = len(sigs) - 1
    while start > 0 and sigs[start - 1] == repeat:
        start -= 1
    loop_n = len(sigs) - start
    loop_cells = tuple(period_cells[start])
    loop_series = {sid for sid in repeat if sid is not None}
    if (
        loop_n < 1
        or not set(scc).issubset(loop_series)
        or len(loop_cells) != len(scc)
        or set(repeat) != set(scc)
    ):
        raise InvertedTreeExportError(
            f"zipper series {list(scc)!r} are not a homogeneous year loop"
        )
    seed_cells = tuple(cell for cells in period_cells[:start] for cell in cells)
    series_start: dict[str, int] = {}
    for cell in loop_cells:
        sid = catalog.series_id_for(cell)
        if sid is None:
            continue
        idx = catalog.get(sid).index_of(cell)
        if idx is None:
            continue
        series_start[sid] = idx
    return ZipperPlan(
        scc=scc,
        seed_cells=seed_cells,
        loop_cells=loop_cells,
        loop_n=loop_n,
        series_start=series_start,
    )
