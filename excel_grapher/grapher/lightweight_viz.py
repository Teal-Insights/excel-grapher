"""Lightweight workbook graph visualization: core payload, overlays, JSON, and HTML."""

from __future__ import annotations

import heapq
import json
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal, TypeAlias

from excel_grapher.core.address_keys import (
    CellKey,
    normalize_key,
    parse_node_key,
)

from .formula_label import (
    display_formula,
    truncate_formula_display,
    validate_max_formula_length,
)
from .graph import DependencyGraph, GraphReadView
from .node import NodeKey, NodeView

VizGraph: TypeAlias = GraphReadView | DependencyGraph

# --- Constants ----------------------------------------------------------------

VIZ_PAYLOAD_VERSION = 3
WEBVIZ_LOUVAIN_DIRECTED_OVERLAY_ID = "webviz.louvain_directed"

# --- CSR / edge extraction ----------------------------------------------------


def _resolve_viz_endpoint(graph: VizGraph, dep: NodeKey) -> NodeKey | None:
    """Map an edge endpoint to a stored graph key when present."""
    nk = normalize_key(dep)
    if nk in graph:
        return nk
    return None


def _viz_label_geometry(node: NodeView) -> tuple[str, int, str]:
    """Return `(sheet, row, column)` used as the viz node label anchor."""
    parsed = parse_node_key(node.key)
    if not isinstance(parsed, CellKey):
        raise ValueError(f"Visualization requires cell nodes; got: {node.key!r}")
    return parsed.sheet, parsed.row, parsed.column


def _node_sheets(node: NodeView) -> set[str]:
    parsed = parse_node_key(node.key)
    if not isinstance(parsed, CellKey):
        raise ValueError(f"Visualization requires cell nodes; got: {node.key!r}")
    return {parsed.sheet}


def _build_int_adjacencies(
    graph: VizGraph, keys: list[NodeKey], key_id: dict[NodeKey, int]
) -> tuple[list[list[int]], list[list[int]]]:
    n = len(keys)
    uncond: list[list[int]] = [[] for _ in range(n)]
    all_e: list[list[int]] = [[] for _ in range(n)]
    for i, fk in enumerate(keys):
        for tk in graph.keys(order="workbook", source=graph.get_dependencies(fk)):
            resolved = _resolve_viz_endpoint(graph, tk)
            tid = None if resolved is None else key_id.get(resolved)
            if tid is None:
                continue
            all_e[i].append(tid)
            if not graph.is_guarded(fk, tk):
                uncond[i].append(tid)
    return uncond, all_e


def _reverse_adj(adj: list[list[int]], n: int) -> list[list[int]]:
    rev = [[] for _ in range(n)]
    for u in range(n):
        for v in adj[u]:
            rev[v].append(u)
    for row in rev:
        row.sort()
    return rev


def _edge_list_filtered(
    graph: VizGraph,
    keys: list[NodeKey],
    key_id: dict[NodeKey, int],
    *,
    include_guarded: bool,
) -> list[tuple[int, int, bool]]:
    out: list[tuple[int, int, bool]] = []
    for fk in keys:
        fi = key_id[fk]
        for tk in graph.keys(order="workbook", source=graph.get_dependencies(fk)):
            resolved = _resolve_viz_endpoint(graph, tk)
            ti = None if resolved is None else key_id.get(resolved)
            if ti is None:
                continue
            g = graph.is_guarded(fk, tk)
            if g and not include_guarded:
                continue
            out.append((fi, ti, g))
    return out


def _neighbor_sort_key(
    target: int,
    guarded: bool,
    module_of: list[int],
    src_module: int,
    out_deg: list[int],
) -> tuple[int, int, int, int]:
    same_mod = 0 if module_of[target] == src_module else 1
    return (1 if guarded else 0, same_mod, -out_deg[target], target)


def _build_local_csr(
    n: int,
    module_of: list[int],
    out_edges_by_src: list[list[tuple[int, bool]]],
    mod_node_count: list[int],
    mod_internal_edges: list[int],
    max_local_nodes: int,
    max_local_edges: int,
) -> tuple[list[int], list[int], list[bool], list[bool]]:
    offsets = [0] * (n + 1)
    targets: list[int] = []
    guarded_flags: list[bool] = []
    complete = [True] * n
    out_deg = [len(out_edges_by_src[i]) for i in range(n)]

    for src in range(n):
        m = module_of[src]
        small_module = (
            mod_node_count[m] <= max_local_nodes and mod_internal_edges[m] <= max_local_edges
        )
        raw = list(out_edges_by_src[src])
        if small_module:
            raw.sort(key=lambda t: _neighbor_sort_key(t[0], t[1], module_of, m, out_deg))
            for tgt, g in raw:
                targets.append(tgt)
                guarded_flags.append(g)
        else:
            raw.sort(key=lambda t: _neighbor_sort_key(t[0], t[1], module_of, m, out_deg))
            for k, (tgt, g) in enumerate(raw):
                if k >= max_local_edges:
                    complete[src] = False
                    break
                targets.append(tgt)
                guarded_flags.append(g)
            if len(raw) > max_local_edges:
                complete[src] = False
        offsets[src + 1] = len(targets)

    return offsets, targets, guarded_flags, complete


def _resolve_local_limits(
    n: int,
    total_out_edges: int,
    max_local_nodes: int | None,
    max_local_edges: int | None,
) -> tuple[int, int]:
    max_nodes_eff = n if max_local_nodes is None else max(0, max_local_nodes)
    max_edges_eff = total_out_edges if max_local_edges is None else max(0, max_local_edges)
    return max_nodes_eff, max_edges_eff


# --- Shared edge column type --------------------------------------------------


@dataclass(frozen=True, slots=True)
class LightweightVizLocalEdges:
    offsets: tuple[int, ...]
    targets: tuple[int, ...]
    guarded: tuple[bool, ...]
    complete: tuple[bool, ...]


# --- Core payload (structural + layout) ---------------------------------------


@dataclass(frozen=True, slots=True)
class VizLimits:
    max_local_nodes: int | None = None
    max_local_edges: int | None = None


@dataclass(frozen=True, slots=True)
class LightweightVizLayoutInput:
    module_of: tuple[int, ...]
    node_rank: tuple[int, ...]


@dataclass(frozen=True, slots=True)
class LightweightVizCoreStats:
    node_count: int
    local_edge_count: int
    truncated_local_nodes: int


@dataclass(frozen=True, slots=True)
class LightweightVizCoreNodeColumns:
    sheet_index: tuple[int, ...]
    row: tuple[int, ...]
    column: tuple[str, ...]
    is_leaf: tuple[bool, ...]
    formula: tuple[str | None, ...]
    in_degree: tuple[int, ...]
    out_degree: tuple[int, ...]
    rank: tuple[int, ...]
    depth: tuple[int, ...]
    x: tuple[float, ...]
    y: tuple[float, ...]


@dataclass(frozen=True, slots=True)
class LightweightVizCore:
    stats: LightweightVizCoreStats
    sheets: tuple[str, ...]
    nodes: LightweightVizCoreNodeColumns
    local_edges: LightweightVizLocalEdges
    max_local_nodes: int | None
    max_local_edges: int | None


def _dfs_postorder_finish(adj: list[list[int]], n: int) -> list[int]:
    visited = [False] * n
    order: list[int] = []
    for start in range(n):
        if visited[start]:
            continue
        stack: list[tuple[int, int]] = [(start, 0)]
        visited[start] = True
        while stack:
            v, ni = stack[-1]
            nbrs = adj[v]
            if ni < len(nbrs):
                w = nbrs[ni]
                stack[-1] = (v, ni + 1)
                if not visited[w]:
                    visited[w] = True
                    stack.append((w, 0))
            else:
                stack.pop()
                order.append(v)
    return order


def _assign_components_reverse(adj_rev: list[list[int]], order_rev: list[int], n: int) -> list[int]:
    comp = [-1] * n
    label = 0
    for start in order_rev:
        if comp[start] >= 0:
            continue
        stack = [start]
        comp[start] = label
        while stack:
            v = stack.pop()
            for w in adj_rev[v]:
                if comp[w] < 0:
                    comp[w] = label
                    stack.append(w)
        label += 1
    return comp


def _iterative_kosaraju_scc(adj_out: list[list[int]], n: int) -> list[int]:
    order = _dfs_postorder_finish(adj_out, n)
    adj_rev = [[] for _ in range(n)]
    for u in range(n):
        for v in adj_out[u]:
            adj_rev[v].append(u)
    for row in adj_rev:
        row.sort()
    return _assign_components_reverse(adj_rev, list(reversed(order)), n)


def _remap_components(comp: list[int]) -> tuple[list[int], int]:
    mapping: dict[int, int] = {}
    out = [0] * len(comp)
    nxt = 0
    for i, c in enumerate(comp):
        if c not in mapping:
            mapping[c] = nxt
            nxt += 1
        out[i] = mapping[c]
    return out, nxt


def _build_condensation_edges(
    adj: list[list[int]], n: int, comp: list[int], n_comp: int
) -> list[list[int]]:
    edges: set[tuple[int, int]] = set()
    for u in range(n):
        cu = comp[u]
        for v in adj[u]:
            cv = comp[v]
            if cu != cv:
                edges.add((cu, cv))
    cond = [[] for _ in range(n_comp)]
    for a, b in sorted(edges):
        cond[a].append(b)
    return cond


def _condensation_indegree(adj_cond: list[list[int]], n_comp: int) -> list[int]:
    indeg = [0] * n_comp
    for u in range(n_comp):
        for v in adj_cond[u]:
            indeg[v] += 1
    return indeg


def _kahn_toposort(adj: list[list[int]], n: int) -> list[int] | None:
    indeg = _condensation_indegree(adj, n)
    heap = [i for i in range(n) if indeg[i] == 0]
    heapq.heapify(heap)
    order: list[int] = []
    while heap:
        u = heapq.heappop(heap)
        order.append(u)
        for v in adj[u]:
            indeg[v] -= 1
            if indeg[v] == 0:
                heapq.heappush(heap, v)
    if len(order) != n:
        return None
    return order


def _longest_path_ranks(adj_cond: list[list[int]], n_comp: int) -> list[int]:
    preds = [[] for _ in range(n_comp)]
    for u in range(n_comp):
        for v in adj_cond[u]:
            preds[v].append(u)
    for row in preds:
        row.sort()
    topo = _kahn_toposort(adj_cond, n_comp)
    if topo is None:
        return [0] * n_comp
    rank = [0] * n_comp
    for v in topo:
        pr = preds[v]
        if not pr:
            rank[v] = 0
        else:
            rank[v] = max(rank[u] + 1 for u in pr)
    return rank


def unconditional_scc_ranks(uncond: list[list[int]], n: int) -> tuple[list[int], int]:
    """Rank nodes by longest path through the SCC condensation of `uncond`.

    Args:
        uncond: Out-adjacency (unconditional edges only), indexed by node id.
        n: Node count; `uncond` must have this length.

    Returns:
        `(ranks, scc_count)`, where `ranks[i]` is the condensation rank of node
        `i` (cycle members share a rank) and `scc_count` is the number of
        strongly connected components.
    """
    if n == 0:
        return [], 0
    comp_raw = _iterative_kosaraju_scc(uncond, n)
    comp, n_comp = _remap_components(comp_raw)
    adj_cond = _build_condensation_edges(uncond, n, comp, n_comp)
    comp_rank = _longest_path_ranks(adj_cond, n_comp)
    return [comp_rank[comp[i]] for i in range(n)], n_comp


def build_lightweight_viz_core(
    graph: VizGraph,
    *,
    limits: VizLimits | None = None,
    layout_input: LightweightVizLayoutInput | None = None,
    positions: Sequence[tuple[float, float]] | None = None,
    include_guarded_edges: bool = True,
    include_formula_on_nodes: bool = True,
    max_formula_length: int | None = 120,
) -> LightweightVizCore:
    """Build the cell-level core columns of a web-viz payload.

    Args:
        graph: Cell graph (or read view) to draw.
        limits: Local-neighborhood export caps.
        layout_input: Module per node and node rank (e.g. Louvain + SCC rank).
            Defaults to one module and SCC-condensation rank.
        positions: `(x, y)` per node in `graph.keys(order="workbook")` order.
            When omitted, runs the shared clustered force layout with modules
            as clusters and the default vertical pull.
        include_guarded_edges: Count guarded edges in degrees, local edges, and
            the default layout.
        include_formula_on_nodes: Store display formulas on nodes.
        max_formula_length: Truncate display formulas to this length.

    Raises:
        ValueError: `layout_input` or `positions` length does not match the
            node count.
    """
    from .viz_layout import DEFAULT_RANK_PULL, clustered_force_layout, input_depths

    validate_max_formula_length(max_formula_length)

    lim = limits or VizLimits()
    keys = graph.keys(order="workbook")
    n = len(keys)
    if positions is not None and len(positions) != n:
        raise ValueError("positions must have one (x, y) per graph node")
    if n == 0:
        return LightweightVizCore(
            stats=LightweightVizCoreStats(
                node_count=0,
                local_edge_count=0,
                truncated_local_nodes=0,
            ),
            sheets=tuple(),
            nodes=LightweightVizCoreNodeColumns(
                sheet_index=tuple(),
                row=tuple(),
                column=tuple(),
                is_leaf=tuple(),
                formula=tuple(),
                in_degree=tuple(),
                out_degree=tuple(),
                rank=tuple(),
                depth=tuple(),
                x=tuple(),
                y=tuple(),
            ),
            local_edges=LightweightVizLocalEdges(
                offsets=(0,),
                targets=tuple(),
                guarded=tuple(),
                complete=tuple(),
            ),
            max_local_nodes=lim.max_local_nodes,
            max_local_edges=lim.max_local_edges,
        )

    key_id = {k: i for i, k in enumerate(keys)}
    present_sheets: set[str] = set()
    for k in keys:
        node = graph.get_node(k)
        if node is not None:
            present_sheets.update(_node_sheets(node))
    if graph.sheet_order is not None:
        sheets_sorted = [sheet for sheet in graph.sheet_order if sheet in present_sheets]
        unknown_sheets = sorted(present_sheets - set(sheets_sorted))
        sheets_sorted.extend(unknown_sheets)
    else:
        sheets_sorted = sorted(present_sheets)
    sheet_index_map = {s: i for i, s in enumerate(sheets_sorted)}
    uncond, all_adj = _build_int_adjacencies(graph, keys, key_id)
    selected_adj = all_adj if include_guarded_edges else uncond
    rev_selected = _reverse_adj(selected_adj, n)

    in_deg = [len(rev_selected[i]) for i in range(n)]
    out_deg = [len(selected_adj[i]) for i in range(n)]

    if layout_input is None:
        module_of = [0] * n
        ranks, _scc_count = unconditional_scc_ranks(uncond, n)
    else:
        if len(layout_input.module_of) != n or len(layout_input.node_rank) != n:
            raise ValueError("layout_input tuple lengths must match graph order")
        module_of = list(layout_input.module_of)
        ranks = list(layout_input.node_rank)
    depths = input_depths(n, [(u, v) for u in range(n) for v in uncond[u]])

    if positions is None:
        pos = clustered_force_layout(
            n,
            [(u, v) for u in range(n) for v in selected_adj[u]],
            module_of,
            depths=depths,
            rank_pull=DEFAULT_RANK_PULL,
        )
        xs = [float(v) for v in pos[:, 0]]
        ys = [float(v) for v in pos[:, 1]]
    else:
        xs = [float(x) for x, _y in positions]
        ys = [float(y) for _x, y in positions]

    n_mod = max(module_of) + 1 if module_of else 0
    mod_node_count = [0] * n_mod
    for m in module_of:
        mod_node_count[m] += 1

    all_edges = _edge_list_filtered(graph, keys, key_id, include_guarded=include_guarded_edges)
    mod_internal_edges = [0] * n_mod
    for u, v, _ in all_edges:
        if module_of[u] == module_of[v]:
            mod_internal_edges[module_of[u]] += 1

    out_edges_by_src: list[list[tuple[int, bool]]] = [[] for _ in range(n)]
    for u, v, g in all_edges:
        out_edges_by_src[u].append((v, g))

    total_out_edges = sum(len(row) for row in out_edges_by_src)
    max_local_nodes_eff, max_local_edges_eff = _resolve_local_limits(
        n, total_out_edges, lim.max_local_nodes, lim.max_local_edges
    )
    offsets, loc_tgts, loc_guarded, loc_complete = _build_local_csr(
        n,
        module_of,
        out_edges_by_src,
        mod_node_count,
        mod_internal_edges,
        max_local_nodes_eff,
        max_local_edges_eff,
    )
    truncated_local = sum(1 for c in loc_complete if not c)
    local_edge_count = len(loc_tgts)

    rows: list[int] = []
    cols: list[str] = []
    sheet_ix: list[int] = []
    is_leaf: list[bool] = []
    formulas: list[str | None] = []
    for k in keys:
        node = graph.get_node(k)
        assert node is not None
        sheet, row, col = _viz_label_geometry(node)
        rows.append(row)
        cols.append(col)
        sheet_ix.append(sheet_index_map[sheet])
        is_leaf.append(node.is_leaf)
        shown_formula = display_formula(node)
        if include_formula_on_nodes and shown_formula:
            formulas.append(truncate_formula_display(shown_formula, max_formula_length))
        else:
            formulas.append(None)

    stats = LightweightVizCoreStats(
        node_count=n,
        local_edge_count=local_edge_count,
        truncated_local_nodes=truncated_local,
    )

    nodes = LightweightVizCoreNodeColumns(
        sheet_index=tuple(sheet_ix),
        row=tuple(rows),
        column=tuple(cols),
        is_leaf=tuple(is_leaf),
        formula=tuple(formulas),
        in_degree=tuple(in_deg),
        out_degree=tuple(out_deg),
        rank=tuple(ranks),
        depth=tuple(depths),
        x=tuple(xs),
        y=tuple(ys),
    )

    local_edges = LightweightVizLocalEdges(
        offsets=tuple(offsets),
        targets=tuple(loc_tgts),
        guarded=tuple(loc_guarded),
        complete=tuple(loc_complete),
    )

    return LightweightVizCore(
        stats=stats,
        sheets=tuple(sheets_sorted),
        nodes=nodes,
        local_edges=local_edges,
        max_local_nodes=lim.max_local_nodes,
        max_local_edges=lim.max_local_edges,
    )


# --- Overlays + wire payload --------------------------------------------------


@dataclass(frozen=True, slots=True)
class LightweightVizOverlay:
    overlay_id: str
    schema_version: int
    kind: str
    data: Mapping[str, Any]
    display_name: str | None = None
    default_visible: bool = True
    supplemental_stats: Mapping[str, Any] | None = None


@dataclass(frozen=True, slots=True)
class LightweightVizPayload:
    version: int
    core: LightweightVizCore
    overlays: tuple[LightweightVizOverlay, ...]
    annotations: Mapping[str, Any] | None = None
    viewer_hints: Mapping[str, Any] | None = None


def assemble_lightweight_viz_payload(
    core: LightweightVizCore,
    overlays: Sequence[LightweightVizOverlay],
    *,
    annotations: Mapping[str, Any] | None = None,
    viewer_hints: Mapping[str, Any] | None = None,
) -> LightweightVizPayload:
    return LightweightVizPayload(
        version=VIZ_PAYLOAD_VERSION,
        core=core,
        overlays=tuple(overlays),
        annotations=annotations,
        viewer_hints=viewer_hints,
    )


# --- Partition modules ----------------------------------------------------------


@dataclass(frozen=True, slots=True)
class LightweightVizModule:
    id: int
    node_count: int
    rank_min: int
    rank_max: int
    centroid_x: float
    centroid_y: float


@dataclass(frozen=True, slots=True)
class LightweightVizModuleEdge:
    source_module_id: int
    target_module_id: int
    unconditional_weight: int
    guarded_weight: int


def derive_partition_modules_table(
    core: LightweightVizCore,
    module_id: tuple[int, ...],
    node_rank: tuple[int, ...],
) -> tuple[LightweightVizModule, ...]:
    n = core.stats.node_count
    if n == 0:
        return tuple()

    n_mod = max(module_id) + 1
    mod_node_count = [0] * n_mod
    for m in module_id:
        mod_node_count[m] += 1

    xs = list(core.nodes.x)
    ys = list(core.nodes.y)

    mod_rank_min = [10**9] * n_mod
    mod_rank_max = [-1] * n_mod
    sum_x = [0.0] * n_mod
    sum_y = [0.0] * n_mod
    for i in range(n):
        m = module_id[i]
        r = node_rank[i]
        mod_rank_min[m] = min(mod_rank_min[m], r)
        mod_rank_max[m] = max(mod_rank_max[m], r)
        sum_x[m] += xs[i]
        sum_y[m] += ys[i]

    modules: list[LightweightVizModule] = []
    for m in range(n_mod):
        c = mod_node_count[m]
        modules.append(
            LightweightVizModule(
                id=m,
                node_count=c,
                rank_min=mod_rank_min[m] if c else 0,
                rank_max=mod_rank_max[m] if c else 0,
                centroid_x=sum_x[m] / c if c else 0.0,
                centroid_y=sum_y[m] / c if c else 0.0,
            )
        )
    return tuple(modules)


# --- Serialization ------------------------------------------------------------


def lightweight_viz_overlay_to_jsonable(overlay: LightweightVizOverlay) -> dict[str, Any]:
    out: dict[str, Any] = {
        "overlay_id": overlay.overlay_id,
        "schema_version": overlay.schema_version,
        "kind": overlay.kind,
        "data": dict(overlay.data),
    }
    if overlay.display_name is not None:
        out["display_name"] = overlay.display_name
    out["default_visible"] = overlay.default_visible
    if overlay.supplemental_stats is not None:
        out["supplemental_stats"] = dict(overlay.supplemental_stats)
    return out


def _core_to_jsonable(c: LightweightVizCore) -> dict[str, Any]:
    nc = c.nodes
    return {
        "stats": {
            "node_count": c.stats.node_count,
            "local_edge_count": c.stats.local_edge_count,
            "truncated_local_nodes": c.stats.truncated_local_nodes,
        },
        "sheets": list(c.sheets),
        "nodes": {
            "sheet_index": list(nc.sheet_index),
            "row": list(nc.row),
            "column": list(nc.column),
            "is_leaf": list(nc.is_leaf),
            "formula": list(nc.formula),
            "in_degree": list(nc.in_degree),
            "out_degree": list(nc.out_degree),
            "rank": list(nc.rank),
            "depth": list(nc.depth),
            "x": list(nc.x),
            "y": list(nc.y),
        },
        "local_edges": {
            "offsets": list(c.local_edges.offsets),
            "targets": list(c.local_edges.targets),
            "guarded": list(c.local_edges.guarded),
            "complete": list(c.local_edges.complete),
        },
        "max_local_nodes": c.max_local_nodes,
        "max_local_edges": c.max_local_edges,
    }


def _payload_to_jsonable(payload: LightweightVizPayload) -> dict[str, Any]:
    d: dict[str, Any] = {
        "version": payload.version,
        "core": _core_to_jsonable(payload.core),
        "overlays": [lightweight_viz_overlay_to_jsonable(o) for o in payload.overlays],
    }
    if payload.annotations is not None:
        d["annotations"] = dict(payload.annotations)
    if payload.viewer_hints is not None:
        d["viewer_hints"] = dict(payload.viewer_hints)
    return d


def estimate_serialized_json_bytes(payload: LightweightVizPayload) -> int:
    n = payload.core.stats.node_count
    e_loc = payload.core.stats.local_edge_count
    est = 4000
    est += n * (6 * 11 + 12 * 2 + 5 + 8 * 4)
    est += e_loc * 12
    est += sum(len(o.data.get("module_edges", ())) for o in payload.overlays) * 40
    est += sum(len(o.data.get("modules", ())) for o in payload.overlays) * 60
    est += sum(len(s) for s in payload.core.sheets) + n * 4
    est += sum(len(f or "") for f in payload.core.nodes.formula)
    return int(est * 1.15)


def serialize_lightweight_viz_json(payload: LightweightVizPayload) -> str:
    if payload.version != VIZ_PAYLOAD_VERSION:
        raise ValueError(f"Unsupported lightweight viz payload version: {payload.version}")
    return json.dumps(_payload_to_jsonable(payload), separators=(",", ":"))


def write_lightweight_viz_data(payload: LightweightVizPayload, path: Path | str) -> None:
    p = Path(path)
    p.write_text(serialize_lightweight_viz_json(payload), encoding="utf-8")


def write_web_viz_html(
    payload: LightweightVizPayload,
    path: Path | str,
    *,
    title: str = "Workbook dependency graph",
    data_mode: Literal["inline", "sidecar", "auto"] = "auto",
    data_path: Path | str | None = None,
    inline_size_budget_mb: int = 50,
    template_path: Path | str | None = None,
) -> None:
    """Write a web visualization HTML bundle from a web-viz payload."""
    from importlib import resources

    from .viz_layout import viz_layout_js_bootstrap

    if payload.version != VIZ_PAYLOAD_VERSION:
        raise ValueError(f"Unsupported lightweight viz payload version: {payload.version}")

    out = Path(path)
    budget = max(0, inline_size_budget_mb) * 1024 * 1024
    json_payload: str | None = None
    sidecar_name: str | None = None

    if data_mode == "inline":
        json_payload = serialize_lightweight_viz_json(payload)
    elif data_mode == "sidecar":
        if data_path is None:
            sidecar_name = out.with_suffix(".viz.json").name
        else:
            sidecar_name = Path(data_path).name
        json_payload = None
    else:
        est = estimate_serialized_json_bytes(payload)
        if est <= budget:
            json_payload = serialize_lightweight_viz_json(payload)
        else:
            sidecar_name = (
                Path(data_path).name if data_path is not None else out.with_suffix(".viz.json").name
            )

    if json_payload is not None and len(json_payload.encode("utf-8")) > budget:
        sidecar_name = (
            Path(data_path).name if data_path is not None else out.with_suffix(".viz.json").name
        )
        json_payload = None

    if json_payload is None:
        if sidecar_name is None:
            sidecar_name = out.with_suffix(".viz.json").name
        data_file = out.parent / sidecar_name
        write_lightweight_viz_data(payload, data_file)

    if template_path is None:
        tpl = (
            resources.files(__package__ or __name__)
            .joinpath("lightweight_viz_template.html")
            .read_text(encoding="utf-8")
        )
    else:
        tpl = Path(template_path).read_text(encoding="utf-8")
    bootstrap = (
        f"window.__VIZ_DATA__ = {json_payload};"
        if json_payload is not None
        else "window.__VIZ_DATA__ = null;"
    )
    sidecar_js = (
        f"window.__VIZ_DATA_URL__ = {json.dumps(sidecar_name)};"
        if json_payload is None
        else "window.__VIZ_DATA_URL__ = null;"
    )
    html = (
        tpl.replace("__TITLE__", title)
        .replace("/*__BOOTSTRAP__*/", bootstrap)
        .replace("/*__SIDECAR__*/", sidecar_js)
        .replace("/*__VIZ_FORCE_JS__*/", viz_layout_js_bootstrap())
    )
    out.write_text(html, encoding="utf-8")
