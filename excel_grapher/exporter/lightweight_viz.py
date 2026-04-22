"""Batteries-included lightweight workbook graph visualization (core + default overlays)."""

from __future__ import annotations

import heapq
from dataclasses import dataclass
from typing import Literal

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.lightweight_viz import (
    MODULE_INFERENCE_OVERLAY_ID,
    LightweightVizLayoutInput,
    LightweightVizModuleEdge,
    LightweightVizOverlay,
    LightweightVizPayload,
    VizLimits,
    _build_int_adjacencies,
    _edge_list_filtered,
    assemble_lightweight_viz_payload,
    build_lightweight_viz_core,
    derive_partition_modules_table,
)
from excel_grapher.grapher.overlay_registry import build_overlays

# --- Graph algorithms (module inference; exporter-local) --------------------


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


def iterative_kosaraju_scc(adj_out: list[list[int]], n: int) -> list[int]:
    order = _dfs_postorder_finish(adj_out, n)
    adj_rev = [[] for _ in range(n)]
    for u in range(n):
        for v in adj_out[u]:
            adj_rev[v].append(u)
    for row in adj_rev:
        row.sort()
    comp = _assign_components_reverse(adj_rev, list(reversed(order)), n)
    return comp


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


def build_condensation_edges(
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


def condensation_indegree(adj_cond: list[list[int]], n_comp: int) -> list[int]:
    indeg = [0] * n_comp
    for u in range(n_comp):
        for v in adj_cond[u]:
            indeg[v] += 1
    return indeg


def kahn_toposort(adj: list[list[int]], n: int) -> list[int] | None:
    indeg = condensation_indegree(adj, n)
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


def longest_path_ranks(adj_cond: list[list[int]], n_comp: int) -> list[int]:
    preds = [[] for _ in range(n_comp)]
    for u in range(n_comp):
        for v in adj_cond[u]:
            preds[v].append(u)
    for row in preds:
        row.sort()

    topo = kahn_toposort(adj_cond, n_comp)
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


def _module_labels_async(
    adj_cond: list[list[int]],
    n_comp: int,
    rank: list[int],
    iterations: int,
) -> list[int]:
    preds = [[] for _ in range(n_comp)]
    for u in range(n_comp):
        for v in adj_cond[u]:
            preds[v].append(u)
    for row in preds:
        row.sort()

    label = list(range(n_comp))
    for _ in range(iterations):
        order = sorted(range(n_comp), key=lambda s: (rank[s], s))
        for c in order:
            neigh = sorted(set(adj_cond[c]) | set(preds[c]))
            best_key: tuple[float, int] | None = None
            best_lbl = label[c]
            for nb in neigh:
                if nb == c:
                    continue
                dr = abs(rank[nb] - rank[c])
                w = 1.0 / (1.0 + dr)
                cand = label[nb]
                key = (-w, cand)
                if best_key is None or key < best_key:
                    best_key = key
                    best_lbl = cand
            label[c] = min(label[c], best_lbl)
    return label


def _compact_module_ids(labels: list[int]) -> tuple[list[int], int]:
    uniq = sorted(set(labels))
    m = {v: i for i, v in enumerate(uniq)}
    return [m[v] for v in labels], len(uniq)


@dataclass(frozen=True, slots=True)
class ModuleAnalysisResult:
    """Shared context for core layout + module overlay (computed once per graph)."""

    scc_count: int
    module_of: tuple[int, ...]
    node_rank: tuple[int, ...]
    module_edges: tuple[LightweightVizModuleEdge, ...]


def analyze_modules_for_viz(
    graph: DependencyGraph,
    *,
    module_iterations: int,
) -> ModuleAnalysisResult:
    keys = sorted(graph)
    n = len(keys)
    if n == 0:
        return ModuleAnalysisResult(
            scc_count=0,
            module_of=tuple(),
            node_rank=tuple(),
            module_edges=tuple(),
        )

    key_id = {k: i for i, k in enumerate(keys)}
    uncond, _all_adj = _build_int_adjacencies(graph, keys, key_id)

    comp_raw = iterative_kosaraju_scc(uncond, n)
    comp, n_comp = _remap_components(comp_raw)

    adj_cond = build_condensation_edges(uncond, n, comp, n_comp)
    scc_rank = longest_path_ranks(adj_cond, n_comp)

    scc_labels = _module_labels_async(adj_cond, n_comp, scc_rank, module_iterations)
    module_of_scc, _n_mod = _compact_module_ids(scc_labels)
    module_of = [module_of_scc[comp[i]] for i in range(n)]

    node_rank = [scc_rank[comp[i]] for i in range(n)]

    all_edges = _edge_list_filtered(graph, keys, key_id, include_guarded=True)

    mod_edge_map: dict[tuple[int, int], list[int]] = {}
    for u, v, g in all_edges:
        mu, mv = module_of[u], module_of[v]
        if mu == mv:
            continue
        key = (mu, mv)
        mod_edge_map.setdefault(key, [0, 0])
        if g:
            mod_edge_map[key][1] += 1
        else:
            mod_edge_map[key][0] += 1

    module_edges = tuple(
        LightweightVizModuleEdge(
            source_module_id=a,
            target_module_id=b,
            unconditional_weight=pair[0],
            guarded_weight=pair[1],
        )
        for (a, b), pair in sorted(mod_edge_map.items())
    )

    return ModuleAnalysisResult(
        scc_count=n_comp,
        module_of=tuple(module_of),
        node_rank=tuple(node_rank),
        module_edges=module_edges,
    )


def ensure_default_overlay_builders() -> None:
    """Register exporter overlay builders (idempotent)."""
    from excel_grapher.grapher.overlay_registry import (
        list_overlay_builders,
        register_overlay_builder,
    )

    if MODULE_INFERENCE_OVERLAY_ID in list_overlay_builders():
        return

    def _module_inference_overlay_builder(
        graph: DependencyGraph,
        core,
        *,
        context,
    ) -> LightweightVizOverlay:
        if isinstance(context, ModuleAnalysisResult):
            ma = context
        elif context is None:
            ma = analyze_modules_for_viz(graph, module_iterations=8)
        else:
            raise TypeError(
                "exporter.module_inference overlay expects ModuleAnalysisResult context or None"
            )
        modules = derive_partition_modules_table(core, ma.module_of, ma.node_rank)
        data = {
            "scc_count": ma.scc_count,
            "node_module_id": list(ma.module_of),
            "node_rank": list(ma.node_rank),
            "module_edges": [
                {
                    "source_module_id": e.source_module_id,
                    "target_module_id": e.target_module_id,
                    "unconditional_weight": e.unconditional_weight,
                    "guarded_weight": e.guarded_weight,
                }
                for e in ma.module_edges
            ],
            "modules": [
                {
                    "id": m.id,
                    "node_count": m.node_count,
                    "rank_min": m.rank_min,
                    "rank_max": m.rank_max,
                    "centroid_x": m.centroid_x,
                    "centroid_y": m.centroid_y,
                    "density_mode": m.density_mode,
                }
                for m in modules
            ],
        }
        return LightweightVizOverlay(
            overlay_id=MODULE_INFERENCE_OVERLAY_ID,
            schema_version=1,
            kind="partition",
            data=data,
            display_name="Module inference",
            supplemental_stats={"module_edge_count": len(ma.module_edges)},
        )

    register_overlay_builder(MODULE_INFERENCE_OVERLAY_ID, _module_inference_overlay_builder)


def to_lightweight_viz(
    graph: DependencyGraph,
    *,
    max_local_nodes: int | None = None,
    max_local_edges: int | None = None,
    include_guarded_edges: bool = True,
    module_iterations: int = 8,
    inline_size_budget_mb: int = 50,
    layout_mode: Literal["bfs", "layered", "grid", "force"] | None = None,
) -> LightweightVizPayload:
    """
    Build the columnar visualization payload used by ``write_lightweight_viz_html``.

    Runs module inference in the exporter layer and attaches the default partition overlay.
    """
    _ = inline_size_budget_mb
    ensure_default_overlay_builders()

    ma = analyze_modules_for_viz(graph, module_iterations=module_iterations)
    limits = VizLimits(max_local_nodes=max_local_nodes, max_local_edges=max_local_edges)
    layout = LightweightVizLayoutInput(ma.module_of, ma.node_rank)
    core_mode: Literal["bfs", "layered", "grid", "force"] = layout_mode or "bfs"
    core = build_lightweight_viz_core(
        graph,
        limits=limits,
        layout_input=layout,
        layout_mode=core_mode,
        include_guarded_edges=include_guarded_edges,
    )

    overlays = build_overlays(
        graph,
        core,
        [MODULE_INFERENCE_OVERLAY_ID],
        context=ma,
    )
    return assemble_lightweight_viz_payload(core, overlays)
