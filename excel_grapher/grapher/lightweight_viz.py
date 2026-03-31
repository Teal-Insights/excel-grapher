from __future__ import annotations

import heapq
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

from .graph import DependencyGraph
from .node import NodeKey

# --- Public payload types -----------------------------------------------------

VIZ_PAYLOAD_VERSION = 1

# Overview layout: nodes in the same (rank, module) bucket beyond this count use density metadata.
DENSE_BUCKET_THRESHOLD = 12


@dataclass(frozen=True, slots=True)
class LightweightVizStats:
    node_count: int
    scc_count: int
    module_count: int
    module_edge_count: int
    local_edge_count: int
    truncated_local_nodes: int
    dense_bucket_count: int


@dataclass(frozen=True, slots=True)
class LightweightVizNodeColumns:
    sheet_index: tuple[int, ...]
    row: tuple[int, ...]
    column: tuple[str, ...]
    is_leaf: tuple[bool, ...]
    in_degree: tuple[int, ...]
    out_degree: tuple[int, ...]
    module_id: tuple[int, ...]
    rank: tuple[int, ...]
    x: tuple[float, ...]
    y: tuple[float, ...]
    bucket_density: tuple[int, ...]


@dataclass(frozen=True, slots=True)
class LightweightVizModule:
    id: int
    node_count: int
    rank_min: int
    rank_max: int
    centroid_x: float
    centroid_y: float
    density_mode: bool


@dataclass(frozen=True, slots=True)
class LightweightVizModuleEdge:
    source_module_id: int
    target_module_id: int
    unconditional_weight: int
    guarded_weight: int


@dataclass(frozen=True, slots=True)
class LightweightVizLocalEdges:
    offsets: tuple[int, ...]
    targets: tuple[int, ...]
    guarded: tuple[bool, ...]
    complete: tuple[bool, ...]


@dataclass(frozen=True, slots=True)
class LightweightVizPayload:
    version: int
    stats: LightweightVizStats
    sheets: tuple[str, ...]
    nodes: LightweightVizNodeColumns
    modules: tuple[LightweightVizModule, ...]
    module_edges: tuple[LightweightVizModuleEdge, ...]
    local_edges: LightweightVizLocalEdges
    max_local_nodes: int
    max_local_edges: int


@dataclass(frozen=True, slots=True)
class LocalForceSubgraph:
    """Explicit bounded subgraph for client-side force layout."""

    node_ids: tuple[int, ...]
    edges_from: tuple[int, ...]
    edges_to: tuple[int, ...]
    edges_guarded: tuple[bool, ...]
    is_module_scope: bool
    truncated: bool


# --- Iterative graph algorithms (no recursion) -------------------------------


def _dfs_postorder_finish(adj: list[list[int]], n: int) -> list[int]:
    """Iterative postorder; adjacency lists must be sorted for determinism."""
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
    """
    Strongly connected components as integer labels 0..k-1 (not necessarily topo order).
    Uses Kosaraju with iterative DFS only.
    """
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
    """Remap arbitrary SCC labels to 0..c-1 in order of first appearance in node order."""
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
    """
    For a DAG, rank[v] = max(rank[u]+1) over predecessors u (sources rank 0).
    Deterministic: process in Kahn topological order.
    """
    preds = [[] for _ in range(n_comp)]
    for u in range(n_comp):
        for v in adj_cond[u]:
            preds[v].append(u)
    for row in preds:
        row.sort()

    topo = kahn_toposort(adj_cond, n_comp)
    if topo is None:
        # Should not happen for condensation of SCCs; fall back to zero ranks.
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


# --- Core export --------------------------------------------------------------


def _build_int_adjacencies(
    graph: DependencyGraph, keys: list[NodeKey], key_id: dict[NodeKey, int]
) -> tuple[list[list[int]], list[list[int]]]:
    n = len(keys)
    uncond: list[list[int]] = [[] for _ in range(n)]
    all_e: list[list[int]] = [[] for _ in range(n)]
    for i, fk in enumerate(keys):
        for tk in sorted(graph.dependencies(fk)):
            tid = key_id.get(tk)
            if tid is None:
                continue
            all_e[i].append(tid)
            if graph.edge_attrs(fk, tk).get("guard") is None:
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


def _edge_list_all(
    graph: DependencyGraph, keys: list[NodeKey], key_id: dict[NodeKey, int]
) -> list[tuple[int, int, bool]]:
    out: list[tuple[int, int, bool]] = []
    for fk in keys:
        fi = key_id[fk]
        for tk in sorted(graph.dependencies(fk)):
            ti = key_id.get(tk)
            if ti is None:
                continue
            g = graph.edge_attrs(fk, tk).get("guard") is not None
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

    for src in range(n):
        m = module_of[src]
        small_module = (
            mod_node_count[m] <= max_local_nodes and mod_internal_edges[m] <= max_local_edges
        )
        raw = list(out_edges_by_src[src])
        out_deg = [len(out_edges_by_src[i]) for i in range(n)]
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


def to_lightweight_viz(
    graph: DependencyGraph,
    *,
    max_local_nodes: int = 5000,
    max_local_edges: int = 20000,
    module_iterations: int = 8,
    inline_size_budget_mb: int = 50,
) -> LightweightVizPayload:
    """
    Build a columnar, deterministic visualization payload for large dependency graphs.

    ``inline_size_budget_mb`` is reserved for HTML writers that choose inline vs sidecar mode;
    it does not change the payload structure.
    """
    _ = inline_size_budget_mb
    keys = sorted(graph)
    n = len(keys)
    if n == 0:
        return LightweightVizPayload(
            version=VIZ_PAYLOAD_VERSION,
            stats=LightweightVizStats(
                node_count=0,
                scc_count=0,
                module_count=0,
                module_edge_count=0,
                local_edge_count=0,
                truncated_local_nodes=0,
                dense_bucket_count=0,
            ),
            sheets=tuple(),
            nodes=LightweightVizNodeColumns(
                sheet_index=tuple(),
                row=tuple(),
                column=tuple(),
                is_leaf=tuple(),
                in_degree=tuple(),
                out_degree=tuple(),
                module_id=tuple(),
                rank=tuple(),
                x=tuple(),
                y=tuple(),
                bucket_density=tuple(),
            ),
            modules=tuple(),
            module_edges=tuple(),
            local_edges=LightweightVizLocalEdges(
                offsets=(0,),
                targets=tuple(),
                guarded=tuple(),
                complete=tuple(),
            ),
            max_local_nodes=max_local_nodes,
            max_local_edges=max_local_edges,
        )
    key_id = {k: i for i, k in enumerate(keys)}

    sheets_sorted = sorted({node.sheet for k in keys if (node := graph.get_node(k)) is not None})
    sheet_index_map = {s: i for i, s in enumerate(sheets_sorted)}

    uncond, all_adj = _build_int_adjacencies(graph, keys, key_id)
    rev_all = _reverse_adj(all_adj, n)

    in_deg = [len(rev_all[i]) for i in range(n)]
    out_deg = [len(all_adj[i]) for i in range(n)]

    comp_raw = iterative_kosaraju_scc(uncond, n) if n else []
    comp, n_comp = _remap_components(comp_raw) if n else ([], 0)

    adj_cond = build_condensation_edges(uncond, n, comp, n_comp) if n else []
    scc_rank = longest_path_ranks(adj_cond, n_comp) if n_comp else []

    scc_labels = (
        _module_labels_async(adj_cond, n_comp, scc_rank, module_iterations) if n_comp else []
    )
    module_of_scc, _n_mod = _compact_module_ids(scc_labels) if n_comp else ([], 0)
    module_of = [module_of_scc[comp[i]] for i in range(n)] if n else []

    node_rank = [scc_rank[comp[i]] for i in range(n)] if n else []

    n_mod = _n_mod if n_comp else 0
    mod_node_count = [0] * n_mod
    for m in module_of:
        mod_node_count[m] += 1

    all_edges = _edge_list_all(graph, keys, key_id)
    mod_internal_edges = [0] * n_mod
    for u, v, _ in all_edges:
        if module_of[u] == module_of[v]:
            mod_internal_edges[module_of[u]] += 1

    out_edges_by_src: list[list[tuple[int, bool]]] = [[] for _ in range(n)]
    for u, v, g in all_edges:
        out_edges_by_src[u].append((v, g))

    offsets, loc_tgts, loc_guarded, loc_complete = _build_local_csr(
        n,
        module_of,
        out_edges_by_src,
        mod_node_count,
        mod_internal_edges,
        max_local_nodes,
        max_local_edges,
    )
    truncated_local = sum(1 for c in loc_complete if not c)
    local_edge_count = len(loc_tgts)

    x_scale = 120.0
    y_band = 36.0

    bucket_counts: dict[tuple[int, int], int] = {}
    for i in range(n):
        b = (node_rank[i], module_of[i])
        bucket_counts[b] = bucket_counts.get(b, 0) + 1

    dense_bucket_count = sum(1 for _b, c in bucket_counts.items() if c > DENSE_BUCKET_THRESHOLD)

    xs = [0.0] * n
    ys = [0.0] * n
    bucket_density = [0] * n

    for i in range(n):
        rnk = node_rank[i]
        mid = module_of[i]
        xs[i] = float(rnk) * x_scale
        base_y = float(mid) * y_band
        bkey = (rnk, mid)
        cnt = bucket_counts[bkey]
        bucket_density[i] = cnt
        idx_in_bucket = sum(1 for j in range(i) if node_rank[j] == rnk and module_of[j] == mid)
        if cnt <= DENSE_BUCKET_THRESHOLD:
            ys[i] = base_y + float(idx_in_bucket) * 4.0
        else:
            # Deterministic jitter inside the band for overloaded buckets.
            t = (i * 1103515245 + 12345) & 0x7FFFFFFF
            jx = ((t % 10000) / 10000.0 - 0.5) * y_band * 0.85
            jy = (((t // 10000) % 10000) / 10000.0 - 0.5) * 8.0
            ys[i] = base_y + jx + jy

    mod_rank_min = [10**9] * n_mod
    mod_rank_max = [-1] * n_mod
    sum_x = [0.0] * n_mod
    sum_y = [0.0] * n_mod
    for i in range(n):
        m = module_of[i]
        r = node_rank[i]
        mod_rank_min[m] = min(mod_rank_min[m], r)
        mod_rank_max[m] = max(mod_rank_max[m], r)
        sum_x[m] += xs[i]
        sum_y[m] += ys[i]

    modules: list[LightweightVizModule] = []
    for m in range(n_mod):
        c = mod_node_count[m]
        density_mode = any(
            bucket_counts.get((r, m), 0) > DENSE_BUCKET_THRESHOLD
            for r in range(mod_rank_min[m], mod_rank_max[m] + 1)
        )
        modules.append(
            LightweightVizModule(
                id=m,
                node_count=c,
                rank_min=mod_rank_min[m] if c else 0,
                rank_max=mod_rank_max[m] if c else 0,
                centroid_x=sum_x[m] / c if c else 0.0,
                centroid_y=sum_y[m] / c if c else 0.0,
                density_mode=density_mode,
            )
        )

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

    rows: list[int] = []
    cols: list[str] = []
    sheet_ix: list[int] = []
    is_leaf: list[bool] = []
    for k in keys:
        node = graph.get_node(k)
        assert node is not None
        rows.append(node.row)
        cols.append(node.column)
        sheet_ix.append(sheet_index_map[node.sheet])
        is_leaf.append(node.is_leaf)

    stats = LightweightVizStats(
        node_count=n,
        scc_count=n_comp,
        module_count=n_mod,
        module_edge_count=len(module_edges),
        local_edge_count=local_edge_count,
        truncated_local_nodes=truncated_local,
        dense_bucket_count=dense_bucket_count,
    )

    nodes = LightweightVizNodeColumns(
        sheet_index=tuple(sheet_ix),
        row=tuple(rows),
        column=tuple(cols),
        is_leaf=tuple(is_leaf),
        in_degree=tuple(in_deg),
        out_degree=tuple(out_deg),
        module_id=tuple(module_of),
        rank=tuple(node_rank),
        x=tuple(xs),
        y=tuple(ys),
        bucket_density=tuple(bucket_density),
    )

    local_edges = LightweightVizLocalEdges(
        offsets=tuple(offsets),
        targets=tuple(loc_tgts),
        guarded=tuple(loc_guarded),
        complete=tuple(loc_complete),
    )

    return LightweightVizPayload(
        version=VIZ_PAYLOAD_VERSION,
        stats=stats,
        sheets=tuple(sheets_sorted),
        nodes=nodes,
        modules=tuple(modules),
        module_edges=module_edges,
        local_edges=local_edges,
        max_local_nodes=max_local_nodes,
        max_local_edges=max_local_edges,
    )


def select_local_force_subgraph(
    payload: LightweightVizPayload,
    *,
    node_id: int,
) -> LocalForceSubgraph:
    """
    Choose a node-induced subgraph for d3-force: whole module if small enough,
    otherwise a bounded 1-hop neighborhood around ``node_id``.
    """
    n = payload.stats.node_count
    if not (0 <= node_id < n):
        raise ValueError(f"node_id out of range: {node_id}")

    mid = payload.nodes.module_id[node_id]
    mod = payload.modules[mid]
    if mod.node_count <= payload.max_local_nodes:
        nodes = [i for i in range(n) if payload.nodes.module_id[i] == mid]
        nodes.sort()
        node_set = set(nodes)
        ef: list[int] = []
        et: list[int] = []
        eg: list[bool] = []
        off = payload.local_edges.offsets
        tg = payload.local_edges.targets
        gd = payload.local_edges.guarded
        for u in nodes:
            for k in range(off[u], off[u + 1]):
                v = tg[k]
                if v in node_set:
                    ef.append(u)
                    et.append(v)
                    eg.append(gd[k])
        return LocalForceSubgraph(
            node_ids=tuple(nodes),
            edges_from=tuple(ef),
            edges_to=tuple(et),
            edges_guarded=tuple(eg),
            is_module_scope=True,
            truncated=not payload.local_edges.complete[node_id],
        )

    # k-hop: start with 1-hop from node_id using exported local edges only.
    off = payload.local_edges.offsets
    tg = payload.local_edges.targets
    gd = payload.local_edges.guarded
    seeds = {node_id}
    expanded: set[int] = set()
    edges_from: list[int] = []
    edges_to: list[int] = []
    edges_guarded: list[bool] = []
    while seeds and len(expanded) < payload.max_local_nodes:
        u = min(seeds)
        seeds.discard(u)
        if u in expanded:
            continue
        expanded.add(u)
        for k in range(off[u], off[u + 1]):
            v = tg[k]
            edges_from.append(u)
            edges_to.append(v)
            edges_guarded.append(gd[k])
            if v not in expanded and len(expanded) + len(seeds) < payload.max_local_nodes:
                seeds.add(v)
            if len(edges_from) >= payload.max_local_edges:
                return LocalForceSubgraph(
                    node_ids=tuple(sorted(expanded)),
                    edges_from=tuple(edges_from),
                    edges_to=tuple(edges_to),
                    edges_guarded=tuple(edges_guarded),
                    is_module_scope=False,
                    truncated=True,
                )
    return LocalForceSubgraph(
        node_ids=tuple(sorted(expanded)),
        edges_from=tuple(edges_from),
        edges_to=tuple(edges_to),
        edges_guarded=tuple(edges_guarded),
        is_module_scope=False,
        truncated=not payload.local_edges.complete[node_id],
    )


# --- Serialization ------------------------------------------------------------


def _payload_to_jsonable(payload: LightweightVizPayload) -> dict[str, Any]:
    return {
        "version": payload.version,
        "stats": {
            "node_count": payload.stats.node_count,
            "scc_count": payload.stats.scc_count,
            "module_count": payload.stats.module_count,
            "module_edge_count": payload.stats.module_edge_count,
            "local_edge_count": payload.stats.local_edge_count,
            "truncated_local_nodes": payload.stats.truncated_local_nodes,
            "dense_bucket_count": payload.stats.dense_bucket_count,
        },
        "max_local_nodes": payload.max_local_nodes,
        "max_local_edges": payload.max_local_edges,
        "sheets": list(payload.sheets),
        "nodes": {
            "sheet_index": list(payload.nodes.sheet_index),
            "row": list(payload.nodes.row),
            "column": list(payload.nodes.column),
            "is_leaf": list(payload.nodes.is_leaf),
            "in_degree": list(payload.nodes.in_degree),
            "out_degree": list(payload.nodes.out_degree),
            "module_id": list(payload.nodes.module_id),
            "rank": list(payload.nodes.rank),
            "x": list(payload.nodes.x),
            "y": list(payload.nodes.y),
            "bucket_density": list(payload.nodes.bucket_density),
        },
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
            for m in payload.modules
        ],
        "module_edges": [
            {
                "source_module_id": e.source_module_id,
                "target_module_id": e.target_module_id,
                "unconditional_weight": e.unconditional_weight,
                "guarded_weight": e.guarded_weight,
            }
            for e in payload.module_edges
        ],
        "local_edges": {
            "offsets": list(payload.local_edges.offsets),
            "targets": list(payload.local_edges.targets),
            "guarded": list(payload.local_edges.guarded),
            "complete": list(payload.local_edges.complete),
        },
    }


def estimate_serialized_json_bytes(payload: LightweightVizPayload) -> int:
    """Fast structural upper bound without a full ``json.dumps`` of huge payloads."""
    n = payload.stats.node_count
    e_loc = payload.stats.local_edge_count
    # Rough bracketing: integers ~6 chars avg, floats ~12, bools 5, brackets overhead ~15%.
    est = 2000
    est += n * (6 * 11 + 12 * 2 + 5 + 8 * 4)
    est += e_loc * 12
    est += len(payload.module_edges) * 40
    est += len(payload.modules) * 60
    est += sum(len(s) for s in payload.sheets) + n * 4
    return int(est * 1.15)


def serialize_lightweight_viz_json(payload: LightweightVizPayload) -> str:
    return json.dumps(_payload_to_jsonable(payload), separators=(",", ":"))


def write_lightweight_viz_data(payload: LightweightVizPayload, path: Path | str) -> None:
    p = Path(path)
    data = serialize_lightweight_viz_json(payload)
    p.write_text(data, encoding="utf-8")


def write_lightweight_viz_html(
    payload: LightweightVizPayload,
    path: Path | str,
    *,
    title: str = "Workbook dependency graph",
    data_mode: Literal["inline", "sidecar", "auto"] = "auto",
    data_path: Path | str | None = None,
    inline_size_budget_mb: int = 50,
) -> None:
    from importlib import resources

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

    tpl = (
        resources.files(__package__ or __name__)
        .joinpath("lightweight_viz_template.html")
        .read_text(encoding="utf-8")
    )
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
    )
    out.write_text(html, encoding="utf-8")
