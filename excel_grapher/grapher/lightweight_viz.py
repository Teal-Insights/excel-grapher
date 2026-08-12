"""Lightweight workbook graph visualization: core payload, overlays, JSON, and HTML."""

from __future__ import annotations

import heapq
import json
import math
from collections import deque
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

from excel_grapher.core.address_keys import (
    CellKey,
    RangeKey,
    UnionKey,
    normalize_key,
    parse_node_key,
)

from .formula_label import (
    display_formula,
    truncate_formula_display,
    validate_max_formula_length,
)
from .graph import DependencyGraph
from .node import NodeKey, NodeView

# --- Constants ----------------------------------------------------------------

DENSE_BUCKET_THRESHOLD = 12
VIZ_PAYLOAD_VERSION = 2
WEBVIZ_LOUVAIN_DIRECTED_OVERLAY_ID = "webviz.louvain_directed"

# BFS overview: weighted barycentric ordering within each (rank, module) bucket
BFS_HORIZONTAL_UNGUARDED_WEIGHT = 1.0
BFS_HORIZONTAL_GUARDED_WEIGHT = 0.35
BFS_HORIZONTAL_SWEEP_COUNT = 6
BFS_HORIZONTAL_MIN_SLOT_GAP = 1.0

# --- CSR / edge extraction ----------------------------------------------------


def _resolve_viz_endpoint(graph: DependencyGraph, dep: NodeKey) -> NodeKey | None:
    """Map an edge endpoint to a stored graph key (exact or occupancy owner)."""
    nk = normalize_key(dep)
    if nk in graph:
        return nk
    try:
        return graph.cell_owner(nk)
    except ValueError:
        return None


def _viz_label_geometry(node: NodeView) -> tuple[str, int, str]:
    """Return `(sheet, row, column)` used as the viz node label anchor.

    Multi-cell nodes use top-left / first-canonical-member geometry so unions
    and non-row ranges never require scalar `Node.column` / `Node.row`.
    """
    parsed = parse_node_key(node.key)
    if isinstance(parsed, CellKey):
        return parsed.sheet, parsed.row, parsed.column
    if isinstance(parsed, RangeKey):
        return parsed.sheet, parsed.min_row, parsed.min_col
    assert isinstance(parsed, UnionKey)
    first = parsed.members[0]
    if isinstance(first, CellKey):
        return first.sheet, first.row, first.column
    return first.sheet, first.min_row, first.min_col


def _node_sheets(node: NodeView) -> set[str]:
    parsed = parse_node_key(node.key)
    if isinstance(parsed, (CellKey, RangeKey)):
        return {parsed.sheet}
    return {m.sheet for m in parsed.members}


def _build_int_adjacencies(
    graph: DependencyGraph, keys: list[NodeKey], key_id: dict[NodeKey, int]
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
    graph: DependencyGraph,
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
    dense_bucket_count: int


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
    x: tuple[float, ...]
    y: tuple[float, ...]
    bucket_density: tuple[int, ...]


@dataclass(frozen=True, slots=True)
class LightweightVizCore:
    stats: LightweightVizCoreStats
    sheets: tuple[str, ...]
    nodes: LightweightVizCoreNodeColumns
    local_edges: LightweightVizLocalEdges
    max_local_nodes: int | None
    max_local_edges: int | None


def _build_out_adj_guarded(
    graph: DependencyGraph,
    keys: list[NodeKey],
    key_id: dict[NodeKey, int],
    *,
    include_guarded: bool,
) -> list[list[tuple[int, bool]]]:
    """Outgoing adjacency with guarded flags, aligned with `selected_adj` edge filtering."""
    n = len(keys)
    out: list[list[tuple[int, bool]]] = [[] for _ in range(n)]
    for fk in keys:
        fi = key_id[fk]
        for tk in graph.keys(order="workbook", source=graph.get_dependencies(fk)):
            resolved = _resolve_viz_endpoint(graph, tk)
            ti = None if resolved is None else key_id.get(resolved)
            if ti is None:
                continue
            guarded = graph.is_guarded(fk, tk)
            if guarded and not include_guarded:
                continue
            out[fi].append((ti, guarded))
    for row in out:
        row.sort(key=lambda t: t[0])
    return out


def _reverse_adj_flagged(
    adj_flagged: list[list[tuple[int, bool]]], n: int
) -> list[list[tuple[int, bool]]]:
    rev: list[list[tuple[int, bool]]] = [[] for _ in range(n)]
    for u in range(n):
        for v, g in adj_flagged[u]:
            rev[v].append((u, g))
    for row in rev:
        row.sort(key=lambda t: t[0])
    return rev


def _edge_weight(guarded: bool) -> float:
    return BFS_HORIZONTAL_GUARDED_WEIGHT if guarded else BFS_HORIZONTAL_UNGUARDED_WEIGHT


def _bucket_keys_sorted(ranks: list[int], module_of: list[int]) -> list[tuple[int, int]]:
    buckets = {(ranks[i], module_of[i]) for i in range(len(ranks))}
    return sorted(buckets)


def _bfs_bucket_sort_key(
    vid: int,
    ranks: list[int],
    adj_flagged: list[list[tuple[int, bool]]],
    rev_flagged: list[list[tuple[int, bool]]],
) -> tuple[int, int]:
    """Stable tie-break for initial bucket order using the closest rank neighbor when present."""
    r = ranks[vid]
    preds = [u for u, _ in rev_flagged[vid] if ranks[u] == r - 1]
    if preds:
        return (min(preds), vid)
    succs = [w for w, _ in adj_flagged[vid] if ranks[w] == r + 1]
    if succs:
        return (min(succs), vid)
    return (vid, vid)


def _bfs_horizontal_iteration_order(
    n: int,
    ranks: list[int],
    module_of: list[int],
    adj_flagged: list[list[tuple[int, bool]]],
    rev_flagged: list[list[tuple[int, bool]]],
) -> list[int]:
    """Deterministic permutation for `_rank_band_xy`: within each (rank, module) bucket, order by weighted barycentric sweeps."""
    if n == 0:
        return []
    bucket_order: dict[tuple[int, int], list[int]] = {}
    for bk in _bucket_keys_sorted(ranks, module_of):
        members = [i for i in range(n) if ranks[i] == bk[0] and module_of[i] == bk[1]]
        members.sort(key=lambda vid: _bfs_bucket_sort_key(vid, ranks, adj_flagged, rev_flagged))
        bucket_order[bk] = members

    slot_x = [0.0] * n
    for members in bucket_order.values():
        for j, nid in enumerate(members):
            slot_x[nid] = float(j) * BFS_HORIZONTAL_MIN_SLOT_GAP

    rank_min = min(ranks)
    rank_max = max(ranks)

    def respace_bucket(bk: tuple[int, int]) -> None:
        for j, nid in enumerate(bucket_order[bk]):
            slot_x[nid] = float(j) * BFS_HORIZONTAL_MIN_SLOT_GAP

    def down_sweep() -> None:
        for r in range(rank_min, rank_max + 1):
            for mid in sorted({module_of[i] for i in range(n) if ranks[i] == r}):
                bk = (r, mid)
                members = bucket_order[bk]
                scores: list[tuple[float, int]] = []
                for v in members:
                    num = 0.0
                    den = 0.0
                    for u, guarded in rev_flagged[v]:
                        if ranks[u] != r - 1:
                            continue
                        if module_of[u] != module_of[v]:
                            continue
                        w = _edge_weight(guarded)
                        num += w * slot_x[u]
                        den += w
                    sc = (num / den) if den > 0.0 else slot_x[v]
                    scores.append((sc, v))
                scores.sort(key=lambda t: (t[0], t[1]))
                bucket_order[bk] = [v for _, v in scores]
                respace_bucket(bk)

    def up_sweep() -> None:
        for r in range(rank_max, rank_min - 1, -1):
            for mid in sorted({module_of[i] for i in range(n) if ranks[i] == r}):
                bk = (r, mid)
                members = bucket_order[bk]
                scores: list[tuple[float, int]] = []
                for v in members:
                    num = 0.0
                    den = 0.0
                    for wn, guarded in adj_flagged[v]:
                        if ranks[wn] != r + 1:
                            continue
                        if module_of[wn] != module_of[v]:
                            continue
                        w = _edge_weight(guarded)
                        num += w * slot_x[wn]
                        den += w
                    sc = (num / den) if den > 0.0 else slot_x[v]
                    scores.append((sc, v))
                scores.sort(key=lambda t: (t[0], t[1]))
                bucket_order[bk] = [v for _, v in scores]
                respace_bucket(bk)

    for _ in range(BFS_HORIZONTAL_SWEEP_COUNT):
        down_sweep()
        up_sweep()

    return sorted(range(n), key=lambda i: (ranks[i], module_of[i], slot_x[i], i))


def _balance_overview_layout_spans(xs: list[float], ys: list[float]) -> None:
    if not xs:
        return
    min_x = min(xs)
    max_x = max(xs)
    min_y = min(ys)
    max_y = max(ys)
    cx = 0.5 * (min_x + max_x)
    cy = 0.5 * (min_y + max_y)
    span_x = max(max_x - min_x, 1.0)
    span_y = max(max_y - min_y, 1.0)
    target = max(span_x, span_y)
    sx = target / span_x
    sy = target / span_y
    for i in range(len(xs)):
        xs[i] = cx + (xs[i] - cx) * sx
        ys[i] = cy + (ys[i] - cy) * sy


def _rank_band_xy(
    n: int,
    module_of: list[int],
    node_rank: list[int],
    *,
    iteration_order: Sequence[int] | None = None,
) -> tuple[list[float], list[float], list[int], int]:
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
    bucket_running_idx: dict[tuple[int, int], int] = {}

    order = list(range(n)) if iteration_order is None else list(iteration_order)
    if len(order) != n:
        raise ValueError("iteration_order length must equal node count")
    if set(order) != set(range(n)):
        raise ValueError("iteration_order must be a permutation of range(n)")

    for i in order:
        rnk = node_rank[i]
        mid = module_of[i]
        xs[i] = float(rnk) * x_scale
        base_y = float(mid) * y_band
        bkey = (rnk, mid)
        cnt = bucket_counts[bkey]
        bucket_density[i] = cnt
        idx_in_bucket = bucket_running_idx.get(bkey, 0)
        bucket_running_idx[bkey] = idx_in_bucket + 1
        if cnt <= DENSE_BUCKET_THRESHOLD:
            centered = float(idx_in_bucket) - 0.5 * float(cnt - 1)
            ys[i] = base_y + centered * 4.0
        else:
            t = (i * 1103515245 + 12345) & 0x7FFFFFFF
            jx = ((t % 10000) / 10000.0 - 0.5) * y_band * 0.85
            jy = (((t // 10000) % 10000) / 10000.0 - 0.5) * 8.0
            ys[i] = base_y + jx + jy

    _balance_overview_layout_spans(xs, ys)
    xs, ys = ys, xs
    return xs, ys, bucket_density, dense_bucket_count


# Whole-graph force layout (export-time); pairwise repulsion only for modest n.
_FORCE_PAIRWISE_REPULSION_MAX_N = 512
_FORCE_LINK_DISTANCE = 40.0
_FORCE_LINK_STRENGTH = 0.06
_FORCE_CHARGE = 120.0
_FORCE_CENTER_STRENGTH = 0.05
_FORCE_TICKS_MIN = 40
_FORCE_TICKS_MAX = 120
_FORCE_GRID_REPULSE_CELL_DIVISOR = 48.0


def _force_directed_xy(n: int, adj: list[list[int]]) -> tuple[list[float], list[float]]:
    """Deterministic force-directed placement using link springs, pairwise repulsion, and weak centering."""
    if n == 0:
        return [], []
    radius = 100.0 * math.sqrt(float(max(n, 1)))
    xs = [0.0] * n
    ys = [0.0] * n
    for i in range(n):
        ang = 2.0 * math.pi * float(i) / float(n)
        xs[i] = math.cos(ang) * radius
        ys[i] = math.sin(ang) * radius

    edges: list[tuple[int, int]] = []
    for u in range(n):
        for v in adj[u]:
            edges.append((u, v))

    ticks = min(_FORCE_TICKS_MAX, max(_FORCE_TICKS_MIN, n // 50))
    if n > 50_000:
        ticks = min(ticks, 80)

    for tick in range(ticks):
        tnorm = tick / max(ticks - 1, 1) if ticks > 1 else 1.0
        alpha = (1.0 - tnorm) ** 0.5

        fx = [0.0] * n
        fy = [0.0] * n

        for u, v in edges:
            dx = xs[v] - xs[u]
            dy = ys[v] - ys[u]
            dist = math.hypot(dx, dy)
            if dist < 1e-9:
                dist = 1e-9
            fmag = _FORCE_LINK_STRENGTH * (dist - _FORCE_LINK_DISTANCE)
            fx_u = fmag * dx / dist
            fy_u = fmag * dy / dist
            fx[u] += fx_u
            fy[u] += fy_u
            fx[v] -= fx_u
            fy[v] -= fy_u

        if n <= _FORCE_PAIRWISE_REPULSION_MAX_N:
            for i in range(n):
                xi, yi = xs[i], ys[i]
                for j in range(i + 1, n):
                    dx = xs[j] - xi
                    dy = ys[j] - yi
                    dist2 = dx * dx + dy * dy + 0.01
                    dist = math.sqrt(dist2)
                    inv_cubed = _FORCE_CHARGE / (dist2 * dist)
                    fx[i] -= inv_cubed * dx
                    fy[i] -= inv_cubed * dy
                    fx[j] += inv_cubed * dx
                    fy[j] += inv_cubed * dy
        else:
            min_x = min(xs)
            max_x = max(xs)
            min_y = min(ys)
            max_y = max(ys)
            span = max(max_x - min_x, max_y - min_y, 1.0)
            cell = span / _FORCE_GRID_REPULSE_CELL_DIVISOR
            if cell < 1e-9:
                cell = 1e-9
            buckets: dict[tuple[int, int], list[int]] = {}
            for i in range(n):
                bx = int(xs[i] / cell)
                by = int(ys[i] / cell)
                buckets.setdefault((bx, by), []).append(i)
            for i in range(n):
                bx = int(xs[i] / cell)
                by = int(ys[i] / cell)
                for ox in (-1, 0, 1):
                    for oy in (-1, 0, 1):
                        for j in buckets.get((bx + ox, by + oy), []):
                            if j <= i:
                                continue
                            dx = xs[j] - xs[i]
                            dy = ys[j] - ys[i]
                            dist2 = dx * dx + dy * dy + 0.01
                            dist = math.sqrt(dist2)
                            inv_cubed = _FORCE_CHARGE / (dist2 * dist)
                            fx[i] -= inv_cubed * dx
                            fy[i] -= inv_cubed * dy
                            fx[j] += inv_cubed * dx
                            fy[j] += inv_cubed * dy

        for i in range(n):
            fx[i] -= _FORCE_CENTER_STRENGTH * xs[i]
            fy[i] -= _FORCE_CENTER_STRENGTH * ys[i]

        for i in range(n):
            xs[i] += fx[i] * alpha
            ys[i] += fy[i] * alpha

    return xs, ys


def _grid_xy(n: int) -> tuple[list[float], list[float], list[int], int]:
    x_scale = 120.0
    y_band = 36.0
    cols = max(1, int(math.ceil(math.sqrt(max(n, 1)))))
    xs = [0.0] * n
    ys = [0.0] * n
    bucket_density = [1] * n
    for i in range(n):
        r = i // cols
        c = i % cols
        xs[i] = float(c) * x_scale
        ys[i] = float(r) * y_band
    _balance_overview_layout_spans(xs, ys)
    xs, ys = ys, xs
    return xs, ys, bucket_density, 0


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


def _default_bfs_target_ranks(adj: list[list[int]], rev_adj: list[list[int]], n: int) -> list[int]:
    if n == 0:
        return []
    target_like = [i for i in range(n) if not rev_adj[i]]
    if not target_like:
        target_like = list(range(n))
    target_like.sort()

    dist = [-1] * n
    q: deque[int] = deque()
    for s in target_like:
        dist[s] = 0
        q.append(s)
    while q:
        u = q.popleft()
        du = dist[u]
        for v in adj[u]:
            if dist[v] >= 0:
                continue
            dist[v] = du + 1
            q.append(v)
    return [d if d >= 0 else 0 for d in dist]


def _bfs_distances_from_seed_ids(
    adj: list[list[int]], n: int, seed_ids: Sequence[int]
) -> list[int]:
    dist = [-1] * n
    q: deque[int] = deque()
    for s in seed_ids:
        if not (0 <= s < n):
            continue
        if dist[s] >= 0:
            continue
        dist[s] = 0
        q.append(s)
    while q:
        u = q.popleft()
        du = dist[u]
        for v in adj[u]:
            if dist[v] >= 0:
                continue
            dist[v] = du + 1
            q.append(v)
    return dist


def _induced_dependency_subgraph(
    graph: DependencyGraph,
    keep_keys: set[NodeKey],
) -> DependencyGraph:
    sub = DependencyGraph()
    if graph.sheet_order is not None:
        sub.sheet_order = list(graph.sheet_order)
    sub.leaf_classification = graph.leaf_classification
    for k in graph.keys(order="workbook", source=keep_keys):
        node = graph._get_internal_node(k)
        if node is None:
            continue
        sub.add_node(node)
    for fk in graph.keys(order="workbook", source=keep_keys):
        for tk in graph.keys(order="workbook", source=graph.get_dependencies(fk)):
            resolved = _resolve_viz_endpoint(graph, tk)
            if resolved is None or resolved not in keep_keys:
                continue
            edge = graph.get_edge_attrs(fk, tk)
            edge_kwargs: dict[str, Any] = {}
            if edge.provenance is not None:
                edge_kwargs["provenance"] = edge.provenance
            sub.add_edge(fk, resolved, guard=edge.guard, **edge_kwargs)
    return sub


def build_lightweight_viz_core(
    graph: DependencyGraph,
    *,
    limits: VizLimits | None = None,
    layout_input: LightweightVizLayoutInput | None = None,
    layout_mode: Literal["bfs", "layered", "grid", "force"] = "bfs",
    include_guarded_edges: bool = True,
    bfs_seed_keys: Sequence[NodeKey] | None = None,
    exclude_unreachable_from_bfs: bool = False,
    include_formula_on_nodes: bool = True,
    max_formula_length: int | None = 120,
) -> LightweightVizCore:
    validate_max_formula_length(max_formula_length)

    lim = limits or VizLimits()
    keys = graph.keys(order="workbook")
    n = len(keys)
    if n == 0:
        return LightweightVizCore(
            stats=LightweightVizCoreStats(
                node_count=0,
                local_edge_count=0,
                truncated_local_nodes=0,
                dense_bucket_count=0,
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
                x=tuple(),
                y=tuple(),
                bucket_density=tuple(),
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
        if layout_mode == "grid":
            xs, ys, bucket_density, dense_bucket_count = _grid_xy(n)
            cols = max(1, int(math.ceil(math.sqrt(n))))
            ranks = [i // cols for i in range(n)]
        elif layout_mode == "bfs":
            if bfs_seed_keys is None:
                seed_ids = [i for i in range(n) if not rev_selected[i]]
                if not seed_ids:
                    seed_ids = list(range(n))
            else:
                seed_ids = [key_id[k] for k in bfs_seed_keys if k in key_id]
                if not seed_ids:
                    seed_ids = [i for i in range(n) if not rev_selected[i]]
                    if not seed_ids:
                        seed_ids = list(range(n))
            dist = _bfs_distances_from_seed_ids(selected_adj, n, seed_ids)
            should_exclude_unreachable = exclude_unreachable_from_bfs or not include_guarded_edges
            if should_exclude_unreachable:
                keep_ids = [i for i, d in enumerate(dist) if d >= 0]
                if keep_ids and len(keep_ids) < n:
                    keep_keys = {keys[i] for i in keep_ids}
                    subgraph = _induced_dependency_subgraph(graph, keep_keys)
                    sub_seeds = (
                        None
                        if bfs_seed_keys is None
                        else tuple(k for k in bfs_seed_keys if k in keep_keys)
                    )
                    return build_lightweight_viz_core(
                        subgraph,
                        limits=lim,
                        layout_input=None,
                        layout_mode=layout_mode,
                        include_guarded_edges=include_guarded_edges,
                        bfs_seed_keys=sub_seeds,
                        exclude_unreachable_from_bfs=should_exclude_unreachable,
                        include_formula_on_nodes=include_formula_on_nodes,
                        max_formula_length=max_formula_length,
                    )
            ranks = [d if d >= 0 else 0 for d in dist]
            adj_flagged = _build_out_adj_guarded(
                graph, keys, key_id, include_guarded=include_guarded_edges
            )
            rev_flagged = _reverse_adj_flagged(adj_flagged, n)
            iteration_order = _bfs_horizontal_iteration_order(
                n, ranks, module_of, adj_flagged, rev_flagged
            )
            xs, ys, bucket_density, dense_bucket_count = _rank_band_xy(
                n, module_of, ranks, iteration_order=iteration_order
            )
        elif layout_mode == "layered":
            ranks, _scc_count = unconditional_scc_ranks(uncond, n)
            xs, ys, bucket_density, dense_bucket_count = _rank_band_xy(n, module_of, ranks)
        elif layout_mode == "force":
            ranks = _default_bfs_target_ranks(selected_adj, rev_selected, n)
            _, _, bucket_density, dense_bucket_count = _rank_band_xy(n, module_of, ranks)
            xs, ys = _force_directed_xy(n, selected_adj)
            _balance_overview_layout_spans(xs, ys)
            xs, ys = ys, xs
        else:
            raise ValueError(f"Unsupported layout_mode: {layout_mode!r}")
    else:
        if len(layout_input.module_of) != n or len(layout_input.node_rank) != n:
            raise ValueError("layout_input tuple lengths must match graph order")
        module_of = list(layout_input.module_of)
        node_rank = list(layout_input.node_rank)
        ranks = node_rank
        if layout_mode == "force":
            _, _, bucket_density, dense_bucket_count = _rank_band_xy(n, module_of, node_rank)
            xs, ys = _force_directed_xy(n, selected_adj)
            _balance_overview_layout_spans(xs, ys)
            xs, ys = ys, xs
        else:
            xs, ys, bucket_density, dense_bucket_count = _rank_band_xy(n, module_of, node_rank)

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
        dense_bucket_count=dense_bucket_count,
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


# --- Flat view (core + partition overlay) for tools, tests, force layout ------


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
    formula: tuple[str | None, ...]
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
class LightweightVizFlat:
    stats: LightweightVizStats
    sheets: tuple[str, ...]
    nodes: LightweightVizNodeColumns
    modules: tuple[LightweightVizModule, ...]
    module_edges: tuple[LightweightVizModuleEdge, ...]
    local_edges: LightweightVizLocalEdges
    max_local_nodes: int | None
    max_local_edges: int | None


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

    bucket_counts: dict[tuple[int, int], int] = {}
    for i in range(n):
        b = (node_rank[i], module_id[i])
        bucket_counts[b] = bucket_counts.get(b, 0) + 1

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
    return tuple(modules)


def lightweight_viz_flat(payload: LightweightVizPayload) -> LightweightVizFlat:
    if payload.version != VIZ_PAYLOAD_VERSION:
        raise ValueError(f"Unsupported lightweight viz payload version: {payload.version}")
    core = payload.core
    n = core.stats.node_count

    mod_ov: LightweightVizOverlay | None = None
    for ov in payload.overlays:
        if ov.overlay_id == WEBVIZ_LOUVAIN_DIRECTED_OVERLAY_ID:
            mod_ov = ov
            break

    if mod_ov is not None:
        data = mod_ov.data
        module_id = tuple(int(x) for x in data["node_module_id"])
        node_rank_out = tuple(int(x) for x in data["node_rank"])
        raw_edges = data["module_edges"]
        module_edges = tuple(
            LightweightVizModuleEdge(
                source_module_id=int(e["source_module_id"]),
                target_module_id=int(e["target_module_id"]),
                unconditional_weight=int(e["unconditional_weight"]),
                guarded_weight=int(e["guarded_weight"]),
            )
            for e in raw_edges
        )
        raw_mods = data["modules"]
        modules = tuple(
            LightweightVizModule(
                id=int(m["id"]),
                node_count=int(m["node_count"]),
                rank_min=int(m["rank_min"]),
                rank_max=int(m["rank_max"]),
                centroid_x=float(m["centroid_x"]),
                centroid_y=float(m["centroid_y"]),
                density_mode=bool(m["density_mode"]),
            )
            for m in raw_mods
        )
        stats = LightweightVizStats(
            node_count=n,
            scc_count=int(data["scc_count"]),
            module_count=len(modules),
            module_edge_count=len(module_edges),
            local_edge_count=core.stats.local_edge_count,
            truncated_local_nodes=core.stats.truncated_local_nodes,
            dense_bucket_count=core.stats.dense_bucket_count,
        )
    else:
        module_id = (0,) * n if n else tuple()
        node_rank_out = tuple(core.nodes.rank)
        modules = derive_partition_modules_table(core, module_id, node_rank_out)
        module_edges = tuple()
        stats = LightweightVizStats(
            node_count=n,
            scc_count=0,
            module_count=len(modules) if n else 0,
            module_edge_count=0,
            local_edge_count=core.stats.local_edge_count,
            truncated_local_nodes=core.stats.truncated_local_nodes,
            dense_bucket_count=core.stats.dense_bucket_count,
        )

    nodes = LightweightVizNodeColumns(
        sheet_index=core.nodes.sheet_index,
        row=core.nodes.row,
        column=core.nodes.column,
        is_leaf=core.nodes.is_leaf,
        formula=core.nodes.formula,
        in_degree=core.nodes.in_degree,
        out_degree=core.nodes.out_degree,
        module_id=module_id,
        rank=node_rank_out,
        x=core.nodes.x,
        y=core.nodes.y,
        bucket_density=core.nodes.bucket_density,
    )

    return LightweightVizFlat(
        stats=stats,
        sheets=core.sheets,
        nodes=nodes,
        modules=modules,
        module_edges=module_edges,
        local_edges=core.local_edges,
        max_local_nodes=core.max_local_nodes,
        max_local_edges=core.max_local_edges,
    )


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
            "dense_bucket_count": c.stats.dense_bucket_count,
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
            "x": list(nc.x),
            "y": list(nc.y),
            "bucket_density": list(nc.bucket_density),
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
    )
    out.write_text(html, encoding="utf-8")
