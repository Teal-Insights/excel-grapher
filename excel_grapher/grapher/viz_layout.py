"""Shared multilevel clustered force layout for every graph viewer.

The cell viewer (`to_web_viz_payload`) and the statement viewer
(`to_semantic_viz_payload`) both precompute node positions here, and both HTML
templates receive `viz_layout_js_config()` so in-browser relayout
(`viz_force.js`) runs with the same settings.

The layout works in levels, the approach sfdp and ForceAtlas2 use for large
graphs:

1. Force layout on the graph of clusters, one node per cluster.
2. Force layout of every node, anchored to its cluster's position.
3. For the large tier only, a few global refinement ticks.

Repulsion is exact (all pairs, or all pairs inside each cluster) while the
pair count fits `FORCE_EXACT_PAIR_BUDGET`, and grid-based (pairs in
neighboring cells only) beyond that.

An optional weak vertical pull towards input depth keeps the input -> output
direction readable: inputs sit at the top (smallest `y`).
"""

from __future__ import annotations

import json
import math
from collections.abc import Hashable, Sequence
from typing import TYPE_CHECKING, Any, Literal, TypeAlias

if TYPE_CHECKING:
    import numpy as np

VizTier: TypeAlias = Literal["small", "medium", "large"]
RankPull: TypeAlias = Literal["none", "between", "everywhere"]

RANK_PULL_MODES: tuple[RankPull, ...] = ("none", "between", "everywhere")
DEFAULT_RANK_PULL: RankPull = "between"

# --- Size tiers -------------------------------------------------------------
# small: exact force, labeled boxes, whole-graph relayout in the browser.
# medium: clustered force, boxes with zoom-dependent labels, relayout a cluster
#   or neighborhood in the browser.
# large: all levels, opens on the cluster overview, relayout a cluster only.
VIZ_TIER_SMALL_MAX = 2_000
VIZ_TIER_MEDIUM_MAX = 20_000

# Widest opening camera, in layout units, for the statement viewer.
VIZ_CAMERA_MAX_WIDTH = 2800
# Overview edge draw caps for the WebGL cell viewer, by tier.
VIZ_MEDIUM_MAX_DRAWN_EDGES = 500_000
VIZ_LARGE_MAX_DRAWN_EDGES = 200_000

# --- Force settings (mirrored in viz_force.js via viz_layout_js_config) -----
FORCE_LINK_DISTANCE = 40.0
FORCE_LINK_STRENGTH = 0.06
# Repulsion is `charge / dist**2`; this balances a spring at its rest length.
FORCE_CHARGE = 0.5 * FORCE_LINK_STRENGTH * FORCE_LINK_DISTANCE**3
FORCE_CLUSTER_ATTRACT = 0.02
FORCE_CENTER_STRENGTH = 0.01
FORCE_INTER_CLUSTER_LINK_SCALE = 0.1
FORCE_RANK_PULL = 0.05
FORCE_RANK_PULL_WITHIN = 0.02
FORCE_TICKS = 80
FORCE_REFINE_TICKS = 20
FORCE_MAX_STEP = FORCE_LINK_DISTANCE
FORCE_REPULSION_CUTOFF = 2.0 * FORCE_LINK_DISTANCE
# Grid repulsion keeps a neighbor list of pairs within cutoff + skin and
# rebuilds it every `FORCE_NEIGHBOR_REBUILD_TICKS` ticks.
FORCE_NEIGHBOR_SKIN = 0.5 * FORCE_LINK_DISTANCE
FORCE_NEIGHBOR_REBUILD_TICKS = 5
FORCE_LARGE_TICKS = 50
# Spiral spacing of the initial disc; a cluster of `s` nodes spans radius
# about `FORCE_DISC_SPACING * sqrt(s)`.
FORCE_DISC_SPACING = 0.5 * FORCE_LINK_DISTANCE
# Most repulsion pairs computed exactly per tick; beyond it, grid pairs.
FORCE_EXACT_PAIR_BUDGET = 250_000

_GOLDEN_ANGLE = math.pi * (3.0 - math.sqrt(5.0))

__all__ = [
    "DEFAULT_RANK_PULL",
    "FORCE_CHARGE",
    "FORCE_LINK_DISTANCE",
    "FORCE_LINK_STRENGTH",
    "RANK_PULL_MODES",
    "VIZ_CAMERA_MAX_WIDTH",
    "VIZ_TIER_MEDIUM_MAX",
    "VIZ_TIER_SMALL_MAX",
    "RankPull",
    "VizTier",
    "clustered_force_layout",
    "input_depths",
    "viz_force_js_source",
    "viz_layout_js_bootstrap",
    "viz_layout_js_config",
    "viz_tier",
]


def viz_tier(size: int) -> VizTier:
    """Return the size tier for a graph with `size` drawn primitives.

    Args:
        size: Node count for the cell viewer, or statements plus twice the
            bundles for the statement viewer.
    """
    if size <= VIZ_TIER_SMALL_MAX:
        return "small"
    if size <= VIZ_TIER_MEDIUM_MAX:
        return "medium"
    return "large"


def viz_layout_js_config() -> dict[str, Any]:
    """Return the shared constants injected into both HTML viewers."""
    return {
        "tier_small_max": VIZ_TIER_SMALL_MAX,
        "tier_medium_max": VIZ_TIER_MEDIUM_MAX,
        "camera_max_width": VIZ_CAMERA_MAX_WIDTH,
        "medium_max_drawn_edges": VIZ_MEDIUM_MAX_DRAWN_EDGES,
        "large_max_drawn_edges": VIZ_LARGE_MAX_DRAWN_EDGES,
        "link_distance": FORCE_LINK_DISTANCE,
        "link_strength": FORCE_LINK_STRENGTH,
        "charge": FORCE_CHARGE,
        "cluster_attract": FORCE_CLUSTER_ATTRACT,
        "center_strength": FORCE_CENTER_STRENGTH,
        "inter_cluster_link_scale": FORCE_INTER_CLUSTER_LINK_SCALE,
        "rank_pull": FORCE_RANK_PULL,
        "rank_pull_within": FORCE_RANK_PULL_WITHIN,
        "ticks": FORCE_TICKS,
        "max_step": FORCE_MAX_STEP,
        "disc_spacing": FORCE_DISC_SPACING,
    }


def viz_force_js_source() -> str:
    """Return the shared in-browser relayout module (`viz_force.js`)."""
    from importlib import resources

    return (
        resources.files(__package__ or "excel_grapher.grapher")
        .joinpath("viz_force.js")
        .read_text(encoding="utf-8")
    )


def viz_layout_js_bootstrap() -> str:
    """Return JS that defines `VIZ_LAYOUT_CONFIG` and `VizForce` for a viewer."""
    config = json.dumps(viz_layout_js_config(), separators=(",", ":"))
    return f"window.VIZ_LAYOUT_CONFIG = {config};\n{viz_force_js_source()}"


def input_depths(n: int, edges: Sequence[tuple[int, int]]) -> list[int]:
    """Return each node's longest-path depth from the inputs.

    Args:
        n: Node count.
        edges: `(consumer, producer)` pairs. Nodes that consume nothing have
            depth 0; cycle members share a depth.
    """
    from excel_grapher.grapher.lightweight_viz import unconditional_scc_ranks

    feeds: list[list[int]] = [[] for _ in range(n)]
    for consumer, producer in edges:
        if consumer != producer:
            feeds[producer].append(consumer)
    for row in feeds:
        row.sort()
    depths, _ = unconditional_scc_ranks(feeds, n)
    return depths


def _numpy():
    try:
        import numpy
    except ImportError as e:  # pragma: no cover
        raise ImportError("numpy is required for graph viewer layouts") from e
    return numpy


def _accumulate(np, force, idx, vec) -> None:
    n = force.shape[0]
    force[:, 0] += np.bincount(idx, weights=vec[:, 0], minlength=n)
    force[:, 1] += np.bincount(idx, weights=vec[:, 1], minlength=n)


def _grid_pairs(pos: np.ndarray, cutoff: float) -> tuple[np.ndarray, np.ndarray]:
    """Return index pairs `(i, j)` for nodes in the same or adjacent grid cells.

    Every pair closer than `cutoff` is included exactly once; farther pairs may
    be included too.
    """
    np = _numpy()
    n = pos.shape[0]
    empty = np.zeros(0, dtype=np.int64)
    if n < 2:
        return empty, empty
    cell = np.floor(pos / cutoff).astype(np.int64)
    cx = cell[:, 0] - cell[:, 0].min()
    cy = cell[:, 1] - cell[:, 1].min() + 1
    stride = int(cy.max()) + 2
    key = cx * stride + cy
    order = np.argsort(key, kind="stable")
    cells, start, counts = np.unique(key[order], return_index=True, return_counts=True)

    out_i: list[np.ndarray] = []
    out_j: list[np.ndarray] = []

    def cross(a: np.ndarray, b: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
        ca = counts[a]
        cb = counts[b]
        tot = ca * cb
        total = int(tot.sum())
        if total == 0:
            return empty, empty
        base = np.repeat(np.cumsum(tot) - tot, tot)
        local = np.arange(total, dtype=np.int64) - base
        cbr = np.repeat(cb, tot)
        i = order[np.repeat(start[a], tot) + local // cbr]
        j = order[np.repeat(start[b], tot) + local % cbr]
        return i, j

    same = np.nonzero(counts > 1)[0]
    if same.size:
        i, j = cross(same, same)
        keep = i < j
        out_i.append(i[keep])
        out_j.append(j[keep])
    for off in (1, stride - 1, stride, stride + 1):
        target = cells + off
        idx = np.searchsorted(cells, target)
        idx_clip = np.minimum(idx, cells.size - 1)
        valid = (idx < cells.size) & (cells[idx_clip] == target)
        a = np.nonzero(valid)[0]
        if a.size:
            i, j = cross(a, idx_clip[valid])
            out_i.append(i)
            out_j.append(j)
    if not out_i:
        return empty, empty
    return np.concatenate(out_i), np.concatenate(out_j)


def _simulate(
    pos: np.ndarray,
    *,
    edge_i: np.ndarray,
    edge_j: np.ndarray,
    link_len: np.ndarray | float,
    link_k: np.ndarray | float,
    pairs: tuple[np.ndarray, np.ndarray] | None,
    cutoff: float,
    charge_weight: np.ndarray | None = None,
    anchors: np.ndarray | None = None,
    anchor_k: float = 0.0,
    center_k: float = 0.0,
    target_y: np.ndarray | None = None,
    pull_k: float = 0.0,
    ticks: int = FORCE_TICKS,
    alpha0: float = 1.0,
) -> np.ndarray:
    """Run `ticks` force steps in place and return `pos`.

    `pairs=None` uses grid pairs within `cutoff`, from a neighbor list rebuilt
    every `FORCE_NEIGHBOR_REBUILD_TICKS` ticks.
    """
    np = _numpy()
    neighbors = pairs
    for tick in range(ticks):
        if pairs is None and tick % FORCE_NEIGHBOR_REBUILD_TICKS == 0:
            reach = cutoff + FORCE_NEIGHBOR_SKIN
            gi, gj = _grid_pairs(pos, reach)
            gd = pos[gj] - pos[gi]
            near = (gd * gd).sum(axis=1) < reach * reach
            neighbors = (gi[near], gj[near])
        alpha = alpha0 * (1.0 - tick / ticks) ** 0.5
        force = np.zeros_like(pos)
        if edge_i.size:
            d = pos[edge_j] - pos[edge_i]
            dist = np.maximum(np.sqrt((d * d).sum(axis=1)), 1e-9)
            vec = d * (link_k * (dist - link_len) / dist)[:, None]
            _accumulate(np, force, edge_i, vec)
            _accumulate(np, force, edge_j, -vec)
        assert neighbors is not None
        pi, pj = neighbors
        if pi.size:
            d = pos[pj] - pos[pi]
            dist2 = (d * d).sum(axis=1) + 0.01
            mag = FORCE_CHARGE / (dist2 * np.sqrt(dist2))
            if charge_weight is not None:
                mag = mag * charge_weight[pi] * charge_weight[pj]
            vec = d * mag[:, None]
            _accumulate(np, force, pi, -vec)
            _accumulate(np, force, pj, vec)
        if anchors is not None and anchor_k:
            force += anchor_k * (anchors - pos)
        if center_k:
            force -= center_k * (pos - pos.mean(axis=0))
        if target_y is not None and pull_k:
            force[:, 1] += pull_k * (target_y - pos[:, 1])
        step = force * alpha
        norm = np.sqrt((step * step).sum(axis=1))
        over = norm > FORCE_MAX_STEP
        if over.any():
            step[over] *= (FORCE_MAX_STEP / norm[over])[:, None]
        pos += step
    return pos


def _spiral(np, count: int, spacing: float) -> np.ndarray:
    k = np.arange(count, dtype=np.float64)
    r = spacing * np.sqrt(k + 0.5)
    theta = k * _GOLDEN_ANGLE
    return np.stack([r * np.cos(theta), r * np.sin(theta)], axis=1)


def _all_pairs(np, members: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
    a, b = np.triu_indices(members.size, k=1)
    return members[a], members[b]


def _separate_discs(np, centers: np.ndarray, radii: np.ndarray, gap: float) -> None:
    """Push overlapping cluster discs apart in place (exact pairs)."""
    k = centers.shape[0]
    if k < 2 or k * (k - 1) // 2 > FORCE_EXACT_PAIR_BUDGET:
        return
    a, b = np.triu_indices(k, k=1)
    need = radii[a] + radii[b] + gap
    for _ in range(40):
        d = centers[b] - centers[a]
        dist = np.sqrt((d * d).sum(axis=1))
        overlap = need - dist
        hit = overlap > 0
        if not hit.any():
            return
        dist_hit = np.maximum(dist[hit], 1e-9)
        push = d[hit] * (0.5 * overlap[hit] / dist_hit)[:, None]
        zero = dist[hit] < 1e-9
        if zero.any():
            push[zero] = np.array([0.5 * gap, 0.0])
        move = np.zeros_like(centers)
        _accumulate(np, move, a[hit], -push)
        _accumulate(np, move, b[hit], push)
        centers += move


def _cluster_ids(clusters: Sequence[Hashable]) -> tuple[list[int], int]:
    seen: dict[Hashable, int] = {}
    ids = [seen.setdefault(c, len(seen)) for c in clusters]
    return ids, len(seen)


def clustered_force_layout(
    n: int,
    edges: Sequence[tuple[int, int]],
    clusters: Sequence[Hashable],
    *,
    depths: Sequence[int] | None = None,
    rank_pull: RankPull = "none",
    tier: VizTier | None = None,
    seed: int = 0,
) -> np.ndarray:
    """Lay out a graph by clusters with a multilevel force simulation.

    Args:
        n: Node count.
        edges: `(u, v)` node-index pairs. Direction only matters for `depths`.
        clusters: One hashable cluster label per node (Louvain module, series,
            role, ...). A single label gives a plain force layout.
        depths: Input depth per node (see `input_depths`). Required unless
            `rank_pull` is `"none"`.
        rank_pull: Weak vertical pull towards depth: `"none"`, `"between"`
            clusters only, or `"everywhere"` (between clusters and, weaker,
            inside each cluster). Inputs are pulled to the top.
        tier: Size tier; defaults to `viz_tier(n)`. Repulsion is exact over
            all pairs while they fit `FORCE_EXACT_PAIR_BUDGET`, else exact
            inside clusters when those fit (not for `large`), else grid-based.
            `large` adds a global refinement pass.
        seed: Seed for the tiny symmetry-breaking jitter.

    Returns:
        `(n, 2)` float array of positions centered on the origin, in layout
        units (`FORCE_LINK_DISTANCE` is the spring rest length).

    Raises:
        ValueError: `clusters` or `depths` length mismatch, unknown
            `rank_pull`, or a pull without `depths`.
    """
    np = _numpy()
    if rank_pull not in RANK_PULL_MODES:
        raise ValueError(f"rank_pull must be one of {RANK_PULL_MODES}, got {rank_pull!r}")
    if len(clusters) != n:
        raise ValueError("clusters must have one label per node")
    if rank_pull != "none":
        if depths is None:
            raise ValueError("rank_pull needs depths")
        if len(depths) != n:
            raise ValueError("depths must have one entry per node")
    if n == 0:
        return np.zeros((0, 2), dtype=np.float64)
    tier = tier or viz_tier(n)

    cid_list, k = _cluster_ids(clusters)
    cid = np.asarray(cid_list, dtype=np.int64)
    sizes = np.bincount(cid, minlength=k).astype(np.float64)
    radii = FORCE_DISC_SPACING * np.sqrt(sizes)
    depth = None if rank_pull == "none" else np.asarray(depths, dtype=np.float64)

    edge_arr = np.asarray(
        [(u, v) for u, v in edges if u != v] or np.zeros((0, 2)), dtype=np.int64
    ).reshape(-1, 2)
    eu, ev = edge_arr[:, 0], edge_arr[:, 1]

    # Level 1: one node per cluster.
    centers = _spiral(np, k, 2.0 * float(radii.mean()) + FORCE_LINK_DISTANCE)
    if k > 1:
        cu, cv = cid[eu], cid[ev]
        cross = cu != cv
        pair = np.unique(np.sort(np.stack([cu[cross], cv[cross]], axis=1), axis=1), axis=0)
        ci = pair[:, 0] if pair.size else np.zeros(0, dtype=np.int64)
        cj = pair[:, 1] if pair.size else np.zeros(0, dtype=np.int64)
        target = None
        if depth is not None:
            mean_depth = np.bincount(cid, weights=depth, minlength=k) / sizes
            span = float(mean_depth.max() - mean_depth.min())
            extent = 3.0 * math.sqrt(float((radii * radii).sum()))
            if span > 0:
                target = (mean_depth - mean_depth.mean()) / span * extent
        cluster_pairs = (
            _all_pairs(np, np.arange(k)) if k * (k - 1) // 2 <= FORCE_EXACT_PAIR_BUDGET else None
        )
        cutoff = 2.0 * float(radii.max()) + FORCE_REPULSION_CUTOFF
        _simulate(
            centers,
            edge_i=ci,
            edge_j=cj,
            link_len=radii[ci] + radii[cj] + FORCE_LINK_DISTANCE,
            link_k=FORCE_LINK_STRENGTH,
            pairs=cluster_pairs,
            cutoff=cutoff,
            charge_weight=np.sqrt(sizes),
            center_k=FORCE_CENTER_STRENGTH,
            target_y=target,
            pull_k=FORCE_RANK_PULL,
        )
        _separate_discs(np, centers, radii, FORCE_LINK_DISTANCE)

    # Level 2: every node, anchored to its cluster.
    order = np.argsort(cid, kind="stable")
    rank_in_cluster = np.empty(n, dtype=np.int64)
    starts = np.concatenate([[0], np.cumsum(sizes.astype(np.int64))[:-1]])
    rank_in_cluster[order] = np.arange(n) - np.repeat(starts, sizes.astype(np.int64))
    pos = centers[cid] + _spiral(np, int(sizes.max()), FORCE_DISC_SPACING)[rank_in_cluster]
    rng = np.random.default_rng(seed)
    pos += rng.uniform(-1e-3, 1e-3, size=pos.shape) * FORCE_LINK_DISTANCE

    same = cid[eu] == cid[ev]
    link_k = np.where(
        same, FORCE_LINK_STRENGTH, FORCE_LINK_STRENGTH * FORCE_INTER_CLUSTER_LINK_SCALE
    )

    pairs: tuple[np.ndarray, np.ndarray] | None
    if n * (n - 1) // 2 <= FORCE_EXACT_PAIR_BUDGET:
        pairs = _all_pairs(np, np.arange(n))
    elif tier != "large" and float((sizes * (sizes - 1) / 2).sum()) <= FORCE_EXACT_PAIR_BUDGET:
        parts = [_all_pairs(np, order[s : s + int(c)]) for s, c in zip(starts, sizes, strict=True)]
        pairs = (np.concatenate([p[0] for p in parts]), np.concatenate([p[1] for p in parts]))
    else:
        pairs = None

    target_y = None
    if depth is not None and rank_pull == "everywhere":
        mean_depth = np.bincount(cid, weights=depth, minlength=k) / sizes
        lo = np.full(k, np.inf)
        hi = np.full(k, -np.inf)
        np.minimum.at(lo, cid, depth)
        np.maximum.at(hi, cid, depth)
        span = np.maximum(hi - lo, 1.0)
        target_y = centers[cid, 1] + (depth - mean_depth[cid]) / span[cid] * 2.0 * radii[cid]

    _simulate(
        pos,
        edge_i=eu,
        edge_j=ev,
        link_len=FORCE_LINK_DISTANCE,
        link_k=link_k,
        pairs=pairs,
        cutoff=FORCE_REPULSION_CUTOFF,
        anchors=centers[cid],
        anchor_k=FORCE_CLUSTER_ATTRACT,
        target_y=target_y,
        pull_k=FORCE_RANK_PULL_WITHIN,
        ticks=FORCE_LARGE_TICKS if tier == "large" else FORCE_TICKS,
    )

    # Level 3: global refinement with grid repulsion across cluster borders.
    if tier == "large":
        _simulate(
            pos,
            edge_i=eu,
            edge_j=ev,
            link_len=FORCE_LINK_DISTANCE,
            link_k=link_k,
            pairs=None,
            cutoff=FORCE_REPULSION_CUTOFF,
            anchors=centers[cid],
            anchor_k=0.25 * FORCE_CLUSTER_ATTRACT,
            target_y=target_y,
            pull_k=FORCE_RANK_PULL_WITHIN,
            ticks=FORCE_REFINE_TICKS,
            alpha0=0.5,
        )

    pos -= pos.mean(axis=0)
    return pos
