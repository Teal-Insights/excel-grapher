"""Shared multilevel clustered force layout used by every graph viewer."""

from __future__ import annotations

import pytest

np = pytest.importorskip("numpy")

from excel_grapher.grapher import viz_layout as vl  # noqa: E402


def _clique_edges(members: list[int]) -> list[tuple[int, int]]:
    return [(a, b) for i, a in enumerate(members) for b in members[i + 1 :]]


def _two_cliques(size: int = 12) -> tuple[int, list[tuple[int, int]], list[int]]:
    a = list(range(size))
    b = list(range(size, 2 * size))
    edges = _clique_edges(a) + _clique_edges(b) + [(0, size)]
    return 2 * size, edges, [0] * size + [1] * size


def test_viz_tier_boundaries() -> None:
    assert vl.viz_tier(0) == "small"
    assert vl.viz_tier(vl.VIZ_TIER_SMALL_MAX) == "small"
    assert vl.viz_tier(vl.VIZ_TIER_SMALL_MAX + 1) == "medium"
    assert vl.viz_tier(vl.VIZ_TIER_MEDIUM_MAX) == "medium"
    assert vl.viz_tier(vl.VIZ_TIER_MEDIUM_MAX + 1) == "large"


def test_empty_and_singleton_layouts() -> None:
    assert vl.clustered_force_layout(0, [], []).shape == (0, 2)
    pos = vl.clustered_force_layout(1, [], [0])
    assert pos.shape == (1, 2)
    assert np.isfinite(pos).all()


def test_layout_is_deterministic() -> None:
    n, edges, clusters = _two_cliques()
    a = vl.clustered_force_layout(n, edges, clusters, seed=3)
    b = vl.clustered_force_layout(n, edges, clusters, seed=3)
    assert np.array_equal(a, b)


def test_clusters_occupy_separate_regions() -> None:
    n, edges, clusters = _two_cliques()
    pos = vl.clustered_force_layout(n, edges, clusters)
    labels = np.asarray(clusters)
    ca = pos[labels == 0].mean(axis=0)
    cb = pos[labels == 1].mean(axis=0)
    for i in range(n):
        own, other = (ca, cb) if labels[i] == 0 else (cb, ca)
        assert np.linalg.norm(pos[i] - own) < np.linalg.norm(pos[i] - other)


def test_nodes_do_not_collapse() -> None:
    n, edges, clusters = _two_cliques()
    pos = vl.clustered_force_layout(n, edges, clusters)
    d = np.linalg.norm(pos[:, None, :] - pos[None, :, :], axis=-1)
    d[np.arange(n), np.arange(n)] = np.inf
    assert d.min() > 0.2 * vl.FORCE_LINK_DISTANCE


def _stacked_clusters() -> tuple[int, list[tuple[int, int]], list[int], list[int]]:
    """Three chained clusters; cluster k holds depth-k nodes."""
    size = 6
    n = 3 * size
    clusters = [i // size for i in range(n)]
    depths = list(clusters)
    edges: list[tuple[int, int]] = []
    for k in range(3):
        edges += _clique_edges(list(range(k * size, (k + 1) * size)))
    # Consumer -> producer: deeper clusters consume shallower ones.
    edges += [(size, 0), (2 * size, size)]
    return n, edges, clusters, depths


def test_rank_pull_between_clusters_orders_clusters_top_down() -> None:
    n, edges, clusters, depths = _stacked_clusters()
    pos = vl.clustered_force_layout(n, edges, clusters, depths=depths, rank_pull="between")
    labels = np.asarray(clusters)
    ys = [pos[labels == k, 1].mean() for k in range(3)]
    assert ys[0] < ys[1] < ys[2]


def test_rank_pull_everywhere_orders_chain_inside_one_cluster() -> None:
    n = 8
    edges = [(i + 1, i) for i in range(n - 1)]
    depths = list(range(n))
    pos = vl.clustered_force_layout(n, edges, [0] * n, depths=depths, rank_pull="everywhere")
    assert pos[-1, 1] > pos[0, 1]


def test_rank_pull_requires_depths() -> None:
    with pytest.raises(ValueError, match="depths"):
        vl.clustered_force_layout(2, [(0, 1)], [0, 0], rank_pull="between")


def test_rank_pull_none_ignores_depths() -> None:
    n, edges, clusters, depths = _stacked_clusters()
    pos = vl.clustered_force_layout(n, edges, clusters, depths=depths, rank_pull="none")
    assert np.isfinite(pos).all()


def test_input_depths_put_inputs_at_zero() -> None:
    # 2 consumes 1 consumes 0; 3 consumes 0.
    depths = vl.input_depths(4, [(2, 1), (1, 0), (3, 0)])
    assert list(depths) == [0, 1, 2, 1]


def test_input_depths_share_rank_inside_cycles() -> None:
    depths = vl.input_depths(3, [(0, 1), (1, 0), (2, 0)])
    assert depths[0] == depths[1] == 0
    assert depths[2] == 1


def test_grid_pairs_match_exact_pairs_within_cutoff() -> None:
    rng = np.random.default_rng(0)
    pos = rng.uniform(-200, 200, size=(300, 2))
    cutoff = 50.0
    i, j = vl._grid_pairs(pos, cutoff)
    got = {(min(a, b), max(a, b)) for a, b in zip(i.tolist(), j.tolist(), strict=True)}
    assert len(got) == len(i)
    d = np.linalg.norm(pos[:, None, :] - pos[None, :, :], axis=-1)
    want = {(a, b) for a in range(300) for b in range(a + 1, 300) if d[a, b] < cutoff}
    assert want <= got


def test_medium_and_large_tiers_run_on_sparse_graphs() -> None:
    rng = np.random.default_rng(1)
    n = 3000
    clusters = (np.arange(n) // 50).tolist()
    edges = [(i, i - 1) for i in range(1, n) if clusters[i] == clusters[i - 1]]
    edges += [(int(a), int(b)) for a, b in rng.integers(0, n, size=(200, 2)) if a != b]
    depths = vl.input_depths(n, edges)
    for tier in ("medium", "large"):
        pos = vl.clustered_force_layout(
            n, edges, clusters, depths=depths, rank_pull="between", tier=tier
        )
        assert pos.shape == (n, 2)
        assert np.isfinite(pos).all()
        assert np.unique(np.round(pos, 3), axis=0).shape[0] == n


def test_js_config_carries_shared_constants() -> None:
    cfg = vl.viz_layout_js_config()
    assert cfg["tier_small_max"] == vl.VIZ_TIER_SMALL_MAX
    assert cfg["tier_medium_max"] == vl.VIZ_TIER_MEDIUM_MAX
    assert cfg["link_distance"] == vl.FORCE_LINK_DISTANCE
    assert cfg["link_strength"] == vl.FORCE_LINK_STRENGTH
    assert cfg["charge"] == vl.FORCE_CHARGE
    assert cfg["camera_max_width"] == vl.VIZ_CAMERA_MAX_WIDTH
