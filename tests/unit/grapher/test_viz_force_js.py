"""Shared in-browser relayout (`viz_force.js`) run under node."""

from __future__ import annotations

import json
import math
import shutil
import subprocess

import pytest

from excel_grapher.grapher.viz_layout import viz_layout_js_bootstrap, viz_layout_js_config


def _run(script: str, payload: dict) -> dict:
    node = shutil.which("node")
    if node is None:
        pytest.skip("node is required to execute viz_force.js")
    runner = (
        "globalThis.window = globalThis;\n"
        + viz_layout_js_bootstrap()
        + "\nconst input = JSON.parse(require('fs').readFileSync(0, 'utf8'));\n"
        + "const cfg = window.VIZ_LAYOUT_CONFIG;\n"
        + script
    )
    proc = subprocess.run(
        [node, "-e", runner], input=json.dumps(payload), capture_output=True, text=True
    )
    if proc.returncode != 0:
        raise AssertionError(proc.stderr or proc.stdout)
    return json.loads(proc.stdout)


_LAYOUT = """
const out = VizForce.clusteredLayout(input, cfg);
process.stdout.write(JSON.stringify({x: Array.from(out.x), y: Array.from(out.y)}));
"""


def _cliques(size: int = 8) -> dict:
    edges = []
    for base in (0, size):
        edges += [[base + a, base + b] for a in range(size) for b in range(a + 1, size)]
    edges.append([0, size])
    return {"n": 2 * size, "edges": edges, "clusters": [0] * size + [1] * size}


def test_bootstrap_injects_python_constants() -> None:
    out = _run("process.stdout.write(JSON.stringify(cfg));", {})
    assert out == json.loads(json.dumps(viz_layout_js_config()))


def test_tier_uses_injected_boundaries() -> None:
    cfg = viz_layout_js_config()
    sizes = [cfg["tier_small_max"], cfg["tier_small_max"] + 1, cfg["tier_medium_max"] + 1]
    out = _run(
        "process.stdout.write(JSON.stringify(input.sizes.map((s) => VizForce.tier(s, cfg))));",
        {"sizes": sizes},
    )
    assert out == ["small", "medium", "large"]


def test_clustered_layout_separates_clusters_deterministically() -> None:
    payload = _cliques()
    a = _run(_LAYOUT, payload)
    b = _run(_LAYOUT, payload)
    assert a == b
    n = payload["n"]
    xs, ys = a["x"], a["y"]
    assert all(math.isfinite(v) for v in xs + ys)
    labels = payload["clusters"]

    def centroid(k: int) -> tuple[float, float]:
        idx = [i for i in range(n) if labels[i] == k]
        return sum(xs[i] for i in idx) / len(idx), sum(ys[i] for i in idx) / len(idx)

    cents = [centroid(0), centroid(1)]
    for i in range(n):
        own, other = cents[labels[i]], cents[1 - labels[i]]
        assert math.dist((xs[i], ys[i]), own) < math.dist((xs[i], ys[i]), other)


def test_rank_pull_between_orders_clusters() -> None:
    size = 5
    edges = []
    for k in range(3):
        base = k * size
        edges += [[base + a, base + b] for a in range(size) for b in range(a + 1, size)]
    edges += [[size, 0], [2 * size, size]]
    clusters = [i // size for i in range(3 * size)]
    out = _run(
        _LAYOUT,
        {
            "n": 3 * size,
            "edges": edges,
            "clusters": clusters,
            "depths": clusters,
            "rankPull": "between",
        },
    )
    ys = [sum(out["y"][k * size : (k + 1) * size]) / size for k in range(3)]
    assert ys[0] < ys[1] < ys[2]


def _role_dag() -> dict:
    size = 5
    edges = []
    for k in range(3):
        base = k * size
        edges += [[base + a, base + b] for a in range(size) for b in range(a + 1, size)]
    edges += [[size, 0], [2 * size, size], [2 * size + 1, 1]]
    clusters = [i // size for i in range(3 * size)]
    return {"n": 3 * size, "edges": edges, "clusters": clusters, "depths": clusters}


def _cluster_x(out: dict, clusters: list[int]) -> list[float]:
    return [
        sum(x for x, c in zip(out["x"], clusters, strict=True) if c == k) / clusters.count(k)
        for k in range(3)
    ]


@pytest.mark.parametrize("pull", ["between", "everywhere"])
def test_rank_pull_aligns_cluster_centroids_on_cross_axis(pull: str) -> None:
    payload = {**_role_dag(), "rankPull": pull}
    out = _run(_LAYOUT, payload)
    assert all(abs(x) < 1e-9 for x in _cluster_x(out, payload["clusters"]))
    size = 5
    ys = [sum(out["y"][k * size : (k + 1) * size]) / size for k in range(3)]
    gap = viz_layout_js_config()["link_distance"]
    assert ys[0] + gap < ys[1] and ys[1] + gap < ys[2]


def test_align_clusters_can_be_turned_off() -> None:
    payload = {**_role_dag(), "rankPull": "everywhere", "alignClusters": False}
    xs = _cluster_x(_run(_LAYOUT, payload), payload["clusters"])
    assert max(xs) - min(xs) > 1.0


def test_relayout_group_moves_only_members_and_keeps_centroid() -> None:
    script = """
    const x = input.x.slice();
    const y = input.y.slice();
    VizForce.relayoutGroup({ids: input.ids, edges: input.edges, x: x, y: y}, cfg);
    process.stdout.write(JSON.stringify({x: x, y: y}));
    """
    x = [0.0, 0.0, 0.0, 500.0]
    y = [0.0, 1.0, 2.0, 500.0]
    out = _run(script, {"ids": [0, 1, 2], "edges": [[0, 1], [1, 2], [2, 3]], "x": x, "y": y})
    assert out["x"][3] == 500.0 and out["y"][3] == 500.0
    assert abs(sum(out["x"][:3]) / 3 - 0.0) < 1e-6
    assert abs(sum(out["y"][:3]) / 3 - 1.0) < 1e-6
    assert math.dist((out["x"][0], out["y"][0]), (out["x"][2], out["y"][2])) > 10
