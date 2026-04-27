"""Sample LIC-DSF viz HTML: inline payload contract after regeneration (integration, slow).

Rebuilds or validates cached graph artifacts and checks ``__VIZ_DATA__`` in emitted
HTML matches export range expectations for the sample visualization contract.
"""

from __future__ import annotations

import json
import pickle
import subprocess
from pathlib import Path

import pytest


def _extract_inline_payload(html_path: Path) -> dict:
    prefix = "window.__VIZ_DATA__ = "
    with html_path.open("r", encoding="utf-8") as f:
        for line in f:
            stripped = line.strip()
            if stripped.startswith(prefix):
                payload = stripped[len(prefix) :]
                if payload.endswith(";"):
                    payload = payload[:-1]
                return json.loads(payload)
    raise AssertionError(f"Could not find inline __VIZ_DATA__ in {html_path}")


@pytest.mark.slow
def test_regenerate_sample_viz_full_contract() -> None:
    repo_root = Path(__file__).resolve().parents[3]
    cache_path = (
        repo_root
        / "examples"
        / "lic_dsf"
        / ".cache"
        / "lic-dsf-template-2025-08-12-dependency-graph.pkl"
    )
    if not cache_path.is_file():
        pytest.skip("sample cache pickle missing; skipping full regenerate contract test")

    html_path = (
        repo_root / "examples" / "lic_dsf" / "data" / "lic-dsf-template-sample-exported-viz.html"
    )
    subprocess.run(
        ["uv", "run", "examples/lic_dsf/regenerate_sample_viz.py", "--full"],
        cwd=repo_root,
        check=True,
        capture_output=True,
        text=True,
    )
    payload = _extract_inline_payload(html_path)
    core = payload["core"]

    # 2) Expected full-node graph size for this sample cache.
    assert core["stats"]["node_count"] == 132068

    # Build index->key mapping from the same cached graph that the script uses.
    with cache_path.open("rb") as f:
        _, graph = pickle.load(f)
    keys = sorted(graph)
    assert len(keys) == core["stats"]["node_count"]

    # 1) Rank metadata should remain populated and non-degenerate.
    ranks = core["nodes"]["rank"]
    assert len(ranks) == len(keys)
    assert min(ranks) == 0
    assert max(ranks) >= 1

    # 3) There should be incoming edges to deepest-rank nodes in the exported graph.
    max_rank = max(ranks)
    deepest_ids = {i for i, r in enumerate(ranks) if r == max_rank}
    offsets = core["local_edges"]["offsets"]
    targets = core["local_edges"]["targets"]
    incoming_to_deepest = 0
    for src in range(core["stats"]["node_count"]):
        for k in range(offsets[src], offsets[src + 1]):
            if targets[k] in deepest_ids:
                incoming_to_deepest += 1
    assert incoming_to_deepest > 0
