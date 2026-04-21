from __future__ import annotations

import json
import pickle
import subprocess
from pathlib import Path

import pytest

from example.extract_graph_cached import EXPORT_RANGES, cells_in_range, parse_range_spec


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


def _configured_targets() -> list[str]:
    targets: list[str] = []
    seen: set[str] = set()
    for entry in EXPORT_RANGES:
        sheet, a1 = parse_range_spec(entry["range_spec"])
        for key in cells_in_range(sheet, a1):
            if key in seen:
                continue
            seen.add(key)
            targets.append(key)
    return targets


@pytest.mark.slow
def test_regenerate_sample_viz_full_contract() -> None:
    repo_root = Path(__file__).resolve().parents[2]
    cache_path = repo_root / "example" / ".cache" / "lic-dsf-template-2025-08-12-dependency-graph.pkl"
    if not cache_path.is_file():
        pytest.skip("sample cache pickle missing; skipping full regenerate contract test")

    html_path = repo_root / "example" / "data" / "lic-dsf-template-sample-exported-viz.html"
    subprocess.run(
        ["uv", "run", "example/regenerate_sample_viz.py", "--full"],
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

    # 1) Rank-0 nodes should be exactly the configured extraction targets.
    ranks = core["nodes"]["rank"]
    rank0_ids = [i for i, r in enumerate(ranks) if r == min(ranks)]
    rank0_keys = {keys[i] for i in rank0_ids}
    target_keys = set(_configured_targets())
    assert rank0_keys <= target_keys
    assert target_keys <= rank0_keys

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

