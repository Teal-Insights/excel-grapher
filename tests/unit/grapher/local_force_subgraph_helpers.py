"""Reference implementation of lightweight-viz local force subgraph selection.

Mirrors ``localForceSubgraph()`` in ``lightweight_viz_template.html`` so tests can
assert the viewer's neighborhood expansion semantics without a browser runtime.
"""

from __future__ import annotations

from dataclasses import dataclass

from excel_grapher.grapher.lightweight_viz import (
    LightweightVizPayload,
    lightweight_viz_flat,
)


@dataclass(frozen=True, slots=True)
class LocalForceSubgraph:
    node_ids: tuple[int, ...]
    edges_from: tuple[int, ...]
    edges_to: tuple[int, ...]
    edges_guarded: tuple[bool, ...]
    truncated: bool


def select_local_force_subgraph(
    payload: LightweightVizPayload,
    *,
    node_id: int,
) -> LocalForceSubgraph:
    data = lightweight_viz_flat(payload)
    n = data.stats.node_count
    max_nodes = data.max_local_nodes if data.max_local_nodes is not None else n
    max_edges = data.max_local_edges
    if not (0 <= node_id < n):
        raise ValueError(f"node_id out of range: {node_id}")

    off = data.local_edges.offsets
    tg = data.local_edges.targets
    gd = data.local_edges.guarded
    incoming: list[list[tuple[int, bool]]] = [[] for _ in range(n)]
    for u in range(n):
        for k in range(off[u], off[u + 1]):
            incoming[tg[k]].append((u, gd[k]))

    seeds = {node_id}
    expanded: set[int] = set()
    while seeds and len(expanded) < max_nodes:
        u = min(seeds)
        seeds.discard(u)
        if u in expanded:
            continue
        expanded.add(u)
        for k in range(off[u], off[u + 1]):
            v = tg[k]
            if v not in expanded and len(expanded) + len(seeds) < max_nodes:
                seeds.add(v)
        for w, _ in incoming[u]:
            if w not in expanded and len(expanded) + len(seeds) < max_nodes:
                seeds.add(w)

    node_set = expanded
    edges_from: list[int] = []
    edges_to: list[int] = []
    edges_guarded: list[bool] = []
    truncated = not data.local_edges.complete[node_id]
    for u in sorted(node_set):
        for k in range(off[u], off[u + 1]):
            v = tg[k]
            if v not in node_set:
                continue
            edges_from.append(u)
            edges_to.append(v)
            edges_guarded.append(gd[k])
            if max_edges is not None and len(edges_from) >= max_edges:
                return LocalForceSubgraph(
                    node_ids=tuple(sorted(node_set)),
                    edges_from=tuple(edges_from),
                    edges_to=tuple(edges_to),
                    edges_guarded=tuple(edges_guarded),
                    truncated=True,
                )
    return LocalForceSubgraph(
        node_ids=tuple(sorted(node_set)),
        edges_from=tuple(edges_from),
        edges_to=tuple(edges_to),
        edges_guarded=tuple(edges_guarded),
        truncated=truncated,
    )
