"""Loose time budget for to_web_viz_payload (catches catastrophic slowdowns, not micro-benchmarks)."""

from __future__ import annotations

import time

import networkx as nx

from excel_grapher.exporter import to_web_viz_payload


def _chain_graph(n: int) -> nx.DiGraph:
    g = nx.DiGraph()
    for r in range(1, n + 1):
        is_leaf = r == 1
        formula = None if is_leaf else f"=A{r - 1}"
        g.add_node(
            f"S!A{r}",
            formula=formula,
            is_leaf=is_leaf,
            value=1 if is_leaf else None,
        )
    for r in range(2, n + 1):
        g.add_edge(f"S!A{r}", f"S!A{r - 1}")
    return g


def test_to_web_viz_payload_100_node_chain_completes_under_time_budget() -> None:
    # NetworkX multipartite layout imports NumPy internally.
    import pytest

    pytest.importorskip("numpy")
    g = _chain_graph(100)
    t0 = time.perf_counter()
    p = to_web_viz_payload(
        g,
        layout="multipartite",
        seed=0,
        include_module_overlay=True,
    )
    elapsed = time.perf_counter() - t0
    assert p.core.stats.node_count == 100
    assert elapsed < 30.0, f"to_web_viz_payload took {elapsed:.2f}s, expected < 30s"
