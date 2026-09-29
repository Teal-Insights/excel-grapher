"""Operation counts for to_web_viz_payload (layout once per payload)."""

from __future__ import annotations

import networkx as nx
import pytest

from excel_grapher.exporter import to_web_viz_payload
from excel_grapher.exporter.web_viz_layout import run_web_viz_layout


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


def test_to_web_viz_payload_runs_layout_once_for_100_node_chain(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    pytest.importorskip("numpy")
    import excel_grapher.grapher.viz_layout as viz_layout

    g = _chain_graph(100)
    layout_ops = {"plugin": 0, "force": 0}
    original = run_web_viz_layout
    original_force = viz_layout.clustered_force_layout

    def counting(ctx: object, layout: object, layout_config: object) -> object:
        layout_ops["plugin"] += 1
        return original(ctx, layout, layout_config)

    def counting_force(*args: object, **kwargs: object) -> object:
        layout_ops["force"] += 1
        return original_force(*args, **kwargs)

    monkeypatch.setattr("excel_grapher.exporter.web_viz_layout.run_web_viz_layout", counting)
    monkeypatch.setattr(viz_layout, "clustered_force_layout", counting_force)
    p = to_web_viz_payload(g, seed=0, include_module_overlay=True)
    assert p.core.stats.node_count == 100
    assert layout_ops == {"plugin": 1, "force": 1}
