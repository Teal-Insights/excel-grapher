"""Regression tests for lightweight-viz local force neighborhood selection."""

from __future__ import annotations

import importlib.resources
import re

import networkx as nx

from excel_grapher.exporter import to_web_viz_payload
from excel_grapher.grapher.lightweight_viz import lightweight_viz_flat


def _local_force_subgraph_source() -> str:
    text = (
        importlib.resources.files("excel_grapher.grapher")
        .joinpath("lightweight_viz_template.html")
        .read_text(encoding="utf-8")
    )
    match = re.search(
        r"function localForceSubgraph\(data, nodeId\) \{(.*?)\n  \}\n\n  async function loadData",
        text,
        flags=re.DOTALL,
    )
    assert match is not None, "localForceSubgraph definition missing from viewer template"
    return match.group(1)


def test_local_force_template_avoids_module_scope_shortcut() -> None:
    source = _local_force_subgraph_source()
    assert "mod.node_count <= maxN" not in source
    assert "moduleScope: true" not in source
    assert "incoming[tg[k]].push" in source
    assert "moduleScope: false" in source


def _chain_nx() -> nx.DiGraph:
    g = nx.DiGraph()
    g.add_node("S!A1", formula=None, value=1, is_leaf=True, sheet="S", column="A", row=1)
    g.add_node("S!A2", formula="=A1", value=None, is_leaf=False, sheet="S", column="A", row=2)
    g.add_node("S!A3", formula="=A2", value=None, is_leaf=False, sheet="S", column="A", row=3)
    g.add_edge("S!A3", "S!A2")
    g.add_edge("S!A2", "S!A1")
    return g


def test_louvain_chain_exports_cross_module_local_edges() -> None:
    """Tail-node selection needs the exported CSR edge across the module split."""
    flat = lightweight_viz_flat(to_web_viz_payload(_chain_nx(), layout="stratified_multipartite"))
    assert flat.nodes.module_id == (0, 0, 1)

    off = flat.local_edges.offsets
    tg = flat.local_edges.targets
    assert list(tg[off[2] : off[3]]) == [1]
    assert list(tg[off[1] : off[2]]) == [0]
