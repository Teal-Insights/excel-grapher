"""Regression tests for lightweight-viz viewer styling enhancements (issue #110)."""

from __future__ import annotations

import importlib.resources
import re

import networkx as nx

from excel_grapher.exporter import to_web_viz_payload
from excel_grapher.grapher.export import to_networkx
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.lightweight_viz import lightweight_viz_flat
from excel_grapher.grapher.node import Node


def _viewer_template() -> str:
    return (
        importlib.resources.files("excel_grapher.grapher")
        .joinpath("lightweight_viz_template.html")
        .read_text(encoding="utf-8")
    )


def test_template_draws_guarded_edges_with_distinct_style() -> None:
    text = _viewer_template()
    assert "setLineDash" in text
    assert re.search(r"ed\.guarded|edge\.guarded", text)
    # Batched passes: unguarded solid, guarded dashed (avoid per-edge dash thrash).
    assert "guardedEdges" in text or "unguarded" in text.lower()


def test_template_colors_formula_and_target_nodes_distinctly() -> None:
    text = _viewer_template()
    assert "is_leaf" in text
    assert "is_target" in text
    assert "nodeBaseColor" in text or "nodeColor" in text


def test_template_lights_up_selection_neighborhood_edges() -> None:
    text = _viewer_template()
    assert "highlight" in text
    draw_overlay = re.search(
        r"function drawOverlay\(\) \{(.*?)\n    \}",
        text,
        flags=re.DOTALL,
    )
    assert draw_overlay is not None, "drawOverlay definition missing from viewer template"
    body = draw_overlay.group(1)
    assert "highlight" in body
    assert "neigh" in body or "highlight.has" in body


def test_to_networkx_exports_is_target() -> None:
    graph = DependencyGraph()
    graph.add_node(
        Node(
            sheet="S",
            column="A",
            row=1,
            formula=None,
            normalized_formula=None,
            value=1,
            is_leaf=True,
            is_target=False,
        )
    )
    graph.add_node(
        Node(
            sheet="S",
            column="A",
            row=2,
            formula="=A1",
            normalized_formula="=A1",
            value=None,
            is_leaf=False,
            is_target=True,
        )
    )
    graph.add_edge("S!A2", "S!A1")
    nx_g = to_networkx(graph)
    assert nx_g.nodes["S!A2"]["is_target"] is True
    assert nx_g.nodes["S!A1"]["is_target"] is False


def test_web_viz_payload_preserves_is_target() -> None:
    g = nx.DiGraph()
    g.add_node(
        "S!A1", formula=None, value=1, is_leaf=True, is_target=False, sheet="S", column="A", row=1
    )
    g.add_node(
        "S!A2",
        formula="=A1",
        value=None,
        is_leaf=False,
        is_target=True,
        sheet="S",
        column="A",
        row=2,
    )
    g.add_edge("S!A2", "S!A1")
    flat = lightweight_viz_flat(to_web_viz_payload(g, layout="stratified_multipartite"))
    assert flat.nodes.is_target == (False, True)
    assert flat.nodes.is_leaf == (True, False)
