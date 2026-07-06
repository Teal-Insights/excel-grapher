from __future__ import annotations

import networkx as nx
import pytest

from excel_grapher.exporter import to_web_viz_payload
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.lightweight_viz import (
    LightweightVizPayload,
    VizLimits,
    assemble_lightweight_viz_payload,
    build_lightweight_viz_core,
    lightweight_viz_flat,
)
from excel_grapher.grapher.node import Node
from tests.unit.grapher.local_force_subgraph_helpers import (
    LocalForceSubgraph,
    select_local_force_subgraph,
)


def _leaf(key: str, value: object = 0) -> Node:
    sheet, addr = key.split("!")
    col = "".join(ch for ch in addr if ch.isalpha())
    row = int("".join(ch for ch in addr if ch.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=None,
        normalized_formula=None,
        value=value,
        is_leaf=True,
    )


def _formula(key: str, formula: str) -> Node:
    sheet, addr = key.split("!")
    col = "".join(ch for ch in addr if ch.isalpha())
    row = int("".join(ch for ch in addr if ch.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=None,
        is_leaf=False,
    )


def _chain_graph() -> DependencyGraph:
    g = DependencyGraph()
    for node in [
        _leaf("S!A1", 1),
        _formula("S!A2", "=A1"),
        _formula("S!A3", "=A2"),
    ]:
        g.add_node(node)
    g.add_edge("S!A3", "S!A2")
    g.add_edge("S!A2", "S!A1")
    return g


def _chain_nx() -> nx.DiGraph:
    g = nx.DiGraph()
    g.add_node("S!A1", formula=None, value=1, is_leaf=True, sheet="S", column="A", row=1)
    g.add_node("S!A2", formula="=A1", value=None, is_leaf=False, sheet="S", column="A", row=2)
    g.add_node("S!A3", formula="=A2", value=None, is_leaf=False, sheet="S", column="A", row=3)
    g.add_edge("S!A3", "S!A2")
    g.add_edge("S!A2", "S!A1")
    return g


def _payload(
    g: DependencyGraph,
    *,
    limits: VizLimits | None = None,
) -> LightweightVizPayload:
    core = build_lightweight_viz_core(g, limits=limits or VizLimits(), layout_input=None)
    return assemble_lightweight_viz_payload(core, [])


def _edge_pairs(sub: LocalForceSubgraph) -> list[tuple[int, int]]:
    return list(zip(sub.edges_from, sub.edges_to, strict=True))


def test_select_local_force_subgraph_expands_across_module_boundary() -> None:
    """Louvain-style splits must not isolate the selected node from its local deps."""
    payload = to_web_viz_payload(_chain_nx(), layout="stratified_multipartite")
    flat = lightweight_viz_flat(payload)
    assert flat.nodes.module_id == (0, 0, 1)

    sub = select_local_force_subgraph(payload, node_id=2)

    assert sub.node_ids == (0, 1, 2)
    assert _edge_pairs(sub) == [(1, 0), (2, 1)]
    assert sub.truncated is False


def test_select_local_force_subgraph_includes_precedents_from_leaf() -> None:
    payload = to_web_viz_payload(_chain_nx(), layout="stratified_multipartite")

    sub = select_local_force_subgraph(payload, node_id=0)

    assert sub.node_ids == (0, 1, 2)
    assert _edge_pairs(sub) == [(1, 0), (2, 1)]


def test_select_local_force_subgraph_respects_max_nodes() -> None:
    payload = _payload(_chain_graph(), limits=VizLimits(max_local_nodes=2))

    sub = select_local_force_subgraph(payload, node_id=2)

    assert sub.node_ids == (1, 2)
    assert _edge_pairs(sub) == [(2, 1)]


def test_select_local_force_subgraph_respects_max_edges() -> None:
    g = DependencyGraph()
    nodes = [_leaf("S!A1", 1)]
    for i in range(2, 7):
        nodes.append(_formula(f"S!A{i}", f"=A{i - 1}"))
    for node in nodes:
        g.add_node(node)
    for i in range(2, 7):
        g.add_edge(f"S!A{i}", f"S!A{i - 1}")

    payload = _payload(g, limits=VizLimits(max_local_edges=2))

    sub = select_local_force_subgraph(payload, node_id=5)

    assert len(sub.edges_from) == 2
    assert sub.truncated is True


def test_select_local_force_subgraph_reports_truncated_when_local_edges_incomplete() -> None:
    g = DependencyGraph()
    g.add_node(_leaf("S!A1"))
    g.add_node(_formula("S!A2", "=A1"))
    g.add_node(_formula("S!A3", "=A1"))
    g.add_node(_formula("S!A4", "=A1"))
    g.add_node(_formula("S!A5", "=A2+A3+A4"))
    g.add_edge("S!A2", "S!A1")
    g.add_edge("S!A3", "S!A1")
    g.add_edge("S!A4", "S!A1")
    g.add_edge("S!A5", "S!A2")
    g.add_edge("S!A5", "S!A3")
    g.add_edge("S!A5", "S!A4")

    payload = _payload(g, limits=VizLimits(max_local_edges=1))
    flat = lightweight_viz_flat(payload)
    assert flat.local_edges.complete[4] is False

    sub = select_local_force_subgraph(payload, node_id=4)

    assert sub.truncated is True


def test_select_local_force_subgraph_raises_for_invalid_node_id() -> None:
    payload = _payload(_chain_graph())

    with pytest.raises(ValueError, match="node_id out of range"):
        select_local_force_subgraph(payload, node_id=3)


def test_select_local_force_subgraph_is_deterministic() -> None:
    payload = to_web_viz_payload(_chain_nx(), layout="stratified_multipartite")

    sub_a = select_local_force_subgraph(payload, node_id=1)
    sub_b = select_local_force_subgraph(payload, node_id=1)

    assert sub_a == sub_b
