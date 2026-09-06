from __future__ import annotations

import pytest

from excel_grapher.grapher import select_shortest_path_subgraph
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.guard import CellRef as GuardCellRef
from excel_grapher.grapher.guard import Compare, Literal
from excel_grapher.grapher.node import Node


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


def _add_nodes(g: DependencyGraph, nodes: list[Node]) -> None:
    for node in nodes:
        g.add_node(node)


def _sorted_edges(g: DependencyGraph) -> list[tuple[str, str]]:
    edges: list[tuple[str, str]] = []
    for key in sorted(g):
        for dep in sorted(g.get_dependencies(key)):
            edges.append((key, dep))
    return edges


def _directed_chain() -> DependencyGraph:
    g = DependencyGraph()
    _add_nodes(
        g,
        [
            _formula("S!F1", "=D1"),
            _formula("S!D1", "=A1"),
            _leaf("S!A1", 1),
        ],
    )
    g.add_edge("S!F1", "S!D1")
    g.add_edge("S!D1", "S!A1")
    return g


def _directed_diamond_with_bypass() -> DependencyGraph:
    g = DependencyGraph()
    _add_nodes(
        g,
        [
            _formula("S!F1", "=D1+X1"),
            _formula("S!D1", "=B1+C1"),
            _formula("S!B1", "=A1"),
            _formula("S!C1", "=A1"),
            _formula("S!X1", "=Y1"),
            _formula("S!Y1", "=Z1"),
            _formula("S!Z1", "=A1"),
            _leaf("S!A1", 1),
        ],
    )
    for frm, to in [
        ("S!F1", "S!D1"),
        ("S!F1", "S!X1"),
        ("S!D1", "S!B1"),
        ("S!D1", "S!C1"),
        ("S!B1", "S!A1"),
        ("S!C1", "S!A1"),
        ("S!X1", "S!Y1"),
        ("S!Y1", "S!Z1"),
        ("S!Z1", "S!A1"),
    ]:
        g.add_edge(frm, to)
    return g


def _siblings() -> DependencyGraph:
    g = DependencyGraph()
    _add_nodes(
        g,
        [
            _formula("S!A1", "=B1+C1"),
            _leaf("S!B1", 1),
            _leaf("S!C1", 2),
        ],
    )
    g.add_edge("S!A1", "S!B1")
    g.add_edge("S!A1", "S!C1")
    return g


def _shared_input_and_distant_consumer() -> DependencyGraph:
    """B and C share a nearby input and a distant consumer.

    Undirected hop-shortest B<->C is through SharedInput (2 hops), not
    through FarConsumer (6 hops).
    """
    g = DependencyGraph()
    _add_nodes(
        g,
        [
            _formula("S!F9", "=M1+M3"),
            _formula("S!M1", "=M2"),
            _formula("S!M2", "=B1"),
            _formula("S!M3", "=M4"),
            _formula("S!M4", "=C1"),
            _formula("S!B1", "=In1"),
            _formula("S!C1", "=In1"),
            _leaf("S!In1", 1),
        ],
    )
    for frm, to in [
        ("S!F9", "S!M1"),
        ("S!F9", "S!M3"),
        ("S!M1", "S!M2"),
        ("S!M2", "S!B1"),
        ("S!M3", "S!M4"),
        ("S!M4", "S!C1"),
        ("S!B1", "S!In1"),
        ("S!C1", "S!In1"),
    ]:
        g.add_edge(frm, to)
    return g


def test_select_shortest_path_subgraph_directed_unique_path() -> None:
    g = _directed_chain()

    sub = select_shortest_path_subgraph(g, source_key="S!F1", target_key="S!A1")

    assert sorted(list(sub)) == ["S!A1", "S!D1", "S!F1"]
    assert _sorted_edges(sub) == [("S!D1", "S!A1"), ("S!F1", "S!D1")]


def test_select_shortest_path_subgraph_keeps_all_shortest_excludes_longer() -> None:
    g = _directed_diamond_with_bypass()

    sub = select_shortest_path_subgraph(g, source_key="S!F1", target_key="S!A1")

    assert sorted(list(sub)) == ["S!A1", "S!B1", "S!C1", "S!D1", "S!F1"]
    assert _sorted_edges(sub) == [
        ("S!B1", "S!A1"),
        ("S!C1", "S!A1"),
        ("S!D1", "S!B1"),
        ("S!D1", "S!C1"),
        ("S!F1", "S!D1"),
    ]


def test_select_shortest_path_subgraph_siblings_require_undirected() -> None:
    g = _siblings()

    with pytest.raises(ValueError, match="directed=False"):
        select_shortest_path_subgraph(g, source_key="S!B1", target_key="S!C1")

    sub = select_shortest_path_subgraph(g, source_key="S!B1", target_key="S!C1", directed=False)

    assert sorted(list(sub)) == ["S!A1", "S!B1", "S!C1"]
    assert _sorted_edges(sub) == [("S!A1", "S!B1"), ("S!A1", "S!C1")]


def test_select_shortest_path_subgraph_undirected_prefers_hop_shortest() -> None:
    g = _shared_input_and_distant_consumer()

    sub = select_shortest_path_subgraph(g, source_key="S!B1", target_key="S!C1", directed=False)

    assert sorted(list(sub)) == ["S!B1", "S!C1", "S!In1"]
    assert _sorted_edges(sub) == [("S!B1", "S!In1"), ("S!C1", "S!In1")]
    assert "S!F9" not in sub


def test_select_shortest_path_subgraph_same_key_is_single_node() -> None:
    g = _siblings()

    sub = select_shortest_path_subgraph(g, source_key="S!B1", target_key="S!B1")
    same = select_shortest_path_subgraph(g, source_key="S!B1", target_key="S!B1", max_path_length=0)

    assert sorted(list(sub)) == ["S!B1"]
    assert _sorted_edges(sub) == []
    assert sorted(list(same)) == ["S!B1"]


def test_select_shortest_path_subgraph_raises_for_missing_keys() -> None:
    g = _siblings()

    with pytest.raises(ValueError, match="not present"):
        select_shortest_path_subgraph(g, source_key="S!MISSING", target_key="S!B1")
    with pytest.raises(ValueError, match="not present"):
        select_shortest_path_subgraph(g, source_key="S!B1", target_key="S!MISSING")


def test_select_shortest_path_subgraph_raises_when_disconnected() -> None:
    g = _siblings()
    g.add_node(_leaf("S!X1", 9))

    with pytest.raises(ValueError, match="no path"):
        select_shortest_path_subgraph(g, source_key="S!B1", target_key="S!X1", directed=False)


def test_select_shortest_path_subgraph_enforces_max_path_length() -> None:
    g = _directed_chain()

    with pytest.raises(ValueError, match="max_path_length"):
        select_shortest_path_subgraph(g, source_key="S!F1", target_key="S!A1", max_path_length=1)
    with pytest.raises(ValueError, match="max_path_length"):
        select_shortest_path_subgraph(g, source_key="S!F1", target_key="S!A1", max_path_length=-1)

    sub = select_shortest_path_subgraph(g, source_key="S!F1", target_key="S!A1", max_path_length=2)
    assert sorted(list(sub)) == ["S!A1", "S!D1", "S!F1"]


def test_select_shortest_path_subgraph_preserves_guard_and_provenance() -> None:
    g = DependencyGraph()
    g.add_node(_formula("S!F1", "=D1"))
    g.add_node(_formula("S!D1", "=B1"))
    g.add_node(_leaf("S!B1", 1))
    guard = Compare(left=GuardCellRef(key="S!A1"), op="=", right=Literal(1))
    provenance = EdgeProvenance(causes=DependencyCause.direct_ref)
    g.add_edge("S!F1", "S!D1")
    g.add_edge("S!D1", "S!B1", guard=guard, provenance=provenance)

    sub = select_shortest_path_subgraph(g, source_key="S!F1", target_key="S!B1")

    attrs = sub.get_edge_attrs("S!D1", "S!B1")
    assert attrs.guard == guard
    assert attrs.provenance == provenance


def test_select_shortest_path_subgraph_is_deterministic_across_insertion_order() -> None:
    g1 = _directed_diamond_with_bypass()
    g2 = DependencyGraph()
    for node in [
        _leaf("S!A1", 1),
        _formula("S!Z1", "=A1"),
        _formula("S!Y1", "=Z1"),
        _formula("S!X1", "=Y1"),
        _formula("S!C1", "=A1"),
        _formula("S!B1", "=A1"),
        _formula("S!D1", "=B1+C1"),
        _formula("S!F1", "=D1+X1"),
    ]:
        g2.add_node(node)
    for frm, to in [
        ("S!Z1", "S!A1"),
        ("S!C1", "S!A1"),
        ("S!B1", "S!A1"),
        ("S!Y1", "S!Z1"),
        ("S!X1", "S!Y1"),
        ("S!D1", "S!C1"),
        ("S!D1", "S!B1"),
        ("S!F1", "S!X1"),
        ("S!F1", "S!D1"),
    ]:
        g2.add_edge(frm, to)

    sub1 = select_shortest_path_subgraph(g1, source_key="S!F1", target_key="S!A1")
    sub2 = select_shortest_path_subgraph(g2, source_key="S!F1", target_key="S!A1")

    assert sorted(list(sub1)) == sorted(list(sub2))
    assert _sorted_edges(sub1) == _sorted_edges(sub2)
