from __future__ import annotations

import pytest

from excel_grapher.grapher import select_path_induced_subgraph
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


def _build_branching_graph() -> DependencyGraph:
    g = DependencyGraph()
    for node in [
        _formula("S!F1", "=D1"),
        _formula("S!D1", "=B1+C1"),
        _formula("S!C1", "=A1"),
        _formula("S!B1", "=A1"),
        _leaf("S!A1", 1),
        _formula("S!X1", "=Y1"),
        _leaf("S!Y1", 2),
    ]:
        g.add_node(node)
    g.add_edge("S!F1", "S!D1")
    g.add_edge("S!D1", "S!B1")
    g.add_edge("S!D1", "S!C1")
    g.add_edge("S!B1", "S!A1")
    g.add_edge("S!C1", "S!A1")
    g.add_edge("S!X1", "S!Y1")
    return g


def _sorted_edges(g: DependencyGraph) -> list[tuple[str, str]]:
    edges: list[tuple[str, str]] = []
    for key in sorted(g):
        for dep in sorted(g.get_dependencies(key)):
            edges.append((key, dep))
    return edges


def test_select_path_induced_subgraph_single_source_target() -> None:
    g = _build_branching_graph()

    sub = select_path_induced_subgraph(g, source_keys=["S!F1"], target_keys=["S!A1"])

    assert sorted(list(sub)) == ["S!A1", "S!B1", "S!C1", "S!D1", "S!F1"]
    assert _sorted_edges(sub) == [
        ("S!B1", "S!A1"),
        ("S!C1", "S!A1"),
        ("S!D1", "S!B1"),
        ("S!D1", "S!C1"),
        ("S!F1", "S!D1"),
    ]


def test_select_path_induced_subgraph_many_to_many_sets() -> None:
    g = _build_branching_graph()

    sub = select_path_induced_subgraph(
        g,
        source_keys=["S!F1", "S!C1"],
        target_keys=["S!A1", "S!B1"],
    )

    assert sorted(list(sub)) == ["S!A1", "S!B1", "S!C1", "S!D1", "S!F1"]
    assert ("S!X1", "S!Y1") not in _sorted_edges(sub)


def test_select_path_induced_subgraph_can_exclude_endpoints() -> None:
    g = _build_branching_graph()

    sub = select_path_induced_subgraph(
        g,
        source_keys=["S!F1"],
        target_keys=["S!A1"],
        include_endpoints=False,
    )

    assert sorted(list(sub)) == ["S!B1", "S!C1", "S!D1"]
    assert _sorted_edges(sub) == [("S!D1", "S!B1"), ("S!D1", "S!C1")]


def test_select_path_induced_subgraph_preserves_guard_and_provenance() -> None:
    g = DependencyGraph()
    g.add_node(_formula("S!F1", "=D1"))
    g.add_node(_formula("S!D1", "=B1"))
    g.add_node(_leaf("S!B1", 1))
    guard = Compare(left=GuardCellRef(key="S!A1"), op="=", right=Literal(1))
    provenance = EdgeProvenance(causes=frozenset({DependencyCause.direct_ref}))
    g.add_edge("S!F1", "S!D1")
    g.add_edge("S!D1", "S!B1", guard=guard, provenance=provenance)

    sub = select_path_induced_subgraph(g, source_keys=["S!F1"], target_keys=["S!B1"])

    attrs = sub.get_edge_attrs("S!D1", "S!B1")
    assert attrs.guard == guard
    assert attrs.provenance == provenance


def test_select_path_induced_subgraph_is_deterministic_across_insertion_order() -> None:
    g1 = _build_branching_graph()
    g2 = DependencyGraph()
    for node in [
        _leaf("S!Y1", 2),
        _formula("S!X1", "=Y1"),
        _leaf("S!A1", 1),
        _formula("S!B1", "=A1"),
        _formula("S!C1", "=A1"),
        _formula("S!D1", "=B1+C1"),
        _formula("S!F1", "=D1"),
    ]:
        g2.add_node(node)
    for frm, to in [
        ("S!C1", "S!A1"),
        ("S!X1", "S!Y1"),
        ("S!D1", "S!C1"),
        ("S!F1", "S!D1"),
        ("S!B1", "S!A1"),
        ("S!D1", "S!B1"),
    ]:
        g2.add_edge(frm, to)

    sub1 = select_path_induced_subgraph(g1, source_keys=["S!F1"], target_keys=["S!A1"])
    sub2 = select_path_induced_subgraph(g2, source_keys=["S!F1"], target_keys=["S!A1"])

    assert sorted(list(sub1)) == sorted(list(sub2))
    assert _sorted_edges(sub1) == _sorted_edges(sub2)


def test_select_path_induced_subgraph_enforces_max_path_length() -> None:
    g = _build_branching_graph()

    with pytest.raises(ValueError, match="max_path_length"):
        select_path_induced_subgraph(
            g,
            source_keys=["S!F1"],
            target_keys=["S!A1"],
            max_path_length=2,
        )


def test_select_path_induced_subgraph_enforces_max_paths() -> None:
    g = _build_branching_graph()

    with pytest.raises(ValueError, match="max_paths"):
        select_path_induced_subgraph(
            g,
            source_keys=["S!F1"],
            target_keys=["S!A1"],
            max_paths=1,
        )


def test_select_path_induced_subgraph_raises_for_missing_or_empty_inputs() -> None:
    g = _build_branching_graph()

    with pytest.raises(ValueError, match="source_keys"):
        select_path_induced_subgraph(g, source_keys=[], target_keys=["S!A1"])
    with pytest.raises(ValueError, match="target_keys"):
        select_path_induced_subgraph(g, source_keys=["S!F1"], target_keys=[])
    with pytest.raises(ValueError, match="not present"):
        select_path_induced_subgraph(g, source_keys=["S!MISSING"], target_keys=["S!A1"])
