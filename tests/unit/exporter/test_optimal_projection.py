"""Unit tests for optimal compression projection."""

from __future__ import annotations

from excel_grapher.exporter.projection import (
    BaseProjectionManifest,
    OptimalCompression,
    resolve_projection_manifest,
)
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node


def _make_node(
    key: str,
    formula: str | None,
    normalized: str | None,
    *,
    is_leaf: bool = False,
    is_target: bool = False,
) -> Node:
    sheet, rest = key.split("!", 1)
    col = "".join(c for c in rest if c.isalpha())
    row = int("".join(c for c in rest if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=normalized,
        value=None,
        is_leaf=is_leaf,
        is_target=is_target,
    )


def _direct_edge(graph: DependencyGraph, dependent: str, precedent: str) -> None:
    dr = DependencyCause.direct_ref
    dep_node = graph.get_node(dependent)
    assert dep_node is not None
    f = dep_node.formula
    n = dep_node.normalized_formula
    assert f is not None and n is not None
    ref = precedent
    i_f = f.index(ref)
    i_n = n.index(ref)
    graph.add_edge(
        dependent,
        precedent,
        provenance=EdgeProvenance(
            causes=frozenset({dr}),
            direct_sites_formula=((i_f, i_f + len(ref)),),
            direct_sites_normalized=((i_n, i_n + len(ref)),),
        ),
    )


def test_optimal_projection_does_not_mutate_original_graph() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    for n in (d, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    original_keys = set(graph)
    projection = OptimalCompression().project(graph)

    assert set(graph) == original_keys
    assert "Sheet1!B1" in graph
    assert "Sheet1!B1" not in projection


def test_optimal_projection_manifest_kind_and_forwarding() -> None:
    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (c, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    manifest = OptimalCompression().project(graph).manifest
    projection = OptimalCompression().project(graph)
    assert isinstance(manifest, BaseProjectionManifest)
    assert manifest.kind == "optimal_compression"
    assert manifest.map_to_projected("Sheet1!B1") == "Sheet1!C1"
    assert "Sheet1!C1" in projection


def test_optimal_projection_inline_lineage_without_forwarding() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    for n in (d, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    manifest = OptimalCompression().project(graph).manifest
    assert isinstance(manifest, BaseProjectionManifest)
    assert manifest.map_to_projected("Sheet1!B1") == "Sheet1!B1"
    assert manifest.retained_to_collapsed_sources["Sheet1!A1"] == ("Sheet1!B1",)
    assert "Sheet1!B1" in manifest.removed_node_snapshots


def test_optimal_projection_resolves_chained_inline_to_final_retained() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    c = _make_node("Sheet1!C1", "=Sheet1!D1+1", "=Sheet1!D1+1")
    b = _make_node("Sheet1!B1", "=Sheet1!C1*2", "=Sheet1!C1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+3", "=Sheet1!B1+3")
    for n in (d, c, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!C1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    manifest = OptimalCompression().project(graph).manifest
    assert isinstance(manifest, BaseProjectionManifest)
    collapsed = set(manifest.retained_to_collapsed_sources["Sheet1!A1"])
    assert collapsed == {"Sheet1!B1", "Sheet1!C1"}


def test_optimal_projection_manifest_round_trips() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    for n in (d, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    manifest = OptimalCompression().project(graph).manifest
    assert isinstance(manifest, BaseProjectionManifest)
    restored = resolve_projection_manifest(manifest.to_dict())
    assert isinstance(restored, BaseProjectionManifest)
    assert restored.kind == "optimal_compression"
    assert restored.forwarding_map == manifest.forwarding_map
    assert restored.retained_to_collapsed_sources == manifest.retained_to_collapsed_sources
