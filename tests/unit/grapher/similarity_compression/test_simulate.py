"""Unit tests for packing collapse simulation (issue #282 phase C)."""

from __future__ import annotations

from pathlib import Path

from excel_grapher import create_dependency_graph
from excel_grapher.evaluator.parser import parse
from excel_grapher.exporter import OptimalCompression
from excel_grapher.exporter.projection import (
    BaseProjectionManifest,
    SimilarityPackingProjection,
    build_similarity_projection_manifest,
    project_similarity_packing,
)
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.similarity_compression import (
    enumerate_compressible_candidates,
    enumerate_packings,
    simulate_packing,
)
from tests.fixtures.tiny_dsa.workbook import (
    TINY_DSA_GROUPS,
    TINY_DSA_TARGETS,
    build_tiny_dsa_workbook,
)
from tests.unit.grapher.similarity_compression.conftest import direct_edge, make_node


def test_simulate_does_not_mutate_canonical_graph(tmp_path: Path) -> None:
    path = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(path)
    graph = create_dependency_graph(
        path,
        list(TINY_DSA_TARGETS),
        load_values=True,
        capture_dependency_provenance=True,
    )
    original_keys = set(graph)
    candidates = enumerate_compressible_candidates(graph)
    packing = enumerate_packings(candidates)[0]

    simulate_packing(graph, packing)

    assert set(graph) == original_keys


def test_issue_277_single_group_fully_collapses_candidate(tmp_path: Path) -> None:
    from tests.unit.grapher.test_graph_optimal_compression import _build_issue_277_workbook

    path = tmp_path / "issue_277.xlsx"
    _build_issue_277_workbook(path)
    graph = create_dependency_graph(
        path,
        ["Engine!C20"],
        load_values=True,
        capture_dependency_provenance=True,
    )
    candidate = enumerate_compressible_candidates(graph)[0]
    packing = enumerate_packings((candidate,))[0]

    optimal = OptimalCompression().project(graph)
    simulation = simulate_packing(graph, packing)

    optimal_formula = optimal.projected_graph.get_node("Engine!C20")
    simulated_formula = simulation.projected_graph.get_node("Engine!C20")
    assert optimal_formula is not None and simulated_formula is not None
    assert "Engine!C16" in optimal.projected_graph
    assert "Engine!C16" not in simulation.projected_graph
    assert simulated_formula.normalized_formula is not None
    parse(simulated_formula.normalized_formula.strip())
    assert "/100" in simulated_formula.normalized_formula
    assert "(Inputs!C16+CHOOSE" in simulated_formula.normalized_formula


def test_tiny_dsa_packing_removes_all_group_internals(tmp_path: Path) -> None:
    path = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(path)
    graph = create_dependency_graph(
        path,
        list(TINY_DSA_TARGETS),
        load_values=True,
        capture_dependency_provenance=True,
    )
    candidates = enumerate_compressible_candidates(graph)
    packing = enumerate_packings(candidates)[0]
    simulation = simulate_packing(graph, packing)

    removed = {key for key in graph if key not in simulation.projected_graph}
    expected_removed = {member for group in TINY_DSA_GROUPS for member in group.internal_members}
    assert removed == expected_removed
    assert len(removed) == 18
    for group in TINY_DSA_GROUPS:
        root = simulation.projected_graph.get_node(group.root)
        assert root is not None
        assert root.normalized_formula is not None
        assert simulation.collapsed_roots[group.root] == root.normalized_formula
        parse(root.normalized_formula.strip())


def test_projection_result_carries_similarity_manifest(tmp_path: Path) -> None:
    path = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(path)
    graph = create_dependency_graph(
        path,
        list(TINY_DSA_TARGETS),
        load_values=True,
        capture_dependency_provenance=True,
    )
    packing = enumerate_packings(enumerate_compressible_candidates(graph))[0]

    projection = SimilarityPackingProjection(packing).project(graph)
    assert projection.original_graph is graph
    assert isinstance(projection.manifest, BaseProjectionManifest)
    assert projection.manifest.kind == "similarity_aware_compression"
    shocked_roots = {f"Engine!{column}20" for column in "CDEFG"}
    shocked_retained = {
        group.retained
        for group in projection.manifest.collapsed_groups
        if group.retained in shocked_roots
    }
    assert shocked_retained == shocked_roots
    assert "Engine!H20" in projection.projected_graph
    assert {"Engine!H14", "Engine!H15", "Engine!H16"} <= projection.manifest.forwarding_map.keys()


def test_project_similarity_packing_matches_simulate(tmp_path: Path) -> None:
    path = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(path)
    graph = create_dependency_graph(
        path,
        ["Engine!C20"],
        load_values=True,
        capture_dependency_provenance=True,
    )
    packing = enumerate_packings(enumerate_compressible_candidates(graph))[0]
    simulation = simulate_packing(graph, packing)
    projected, manifest = project_similarity_packing(graph, packing)

    assert set(projected) == set(simulation.projected_graph)
    assert manifest == build_similarity_projection_manifest(graph, simulation.record)


def test_chain_manual_graph_collapses_to_root() -> None:
    graph = DependencyGraph()
    d = make_node("Sheet1!D1", None, None, is_leaf=True)
    object.__setattr__(d, "value", 2)
    c = make_node("Sheet1!C1", "=Sheet1!D1+1", "=Sheet1!D1+1")
    b = make_node("Sheet1!B1", "=Sheet1!C1*2", "=Sheet1!C1*2")
    a = make_node("Sheet1!A1", "=Sheet1!B1+3", "=Sheet1!B1+3", is_target=True)
    for node in (d, c, b, a):
        graph.add_node(node)
    direct_edge(graph, "Sheet1!C1", "Sheet1!D1")
    direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    candidate = enumerate_compressible_candidates(graph)[0]
    simulation = simulate_packing(graph, enumerate_packings((candidate,))[0])

    assert "Sheet1!B1" not in simulation.projected_graph
    assert "Sheet1!C1" not in simulation.projected_graph
    assert "Sheet1!D1" in simulation.projected_graph
    root = simulation.projected_graph.get_node("Sheet1!A1")
    assert root is not None
    assert "Sheet1!D1" in (root.normalized_formula or "")
