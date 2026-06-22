"""Unit tests for compressible candidate enumeration (issue #282 phase A)."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.compression import CompressionProvenanceRequiredError
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.similarity_compression import (
    CompressibleCandidate,
    SimilarityCompressionConfig,
    enumerate_compressible_candidates,
)
from tests.fixtures.tiny_dsa.workbook import (
    TINY_DSA_GROUPS,
    TINY_DSA_TARGETS,
    build_tiny_dsa_workbook,
)
from tests.unit.grapher.similarity_compression.conftest import direct_edge, make_node


def _candidate_by_root(
    candidates: tuple[CompressibleCandidate, ...],
) -> dict[str, CompressibleCandidate]:
    return {candidate.root: candidate for candidate in candidates}


def test_chain_yields_single_candidate() -> None:
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

    candidates = enumerate_compressible_candidates(graph)
    by_root = _candidate_by_root(candidates)
    candidate = by_root["Sheet1!A1"]
    assert candidate.members == frozenset({"Sheet1!A1", "Sheet1!B1", "Sheet1!C1"})
    assert candidate.size_reduction == 2


def test_multi_dependent_node_blocks_candidate() -> None:
    graph = DependencyGraph()
    leaf = make_node("Sheet1!L1", None, None, is_leaf=True)
    object.__setattr__(leaf, "value", 1)
    middle = make_node("Sheet1!M1", "=Sheet1!L1+1", "=Sheet1!L1+1")
    a = make_node("Sheet1!A1", "=Sheet1!M1+1", "=Sheet1!M1+1", is_target=True)
    b = make_node("Sheet1!B1", "=Sheet1!M1+2", "=Sheet1!M1+2", is_target=True)
    for node in (leaf, middle, a, b):
        graph.add_node(node)
    direct_edge(graph, "Sheet1!M1", "Sheet1!L1")
    direct_edge(graph, "Sheet1!A1", "Sheet1!M1")
    direct_edge(graph, "Sheet1!B1", "Sheet1!M1")

    candidates = enumerate_compressible_candidates(graph)
    assert candidates == ()


def test_fan_in_diamond_yields_root_candidate_without_leaf() -> None:
    graph = DependencyGraph()
    leaf = make_node("Sheet1!E1", None, None, is_leaf=True)
    b = make_node("Sheet1!B1", "=Sheet1!E1", "=Sheet1!E1")
    c = make_node("Sheet1!C1", "=Sheet1!E1", "=Sheet1!E1")
    d = make_node("Sheet1!D1", "=Sheet1!B1+Sheet1!C1", "=Sheet1!B1+Sheet1!C1")
    a = make_node("Sheet1!A1", "=Sheet1!D1", "=Sheet1!D1", is_target=True)
    for node in (leaf, b, c, d, a):
        graph.add_node(node)
    direct_edge(graph, "Sheet1!B1", "Sheet1!E1")
    direct_edge(graph, "Sheet1!C1", "Sheet1!E1")
    direct_edge(graph, "Sheet1!D1", "Sheet1!B1")
    direct_edge(graph, "Sheet1!D1", "Sheet1!C1")
    direct_edge(graph, "Sheet1!A1", "Sheet1!D1")

    candidates = enumerate_compressible_candidates(graph)
    by_root = _candidate_by_root(candidates)
    assert by_root["Sheet1!A1"].members == frozenset(
        {"Sheet1!A1", "Sheet1!B1", "Sheet1!C1", "Sheet1!D1"}
    )
    assert "Sheet1!E1" not in by_root["Sheet1!A1"].members


def test_issue_277_workbook_yields_c20_candidate(tmp_path: Path) -> None:
    from tests.unit.grapher.test_graph_optimal_compression import _build_issue_277_workbook

    path = tmp_path / "issue_277.xlsx"
    _build_issue_277_workbook(path)
    graph = create_dependency_graph(
        path,
        ["Engine!C20"],
        load_values=True,
        capture_dependency_provenance=True,
    )

    candidates = enumerate_compressible_candidates(graph)
    by_root = _candidate_by_root(candidates)
    candidate = by_root["Engine!C20"]
    assert candidate.members == frozenset(
        {
            "Engine!C20",
            "Engine!C14",
            "Engine!C15",
            "Engine!C16",
            "Engine!B20",
        }
    )
    assert candidate.size_reduction == 4


def test_tiny_dsa_yields_six_parallel_candidates(tmp_path: Path) -> None:
    path = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(path)
    graph = create_dependency_graph(
        path,
        list(TINY_DSA_TARGETS),
        load_values=True,
        capture_dependency_provenance=True,
    )

    candidates = enumerate_compressible_candidates(graph)
    by_root = _candidate_by_root(candidates)
    assert len(by_root) == 6
    for expected in TINY_DSA_GROUPS:
        candidate = by_root[expected.root]
        assert candidate.members == expected.members
        assert candidate.size_reduction == 3


def test_preserve_blocks_internal_node_from_group() -> None:
    graph = DependencyGraph()
    d = make_node("Sheet1!D1", None, None, is_leaf=True)
    object.__setattr__(d, "value", 2)
    b = make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2", is_target=True)
    a = make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1", is_target=True)
    for node in (d, b, a):
        graph.add_node(node)
    direct_edge(graph, "Sheet1!B1", "Sheet1!D1")
    direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    candidates = enumerate_compressible_candidates(graph)
    assert candidates == ()


def test_explicit_preserve_excludes_node_from_members() -> None:
    graph = DependencyGraph()
    d = make_node("Sheet1!D1", None, None, is_leaf=True)
    object.__setattr__(d, "value", 2)
    b = make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2")
    a = make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1", is_target=True)
    for node in (d, b, a):
        graph.add_node(node)
    direct_edge(graph, "Sheet1!B1", "Sheet1!D1")
    direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    candidates = enumerate_compressible_candidates(graph, preserve={"Sheet1!B1"})
    assert candidates == ()


def test_requires_provenance() -> None:
    graph = DependencyGraph()
    c = make_node("Sheet1!C1", None, None, is_leaf=True)
    b = make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for node in (c, b, a):
        graph.add_node(node)
    graph.add_edge("Sheet1!B1", "Sheet1!C1")
    graph.add_edge("Sheet1!A1", "Sheet1!B1")

    with pytest.raises(CompressionProvenanceRequiredError):
        enumerate_compressible_candidates(graph)


def test_candidate_cap_respected() -> None:
    graph = DependencyGraph()
    leaves = []
    for index in range(5):
        key = f"Sheet1!L{index}"
        leaf = make_node(key, None, None, is_leaf=True)
        object.__setattr__(leaf, "value", index)
        leaves.append(leaf)
        graph.add_node(leaf)

    prev = "Sheet1!L0"
    for index in range(1, 5):
        key = f"Sheet1!N{index}"
        node = make_node(key, f"={prev}+1", f"={prev}+1")
        graph.add_node(node)
        direct_edge(graph, key, prev)
        prev = key

    root = "Sheet1!R1"
    graph.add_node(make_node(root, f"={prev}+1", f"={prev}+1", is_target=True))
    direct_edge(graph, root, prev)

    config = SimilarityCompressionConfig(max_candidates=2)
    candidates = enumerate_compressible_candidates(graph, config=config)
    assert len(candidates) == 2
