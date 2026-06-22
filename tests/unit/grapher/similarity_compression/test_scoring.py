"""Unit tests for embedding and packing scores (issue #282 phases D–E)."""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.similarity_compression import (
    MockEmbeddingProvider,
    Packing,
    SimilarityCompressionConfig,
    cosine_distance,
    embed_texts,
    enumerate_compressible_candidates,
    enumerate_packings,
    score_packing,
    score_packings,
    select_best_packing,
)
from excel_grapher.grapher.similarity_compression.candidates import CompressibleCandidate
from tests.fixtures.tiny_dsa.workbook import (
    TINY_DSA_GROUPS,
    TINY_DSA_TARGETS,
    build_tiny_dsa_workbook,
)


def test_cosine_distance_identical_vectors_are_zero() -> None:
    vector = [1.0, 0.0, 0.0]
    assert cosine_distance(vector, vector) == 0.0


def test_mock_embeddings_are_deterministic() -> None:
    provider = MockEmbeddingProvider()
    first = provider.embed(["kind: shocked_year_block\nformula_normalized: x"])
    second = provider.embed(["kind: shocked_year_block\nformula_normalized: x"])
    assert first == second


def test_score_prefers_tight_parallel_cluster() -> None:
    packing_parallel = Packing(
        groups=(
            _candidate("Engine!C20", ("Engine!C14", "Engine!C15")),
            _candidate("Engine!D20", ("Engine!D14", "Engine!D15")),
        )
    )
    packing_single = Packing(groups=(_candidate("Engine!C20", ("Engine!C14", "Engine!C15")),))
    collapsed_parallel = {
        "Engine!C20": "=CHOOSE(1)*1/100",
        "Engine!D20": "=CHOOSE(1)*1/100",
    }
    collapsed_single = {"Engine!C20": "=CHOOSE(1)*1/100"}
    provider = MockEmbeddingProvider()
    embeddings_parallel = {
        root: embed_texts([formula], provider)[formula]
        for root, formula in collapsed_parallel.items()
    }
    embeddings_single = {
        root: embed_texts([formula], provider)[formula]
        for root, formula in collapsed_single.items()
    }
    graph = _empty_graph_stub()
    parallel_score = score_packing(
        packing_parallel,
        collapsed_parallel,
        graph,
        embeddings_parallel,
    )
    single_score = score_packing(
        packing_single,
        collapsed_single,
        graph,
        embeddings_single,
    )
    assert parallel_score.total_reduction > single_score.total_reduction
    assert parallel_score.final_score > single_score.final_score


def test_tiny_dsa_select_best_packing_is_full_parallel_family(tmp_path: Path) -> None:
    path = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(path)
    graph = create_dependency_graph(
        path,
        list(TINY_DSA_TARGETS),
        load_values=True,
        capture_dependency_provenance=True,
    )
    candidates = enumerate_compressible_candidates(graph)
    packings = enumerate_packings(candidates)
    score, simulation = select_best_packing(
        graph,
        packings,
        provider=MockEmbeddingProvider(),
    )
    assert score.packing.total_reduction == 18
    assert {group.root for group in score.packing.groups} == {
        group.root for group in TINY_DSA_GROUPS
    }
    assert score.mean_cluster_distance == 0.0
    assert len(simulation.collapsed_roots) == 6


def test_score_packings_shares_embedding_work(tmp_path: Path) -> None:
    path = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(path)
    graph = create_dependency_graph(
        path,
        list(TINY_DSA_TARGETS),
        load_values=True,
        capture_dependency_provenance=True,
    )
    provider = _CountingProvider()
    packings = enumerate_packings(enumerate_compressible_candidates(graph))[:3]
    score_packings(graph, packings, provider=provider)
    assert provider.call_count == 1


def test_select_best_packing_flat_scores_falls_back_to_max_reduction(tmp_path: Path) -> None:
    """When similarity scores tie, prefer the packing with greatest node reduction."""
    path = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(path)
    graph = create_dependency_graph(
        path,
        list(TINY_DSA_TARGETS),
        load_values=True,
        capture_dependency_provenance=True,
    )
    packings = enumerate_packings(enumerate_compressible_candidates(graph))
    assert len(packings) > 1
    reductions = {packing.total_reduction for packing in packings}
    assert len(reductions) > 1

    flat_config = SimilarityCompressionConfig(
        alpha=0.0,
        beta=0.0,
        gamma=0.0,
        fallback_to_optimal=True,
        score_flatness_epsilon=1.0,
    )
    score, _simulation = select_best_packing(
        graph,
        packings,
        provider=MockEmbeddingProvider(),
        config=flat_config,
    )
    assert score.total_reduction == max(reductions)
    assert score.total_reduction == 18


def test_select_best_packing_flat_scores_disabled_keeps_score_order(tmp_path: Path) -> None:
    path = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(path)
    graph = create_dependency_graph(
        path,
        list(TINY_DSA_TARGETS),
        load_values=True,
        capture_dependency_provenance=True,
    )
    packings = enumerate_packings(enumerate_compressible_candidates(graph))
    flat_config = SimilarityCompressionConfig(
        alpha=0.0,
        beta=0.0,
        gamma=0.0,
        fallback_to_optimal=False,
        score_flatness_epsilon=1.0,
    )
    with_fallback, _ = select_best_packing(
        graph,
        packings,
        provider=MockEmbeddingProvider(),
        config=replace(flat_config, fallback_to_optimal=True),
    )
    without_fallback, _ = select_best_packing(
        graph,
        packings,
        provider=MockEmbeddingProvider(),
        config=flat_config,
    )
    assert with_fallback.total_reduction == 18
    assert without_fallback.final_score == with_fallback.final_score


def _candidate(root: str, internals: tuple[str, ...]) -> CompressibleCandidate:
    return CompressibleCandidate(root=root, members=frozenset({root, *internals}))


def _empty_graph_stub():
    from excel_grapher.grapher.graph import DependencyGraph

    return DependencyGraph()


class _CountingProvider(MockEmbeddingProvider):
    def __init__(self) -> None:
        super().__init__()
        self.call_count = 0

    def embed(self, texts: list[str]) -> list[list[float]]:
        self.call_count += 1
        return super().embed(texts)
