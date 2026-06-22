"""Integration tests for similarity-aware selection (issue #282)."""

from __future__ import annotations

from pathlib import Path

from excel_grapher import create_dependency_graph
from excel_grapher.exporter import SimilarityAwareCompression
from excel_grapher.exporter.projection import BaseProjectionManifest
from excel_grapher.grapher.similarity_compression import (
    MockEmbeddingProvider,
    select_similarity_projection,
)
from tests.fixtures.tiny_dsa.workbook import (
    TINY_DSA_GROUPS,
    TINY_DSA_TARGETS,
    build_tiny_dsa_workbook,
)


def test_select_similarity_projection_on_tiny_dsa(tmp_path: Path) -> None:
    path = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(path)
    graph = create_dependency_graph(
        path,
        list(TINY_DSA_TARGETS),
        load_values=True,
        capture_dependency_provenance=True,
    )

    selection = select_similarity_projection(graph, provider=MockEmbeddingProvider())
    assert selection.score.total_reduction == 18
    assert {group.root for group in selection.packing.groups} == {
        group.root for group in TINY_DSA_GROUPS
    }
    shocked_roots = [root for root in selection.simulation.collapsed_roots if root.endswith("20")]
    assert len(shocked_roots) == 6


def test_similarity_aware_compression_projection_step(tmp_path: Path) -> None:
    path = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(path)
    graph = create_dependency_graph(
        path,
        list(TINY_DSA_TARGETS),
        load_values=True,
        capture_dependency_provenance=True,
    )

    projection = SimilarityAwareCompression(provider=MockEmbeddingProvider()).project(graph)
    assert isinstance(projection.manifest, BaseProjectionManifest)
    assert projection.manifest.kind == "similarity_aware_compression"
    removed = {key for key in graph if key not in projection.projected_graph}
    assert len(removed) == 18
