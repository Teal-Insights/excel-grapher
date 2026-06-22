"""Unit tests for canonical embedding text (issue #282 phase D)."""

from __future__ import annotations

from pathlib import Path

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.similarity_compression import (
    canonicalize_for_embedding,
    enumerate_compressible_candidates,
    simulate_packing,
)
from excel_grapher.grapher.similarity_compression.packings import Packing, enumerate_packings
from tests.fixtures.tiny_dsa.workbook import SHOCKED_YEAR_COLUMNS, build_tiny_dsa_workbook


def test_shocked_year_blocks_share_role_normalized_formula(tmp_path: Path) -> None:
    path = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(path)
    graph = create_dependency_graph(
        path,
        [f"Engine!{col}20" for col in SHOCKED_YEAR_COLUMNS],
        load_values=True,
        capture_dependency_provenance=True,
    )
    candidates = enumerate_compressible_candidates(graph)
    packing = enumerate_packings(candidates)[0]
    simulation = simulate_packing(graph, packing)

    canonical = {
        root: canonicalize_for_embedding(root, formula, graph)
        for root, formula in simulation.collapsed_roots.items()
        if root.startswith("Engine!") and root.endswith("20") and root != "Engine!H20"
    }
    assert len(canonical) == 5
    blobs = list(canonical.values())
    assert blobs[0] == blobs[1] == blobs[2] == blobs[3] == blobs[4]
    assert "kind: shocked_year_block" in blobs[0]
    assert "{COL}" in blobs[0]
    assert "{BASE}" in blobs[0]


def test_linear_aggregate_kind_is_distinct(tmp_path: Path) -> None:
    path = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(path)
    graph = create_dependency_graph(
        path,
        ["Engine!H20"],
        load_values=True,
        capture_dependency_provenance=True,
    )
    candidate = enumerate_compressible_candidates(graph)[0]
    simulation = simulate_packing(graph, Packing(groups=(candidate,)))
    blob = canonicalize_for_embedding(
        "Engine!H20",
        simulation.collapsed_roots["Engine!H20"],
        graph,
    )
    assert "kind: linear_aggregate" in blob
