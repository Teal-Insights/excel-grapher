"""Unit tests for non-overlapping packing enumeration (issue #282 phase B)."""

from __future__ import annotations

from pathlib import Path

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.similarity_compression import (
    CompressibleCandidate,
    SimilarityCompressionConfig,
    enumerate_compressible_candidates,
    enumerate_packings,
)
from tests.fixtures.tiny_dsa.workbook import (
    TINY_DSA_GROUPS,
    TINY_DSA_TARGETS,
    build_tiny_dsa_workbook,
)


def _make_candidate(root: str, internals: tuple[str, ...]) -> CompressibleCandidate:
    return CompressibleCandidate(root=root, members=frozenset({root, *internals}))


def test_overlapping_candidates_pick_max_non_overlapping() -> None:
    c1 = _make_candidate("Sheet1!A1", ("Sheet1!B1", "Sheet1!C1", "Sheet1!D1"))
    c2 = _make_candidate("Sheet1!A2", ("Sheet1!C1", "Sheet1!E1"))
    c3 = _make_candidate("Sheet1!F1", ("Sheet1!G1", "Sheet1!H1", "Sheet1!I1"))

    packings = enumerate_packings((c1, c2, c3))
    best = packings[0]
    assert best.total_reduction == 6
    assert len(best.groups) == 2
    assert {group.root for group in best.groups} == {"Sheet1!A1", "Sheet1!F1"}


def test_parallel_synthetic_packing_includes_all_families() -> None:
    shocked = [
        _make_candidate(
            f"Engine!{col}20",
            (f"Engine!{col}14", f"Engine!{col}15", f"Engine!{col}16"),
        )
        for col in ("C", "D", "E", "F", "G")
    ]
    linear = _make_candidate("Engine!H20", ("Engine!H14", "Engine!H15", "Engine!H16"))
    candidates = tuple([*shocked, linear])

    packings = enumerate_packings(candidates)
    best = packings[0]
    assert best.total_reduction == 18
    assert len(best.groups) == 6
    assert best.member_nodes == frozenset(
        member for group in TINY_DSA_GROUPS for member in group.members
    )


def test_packing_cap_respected() -> None:
    candidates = tuple(
        _make_candidate(f"Sheet1!R{index}", (f"Sheet1!I{index}",)) for index in range(10)
    )
    config = SimilarityCompressionConfig(top_n_packings=3)
    packings = enumerate_packings(candidates, config=config)
    assert len(packings) == 3
    assert packings[0].total_reduction == 10


def test_tiny_dsa_top_packing_covers_all_groups(tmp_path: Path) -> None:
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
    best = packings[0]
    assert best.total_reduction == 18
    assert {group.root for group in best.groups} == {g.root for g in TINY_DSA_GROUPS}


def test_empty_candidates_yields_empty_packings() -> None:
    assert enumerate_packings(()) == ()
