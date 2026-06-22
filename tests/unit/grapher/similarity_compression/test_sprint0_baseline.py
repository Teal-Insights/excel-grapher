"""Sprint 0 baseline for similarity-aware compression (issue #282).

Records Tiny DSA fixture expectations and ``OptimalCompression`` reference
metrics that later sprints must match or beat on size while improving
parallel-family selection.

Optimal-compression baseline (greedy, as of Sprint 0):

- **14 nodes removed**: shared identity transit ``Engine!B20``, rows 14–15 for
  columns C–G, and all three internals for group 6 (H14–H16).
- **Row 16 survives** for shocked-year columns (C16–G16): incoming-edge
  substitution is blocked on the ``-{col}16`` term in each root formula.
- **Six roots survive**: ``Engine!C20`` … ``Engine!G20`` and ``Engine!H20``.

Similarity-aware compression should eventually collapse full candidate groups
(including row 16) when selected together in a packing, while keeping total
reduction within ~5% of this baseline.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher import create_dependency_graph
from excel_grapher.exporter import OptimalCompression
from excel_grapher.exporter.projection import BaseProjectionManifest
from excel_grapher.grapher.similarity_compression import SimilarityCompressionConfig
from tests.fixtures.tiny_dsa.workbook import (
    GROUP_6_ROOT,
    LINEAR_FAMILY_GROUPS,
    SHOCKED_YEAR_FAMILY_GROUPS,
    TINY_DSA_GROUPS,
    TINY_DSA_TARGETS,
    TinyDsaGroup,
    build_tiny_dsa_workbook,
)

OPTIMAL_BASELINE_REMOVED_COUNT = 14
OPTIMAL_BASELINE_REMOVED = frozenset(
    {
        "Engine!B20",
        "Engine!C14",
        "Engine!C15",
        "Engine!D14",
        "Engine!D15",
        "Engine!E14",
        "Engine!E15",
        "Engine!F14",
        "Engine!F15",
        "Engine!G14",
        "Engine!G15",
        "Engine!H14",
        "Engine!H15",
        "Engine!H16",
    }
)
OPTIMAL_BASELINE_SURVIVING_INTERNALS = frozenset(
    {
        "Engine!C16",
        "Engine!D16",
        "Engine!E16",
        "Engine!F16",
        "Engine!G16",
    }
)
OPTIMAL_BASELINE_SURVIVING_ROOTS = frozenset(group.root for group in TINY_DSA_GROUPS)
OPTIMAL_BASELINE_COLLAPSED_GROUP_COUNT = 9


def test_similarity_compression_config_defaults() -> None:
    config = SimilarityCompressionConfig()
    assert config.max_candidates == 200
    assert config.top_n_packings == 50
    assert config.require_connected_component is True
    assert config.alpha == 0.4
    assert config.beta == 0.6
    assert config.gamma == 0.05
    assert config.fallback_to_optimal is True
    assert config.embedding_model == "text-embedding-3-small"


def test_tiny_dsa_workbook_builds(tmp_path: Path) -> None:
    path = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(path)
    assert path.is_file()
    assert path.stat().st_size > 0


def test_tiny_dsa_fixture_paths(
    tiny_dsa_workbook_path: Path, tiny_dsa_targets: tuple[str, ...]
) -> None:
    assert tiny_dsa_workbook_path.is_file()
    assert len(tiny_dsa_targets) == 6


def test_tiny_dsa_groups_encode_parallel_families(
    tiny_dsa_groups: tuple[TinyDsaGroup, ...],
) -> None:
    assert len(tiny_dsa_groups) == 6
    assert len(SHOCKED_YEAR_FAMILY_GROUPS) == 5
    assert len(LINEAR_FAMILY_GROUPS) == 1
    assert LINEAR_FAMILY_GROUPS[0].root == GROUP_6_ROOT
    shocked_roots = {g.root for g in SHOCKED_YEAR_FAMILY_GROUPS}
    assert shocked_roots == {
        "Engine!C20",
        "Engine!D20",
        "Engine!E20",
        "Engine!F20",
        "Engine!G20",
    }
    for group in tiny_dsa_groups:
        assert group.root in group.members
        assert len(group.internal_members) == 3


@pytest.mark.parametrize("group", TINY_DSA_GROUPS, ids=lambda g: f"group_{g.group_id}")
def test_tiny_dsa_group_metadata(group: TinyDsaGroup) -> None:
    assert group.group_id in range(1, 7)
    if group.group_id <= 5:
        assert group.parallel_family == "shocked_year_block"
        assert group.root.endswith("20")
        col = group.root.split("!")[1][0]
        assert {m.split("!")[1][0] for m in group.members} == {col}
    else:
        assert group.parallel_family == "linear_aggregate"
        assert group.root == GROUP_6_ROOT


def test_tiny_dsa_optimal_compression_baseline(tmp_path: Path) -> None:
    """Document OptimalCompression removal counts for the Tiny DSA fixture."""
    path = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(path)

    graph = create_dependency_graph(
        path,
        list(TINY_DSA_TARGETS),
        load_values=True,
        capture_dependency_provenance=True,
    )
    projection = OptimalCompression().project(graph)
    projected = projection.projected_graph
    manifest = projection.manifest

    removed = {key for key in graph if key not in projected}

    assert len(removed) == OPTIMAL_BASELINE_REMOVED_COUNT
    assert removed == OPTIMAL_BASELINE_REMOVED
    assert OPTIMAL_BASELINE_SURVIVING_INTERNALS.issubset(projected)
    for root in OPTIMAL_BASELINE_SURVIVING_ROOTS:
        assert root in projected
        node = projected.get_node(root)
        assert node is not None
        assert node.normalized_formula is not None

    assert isinstance(manifest, BaseProjectionManifest)
    assert len(manifest.collapsed_groups) == OPTIMAL_BASELINE_COLLAPSED_GROUP_COUNT
