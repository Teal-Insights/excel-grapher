"""Full TACO parity on taco_row_patterns fixture workbook."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.range_compression import PatternKind, build_taco_index
from tests.unit.grapher.range_compression.parity_helpers import assert_taco_parity

FIXTURE = (
    Path(__file__).resolve().parents[3] / "examples" / "micro_workbooks" / "taco_row_patterns.xlsx"
)


@pytest.mark.skipif(not FIXTURE.is_file(), reason="taco_row_patterns.xlsx fixture missing")
def test_taco_row_full_parity_on_fixture_workbook() -> None:
    graph = create_dependency_graph(
        FIXTURE,
        [
            "PatternsRow!F9:J9",
            "PatternsRow!W9:AA9",
            "PatternsRow!AI9:AM9",
            "PatternsRow!BA9:BE9",
            "PatternsRow!AS9:AW9",
        ],
        load_values=False,
        capture_dependency_provenance=True,
    )
    index = build_taco_index(graph)
    kinds = {edge.meta.kind for edge in index.compressed_edges}
    assert PatternKind.rr in kinds
    assert PatternKind.rf in kinds
    assert PatternKind.fr in kinds
    assert PatternKind.ff in kinds
    assert PatternKind.rr_chain in kinds
    row_edges = [e for e in index.compressed_edges if e.dependent.min_row == e.dependent.max_row]
    assert row_edges, "expected row-span compressed edges"
    assert_taco_parity(graph, index)
