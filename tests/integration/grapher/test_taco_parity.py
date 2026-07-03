"""Full TACO parity on taco_patterns fixture workbook."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.range_compression import PatternKind, build_taco_index
from tests.unit.grapher.range_compression.parity_helpers import assert_taco_parity

FIXTURE = (
    Path(__file__).resolve().parents[3] / "examples" / "micro_workbooks" / "taco_patterns.xlsx"
)


@pytest.mark.skipif(not FIXTURE.is_file(), reason="taco_patterns.xlsx fixture missing")
def test_taco_full_parity_on_fixture_workbook() -> None:
    graph = create_dependency_graph(
        FIXTURE,
        [
            "Patterns!D3:D7",
            "Patterns!F3:F7",
            "Patterns!H3:H7",
            "Patterns!K3:K7",
            "Patterns!P3:P7",
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
    assert_taco_parity(graph, index)
    assert len(index.single_edges) == 0
