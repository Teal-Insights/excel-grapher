"""Full parity on cross_sheet_taco_patterns.xlsx."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.range_compression import PatternKind, build_taco_index
from tests.unit.grapher.range_compression.parity_helpers import assert_taco_parity

FIXTURE = (
    Path(__file__).resolve().parents[3]
    / "examples"
    / "micro_workbooks"
    / "cross_sheet_taco_patterns.xlsx"
)


@pytest.mark.skipif(not FIXTURE.is_file(), reason="cross_sheet_taco_patterns.xlsx missing")
def test_cross_sheet_taco_full_parity() -> None:
    graph = create_dependency_graph(
        FIXTURE,
        ["Report!D3:D7", "Report!F3:F7", "Report!H3:H7", "Report!K3:K7"],
        load_values=False,
        store_raw_formula=True,
    )
    index = build_taco_index(graph)
    kinds = {e.meta.kind for e in index.compressed_edges}
    assert PatternKind.rr in kinds
    assert PatternKind.rf in kinds
    assert PatternKind.fr in kinds
    assert PatternKind.ff in kinds
    cross = [e for e in index.compressed_edges if e.precedent.sheet != e.dependent.sheet]
    assert len(cross) >= 4
    assert_taco_parity(graph, index)
