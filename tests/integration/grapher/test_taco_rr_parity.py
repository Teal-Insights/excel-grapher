"""Integration parity tests for TACO index on the taco_patterns fixture (RR only in PR1)."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.range_compression import build_taco_index
from tests.unit.grapher.range_compression.parity_helpers import assert_taco_parity

FIXTURE = (
    Path(__file__).resolve().parents[3] / "examples" / "micro_workbooks" / "taco_patterns.xlsx"
)


@pytest.mark.skipif(not FIXTURE.is_file(), reason="taco_patterns.xlsx fixture missing")
def test_taco_rr_parity_on_fixture_workbook() -> None:
    graph = create_dependency_graph(
        FIXTURE,
        ["Patterns!D3:D7"],
        load_values=False,
        capture_dependency_provenance=True,
    )
    index = build_taco_index(graph)
    assert_taco_parity(graph, index)
    assert len(index.compressed_edges) > 0
