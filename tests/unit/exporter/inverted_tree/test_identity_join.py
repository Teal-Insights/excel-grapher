"""Identity alignment is the inverse of `schedule_coord` (#607 / #608 / §3)."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.exporter.inverted_tree.deps import (
    collect_all_dependence_edges,
    identity_join_indices,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.schedule import schedule_coord
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    inverted_graph_parts,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a17_overlap_take import (
    _overlap_bindings,
    _overlap_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a18_splice import (
    _splice_bindings,
    _splice_workbook,
)


def _time_series(series_id: str, data_range: str, *, internal: bool = False) -> dict:
    return series_entry(
        series_id,
        data_range,
        layout="series",
        direction="internal" if internal else "input",
        header_row=1,
    )


def test_overlap_join_inverts_schedule_coord(tmp_path: Path) -> None:
    catalog, deps, graph = inverted_graph_parts(_overlap_workbook(tmp_path), _overlap_bindings())
    host = catalog.get("revenue_pct_gdp")
    gdp = catalog.get("gdp")
    joined = identity_join_indices(host, gdp, catalog)
    assert joined == (1, 2)
    assert joined == deps["revenue_pct_gdp"].index_maps["gdp"]
    for host_cell, slot in zip(host.cells, joined, strict=True):
        assert schedule_coord(host_cell, catalog) == schedule_coord(gdp.cells[slot], catalog)
    edges = collect_all_dependence_edges(catalog, graph)
    gdp_access = {
        edge.access
        for edge in edges
        if edge.consumer_id == "revenue_pct_gdp" and edge.producer_id == "gdp"
    }
    assert gdp_access == {"identity"}


def test_splice_prefix_is_identity_join_to_last_growth_year(tmp_path: Path) -> None:
    catalog, deps, graph = inverted_graph_parts(_splice_workbook(tmp_path), _splice_bindings())
    path = catalog.get("path")
    growth = catalog.get("growth")
    trajectory = catalog.get("trajectory")
    assert identity_join_indices(path, growth, catalog) == (1, -1, -1, -1)
    assert identity_join_indices(path, trajectory, catalog) == (-1, 0, 1, 2)
    assert "growth" not in deps["path"].aligned_ids
    assert "trajectory" not in deps["path"].aligned_ids
    edges = collect_all_dependence_edges(catalog, graph)
    prefix = next(
        edge for edge in edges if edge.consumer_cell == "Engine!D4" and edge.producer_id == "growth"
    )
    assert prefix.access == "identity"
    assert prefix.producer_cell == "Engine!D3"


def test_duplicate_time_period_identity_join_fails_closed(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "dup_year.xlsx",
        {
            "Engine": {
                "B1": 2010,
                "C1": 2010,
                "D1": 2011,
                "B2": 100,
                "C2": 110,
                "D2": 121,
                "C3": 10,
                "D3": 12,
                "C4": "=C3/C2",
                "D4": "=D3/D2",
            },
        },
    )
    bindings = bindings_document(
        _time_series("gdp", "Engine!B2:D2"),
        _time_series("revenue", "Engine!C3:D3"),
        series_entry(
            "ratio",
            "Engine!C4:D4",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )
    with pytest.raises(InvertedTreeExportError, match="duplicate|join"):
        inverted_graph_parts(workbook, bindings)
