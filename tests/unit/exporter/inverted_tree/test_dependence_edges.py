"""DependenceEdge is the source of truth; SeriesDeps is a derived view."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.exporter.inverted_tree.deps import (
    collect_all_dependence_edges,
    collect_all_deps,
    series_deps_from_edges,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.schedule import (
    DependenceEdge,
    plan_fused_scc,
    residual_body_order,
)
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    inverted_graph_parts,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a1_leaf_closure import (
    _a1_bindings,
    _a1_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a2_first_level_deps import (
    _a2_bindings,
    _a2_dynamic_refs,
    _a2_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a10_other_series_lag import (
    _lag_bindings,
    _lag_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a11_zipper import (
    _simultaneous_workbook,
    _zipper_bindings,
)


def _accesses(
    edges: tuple[DependenceEdge, ...],
    consumer_id: str,
    producer_id: str,
) -> set[str]:
    return {
        edge.access
        for edge in edges
        if edge.consumer_id == consumer_id and edge.producer_id == producer_id
    }


def test_a10_lag_edges_are_identity_and_shift(tmp_path: Path) -> None:
    catalog, deps, graph = inverted_graph_parts(_lag_workbook(tmp_path), _lag_bindings())
    edges = collect_all_dependence_edges(catalog, graph)
    assert _accesses(edges, "direction", "debt") == {"identity", "shift"}
    direction = catalog.get("direction")
    derived = series_deps_from_edges(direction, edges, catalog)
    assert derived.lagged_ids == frozenset({"debt"})
    assert derived.aligned_ids == frozenset()
    assert derived.lookup_ids == frozenset()
    assert derived == deps["direction"]


def test_a2_offset_table_is_dynamic_or_whole(tmp_path: Path) -> None:
    catalog, deps, graph = inverted_graph_parts(
        _a2_workbook(tmp_path),
        _a2_bindings(),
        dynamic_refs=_a2_dynamic_refs(),
    )
    edges = collect_all_dependence_edges(catalog, graph)
    assert _accesses(edges, "shock_magnitude_resolved", "shock_magnitudes") <= {
        "dynamic",
        "whole",
    }
    assert _accesses(edges, "shock_magnitude_resolved", "shock_magnitudes")
    derived = series_deps_from_edges(catalog.get("shock_magnitude_resolved"), edges, catalog)
    assert derived.lookup_ids == frozenset({"shock_magnitudes"})
    assert derived == deps["shock_magnitude_resolved"]


def test_a1_aligned_path_is_identity(tmp_path: Path) -> None:
    catalog, deps, graph = inverted_graph_parts(_a1_workbook(tmp_path), _a1_bindings())
    edges = collect_all_dependence_edges(catalog, graph)
    assert _accesses(edges, "engine_path", "growth") == {"identity"}
    assert _accesses(edges, "engine_path", "interest") == {"identity"}
    assert _accesses(edges, "engine_path", "engine_path") == {"shift"}
    derived = series_deps_from_edges(catalog.get("engine_path"), edges, catalog)
    assert derived.aligned_ids == frozenset({"growth", "interest"})
    assert derived.index_maps == {"interest": (0, 1), "growth": (0, 1)}
    assert derived.is_scan is True
    assert derived.seed_id == "engine_year0"
    assert derived == deps["engine_path"]


def test_derived_series_deps_match_collect_all_deps(tmp_path: Path) -> None:
    catalog, deps, graph = inverted_graph_parts(_a1_workbook(tmp_path), _a1_bindings())
    edges = collect_all_dependence_edges(catalog, graph)
    derived = {
        series.series_id: series_deps_from_edges(series, edges, catalog)
        for series in catalog.formula_series()
    }
    assert derived == deps
    assert collect_all_deps(catalog, graph) == deps


def test_distance_zero_cycle_names_statements_and_index() -> None:
    edges = (
        DependenceEdge("debt", "adjustment", "Engine!B2", "Engine!B3", 0),
        DependenceEdge("adjustment", "debt", "Engine!B3", "Engine!B2", 0),
    )
    with pytest.raises(InvertedTreeExportError, match="distance-zero residual") as exc:
        residual_body_order(("debt", "adjustment"), edges)
    message = str(exc.value)
    assert "debt" in message
    assert "adjustment" in message
    assert "Engine!B2" in message
    assert "Engine!B3" in message
    assert "index 2" in message


def test_simultaneous_workbook_residual_names_cells(tmp_path: Path) -> None:
    catalog, _deps, graph = inverted_graph_parts(
        _simultaneous_workbook(tmp_path), _zipper_bindings()
    )
    with pytest.raises(InvertedTreeExportError, match="distance-zero residual") as exc:
        plan_fused_scc(("debt", "adjustment"), catalog=catalog, graph=graph)
    message = str(exc.value)
    assert "debt" in message
    assert "adjustment" in message
    assert "Engine!B2" in message or "Engine!C2" in message
    assert "index 1" in message or "index 2" in message


def test_collect_dependence_edges_records_guarded_flag(tmp_path: Path) -> None:
    wb = write_workbook(
        tmp_path / "guarded_edges.xlsx",
        {
            "Engine": {
                "A1": 0,
                "B1": "=IF($A$1=1, C1, 10)",
                "C1": "=B1*2",
            },
        },
    )
    bindings = bindings_document(
        series_entry("flag", "Engine!A1", layout="scalar", direction="input"),
        series_entry("b", "Engine!B1", layout="scalar", direction="output"),
        series_entry("c", "Engine!C1", layout="scalar", direction="output"),
    )
    catalog, _deps, graph = inverted_graph_parts(wb, bindings)
    edges = collect_all_dependence_edges(catalog, graph)
    b_to_c = next(e for e in edges if e.consumer_id == "b" and e.producer_id == "c")
    c_to_b = next(e for e in edges if e.consumer_id == "c" and e.producer_id == "b")
    assert b_to_c.guarded is True
    assert c_to_b.guarded is False
