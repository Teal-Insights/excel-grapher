"""Statement-graph legality and fused-SCC classification."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.exporter.inverted_tree.deps import collect_all_dependence_edges
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.schedule import (
    DependenceEdge,
    FusedRegion,
    plan_fused_scc,
    residual_body_order,
    schedule_coord,
)
from tests.unit.exporter.inverted_tree.helpers import inverted_graph_parts
from tests.unit.exporter.inverted_tree.test_shape_a11_zipper import (
    _cross_sheet_zipper_bindings,
    _cross_sheet_zipper_workbook,
    _offset_zipper_bindings,
    _offset_zipper_workbook,
    _simultaneous_workbook,
    _vertical_zipper_bindings,
    _vertical_zipper_workbook,
    _zipper_bindings,
    _zipper_workbook,
)


def test_lagged_zipper_residual_orders_adjustment_before_debt() -> None:
    edges = (
        DependenceEdge("adjustment", "debt", "Engine!B3", "Engine!A2", 1),
        DependenceEdge("debt", "adjustment", "Engine!B2", "Engine!B3", 0),
        DependenceEdge("debt", "debt", "Engine!B2", "Engine!A2", 1),
    )
    assert residual_body_order(("debt", "adjustment"), edges) == ("adjustment", "debt")


def test_distance_zero_cycle_is_illegal() -> None:
    edges = (
        DependenceEdge("debt", "adjustment", "Engine!B2", "Engine!B3", 0),
        DependenceEdge("adjustment", "debt", "Engine!B3", "Engine!B2", 0),
    )
    with pytest.raises(InvertedTreeExportError, match="distance-zero residual"):
        residual_body_order(("debt", "adjustment"), edges)


def test_identity_flip_has_no_single_body_order() -> None:
    edges = (
        DependenceEdge("x", "y", "Engine!A2", "Engine!A4", 0),
        DependenceEdge("y", "x", "Engine!B4", "Engine!B2", 0),
    )
    assert residual_body_order(("x", "y"), edges) is None


def test_plan_fused_scc_for_corrected_zipper(tmp_path: Path) -> None:
    catalog, _deps, graph = inverted_graph_parts(_zipper_workbook(tmp_path), _zipper_bindings())
    plan = plan_fused_scc(("debt", "adjustment"), catalog=catalog, graph=graph)
    assert plan is not None
    assert plan.body_order == ("adjustment", "debt")
    assert plan.schedule == (0, 1, 2)
    assert plan.domain["debt"] == (0, 3)
    assert plan.domain["adjustment"] == (1, 3)
    assert plan.peel_stop == 1
    assert plan.regions == (
        FusedRegion(start=0, stop=1, body_order=("debt",)),
        FusedRegion(start=1, stop=3, body_order=("adjustment", "debt")),
    )


def test_plan_fused_scc_rejects_same_year_cycle(tmp_path: Path) -> None:
    catalog, _deps, graph = inverted_graph_parts(
        _simultaneous_workbook(tmp_path), _zipper_bindings()
    )
    with pytest.raises(InvertedTreeExportError, match="distance-zero residual"):
        plan_fused_scc(("debt", "adjustment"), catalog=catalog, graph=graph)


def test_schedule_coord_joins_resolved_time_period(tmp_path: Path) -> None:
    catalog, _deps, _graph = inverted_graph_parts(_zipper_workbook(tmp_path), _zipper_bindings())
    assert [point["TIME_PERIOD"] for point in catalog.get("debt").domain] == [2009, 2010, 2011]
    assert [point["TIME_PERIOD"] for point in catalog.get("adjustment").domain] == [2010, 2011]
    assert schedule_coord("Engine!A2", catalog) == 0
    assert schedule_coord("Engine!B2", catalog) == schedule_coord("Engine!B3", catalog) == 1
    assert schedule_coord("Engine!C2", catalog) == schedule_coord("Engine!C3", catalog) == 2


def test_vertical_zipper_lag_is_not_a_same_index_cycle(tmp_path: Path) -> None:
    catalog, _deps, graph = inverted_graph_parts(
        _vertical_zipper_workbook(tmp_path), _vertical_zipper_bindings()
    )
    assert [point["TIME_PERIOD"] for point in catalog.get("debt").domain] == [2009, 2010, 2011]
    assert schedule_coord("Engine!B2", catalog) == 1
    assert schedule_coord("Engine!B1", catalog) == 0
    edges = collect_all_dependence_edges(catalog, graph)
    lag = next(
        edge
        for edge in edges
        if edge.consumer_id == "debt"
        and edge.producer_id == "debt"
        and edge.consumer_cell == "Engine!B2"
        and edge.producer_cell == "Engine!B1"
    )
    assert lag.distance == 1
    plan = plan_fused_scc(("debt", "adjustment"), catalog=catalog, graph=graph)
    assert plan is not None
    assert plan.body_order == ("adjustment", "debt")
    assert plan.schedule == (0, 1, 2)
    assert plan.domain["debt"] == (0, 3)
    assert plan.domain["adjustment"] == (1, 3)


def test_offset_helper_block_is_same_index_not_look_ahead(tmp_path: Path) -> None:
    catalog, _deps, graph = inverted_graph_parts(
        _offset_zipper_workbook(tmp_path), _offset_zipper_bindings()
    )
    assert schedule_coord("Engine!B2", catalog) == schedule_coord("Engine!E3", catalog)
    edges = collect_all_dependence_edges(catalog, graph)
    same_year = next(
        edge
        for edge in edges
        if edge.consumer_cell == "Engine!B2" and edge.producer_cell == "Engine!E3"
    )
    assert same_year.distance == 0
    plan = plan_fused_scc(("debt", "adjustment"), catalog=catalog, graph=graph)
    assert plan is not None
    assert plan.body_order == ("adjustment", "debt")
    assert plan.domain["adjustment"] == (1, 3)


def test_cross_sheet_coords_are_not_column_subtraction(tmp_path: Path) -> None:
    catalog, _deps, graph = inverted_graph_parts(
        _cross_sheet_zipper_workbook(tmp_path), _cross_sheet_zipper_bindings()
    )
    assert schedule_coord("Engine!B2", catalog) == schedule_coord("Helper!C2", catalog) == 1
    edges = collect_all_dependence_edges(catalog, graph)
    same_year = next(
        edge
        for edge in edges
        if edge.consumer_cell == "Engine!B2" and edge.producer_cell == "Helper!C2"
    )
    assert same_year.distance == 0
    plan = plan_fused_scc(("debt", "adjustment"), catalog=catalog, graph=graph)
    assert plan is not None
    assert "eval_instance" not in str(plan)
