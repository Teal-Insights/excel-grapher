"""Statement-graph legality and fused-SCC classification."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.schedule import (
    DependenceEdge,
    plan_fused_scc,
    residual_body_order,
)
from tests.unit.exporter.inverted_tree.helpers import inverted_graph_parts
from tests.unit.exporter.inverted_tree.test_shape_a11_zipper import (
    _simultaneous_workbook,
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


def test_plan_fused_scc_for_corrected_zipper(tmp_path: Path) -> None:
    catalog, _deps, graph = inverted_graph_parts(_zipper_workbook(tmp_path), _zipper_bindings())
    plan = plan_fused_scc(("debt", "adjustment"), catalog=catalog, graph=graph)
    assert plan is not None
    assert plan.body_order == ("adjustment", "debt")
    assert plan.schedule == (1, 2, 3)
    assert plan.domain["debt"] == (0, 3)
    assert plan.domain["adjustment"] == (1, 3)
    assert plan.peel_stop == 1


def test_plan_fused_scc_rejects_same_year_cycle(tmp_path: Path) -> None:
    catalog, _deps, graph = inverted_graph_parts(
        _simultaneous_workbook(tmp_path), _zipper_bindings()
    )
    with pytest.raises(InvertedTreeExportError, match="distance-zero residual"):
        plan_fused_scc(("debt", "adjustment"), catalog=catalog, graph=graph)
