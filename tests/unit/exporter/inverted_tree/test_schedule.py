"""Statement-graph legality: drop positive-distance edges, residual must be a DAG."""

from __future__ import annotations

import pytest

from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.schedule import (
    DependenceEdge,
    residual_body_order,
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
