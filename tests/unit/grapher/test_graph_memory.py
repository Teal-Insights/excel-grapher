"""Tests for the graph memory measurement harness (#490)."""

from __future__ import annotations

import json

import pytest

from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.guard import CellRef, Compare, Literal
from excel_grapher.grapher.node import make_cell_node
from scripts.measure_graph_memory import (
    GraphMemoryReport,
    deep_size,
    measure_graph_memory,
)

ROWS = 200

# Baselines measured on CPython 3.13 (64-bit) with `scripts/measure_graph_memory.py`
# against `_fixture_graph()`. Re-measure (do not hand-tune) when a change moves
# them out of band, and say in the commit which component moved.
_BYTES_PER_NODE = 1297.0
_BYTES_PER_EDGE = 2595.0
_NODE_BYTES_PER_NODE = 455.0
_PROVENANCE_BYTES_PER_EDGE = 469.0


def _fixture_graph(rows: int = ROWS, *, distinct_guards: bool = False) -> DependencyGraph:
    """Build a deterministic two-column graph with guards and provenance."""
    graph = DependencyGraph()
    shared_guard = Compare(CellRef("Sheet1!A1"), ">", Literal(0))
    for row in range(1, rows + 1):
        guard = (
            Compare(CellRef(f"Sheet1!C{row}"), ">", Literal(0)) if distinct_guards else shared_guard
        )
        graph.add_node(make_cell_node("Sheet1", "A", row, value=float(row), is_leaf=True))
        graph.add_node(
            make_cell_node(
                "Sheet1",
                "B",
                row,
                formula=f"=A{row}*2",
                normalized_formula=f"=Sheet1!A{row}*2",
                is_leaf=False,
                is_target=row == rows,
            )
        )
        graph.add_edge(
            f"Sheet1!B{row}",
            f"Sheet1!A{row}",
            # One shared guard object across every edge: the walk must count it once.
            guard=guard,
            # Distinct site offsets per edge: a literal tuple would be folded to
            # one shared constant and understate real provenance payload.
            provenance=EdgeProvenance(
                causes=DependencyCause.direct_ref,
                direct_sites_normalized=((row, row + 10),),
            ),
        )
    return graph


def _edge_count(graph: DependencyGraph) -> int:
    return sum(len(deps) for deps in graph._edges.values())


# ---- deep_size ------------------------------------------------------------


def test_deep_size_counts_each_distinct_object_once() -> None:
    payload = "x" * 5000
    assert deep_size(payload, payload) == deep_size(payload)
    holder = {"first": payload, "second": payload}
    assert deep_size(holder) < 2 * deep_size(payload)


def test_deep_size_walks_into_slotted_dataclasses() -> None:
    node = make_cell_node("Sheet1", "A", 1, formula="=" + "1" * 4000, is_leaf=False)
    assert deep_size(node) > 4000


def test_deep_size_does_not_walk_into_classes_or_modules() -> None:
    node = make_cell_node("Sheet1", "A", 1, value=1, is_leaf=True)
    # Reaching Node's type would pull in the class dict, module globals, and more.
    assert deep_size(node) < 2048


def test_deep_size_excludes_process_singletons_by_default() -> None:
    assert deep_size(DependencyCause.direct_ref) == 0
    assert deep_size(DependencyCause.direct_ref, include_singletons=True) > 0
    assert deep_size(None) == 0
    assert deep_size(7) == 0


def test_deep_size_of_empty_roots_is_zero() -> None:
    assert deep_size() == 0


# ---- component breakdown --------------------------------------------------


def test_report_covers_the_graph_components() -> None:
    report = measure_graph_memory(_fixture_graph())
    names = [component.name for component in report.components]
    assert names[:5] == [
        "nodes",
        "edges_forward",
        "edges_reverse",
        "guards",
        "provenance",
    ]
    assert report.node_count == 2 * ROWS
    assert report.edge_count == ROWS


def test_component_totals_reconcile_with_the_distinct_total() -> None:
    report = measure_graph_memory(_fixture_graph())
    exclusive = sum(component.exclusive_bytes for component in report.components)
    # Every reachable object is either exclusive to one component or shared.
    assert exclusive + report.shared_bytes == report.total_bytes
    # Naive per-component summing over-counts precisely because of sharing.
    assert sum(component.total_bytes for component in report.components) > report.total_bytes


def test_interned_keys_are_reported_as_shared_not_owned() -> None:
    report = measure_graph_memory(_fixture_graph())
    # Node keys live in _nodes, _edges, and _reverse_edges alike.
    assert report.component("nodes").shared_bytes > 0
    assert report.component("edges_forward").shared_bytes > 0
    assert report.shared_bytes > 0
    naive = sum(component.total_bytes for component in report.components)
    # Shared objects appear in >= 2 components, so naive summing over-counts them
    # at least once each.
    assert naive - report.total_bytes >= report.shared_bytes


def test_shared_guard_object_is_counted_once() -> None:
    shared = measure_graph_memory(_fixture_graph(rows=64)).component("guards")
    distinct = measure_graph_memory(_fixture_graph(rows=64, distinct_guards=True)).component(
        "guards"
    )
    # One GuardExpr tree reused by every edge must not be charged 64 times.
    assert distinct.exclusive_bytes > 3 * shared.exclusive_bytes


def test_scaffolding_is_a_subset_of_each_component_total() -> None:
    report = measure_graph_memory(_fixture_graph())
    for component in report.components:
        assert 0 <= component.scaffolding_bytes <= component.total_bytes
        assert component.exclusive_bytes <= component.total_bytes
        assert component.object_count > 0


def test_per_edge_provenance_outweighs_a_shared_guard() -> None:
    report = measure_graph_memory(_fixture_graph())
    provenance = report.component("provenance")
    guards = report.component("guards")
    # Both maps hold one entry per edge; provenance also owns a per-edge payload.
    assert provenance.exclusive_bytes > 3 * guards.exclusive_bytes


def test_per_node_and_per_edge_averages() -> None:
    graph = _fixture_graph()
    report = measure_graph_memory(graph)
    assert report.bytes_per_node == pytest.approx(report.total_bytes / len(graph))
    assert report.bytes_per_edge == pytest.approx(report.total_bytes / _edge_count(graph))
    nodes = report.component("nodes")
    assert nodes.bytes_per_node == pytest.approx(nodes.total_bytes / len(graph))


def test_empty_graph_averages_are_zero() -> None:
    report = measure_graph_memory(DependencyGraph())
    assert report.node_count == 0
    assert report.edge_count == 0
    assert report.bytes_per_node == 0.0
    assert report.bytes_per_edge == 0.0


def test_optional_maps_are_reported_when_present() -> None:
    graph = _fixture_graph(rows=2)
    graph.leaf_classification = {"Sheet1!A1": "input"}
    graph.sheet_order = ["Sheet1"]
    report = measure_graph_memory(graph)
    metadata = report.component("workbook_metadata")
    assert metadata.total_bytes > 0


# ---- rendering ------------------------------------------------------------


def test_render_marks_shared_and_owned_bytes() -> None:
    text = measure_graph_memory(_fixture_graph()).render()
    assert "component" in text
    assert "exclusive" in text
    assert "shared" in text
    assert "scaffold" in text
    assert "nodes" in text
    assert "provenance" in text
    # A reader must be able to tell re-attribution from a real drop.
    assert "shared with another component" in text


def test_to_dict_is_json_serializable() -> None:
    report = measure_graph_memory(_fixture_graph(rows=4))
    payload = report.to_dict()
    assert payload["node_count"] == 8
    assert payload["edge_count"] == 4
    components = payload["components"]
    assert isinstance(components, list)
    assert {c["name"] for c in components} >= {"nodes", "provenance"}
    assert json.loads(json.dumps(payload)) == payload


# ---- regression band ------------------------------------------------------


def _band(value: float, expected: float, tolerance: float = 0.25) -> None:
    assert expected * (1 - tolerance) <= value <= expected * (1 + tolerance), (
        f"{value:.1f} outside {tolerance:.0%} band around {expected:.1f}"
    )


def test_graph_memory_stays_within_the_measured_band() -> None:
    """Guard against size regressions; update the baselines with a measured number."""
    report = measure_graph_memory(_fixture_graph())
    _band(report.bytes_per_node, _BYTES_PER_NODE)
    _band(report.bytes_per_edge, _BYTES_PER_EDGE)
    _band(report.component("nodes").bytes_per_node, _NODE_BYTES_PER_NODE)
    _band(report.component("provenance").bytes_per_edge, _PROVENANCE_BYTES_PER_EDGE)


def test_node_component_dominated_by_strings_not_instances() -> None:
    """A `Node` instance is 152 bytes; its formula strings are the real cost."""
    report = measure_graph_memory(_fixture_graph())
    nodes = report.component("nodes")
    instance_bytes = 152 * report.node_count
    assert nodes.total_bytes > instance_bytes


def test_report_is_reproducible_for_equivalent_graphs() -> None:
    first = measure_graph_memory(_fixture_graph(rows=32))
    second = measure_graph_memory(_fixture_graph(rows=32))
    assert first.total_bytes == second.total_bytes
    assert [c.total_bytes for c in first.components] == [c.total_bytes for c in second.components]


def test_report_type_is_frozen() -> None:
    report = measure_graph_memory(DependencyGraph())
    assert isinstance(report, GraphMemoryReport)
    with pytest.raises(AttributeError):
        object.__setattr__(report, "unexpected_field", 1)
