"""Tests for the graph memory measurement harness (#490 / #906)."""

from __future__ import annotations

import json
import sys

import pytest

from excel_grapher.core.formula_ast import parse_formula_text
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.formula_shapes import warm_formula_shapes
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
# against `_fixture_graph()`, after CSR intern-id edge metadata (#920). Re-measure
# (do not hand-tune) when a change moves them out of band, and say in the commit
# which component moved.
_BYTES_PER_NODE = 645.1
_BYTES_PER_EDGE = 1290.2
_NODE_BYTES_PER_NODE = 495.4
_PROVENANCE_BYTES_PER_EDGE = 166.0


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
    graph.rebuild_adjacency()
    return graph


def _edge_count(graph: DependencyGraph) -> int:
    return graph.edge_count()


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


def test_report_covers_the_current_graph_components() -> None:
    report = measure_graph_memory(_fixture_graph())
    names = [component.name for component in report.components]
    assert names[0] == "nodes"
    assert "adjacency_csr" in names
    assert "node_index" in names
    assert "edges_forward" not in names
    assert "edges_reverse" not in names
    assert "guards" in names
    assert "provenance" in names
    assert "occupancy" not in names
    # Compact CSR has no per-node empty neighbor sets.
    assert "edges_forward_empty" not in names
    assert "edges_reverse_empty" not in names
    assert report.node_count == 2 * ROWS
    assert report.edge_count == ROWS
    assert report.empty_forward_sets == 0
    assert report.empty_reverse_sets == 0


def test_component_totals_reconcile_with_the_distinct_total() -> None:
    report = measure_graph_memory(_fixture_graph())
    exclusive = sum(component.exclusive_bytes for component in report.components)
    # Every reachable object is either exclusive to one component or shared.
    assert exclusive + report.shared_bytes == report.total_bytes
    # Naive per-component summing over-counts precisely because of sharing.
    assert sum(component.total_bytes for component in report.components) > report.total_bytes


def test_edge_keys_are_reported_as_shared_not_owned() -> None:
    report = measure_graph_memory(_fixture_graph())
    # CSR `node_index` reuses `_nodes` key objects. Compact intern-id metadata
    # no longer stores `EdgeKey` tuples, so those strings are not charged to
    # guards/provenance.
    assert report.component("node_index").shared_bytes > 0
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
    # Compact intern-id provenance still owns a per-edge payload; a shared guard
    # intern table does not.
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
    assert report.empty_forward_sets == 0
    assert report.empty_reverse_sets == 0
    assert report.formula_ast_intern_count == 0


def test_optional_maps_are_reported_when_present() -> None:
    graph = _fixture_graph(rows=2)
    graph.leaf_classification = {"Sheet1!A1": "input"}
    graph.sheet_order = ["Sheet1"]
    report = measure_graph_memory(graph)
    metadata = report.component("workbook_metadata")
    assert metadata.total_bytes > 0


def test_formula_shapes_overlay_is_reported_when_present() -> None:
    graph = _fixture_graph(rows=4)
    graph.formula_shapes = warm_formula_shapes(graph)
    report = measure_graph_memory(graph)
    shapes = report.component("formula_shapes")
    assert shapes.total_bytes > 0


def test_isolated_nodes_do_not_pay_an_empty_adjacency_set_tax() -> None:
    graph = DependencyGraph()
    n_nodes = 50
    for row in range(1, n_nodes + 1):
        graph.add_node(make_cell_node("Sheet1", "A", row, value=float(row), is_leaf=True))
    report = measure_graph_memory(graph)
    assert report.empty_forward_sets == 0
    assert report.empty_reverse_sets == 0
    assert report.empty_forward_bytes == 0
    assert report.empty_reverse_bytes == 0
    names = [component.name for component in report.components]
    assert "edges_forward_empty" not in names
    assert "edges_reverse_empty" not in names


def test_harness_splits_empty_adjacency_sets_when_present() -> None:
    """Empty-set rows are a partition of the parent walk, not a second walk."""
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Sheet1", "A", 1, value=1.0, is_leaf=True))
    graph._edges["Sheet1!A1"] = set()
    graph._reverse_edges["Sheet1!A1"] = set()
    report = measure_graph_memory(graph)
    assert report.empty_forward_sets == 1
    assert report.empty_reverse_sets == 1
    assert report.empty_forward_bytes == report.component("edges_forward_empty").exclusive_bytes
    assert report.empty_reverse_bytes == report.component("edges_reverse_empty").exclusive_bytes
    assert report.empty_forward_bytes >= 200
    assert report.empty_reverse_bytes >= 200


def test_guard_intern_pool_and_guarded_edge_counts() -> None:
    shared = measure_graph_memory(_fixture_graph(rows=32))
    assert shared.guarded_edge_count == 32
    assert shared.identity_distinct_guards == 1
    assert shared.guard_intern_pool_size >= shared.identity_distinct_guards

    distinct = measure_graph_memory(_fixture_graph(rows=32, distinct_guards=True))
    assert distinct.guarded_edge_count == 32
    assert distinct.identity_distinct_guards == 32


def test_formula_ast_intern_pool_counts_shared_trees_once() -> None:
    graph = DependencyGraph()
    ast = parse_formula_text("=Sheet1!A1*2", anchor="Sheet1!B1")
    assert ast is not None
    graph.add_node(make_cell_node("Sheet1", "A", 1, value=1.0, is_leaf=True))
    graph.add_node(make_cell_node("Sheet1", "B", 1, formula_ast=ast, is_leaf=False))
    graph.add_node(make_cell_node("Sheet1", "B", 2, formula_ast=ast, is_leaf=False))
    report = measure_graph_memory(graph)
    assert report.formula_nodes_with_ast == 2
    assert report.formula_ast_intern_count == 1
    assert report.formula_ast_intern_bytes > 0


# ---- rendering ------------------------------------------------------------


def test_render_marks_shared_and_owned_bytes() -> None:
    text = measure_graph_memory(_fixture_graph()).render()
    assert "component" in text
    assert "exclusive" in text
    assert "shared" in text
    assert "scaffold" in text
    assert "nodes" in text
    assert "provenance" in text
    assert "empty-set" in text
    assert "guard intern pool" in text
    assert "formula_ast intern" in text
    # A reader must be able to tell re-attribution from a real drop.
    assert "shared with another component" in text
    assert "occupancy" not in text
    assert "direct_sites_formula" not in text


def test_to_dict_is_json_serializable() -> None:
    report = measure_graph_memory(_fixture_graph(rows=4))
    payload = report.to_dict()
    assert payload["node_count"] == 8
    assert payload["edge_count"] == 4
    assert payload["guarded_edge_count"] == 4
    assert payload["empty_forward_sets"] >= 0
    assert payload["formula_ast_intern_count"] > 0
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


def test_node_component_dominated_by_payload_not_instances() -> None:
    """A slotted `Node` instance is small; formula ASTs and addresses dominate."""
    report = measure_graph_memory(_fixture_graph())
    nodes = report.component("nodes")
    instance_bytes = sys.getsizeof(make_cell_node("Sheet1", "A", 1, value=1.0, is_leaf=True))
    assert nodes.total_bytes > instance_bytes * report.node_count


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
