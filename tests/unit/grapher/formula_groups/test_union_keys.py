"""Regression tests for PR #375 review blockers (workbook sort, viz, locate, order)."""

from __future__ import annotations

from excel_grapher.core.address_keys import sort_node_keys
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.lightweight_viz import (
    assemble_lightweight_viz_payload,
    build_lightweight_viz_core,
    write_lightweight_viz_data,
)
from excel_grapher.grapher.node import (
    NodeKind,
    locate_cell,
    locate_range,
    make_cell_node,
    make_union_node,
)


def test_workbook_sort_handles_column_range_and_union_keys() -> None:
    sheet_order = ["Sheet1"]
    keys = [
        "Sheet1!B2",
        "Sheet1!A1:A5",
        "Sheet1!A1:D5",
        "Sheet1!A1,C1",
        "Sheet1!Z1",
    ]
    # Must not raise; ordered by sheet, then top-left (min_row, min_col).
    ordered = sort_node_keys(keys, sheet_order=sheet_order)
    assert ordered[0] in {"Sheet1!A1:A5", "Sheet1!A1:D5", "Sheet1!A1,C1"}
    assert "Sheet1!Z1" in ordered
    assert "Sheet1!B2" in ordered


def test_graph_keys_workbook_order_with_union_and_column() -> None:
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(make_cell_node("Sheet1", "Z", 1))
    g.add_node(make_union_node(["Sheet1!A1", "Sheet1!C1"]))
    g.add_node(make_union_node([f"Sheet1!A{r}" for r in range(10, 15)]))
    # Must not raise.
    keys = g.keys(order="workbook")
    assert len(keys) == 3


def test_evaluation_order_follows_owner_when_edge_targets_member() -> None:
    g = DependencyGraph()
    union = make_union_node(["Sheet1!D63", "Sheet1!E63", "Sheet1!Y63"], is_leaf=True)
    dependent = make_cell_node(
        "Sheet1",
        "A",
        1,
        formula="=D63",
        normalized_formula="=Sheet1!D63",
        is_leaf=False,
    )
    g.add_node(union)
    g.add_node(dependent)
    # Edge names the member cell; owner resolution is explicit at read time.
    g.add_edge(dependent.key, "Sheet1!D63")

    assert g.get_dependencies(dependent.key) == frozenset({"Sheet1!D63"})
    assert g.get_dependency_nodes(dependent.key) == frozenset({union.key})
    assert g.resolve_endpoint("Sheet1!D63") == union.key
    assert g.get_dependents(union.key) == frozenset()
    assert g.get_dependents("Sheet1!D63") == frozenset({dependent.key})

    order = g.evaluation_order()
    assert order.index(union.key) < order.index(dependent.key)


def test_add_edge_keeps_raw_member_endpoints_and_resolves_dependencies() -> None:
    g = DependencyGraph()
    left = make_union_node(["Sheet1!A1", "Sheet1!B1"], is_leaf=False)
    right = make_union_node(["Sheet1!D63", "Sheet1!Y63"], is_leaf=True)
    g.add_node(left)
    g.add_node(right)
    g.add_edge("Sheet1!A1", "Sheet1!D63")

    assert g.get_dependencies(left.key) == frozenset()
    assert g.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!D63"})
    assert g.get_dependency_nodes("Sheet1!A1") == frozenset({right.key})
    assert g.get_dependents(right.key) == frozenset()
    assert g.get_dependents("Sheet1!D63") == frozenset({"Sheet1!A1"})
    attrs = g.get_edge_attrs("Sheet1!A1", "Sheet1!D63")
    assert attrs.guard is None
    assert g.resolve_endpoint("Sheet1!A1") == left.key
    assert g.resolve_endpoint("Sheet1!D63") == right.key
    assert g.get_edge_attrs(left.key, right.key).guard is None


def test_to_networkx_has_no_dangling_member_endpoints() -> None:
    from excel_grapher.grapher.export import to_networkx

    g = DependencyGraph()
    union = make_union_node(["Sheet1!D63", "Sheet1!E63"], is_leaf=True)
    dependent = make_cell_node("Sheet1", "A", 1, formula="=D63", is_leaf=False)
    g.add_node(union)
    g.add_node(dependent)
    g.add_edge(dependent.key, "Sheet1!D63")

    nx_graph = to_networkx(g)
    assert set(nx_graph.nodes) == {dependent.key, union.key}
    assert list(nx_graph.edges) == [(dependent.key, union.key)]


def test_locate_range_finds_union_owner_of_subspan() -> None:
    g = DependencyGraph()
    g.add_node(make_union_node(["Sheet1!D63", "Sheet1!E63", "Sheet1!F63", "Sheet1!Y63"]))
    loc = locate_range(g, "Sheet1!D63:F63")
    assert loc is not None
    assert loc.node_key == "Sheet1!D63:F63,Y63"
    assert loc.kind is NodeKind.union
    assert locate_cell(g, "Sheet1!D63") is not None


def test_lightweight_viz_accepts_union_nodes(tmp_path) -> None:
    g = DependencyGraph()
    g.sheet_order = ["Sheet1", "Sheet2"]
    g.add_node(make_union_node(["Sheet1!A1", "Sheet2!B2"]))
    g.add_node(make_cell_node("Sheet1", "Z", 1, formula="=A1", is_leaf=False))
    g.add_edge("Sheet1!Z1", "Sheet1!A1")
    out = tmp_path / "viz.json"
    # Must not raise KeyError / AssertionError on None sheet/column/row.
    core = build_lightweight_viz_core(g)
    payload = assemble_lightweight_viz_payload(core, [])
    write_lightweight_viz_data(payload, out)
    assert out.exists() and out.stat().st_size > 0
