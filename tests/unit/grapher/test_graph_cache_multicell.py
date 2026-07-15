"""Graph JSON cache round-trips for multi-cell range / union nodes."""

from __future__ import annotations

from excel_grapher.core.address_keys import parse_node_key
from excel_grapher.grapher.cache import (
    GRAPH_CACHE_SCHEMA_VERSION,
    dependency_graph_from_json,
    dependency_graph_to_json,
)
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node, NodeKind, locate_cell, make_cell_node, make_union_node


def test_graph_cache_schema_version_is_at_least_3() -> None:
    assert GRAPH_CACHE_SCHEMA_VERSION >= 3


def test_json_roundtrip_preserves_union_row_and_member_edges() -> None:
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    union = make_union_node(["Sheet1!D63", "Sheet1!E63", "Sheet1!Y63"], is_leaf=False)
    row = make_union_node(["Sheet1!A10", "Sheet1!B10", "Sheet1!C10"], is_leaf=True)
    leaf = make_cell_node("Sheet1", "Z", 1, value=7, is_leaf=True)
    dependent = make_cell_node(
        "Sheet1",
        "A",
        1,
        formula="=D63",
        normalized_formula="=Sheet1!D63",
        is_leaf=False,
    )
    g.add_node(union)
    g.add_node(row)
    g.add_node(leaf)
    g.add_node(dependent)
    g.add_edge(union.key, leaf.key)
    g.add_edge(dependent.key, "Sheet1!D63")

    payload = dependency_graph_to_json(g)
    restored = dependency_graph_from_json(payload)

    assert set(restored) == set(g)
    restored_union = restored.get_node(union.key)
    assert restored_union is not None
    assert restored_union.kind is NodeKind.union
    assert restored_union.address is not None
    assert str(restored_union.address) == union.key
    assert restored_union.formula is None
    assert restored_union.normalized_formula is None
    assert restored_union.value is None

    restored_row = restored.get_node(row.key)
    assert restored_row is not None
    assert restored_row.kind is NodeKind.union
    assert restored_row.min_col == "A"
    assert restored_row.max_col == "C"
    assert restored_row.row == 10

    loc = locate_cell(restored, "Sheet1!D63")
    assert loc is not None
    assert loc.node_key == union.key

    assert restored.get_dependencies(dependent.key) == frozenset({"Sheet1!D63"})
    assert restored.get_dependency_nodes(dependent.key) == frozenset({union.key})
    assert restored.get_dependencies(union.key) == frozenset({leaf.key})
    restored_leaf = restored.get_node(leaf.key)
    assert restored_leaf is not None
    assert restored_leaf.value == 7


def test_json_roundtrip_one_by_one_address_collapses_to_cell() -> None:
    g = DependencyGraph()
    node = Node(
        sheet=None,
        column=None,
        row=None,
        formula=None,
        normalized_formula=None,
        value=3,
        is_leaf=True,
        address=parse_node_key("Sheet1!D63:D63"),
    )
    g.add_node(node)
    assert node.key == "Sheet1!D63"
    assert node.kind is NodeKind.cell

    restored = dependency_graph_from_json(dependency_graph_to_json(g))
    restored_node = restored.get_node("Sheet1!D63:D63")
    assert restored_node is not None
    assert restored_node.kind is NodeKind.cell
    assert restored_node.key == "Sheet1!D63"
    loc = locate_cell(restored, "Sheet1!D63")
    assert loc is not None
    assert loc.node_key == "Sheet1!D63"


def test_json_payload_preserves_union_addresses_and_kinds() -> None:
    g = DependencyGraph()
    g.add_node(make_union_node(["Sheet1!A1", "Sheet1!C1"]))
    g.add_node(make_cell_node("Sheet1", "B", 2, value=1))
    payload = dependency_graph_to_json(g)
    by_key = {n["key"]: n for n in payload["nodes"]}
    assert by_key["Sheet1!A1,C1"]["address"] == "Sheet1!A1,C1"
    assert by_key["Sheet1!B2"]["address"] == "Sheet1!B2"
    assert by_key["Sheet1!A1,C1"]["kind"] == "union"
    assert by_key["Sheet1!B2"]["kind"] == "cell"
