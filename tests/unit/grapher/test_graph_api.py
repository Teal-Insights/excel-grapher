"""Tests for the DependencyGraph public API contract.

These tests pin the Phase 1 contracts:

* `get_node`, `get_dependencies`, `get_dependents`, `get_edge_guard`,
  `get_edge_attrs` all normalize their key arguments.
* Readers return immutable/snapshot containers so that mutating the returned
  value cannot silently mutate the graph's internal state.
* `set_node_value` and `set_node_metadata` are the durable node mutation path
  and raise `KeyError` when the key is missing.
* `get_node` returns a `NodeView` snapshot (not the live storage `Node`).
"""

from __future__ import annotations

from collections.abc import MutableMapping
from operator import setitem
from typing import cast

import pytest

from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph, EdgeAttrs
from excel_grapher.grapher.guard import CellRef as GuardCellRef
from excel_grapher.grapher.guard import Compare, Literal
from excel_grapher.grapher.node import Node, NodeView


def _leaf(sheet: str, col: str, row: int, value: object = 0) -> Node:
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=None,
        normalized_formula=None,
        value=value,
        is_leaf=True,
    )


def _formula(sheet: str, col: str, row: int, formula: str, value: object = None) -> Node:
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=value,
        is_leaf=False,
    )


# -------------------------------------------------------------------
# get_node: NodeView snapshot + normalization
# -------------------------------------------------------------------


def test_get_node_returns_node_view_not_storage_node() -> None:
    """`get_node` must return a `NodeView` snapshot, not the internal `Node` object."""
    g = DependencyGraph()
    g.add_node(_leaf("S", "A", 1, value=7))

    view = g.get_node("S!A1")
    assert view is not None
    assert isinstance(view, NodeView)
    assert not isinstance(view, Node)
    assert view.value == 7
    assert view.sheet == "S"
    assert view.column == "A"
    assert view.row == 1
    assert view.is_leaf is True


def test_get_node_view_is_frozen() -> None:
    """Attempting to mutate a NodeView must raise (dataclass frozen)."""
    g = DependencyGraph()
    g.add_node(_leaf("S", "A", 1, value=7))
    view = g.get_node("S!A1")
    assert view is not None
    attr_name = "value"
    with pytest.raises((AttributeError, TypeError)):
        setattr(view, attr_name, 42)


def test_get_node_view_metadata_is_read_only() -> None:
    """Mutating the metadata mapping on a NodeView must fail."""
    g = DependencyGraph()
    node = _leaf("S", "A", 1, value=7)
    node.set_metadata({"k": "v"})
    g.add_node(node)

    view = g.get_node("S!A1")
    assert view is not None
    assert view.metadata["k"] == "v"
    with pytest.raises((TypeError, AttributeError)):
        setitem(cast(MutableMapping[str, object], view.metadata), "k", "other")


def test_get_node_normalizes_quoted_sheet() -> None:
    """Quoting style on input must not affect lookup."""
    g = DependencyGraph()
    g.add_node(_leaf("S", "A", 1))

    assert g.get_node("S!A1") is not None
    assert g.get_node("'S'!A1") is not None  # unnecessary quotes normalized away


def test_get_node_normalizes_sheet_with_space() -> None:
    g = DependencyGraph()
    g.add_node(_leaf("My Sheet", "B", 2))

    view = g.get_node("'My Sheet'!B2")
    assert view is not None
    assert view.sheet == "My Sheet"


def test_get_node_normalizes_sheet_with_apostrophe() -> None:
    g = DependencyGraph()
    g.add_node(_leaf("It's Data", "A", 1))

    view = g.get_node("'It''s Data'!A1")
    assert view is not None
    assert view.sheet == "It's Data"


def test_get_node_missing_returns_none() -> None:
    g = DependencyGraph()
    assert g.get_node("S!A1") is None


# -------------------------------------------------------------------
# get_dependencies / get_dependents: immutable snapshots + normalization
# -------------------------------------------------------------------


def test_get_dependencies_returns_frozenset() -> None:
    g = DependencyGraph()
    g.add_node(_leaf("S", "A", 1))
    g.add_node(_formula("S", "B", 1, "=S!A1"))
    g.add_edge("S!B1", "S!A1")

    deps = g.get_dependencies("S!B1")
    assert isinstance(deps, frozenset)
    assert deps == frozenset({"S!A1"})


def test_get_dependents_returns_frozenset() -> None:
    g = DependencyGraph()
    g.add_node(_leaf("S", "A", 1))
    g.add_node(_formula("S", "B", 1, "=S!A1"))
    g.add_edge("S!B1", "S!A1")

    dependents = g.get_dependents("S!A1")
    assert isinstance(dependents, frozenset)
    assert dependents == frozenset({"S!B1"})


def test_get_dependencies_snapshot_is_decoupled_from_graph() -> None:
    """Modifying the returned container must not affect the graph."""
    g = DependencyGraph()
    g.add_node(_leaf("S", "A", 1))
    g.add_node(_formula("S", "B", 1, "=S!A1"))
    g.add_edge("S!B1", "S!A1")

    deps = g.get_dependencies("S!B1")
    assert deps == frozenset({"S!A1"})
    # Returned frozenset cannot be mutated; but even re-binding shouldn't leak.
    # Verify subsequent calls still see the original state.
    assert g.get_dependencies("S!B1") == frozenset({"S!A1"})


def test_get_dependencies_missing_key_returns_empty() -> None:
    g = DependencyGraph()
    assert g.get_dependencies("S!A1") == frozenset()


def test_get_dependencies_normalizes_key() -> None:
    g = DependencyGraph()
    g.add_node(_leaf("My Sheet", "A", 1))
    g.add_node(_formula("My Sheet", "B", 1, "='My Sheet'!A1"))
    g.add_edge("'My Sheet'!B1", "'My Sheet'!A1")

    assert g.get_dependencies("'My Sheet'!B1") == frozenset({"'My Sheet'!A1"})


def test_keys_order_workbook_uses_sheet_order_then_row_then_column() -> None:
    g = DependencyGraph(sheet_order=["Later", "Earlier"])
    g.add_node(_leaf("Earlier", "B", 2))
    g.add_node(_leaf("Earlier", "A", 1))
    g.add_node(_leaf("Later", "A", 2))

    assert g.keys(order="workbook") == ["Later!A2", "Earlier!A1", "Earlier!B2"]


def _build_diamond_graph(*, reverse_insertion: bool) -> DependencyGraph:
    """Diamond: A1 leaf; B1 and C1 depend on A1; D1 depends on B1 and C1."""
    nodes = [
        _leaf("Sheet1", "A", 1, value=1),
        _formula("Sheet1", "B", 1, "=Sheet1!A1+1"),
        _formula("Sheet1", "C", 1, "=Sheet1!A1*2"),
        _formula("Sheet1", "D", 1, "=Sheet1!B1+Sheet1!C1"),
    ]
    if reverse_insertion:
        nodes = list(reversed(nodes))

    graph = DependencyGraph(sheet_order=["Sheet1"])
    for node in nodes:
        graph.add_node(node)
    graph.add_edge("Sheet1!B1", "Sheet1!A1")
    graph.add_edge("Sheet1!C1", "Sheet1!A1")
    graph.add_edge("Sheet1!D1", "Sheet1!B1")
    graph.add_edge("Sheet1!D1", "Sheet1!C1")
    return graph


def test_evaluation_order_breaks_ties_by_workbook_order() -> None:
    graph = _build_diamond_graph(reverse_insertion=True)

    order = graph.evaluation_order()

    assert order.index("Sheet1!A1") < order.index("Sheet1!B1")
    assert order.index("Sheet1!A1") < order.index("Sheet1!C1")
    assert order.index("Sheet1!B1") < order.index("Sheet1!C1")
    assert order.index("Sheet1!C1") < order.index("Sheet1!D1")


def test_evaluation_order_is_independent_of_node_insertion_order() -> None:
    graph_forward = _build_diamond_graph(reverse_insertion=False)
    graph_reverse = _build_diamond_graph(reverse_insertion=True)

    assert graph_forward.evaluation_order() == graph_reverse.evaluation_order()


# -------------------------------------------------------------------
# get_edge_attrs / get_edge_guard: typed container + normalization
# -------------------------------------------------------------------


def test_get_edge_attrs_returns_typed_container() -> None:
    g = DependencyGraph()
    g.add_node(_leaf("S", "A", 1))
    g.add_node(_formula("S", "B", 1, "=S!A1"))
    prov = EdgeProvenance(
        causes=DependencyCause.direct_ref,
        direct_sites_normalized=((1, 5),),
    )
    g.add_edge("S!B1", "S!A1", provenance=prov)

    attrs = g.get_edge_attrs("S!B1", "S!A1")
    assert isinstance(attrs, EdgeAttrs)
    assert attrs.provenance is not None
    assert attrs.provenance.causes == DependencyCause.direct_ref
    assert attrs.guard is None


def test_get_edge_attrs_for_missing_edge_returns_empty_container() -> None:
    g = DependencyGraph()
    g.add_node(_leaf("S", "A", 1))
    g.add_node(_leaf("S", "B", 1))

    attrs = g.get_edge_attrs("S!B1", "S!A1")
    assert isinstance(attrs, EdgeAttrs)
    assert attrs.guard is None
    assert attrs.provenance is None


def test_get_edge_attrs_normalizes_keys() -> None:
    g = DependencyGraph()
    g.add_node(_leaf("My Sheet", "A", 1))
    g.add_node(_formula("My Sheet", "B", 1, "='My Sheet'!A1"))
    prov = EdgeProvenance(causes=DependencyCause.direct_ref)
    g.add_edge("'My Sheet'!B1", "'My Sheet'!A1", provenance=prov)

    attrs = g.get_edge_attrs("'My Sheet'!B1", "'My Sheet'!A1")
    assert attrs.provenance is not None


def test_get_edge_guard_returns_guard_expr() -> None:
    g = DependencyGraph()
    g.add_node(_leaf("S", "A", 1))
    g.add_node(_leaf("S", "C", 1))
    g.add_node(_formula("S", "B", 1, "=IF(S!C1, S!A1, 0)"))
    guard = Compare(left=GuardCellRef(key="S!C1"), op="=", right=Literal(True))
    g.add_edge("S!B1", "S!A1", guard=guard)

    assert g.get_edge_guard("S!B1", "S!A1") == guard


def test_get_edge_guard_missing_is_none() -> None:
    g = DependencyGraph()
    g.add_node(_leaf("S", "A", 1))
    g.add_node(_formula("S", "B", 1, "=S!A1"))
    g.add_edge("S!B1", "S!A1")

    assert g.get_edge_guard("S!B1", "S!A1") is None


def test_add_edge_rejects_unknown_edge_attrs() -> None:
    g = DependencyGraph()
    g.add_node(_leaf("S", "A", 1))
    g.add_node(_formula("S", "B", 1, "=S!A1"))
    with pytest.raises(TypeError):
        g.add_edge("S!B1", "S!A1", weight=3)  # ty: ignore[unknown-argument]


def test_edge_provenance_stored_in_typed_map() -> None:
    from dataclasses import fields as dc_fields

    g = DependencyGraph()
    g.add_node(_leaf("S", "A", 1))
    g.add_node(_formula("S", "B", 1, "=S!A1"))
    prov = EdgeProvenance(
        causes=DependencyCause.direct_ref,
        direct_sites_normalized=((1, 5),),
    )
    g.add_edge("S!B1", "S!A1", provenance=prov)

    assert ("S!B1", "S!A1") in g._edge_provenance
    assert g._edge_provenance[("S!B1", "S!A1")] == prov
    field_names = {f.name for f in dc_fields(DependencyGraph)}
    assert "_edge_provenance" in field_names
    assert "_edge_extra" not in field_names


def test_add_edge_merges_provenance_on_existing_edge() -> None:
    g = DependencyGraph()
    g.add_node(_leaf("S", "A", 1))
    g.add_node(_formula("S", "B", 1, "=S!A1"))
    g.add_edge(
        "S!B1",
        "S!A1",
        provenance=EdgeProvenance(causes=DependencyCause.direct_ref),
    )
    g.add_edge(
        "S!B1",
        "S!A1",
        provenance=EdgeProvenance(causes=DependencyCause.static_range),
    )

    merged = g._edge_provenance[("S!B1", "S!A1")]
    assert merged.causes == (DependencyCause.direct_ref | DependencyCause.static_range)


def test_add_edge_without_provenance_preserves_existing() -> None:
    g = DependencyGraph()
    g.add_node(_leaf("S", "A", 1))
    g.add_node(_formula("S", "B", 1, "=S!A1"))
    prov = EdgeProvenance(causes=DependencyCause.direct_ref)
    g.add_edge("S!B1", "S!A1", provenance=prov)
    g.add_edge("S!B1", "S!A1")

    assert g._edge_provenance[("S!B1", "S!A1")] == prov


def test_remove_edge_clears_provenance() -> None:
    g = DependencyGraph()
    g.add_node(_leaf("S", "A", 1))
    g.add_node(_formula("S", "B", 1, "=S!A1"))
    g.add_edge(
        "S!B1",
        "S!A1",
        provenance=EdgeProvenance(causes=DependencyCause.direct_ref),
    )
    g._remove_edge("S!B1", "S!A1")

    assert ("S!B1", "S!A1") not in g._edge_provenance
    assert g.get_edge_attrs("S!B1", "S!A1").provenance is None


# -------------------------------------------------------------------
# set_node_value: durable mutation of leaf/formula value
# -------------------------------------------------------------------


def test_set_node_value_updates_node() -> None:
    g = DependencyGraph()
    g.add_node(_leaf("S", "A", 1, value=10))

    g.set_node_value("S!A1", 42)

    view = g.get_node("S!A1")
    assert view is not None
    assert view.value == 42


def test_set_node_value_normalizes_key() -> None:
    g = DependencyGraph()
    g.add_node(_leaf("My Sheet", "A", 1, value=10))

    g.set_node_value("'My Sheet'!A1", 99)

    view = g.get_node("'My Sheet'!A1")
    assert view is not None
    assert view.value == 99


def test_set_node_value_missing_key_raises_key_error() -> None:
    g = DependencyGraph()
    with pytest.raises(KeyError):
        g.set_node_value("S!A1", 1)


def test_set_node_value_does_not_change_other_fields() -> None:
    g = DependencyGraph()
    g.add_node(_formula("S", "B", 1, "=S!A1", value=20))

    g.set_node_value("S!B1", 7)

    view = g.get_node("S!B1")
    assert view is not None
    assert view.value == 7
    assert view.formula == "=S!A1"
    assert view.is_leaf is False


# -------------------------------------------------------------------
# set_node_metadata: durable mutation of metadata
# -------------------------------------------------------------------


def test_set_node_metadata_replaces_mapping() -> None:
    g = DependencyGraph()
    node = _leaf("S", "A", 1)
    node.set_metadata({"existing": "before"})
    g.add_node(node)

    g.set_node_metadata("S!A1", {"row_labels": ["Revenue"], "column_labels": ["2024"]})

    view = g.get_node("S!A1")
    assert view is not None
    assert dict(view.metadata) == {
        "row_labels": ["Revenue"],
        "column_labels": ["2024"],
    }


def test_set_node_metadata_copies_mapping_input() -> None:
    """Mutating the caller's dict after set_node_metadata must not affect the graph."""
    g = DependencyGraph()
    g.add_node(_leaf("S", "A", 1))

    source: dict[str, object] = {"k": 1}
    g.set_node_metadata("S!A1", source)
    source["k"] = 2
    source["added_after"] = 99

    view = g.get_node("S!A1")
    assert view is not None
    assert dict(view.metadata) == {"k": 1}


def test_set_node_metadata_normalizes_key() -> None:
    g = DependencyGraph()
    g.add_node(_leaf("My Sheet", "A", 1))

    g.set_node_metadata("'My Sheet'!A1", {"label": "x"})

    view = g.get_node("'My Sheet'!A1")
    assert view is not None
    assert dict(view.metadata) == {"label": "x"}


def test_set_node_metadata_missing_key_raises_key_error() -> None:
    g = DependencyGraph()
    with pytest.raises(KeyError):
        g.set_node_metadata("S!A1", {"k": "v"})


# -------------------------------------------------------------------
# Evaluator API: FormulaEvaluator.set_value must be removed
# -------------------------------------------------------------------


def test_formula_evaluator_has_no_set_value_method() -> None:
    from excel_grapher.evaluator.evaluator import FormulaEvaluator

    assert not hasattr(FormulaEvaluator, "set_value"), (
        "FormulaEvaluator.set_value must be removed; use DependencyGraph.set_node_value."
    )
