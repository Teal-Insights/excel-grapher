"""Tests for slotted Node storage and address-keyed derived-field LRU (#476)."""

from __future__ import annotations

import copy
import sys

import pytest

from excel_grapher.core.address_keys import CellKey, NodeShape, parse_node_key
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import (
    Node,
    NodeKind,
    _derived_fields_cache_clear,
    _derived_fields_cache_info,
    make_cell_node,
    make_union_node,
)


def _leaf(sheet: str = "Sheet1", column: str = "A", row: int = 1, value: object = 1) -> Node:
    return make_cell_node(sheet, column, row, value=value, is_leaf=True)


def test_node_uses_slots_without_instance_dict() -> None:
    node = _leaf()
    assert hasattr(Node, "__slots__")
    assert not hasattr(node, "__dict__")


def test_node_rejects_arbitrary_attribute_assignment() -> None:
    node = _leaf()
    with pytest.raises(AttributeError):
        # Bypass static attribute checks; slots must still reject unknown names.
        object.__setattr__(node, "not_a_field", 123)


def test_node_public_fields_and_derived_properties() -> None:
    node = make_cell_node(
        "Sheet1",
        "B",
        2,
        formula="=A1",
        normalized_formula="=Sheet1!A1",
        value=None,
        is_leaf=False,
        is_target=True,
        metadata={"k": 1},
    )
    assert node.sheet == "Sheet1"
    assert node.column == "B"
    assert node.row == 2
    assert node.formula == "=A1"
    assert node.normalized_formula == "=Sheet1!A1"
    assert node.value is None
    assert node.is_leaf is False
    assert node.is_target is True
    assert node.metadata == {"k": 1}
    assert node.kind is NodeKind.cell
    assert node.min_col == "B"
    assert node.max_col == "B"
    assert node.min_row == 2
    assert node.max_row == 2
    assert isinstance(node.address, CellKey)
    assert node.key == "Sheet1!B2"
    assert node.shape is NodeShape.cell
    assert node.column_index == 2


def test_node_extent_and_address_construction() -> None:
    extent = Node(
        sheet="Sheet1",
        column=None,
        row=None,
        formula=None,
        normalized_formula=None,
        value=None,
        is_leaf=True,
        min_col="D",
        max_col="Y",
        min_row=63,
        max_row=63,
    )
    assert extent.key == "Sheet1!D63:Y63"
    assert extent.shape is NodeShape.row
    assert extent.kind is NodeKind.union

    by_address = Node(
        sheet=None,
        column=None,
        row=None,
        formula=None,
        normalized_formula=None,
        value=None,
        is_leaf=True,
        address=parse_node_key("Sheet1!E5"),
    )
    assert by_address.key == "Sheet1!E5"
    assert by_address.column_index == 5


def test_graph_mutations_on_slotted_node() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf(value=1))
    graph.set_node_value("Sheet1!A1", 9)
    graph.set_node_metadata("Sheet1!A1", {"tag": "x"})
    graph.set_node_formula("Sheet1!A1", "=1", "=1")
    view = graph.get_node("Sheet1!A1")
    assert view is not None
    assert view.value == 9
    assert dict(view.metadata) == {"tag": "x"}
    assert view.formula == "=1"
    assert view.normalized_formula == "=1"


def test_deepcopy_and_projection_clone_preserve_slotted_nodes() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf(value=3))
    graph.add_node(
        make_cell_node(
            "Sheet1",
            "B",
            1,
            formula="=A1",
            normalized_formula="=Sheet1!A1",
            is_leaf=False,
        )
    )
    graph.add_edge("Sheet1!B1", "Sheet1!A1")

    cloned = copy.deepcopy(graph)
    cloned_view = cloned.get_node("Sheet1!A1")
    assert cloned_view is not None
    assert cloned_view.value == 3
    assert not hasattr(cloned._nodes["Sheet1!A1"], "__dict__")

    projected = graph._copy_for_projection()
    assert not hasattr(projected._nodes["Sheet1!A1"], "__dict__")
    assert projected.get_node("Sheet1!B1") is not None


def test_derived_fields_lru_keyed_on_address() -> None:
    from excel_grapher.grapher.node import _lookup_derived_fields

    _derived_fields_cache_clear()
    info0 = _derived_fields_cache_info()
    assert info0.hits == 0
    assert info0.misses == 0
    assert info0.currsize == 0

    a = _leaf("Sheet1", "C", 3)
    b = _leaf("Sheet1", "C", 3)
    assert a.key == "Sheet1!C3"
    assert b.key == "Sheet1!C3"
    assert a.shape is NodeShape.cell
    assert b.column_index == 3

    info = _derived_fields_cache_info()
    assert info.misses == 1
    assert info.hits == 3  # b.key, a.shape, b.column_index
    assert info.currsize == 1
    assert info.maxsize >= 1

    # Plain str and AddressKey must share one dict entry (unlike functools.lru_cache,
    # which keys str subclasses distinctly via _make_key).
    before = _derived_fields_cache_info()
    again = _lookup_derived_fields("Sheet1!C3")
    after = _derived_fields_cache_info()
    assert again.key == "Sheet1!C3"
    assert after.hits == before.hits + 1
    assert after.misses == before.misses
    assert after.currsize == 1


def test_slotted_node_is_smaller_than_dict_backed_baseline() -> None:
    """Slots should drop the per-instance __dict__ (~300 bytes on CPython)."""
    node = _leaf()
    # Instance without __dict__ is the win; size alone can vary by allocator.
    assert not hasattr(node, "__dict__")
    assert sys.getsizeof(node) < 256


def test_union_node_derived_fields_use_cache() -> None:
    _derived_fields_cache_clear()
    node = make_union_node(["Sheet1!A1", "Sheet1!E5"])
    assert node.key == "Sheet1!A1,E5"
    assert node.shape is NodeShape.union
    with pytest.raises(ValueError, match="column_index"):
        _ = node.column_index
    info = _derived_fields_cache_info()
    assert info.currsize >= 1
