"""Sprint 3 unit tests for occupancy, remove rules, mixed edges, and copies."""

from __future__ import annotations

import pickle

import pytest

from excel_grapher.core.address_keys import parse_node_key
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.guard import CellRef, Compare, Literal
from excel_grapher.grapher.node import (
    Node,
    NodeKind,
    locate_cell,
    make_cell_node,
    make_union_node,
)
from excel_grapher.grapher.subgraph import _induced_dependency_subgraph


def _cell(sheet: str, col: str, row: int, *, formula: str | None = None) -> Node:
    return make_cell_node(
        sheet,
        col,
        row,
        formula=formula,
        normalized_formula=formula,
        value=None if formula else 0,
        is_leaf=formula is None,
    )


def test_get_node_normalizes_union_keys() -> None:
    g = DependencyGraph()
    g.add_node(make_union_node(["Sheet1!E5", "Sheet1!A1", "Sheet1!B1", "Sheet1!C1", "Sheet1!D1"]))

    for lookup in (
        "Sheet1!A1:D1,E5",
        "Sheet1!E5,A1:D1",
        "Sheet1!$A$1:$D$1,$E$5",
    ):
        assert lookup in g
        view = g.get_node(lookup)
        assert view is not None
        assert view.key == "Sheet1!A1:D1,E5"
        assert view.kind is NodeKind.union


def test_add_node_rejects_overlapping_occupancy_cell_then_union() -> None:
    g = DependencyGraph()
    g.add_node(_cell("Sheet1", "E", 5))
    with pytest.raises(ValueError, match="occupancy|already|owned"):
        g.add_node(make_union_node(["Sheet1!A1", "Sheet1!E5"]))


def test_add_node_rejects_overlapping_occupancy_union_then_cell() -> None:
    g = DependencyGraph()
    g.add_node(make_union_node(["Sheet1!A1", "Sheet1!E5"]))
    with pytest.raises(ValueError, match="occupancy|already|owned"):
        g.add_node(_cell("Sheet1", "E", 5))


def test_add_node_rejects_overlapping_unions() -> None:
    g = DependencyGraph()
    g.add_node(make_union_node(["Sheet1!A1", "Sheet1!B1", "Sheet1!E5"]))
    with pytest.raises(ValueError, match="occupancy|already|owned"):
        g.add_node(make_union_node(["Sheet1!E5", "Sheet1!Z9"]))


def test_remove_union_clears_occupancy() -> None:
    g = DependencyGraph()
    union = make_union_node(["Sheet1!A1", "Sheet1!B1", "Sheet1!E5"])
    g.add_node(union)
    assert locate_cell(g, "Sheet1!E5") is not None

    g.remove_node(union.key)
    assert union.key not in g
    assert locate_cell(g, "Sheet1!E5") is None
    # Cell becomes free for a new owner.
    g.add_node(_cell("Sheet1", "E", 5))
    loc = locate_cell(g, "Sheet1!E5")
    assert loc is not None
    assert loc.kind is NodeKind.cell


def test_remove_member_cell_while_owned_errors() -> None:
    g = DependencyGraph()
    g.add_node(make_union_node(["Sheet1!A1", "Sheet1!B1", "Sheet1!E5"]))
    with pytest.raises(ValueError, match="member|owned|occupancy"):
        g.remove_node("Sheet1!E5")
    assert "Sheet1!A1:B1,E5" in g
    assert locate_cell(g, "Sheet1!E5") is not None


def test_remove_absent_cell_is_noop() -> None:
    g = DependencyGraph()
    g.add_node(_cell("Sheet1", "A", 1))
    g.remove_node("Sheet1!Z99")
    assert "Sheet1!A1" in g


def test_cell_to_union_edge() -> None:
    g = DependencyGraph()
    cell = _cell("Sheet1", "Z", 1, formula="=A1+E5")
    union = make_union_node(["Sheet1!A1", "Sheet1!E5"])
    g.add_node(cell)
    g.add_node(union)
    g.add_edge(cell.key, union.key)

    assert g.get_dependencies(cell.key) == frozenset({union.key})
    assert g.get_dependents(union.key) == frozenset({cell.key})


def test_union_to_cell_edge() -> None:
    g = DependencyGraph()
    union = make_union_node(["Sheet1!A1", "Sheet1!E5"], is_leaf=False)
    cell = _cell("Sheet1", "Z", 1)
    g.add_node(union)
    g.add_node(cell)
    g.add_edge(union.key, cell.key)

    assert g.get_dependencies(union.key) == frozenset({cell.key})
    assert g.get_dependents(cell.key) == frozenset({union.key})


def test_union_to_union_edge() -> None:
    g = DependencyGraph()
    left = make_union_node(["Sheet1!A1", "Sheet1!B1"])
    right = make_union_node(["Sheet1!E5", "Sheet1!F5"])
    g.add_node(left)
    g.add_node(right)
    g.add_edge(left.key, right.key)

    assert g.get_dependencies(left.key) == frozenset({right.key})
    assert g.get_dependents(right.key) == frozenset({left.key})


def test_edge_attrs_on_cell_to_union() -> None:
    g = DependencyGraph()
    cell = _cell("Sheet1", "Z", 1, formula="=IF(C1,A1,0)")
    union = make_union_node(["Sheet1!A1", "Sheet1!E5"])
    flag = _cell("Sheet1", "C", 1)
    g.add_node(cell)
    g.add_node(union)
    g.add_node(flag)

    guard = Compare(CellRef("Sheet1!C1"), "=", Literal(True))
    prov = EdgeProvenance(causes=DependencyCause.direct_ref)
    g.add_edge(cell.key, union.key, guard=guard, provenance=prov)

    attrs = g.get_edge_attrs(cell.key, union.key)
    assert attrs.guard == guard
    assert attrs.provenance is not None
    assert attrs.provenance.causes == DependencyCause.direct_ref


def test_locate_cell_member_of_union() -> None:
    g = DependencyGraph()
    g.add_node(make_union_node(["Sheet1!A1", "Sheet1!B1", "Sheet1!E5"]))
    loc = locate_cell(g, "Sheet1!E5")
    assert loc is not None
    assert loc.cell_key == "Sheet1!E5"
    assert loc.node_key == "Sheet1!A1:B1,E5"
    assert loc.kind is NodeKind.union
    assert loc.column == "E"


def test_locate_cell_cross_sheet_members_same_owner() -> None:
    g = DependencyGraph()
    union = make_union_node(["Sheet1!A1", "Sheet2!B2"])
    g.add_node(union)
    a = locate_cell(g, "Sheet1!A1")
    b = locate_cell(g, "Sheet2!B2")
    assert a is not None and b is not None
    assert a.node_key == b.node_key == union.key
    assert a.kind is NodeKind.union


def test_locate_cell_uses_occupancy_not_scan_alone() -> None:
    """Occupancy must answer after remove of a covering node."""
    g = DependencyGraph()
    g.add_node(make_union_node(["Sheet1!A1", "Sheet1!E5"]))
    assert g.cell_owner("Sheet1!E5") == "Sheet1!A1,E5"
    g.remove_node("Sheet1!A1,E5")
    assert g.cell_owner("Sheet1!E5") is None


def test_pickle_roundtrip_mixed_union_graph() -> None:
    g = DependencyGraph()
    cell = _cell("Sheet1", "Z", 1, formula="=A1")
    union = make_union_node(
        ["Sheet1!A1", "Sheet1!B1", "Sheet1!E5"],
        metadata={"tag": "group"},
    )
    other = Node(
        sheet=None,
        column=None,
        row=None,
        formula=None,
        normalized_formula=None,
        value=None,
        is_leaf=True,
        address=parse_node_key("Sheet1!D64:Y64"),
    )
    g.add_node(cell)
    g.add_node(union)
    g.add_node(other)
    g.add_edge(cell.key, union.key)
    g.add_edge(union.key, other.key)

    restored: DependencyGraph = pickle.loads(pickle.dumps(g))
    view = restored.get_node(union.key)
    assert view is not None
    assert view.kind is NodeKind.union
    assert view.address == union.address
    assert view.metadata["tag"] == "group"
    assert restored.get_dependencies(cell.key) == frozenset({union.key})
    assert restored.get_dependencies(union.key) == frozenset({other.key})
    assert restored.cell_owner("Sheet1!E5") == union.key
    loc = locate_cell(restored, "Sheet1!E5")
    assert loc is not None
    assert loc.node_key == union.key


def test_copy_for_projection_preserves_union_and_occupancy() -> None:
    g = DependencyGraph()
    cell = _cell("Sheet1", "Z", 1, formula="=A1")
    union = make_union_node(["Sheet1!A1", "Sheet1!E5"])
    g.add_node(cell)
    g.add_node(union)
    g.add_edge(cell.key, union.key)

    cloned = g._copy_for_projection()
    view = cloned.get_node(union.key)
    assert view is not None
    assert view.kind is NodeKind.union
    assert view.address == union.address
    assert cloned.get_dependencies(cell.key) == frozenset({union.key})
    assert cloned.cell_owner("Sheet1!E5") == union.key

    cloned._nodes[union.key].metadata["x"] = 1
    original = g.get_node(union.key)
    assert original is not None
    assert "x" not in original.metadata


def test_induce_subgraph_rebuilds_occupancy() -> None:
    g = DependencyGraph()
    cell = _cell("Sheet1", "Z", 1, formula="=A1")
    union = make_union_node(["Sheet1!A1", "Sheet1!E5"])
    g.add_node(cell)
    g.add_node(union)
    g.add_edge(cell.key, union.key)

    sub = _induced_dependency_subgraph(g, {cell.key, union.key})
    assert sub.cell_owner("Sheet1!A1") == union.key
    assert locate_cell(sub, "Sheet1!E5") is not None


def test_one_row_union_registers_occupancy() -> None:
    g = DependencyGraph()
    g.add_node(
        Node(
            sheet=None,
            column=None,
            row=None,
            formula=None,
            normalized_formula=None,
            value=None,
            is_leaf=True,
            address=parse_node_key("Sheet1!D63:Y63"),
        )
    )
    assert g.cell_owner("Sheet1!E63") == "Sheet1!D63:Y63"
    with pytest.raises(ValueError, match="occupancy|already|owned"):
        g.add_node(_cell("Sheet1", "E", 63))
