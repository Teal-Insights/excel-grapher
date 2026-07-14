"""Unit tests for row nodes in DependencyGraph (issue #374 sprint 3)."""

from __future__ import annotations

import pickle

from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.guard import CellRef, Compare, Literal
from excel_grapher.grapher.node import Node, NodeKind, make_row_node


def _cell(sheet: str, col: str, row: int, *, formula: str | None = None) -> Node:
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=None if formula else 0,
        is_leaf=formula is None,
    )


def test_add_and_get_row_node() -> None:
    g = DependencyGraph()
    row = make_row_node("Sheet1", 63, "D", "Y")
    g.add_node(row)

    assert "Sheet1!D63:Y63" in g
    view = g.get_node("Sheet1!D63:Y63")
    assert view is not None
    assert view.kind is NodeKind.row
    assert view.min_col == "D"
    assert view.max_col == "Y"
    assert view.row == 63
    assert list(g) == ["Sheet1!D63:Y63"]


def test_get_node_normalizes_noncanonical_row_keys() -> None:
    g = DependencyGraph()
    g.add_node(make_row_node("Sheet1", 63, "D", "Y"))

    for lookup in (
        "'Sheet1'!D63:Y63",
        "Sheet1!Y63:D63",
        "Sheet1!D63:Sheet1!Y63",
        "Sheet1!$D$63:$Y$63",
    ):
        assert lookup in g
        view = g.get_node(lookup)
        assert view is not None
        assert view.key == "Sheet1!D63:Y63"


def test_cell_to_row_edge() -> None:
    g = DependencyGraph()
    cell = _cell("Sheet1", "A", 63, formula="=SUM(D63:Y63)")
    row = make_row_node("Sheet1", 63, "D", "Y")
    g.add_node(cell)
    g.add_node(row)
    g.add_edge(cell.key, row.key)

    assert g.get_dependencies(cell.key) == frozenset({row.key})
    assert g.get_dependents(row.key) == frozenset({cell.key})


def test_row_to_cell_edge() -> None:
    g = DependencyGraph()
    row = make_row_node("Sheet1", 63, "D", "Y")
    cell = _cell("Sheet1", "Z", 1)
    g.add_node(row)
    g.add_node(cell)
    g.add_edge(row.key, cell.key)

    assert g.get_dependencies(row.key) == frozenset({cell.key})
    assert g.get_dependents(cell.key) == frozenset({row.key})


def test_row_to_row_edge() -> None:
    g = DependencyGraph()
    left = make_row_node("Sheet1", 63, "D", "Y")
    right = make_row_node("Sheet1", 64, "D", "Y")
    g.add_node(left)
    g.add_node(right)
    g.add_edge(left.key, right.key)

    assert g.get_dependencies(left.key) == frozenset({right.key})
    assert g.get_dependents(right.key) == frozenset({left.key})


def test_edge_attrs_on_cell_to_row() -> None:
    g = DependencyGraph()
    cell = _cell("Sheet1", "A", 63, formula="=IF(C1,SUM(D63:Y63),0)")
    row = make_row_node("Sheet1", 63, "D", "Y")
    flag = _cell("Sheet1", "C", 1)
    g.add_node(cell)
    g.add_node(row)
    g.add_node(flag)

    guard = Compare(CellRef("Sheet1!C1"), "=", Literal(True))
    prov = EdgeProvenance(causes=frozenset({DependencyCause.direct_ref}))
    g.add_edge(cell.key, row.key, guard=guard, provenance=prov)

    attrs = g.get_edge_attrs(cell.key, row.key)
    assert attrs.guard == guard
    assert attrs.provenance is not None
    assert attrs.provenance.causes == frozenset({DependencyCause.direct_ref})
    assert g.get_edge_guard(cell.key, row.key) == guard


def test_pickle_roundtrip_mixed_graph() -> None:
    g = DependencyGraph()
    cell = _cell("Sheet1", "A", 63, formula="=SUM(D63:Y63)")
    row = make_row_node("Sheet1", 63, "D", "Y", metadata={"tag": "inputs"})
    other = make_row_node("Sheet1", 64, "D", "Y")
    g.add_node(cell)
    g.add_node(row)
    g.add_node(other)
    g.add_edge(cell.key, row.key)
    g.add_edge(row.key, other.key)

    restored: DependencyGraph = pickle.loads(pickle.dumps(g))
    view = restored.get_node(row.key)
    assert view is not None
    assert view.kind is NodeKind.row
    assert view.min_col == "D"
    assert view.max_col == "Y"
    assert view.metadata["tag"] == "inputs"
    assert restored.get_dependencies(cell.key) == frozenset({row.key})
    assert restored.get_dependencies(row.key) == frozenset({other.key})
    assert restored.get_dependents(row.key) == frozenset({cell.key})


def test_copy_for_projection_preserves_rows() -> None:
    g = DependencyGraph()
    cell = _cell("Sheet1", "A", 63, formula="=SUM(D63:Y63)")
    row = make_row_node("Sheet1", 63, "D", "Y")
    g.add_node(cell)
    g.add_node(row)
    g.add_edge(cell.key, row.key)

    cloned = g._copy_for_projection()
    view = cloned.get_node(row.key)
    assert view is not None
    assert view.kind is NodeKind.row
    assert view.min_col == "D"
    assert view.max_col == "Y"
    assert cloned.get_dependencies(cell.key) == frozenset({row.key})

    cloned._nodes[row.key].metadata["x"] = 1
    original = g.get_node(row.key)
    assert original is not None
    assert "x" not in original.metadata


def test_workbook_key_order_includes_row_nodes() -> None:
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(_cell("Sheet1", "A", 63))
    g.add_node(make_row_node("Sheet1", 63, "D", "Y"))
    g.add_node(_cell("Sheet1", "B", 64))

    assert g.keys(order="workbook") == [
        "Sheet1!A63",
        "Sheet1!D63:Y63",
        "Sheet1!B64",
    ]
