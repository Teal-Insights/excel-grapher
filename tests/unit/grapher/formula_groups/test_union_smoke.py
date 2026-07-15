"""Sprint 4 smoke tests: scenarios A–D, builder cell-only, and locate perf."""

from __future__ import annotations

import time
from pathlib import Path

import xlsxwriter

from excel_grapher.core.address_keys import NodeShape
from excel_grapher.exporter.projection import OptimalCompression
from excel_grapher.grapher.builder import create_dependency_graph
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import (
    Node,
    NodeKind,
    locate_cell,
    make_cell_node,
    make_union_node,
)
from excel_grapher.grapher.range_compression.grouping import column_adjacent_groups


def _formula_cell(sheet: str, col: str, row: int, formula: str) -> Node:
    return make_cell_node(
        sheet,
        col,
        row,
        formula=formula,
        normalized_formula=formula,
        is_leaf=False,
        is_target=True,
    )


def test_scenario_a_outside_formula_depends_on_union() -> None:
    """Outside cell formula depends on a union; dependents / evaluation_order OK."""
    g = DependencyGraph()
    union = make_union_node(
        ["Sheet1!A1", "Sheet1!B1", "Sheet1!C1", "Sheet1!D1", "Sheet1!E5"],
        is_leaf=True,
    )
    cell = _formula_cell("Sheet1", "Z", 1, "=SUM(A1:D1)+E5")
    g.add_node(union)
    g.add_node(cell)
    g.add_edge(cell.key, union.key)

    assert g.get_dependencies(cell.key) == frozenset({union.key})
    assert g.get_dependents(union.key) == frozenset({cell.key})
    assert g.evaluation_order() == [union.key, cell.key]
    assert g.leaf_keys() == [union.key]
    assert g.formula_keys() == [cell.key]
    assert g.get_node("Sheet1!E5") is None
    loc = locate_cell(g, "Sheet1!E5")
    assert loc is not None
    assert loc.node_key == union.key
    assert loc.kind is NodeKind.union


def test_scenario_b_cross_sheet_members_share_owner() -> None:
    g = DependencyGraph()
    union = make_union_node(["Sheet1!A1", "Sheet2!B2", "Sheet3!C3"])
    g.add_node(union)

    owners = {
        locate_cell(g, "Sheet1!A1"),
        locate_cell(g, "Sheet2!B2"),
        locate_cell(g, "Sheet3!C3"),
    }
    assert None not in owners
    node_keys = {loc.node_key for loc in owners if loc is not None}
    assert node_keys == {union.key}
    for loc in owners:
        assert loc is not None
        assert loc.kind is NodeKind.union


def test_scenario_c_builder_still_emits_only_cells(tmp_path: Path) -> None:
    """create_dependency_graph expands ranges to cells (no multi-cell nodes)."""
    path = tmp_path / "range_sum.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write("B1", 1)
    ws.write("C1", 2)
    ws.write("D1", 3)
    ws.write_formula("A1", "=SUM(B1:D1)", None, 6)
    wb.close()

    graph = create_dependency_graph(path, ["Sheet1!A1"], load_values=True)
    assert set(graph) == {"Sheet1!A1", "Sheet1!B1", "Sheet1!C1", "Sheet1!D1"}
    for key in graph:
        node = graph.get_node(key)
        assert node is not None
        assert node.shape is NodeShape.cell
        assert node.kind is NodeKind.cell
    assert graph.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!B1", "Sheet1!C1", "Sheet1!D1"})


def test_scenario_d_taco_and_optimal_skip_hand_built_union() -> None:
    g = DependencyGraph()
    union = make_union_node(
        ["Sheet1!Z1", "Sheet1!Z2", "Sheet1!Z9"],
        is_leaf=False,
    )
    a1 = _formula_cell("Sheet1", "A", 1, "=1")
    a2 = _formula_cell("Sheet1", "A", 2, "=1")
    g.add_node(union)
    g.add_node(a1)
    g.add_node(a2)

    groups = column_adjacent_groups(g, min_len=2)
    flat = {key for group in groups for key in group}
    assert union.key not in flat
    assert flat == {a1.key, a2.key}

    projection = OptimalCompression().project(g)
    assert union.key in projection.projected_graph
    assert projection.projected_graph.get_node(union.key) is not None
    assert projection.projected_graph.cell_owner("Sheet1!Z9") == union.key


def test_perf_smoke_locate_cell_1000_members() -> None:
    members = [f"Sheet1!A{i}" for i in range(1, 1001)]
    union = make_union_node(members)
    g = DependencyGraph()
    g.add_node(union)

    assert union.shape is NodeShape.column
    assert g.cell_owner("Sheet1!A500") == union.key

    t0 = time.perf_counter()
    for _ in range(2000):
        loc = locate_cell(g, "Sheet1!A500")
        assert loc is not None
        assert loc.node_key == union.key
    elapsed = time.perf_counter() - t0
    # Occupancy index should keep this well under a scan of 1000 members.
    assert elapsed < 0.5, f"locate_cell too slow: {elapsed:.3f}s"


def test_get_node_vs_locate_cell_lookup_story() -> None:
    """get_node is exact-key only; members must use locate_cell."""
    g = DependencyGraph()
    union = make_union_node(["Sheet1!A1", "Sheet1!E5"])
    g.add_node(union)

    assert g.get_node(union.key) is not None
    assert g.get_node("Sheet1!E5") is None
    loc = locate_cell(g, "Sheet1!E5")
    assert loc is not None
    assert loc.node_key == union.key
    owned = g.get_node(loc.node_key)
    assert owned is not None
    assert owned.key == union.key
