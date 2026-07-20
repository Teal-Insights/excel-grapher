"""Builder `formula_groups=` wiring and cell-only coalesce fixtures (Issue 3 sprint 3)."""

from __future__ import annotations

from pathlib import Path

from excel_grapher.grapher.builder import create_dependency_graph
from excel_grapher.grapher.formula_groups import coalesce_formula_groups
from excel_grapher.grapher.node import NodeKind, locate_cell
from tests.fixtures.formula_groups.cell_only import (
    build_cross_sheet_scale_cell_only,
    write_cross_sheet_scale_workbook,
)


def test_cell_only_fixture_coalesces_to_group() -> None:
    fx = build_cross_sheet_scale_cell_only()
    for member in fx.members:
        assert fx.graph.get_node(member) is not None
        assert fx.graph.get_node(member).kind is NodeKind.cell

    report = coalesce_formula_groups(fx.graph)
    assert len(report.created_groups) == 1
    group_key = report.created_groups[0]
    for member in fx.members:
        assert fx.graph.get_node(member) is None
        loc = locate_cell(fx.graph, member)
        assert loc is not None
        assert loc.node_key == group_key
    group = fx.graph.get_node(group_key)
    assert group is not None
    assert group.skeleton is not None
    assert set(group.member_bindings or {}) == set(fx.members)
    assert fx.graph.target_keys() == list(fx.members)


def test_create_dependency_graph_formula_groups_default_off(tmp_path: Path) -> None:
    wb_path = write_cross_sheet_scale_workbook(tmp_path / "scale.xlsx")
    graph = create_dependency_graph(
        wb_path,
        ["Sheet1!B1", "Sheet2!B1"],
        load_values=True,
        use_cached_dynamic_refs=True,
    )
    assert graph.get_node("Sheet1!B1") is not None
    assert graph.get_node("Sheet2!B1") is not None
    assert graph.get_node("Sheet1!B1").kind is NodeKind.cell
    assert graph.get_node("Sheet2!B1").kind is NodeKind.cell


def test_create_dependency_graph_formula_groups_true(tmp_path: Path) -> None:
    wb_path = write_cross_sheet_scale_workbook(tmp_path / "scale.xlsx")
    graph = create_dependency_graph(
        wb_path,
        ["Sheet1!B1", "Sheet2!B1"],
        load_values=True,
        use_cached_dynamic_refs=True,
        formula_groups=True,
    )
    assert graph.get_node("Sheet1!B1") is None
    assert graph.get_node("Sheet2!B1") is None
    loc1 = locate_cell(graph, "Sheet1!B1")
    loc2 = locate_cell(graph, "Sheet2!B1")
    assert loc1 is not None and loc2 is not None
    assert loc1.node_key == loc2.node_key
    group = graph.get_node(loc1.node_key)
    assert group is not None
    assert group.kind is not NodeKind.cell
    assert group.skeleton is not None
    assert graph.target_keys() == ["Sheet1!B1", "Sheet2!B1"]
