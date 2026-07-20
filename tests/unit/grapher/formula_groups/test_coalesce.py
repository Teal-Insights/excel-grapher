"""Unit tests for coalesce_formula_groups (Issue 3 sprint 2)."""

from __future__ import annotations

from excel_grapher.core.address_keys import members_to_node_key
from excel_grapher.grapher.formula_groups import (
    TARGET_MEMBERS_METADATA_KEY,
    coalesce_formula_groups,
)
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import NodeKind, locate_cell, make_cell_node


def _add_formula(
    g: DependencyGraph,
    address: str,
    formula: str,
    *,
    is_target: bool = False,
) -> None:
    sheet, cell = address.split("!", 1)
    col = "".join(ch for ch in cell if ch.isalpha())
    row = int("".join(ch for ch in cell if ch.isdigit()))
    g.add_node(
        make_cell_node(
            sheet,
            col,
            row,
            formula=formula,
            normalized_formula=formula,
            is_leaf=False,
            is_target=is_target,
        )
    )


def _scale_pair_graph(*, b1_target: bool = False, b2_target: bool = False) -> DependencyGraph:
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(make_cell_node("Sheet1", "A", 1, value=1.0, is_leaf=True))
    g.add_node(make_cell_node("Sheet1", "A", 2, value=2.0, is_leaf=True))
    _add_formula(g, "Sheet1!B1", "=Sheet1!A1*10", is_target=b1_target)
    _add_formula(g, "Sheet1!B2", "=Sheet1!A2*10", is_target=b2_target)
    g.add_edge("Sheet1!B1", "Sheet1!A1")
    g.add_edge("Sheet1!B2", "Sheet1!A2")
    return g


def test_coalesce_replaces_members_with_group() -> None:
    g = _scale_pair_graph()
    report = coalesce_formula_groups(g)
    assert len(report.created_groups) == 1
    group_key = report.created_groups[0]
    assert group_key == members_to_node_key(["Sheet1!B1", "Sheet1!B2"])
    assert str(group_key) == "Sheet1!B1:B2"

    assert g.get_node("Sheet1!B1") is None
    assert g.get_node("Sheet1!B2") is None
    group = g.get_node(group_key)
    assert group is not None
    assert group.kind is not NodeKind.cell
    assert group.skeleton is not None
    assert group.member_bindings is not None
    assert set(group.member_bindings) == {"Sheet1!B1", "Sheet1!B2"}
    assert group.value is None

    for member in ("Sheet1!B1", "Sheet1!B2"):
        loc = locate_cell(g, member)
        assert loc is not None
        assert loc.node_key == group_key
        assert g.cell_owner(member) == group_key


def test_coalesce_rewrites_inbound_and_outbound_edges() -> None:
    g = _scale_pair_graph()
    _add_formula(g, "Sheet1!C1", "=Sheet1!B1+Sheet1!B2")
    g.add_edge("Sheet1!C1", "Sheet1!B1")
    g.add_edge("Sheet1!C1", "Sheet1!B2")

    report = coalesce_formula_groups(g)
    group_key = report.created_groups[0]

    # Outbound: group depends on A1 and A2
    deps = g.get_dependencies(group_key)
    assert deps == frozenset({"Sheet1!A1", "Sheet1!A2"})

    # Inbound: C1 depends on the group (not the removed members)
    assert g.get_dependencies("Sheet1!C1") == frozenset({group_key})
    assert "Sheet1!C1" in g.get_dependents(group_key)

    # Dependent formula text unchanged
    c1 = g.get_node("Sheet1!C1")
    assert c1 is not None
    assert c1.normalized_formula == "=Sheet1!B1+Sheet1!B2"


def test_coalesce_preserves_target_members_in_target_keys() -> None:
    g = _scale_pair_graph(b1_target=True, b2_target=False)
    assert g.target_keys() == ["Sheet1!B1"]

    report = coalesce_formula_groups(g)
    group_key = report.created_groups[0]
    group = g.get_node(group_key)
    assert group is not None
    assert group.is_target is True
    assert group.metadata.get(TARGET_MEMBERS_METADATA_KEY) == ("Sheet1!B1",)
    # Public targets remain member addresses, not the group key.
    assert g.target_keys() == ["Sheet1!B1"]
    assert group_key not in g.target_keys()


def test_coalesce_second_pass_is_idempotent() -> None:
    g = _scale_pair_graph()
    first = coalesce_formula_groups(g)
    assert len(first.created_groups) == 1
    second = coalesce_formula_groups(g)
    assert second.created_groups == ()
    assert g.get_node(first.created_groups[0]) is not None
    assert g.get_node("Sheet1!B1") is None


def test_coalesce_reports_skipped_intra_family() -> None:
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(make_cell_node("Sheet1", "A", 1, value=1.0, is_leaf=True))
    _add_formula(g, "Sheet1!B1", "=Sheet1!A1+1")
    _add_formula(g, "Sheet1!B2", "=Sheet1!B1+1")
    g.add_edge("Sheet1!B1", "Sheet1!A1")
    g.add_edge("Sheet1!B2", "Sheet1!B1")

    report = coalesce_formula_groups(g)
    assert report.created_groups == ()
    assert len(report.skipped_families) == 1
    assert report.skipped_families[0].reason == "intra_family_edge"
    assert g.get_node("Sheet1!B1") is not None
    assert g.get_node("Sheet1!B2") is not None


def test_coalesce_cross_sheet_family() -> None:
    g = DependencyGraph()
    g.sheet_order = ["Sheet1", "Sheet2"]
    g.add_node(make_cell_node("Sheet1", "A", 1, value=1.0, is_leaf=True))
    g.add_node(make_cell_node("Sheet2", "A", 1, value=2.0, is_leaf=True))
    _add_formula(g, "Sheet1!B1", "=Sheet1!A1*10")
    _add_formula(g, "Sheet2!B1", "=Sheet2!A1*10")
    g.add_edge("Sheet1!B1", "Sheet1!A1")
    g.add_edge("Sheet2!B1", "Sheet2!A1")

    report = coalesce_formula_groups(g)
    assert len(report.created_groups) == 1
    group_key = report.created_groups[0]
    assert str(group_key) == members_to_node_key(["Sheet1!B1", "Sheet2!B1"])
    assert locate_cell(g, "Sheet1!B1") is not None
    assert locate_cell(g, "Sheet2!B1") is not None
    assert locate_cell(g, "Sheet1!B1").node_key == group_key
    assert locate_cell(g, "Sheet2!B1").node_key == group_key
