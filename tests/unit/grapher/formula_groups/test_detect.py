"""Unit tests for formula-family discovery and group template building (Issue 3 sprint 1)."""

from __future__ import annotations

import pytest

from excel_grapher.core.formula_ast import (
    AddressHoleNode,
    AddressLeafKind,
    BinaryOpNode,
    CellRefNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
)
from excel_grapher.evaluator.parser import parse
from excel_grapher.grapher.formula_groups import (
    ReadyFamily,
    SkippedFamily,
    TemplateBuildError,
    build_group_template,
    iter_formula_families,
    shape_fingerprint,
    specialize_group,
)
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import make_cell_node, make_union_node


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


def test_build_group_template_holes_differing_cell_refs() -> None:
    members = ("Sheet1!B1", "Sheet1!B2")
    formulas = {
        "Sheet1!B1": "=Sheet1!A1+1",
        "Sheet1!B2": "=Sheet1!A2+1",
    }
    template = build_group_template(members, formulas)
    assert template.shape_fingerprint == shape_fingerprint(parse("=Sheet1!A1+1"))
    assert isinstance(template.skeleton, BinaryOpNode)
    assert isinstance(template.skeleton.left, AddressHoleNode)
    assert template.skeleton.left.kind is AddressLeafKind.cell
    assert isinstance(template.skeleton.right, NumberNode)
    assert template.member_bindings["Sheet1!B1"] == (CellRefNode(address="Sheet1!A1"),)
    assert template.member_bindings["Sheet1!B2"] == (CellRefNode(address="Sheet1!A2"),)
    for member in members:
        specialized = specialize_group(template.skeleton, template.member_bindings[member])
        assert specialized == parse(formulas[member])


def test_build_group_template_bakes_shared_range_refs() -> None:
    members = ("Sheet1!D63", "Sheet2!B10")
    formulas = {
        "Sheet1!D63": (
            "=INDEX(Sheet1!D40:AJ50,MATCH(1,Sheet1!AJ40:AJ50,0),MATCH(Sheet1!D35,Sheet1!D35:Y35,0))"
        ),
        "Sheet2!B10": (
            "=INDEX(Sheet1!D40:AJ50,MATCH(1,Sheet1!AJ40:AJ50,0),MATCH(Sheet2!Z9,Sheet1!D35:Y35,0))"
        ),
    }
    template = build_group_template(members, formulas)
    # Only the differing MATCH lookup cell is a hole; ranges bake.
    assert isinstance(template.skeleton, FunctionCallNode)
    assert template.skeleton.name.upper() == "INDEX"
    col_match = template.skeleton.args[2]
    assert isinstance(col_match, FunctionCallNode)
    assert isinstance(col_match.args[0], AddressHoleNode)
    assert col_match.args[0].kind is AddressLeafKind.cell
    assert isinstance(col_match.args[1], RangeNode)
    assert template.member_bindings["Sheet1!D63"] == (CellRefNode(address="Sheet1!D35"),)
    assert template.member_bindings["Sheet2!B10"] == (CellRefNode(address="Sheet2!Z9"),)


def test_build_group_template_rejects_divergent_fingerprints() -> None:
    """Different fingerprints must not be passed to build_group_template."""
    formulas = {
        "Sheet1!B1": "=Sheet1!A1+1",
        "Sheet1!B2": "=Sheet1!A1+2",
    }
    assert shape_fingerprint(parse(formulas["Sheet1!B1"])) != shape_fingerprint(
        parse(formulas["Sheet1!B2"])
    )
    with pytest.raises(TemplateBuildError, match="shape fingerprint"):
        build_group_template(("Sheet1!B1", "Sheet1!B2"), formulas)


def test_iter_clusters_same_shape_with_cell_hole() -> None:
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(make_cell_node("Sheet1", "A", 1, value=1.0, is_leaf=True))
    g.add_node(make_cell_node("Sheet1", "A", 2, value=2.0, is_leaf=True))
    _add_formula(g, "Sheet1!B1", "=Sheet1!A1+1")
    _add_formula(g, "Sheet1!B2", "=Sheet1!A2+1")
    g.add_edge("Sheet1!B1", "Sheet1!A1")
    g.add_edge("Sheet1!B2", "Sheet1!A2")

    results = list(iter_formula_families(g))
    ready = [r for r in results if isinstance(r, ReadyFamily)]
    skipped = [r for r in results if isinstance(r, SkippedFamily)]
    assert skipped == []
    assert len(ready) == 1
    family = ready[0]
    assert family.members == ("Sheet1!B1", "Sheet1!B2")
    assert len(family.member_bindings["Sheet1!B1"]) == 1
    assert family.member_bindings["Sheet1!B1"][0] == CellRefNode(address="Sheet1!A1")


def test_iter_does_not_cluster_different_literals() -> None:
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(make_cell_node("Sheet1", "A", 1, value=1.0, is_leaf=True))
    _add_formula(g, "Sheet1!B1", "=Sheet1!A1+1")
    _add_formula(g, "Sheet1!B2", "=Sheet1!A1+2")
    g.add_edge("Sheet1!B1", "Sheet1!A1")
    g.add_edge("Sheet1!B2", "Sheet1!A1")

    results = list(iter_formula_families(g))
    ready = [r for r in results if isinstance(r, ReadyFamily)]
    skipped = [r for r in results if isinstance(r, SkippedFamily)]
    assert ready == []
    assert {s.reason for s in skipped} == {"below_min_size"}
    assert all(len(s.members) == 1 for s in skipped)


def test_iter_allows_cross_sheet_noncontiguous_family() -> None:
    g = DependencyGraph()
    g.sheet_order = ["Sheet1", "Sheet2"]
    g.add_node(make_cell_node("Sheet1", "D", 35, value="D", is_leaf=True))
    g.add_node(make_cell_node("Sheet2", "Z", 9, value="D", is_leaf=True))
    formula_d = (
        "=INDEX(Sheet1!D40:AJ50,MATCH(1,Sheet1!AJ40:AJ50,0),MATCH(Sheet1!D35,Sheet1!D35:Y35,0))"
    )
    formula_b = (
        "=INDEX(Sheet1!D40:AJ50,MATCH(1,Sheet1!AJ40:AJ50,0),MATCH(Sheet2!Z9,Sheet1!D35:Y35,0))"
    )
    _add_formula(g, "Sheet1!D63", formula_d)
    _add_formula(g, "Sheet2!B10", formula_b)
    g.add_edge("Sheet1!D63", "Sheet1!D35")
    g.add_edge("Sheet2!B10", "Sheet2!Z9")

    results = list(iter_formula_families(g))
    ready = [r for r in results if isinstance(r, ReadyFamily)]
    assert len(ready) == 1
    assert ready[0].members == ("Sheet1!D63", "Sheet2!B10")


def test_iter_skips_intra_family_edge() -> None:
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(make_cell_node("Sheet1", "A", 1, value=1.0, is_leaf=True))
    _add_formula(g, "Sheet1!B1", "=Sheet1!A1+1")
    _add_formula(g, "Sheet1!B2", "=Sheet1!B1+1")
    g.add_edge("Sheet1!B1", "Sheet1!A1")
    g.add_edge("Sheet1!B2", "Sheet1!B1")  # intra-family

    results = list(iter_formula_families(g))
    ready = [r for r in results if isinstance(r, ReadyFamily)]
    skipped = [r for r in results if isinstance(r, SkippedFamily)]
    assert ready == []
    assert len(skipped) == 1
    assert skipped[0].reason == "intra_family_edge"
    assert set(skipped[0].members) == {"Sheet1!B1", "Sheet1!B2"}


def test_iter_ignores_existing_multi_cell_groups() -> None:
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(make_cell_node("Sheet1", "Z", 1, value=1.0, is_leaf=True))
    g.add_node(make_cell_node("Sheet1", "Z", 2, value=2.0, is_leaf=True))
    skeleton = AddressHoleNode(kind=AddressLeafKind.cell, slot=0)
    fp = shape_fingerprint(skeleton)
    g.add_node(
        make_union_node(
            ["Sheet1!A1", "Sheet1!B1"],
            is_leaf=False,
            shape_fingerprint=fp,
            skeleton=skeleton,
            member_bindings={
                "Sheet1!A1": (CellRefNode(address="Sheet1!Z1"),),
                "Sheet1!B1": (CellRefNode(address="Sheet1!Z2"),),
            },
        )
    )
    # Also a lone formula cell — below min size alone.
    _add_formula(g, "Sheet1!C1", "=Sheet1!Z1+1")
    g.add_edge("Sheet1!C1", "Sheet1!Z1")

    results = list(iter_formula_families(g))
    ready = [r for r in results if isinstance(r, ReadyFamily)]
    assert ready == []
    # Existing group must not appear as a family member candidate.
    for item in results:
        assert "Sheet1!A1" not in item.members
        assert "Sheet1!B1" not in item.members


def test_iter_omits_unparseable_from_clusters() -> None:
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(make_cell_node("Sheet1", "A", 1, value=1.0, is_leaf=True))
    _add_formula(g, "Sheet1!B1", "=Sheet1!A1+1")
    # Intentionally broken formula text; still a formula cell node.
    g.add_node(
        make_cell_node(
            "Sheet1",
            "B",
            2,
            formula="=!!!",
            normalized_formula="=!!!",
            is_leaf=False,
        )
    )
    g.add_edge("Sheet1!B1", "Sheet1!A1")

    results = list(iter_formula_families(g))
    ready = [r for r in results if isinstance(r, ReadyFamily)]
    assert ready == []
    # Unparseable cell is not clustered with B1; B1 alone is below_min_size.
    skipped = [r for r in results if isinstance(r, SkippedFamily)]
    assert any(s.reason == "below_min_size" and s.members == ("Sheet1!B1",) for s in skipped)
    assert all("Sheet1!B2" not in s.members for s in skipped)


def test_iter_families_are_deterministic() -> None:
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(make_cell_node("Sheet1", "A", 1, value=1.0, is_leaf=True))
    g.add_node(make_cell_node("Sheet1", "A", 2, value=2.0, is_leaf=True))
    g.add_node(make_cell_node("Sheet1", "C", 1, value=3.0, is_leaf=True))
    g.add_node(make_cell_node("Sheet1", "C", 2, value=4.0, is_leaf=True))
    _add_formula(g, "Sheet1!B2", "=Sheet1!A2+1")
    _add_formula(g, "Sheet1!B1", "=Sheet1!A1+1")
    _add_formula(g, "Sheet1!D2", "=Sheet1!C2*2")
    _add_formula(g, "Sheet1!D1", "=Sheet1!C1*2")
    for a, b in (
        ("Sheet1!B1", "Sheet1!A1"),
        ("Sheet1!B2", "Sheet1!A2"),
        ("Sheet1!D1", "Sheet1!C1"),
        ("Sheet1!D2", "Sheet1!C2"),
    ):
        g.add_edge(a, b)

    first = list(iter_formula_families(g))
    second = list(iter_formula_families(g))
    assert first == second
    ready = [r for r in first if isinstance(r, ReadyFamily)]
    assert len(ready) == 2
    # Fingerprints sorted lexicographically; members workbook-ordered within family.
    assert [r.members for r in ready] == [
        ("Sheet1!D1", "Sheet1!D2"),  # O:*(...) before O:+(...)
        ("Sheet1!B1", "Sheet1!B2"),
    ]
