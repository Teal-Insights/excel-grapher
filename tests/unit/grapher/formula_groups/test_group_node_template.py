"""Tests for formula-group template fields on multi-cell nodes."""

from __future__ import annotations

import pickle

import pytest

from excel_grapher.core.formula_ast import (
    AddressHoleNode,
    AddressLeafKind,
    CellRefNode,
    RangeNode,
)
from excel_grapher.grapher.formula_groups import shape_fingerprint
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import (
    NodeKind,
    make_cell_node,
    make_union_node,
    member_keys,
)
from tests.fixtures.formula_groups.option_b import (
    assert_option_b_occupancy,
    build_cross_sheet_cell_only_twin,
    build_cross_sheet_union_option_b,
    build_row_stripe_cell_only_twin,
    build_row_stripe_option_b,
    index_match_fingerprint,
    index_match_skeleton,
)


def test_make_union_node_attaches_validated_template() -> None:
    skeleton = index_match_skeleton()
    fp = shape_fingerprint(skeleton)
    node = make_union_node(
        ["Sheet1!D63", "Sheet1!E63"],
        shape_fingerprint=fp,
        skeleton=skeleton,
        member_bindings={
            "Sheet1!D63": (CellRefNode(address="Sheet1!D35"),),
            "Sheet1!E63": (CellRefNode(address="Sheet1!E35"),),
        },
    )
    assert node.kind is not NodeKind.cell
    assert node.value is None
    assert node.shape_fingerprint == fp
    assert node.skeleton == skeleton
    assert set(node.member_bindings or {}) == {"Sheet1!D63", "Sheet1!E63"}


def test_cell_node_rejects_template_fields() -> None:
    skeleton = index_match_skeleton()
    with pytest.raises(ValueError, match="multi-cell"):
        make_union_node(
            ["Sheet1!D63"],
            shape_fingerprint=shape_fingerprint(skeleton),
            skeleton=skeleton,
            member_bindings={"Sheet1!D63": (CellRefNode(address="Sheet1!D35"),)},
        )


def test_template_requires_all_three_fields() -> None:
    skeleton = index_match_skeleton()
    with pytest.raises(ValueError, match="template requires"):
        make_union_node(
            ["Sheet1!D63", "Sheet1!E63"],
            skeleton=skeleton,
            member_bindings={
                "Sheet1!D63": (CellRefNode(address="Sheet1!D35"),),
                "Sheet1!E63": (CellRefNode(address="Sheet1!E35"),),
            },
        )


def test_template_rejects_missing_member_binding() -> None:
    skeleton = index_match_skeleton()
    with pytest.raises(ValueError, match="Missing member_bindings"):
        make_union_node(
            ["Sheet1!D63", "Sheet1!E63"],
            shape_fingerprint=shape_fingerprint(skeleton),
            skeleton=skeleton,
            member_bindings={"Sheet1!D63": (CellRefNode(address="Sheet1!D35"),)},
        )


def test_template_rejects_kind_mismatch_binding() -> None:
    skeleton = index_match_skeleton()
    with pytest.raises(ValueError, match="Invalid bindings"):
        make_union_node(
            ["Sheet1!D63", "Sheet1!E63"],
            shape_fingerprint=shape_fingerprint(skeleton),
            skeleton=skeleton,
            member_bindings={
                "Sheet1!D63": (RangeNode(start="Sheet1!A1", end="Sheet1!B2"),),
                "Sheet1!E63": (CellRefNode(address="Sheet1!E35"),),
            },
        )


def test_template_rejects_wrong_fingerprint() -> None:
    skeleton = index_match_skeleton()
    with pytest.raises(ValueError, match="shape_fingerprint"):
        make_union_node(
            ["Sheet1!D63", "Sheet1!E63"],
            shape_fingerprint="not-the-fingerprint",
            skeleton=skeleton,
            member_bindings={
                "Sheet1!D63": (CellRefNode(address="Sheet1!D35"),),
                "Sheet1!E63": (CellRefNode(address="Sheet1!E35"),),
            },
        )


def test_cell_nodes_keep_template_fields_none() -> None:
    node = make_cell_node("Sheet1", "A", 1, value=1)
    assert node.shape_fingerprint is None
    assert node.skeleton is None
    assert node.member_bindings is None


def test_row_stripe_option_b_fixture() -> None:
    fx = build_row_stripe_option_b()
    group = fx.graph.get_node(fx.group_key)
    assert group is not None
    assert group.shape_fingerprint == index_match_fingerprint()
    assert group.skeleton is not None
    assert group.member_bindings is not None
    assert set(group.member_bindings) == set(fx.members)
    assert fx.group_key == "Sheet1!D63:F63"
    for m in fx.members:
        assert fx.graph.get_node(m) is None
        assert fx.graph.cell_owner(m) == fx.group_key
    internal = fx.graph._get_internal_node(fx.group_key)
    assert internal is not None
    assert_option_b_occupancy(internal)


def test_cross_sheet_option_b_fixture() -> None:
    fx = build_cross_sheet_union_option_b()
    group = fx.graph.get_node(fx.group_key)
    assert group is not None
    assert group.member_bindings is not None
    assert set(group.member_bindings) == {"Sheet1!D63", "Sheet2!B10"}
    assert "Sheet1!D63" in fx.group_key and "Sheet2!B10" in fx.group_key


def test_cell_only_twins_have_member_cells_not_groups() -> None:
    twin = build_row_stripe_cell_only_twin()
    for key in ("Sheet1!D63", "Sheet1!E63", "Sheet1!F63"):
        node = twin.get_node(key)
        assert node is not None
        assert node.kind is NodeKind.cell
        assert node.normalized_formula is not None

    cross = build_cross_sheet_cell_only_twin()
    assert cross.get_node("Sheet1!D63") is not None
    assert cross.get_node("Sheet2!B10") is not None
    # No multi-cell keys.
    for k in cross:
        node = cross.get_node(k)
        assert node is not None
        assert node.kind is NodeKind.cell


def test_pickle_preserves_template_fields() -> None:
    fx = build_row_stripe_option_b()
    restored: DependencyGraph = pickle.loads(pickle.dumps(fx.graph))
    view = restored.get_node(fx.group_key)
    assert view is not None
    assert view.shape_fingerprint == index_match_fingerprint()
    assert view.skeleton == index_match_skeleton()
    assert view.member_bindings is not None
    assert view.member_bindings["Sheet1!E63"] == (CellRefNode(address="Sheet1!E35"),)
    assert restored.cell_owner("Sheet1!E63") == fx.group_key


def test_copy_for_projection_preserves_template_fields() -> None:
    fx = build_cross_sheet_union_option_b()
    cloned = fx.graph._copy_for_projection()
    view = cloned.get_node(fx.group_key)
    assert view is not None
    assert view.skeleton is not None
    hole = view.skeleton.args[2].args[0]
    assert isinstance(hole, AddressHoleNode)
    assert hole.kind is AddressLeafKind.cell
    assert view.member_bindings is not None
    assert set(view.member_bindings) == set(fx.members)
    orig = fx.graph._get_internal_node(fx.group_key)
    copy_node = cloned._get_internal_node(fx.group_key)
    assert orig is not None and copy_node is not None
    assert member_keys(copy_node) == member_keys(orig)
