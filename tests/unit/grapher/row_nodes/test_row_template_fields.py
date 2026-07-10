"""Tests for row-node template fields and Option B fixtures (issue #377 sprint 2)."""

from __future__ import annotations

import pytest

from excel_grapher.grapher.node import NodeKind, make_row_node, node_to_view
from excel_grapher.grapher.specialize_template import specialize_template
from tests.fixtures.row_nodes.option_b_stripe import (
    OPTION_B_ROW_KEY,
    OPTION_B_TEMPLATE,
    OPTION_B_VARYING_REF_SLOTS,
    assert_unique_occupancy_for_row,
    build_cell_only_product_twin,
    build_option_b_product_graph,
    build_option_b_stripe_fixture,
)


def test_make_row_node_stores_varying_ref_slots() -> None:
    node = make_row_node(
        "Sheet1",
        63,
        "D",
        "E",
        formula=OPTION_B_TEMPLATE,
        normalized_formula=OPTION_B_TEMPLATE,
        varying_ref_slots=OPTION_B_VARYING_REF_SLOTS,
    )
    assert node.varying_ref_slots == (0,)
    view = node_to_view(node)
    assert view.varying_ref_slots == (0,)
    assert view.normalized_formula == OPTION_B_TEMPLATE


def test_make_row_node_rejects_absolute_column_varying_slot() -> None:
    with pytest.raises(ValueError, match="absolute column"):
        make_row_node(
            "Sheet1",
            63,
            "D",
            "E",
            formula="=$D$35*2",
            normalized_formula="=$D$35*2",
            varying_ref_slots=(0,),
        )


def test_make_row_node_rejects_slot_without_template() -> None:
    with pytest.raises(ValueError, match="normalized_formula"):
        make_row_node(
            "Sheet1",
            63,
            "D",
            "E",
            varying_ref_slots=(0,),
        )


def test_make_row_node_rejects_out_of_range_slot() -> None:
    with pytest.raises(ValueError, match="out of range"):
        make_row_node(
            "Sheet1",
            63,
            "D",
            "E",
            formula="=1+1",
            normalized_formula="=1+1",
            varying_ref_slots=(0,),
        )


def test_option_b_fixture_has_row_template_and_no_member_cells() -> None:
    g = build_option_b_product_graph()
    row = g.get_node(OPTION_B_ROW_KEY)
    assert row is not None
    assert row.kind is NodeKind.row
    assert row.normalized_formula == OPTION_B_TEMPLATE
    assert row.varying_ref_slots == OPTION_B_VARYING_REF_SLOTS
    assert g.get_node("Sheet1!D63") is None
    assert g.get_node("Sheet1!E63") is None
    assert_unique_occupancy_for_row(g, OPTION_B_ROW_KEY)


def test_option_b_fixture_specializes_like_cell_twin_formulas() -> None:
    fixture = build_option_b_stripe_fixture()
    specialized_d = specialize_template(
        fixture.template,
        varying_ref_slots=fixture.varying_ref_slots,
        column="D",
    )
    specialized_e = specialize_template(
        fixture.template,
        varying_ref_slots=fixture.varying_ref_slots,
        column="E",
    )
    twin = build_cell_only_product_twin()
    d63 = twin.get_node("Sheet1!D63")
    e63 = twin.get_node("Sheet1!E63")
    assert d63 is not None and e63 is not None
    assert specialized_d == d63.normalized_formula
    assert specialized_e == e63.normalized_formula


def test_cell_only_twin_has_member_cells_not_row_node() -> None:
    twin = build_cell_only_product_twin()
    assert twin.get_node(OPTION_B_ROW_KEY) is None
    assert twin.get_node("Sheet1!D63") is not None
    assert twin.get_node("Sheet1!E63") is not None
    assert twin.get_node("Sheet1!D63").kind is NodeKind.cell  # type: ignore[union-attr]
