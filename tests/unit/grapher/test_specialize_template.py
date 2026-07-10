"""Tests for row-template specialization (varying cell-ref slots)."""

from __future__ import annotations

import pytest

from excel_grapher.grapher.specialize_template import (
    specialize_template,
    walk_template_cell_refs,
)


def test_walk_skips_range_endpoints() -> None:
    template = "=INDEX($D40:$AJ50,MATCH(1,$AJ40:$AJ50,0),MATCH(D$35,$D$35:$Y$35,0))"
    refs = walk_template_cell_refs(template)
    assert len(refs) == 1
    assert refs[0].column == "D"
    assert refs[0].row == 35
    assert refs[0].is_absolute_row
    assert not refs[0].is_absolute_col


def test_specialize_rewrites_relative_column_preserves_row_abs() -> None:
    template = "=MATCH(D$35,$D$35:$Y$35,0)"
    assert specialize_template(template, varying_ref_slots=(0,), column="E") == (
        "=MATCH(E$35,$D$35:$Y$35,0)"
    )


def test_specialize_preserves_sheet_qualifier() -> None:
    template = "=Sheet1!D$35+$A$1"
    assert specialize_template(template, varying_ref_slots=(0,), column="E") == (
        "=Sheet1!E$35+$A$1"
    )


def test_specialize_preserves_quoted_sheet() -> None:
    template = "='My Sheet'!D$35"
    assert specialize_template(template, varying_ref_slots=(0,), column="E") == ("='My Sheet'!E$35")


def test_specialize_leaves_static_range_unchanged() -> None:
    template = "=SUM($D$35:$Y$35)+D$35"
    out = specialize_template(template, varying_ref_slots=(0,), column="E")
    assert out == "=SUM($D$35:$Y$35)+E$35"


def test_specialize_multiple_slots_right_to_left_safe() -> None:
    template = "=A1+B$2+C3"
    out = specialize_template(template, varying_ref_slots=(0, 2), column="E")
    assert out == "=E1+B$2+E3"


def test_specialize_rejects_absolute_column_varying_slot() -> None:
    template = "=$D$35+A1"
    with pytest.raises(ValueError, match="absolute column"):
        specialize_template(template, varying_ref_slots=(0,), column="E")


def test_specialize_rejects_out_of_range_slot() -> None:
    template = "=D$35"
    with pytest.raises(ValueError, match="out of range"):
        specialize_template(template, varying_ref_slots=(1,), column="E")


def test_specialize_empty_slots_returns_template() -> None:
    template = "=D$35+$A$1"
    assert specialize_template(template, varying_ref_slots=(), column="E") == template


def test_walk_order_matches_specialize_indices() -> None:
    template = "=INDEX($D40:$AJ50,MATCH(1,$AJ40:$AJ50,0),MATCH(D$35,$D$35:$Y$35,0))"
    refs = walk_template_cell_refs(template)
    assert (
        specialize_template(template, varying_ref_slots=(0,), column="E")
        == "=INDEX($D40:$AJ50,MATCH(1,$AJ40:$AJ50,0),MATCH(E$35,$D$35:$Y$35,0))"
    )
    assert refs[0].column == "D"
