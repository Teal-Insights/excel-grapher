"""Unit tests for the range-/element-aware guard representation (issue #483).

Array-context `IF` evaluates its condition element-wise, so a condition over a
range is a *template*: `RangeRef` stands for "the element aligned with the value
element being guarded", and `instantiate_element_guard` resolves it per element.
"""

from __future__ import annotations

from excel_grapher.grapher.guard import (
    And,
    CellRef,
    Compare,
    Literal,
    Not,
    Or,
    RangeRef,
    guard_range_shape,
    instantiate_element_guard,
)
from excel_grapher.grapher.parser import element_aligned_range_cells, parse_guard_expr


def test_range_ref_reports_shape_and_resolves_elements() -> None:
    ref = RangeRef("Sheet1!A1:A10")
    assert ref.shape == (10, 1)
    assert ref.element(0, 0) == CellRef("Sheet1!A1")
    assert ref.element(2, 0) == CellRef("Sheet1!A3")


def test_range_ref_resolves_two_dimensional_elements() -> None:
    ref = RangeRef("Sheet1!B2:D4")
    assert ref.shape == (3, 3)
    assert ref.element(1, 2) == CellRef("Sheet1!D3")


def test_range_ref_str_is_the_range_address() -> None:
    assert str(RangeRef("'My Sheet'!A1:A3")) == "'My Sheet'!A1:A3"


def test_guard_range_shape_returns_uniform_shape() -> None:
    guard = Compare(RangeRef("Sheet1!A1:A10"), "=", RangeRef("Sheet1!C1:C10"))
    assert guard_range_shape(guard) == (10, 1)


def test_guard_range_shape_is_none_without_ranges() -> None:
    assert guard_range_shape(Compare(CellRef("Sheet1!A1"), ">", Literal(0))) is None


def test_guard_range_shape_is_none_for_mixed_shapes() -> None:
    guard = Compare(RangeRef("Sheet1!A1:A10"), "=", RangeRef("Sheet1!C1:C5"))
    assert guard_range_shape(guard) is None


def test_instantiate_element_guard_substitutes_the_aligned_element() -> None:
    guard = Compare(RangeRef("Sheet1!A1:A10"), ">", Literal(0))
    assert instantiate_element_guard(guard, row_offset=2, col_offset=0) == Compare(
        CellRef("Sheet1!A3"), ">", Literal(0)
    )


def test_instantiate_element_guard_keeps_scalar_operands() -> None:
    guard = Not(Compare(RangeRef("Sheet1!A1:A4"), "=", CellRef("Sheet1!C1")))
    assert instantiate_element_guard(guard, row_offset=1, col_offset=0) == Not(
        Compare(CellRef("Sheet1!A2"), "=", CellRef("Sheet1!C1"))
    )


def test_instantiate_element_guard_returns_none_when_offset_is_out_of_range() -> None:
    guard = Compare(RangeRef("Sheet1!A1:A4"), ">", Literal(0))
    assert instantiate_element_guard(guard, row_offset=9, col_offset=0) is None


def test_parse_guard_expr_rejects_ranges_by_default() -> None:
    assert parse_guard_expr("A1:A10>0", current_sheet="Sheet1") is None


def test_parse_guard_expr_builds_range_refs_when_ranges_are_allowed() -> None:
    guard = parse_guard_expr("A1:A10>0", current_sheet="Sheet1", allow_ranges=True)
    assert guard == Compare(RangeRef("Sheet1!A1:A10"), ">", Literal(0))


def test_parse_guard_expr_allows_ranges_under_not() -> None:
    guard = parse_guard_expr("NOT(A1:A3=C1)", current_sheet="Sheet1", allow_ranges=True)
    assert guard == Not(Compare(RangeRef("Sheet1!A1:A3"), "=", CellRef("Sheet1!C1")))


def test_parse_guard_expr_rejects_ranges_under_aggregating_logicals() -> None:
    # AND/OR collapse an array to a single boolean; element-wise instantiation of
    # their operands would not describe Excel's evaluation.
    assert parse_guard_expr("AND(A1:A3>0,C1>0)", current_sheet="Sheet1", allow_ranges=True) is None
    assert parse_guard_expr("OR(A1:A3>0)", current_sheet="Sheet1", allow_ranges=True) is None


def test_parse_guard_expr_still_parses_scalar_logicals_when_ranges_are_allowed() -> None:
    guard = parse_guard_expr("AND(A1>0,C1>0)", current_sheet="Sheet1", allow_ranges=True)
    assert guard == And(
        (
            Compare(CellRef("Sheet1!A1"), ">", Literal(0)),
            Compare(CellRef("Sheet1!C1"), ">", Literal(0)),
        )
    )


def test_parse_guard_expr_rejects_aggregated_ranges() -> None:
    assert parse_guard_expr("SUM(A1:A3)>0", current_sheet="Sheet1", allow_ranges=True) is None


def test_or_guard_of_element_guards_is_unaffected_by_range_refs() -> None:
    # Sanity: instantiated guards stay ordinary scalar guards downstream.
    left = Compare(CellRef("Sheet1!A1"), ">", Literal(0))
    right = Compare(CellRef("Sheet1!A2"), ">", Literal(0))
    assert guard_range_shape(Or((left, right))) is None


def test_element_alignment_maps_range_cells_to_offsets() -> None:
    mapping = element_aligned_range_cells(
        "B1:B4", current_sheet="Sheet1", shape=(4, 1), max_cells=1000
    )
    assert mapping[("Sheet1", "B3")] == (2, 0)
    assert len(mapping) == 4


def test_element_alignment_skips_ranges_inside_aggregating_calls() -> None:
    assert (
        element_aligned_range_cells(
            "SUM(B1:B4)", current_sheet="Sheet1", shape=(4, 1), max_cells=1000
        )
        == {}
    )


def test_element_alignment_allows_ranges_inside_element_wise_calls() -> None:
    mapping = element_aligned_range_cells(
        "IF(B1:B4>0,C1:C4,0)", current_sheet="Sheet1", shape=(4, 1), max_cells=1000
    )
    assert mapping[("Sheet1", "B2")] == (1, 0)
    assert mapping[("Sheet1", "C2")] == (1, 0)


def test_element_alignment_skips_shape_mismatches() -> None:
    assert (
        element_aligned_range_cells("B1:B5", current_sheet="Sheet1", shape=(4, 1), max_cells=1000)
        == {}
    )


def test_element_alignment_skips_ranges_inside_dynamic_refs() -> None:
    assert (
        element_aligned_range_cells(
            "OFFSET(B1:B4,0,0)", current_sheet="Sheet1", shape=(4, 1), max_cells=1000
        )
        == {}
    )


def test_element_alignment_drops_cells_with_ambiguous_offsets() -> None:
    # B2 is element 1 of B1:B4 and element 0 of B2:B5 — no single aligned index.
    mapping = element_aligned_range_cells(
        "B1:B4+B2:B5", current_sheet="Sheet1", shape=(4, 1), max_cells=1000
    )
    assert ("Sheet1", "B2") not in mapping
    assert mapping[("Sheet1", "B1")] == (0, 0)
    assert mapping[("Sheet1", "B5")] == (3, 0)


def test_element_alignment_skips_ranges_over_the_cell_budget() -> None:
    assert (
        element_aligned_range_cells("B1:B4", current_sheet="Sheet1", shape=(4, 1), max_cells=2)
        == {}
    )
