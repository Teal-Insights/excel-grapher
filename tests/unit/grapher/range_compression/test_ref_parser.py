"""Tests for raw formula reference parsing with absolute markers."""

from __future__ import annotations

from excel_grapher.grapher.range_compression.ref_parser import parse_cell_refs_with_abs


def test_parse_relative_cell_refs() -> None:
    refs = parse_cell_refs_with_abs("=B3*C3", default_sheet="Patterns")
    assert len(refs) == 2
    assert refs[0].column == "B"
    assert refs[0].row == 3
    assert not refs[0].is_absolute_col
    assert not refs[0].is_absolute_row
    assert refs[1].column == "C"
    assert refs[1].row == 3


def test_parse_absolute_markers() -> None:
    refs = parse_cell_refs_with_abs("=E3+$E$11", default_sheet="Patterns")
    assert len(refs) == 2
    head, tail = refs
    assert head.column == "E" and head.row == 3
    assert not head.is_absolute_col and not head.is_absolute_row
    assert tail.column == "E" and tail.row == 11
    assert tail.is_absolute_col and tail.is_absolute_row


def test_masks_range_endpoints() -> None:
    refs = parse_cell_refs_with_abs("=SUM(B3:C3)", default_sheet="Sheet1")
    assert refs == []
