from __future__ import annotations

from excel_grapher.core.address_keys import format_key, parse_address
from excel_grapher.core.addressing import split_sheet_qualified_address
from excel_grapher.core.cell_types import CellKind, CellType, IntIntervalDomain
from excel_grapher.grapher.blank_ranges import address_in_blank_ranges, parse_blank_range_spec
from excel_grapher.grapher.builder import _workbook_sorted_sheet_a1_pairs
from excel_grapher.grapher.dynamic_refs import (
    DynamicRefLimits,
    _ast_address_to_ref_key,
    _sheet_from_addr,
    _split_addr_sheet_coord,
    _split_qualified_to_sheet_a1,
    expand_leaf_env_to_argument_env,
)
from excel_grapher.grapher.parser import parse_cell_refs
from excel_grapher.grapher.validation import _sheet_name_from_key

_APOSTROPHE_SHEET = "O'Neil"
_APOSTROPHE_KEY = "'O''Neil'!A1"


def test_format_key_round_trip_handles_apostrophe_sheet_names() -> None:
    sheet = _APOSTROPHE_SHEET
    a1 = "B2"
    key = format_key(sheet, a1)
    assert key == "'O''Neil'!B2"
    assert parse_address(key) == (sheet, a1)


def test_workbook_sorted_sheet_a1_pairs_handles_apostrophe_sheet_names() -> None:
    pairs = [(_APOSTROPHE_SHEET, "A1"), ("Main", "B2")]
    assert _workbook_sorted_sheet_a1_pairs(pairs, sheet_order=["Main", _APOSTROPHE_SHEET]) == [
        ("Main", "B2"),
        (_APOSTROPHE_SHEET, "A1"),
    ]


def test_parse_blank_range_spec_handles_apostrophe_sheet_names() -> None:
    assert parse_blank_range_spec("'It''s Data'!A1:C2") == ("It's Data", 1, 1, 2, 3)


def test_address_in_blank_ranges_handles_apostrophe_sheet_names() -> None:
    rects = (parse_blank_range_spec("'It''s Data'!A1:B2"),)
    assert address_in_blank_ranges("'It''s Data'!A1", rects)
    assert not address_in_blank_ranges("'It''s Data'!C3", rects)


def test_split_sheet_qualified_address_soft_wrapper_handles_apostrophe_sheet_names() -> None:
    assert split_sheet_qualified_address(_APOSTROPHE_KEY) == (_APOSTROPHE_SHEET, "A1")
    assert split_sheet_qualified_address("A1") is None
    assert split_sheet_qualified_address("'Broken Sheet'") is None


def test_split_addr_sheet_coord_handles_apostrophe_sheet_names() -> None:
    assert _split_addr_sheet_coord(_APOSTROPHE_KEY) == (_APOSTROPHE_SHEET, "A1")


def test_sheet_from_addr_handles_apostrophe_sheet_names() -> None:
    assert _sheet_from_addr(_APOSTROPHE_KEY) == _APOSTROPHE_SHEET


def test_split_qualified_to_sheet_a1_handles_apostrophe_sheet_names() -> None:
    assert _split_qualified_to_sheet_a1(_APOSTROPHE_KEY) == (_APOSTROPHE_SHEET, "A1")


def test_ast_address_to_ref_key_handles_apostrophe_sheet_names() -> None:
    assert _ast_address_to_ref_key(_APOSTROPHE_KEY) == _APOSTROPHE_KEY


def test_sheet_name_from_key_handles_apostrophe_sheet_names() -> None:
    assert _sheet_name_from_key(_APOSTROPHE_KEY) == _APOSTROPHE_SHEET


def test_expand_leaf_env_passes_correct_sheet_for_apostrophe_sheet_names() -> None:
    formula_cell = format_key(_APOSTROPHE_SHEET, "B1")
    leaf_env = {
        f"{_APOSTROPHE_SHEET}!A1": CellType(
            kind=CellKind.NUMBER,
            interval=IntIntervalDomain(min=0, max=10),
        )
    }
    seen_sheets: list[str] = []

    def _get_cell_formula(addr: str) -> str | None:
        if addr == formula_cell:
            return "=A1"
        return None

    def _get_refs_from_formula(formula: str, current_sheet: str) -> set[str]:
        seen_sheets.append(current_sheet)
        return {format_key(_APOSTROPHE_SHEET, "A1")}

    env = expand_leaf_env_to_argument_env(
        {formula_cell},
        _get_cell_formula,
        _get_refs_from_formula,
        leaf_env,
        DynamicRefLimits(),
    )
    assert seen_sheets == [_APOSTROPHE_SHEET]
    assert formula_cell in env


def test_parse_cell_refs_handles_apostrophe_sheet_names() -> None:
    refs = parse_cell_refs("='O''Neil'!A1*2")
    assert len(refs) == 1
    assert refs[0].sheet == _APOSTROPHE_SHEET
    assert refs[0].column == "A"
    assert refs[0].row == 1
