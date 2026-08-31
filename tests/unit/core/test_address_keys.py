from __future__ import annotations

import re
import typing

import pytest

from excel_grapher.core.address_keys import (
    NormalizedAddress,
    canonical_cell_coord,
    format_cell_key,
    format_key,
    make_node_key_sort_key,
    normalize_key,
    parse_cell_coords,
    quoted_sheet_prefix_regex,
    sort_node_keys,
    split_address_on_colon,
    unescape_formula_sheet_name,
)
from excel_grapher.grapher.node import NodeKey
from excel_grapher.grapher.target_expansion import split_range_target_on_colon


def test_normalized_address_is_str_type_alias() -> None:
    assert NormalizedAddress is str


def test_normalize_key_returns_normalized_address() -> None:
    hints = typing.get_type_hints(normalize_key)
    assert hints["return"] is NormalizedAddress
    result: NormalizedAddress = normalize_key("'Sheet1'!A1")
    assert result == "Sheet1!A1"


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("Sheet1!$A$1", "Sheet1!A1"),
        ("Sheet1!a1", "Sheet1!A1"),
        ("Sheet1!A1:$A$3", "Sheet1!A1:A3"),
        ("Sheet1!$A$1:$B$2", "Sheet1!A1:B2"),
        ("Sheet1!A1:Sheet1!$A$3", "Sheet1!A1:A3"),
        ("Sheet1!a1:b2", "Sheet1!A1:B2"),
        ("Sheet1!$A:$A", "Sheet1!A:A"),
        ("Sheet1!$1:$1", "Sheet1!1:1"),
        ("Sheet1!$A$1:Sheet2!$B$2", "Sheet1!A1:Sheet2!B2"),
    ],
)
def test_normalize_key_canonicalizes_cell_coords(raw: str, expected: str) -> None:
    assert normalize_key(raw) == expected


@pytest.mark.parametrize(
    ("cell", "expected"),
    [
        ("A1", "A1"),
        ("$A$1", "A1"),
        ("a1", "A1"),
        ("$b$10", "B10"),
        ("A", "A"),
        ("$A", "A"),
        ("a", "A"),
        ("1", "1"),
        ("$1", "1"),
        ("01", "1"),
    ],
)
def test_canonical_cell_coord(cell: str, expected: str) -> None:
    assert canonical_cell_coord(cell) == expected


@pytest.mark.parametrize(
    "address",
    [
        "Sheet1!A1:A3",
        "Sheet1!A1:Sheet1!A3",
        "'My Sheet'!A1:B2",
        "'A:B'!A1",
        "'O''Neil'!A1:B2",
        "Sheet1!A1",
    ],
)
def test_split_range_target_on_colon_delegates_to_split_address_on_colon(
    address: str,
) -> None:
    assert split_range_target_on_colon(address) == split_address_on_colon(address)


def test_split_address_on_colon_ignores_colon_inside_quoted_sheet() -> None:
    assert split_address_on_colon("'A:B'!C1") is None
    assert split_address_on_colon("'A:B'!C1:D2") == ("'A:B'!C1", "D2")
    assert split_address_on_colon("'O''Neil'!A1:B2") == ("'O''Neil'!A1", "B2")


def test_format_helpers_return_normalized_address() -> None:
    assert typing.get_type_hints(format_key)["return"] is NormalizedAddress
    assert typing.get_type_hints(format_cell_key)["return"] is NormalizedAddress


def test_node_key_aliases_normalized_address() -> None:
    assert NodeKey is NormalizedAddress


def test_sort_node_keys_respects_workbook_sheet_order_then_row_then_column() -> None:
    sheet_order = ["Inputs", "Calc", "Summary"]
    keys = [
        "Calc!B2",
        "Inputs!C1",
        "Inputs!A1",
        "Calc!A2",
        "Summary!A1",
        "Calc!A1",
    ]

    assert sort_node_keys(keys, sheet_order=sheet_order) == [
        "Inputs!A1",
        "Inputs!C1",
        "Calc!A1",
        "Calc!A2",
        "Calc!B2",
        "Summary!A1",
    ]


def test_sort_node_keys_handles_quoted_sheet_names_in_workbook_order() -> None:
    sheet_order = ["Main", "Data Set", "O'Neil"]
    keys = [
        "'Data Set'!A1",
        "'O''Neil'!A1",
        "Main!A1",
    ]

    assert sort_node_keys(keys, sheet_order=sheet_order) == [
        "Main!A1",
        "'Data Set'!A1",
        "'O''Neil'!A1",
    ]


def test_make_node_key_sort_key_places_unknown_sheets_after_known() -> None:
    key_fn = make_node_key_sort_key(sheet_order=["Known"])
    keys = ["Known!A1", "Other!A1", "Another!A1"]

    assert sorted(keys, key=key_fn) == ["Known!A1", "Another!A1", "Other!A1"]


def test_quoted_sheet_prefix_regex_matches_doubled_apostrophe_escape() -> None:
    pattern = quoted_sheet_prefix_regex()
    match = re.match(pattern + r"A1", "'O''Neil'!A1")
    assert match is not None
    assert unescape_formula_sheet_name(match.group("sheet")) == "O'Neil"


def test_normalize_key_canonicalizes_one_row_ranges() -> None:
    assert normalize_key("Sheet1!Y63:D63") == "Sheet1!D63:Y63"
    assert normalize_key("Sheet1!D63:Sheet1!Y63") == "Sheet1!D63:Y63"
    assert normalize_key("Sheet1!$D$63:$Y$63") == "Sheet1!D63:Y63"
    assert normalize_key("'Sheet1'!D63:Y63") == "Sheet1!D63:Y63"


@pytest.mark.parametrize(
    ("address", "expected"),
    [
        ("Sheet1!A1", ("Sheet1", 1, 1)),
        ("Sheet1!B2", ("Sheet1", 2, 2)),
        ("Sheet1!$C$10", ("Sheet1", 10, 3)),
        ("'My Sheet'!AA3", ("My Sheet", 3, 27)),
        ("'O''Neil'!Z1", ("O'Neil", 1, 26)),
    ],
)
def test_parse_cell_coords_returns_sheet_row_col(
    address: str, expected: tuple[str, int, int]
) -> None:
    assert parse_cell_coords(address) == expected


@pytest.mark.parametrize("address", ["A1", "Sheet1!A1:B2", "Sheet1!A:A"])
def test_parse_cell_coords_rejects_non_cells(address: str) -> None:
    with pytest.raises(ValueError):
        parse_cell_coords(address)


def test_sort_node_keys_orders_row_keys_by_min_col() -> None:
    sheet_order = ["Sheet1"]
    keys = [
        "Sheet1!B64",
        "Sheet1!D63:Y63",
        "Sheet1!A63",
    ]
    assert sort_node_keys(keys, sheet_order=sheet_order) == [
        "Sheet1!A63",
        "Sheet1!D63:Y63",
        "Sheet1!B64",
    ]
