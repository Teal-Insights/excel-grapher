"""Orientation helper: row ↔ column and `column_header` ↔ `row_label`."""

from __future__ import annotations

from pathlib import Path

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    oriented_addresses,
    oriented_document,
    series_entry,
    transpose_address,
    transpose_bindings,
    transpose_cell_coord,
    transpose_formula,
    transpose_sheets,
    write_oriented_workbook,
)


def test_transpose_cell_coord_swaps_axes() -> None:
    assert transpose_cell_coord("A1") == "A1"
    assert transpose_cell_coord("B1") == "A2"
    assert transpose_cell_coord("A2") == "B1"
    assert transpose_cell_coord("C2") == "B3"
    assert transpose_cell_coord("$B$1") == "$A$2"
    assert transpose_cell_coord("$A2") == "B$1"


def test_transpose_address_orders_range_endpoints() -> None:
    assert transpose_address("Engine!A2:C2") == "Engine!B1:B3"
    assert transpose_address("Engine!B1:B3") == "Engine!A2:C2"
    assert transpose_address("C2:A2") == "B1:B3"


def test_transpose_formula_rewrites_refs_not_function_names() -> None:
    assert transpose_formula("=A2+B3") == "=B1+C2"
    assert transpose_formula("=LOG10(A1)") == "=LOG10(A1)"
    assert transpose_formula('="A2"&B1') == '="A2"&A2'
    assert transpose_formula("=Engine!C2*0.9") == "=Engine!B3*0.9"
    assert transpose_formula("=INDEX(Inputs!$A$10:$C$12,1,2)") == ("=INDEX(Inputs!$J$1:$L$3,1,2)")


def test_transpose_sheets_moves_values_and_formulas() -> None:
    sheets = {
        "Engine": {
            "A1": 2009,
            "B1": 2010,
            "A2": "=100",
            "B2": "=A2+1",
        }
    }
    assert transpose_sheets(sheets) == {
        "Engine": {
            "A1": 2009,
            "A2": 2010,
            "B1": "=100",
            "B2": "=B1+1",
        }
    }


def test_transpose_bindings_swaps_header_kind_and_range() -> None:
    document = bindings_document(
        series_entry(
            "debt",
            "Engine!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
        )
    )
    transposed = transpose_bindings(document)
    series = transposed["series"][0]
    assert series["data_range"] == "Engine!B1:B3"
    bind = series["structure"]["dimensions"][0]["bind"]
    assert bind["kind"] == "row_label"
    assert bind["label_column"] == "A"
    restored = transpose_bindings(transposed)["series"][0]
    assert restored["data_range"] == "Engine!A2:C2"
    assert restored["structure"]["dimensions"][0]["bind"]["kind"] == "column_header"
    assert restored["structure"]["dimensions"][0]["bind"]["header_row"] == 1


def test_write_oriented_workbook_vertical_round_trips(tmp_path: Path) -> None:
    sheets = {"Engine": {"A1": 2009, "B1": 2010, "A2": 1, "B2": "=A2+1"}}
    path = write_oriented_workbook(tmp_path / "v.xlsx", sheets, orientation="vertical")
    assert path.is_file()
    document = oriented_document(
        bindings_document(
            series_entry(
                "value",
                "Engine!A2:B2",
                layout="series",
                direction="output",
                header_row=1,
            )
        ),
        "vertical",
    )
    assert document["series"][0]["data_range"] == "Engine!B1:B2"
    assert oriented_addresses(("Engine!A2", "Engine!B2"), "vertical") == (
        "Engine!B1",
        "Engine!B2",
    )


def test_orientation_fixture_is_horizontal_or_vertical(orientation: str) -> None:
    assert orientation in {"horizontal", "vertical"}
