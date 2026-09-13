"""INDEX preserves whole-row and whole-column lookup operands."""

from pathlib import Path

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a20_matrix_join import _matrix_entry


@pytest.mark.parametrize("column", ["", "0"])
def test_index_header_row_is_not_reduced_to_first_cell(tmp_path: Path, column: str) -> None:
    workbook = write_workbook(
        tmp_path / "index_vectors.xlsx",
        {
            "M": {
                "A1": "code",
                "B1": "market",
                "C1": "other",
                "D1": "header",
                "A2": "a",
                "B2": "Yes",
                "C2": "No",
                "D2": "a",
                "A3": "b",
                "B3": "No",
                "C3": "Yes",
                "D3": "b",
                "A4": 2020,
                "B4": 2021,
                "C4": 2022,
                "F1": f'=INDEX(A1:C3,MATCH("b",INDEX(A1:C3,,1),0),MATCH("other",INDEX(A1:C3,1,{column}),0))',
            }
        },
    )
    table = _matrix_entry("table", "M!A1:C3", header_row=4, label_column="D")
    table["structure"]["measure"]["dtype"] = "string"
    table["structure"]["measure"]["bind"]["read"] = "string"
    document = bindings_document(
        table, series_entry("result", "M!F1", direction="output", dtype="string")
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="index_vectors")
    assert pkg.compute_result() == "Yes"
    assert pkg.internals.result(table=pkg.data.TABLE) == "Yes"


def test_reordered_indicator_references_use_the_producer_coordinate(tmp_path: Path) -> None:
    cells = {}
    for row, label in enumerate(("interest", "discount", "maturity", "grace"), 1):
        cells[f"A{row}"] = label
        cells[f"B{row}"] = float(row * 10)
        cells[f"D{row}"] = f"result {row}"
    for row, producer in enumerate((1, 2, 4, 3), 1):
        cells[f"E{row}"] = f"=B{producer}"
    workbook = write_workbook(tmp_path / "reordered.xlsx", {"M": cells})
    document = bindings_document(
        series_entry(
            "source",
            "M!B1:B4",
            layout="series",
            direction="input",
            label_column="A",
            key_concept="INDICATOR",
            key_read="string",
        ),
        series_entry(
            "result",
            "M!E1:E4",
            layout="series",
            direction="output",
            label_column="D",
            key_concept="INDICATOR",
            key_read="string",
        ),
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="reordered")
    expected = {(f"result {row}",): float(value) for row, value in enumerate((10, 20, 40, 30), 1)}
    assert dict(pkg.compute_result(source=pkg.data.SOURCE_DEFAULT).items()) == expected
    assert dict(pkg.internals.result(source=pkg.data.SOURCE_DEFAULT).items()) == expected


@pytest.mark.parametrize("value", ["IDA Information", 12.5])
def test_offset_of_scalar_reads_the_cell_value(tmp_path: Path, value) -> None:
    dtype = "string" if isinstance(value, str) else "float"
    workbook = write_workbook(
        tmp_path / "scalar_offset.xlsx", {"M": {"A1": value, "B1": "=OFFSET(A1,0,0)"}}
    )
    document = bindings_document(
        series_entry("source", "M!A1", direction="input", dtype=dtype),
        series_entry("result", "M!B1", direction="output", dtype=dtype),
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="scalar_offset")
    assert pkg.compute_result(source=value) == value
    assert pkg.internals.result(source=value) == value
