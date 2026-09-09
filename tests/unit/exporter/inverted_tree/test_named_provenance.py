"""Regular coordinate-to-cell provenance is a worksheet rectangle, not a dictionary."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from fastpyxl.utils.cell import get_column_letter

from excel_grapher.exporter.export_runtime.provenance import block_cells, column_cells, row_cells
from excel_grapher.exporter.export_runtime.tensor import Axis
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)

YEARS = Axis("TIME_PERIOD", (2020, 2021, 2022), int)
COUNTRIES = Axis("COUNTRY", ("France", "Kenya"), str)


def test_row_cells_map_axis_keys_to_consecutive_columns() -> None:
    cells = row_cells("Engine", 5, "C", YEARS)
    assert dict(cells) == {(2020,): "Engine!C5", (2021,): "Engine!D5", (2022,): "Engine!E5"}
    assert cells[(2021,)] == "Engine!D5"
    assert len(cells) == 3
    assert (2023,) not in cells


def test_column_cells_and_block_cells_follow_worksheet_order() -> None:
    column = column_cells("Inputs", "A", 10, COUNTRIES)
    assert list(column.items()) == [(("France",), "Inputs!A10"), (("Kenya",), "Inputs!A11")]
    block = block_cells("Engine", 10, "C", COUNTRIES, YEARS)
    assert block[("Kenya", 2021)] == "Engine!D11"
    assert list(block)[:3] == [("France", 2020), ("France", 2021), ("France", 2022)]
    transposed = block_cells("Engine", 10, "C", COUNTRIES, YEARS, cols_first=True)
    assert transposed[(2021, "Kenya")] == "Engine!D11"
    assert list(transposed)[:2] == [(2020, "France"), (2020, "Kenya")]


def test_block_cells_keep_explicit_exceptions() -> None:
    cells = block_cells(
        "Engine", 10, "C", COUNTRIES, YEARS, exceptions={("Kenya", 2022): "Other!Z1"}
    )
    assert cells[("Kenya", 2022)] == "Other!Z1"
    assert cells[("Kenya", 2021)] == "Engine!D11"
    assert dict(cells)[("Kenya", 2022)] == "Other!Z1"


def _horizon_workbook(tmp_path: Path, years: int) -> Path:
    cells: dict[str, object] = {}
    for index in range(years):
        column = get_column_letter(index + 2)
        cells[f"{column}1"] = 2020 + index
        cells[f"{column}2"] = float(index + 1)
        cells[f"{column}3"] = f"={column}2*2"
    return write_workbook(tmp_path / f"horizon_{years}.xlsx", {"Sheet": cells})


def _horizon_bindings(years: int) -> dict[str, Any]:
    last = get_column_letter(years + 1)
    return bindings_document(
        series_entry("flow", f"Sheet!B2:{last}2", layout="series", direction="input", header_row=1),
        series_entry(
            "twice", f"Sheet!B3:{last}3", layout="series", direction="output", header_row=1
        ),
    )


def test_data_module_provenance_does_not_grow_with_the_horizon(tmp_path: Path) -> None:
    sizes = {}
    for years in (10, 80):
        modules = generate_inverted(_horizon_workbook(tmp_path, years), _horizon_bindings(years))
        data = modules["data.py"]
        assert "row_cells('Sheet', 3, 'B', TIME_PERIOD_AXIS)" in data
        cells_lines = [line for line in data.splitlines() if "_CELLS = " in line]
        sizes[years] = sum(len(line) for line in cells_lines)
    assert sizes[80] == sizes[10], sizes
    pkg = load_package(modules, tmp_path, name="horizon_provenance")
    assert pkg.compute_twice.__cells__[(2025,)] == "Sheet!G3"
    assert dict(pkg.data.FLOW_CELLS)[(2020,)] == "Sheet!B2"


def test_series_facades_declare_their_schema_once(tmp_path: Path) -> None:
    modules = generate_inverted(_horizon_workbook(tmp_path, 5), _horizon_bindings(5))
    data = modules["data.py"]
    assert "class Twice(Series[T]):\n" in data
    assert "    schema = TWICE_SCHEMA\n" in data
    assert "def __post_init__" not in data
    assert "def __getitem__" not in data
    pkg = load_package(modules, tmp_path, name="horizon_facades")
    tensor = pkg.compute_twice(flow=pkg.data.FLOW_DEFAULT)
    assert isinstance(tensor, pkg.data.Twice)
    assert tensor[2022] == 6.0
    try:
        pkg.data.Twice(pkg.data.TWICE_DOMAIN, (1.0, 2.0, 3.0, 4.0, "#N/A"))
    except pkg.tensor.SchemaError:
        raise AssertionError("errors are valid observations") from None
    import pytest

    with pytest.raises(pkg.tensor.SchemaError, match="twice"):
        pkg.data.Twice(pkg.tensor.Domain.product(pkg.tensor.Axis("year", (1, 2), int)), (1.0, 2.0))
