"""Issue 855 — bind series with define_series instead of empty subclasses."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from fastpyxl import Workbook

from excel_grapher.exporter.inverted_tree.emit import generate_inverted_tree_modules
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import validate_bindings_document
from tests.unit.exporter.inverted_tree.helpers import load_package


def _series(
    sid: str, data_range: str, key: list[str], dims: list[dict[str, Any]]
) -> dict[str, Any]:
    return {
        "id": sid,
        "sheet": "Data",
        "data_range": data_range,
        "layout": "series",
        "input": {},
        "key": key,
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": dims,
        },
    }


def _dim_row(column: str) -> dict[str, Any]:
    return {
        "id": "VARIANT",
        "concept": "VARIANT",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "row_label", "label_column": column, "read": "string"},
    }


def _dim_col(header_row: int) -> dict[str, Any]:
    return {
        "id": "TIME_PERIOD",
        "concept": "TIME_PERIOD",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "column_header", "header_row": header_row, "read": "int"},
    }


def _mcve_workbook(tmp_path: Path) -> Path:
    path = tmp_path / "series_blocks.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "Data"
    sheet["A1"] = "item"
    sheet["B1"] = 2024
    sheet["C1"] = 2025
    sheet["D1"] = 2026
    sheet["A2"] = "alpha"
    sheet["B2"] = 1
    sheet["C2"] = 2
    sheet["D2"] = 3
    sheet["A4"] = "item"
    sheet["C4"] = 2025
    sheet["D4"] = 2026
    sheet["A5"] = "beta"
    sheet["C5"] = 10
    sheet["D5"] = 11
    sheet["B7"] = "=SUM(B2:D2)+SUM(C5:D5)"
    book.save(path)
    return path


def _mcve_bindings() -> dict[str, Any]:
    return validate_bindings_document(
        {
            "schema_version": "1.16.0",
            "concept_scheme": {
                "id": "example",
                "concepts": [
                    {"id": "OBS_VALUE", "dtype": "number"},
                    {"id": "VARIANT", "dtype": "string"},
                    {"id": "TIME_PERIOD", "dtype": "int"},
                ],
            },
            "series": [
                _series(
                    "prices_full",
                    "Data!B2:D2",
                    ["VARIANT", "TIME_PERIOD"],
                    [_dim_row("A"), _dim_col(1)],
                ),
                _series(
                    "prices_short",
                    "Data!C5:D5",
                    ["VARIANT", "TIME_PERIOD"],
                    [_dim_row("A"), _dim_col(4)],
                ),
                {
                    "id": "total",
                    "sheet": "Data",
                    "data_range": "Data!B7",
                    "layout": "scalar",
                    "output": {"compute": {"name": "compute_total"}},
                    "key": [],
                    "structure": {
                        "measure": {
                            "concept": "OBS_VALUE",
                            "dtype": "float",
                            "bind": {"kind": "data_cell", "read": "float"},
                        },
                        "dimensions": [],
                    },
                },
            ],
        }
    )


def _generate_mcve(tmp_path: Path) -> dict[str, str]:
    path = _mcve_workbook(tmp_path)
    bindings = _mcve_bindings()
    graph = create_dependency_graph(path, ["Data!B7"], load_values=True)
    return generate_inverted_tree_modules(graph, series_bindings=bindings, bindings_workbook=path)


def test_mcve_data_module_has_no_empty_series_subclasses(tmp_path: Path) -> None:
    data = _generate_mcve(tmp_path)["data.py"]
    assert "class PricesFull" not in data
    assert "class PricesShort" not in data
    assert "PRICES_FULL_REQUIRED" not in data
    assert "PRICES_SHORT_REQUIRED" not in data
    assert "PRICES_FULL_SCHEMA" not in data
    assert "required=" not in data
    assert data.count("define_series(") == 2
    assert "PRICES_FULL: Series[" in data
    assert "PRICES_SHORT: Series[" in data
    assert "TOTAL_CELLS = {(): 'Data!B7'}" in data


def test_mcve_bindings_compute_and_carry_schema(tmp_path: Path) -> None:
    modules = _generate_mcve(tmp_path)
    pkg = load_package(modules, tmp_path, name="series_define_mcve")
    full = pkg.data.PRICES_FULL
    short = pkg.data.PRICES_SHORT
    assert full[("alpha", 2024)] == 1.0
    assert short[("beta", 2026)] == 11.0
    assert full.schema.series_id == "prices_full"
    assert full.required is full.domain
    assert dict(full.cells)[("alpha", 2025)] == "Data!C2"
    assert pkg.compute_total(prices_full=full, prices_short=short) == 27.0
    assert isinstance(full, pkg.Series)
    full.schema.validate(full)
    short.schema.validate(short)
