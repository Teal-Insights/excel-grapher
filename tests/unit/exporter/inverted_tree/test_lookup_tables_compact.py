"""Lookup tables and CHOOSE lists over bound cells lower to views, not cell lambdas."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from excel_grapher.exporter.export_runtime.tensor import Axis, Domain, Tensor
from excel_grapher.exporter.inverted_tree import runtime
from tests.unit.exporter.inverted_tree.helpers import (
    assert_package_matches_evaluator,
    bindings_document,
    generate_inverted,
    series_entry,
    write_workbook,
)

YEARS = Axis("TIME_PERIOD", (2020, 2021, 2022), int)


def test_lazy_table_rows_accept_views_beside_cell_callbacks() -> None:
    flow = Tensor(Domain.product(YEARS), (1.0, 2.0, 3.0))
    table = runtime.lazy_table(
        (
            (lambda: "a", runtime.view(flow, cols=YEARS.keys)),
            (lambda: "b", lambda: 7.0, lambda: 8.0, lambda: 9.0),
        )
    )
    assert table.shape == (2, 4)
    assert runtime.xl_index(table, 1, 3) == 2.0
    assert runtime.xl_vlookup("b", table, 4, False) == 9.0


def test_choose_range_selects_one_cell_of_a_view() -> None:
    flow = Tensor(Domain.product(YEARS), (1.0, 2.0, 3.0))
    cells = runtime.view(flow, cols=YEARS.keys)
    assert runtime.xl_choose_range(2, cells) == 2.0
    try:
        runtime.xl_choose_range(4, cells)
    except runtime.XlError as error:
        assert error.code == "#VALUE!"
    else:
        raise AssertionError("expected #VALUE!")


def _choose_workbook(tmp_path: Path) -> Path:
    cells: dict[str, object] = {"A4": 2}
    for index, column in enumerate("BCD"):
        cells[f"{column}1"] = 2020 + index
        cells[f"{column}2"] = float(index + 1)
        cells[f"{column}3"] = "=CHOOSE($A$4, $B$2, $C$2, $D$2)"
    return write_workbook(tmp_path / "choose.xlsx", {"Sheet": cells})


def _choose_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("flow", "Sheet!B2:D2", layout="series", direction="input", header_row=1),
        series_entry("picked", "Sheet!B3:D3", layout="series", direction="output", header_row=1),
        series_entry("which", "Sheet!A4", layout="scalar", direction="input", dtype="int"),
    )


def test_choose_over_a_run_of_cells_is_a_view(tmp_path: Path) -> None:
    modules = generate_inverted(_choose_workbook(tmp_path), _choose_bindings())
    internals = modules["internals.py"]
    assert "_picked_table_0 = view(flow, cols=data.TIME_PERIOD_AXIS.keys)" in internals
    assert "xl_choose_range(which, _picked_table_0)" in internals
    assert "lambda" not in internals
    pkg = assert_package_matches_evaluator(
        _choose_workbook(tmp_path), _choose_bindings(), tmp_path, "choose_view"
    )
    picked = pkg.compute_picked(flow=pkg.data.FLOW_DEFAULT, which=3)
    assert picked[2021] == 3.0


def _table_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "table.xlsx",
        {
            "Table": {
                "A1": "code",
                "B1": 2020,
                "C1": 2021,
                "D1": 2022,
                "A2": "AF",
                "B2": 1.0,
                "C2": 2.0,
                "D2": 3.0,
                "A3": "BR",
                "B3": 4.0,
                "C3": 5.0,
                "D3": 6.0,
            },
            "Outputs": {"A1": '=VLOOKUP("BR", Table!$A$2:$D$3, 3, FALSE)'},
        },
    )


def _table_bindings() -> dict[str, Any]:
    rates = {
        "id": "rates",
        "sheet": "Table",
        "data_range": "Table!B2:D3",
        "layout": "matrix",
        "input": {"setter": {"name": "set_rates"}},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "id": "COUNTRY",
                    "concept": "COUNTRY",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "row_label", "label_column": "A", "read": "string"},
                },
                {
                    "id": "TIME_PERIOD",
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                },
            ],
        },
        "key": ["COUNTRY", "TIME_PERIOD"],
    }
    return bindings_document(
        series_entry(
            "codes",
            "Table!A2:A3",
            layout="series",
            direction="constant",
            dtype="string",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
        rates,
        series_entry("out", "Outputs!A1", layout="scalar", direction="output"),
    )


def test_multi_series_lookup_tables_keep_one_view_per_row_run(tmp_path: Path) -> None:
    modules = generate_inverted(_table_workbook(tmp_path), _table_bindings())
    internals = modules["internals.py"]
    assert (
        "lazy_table(((lambda: codes['AF'], view(rates, rows=('AF',), "
        "cols=data.TIME_PERIOD_AXIS.keys)), (lambda: codes['BR'], view(rates, "
        "rows=('BR',), cols=data.TIME_PERIOD_AXIS.keys))))"
    ) in internals
    assert "_out_table_1" not in internals
    pkg = assert_package_matches_evaluator(
        _table_workbook(tmp_path), _table_bindings(), tmp_path, "table_runs"
    )
    assert pkg.compute_out(rates=pkg.data.RATES_DEFAULT) == 5.0
