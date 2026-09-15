"""Lookup tables and CHOOSE lists over bound cells lower to views, not cell lambdas."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.exporter.export_runtime.tensor import Axis, Domain, Tensor
from excel_grapher.exporter.inverted_tree import excel, runtime
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from tests.unit.exporter.inverted_tree.helpers import (
    assert_package_matches_evaluator,
    bindings_document,
    generate_inverted,
    series_entry,
    write_workbook,
)

YEARS = Axis("TIME_PERIOD", (2020, 2021, 2022), int)
COUNTRIES = Axis("COUNTRY", ("AF", "BR"), str)


def test_lazy_table_rows_accept_views_beside_cell_callbacks() -> None:
    flow = Tensor(Domain.product(YEARS), (1.0, 2.0, 3.0))
    table = runtime.lazy_table(
        (
            (lambda: "a", runtime.view(flow, cols=YEARS.keys)),
            (lambda: "b", lambda: 7.0, lambda: 8.0, lambda: 9.0),
        )
    )
    assert table.shape == (2, 4)
    assert excel.xl_index(table, 1, 3) == 2.0
    assert excel.xl_vlookup("b", table, 4, False) == 9.0


def test_lazy_table_strips_accept_column_views() -> None:
    codes = Tensor(Domain.product(COUNTRIES), ("AF", "BR"))
    rates = Tensor(Domain.product(COUNTRIES, YEARS), (1.0, 2.0, 3.0, 4.0, 5.0, 6.0))
    headers = Tensor(Domain.product(YEARS), (2020, 2021, 2022))
    table = runtime.lazy_table(
        (
            (lambda: "code", runtime.view(headers, cols=YEARS.keys)),
            (
                runtime.view(codes, rows=COUNTRIES.keys),
                runtime.view(rates, rows=COUNTRIES.keys, cols=YEARS.keys),
            ),
        )
    )
    assert table.shape == (3, 4)
    assert excel.xl_index(table, 1, 1) == "code"
    assert excel.xl_index(table, 1, 3) == 2021
    assert excel.xl_index(table, 3, 1) == "BR"
    assert excel.xl_vlookup("BR", table, 3, False) == 5.0


def test_lazy_table_rejects_strip_height_mismatch() -> None:
    column = runtime.view(Tensor(Domain.product(YEARS), (1.0, 2.0, 3.0)), rows=YEARS.keys)
    with pytest.raises(ValueError, match="height"):
        runtime.lazy_table(((column, lambda: 1.0),))


def test_choose_range_selects_one_cell_of_a_view() -> None:
    flow = Tensor(Domain.product(YEARS), (1.0, 2.0, 3.0))
    cells = runtime.view(flow, cols=YEARS.keys)
    assert excel.xl_choose_range(2, cells) == 2.0
    try:
        excel.xl_choose_range(4, cells)
    except excel.XlError as error:
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
        "input": {},
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


def test_multi_series_lookup_tables_keep_one_view_per_series_block(tmp_path: Path) -> None:
    modules = generate_inverted(_table_workbook(tmp_path), _table_bindings())
    internals = modules["internals.py"]
    assert (
        "lazy_table(((view(codes, rows=data.COUNTRY_AXIS.keys), "
        "view(rates, rows=data.COUNTRY_AXIS.keys, cols=data.TIME_PERIOD_AXIS.keys)),))"
    ) in internals
    assert "lambda" not in internals
    assert "_out_table_1" not in internals
    pkg = assert_package_matches_evaluator(
        _table_workbook(tmp_path), _table_bindings(), tmp_path, "table_runs"
    )
    assert pkg.compute_out(rates=pkg.data.RATES_DEFAULT) == 5.0


def _catalog_workbook(tmp_path: Path) -> Path:
    rows = (("AFG", 512, "Afghanistan"), ("BGD", 513, "Bangladesh"), ("BEN", 638, "Benin"))
    cells: dict[str, object] = {
        "A1": "Code2",
        "B1": "IMF",
        "C1": "Name",
        "E1": "iso",
        "F1": "ifs",
        "G1": "header",
        "H1": "code",
    }
    for index, (code, ifs, name) in enumerate(rows, start=2):
        cells[f"A{index}"] = code
        cells[f"B{index}"] = ifs
        cells[f"C{index}"] = name
        cells[f"E{index}"] = code
        cells[f"F{index}"] = ifs
        cells[f"G{index}"] = "Code2"
        cells[f"H{index}"] = (
            f"=INDEX($A$1:$C$4,MATCH(F{index},$B$1:$B$4,0),MATCH(G{index},$A$1:$C$1,0))"
        )
    return write_workbook(tmp_path / "index_match_catalog.xlsx", {"Data": cells})


def _catalog_bindings() -> dict[str, Any]:
    document = bindings_document(
        series_entry(
            "catalog_code",
            "Data!A2:A4",
            layout="series",
            direction="constant",
            dtype="string",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
        series_entry(
            "catalog_ifs",
            "Data!B2:B4",
            layout="series",
            direction="constant",
            dtype="int",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
        series_entry(
            "catalog_name",
            "Data!C2:C4",
            layout="series",
            direction="constant",
            dtype="string",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
        series_entry(
            "catalog_headers",
            "Data!A1:C1",
            layout="series",
            direction="constant",
            dtype="string",
            header_row=1,
            key_concept="VARIANT",
            key_read="string",
        ),
        series_entry(
            "trigger_ifs",
            "Data!F2:F4",
            layout="series",
            direction="input",
            dtype="int",
            label_column="E",
            key_concept="COUNTRY",
            key_read="string",
        ),
        series_entry(
            "trigger_header",
            "Data!G2:G4",
            layout="series",
            direction="input",
            dtype="string",
            label_column="E",
            key_concept="COUNTRY",
            key_read="string",
        ),
        series_entry(
            "trigger_code",
            "Data!H2:H4",
            layout="series",
            direction="output",
            dtype="string",
            label_column="E",
            key_concept="COUNTRY",
            key_read="string",
        ),
    )
    document["concept_scheme"]["concepts"].append({"id": "VARIANT", "dtype": "string"})
    return document


def _catalog_dynamic_refs() -> DynamicRefConfig:
    return DynamicRefConfig.from_constraints({}, {})


def test_multi_column_index_match_catalog_is_views_not_per_key_lambdas(tmp_path: Path) -> None:
    workbook = _catalog_workbook(tmp_path)
    bindings = _catalog_bindings()
    dynamic_refs = _catalog_dynamic_refs()
    modules = generate_inverted(workbook, bindings, dynamic_refs=dynamic_refs)
    internals = modules["internals.py"]
    assert "view(catalog_headers, cols=data.VARIANT_AXIS.keys)" in internals
    assert "view(catalog_code, rows=data.COUNTRY_AXIS.keys)" in internals
    assert "view(catalog_ifs, rows=data.COUNTRY_AXIS.keys)" in internals
    assert "view(catalog_name, rows=data.COUNTRY_AXIS.keys)" in internals
    assert "lambda: catalog_ifs['AFG']" not in internals
    assert "lambda: catalog_code['AFG']" not in internals
    assert "lambda: catalog_name['AFG']" not in internals
    assert "xl_index(" in internals
    assert "xl_match(" in internals
    pkg = assert_package_matches_evaluator(
        workbook, bindings, tmp_path, "catalog_index_match", dynamic_refs=dynamic_refs
    )
    result = pkg.compute_trigger_code(
        trigger_ifs=pkg.data.TRIGGER_IFS_DEFAULT,
        trigger_header=pkg.data.TRIGGER_HEADER_DEFAULT,
    )
    assert result["AFG"] == "AFG"
    assert result["BGD"] == "BGD"
    assert result["BEN"] == "BEN"
    missing = type(pkg.data.TRIGGER_IFS_DEFAULT).from_nested(
        domain=pkg.data.TRIGGER_IFS_DEFAULT.domain,
        values=(999, 513, 638),
    )
    unmatched = pkg.compute_trigger_code(
        trigger_ifs=missing,
        trigger_header=pkg.data.TRIGGER_HEADER_DEFAULT,
    )
    assert unmatched["AFG"] == "#N/A"
    assert unmatched["BGD"] == "BGD"
