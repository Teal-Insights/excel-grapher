"""Unit tests for input series derived from series binding manifests."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher.exporter import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import derive_input_series, load_series_bindings
from tests.paths import SERIES_BINDINGS_FIXTURES as FIXTURES


def _write_borvelia_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    ws.write("A2", "Borvelia")
    ws.write("A5", "Primary balance (% of GDP)")
    for col, year in enumerate([1, 2, 3, 4, 5], start=5):
        ws.write(0, col, year)
        ws.write_number(4, col, float(year - 3))
    wb.close()


def test_derive_input_series_from_series_bindings(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        ["Inputs!F5", "Inputs!G5", "Inputs!H5", "Inputs!I5", "Inputs!J5"],
        load_values=True,
    )
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")

    input_series = derive_input_series(graph, bindings, workbook=wb_path)

    assert len(input_series) == 1
    series = input_series[0]
    assert series["id"] == "borvelia_primary_balance"
    assert series["setter_name"] == "set_borvelia_primary_balance"
    assert series["key_fields"] == ["TIME_PERIOD"]
    assert series["requires_address"] is False
    assert [cell["address"] for cell in series["cells"]] == [
        "Inputs!F5",
        "Inputs!G5",
        "Inputs!H5",
        "Inputs!I5",
        "Inputs!J5",
    ]
    assert series["cells"][2]["key"] == {"TIME_PERIOD": 3}
    assert series["cells"][2]["record"]["OBS_VALUE"] == 0.0


def test_derive_input_series_filters_to_graph_leaf_overlap(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!H5"], load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")

    input_series = derive_input_series(graph, bindings, workbook=wb_path)

    assert len(input_series) == 1
    assert [cell["address"] for cell in input_series[0]["cells"]] == ["Inputs!H5"]
    assert input_series[0]["cells"][0]["key"] == {"TIME_PERIOD": 3}


def test_derive_input_series_skips_series_without_graph_leaf_overlap(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!A2"], load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")

    assert derive_input_series(graph, bindings, workbook=wb_path) == []


def test_code_generator_derives_input_series_from_bindings(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!H5"], load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")

    input_series = CodeGenerator(graph).derive_input_series(bindings, workbook=wb_path)

    assert len(input_series) == 1
    assert input_series[0]["id"] == "borvelia_primary_balance"
    assert [cell["address"] for cell in input_series[0]["cells"]] == ["Inputs!H5"]
