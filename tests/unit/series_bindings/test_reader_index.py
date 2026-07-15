"""Tests for Phase 1b reverse address/range → reader call mapping."""

from __future__ import annotations

import importlib
import sys
from pathlib import Path
from typing import Any

import xlsxwriter

from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    build_reader_index,
    expand_data_range,
    format_reader_call_form,
    load_series_bindings,
    resolve_reader_ref,
    resolve_series_binding,
    validate_bindings_document,
)
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


def _write_scalar_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    ws.write("B5", "Borvelia")
    wb.close()


def _write_dup_headers_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write(0, 2, 1)
    ws.write(0, 3, 1)
    ws.write_number(1, 2, 10.0)
    ws.write_number(1, 3, 20.0)
    wb.close()


KEYLESS_SCALAR_BINDING: dict[str, Any] = {
    "schema_version": "1.3.0",
    "series": [
        {
            "id": "country_name",
            "sheet": "Inputs",
            "data_range": "Inputs!B5",
            "layout": "scalar",
            "input": {
                "setter": {
                    "name": "set_country_name",
                    "record_contract": "records",
                    "strict": True,
                }
            },
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "string",
                    "bind": {"kind": "data_cell", "read": "string"},
                },
                "dimensions": [],
            },
            "key": [],
        }
    ],
}


def test_format_reader_call_form_keyed() -> None:
    assert (
        format_reader_call_form(
            "read_borvelia_primary_balance",
            kwargs={"time_period": 3},
        )
        == "read_borvelia_primary_balance(ctx, time_period=3)"
    )


def test_format_reader_call_form_scalar() -> None:
    assert format_reader_call_form("read_country_name") == "read_country_name(ctx)"


def test_format_reader_call_form_address_keyed() -> None:
    assert (
        format_reader_call_form("read_dup_headers", address="Sheet1!C2")
        == "read_dup_headers(ctx, address='Sheet1!C2')"
    )


def test_format_reader_call_form_range() -> None:
    assert (
        format_reader_call_form("read_borvelia_primary_balance_range", range_reader=True)
        == "read_borvelia_primary_balance_range(ctx)"
    )


def test_resolve_keyed_leaf_to_reader_call(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")

    result = resolve_reader_ref(
        "Inputs!H5",
        graph=graph,
        bindings=bindings,
        workbook=wb_path,
    )

    assert result["mode"] == "reader"
    assert result["series_id"] == "borvelia_primary_balance"
    assert result["reader"] == "read_borvelia_primary_balance"
    assert result["keys"] == {"TIME_PERIOD": 3}
    assert result["kwargs"] == {"time_period": 3}
    assert result["call_form"] == "read_borvelia_primary_balance(ctx, time_period=3)"
    assert result["reason"] is None


def test_resolve_binding_aligned_range_to_range_reader(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")

    result = resolve_reader_ref(
        "Inputs!F5:J5",
        graph=graph,
        bindings=bindings,
        workbook=wb_path,
    )

    assert result["mode"] == "reader_range"
    assert result["series_id"] == "borvelia_primary_balance"
    assert result["reader"] == "read_borvelia_primary_balance_range"
    assert result["call_form"] == "read_borvelia_primary_balance_range(ctx)"
    assert result["reason"] is None


def test_resolve_equivalent_range_form_to_range_reader(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")

    result = resolve_reader_ref(
        "Inputs!F5:Inputs!J5",
        graph=graph,
        bindings=bindings,
        workbook=wb_path,
    )

    assert result["mode"] == "reader_range"
    assert result["reader"] == "read_borvelia_primary_balance_range"


def test_resolve_partial_range_falls_back_to_xl_range(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")

    result = resolve_reader_ref(
        "Inputs!F5:H5",
        graph=graph,
        bindings=bindings,
        workbook=wb_path,
    )

    assert result["mode"] == "xl_range"
    assert result["reason"] == "not_binding_aligned_range"
    assert result["call_form"] == "xl_range(ctx, 'Inputs!F5:H5')"
    assert result["reader"] is None


def test_resolve_unbound_cell_falls_back_to_xl_cell(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")

    result = resolve_reader_ref(
        "Inputs!A2",
        graph=graph,
        bindings=bindings,
        workbook=wb_path,
    )

    assert result["mode"] == "xl_cell"
    assert result["reason"] == "unbound"
    assert result["call_form"] == "xl_cell(ctx, 'Inputs!A2')"


def test_resolve_keyless_scalar_leaf(tmp_path: Path) -> None:
    wb_path = tmp_path / "scalar.xlsx"
    _write_scalar_workbook(wb_path)
    bindings = validate_bindings_document(KEYLESS_SCALAR_BINDING)
    graph = create_dependency_graph(wb_path, ["Inputs!B5"], load_values=True)

    result = resolve_reader_ref(
        "Inputs!B5",
        graph=graph,
        bindings=bindings,
        workbook=wb_path,
    )

    assert result["mode"] == "reader"
    assert result["series_id"] == "country_name"
    assert result["reader"] == "read_country_name"
    assert result["keys"] == {}
    assert result["kwargs"] == {}
    assert result["call_form"] == "read_country_name(ctx)"
    assert result["kind"] == "scalar"


def test_resolve_address_keyed_leaf(tmp_path: Path) -> None:
    wb_path = tmp_path / "dup_headers.xlsx"
    _write_dup_headers_workbook(wb_path)
    bindings = validate_bindings_document(
        {
            "schema_version": "1.3.0",
            "series": [
                {
                    "id": "dup_headers",
                    "sheet": "Sheet1",
                    "data_range": "Sheet1!C2:D2",
                    "layout": "series",
                    "setter": {"name": "set_dup_headers", "allow_address": True, "strict": False},
                    "structure": {
                        "measure": {
                            "concept": "OBS_VALUE",
                            "dtype": "float",
                            "bind": {"kind": "data_cell", "read": "float"},
                        },
                        "dimensions": [
                            {
                                "concept": "TIME_PERIOD",
                                "role": "key",
                                "scope": "cell",
                                "bind": {
                                    "kind": "column_header",
                                    "header_row": 1,
                                    "read": "int",
                                },
                            }
                        ],
                    },
                    "key": ["TIME_PERIOD"],
                    "validation": {"require_unique_key": True},
                }
            ],
        }
    )
    graph = create_dependency_graph(wb_path, ["Sheet1!C2", "Sheet1!D2"], load_values=True)
    resolved = resolve_series_binding(graph, wb_path, bindings["series"][0])
    assert resolved["requires_address"] is True

    result = resolve_reader_ref(
        "Sheet1!C2",
        graph=graph,
        bindings=bindings,
        workbook=wb_path,
    )

    assert result["mode"] == "reader"
    assert result["kind"] == "address_keyed"
    assert result["reader"] == "read_dup_headers"
    assert result["call_form"] == "read_dup_headers(ctx, address='Sheet1!C2')"


def test_resolve_ambiguous_owner_falls_back(tmp_path: Path) -> None:
    wb_path = tmp_path / "overlap.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Inputs")
    ws.write_number(0, 0, 1.0)
    wb.close()

    bindings = validate_bindings_document(
        {
            "schema_version": "1.3.0",
            "series": [
                {
                    "id": "alpha",
                    "sheet": "Inputs",
                    "data_range": "Inputs!A1",
                    "layout": "scalar",
                    "setter": {"name": "set_alpha"},
                    "structure": {
                        "measure": {
                            "concept": "OBS_VALUE",
                            "dtype": "float",
                            "bind": {"kind": "data_cell", "read": "float"},
                        },
                        "dimensions": [],
                    },
                    "key": [],
                },
                {
                    "id": "beta",
                    "sheet": "Inputs",
                    "data_range": "Inputs!A1",
                    "layout": "scalar",
                    "setter": {"name": "set_beta"},
                    "structure": {
                        "measure": {
                            "concept": "OBS_VALUE",
                            "dtype": "float",
                            "bind": {"kind": "data_cell", "read": "float"},
                        },
                        "dimensions": [],
                    },
                    "key": [],
                },
            ],
        }
    )
    graph = create_dependency_graph(wb_path, ["Inputs!A1"], load_values=True)

    result = resolve_reader_ref(
        "Inputs!A1",
        graph=graph,
        bindings=bindings,
        workbook=wb_path,
    )

    assert result["mode"] == "xl_cell"
    assert result["reason"] == "ambiguous_owner"
    assert result["call_form"] == "xl_cell(ctx, 'Inputs!A1')"


def test_build_reader_index_excludes_ambiguous_leaves(tmp_path: Path) -> None:
    wb_path = tmp_path / "overlap.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Inputs")
    ws.write_number(0, 0, 1.0)
    wb.close()

    bindings = validate_bindings_document(
        {
            "schema_version": "1.3.0",
            "series": [
                {
                    "id": "alpha",
                    "sheet": "Inputs",
                    "data_range": "Inputs!A1",
                    "layout": "scalar",
                    "setter": {"name": "set_alpha"},
                    "structure": {
                        "measure": {
                            "concept": "OBS_VALUE",
                            "dtype": "float",
                            "bind": {"kind": "data_cell", "read": "float"},
                        },
                        "dimensions": [],
                    },
                    "key": [],
                },
                {
                    "id": "beta",
                    "sheet": "Inputs",
                    "data_range": "Inputs!A1",
                    "layout": "scalar",
                    "setter": {"name": "set_beta"},
                    "structure": {
                        "measure": {
                            "concept": "OBS_VALUE",
                            "dtype": "float",
                            "bind": {"kind": "data_cell", "read": "float"},
                        },
                        "dimensions": [],
                    },
                    "key": [],
                },
            ],
        }
    )
    graph = create_dependency_graph(wb_path, ["Inputs!A1"], load_values=True)
    index = build_reader_index(graph, bindings, workbook=wb_path)

    assert "Inputs!A1" not in index["leaves"]
    assert "Inputs!A1" in index["ambiguous"]


def test_build_reader_index_keyed_and_range(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")

    index = build_reader_index(graph, bindings, workbook=wb_path)

    leaf = index["leaves"]["Inputs!H5"]
    assert leaf["series_id"] == "borvelia_primary_balance"
    assert leaf["reader"] == "read_borvelia_primary_balance"
    assert leaf["keys"] == {"TIME_PERIOD": 3}
    assert leaf["kwargs"] == {"time_period": 3}
    assert leaf["kind"] == "keyed"
    assert leaf["call_form"] == "read_borvelia_primary_balance(ctx, time_period=3)"

    rng = index["ranges"]["Inputs!F5:J5"]
    assert rng["reader"] == "read_borvelia_primary_balance_range"
    assert rng["call_form"] == "read_borvelia_primary_balance_range(ctx)"


def test_generated_modules_emit_list_reader_leaves(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    files = CodeGenerator(graph).generate_modules(
        expand_data_range("Inputs!F5:J5"),
        series_bindings=bindings,
        bindings_workbook=wb_path,
    )
    assert "def list_reader_leaves(" in files["api.py"]
    assert "def list_reader_ranges(" in files["api.py"]
    assert "list_reader_leaves" in files["__init__.py"]
    assert "list_reader_ranges" in files["__init__.py"]

    pkg_dir = tmp_path / "exported_readers"
    for filename, content in files.items():
        pkg_dir.mkdir(parents=True, exist_ok=True)
        (pkg_dir / filename).write_text(content, encoding="utf-8")

    sys.path.insert(0, str(tmp_path))
    try:
        pkg = importlib.import_module("exported_readers")
        leaves = pkg.list_reader_leaves()
        ranges = pkg.list_reader_ranges()
        assert leaves["Inputs!H5"]["call_form"] == (
            "read_borvelia_primary_balance(ctx, time_period=3)"
        )
        assert ranges["Inputs!F5:J5"]["call_form"] == ("read_borvelia_primary_balance_range(ctx)")
    finally:
        sys.path.remove(str(tmp_path))
        for name in list(sys.modules):
            if name == "exported_readers" or name.startswith("exported_readers."):
                sys.modules.pop(name, None)
