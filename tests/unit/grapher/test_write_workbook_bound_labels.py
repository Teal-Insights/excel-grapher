"""Bound row_label / column_header cells in Excel write-back (#895)."""

from __future__ import annotations

from pathlib import Path

import fastpyxl
import pytest

from excel_grapher.grapher import create_dependency_graph, write_workbook
from excel_grapher.series_bindings import bound_label_addresses, validate_bindings_document
from excel_grapher.series_bindings.types import WorkbookSeriesBindings
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.fixtures.series_bindings.grouped_matrix_helpers import (
    grouped_matrix_bindings_document,
    write_grouped_matrix_workbook,
)


def _cell_value(path: Path, sheet: str, coord: str) -> object:
    wb = fastpyxl.load_workbook(path)
    try:
        return wb[sheet][coord].value
    finally:
        wb.close()


def _write_labelled_source(path: Path) -> None:
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = "indicator"
    ws["B1"] = 2024
    ws["C1"] = 2025
    ws["A2"] = "GDP"
    ws["B2"] = 1.1
    ws["C2"] = 1.2
    ws["Z99"] = "stray"
    wb.save(path)
    wb.close()


def _labelled_bindings() -> WorkbookSeriesBindings:
    return validate_bindings_document(
        {
            "schema_version": "1.17.0",
            "concept_scheme": {
                "id": "writeback_labels",
                "concepts": [
                    {"id": "TIME_PERIOD", "dtype": "int"},
                    {"id": "INDICATOR", "dtype": "string"},
                    {"id": "OBS_VALUE", "dtype": "number"},
                ],
            },
            "series": [
                {
                    "id": "gdp",
                    "sheet": "Sheet1",
                    "data_range": "Sheet1!B2:C2",
                    "layout": "series",
                    "input": {},
                    "structure": {
                        "measure": {
                            "concept": "OBS_VALUE",
                            "dtype": "float",
                            "bind": {"kind": "data_cell", "read": "float"},
                        },
                        "dimensions": [
                            {
                                "concept": "INDICATOR",
                                "role": "key",
                                "scope": "cell",
                                "bind": {
                                    "kind": "row_label",
                                    "label_column": "A",
                                    "read": "string",
                                },
                            },
                            {
                                "concept": "TIME_PERIOD",
                                "role": "key",
                                "scope": "cell",
                                "bind": {
                                    "kind": "column_header",
                                    "header_row": 1,
                                    "read": "int",
                                },
                            },
                        ],
                    },
                    "key": ["INDICATOR", "TIME_PERIOD"],
                }
            ],
        }
    )


def _write_formula_header_source(path: Path) -> None:
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 2024
    ws["B1"] = "=A1+1"
    ws["B2"] = 1.5
    ws["Z99"] = "stray"
    wb.save(path)
    wb.close()


def _formula_header_bindings() -> WorkbookSeriesBindings:
    return validate_bindings_document(
        {
            "schema_version": "1.17.0",
            "concept_scheme": {
                "id": "writeback_formula_labels",
                "concepts": [
                    {"id": "TIME_PERIOD", "dtype": "int"},
                    {"id": "OBS_VALUE", "dtype": "number"},
                ],
            },
            "series": [
                {
                    "id": "gdp",
                    "sheet": "Sheet1",
                    "data_range": "Sheet1!B2",
                    "layout": "series",
                    "input": {},
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
                            },
                        ],
                    },
                    "key": ["TIME_PERIOD"],
                }
            ],
        }
    )


def test_write_workbook_without_bindings_omits_unreferenced_labels(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_labelled_source(source)
    bindings = _labelled_bindings()
    graph = create_dependency_graph(
        source,
        all_series_targets(bindings, workbook=source),
        load_values=True,
    )
    assert "Sheet1!A2" not in graph
    assert "Sheet1!B1" not in graph

    write_workbook(graph, dest)

    assert _cell_value(dest, "Sheet1", "B2") == 1.1
    assert _cell_value(dest, "Sheet1", "A2") is None
    assert _cell_value(dest, "Sheet1", "B1") is None
    assert _cell_value(dest, "Sheet1", "Z99") is None


def test_write_workbook_includes_bound_labels_from_bindings(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_labelled_source(source)
    bindings = _labelled_bindings()
    graph = create_dependency_graph(
        source,
        all_series_targets(bindings, workbook=source),
        load_values=True,
    )
    original_keys = set(graph)

    write_workbook(
        graph,
        dest,
        series_bindings=bindings,
        bindings_workbook=source,
    )

    assert set(graph) == original_keys
    assert _cell_value(dest, "Sheet1", "B2") == 1.1
    assert _cell_value(dest, "Sheet1", "C2") == 1.2
    assert _cell_value(dest, "Sheet1", "A2") == "GDP"
    assert _cell_value(dest, "Sheet1", "B1") == 2024
    assert _cell_value(dest, "Sheet1", "C1") == 2025
    assert _cell_value(dest, "Sheet1", "A1") is None
    assert _cell_value(dest, "Sheet1", "Z99") is None


def test_write_workbook_writes_formula_label_closure(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_formula_header_source(source)
    bindings = _formula_header_bindings()
    graph = create_dependency_graph(
        source,
        all_series_targets(bindings, workbook=source),
        load_values=True,
    )
    assert "Sheet1!B1" not in graph
    assert "Sheet1!A1" not in graph

    write_workbook(
        graph,
        dest,
        series_bindings=bindings,
        bindings_workbook=source,
    )

    assert _cell_value(dest, "Sheet1", "B2") == 1.5
    assert _cell_value(dest, "Sheet1", "B1") == "=A1+1"
    assert _cell_value(dest, "Sheet1", "A1") == 2024
    assert _cell_value(dest, "Sheet1", "Z99") is None


def test_write_workbook_bound_labels_require_workbook(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    _write_labelled_source(source)
    bindings = _labelled_bindings()
    graph = create_dependency_graph(
        source,
        all_series_targets(bindings, workbook=source),
        load_values=True,
    )
    with pytest.raises(ValueError, match="bindings_workbook"):
        write_workbook(graph, tmp_path / "out.xlsx", series_bindings=bindings)


def test_write_workbook_include_bound_labels_false_omits_labels(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_labelled_source(source)
    bindings = _labelled_bindings()
    graph = create_dependency_graph(
        source,
        all_series_targets(bindings, workbook=source),
        load_values=True,
    )

    write_workbook(
        graph,
        dest,
        series_bindings=bindings,
        bindings_workbook=source,
        include_bound_labels=False,
    )

    assert _cell_value(dest, "Sheet1", "B2") == 1.1
    assert _cell_value(dest, "Sheet1", "A2") is None
    assert _cell_value(dest, "Sheet1", "B1") is None


def test_write_workbook_projection_still_includes_bound_labels(tmp_path: Path) -> None:
    from excel_grapher.exporter import IdentityTransitCompression

    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 2024
    ws["B1"] = 2025
    ws["A2"] = 10
    ws["B2"] = "=A2"
    ws["B3"] = "=B2+1"
    wb.save(source)
    wb.close()

    bindings = validate_bindings_document(
        {
            "schema_version": "1.17.0",
            "concept_scheme": {
                "id": "writeback_proj_labels",
                "concepts": [
                    {"id": "TIME_PERIOD", "dtype": "int"},
                    {"id": "OBS_VALUE", "dtype": "number"},
                ],
            },
            "series": [
                {
                    "id": "out",
                    "sheet": "Sheet1",
                    "data_range": "Sheet1!B3",
                    "layout": "series",
                    "output": {"compute": {"name": "compute_out"}},
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
                            },
                        ],
                    },
                    "key": ["TIME_PERIOD"],
                }
            ],
        }
    )
    graph = create_dependency_graph(
        source,
        ["Sheet1!B3"],
        capture_dependency_provenance=True,
        load_values=True,
    )
    projection = IdentityTransitCompression().project(graph)
    assert "Sheet1!B2" not in projection

    write_workbook(
        projection,
        dest,
        series_bindings=bindings,
        bindings_workbook=source,
    )

    assert _cell_value(dest, "Sheet1", "B3") == "=A2+1"
    assert _cell_value(dest, "Sheet1", "B2") is None
    assert _cell_value(dest, "Sheet1", "A2") == 10
    assert _cell_value(dest, "Sheet1", "B1") == 2025
    assert _cell_value(dest, "Sheet1", "A1") is None


def test_bound_label_addresses_uses_actual_fill_sources(tmp_path: Path) -> None:
    source = tmp_path / "grouped_inputs.xlsx"
    write_grouped_matrix_workbook(source)
    bindings = validate_bindings_document(grouped_matrix_bindings_document())
    graph = create_dependency_graph(
        source,
        all_series_targets(bindings, workbook=source),
        load_values=True,
    )

    addresses = bound_label_addresses(bindings, workbook=source, graph=graph)

    assert addresses >= {
        "Inputs!A2",
        "Inputs!A3",
        "Inputs!A4",
        "Inputs!A6",
        "Inputs!A7",
        "Inputs!A8",
        "Inputs!C1",
        "Inputs!D1",
    }
    assert "Inputs!A5" not in addresses
    assert "Inputs!A1" not in addresses


def test_write_workbook_grouped_fill_labels(tmp_path: Path) -> None:
    source = tmp_path / "grouped_inputs.xlsx"
    dest = tmp_path / "out.xlsx"
    write_grouped_matrix_workbook(source)
    bindings = validate_bindings_document(grouped_matrix_bindings_document())
    graph = create_dependency_graph(
        source,
        all_series_targets(bindings, workbook=source),
        load_values=True,
    )

    write_workbook(
        graph,
        dest,
        series_bindings=bindings,
        bindings_workbook=source,
    )

    assert _cell_value(dest, "Inputs", "A2") == "Paris"
    assert _cell_value(dest, "Inputs", "A3") == "Revenue"
    assert _cell_value(dest, "Inputs", "A6") == "Moderate"
    assert _cell_value(dest, "Inputs", "C1") == 2024
    assert _cell_value(dest, "Inputs", "D1") == 2025
    assert _cell_value(dest, "Inputs", "A5") is None
    assert _cell_value(dest, "Inputs", "C3") == 1.1
