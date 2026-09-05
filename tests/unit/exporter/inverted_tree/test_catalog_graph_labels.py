"""Catalog must not bind unused sheets when resolving off-graph key labels."""

from __future__ import annotations

from pathlib import Path
from typing import Any
from unittest.mock import patch

import fastpyxl
from fastpyxl.worksheet._reader import WorkSheetParser, WorksheetReader

from excel_grapher.exporter import CodeGenerator
from excel_grapher.exporter.inverted_tree.catalog import build_catalog
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import validate_bindings_document
from tests.unit.exporter.inverted_tree.helpers import write_workbook


def _label_workbook(path: Path) -> Path:
    return write_workbook(
        path,
        {
            "Store": {
                "A1": "country",
                "B1": 2009,
                "A2": "Afghanistan",
                "B2": 1.0,
                "C2": "=B2",
                "A50": "tail-should-not-be-parsed",
            },
            "Macrofiscal": {
                "A1": "unused",
                "B1": 0,
            },
        },
    )


def _label_bindings() -> dict[str, Any]:
    time_dim = {
        "id": "TIME_PERIOD",
        "concept": "TIME_PERIOD",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
    }
    country_dim = {
        "id": "COUNTRY",
        "concept": "COUNTRY",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "row_label", "label_column": "A", "read": "string"},
    }
    measure = {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }
    return {
        "schema_version": "1.13.0",
        "concept_scheme": {
            "id": "catalog_graph_labels",
            "concepts": [
                {"id": "TIME_PERIOD", "dtype": "int"},
                {"id": "COUNTRY", "dtype": "string"},
                {"id": "OBS_VALUE", "dtype": "number"},
            ],
        },
        "series": [
            {
                "id": "store",
                "sheet": "Store",
                "data_range": "Store!B2",
                "layout": "scalar",
                "constant": {},
                "structure": {"measure": measure, "dimensions": [country_dim, time_dim]},
                "key": ["COUNTRY", "TIME_PERIOD"],
            },
            {
                "id": "out",
                "sheet": "Store",
                "data_range": "Store!C2",
                "layout": "scalar",
                "output": {"compute": {"name": "compute_out"}},
                "structure": {"measure": measure, "dimensions": []},
                "key": [],
            },
        ],
    }


def test_generate_modules_does_not_bind_full_workbook(tmp_path: Path) -> None:
    path = _label_workbook(tmp_path / "inverted_tree_catalog_reopens_workbook.xlsx")
    bindings = validate_bindings_document(_label_bindings())
    graph = create_dependency_graph(path, targets=["Store!C2"], load_values=True)
    assert "Store!A2" not in graph
    assert "Store!B1" not in graph

    loads: list[dict[str, Any]] = []
    real_load = fastpyxl.load_workbook

    def wrapped_load(*args: Any, **kwargs: Any) -> Any:
        loads.append(dict(kwargs))
        return real_load(*args, **kwargs)

    with (
        patch("fastpyxl.load_workbook", side_effect=wrapped_load),
        patch.object(WorksheetReader, "bind_all") as bind_all,
        CodeGenerator(graph) as generator,
    ):
        modules = generator.generate_modules(
            series_bindings=bindings,
            bindings_workbook=path,
        )

    assert "api.py" in modules
    assert bind_all.call_count == 0
    assert all(call.get("read_only") is not False for call in loads)


def test_catalog_with_graph_resolves_off_graph_labels(tmp_path: Path) -> None:
    path = _label_workbook(tmp_path / "labels.xlsx")
    bindings = validate_bindings_document(_label_bindings())
    graph = create_dependency_graph(path, targets=["Store!C2"], load_values=True)
    catalog = build_catalog(bindings, workbook=path, graph=graph)
    assert catalog.get("store").domain[0].as_mapping() == {
        "COUNTRY": "Afghanistan",
        "TIME_PERIOD": 2009,
    }


def test_catalog_with_graph_does_not_stream_unused_sheet_body(tmp_path: Path) -> None:
    path = _label_workbook(tmp_path / "unused_sheet.xlsx")
    bindings = validate_bindings_document(_label_bindings())
    graph = create_dependency_graph(path, targets=["Store!C2"], load_values=True)
    streamed: list[str] = []
    original_parse = WorkSheetParser.parse

    def tracking_parse(self: WorkSheetParser) -> Any:
        streamed.append(str(getattr(self.source, "name", "")))
        yield from original_parse(self)

    with (
        patch.object(WorkSheetParser, "parse", new=tracking_parse),
        patch.object(WorksheetReader, "bind_all") as bind_all,
    ):
        build_catalog(bindings, workbook=path, graph=graph)

    assert bind_all.call_count == 0
    assert streamed
    assert all("sheet2" not in name.lower() for name in streamed)


def test_catalog_stops_used_sheet_parse_after_needed_label_rows(tmp_path: Path) -> None:
    path = _label_workbook(tmp_path / "used_sheet_tail.xlsx")
    bindings = validate_bindings_document(_label_bindings())
    graph = create_dependency_graph(path, targets=["Store!C2"], load_values=True)
    rows_seen: list[int] = []
    original_parse = WorkSheetParser.parse

    def tracking_parse(self: WorkSheetParser) -> Any:
        for row_idx, cells in original_parse(self):
            rows_seen.append(row_idx)
            yield row_idx, cells

    with patch.object(WorkSheetParser, "parse", new=tracking_parse):
        catalog = build_catalog(bindings, workbook=path, graph=graph)

    assert catalog.get("store").domain[0]["COUNTRY"] == "Afghanistan"
    assert rows_seen
    assert max(rows_seen) < 50
