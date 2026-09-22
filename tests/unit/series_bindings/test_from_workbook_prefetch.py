"""Issue 964 — from_workbook compilation must not re-parse a sheet per cell."""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest
import xlsxwriter
from fastpyxl.utils.cell import coordinate_from_string

from excel_grapher.core.cell_types import normalize_cell_type_env_key
from excel_grapher.exporter.inverted_tree.emit import generate_inverted_tree_modules
from excel_grapher.exporter.inverted_tree.input_annotations import public_input_annotations
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from excel_grapher.series_bindings import validate_bindings_document
from excel_grapher.series_bindings.domains import SeriesDomainIndex
from excel_grapher.series_bindings.resolve import _stream_sheet_values, _WorkbookValues
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    series_entry,
    write_workbook,
)


def _constant_series(series_id: str, data_range: str, sheet: str) -> dict[str, Any]:
    return {
        "id": series_id,
        "sheet": sheet,
        "data_range": data_range,
        "layout": "scalar",
        "constant": {},
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


def _flag_series() -> dict[str, Any]:
    return {
        "id": "flag",
        "sheet": "Calc",
        "data_range": "Calc!A1",
        "layout": "scalar",
        "input": {},
        "domain": {"enum": ["On", "Off"]},
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


def _write_split_constants(path: Path) -> None:
    workbook = xlsxwriter.Workbook(path)
    data = workbook.add_worksheet("Data")
    for row in range(1, 21):
        data.write_string(row - 1, 0, f"label-{row}")
    climate = workbook.add_worksheet("Climate Data")
    climate.write_string("A1", "north")
    climate.write_string("B3", "corner")
    calc = workbook.add_worksheet("Calc")
    calc.write_string("A1", "On")
    calc.write_formula("B1", "=A1")
    workbook.close()


def _split_bindings() -> dict[str, Any]:
    return validate_bindings_document(
        {
            "schema_version": "1.19.0",
            "series": [
                _flag_series(),
                _constant_series("top", "Data!A1:A10", "Data"),
                _constant_series("bottom", "Data!A11:A20", "Data"),
                _constant_series("wide", "'Climate Data'!A1:B3", "Climate Data"),
            ],
        }
    )


def _record_streams(monkeypatch: pytest.MonkeyPatch) -> list[tuple[str, frozenset[str]]]:
    calls: list[tuple[str, frozenset[str]]] = []

    def counting(
        workbook: Any,
        sheet: str,
        wanted: set[str],
        *,
        data_only: bool = True,
    ) -> dict[str, Any]:
        calls.append((sheet, frozenset(wanted)))
        return _stream_sheet_values(workbook, sheet, wanted, data_only=data_only)

    monkeypatch.setattr(
        "excel_grapher.series_bindings.resolve._stream_sheet_values",
        counting,
    )
    return calls


def test_truthiness_and_input_annotations_do_not_compile_constants(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = tmp_path / "book.xlsx"
    _write_split_constants(path)
    graph = create_dependency_graph(path, ["Calc!B1"], load_values=True)
    index = graph.attach_domains(_split_bindings(), workbook=path)
    catalog = SimpleNamespace(
        series={
            "flag": SimpleNamespace(
                series_id="flag",
                direction="input",
                cells=("Calc!A1",),
                graph_cells=None,
                raw={"id": "flag"},
                python_dtype="str",
                dtype="string",
            )
        }
    )
    calls = _record_streams(monkeypatch)

    assert index
    annotations = public_input_annotations(catalog, graph)

    assert calls == []
    assert index._expanded is None
    assert annotations == {"flag": 'Literal["Off", "On"]'}


def test_empty_domain_index_is_falsy_without_compiling(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    workbook = xlsxwriter.Workbook(path)
    sheet = workbook.add_worksheet("Inputs")
    sheet.write_number("C5", 1)
    workbook.close()
    bindings = validate_bindings_document(
        {
            "schema_version": "1.19.0",
            "series": [
                {
                    "id": "engine_year_labels",
                    "sheet": "Inputs",
                    "data_range": "Inputs!C5",
                    "layout": "scalar",
                    "internal": {},
                    "structure": {
                        "measure": {
                            "concept": "OBS_VALUE",
                            "dtype": "int",
                            "bind": {"kind": "data_cell", "read": "int"},
                        },
                        "dimensions": [],
                    },
                    "key": [],
                }
            ],
        }
    )
    index = SeriesDomainIndex.from_bindings(bindings, workbook=path)
    assert index._expanded is None
    assert not index
    assert index._expanded is None


def test_materialize_prefetches_each_off_graph_sheet_once(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = tmp_path / "book.xlsx"
    _write_split_constants(path)
    graph = create_dependency_graph(path, ["Calc!B1"], load_values=True)
    index = graph.attach_domains(_split_bindings(), workbook=path)
    assert "Data!A1" not in graph
    calls = _record_streams(monkeypatch)

    size = len(index)
    assert size == 23
    assert len(index) == size

    assert [(sheet, max(_row(coord) for coord in wanted)) for sheet, wanted in calls] == [
        ("Data", 20),
        ("Climate Data", 3),
    ]
    assert _enum_values(index, "Data!A1") == frozenset({"label-1"})
    assert _enum_values(index, "Data!A20") == frozenset({"label-20"})
    assert _enum_values(index, "'Climate Data'!B3") == frozenset({"corner"})
    assert normalize_cell_type_env_key("'Climate Data'!A2") not in index


def test_materialize_skips_workbook_reads_for_graph_nodes(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = tmp_path / "book.xlsx"
    workbook = xlsxwriter.Workbook(path)
    data = workbook.add_worksheet("Data")
    for row in range(1, 11):
        data.write_string(row - 1, 0, f"label-{row}")
    workbook.close()
    graph = create_dependency_graph(path, ["Data!A1"], load_values=True)
    bindings = validate_bindings_document(
        {
            "schema_version": "1.19.0",
            "series": [_constant_series("labels", "Data!A1:A10", "Data")],
        }
    )
    index = graph.attach_domains(bindings, workbook=path)
    calls = _record_streams(monkeypatch)

    assert len(index) == 10

    assert len(calls) == 1
    sheet, wanted = calls[0]
    assert sheet == "Data"
    assert "A1" not in wanted
    assert wanted == {f"A{row}" for row in range(2, 11)}
    assert _enum_values(index, "Data!A1") == frozenset({"label-1"})
    assert _enum_values(index, "Data!A10") == frozenset({"label-10"})


def test_prefetch_caches_requested_coordinates_only(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = tmp_path / "book.xlsx"
    workbook = xlsxwriter.Workbook(path)
    sheet = workbook.add_worksheet("Sheet")
    sheet.write_string("A1", "a")
    sheet.write_string("B1", "bee")
    sheet.write_string("A4", "later")
    workbook.close()
    calls: list[int] = []
    original = _stream_sheet_values

    def counting(
        workbook_obj: Any,
        sheet_name: str,
        wanted: set[str],
        *,
        data_only: bool = True,
    ) -> dict[str, Any]:
        calls.append(max(_row(coord) for coord in wanted))
        return original(workbook_obj, sheet_name, wanted, data_only=data_only)

    monkeypatch.setattr("excel_grapher.series_bindings.resolve._stream_sheet_values", counting)
    reader = _WorkbookValues(path)
    reader.prefetch(["Sheet!A1", "Sheet!B1", "Sheet!A4"])
    assert calls == [4]
    assert reader.read("Sheet!A1") == "a"
    assert reader.read("Sheet!B1") == "bee"
    assert reader.read("Sheet!A4") == "later"
    assert calls == [4]
    assert reader.read("Sheet!C1") is None
    assert calls == [4, 1]
    reader.close()
    assert reader.read("Sheet!A1") == "a"
    assert calls == [4, 1, 1]


def test_successive_off_graph_lookups_stream_each_sheet_once(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = tmp_path / "book.xlsx"
    _write_split_constants(path)
    graph = create_dependency_graph(path, ["Calc!B1"], load_values=True)
    index = graph.attach_domains(_split_bindings(), workbook=path)
    calls = _record_streams(monkeypatch)

    assert _enum_values(index, "Data!A1") == frozenset({"label-1"})
    assert [(sheet, max(_row(coord) for coord in wanted)) for sheet, wanted in calls] == [
        ("Data", 20),
        ("Climate Data", 3),
    ]
    assert index._expanded is None
    assert len(index._memo) == 1

    assert _enum_values(index, "Data!A20") == frozenset({"label-20"})
    assert _enum_values(index, "'Climate Data'!B3") == frozenset({"corner"})
    assert len(calls) == 2
    assert index._expanded is None
    assert len(index._memo) == 3


def test_blank_from_workbook_truthiness_does_not_flip(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    workbook = xlsxwriter.Workbook(path)
    workbook.add_worksheet("Data")
    workbook.close()
    bindings = validate_bindings_document(
        {
            "schema_version": "1.19.0",
            "series": [_constant_series("labels", "Data!A1:A3", "Data")],
        }
    )
    index = SeriesDomainIndex.from_bindings(bindings, workbook=path)
    assert index._expanded is None
    assert index
    assert len(index) == 0
    assert index


def test_codegen_does_not_materialize_constant_domains(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = write_workbook(
        tmp_path / "book.xlsx",
        {
            "Data": {"A1": "label-1"},
            "Calc": {"A1": "On", "B1": "=A1"},
        },
    )
    bindings = validate_bindings_document(
        bindings_document(
            series_entry(
                "flag",
                "Calc!A1",
                direction="input",
                dtype="string",
                domain={"enum": ["On", "Off"]},
            ),
            series_entry("labels", "Data!A1", direction="constant", dtype="string"),
            series_entry(
                "result",
                "Calc!B1",
                direction="output",
                dtype="string",
                compute_name="compute_result",
            ),
            schema_version="1.19.0",
        )
    )
    graph = create_dependency_graph(
        path,
        ["Calc!B1"],
        load_values=True,
        dynamic_refs=DynamicRefConfig.from_bindings(bindings, path),
    )
    assert isinstance(graph.cell_type_env, SeriesDomainIndex)
    assert graph.cell_type_env._expanded is None
    calls = {"n": 0}
    original = SeriesDomainIndex._materialize

    def counting(self: SeriesDomainIndex) -> dict[str, Any]:
        calls["n"] += 1
        return original(self)

    monkeypatch.setattr(SeriesDomainIndex, "_materialize", counting)
    modules = generate_inverted_tree_modules(
        graph, series_bindings=bindings, bindings_workbook=path
    )
    assert calls["n"] == 0
    assert graph.cell_type_env._expanded is None
    assert 'Literal["Off", "On"]' in modules["model.py"]


def _enum_values(index: SeriesDomainIndex, address: str) -> frozenset[object]:
    cell = index[normalize_cell_type_env_key(address)]
    assert cell.enum is not None
    return cell.enum.values


def _row(coord: str) -> int:
    return int(coordinate_from_string(coord)[1])
