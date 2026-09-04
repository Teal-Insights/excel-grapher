"""Guards for duplicate series ids, A1 geometry slugs, and Python identifiers."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
import xlsxwriter

from excel_grapher.exporter.inverted_tree.catalog import build_catalog
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    validate_bindings_document,
    validate_series_bindings,
)


def _write_engine_workbook(path: Path, cells: dict[str, float]) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Engine")
    for address, value in cells.items():
        ws.write_number(address, value)
    wb.close()


def _engine_graph(tmp_path: Path, cells: dict[str, float] | None = None):
    cells = cells or {"A1": 1.0, "B1": 2.0}
    path = tmp_path / "engine.xlsx"
    _write_engine_workbook(path, cells)
    addresses = [f"Engine!{address}" for address in cells]
    graph = create_dependency_graph(path, addresses, load_values=True)
    return path, graph


def _scalar_series(
    series_id: str,
    data_range: str,
    *,
    direction: str = "output",
    series_context: dict[str, Any] | None = None,
) -> dict[str, Any]:
    sheet = data_range.split("!", 1)[0]
    entry: dict[str, Any] = {
        "id": series_id,
        "sheet": sheet,
        "data_range": data_range,
        "layout": "scalar",
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [],
        },
        "key": [],
    }
    if direction == "output":
        entry["output"] = {"compute": {"name": f"compute_{series_id}"}}
    elif direction == "input":
        entry["input"] = {"setter": {"name": f"set_{series_id}"}}
    elif direction == "internal":
        entry["internal"] = {}
    elif direction == "constant":
        entry["constant"] = {}
    else:
        raise ValueError(f"unknown direction {direction!r}")
    if series_context is not None:
        entry["series_context"] = series_context
    return entry


def _document(*series: dict[str, Any]) -> dict[str, Any]:
    return {
        "schema_version": "1.14.0",
        "concept_scheme": {
            "id": "series_id_guards",
            "concepts": [{"id": "OBS_VALUE", "dtype": "float"}],
        },
        "series": list(series),
    }


def test_build_catalog_rejects_duplicate_series_id(tmp_path: Path) -> None:
    path, _graph = _engine_graph(tmp_path)
    bindings = validate_bindings_document(
        _document(
            _scalar_series("shared_id", "Engine!A1", direction="input"),
            _scalar_series("shared_id", "Engine!B1", direction="input"),
        )
    )
    with pytest.raises(InvertedTreeExportError, match="Engine!A1") as excinfo:
        build_catalog(bindings, workbook=path)
    message = str(excinfo.value)
    assert "shared_id" in message
    assert "Engine!B1" in message


def test_validate_reports_duplicate_series_id(tmp_path: Path) -> None:
    path, graph = _engine_graph(tmp_path)
    bindings = validate_bindings_document(
        _document(
            _scalar_series("shared_id", "Engine!A1", direction="input"),
            _scalar_series("shared_id", "Engine!B1", direction="input"),
        )
    )
    report = validate_series_bindings(graph, bindings, workbook=path)
    duplicates = [issue for issue in report["issues"] if issue["code"] == "duplicate_series_id"]
    assert duplicates, report["issues"]
    assert duplicates[0]["level"] == "error"
    assert duplicates[0]["series_id"] == "shared_id"
    assert "Engine!A1" in duplicates[0]["message"]
    assert "Engine!B1" in duplicates[0]["message"]
    assert report["ok"] is False


def test_geometry_in_output_id_is_error(tmp_path: Path) -> None:
    path, graph = _engine_graph(tmp_path, {"A1": 1.0})
    bindings = validate_bindings_document(
        _document(
            _scalar_series(
                "constant_demography_bi723_ev920",
                "Engine!A1",
                direction="output",
            )
        )
    )
    report = validate_series_bindings(graph, bindings, workbook=path)
    geometry = [issue for issue in report["issues"] if issue["code"] == "geometry_in_id"]
    assert geometry, report["issues"]
    assert geometry[0]["level"] == "error"
    assert report["ok"] is False


def test_geometry_in_constant_id_is_warning(tmp_path: Path) -> None:
    path, graph = _engine_graph(tmp_path, {"A1": 1.0})
    bindings = validate_bindings_document(
        _document(
            _scalar_series(
                "constant_demography_bi723_ev920",
                "Engine!A1",
                direction="constant",
            )
        )
    )
    report = validate_series_bindings(graph, bindings, workbook=path)
    geometry = [issue for issue in report["issues"] if issue["code"] == "geometry_in_id"]
    assert geometry, report["issues"]
    assert geometry[0]["level"] == "warning"
    assert report["ok"] is True


def test_semantic_ids_without_a1_tokens_are_clean(tmp_path: Path) -> None:
    path, graph = _engine_graph(tmp_path)
    bindings = validate_bindings_document(
        _document(
            _scalar_series(
                "demography_total_population_medium",
                "Engine!A1",
                direction="output",
            ),
            _scalar_series(
                "co2_emissions_2050",
                "Engine!B1",
                direction="output",
            ),
        )
    )
    report = validate_series_bindings(graph, bindings, workbook=path)
    codes = {issue["code"] for issue in report["issues"]}
    assert "geometry_in_id" not in codes
    assert "invalid_python_id" not in codes
    assert report["ok"] is True


def test_sheet_prefix_without_cell_token_is_clean(tmp_path: Path) -> None:
    path, graph = _engine_graph(tmp_path, {"A1": 1.0})
    bindings = validate_bindings_document(
        _document(_scalar_series("macrofiscal_real_gdp", "Engine!A1", direction="output"))
    )
    report = validate_series_bindings(graph, bindings, workbook=path)
    codes = {issue["code"] for issue in report["issues"]}
    assert "geometry_in_id" not in codes
    assert report["ok"] is True


def test_trailing_cell_token_on_public_id_is_error(tmp_path: Path) -> None:
    path, graph = _engine_graph(tmp_path, {"A1": 1.0})
    bindings = validate_bindings_document(
        _document(_scalar_series("macrofiscal_ag67", "Engine!A1", direction="input"))
    )
    report = validate_series_bindings(graph, bindings, workbook=path)
    geometry = [issue for issue in report["issues"] if issue["code"] == "geometry_in_id"]
    assert geometry, report["issues"]
    assert geometry[0]["level"] == "error"
    assert report["ok"] is False


def test_geometry_in_series_context_is_warning(tmp_path: Path) -> None:
    path, graph = _engine_graph(tmp_path, {"A1": 1.0})
    bindings = validate_bindings_document(
        _document(
            _scalar_series(
                "baseline_indicator",
                "Engine!A1",
                direction="constant",
                series_context={"INDICATOR": "constant_baseline_d17_cp17"},
            )
        )
    )
    report = validate_series_bindings(graph, bindings, workbook=path)
    geometry = [issue for issue in report["issues"] if issue["code"] == "geometry_in_id"]
    assert geometry, report["issues"]
    assert geometry[0]["level"] == "warning"
    assert "INDICATOR" in geometry[0]["message"]
    assert "constant_baseline_d17_cp17" in geometry[0]["message"]
    assert report["ok"] is True


def test_python_keyword_series_id_is_invalid(tmp_path: Path) -> None:
    path, graph = _engine_graph(tmp_path, {"A1": 1.0})
    bindings = validate_bindings_document(
        _document(_scalar_series("class", "Engine!A1", direction="output"))
    )
    report = validate_series_bindings(graph, bindings, workbook=path)
    invalid = [issue for issue in report["issues"] if issue["code"] == "invalid_python_id"]
    assert invalid, report["issues"]
    assert invalid[0]["level"] == "error"
    assert report["ok"] is False
