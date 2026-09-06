"""Validate must apply exclude_rows / exclude_columns the same way resolve does (#594)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    resolve_series_binding,
    validate_series_bindings,
)
from excel_grapher.series_bindings.schema import validate_bindings_document
from excel_grapher.series_bindings.workflow import (
    all_series_targets,
    series_binding_public_addresses,
)


def _write_engine_column(path: Path) -> None:
    """Engine!B1 leaf, B2 formula, B3 leaf, with TIME_PERIOD labels in column A."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Engine")
    ws.write_number("A1", 2020)
    ws.write_number("A2", 2021)
    ws.write_number("A3", 2022)
    ws.write_number("B1", 1)
    ws.write_formula("B2", "=B1")
    ws.write_number("B3", 2)
    wb.close()


def _write_engine_row(path: Path) -> None:
    """Engine!A2 leaf, B2 formula, C2 leaf, with TIME_PERIOD headers in row 1."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Engine")
    ws.write_number("A1", 2020)
    ws.write_number("B1", 2021)
    ws.write_number("C1", 2022)
    ws.write_number("A2", 1)
    ws.write_formula("B2", "=A2")
    ws.write_number("C2", 2)
    wb.close()


def _series_doc(
    *,
    data_range: str = "Engine!B1:B3",
    exclude_rows: list[Any] | None = None,
    exclude_columns: list[Any] | None = None,
    constant: bool = False,
    time_bind: dict[str, Any] | None = None,
) -> dict[str, Any]:
    series: dict[str, Any] = {
        "id": "demo",
        "sheet": "Engine",
        "data_range": data_range,
        "layout": "series",
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "id": "TIME_PERIOD",
                    "concept": "TIME_PERIOD",
                    "dtype": "int",
                    "role": "key",
                    "scope": "cell",
                    "bind": time_bind
                    or {
                        "kind": "row_label",
                        "label_column": "A",
                        "read": "int",
                    },
                }
            ],
        },
        "key": ["TIME_PERIOD"],
    }
    if exclude_rows is not None:
        series["exclude_rows"] = exclude_rows
    if exclude_columns is not None:
        series["exclude_columns"] = exclude_columns
    if constant:
        series["constant"] = {}
    else:
        series["input"] = {
            "setter": {"name": "set_demo", "record_contract": "records", "strict": True}
        }
    return {"schema_version": "1.14.0", "workbook": "Book.xlsx", "series": [series]}


def _row_series_doc(**kwargs: Any) -> dict[str, Any]:
    return _series_doc(
        data_range="Engine!A2:C2",
        time_bind={"kind": "column_header", "header_row": 1, "read": "int"},
        **kwargs,
    )


def _graph_for(path: Path, targets: list[str]):
    return create_dependency_graph(path, targets, load_values=True)


def test_validate_exclude_rows_ignores_formula_hole(tmp_path: Path) -> None:
    wb_path = tmp_path / "Book.xlsx"
    _write_engine_column(wb_path)
    graph = _graph_for(wb_path, ["Engine!B1", "Engine!B2", "Engine!B3"])
    bindings = validate_bindings_document(_series_doc(exclude_rows=[2]))

    resolved = resolve_series_binding(graph, wb_path, bindings["series"][0])
    assert resolved["ok"] is True, resolved["issues"]
    assert {leaf["address"] for leaf in resolved["leaves"]} == {"Engine!B1", "Engine!B3"}

    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is True, report["issues"]
    codes = {issue["code"] for issue in report["issues"]}
    assert "non_leaf_input_overlap" not in codes


def test_validate_without_exclude_rows_reports_formula_overlap(tmp_path: Path) -> None:
    wb_path = tmp_path / "Book.xlsx"
    _write_engine_column(wb_path)
    graph = _graph_for(wb_path, ["Engine!B1", "Engine!B2", "Engine!B3"])
    bindings = validate_bindings_document(_series_doc())

    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is False
    assert any(issue["code"] == "non_leaf_input_overlap" for issue in report["issues"])


def test_validate_exclude_rows_ignores_constant_formula_hole(tmp_path: Path) -> None:
    wb_path = tmp_path / "Book.xlsx"
    _write_engine_column(wb_path)
    graph = _graph_for(wb_path, ["Engine!B1", "Engine!B2", "Engine!B3"])
    bindings = validate_bindings_document(_series_doc(exclude_rows=[2], constant=True))

    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is True, report["issues"]
    codes = {issue["code"] for issue in report["issues"]}
    assert "non_leaf_constant_overlap" not in codes


def test_validate_without_exclude_rows_reports_constant_formula_overlap(
    tmp_path: Path,
) -> None:
    wb_path = tmp_path / "Book.xlsx"
    _write_engine_column(wb_path)
    graph = _graph_for(wb_path, ["Engine!B1", "Engine!B2", "Engine!B3"])
    bindings = validate_bindings_document(_series_doc(constant=True))

    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is False
    assert any(issue["code"] == "non_leaf_constant_overlap" for issue in report["issues"])


def test_validate_exclude_columns_ignores_formula_hole(tmp_path: Path) -> None:
    wb_path = tmp_path / "Book.xlsx"
    _write_engine_row(wb_path)
    graph = _graph_for(wb_path, ["Engine!A2", "Engine!B2", "Engine!C2"])
    bindings = validate_bindings_document(_row_series_doc(exclude_columns=["B"]))

    resolved = resolve_series_binding(graph, wb_path, bindings["series"][0])
    assert resolved["ok"] is True, resolved["issues"]
    assert {leaf["address"] for leaf in resolved["leaves"]} == {"Engine!A2", "Engine!C2"}

    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is True, report["issues"]
    codes = {issue["code"] for issue in report["issues"]}
    assert "non_leaf_input_overlap" not in codes


def test_validate_exclude_rows_empty_remainder_is_empty_data_range(tmp_path: Path) -> None:
    wb_path = tmp_path / "Book.xlsx"
    _write_engine_column(wb_path)
    graph = _graph_for(wb_path, ["Engine!B1", "Engine!B2", "Engine!B3"])
    bindings = validate_bindings_document(_series_doc(exclude_rows=[1, "2:3"]))

    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is False
    assert any(issue["code"] == "empty_data_range" for issue in report["issues"])


def test_coverage_helpers_omit_excluded_formula_hole(tmp_path: Path) -> None:
    wb_path = tmp_path / "Book.xlsx"
    _write_engine_column(wb_path)
    graph = _graph_for(wb_path, ["Engine!B1", "Engine!B2", "Engine!B3"])
    bindings = validate_bindings_document(_series_doc(exclude_rows=[2]))

    targets = all_series_targets(bindings, workbook=wb_path)
    assert "Engine!B2" not in targets
    assert set(targets) == {"Engine!B1", "Engine!B3"}

    public = series_binding_public_addresses(graph, bindings, workbook=wb_path)
    assert "Engine!B2" not in public
    assert public == frozenset({"Engine!B1", "Engine!B3"})
