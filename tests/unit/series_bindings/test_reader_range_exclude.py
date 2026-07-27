"""Range readers must honour exclude_rows / exclude_columns (#453).

Keyed readers already filter via resolve; `read_*_range` historically emitted
`xl_range(ctx, data_range)` and ignored exclusions — including complementary
interleaved matrix bindings that then returned identical full blocks.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import xlsxwriter

from excel_grapher.exporter import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    expand_data_range,
    resolve_series_binding,
    validate_bindings_document,
)
from excel_grapher.series_bindings.reader_index import build_reader_index
from excel_grapher.series_bindings.setter_codegen import (
    emit_reader_range_function,
    emit_readers_block,
)


def _interleaved_matrix_doc() -> dict[str, Any]:
    """Two matrix series sharing Risks!B2:D5 with complementary exclude_rows."""
    structure = {
        "measure": {
            "concept": "OBS_VALUE",
            "dtype": "float",
            "bind": {"kind": "data_cell", "read": "float"},
        },
        "dimensions": [
            {
                "id": "BAND",
                "concept": "BAND",
                "dtype": "string",
                "role": "key",
                "scope": "cell",
                "bind": {
                    "kind": "row_label",
                    "label_column": "A",
                    "fill": True,
                    "read": "string",
                },
            },
            {
                "id": "TIME_PERIOD",
                "concept": "TIME_PERIOD",
                "dtype": "int",
                "role": "key",
                "scope": "cell",
                "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
            },
        ],
    }
    return {
        "schema_version": "1.12.0",
        "concept_scheme": {
            "id": "mcve",
            "concepts": [
                {"id": "OBS_VALUE", "name": "Observation value", "dtype": "number"},
                {"id": "TIME_PERIOD", "name": "Time period", "dtype": "int"},
                {"id": "BAND", "name": "Band", "dtype": "string"},
            ],
        },
        "series": [
            {
                "id": "revenue_shocks",
                "sheet": "Risks",
                "data_range": "Risks!B2:D5",
                "layout": "matrix",
                "exclude_rows": [3, 5],
                "input": {"setter": {"name": "set_revenue_shocks"}},
                "structure": structure,
                "key": ["BAND", "TIME_PERIOD"],
            },
            {
                "id": "expenditure_shocks",
                "sheet": "Risks",
                "data_range": "Risks!B2:D5",
                "layout": "matrix",
                "exclude_rows": [2, 4],
                "input": {"setter": {"name": "set_expenditure_shocks"}},
                "structure": structure,
                "key": ["BAND", "TIME_PERIOD"],
            },
        ],
    }


def _write_interleaved_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Risks")
    for col, year in enumerate([2030, 2031, 2032], start=1):
        ws.write(0, col, year)
    for row, label in enumerate(["rev", "exp", "rev", "exp"], start=1):
        ws.write(row, 0, label)
        for col in range(1, 4):
            ws.write_number(row, col, float(row * 10 + col))
    ws.write_formula("E1", "=SUM(B2:D5)")
    wb.close()


def _write_edge_trim_workbook(path: Path) -> None:
    """Contiguous block B2:D5 where excluding edge rows/cols yields one rectangle."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Demo")
    for col, year in enumerate([2028, 2029, 2030], start=1):
        ws.write(0, col, year)
    for row, label in enumerate(["A", "B", "C", "D"], start=1):
        ws.write(row, 0, label)
        for col in range(1, 4):
            ws.write_number(row, col, float(row * 10 + col))
    ws.write_formula("E1", "=SUM(B2:D5)")
    wb.close()


def _matrix_series(
    *,
    series_id: str = "demo",
    data_range: str = "Demo!B2:D5",
    exclude_rows: list[Any] | None = None,
    exclude_columns: list[Any] | None = None,
) -> dict[str, Any]:
    series: dict[str, Any] = {
        "id": series_id,
        "sheet": "Demo",
        "data_range": data_range,
        "layout": "matrix",
        "input": {"setter": {"name": f"set_{series_id}"}},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "id": "BAND",
                    "concept": "BAND",
                    "dtype": "string",
                    "role": "key",
                    "scope": "cell",
                    "bind": {
                        "kind": "row_label",
                        "label_column": "A",
                        "read": "string",
                    },
                },
                {
                    "id": "TIME_PERIOD",
                    "concept": "TIME_PERIOD",
                    "dtype": "int",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                },
            ],
        },
        "key": ["BAND", "TIME_PERIOD"],
    }
    if exclude_rows is not None:
        series["exclude_rows"] = exclude_rows
    if exclude_columns is not None:
        series["exclude_columns"] = exclude_columns
    return {
        "schema_version": "1.12.0",
        "series": [series],
    }


def _fn_body(source: str, name: str) -> str | None:
    match = re.search(rf"^def {name}\(.*?(?=^def |\Z)", source, re.S | re.M)
    return match.group(0).rstrip() if match else None


def test_interleaved_exclude_rows_omits_misleading_range_readers(tmp_path: Path) -> None:
    """Non-contiguous exclusions must not emit identical full-block *_range helpers."""
    wb_path = tmp_path / "mcve.xlsx"
    _write_interleaved_workbook(wb_path)
    bindings = validate_bindings_document(_interleaved_matrix_doc())
    graph = create_dependency_graph(wb_path, expand_data_range("Risks!B2:D5"), load_values=True)

    source = "\n".join(emit_readers_block(graph, wb_path, bindings))
    assert "def read_revenue_shocks(" in source
    assert "def read_expenditure_shocks(" in source
    assert "def read_revenue_shocks_range(" not in source
    assert "def read_expenditure_shocks_range(" not in source
    assert "xl_range(ctx, 'Risks!B2:D5')" not in source


def test_contiguous_exclude_rows_emits_narrowed_range(tmp_path: Path) -> None:
    wb_path = tmp_path / "trim.xlsx"
    _write_edge_trim_workbook(wb_path)
    bindings = validate_bindings_document(
        _matrix_series(exclude_rows=[5], data_range="Demo!B2:D5")
    )
    graph = create_dependency_graph(wb_path, expand_data_range("Demo!B2:D5"), load_values=True)
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_reader_range_function(series, resolved, workbook=wb_path)
    source = "\n".join(lines)
    assert "def read_demo_range(" in source
    assert "return xl_range(ctx, 'Demo!B2:D4')" in source
    assert "Demo!B2:D5" not in source


def test_contiguous_exclude_columns_emits_narrowed_range(tmp_path: Path) -> None:
    wb_path = tmp_path / "trim.xlsx"
    _write_edge_trim_workbook(wb_path)
    bindings = validate_bindings_document(
        _matrix_series(exclude_columns=["D"], data_range="Demo!B2:D5")
    )
    graph = create_dependency_graph(wb_path, expand_data_range("Demo!B2:D5"), load_values=True)
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_reader_range_function(series, resolved, workbook=wb_path)
    source = "\n".join(lines)
    assert "return xl_range(ctx, 'Demo!B2:C5')" in source


def test_hole_exclude_rows_omits_range_reader(tmp_path: Path) -> None:
    """Excluding a middle row leaves a non-rectangular selection — no *_range helper."""
    wb_path = tmp_path / "hole.xlsx"
    _write_edge_trim_workbook(wb_path)
    bindings = validate_bindings_document(
        _matrix_series(exclude_rows=[3], data_range="Demo!B2:D5")
    )
    graph = create_dependency_graph(wb_path, expand_data_range("Demo!B2:D5"), load_values=True)
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    assert emit_reader_range_function(series, resolved, workbook=wb_path) == []


def test_reader_index_keys_narrowed_contiguous_range(tmp_path: Path) -> None:
    wb_path = tmp_path / "trim.xlsx"
    _write_edge_trim_workbook(wb_path)
    bindings = validate_bindings_document(
        _matrix_series(exclude_rows=[5], exclude_columns=["D"], data_range="Demo!B2:D5")
    )
    graph = create_dependency_graph(wb_path, expand_data_range("Demo!B2:D5"), load_values=True)
    index = build_reader_index(graph, bindings, workbook=wb_path)

    assert "Demo!B2:D5" not in index["ranges"]
    assert "Demo!B2:C4" in index["ranges"]
    assert index["ranges"]["Demo!B2:C4"]["reader"] == "read_demo_range"
    assert "Demo!B2:D5" not in index["ambiguous"]


def test_codegen_modules_skip_interleaved_range_readers(tmp_path: Path) -> None:
    wb_path = tmp_path / "mcve.xlsx"
    _write_interleaved_workbook(wb_path)
    bindings = validate_bindings_document(_interleaved_matrix_doc())
    graph = create_dependency_graph(wb_path, ["Risks!E1"], load_values=True)
    with CodeGenerator(graph) as gen:
        modules = gen.generate_modules(
            ["Risks!E1"], series_bindings=bindings, bindings_workbook=wb_path
        )
    readers = modules["_readers.py"]
    assert _fn_body(readers, "read_revenue_shocks_range") is None
    assert _fn_body(readers, "read_expenditure_shocks_range") is None
    assert "def read_revenue_shocks(" in readers
    assert "def read_expenditure_shocks(" in readers
