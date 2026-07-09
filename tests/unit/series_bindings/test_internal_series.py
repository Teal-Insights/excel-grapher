"""Tests for internal (non-I/O) series bindings."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    SeriesBindingsSchemaError,
    derive_input_series,
    derive_internal_series,
    derive_output_series,
    emit_series_bindings_block,
    has_internal_direction,
    load_series_bindings,
    merge_series_binding_documents,
    parse_bindings_file,
    resolve_series_binding,
    validate_bindings_document,
    validate_series_bindings,
)
from tests.paths import SERIES_BINDINGS_FIXTURES as FIXTURES
from tests.unit.series_bindings.test_input_override import (
    _manual_override_graph,
    _write_override_workbook,
)


def _internal_series_doc(**series_overrides: Any) -> dict[str, Any]:
    series: dict[str, Any] = {
        "id": "engine_primary_balance",
        "sheet": "Engine",
        "data_range": "Engine!B2:D2",
        "layout": "series",
        "internal": {},
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
                    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                }
            ],
        },
        "key": ["TIME_PERIOD"],
        "series_context": {"INDICATOR": "Primary balance"},
        "validation": {
            "intersect_graph_formulas": True,
            "require_unique_key": True,
        },
    }
    series.update(series_overrides)
    return {
        "schema_version": "1.7.0",
        "concept_scheme": {
            "id": "example_model",
            "concepts": [
                {"id": "TIME_PERIOD", "dtype": "int"},
                {"id": "INDICATOR", "dtype": "string"},
            ],
        },
        "series": [series],
    }


def test_schema_accepts_internal_only_series() -> None:
    doc = _internal_series_doc()
    bindings = validate_bindings_document(doc)
    series = bindings["series"][0]
    assert has_internal_direction(series)
    assert "internal" in series


def test_schema_rejects_series_without_any_direction() -> None:
    doc = _internal_series_doc()
    del doc["series"][0]["internal"]
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_rejects_internal_with_input() -> None:
    doc = _internal_series_doc(
        input={"setter": {"name": "set_engine_primary_balance"}},
    )
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_rejects_internal_with_output() -> None:
    doc = _internal_series_doc(
        output={"compute": {"name": "compute_engine_primary_balance"}},
    )
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_mcve_internal_document_fails_on_1_5_0() -> None:
    doc = _internal_series_doc()
    del doc["series"][0]["internal"]
    doc["schema_version"] = "1.5.0"
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_resolve_internal_series_includes_formula_cells(tmp_path: Path) -> None:
    wb_path = tmp_path / "formula_override.xlsx"
    _write_override_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        ["Engine!B2", "Engine!C2", "Engine!D2", "Output!B3", "Output!C3", "Output!D3"],
        load_values=True,
    )
    series = _internal_series_doc()["series"][0]

    resolved = resolve_series_binding(graph, wb_path, series, direction="internal")

    assert resolved["ok"] is True
    assert [leaf["address"] for leaf in resolved["leaves"]] == [
        "Engine!B2",
        "Engine!C2",
        "Engine!D2",
    ]
    assert resolved["leaves"][1]["key"] == {"TIME_PERIOD": 2}
    assert resolved["leaves"][1]["record"]["INDICATOR"] == "Primary balance"


def test_resolve_internal_series_skips_non_formula_cells(tmp_path: Path) -> None:
    wb_path = tmp_path / "formula_override.xlsx"
    _write_override_workbook(wb_path)
    graph = _manual_override_graph()
    series = _internal_series_doc(
        sheet="Inputs",
        data_range="Inputs!A1",
        structure={
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
                    "bind": {"kind": "constant", "value": 1},
                }
            ],
        },
    )["series"][0]

    resolved = resolve_series_binding(graph, wb_path, series, direction="internal")

    assert resolved["leaves"] == []
    assert any(i["code"] == "no_resolved_cells" for i in resolved["issues"])


def test_resolve_internal_series_enforces_unique_keys(tmp_path: Path) -> None:
    wb_path = tmp_path / "dup.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Engine")
    ws.write_number("A1", 1)
    ws.write_number("B1", 1)
    ws.write_formula("B2", "=A1+1")
    ws.write_formula("C2", "=A1+2")
    wb.close()

    graph = create_dependency_graph(wb_path, ["Engine!B2", "Engine!C2"], load_values=True)
    series = _internal_series_doc(
        data_range="Engine!B2:C2",
        structure={
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
                    "bind": {"kind": "constant", "value": 1},
                }
            ],
        },
    )["series"][0]

    resolved = resolve_series_binding(graph, wb_path, series, direction="internal")

    assert resolved["ok"] is False
    assert resolved["requires_address"] is True
    assert any(i["code"] == "duplicate_key" for i in resolved["issues"])


def test_derive_internal_series_from_fixture(tmp_path: Path) -> None:
    wb_path = tmp_path / "formula_override.xlsx"
    _write_override_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        ["Engine!B2", "Engine!C2", "Engine!D2", "Output!B3", "Output!C3", "Output!D3"],
        load_values=True,
    )
    bindings = load_series_bindings(FIXTURES / "internal_engine_row.yaml")

    internal_series = derive_internal_series(graph, bindings, workbook=wb_path)

    assert len(internal_series) == 1
    series = internal_series[0]
    assert series["id"] == "engine_primary_balance"
    assert series["key_fields"] == ["TIME_PERIOD"]
    assert [cell["address"] for cell in series["cells"]] == [
        "Engine!B2",
        "Engine!C2",
        "Engine!D2",
    ]


def test_derive_internal_series_does_not_emit_codegen(tmp_path: Path) -> None:
    wb_path = tmp_path / "formula_override.xlsx"
    _write_override_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Engine!B2"], load_values=True)
    bindings = validate_bindings_document(_internal_series_doc())
    lines = emit_series_bindings_block(graph, wb_path, bindings)
    assert lines == []


def test_input_and_output_derive_skip_internal_series(tmp_path: Path) -> None:
    wb_path = tmp_path / "formula_override.xlsx"
    _write_override_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        ["Engine!B2", "Engine!C2", "Engine!D2"],
        load_values=True,
    )
    bindings = validate_bindings_document(_internal_series_doc())

    assert derive_input_series(graph, bindings, workbook=wb_path) == []
    assert derive_output_series(graph, bindings, workbook=wb_path) == []


def test_load_merged_internal_shard_directory(tmp_path: Path) -> None:
    shard_dir = tmp_path / "shards"
    shard_dir.mkdir()
    (shard_dir / "internals.bindings.yaml").write_text(
        (FIXTURES / "internal_engine_row.yaml").read_text(encoding="utf-8"),
        encoding="utf-8",
    )
    bindings = load_series_bindings(shard_dir)
    assert bindings["series"][0]["id"] == "engine_primary_balance"
    assert has_internal_direction(bindings["series"][0])


def test_merge_internal_direction_blocks_across_shards() -> None:
    base_series = parse_bindings_file(FIXTURES / "internal_engine_row.yaml")["series"][0]
    structural = {k: v for k, v in base_series.items() if k != "internal"}
    left_doc = {
        "schema_version": "1.7.0",
        "series": [{**structural, "internal": {}}],
    }
    right_doc = {
        "schema_version": "1.7.0",
        "series": [{**structural, "internal": {}}],
    }
    merged = merge_series_binding_documents([left_doc, right_doc])
    assert has_internal_direction(merged["series"][0])


def test_validate_internal_series_requires_formula_overlap(tmp_path: Path) -> None:
    wb_path = tmp_path / "formula_override.xlsx"
    _write_override_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!A1"], load_values=True)
    bindings = validate_bindings_document(_internal_series_doc())

    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    codes = {issue["code"] for issue in report["issues"]}
    assert "no_formula_internal_targets" in codes
