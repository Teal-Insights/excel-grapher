"""Tests for constant (reader-only graph-leaf) series bindings."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    SeriesBindingsSchemaError,
    build_reader_index,
    derive_constant_series,
    derive_input_series,
    derive_internal_series,
    derive_output_series,
    emit_readers_block,
    emit_series_bindings_block,
    has_constant_direction,
    has_input_direction,
    load_series_bindings,
    merge_series_binding_documents,
    parse_bindings_file,
    resolve_reader_ref,
    resolve_series_binding,
    validate_bindings_document,
    validate_series_bindings,
)
from excel_grapher.series_bindings.workflow import reader_names, setter_names


def _write_constant_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Engine")
    ws.write_number("C5", 2021)
    ws.write_formula("C10", "=IF(C5>=2020,1,0)")
    inputs = wb.add_worksheet("Inputs")
    inputs.write_number("B21", 2020)
    wb.close()


def _constant_series_doc(**series_overrides: Any) -> dict[str, Any]:
    series: dict[str, Any] = {
        "id": "shock_year_anchor",
        "sheet": "Engine",
        "data_range": "Engine!C5",
        "layout": "scalar",
        "constant": {},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [],
        },
        "key": [],
        "validation": {
            "intersect_graph_leaves": True,
            "require_unique_key": True,
        },
    }
    series.update(series_overrides)
    return {
        "schema_version": "1.11.0",
        "concept_scheme": {
            "id": "example_model",
            "concepts": [{"id": "OBS_VALUE", "dtype": "float"}],
        },
        "series": [series],
    }


def test_schema_accepts_constant_only_series() -> None:
    doc = _constant_series_doc()
    bindings = validate_bindings_document(doc)
    series = bindings["series"][0]
    assert has_constant_direction(series)
    assert "constant" in series


def test_schema_rejects_series_without_any_direction() -> None:
    doc = _constant_series_doc()
    del doc["series"][0]["constant"]
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_rejects_constant_with_input() -> None:
    doc = _constant_series_doc(
        input={"setter": {"name": "set_shock_year_anchor"}},
    )
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_rejects_constant_with_output() -> None:
    doc = _constant_series_doc(
        output={"compute": {"name": "compute_shock_year_anchor"}},
    )
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_rejects_constant_with_internal() -> None:
    doc = _constant_series_doc(internal={})
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_rejects_constant_with_legacy_setter() -> None:
    doc = _constant_series_doc(setter={"name": "set_shock_year_anchor"})
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_accepts_constant_reader_name_override() -> None:
    doc = _constant_series_doc(constant={"reader": {"name": "read_engine_c5"}})
    bindings = validate_bindings_document(doc)
    assert bindings["series"][0]["constant"]["reader"]["name"] == "read_engine_c5"


def test_resolve_constant_series_includes_leaf_cells(tmp_path: Path) -> None:
    wb_path = tmp_path / "constants.xlsx"
    _write_constant_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Engine!C10"], load_values=True)
    series = _constant_series_doc()["series"][0]

    resolved = resolve_series_binding(graph, wb_path, series, direction="constant")

    assert resolved["ok"]
    assert [leaf["address"] for leaf in resolved["leaves"]] == ["Engine!C5"]


def test_validate_constant_series_requires_leaf_overlap(tmp_path: Path) -> None:
    wb_path = tmp_path / "constants.xlsx"
    _write_constant_workbook(wb_path)
    # Graph only includes the formula cell; C5 is not a leaf in this extract.
    graph = create_dependency_graph(wb_path, ["Engine!C10"], load_values=False)
    # Drop C5 from the graph by targeting only Inputs!B21 (unrelated leaf).
    graph = create_dependency_graph(wb_path, ["Inputs!B21"], load_values=True)
    bindings = validate_bindings_document(_constant_series_doc())

    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    codes = {issue["code"] for issue in report["issues"]}
    assert "no_leaf_constant_targets" in codes


def test_validate_constant_series_rejects_formula_cells(tmp_path: Path) -> None:
    wb_path = tmp_path / "constants.xlsx"
    _write_constant_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Engine!C10"], load_values=True)
    bindings = validate_bindings_document(
        _constant_series_doc(data_range="Engine!C10", sheet="Engine")
    )

    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    codes = {issue["code"] for issue in report["issues"]}
    assert "non_leaf_constant_overlap" in codes


def test_derive_constant_series(tmp_path: Path) -> None:
    wb_path = tmp_path / "constants.xlsx"
    _write_constant_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Engine!C10"], load_values=True)
    bindings = validate_bindings_document(_constant_series_doc())

    constant_series = derive_constant_series(graph, bindings, workbook=wb_path)

    assert len(constant_series) == 1
    series = constant_series[0]
    assert series["id"] == "shock_year_anchor"
    assert series["reader_name"] == "read_shock_year_anchor"
    assert [cell["address"] for cell in series["cells"]] == ["Engine!C5"]


def test_input_output_internal_derive_skip_constant_series(tmp_path: Path) -> None:
    wb_path = tmp_path / "constants.xlsx"
    _write_constant_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Engine!C10"], load_values=True)
    bindings = validate_bindings_document(_constant_series_doc())

    assert derive_input_series(graph, bindings, workbook=wb_path) == []
    assert derive_output_series(graph, bindings, workbook=wb_path) == []
    assert derive_internal_series(graph, bindings, workbook=wb_path) == []


def test_codegen_emits_reader_without_setter(tmp_path: Path) -> None:
    wb_path = tmp_path / "constants.xlsx"
    _write_constant_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Engine!C10"], load_values=True)
    bindings = validate_bindings_document(_constant_series_doc())

    lines = emit_series_bindings_block(graph, wb_path, bindings)
    text = "\n".join(lines)

    assert "def read_shock_year_anchor(" in text
    assert "def set_shock_year_anchor(" not in text
    assert "def compute_shock_year_anchor(" not in text


def test_emit_readers_block_includes_constant_series(tmp_path: Path) -> None:
    wb_path = tmp_path / "constants.xlsx"
    _write_constant_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Engine!C10"], load_values=True)
    bindings = validate_bindings_document(_constant_series_doc())

    text = "\n".join(emit_readers_block(graph, wb_path, bindings))
    assert "def read_shock_year_anchor(" in text
    assert "_LEAF_INDEX_SHOCK_YEAR_ANCHOR" in text or "xl_cell(ctx, 'Engine!C5')" in text


def test_reader_index_includes_constant_leaves(tmp_path: Path) -> None:
    wb_path = tmp_path / "constants.xlsx"
    _write_constant_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Engine!C10"], load_values=True)
    bindings = validate_bindings_document(_constant_series_doc())

    index = build_reader_index(graph, bindings, workbook=wb_path)
    assert "Engine!C5" in index["leaves"]
    assert index["leaves"]["Engine!C5"]["reader"] == "read_shock_year_anchor"
    assert index["leaves"]["Engine!C5"]["call_form"] == "read_shock_year_anchor(ctx)"

    resolved = resolve_reader_ref("Engine!C5", index=index)
    assert resolved["mode"] == "reader"
    assert resolved["call_form"] == "read_shock_year_anchor(ctx)"


def test_discovery_lists_readers_but_not_setters() -> None:
    bindings = validate_bindings_document(_constant_series_doc())
    assert setter_names(bindings) == []
    assert reader_names(bindings) == ["read_shock_year_anchor"]
    assert not any(has_input_direction(s) for s in bindings["series"])


def test_merge_constant_direction_blocks_across_shards() -> None:
    base = _constant_series_doc()["series"][0]
    structural = {k: v for k, v in base.items() if k != "constant"}
    left_doc = {"schema_version": "1.11.0", "series": [{**structural, "constant": {}}]}
    right_doc = {"schema_version": "1.11.0", "series": [{**structural, "constant": {}}]}
    merged = merge_series_binding_documents([left_doc, right_doc])
    assert has_constant_direction(merged["series"][0])


def test_load_constant_shard_from_yaml(tmp_path: Path) -> None:
    shard = tmp_path / "constants.bindings.yaml"
    shard.write_text(
        """\
schema_version: "1.11.0"
workbook: constants.xlsx
series:
  - id: shock_year_anchor
    sheet: Engine
    data_range: Engine!C5
    layout: scalar
    constant: {}
    structure:
      measure:
        concept: OBS_VALUE
        dtype: float
        bind: { kind: data_cell, read: float }
      dimensions: []
    key: []
""",
        encoding="utf-8",
    )
    bindings = load_series_bindings(shard)
    assert has_constant_direction(bindings["series"][0])
    assert parse_bindings_file(shard)["schema_version"] == "1.11.0"
