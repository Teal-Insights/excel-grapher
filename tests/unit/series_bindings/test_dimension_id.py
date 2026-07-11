"""Tests for dimension id vs concept separation (schema 1.8.0).

A dimension may declare an `id` distinct from its `concept` (SDMX concept
identity), so two dimensions in one series can share a concept (e.g. a
`TIME_PERIOD` axis and a `REFERENCE_TIME_PERIOD` comparison period). Record
fields and `key` entries use the effective dimension id (`id` when declared,
else `concept`); dtype inheritance from `concept_scheme` keys on `concept`.
"""

from __future__ import annotations

from collections.abc import Callable
from datetime import datetime
from pathlib import Path
from typing import Any, cast

import pytest
import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.runtime.cache import EvalContext, coerce_inputs_dict
from excel_grapher.series_bindings import (
    expand_data_range,
    load_series_bindings,
    resolve_series_binding,
    validate_series_bindings,
)
from excel_grapher.series_bindings.docstrings import derive_doc_contract
from excel_grapher.series_bindings.normalize import effective_dimension_id
from excel_grapher.series_bindings.schema import (
    SeriesBindingsSchemaError,
    validate_bindings_document,
)
from excel_grapher.series_bindings.setter_codegen import (
    emit_input_coerce_helpers,
    emit_setter_function,
    emit_setter_helpers,
)
from excel_grapher.series_bindings.versions import SUPPORTED_SCHEMA_VERSIONS
from tests.paths import SERIES_BINDINGS_FIXTURES as FIXTURES


def _write_reference_period_workbook(path: Path) -> None:
    """Row of GDP by year plus a reference-year cell shared by every leaf."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    ws.write_number("B1", 2019)
    date_format = wb.add_format({"num_format": "yyyy-mm-dd"})
    ws.write_datetime("C1", datetime(2019, 1, 15), date_format)
    ws.write("A2", "Borvelia")
    ws.write("A5", "GDP")
    for offset, year in enumerate([2020, 2021, 2022]):
        ws.write_number(0, 5 + offset, year)
        ws.write_number(4, 5 + offset, float(offset + 1))
    wb.close()


def _reference_period_series(**overrides: Any) -> dict[str, Any]:
    series: dict[str, Any] = {
        "id": "gdp_vs_reference",
        "sheet": "Inputs",
        "data_range": "Inputs!F5:H5",
        "layout": "series",
        "input": {"setter": {"name": "set_gdp_vs_reference"}},
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
                },
                {
                    "id": "REFERENCE_TIME_PERIOD",
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "series",
                    "bind": {"kind": "cell", "address": "Inputs!B1"},
                },
            ],
        },
        "key": ["TIME_PERIOD", "REFERENCE_TIME_PERIOD"],
    }
    series.update(overrides)
    return series


def _reference_period_document(**series_overrides: Any) -> dict[str, Any]:
    return {
        "schema_version": "1.8.0",
        "concept_scheme": {
            "concepts": [{"id": "TIME_PERIOD", "name": "Time period", "dtype": "int"}]
        },
        "series": [_reference_period_series(**series_overrides)],
    }


def _reference_period_graph(tmp_path: Path) -> tuple[Path, Any]:
    wb_path = tmp_path / "reference_period.xlsx"
    _write_reference_period_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        expand_data_range("Inputs!F5:H5"),
        load_values=True,
    )
    return wb_path, graph


# --- effective_dimension_id helper ---


def test_effective_dimension_id_defaults_to_concept() -> None:
    assert effective_dimension_id({"concept": "TIME_PERIOD"}) == "TIME_PERIOD"


def test_effective_dimension_id_prefers_declared_id() -> None:
    dim = {"id": "REFERENCE_TIME_PERIOD", "concept": "TIME_PERIOD"}
    assert effective_dimension_id(dim) == "REFERENCE_TIME_PERIOD"


def test_effective_dimension_id_empty_component() -> None:
    assert effective_dimension_id({}) == ""


# --- schema and versions ---


def test_schema_version_1_8_0_supported() -> None:
    assert "1.8.0" in SUPPORTED_SCHEMA_VERSIONS


def test_schema_accepts_dimension_and_attribute_id() -> None:
    doc = _reference_period_document()
    doc["series"][0]["structure"]["attributes"] = [
        {
            "id": "SOURCE_UNIT",
            "concept": "UNIT_MEASURE",
            "role": "attribute",
            "value": "PC_GDP",
        }
    ]
    bindings = validate_bindings_document(doc)
    dims = bindings["series"][0]["structure"]["dimensions"]
    assert dims[1]["id"] == "REFERENCE_TIME_PERIOD"
    assert dims[1]["concept"] == "TIME_PERIOD"


def test_yaml_fixture_loads_and_resolves(tmp_path: Path) -> None:
    wb_path, graph = _reference_period_graph(tmp_path)
    bindings = load_series_bindings(FIXTURES / "reference_period_1_8_0.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(
        graph,
        wb_path,
        series,
        concept_scheme=bindings.get("concept_scheme"),
        direction="input",
    )
    assert resolved["ok"] is True, resolved["issues"]
    keys = {
        (leaf["key"]["TIME_PERIOD"], leaf["key"]["REFERENCE_TIME_PERIOD"])
        for leaf in resolved["leaves"]
    }
    assert keys == {(2020, 2019), (2021, 2019), (2022, 2019)}


# --- validation ---


def test_validate_accepts_key_by_dimension_id(tmp_path: Path) -> None:
    wb_path, graph = _reference_period_graph(tmp_path)
    bindings = validate_bindings_document(_reference_period_document())
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    codes = {issue["code"] for issue in report["issues"]}
    assert "key_not_in_dimensions" not in codes
    assert report["ok"] is True


def test_validate_rejects_duplicate_effective_dimension_ids(tmp_path: Path) -> None:
    wb_path, graph = _reference_period_graph(tmp_path)
    doc = _reference_period_document()
    del doc["series"][0]["structure"]["dimensions"][1]["id"]
    doc["series"][0]["key"] = ["TIME_PERIOD"]
    bindings = validate_bindings_document(doc)
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    duplicate = [i for i in report["issues"] if i["code"] == "duplicate_dimension_id"]
    assert duplicate, report["issues"]
    assert duplicate[0]["level"] == "error"
    assert "TIME_PERIOD" in duplicate[0]["message"]
    assert report["ok"] is False


def test_validate_key_must_use_declared_dimension_id(tmp_path: Path) -> None:
    """A key entry naming the concept fails when the dimension declares an id."""
    wb_path, graph = _reference_period_graph(tmp_path)
    doc = _reference_period_document()
    dims = doc["series"][0]["structure"]["dimensions"]
    dims[0]["id"] = "PROJECTION_TIME_PERIOD"
    doc["series"][0]["key"] = ["TIME_PERIOD", "REFERENCE_TIME_PERIOD"]
    bindings = validate_bindings_document(doc)
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    codes = {issue["code"] for issue in report["issues"]}
    assert "key_not_in_dimensions" in codes
    assert report["ok"] is False


# --- resolution ---


def test_resolve_dimensions_sharing_concept(tmp_path: Path) -> None:
    wb_path, graph = _reference_period_graph(tmp_path)
    doc = _reference_period_document()
    bindings = validate_bindings_document(doc)
    series = bindings["series"][0]

    resolved = resolve_series_binding(
        graph,
        wb_path,
        series,
        concept_scheme=bindings.get("concept_scheme"),
        direction="input",
    )

    assert resolved["ok"] is True, resolved["issues"]
    assert len(resolved["leaves"]) == 3
    by_period = {leaf["key"]["TIME_PERIOD"]: leaf for leaf in resolved["leaves"]}
    assert set(by_period) == {2020, 2021, 2022}
    leaf = by_period[2021]
    assert leaf["address"] == "Inputs!G5"
    assert leaf["coordinates"]["TIME_PERIOD"] == 2021
    # dtype inherited from concept_scheme via the shared TIME_PERIOD concept
    assert leaf["coordinates"]["REFERENCE_TIME_PERIOD"] == 2019
    assert leaf["key"] == {"TIME_PERIOD": 2021, "REFERENCE_TIME_PERIOD": 2019}
    assert leaf["record"]["TIME_PERIOD"] == 2021
    assert leaf["record"]["REFERENCE_TIME_PERIOD"] == 2019
    assert leaf["record"]["OBS_VALUE"] == 2.0


# --- setter codegen round trip ---


def _exec_setters(lines: list[str]) -> dict[str, object]:
    namespace: dict[str, object] = {
        "EvalContext": EvalContext,
        "coerce_inputs_dict": coerce_inputs_dict,
    }
    source_lines = emit_input_coerce_helpers() + emit_setter_helpers() + lines
    exec("\n".join(source_lines), namespace)
    return namespace


def test_emit_setter_round_trips_dimension_id_key(tmp_path: Path) -> None:
    wb_path, graph = _reference_period_graph(tmp_path)
    bindings = validate_bindings_document(_reference_period_document())
    series = bindings["series"][0]
    resolved = resolve_series_binding(
        graph,
        wb_path,
        series,
        concept_scheme=bindings.get("concept_scheme"),
        direction="input",
    )
    ns = _exec_setters(emit_setter_function(series, resolved))
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_gdp_vs_reference"],
    )

    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    setter(
        ctx,
        [{"TIME_PERIOD": 2021, "REFERENCE_TIME_PERIOD": 2019, "OBS_VALUE": 42.0}],
    )
    assert ctx.inputs["Inputs!G5"] == 42.0

    with pytest.raises(ValueError, match="unknown fields"):
        setter(
            ctx,
            [
                {
                    "TIME_PERIOD": 2021,
                    "REFERENCE_TIME_PERIOD": 2019,
                    "COMPARISON_AREA": "x",
                    "OBS_VALUE": 1.0,
                }
            ],
        )


# --- docstring contract ---


def test_doc_contract_fields_use_dimension_id(tmp_path: Path) -> None:
    wb_path, graph = _reference_period_graph(tmp_path)
    bindings = validate_bindings_document(_reference_period_document())
    series = bindings["series"][0]
    resolved = resolve_series_binding(
        graph,
        wb_path,
        series,
        concept_scheme=bindings.get("concept_scheme"),
        direction="input",
    )

    contract = derive_doc_contract(
        series,
        function_kind="setter",
        function_name="set_gdp_vs_reference",
        resolution=resolved,
        bindings=bindings,
    )

    assert "REFERENCE_TIME_PERIOD" in contract.fields
    assert contract.required_fields == ("TIME_PERIOD", "REFERENCE_TIME_PERIOD", "OBS_VALUE")
    reference = contract.fields["REFERENCE_TIME_PERIOD"]
    # dtype and human name resolve through the underlying TIME_PERIOD concept
    assert reference.dtype == "int"
    assert reference.concept_name == "Time period"


# --- per-dimension dtype (same concept, different storage type) ---


def _reference_date_series(**overrides: Any) -> dict[str, Any]:
    """Same-concept dimensions with differing storage: int axis vs datetime cell."""
    series = _reference_period_series(**overrides)
    series["structure"]["dimensions"][1] = {
        "id": "REFERENCE_TIME_PERIOD",
        "concept": "TIME_PERIOD",
        "dtype": "datetime",
        "role": "key",
        "scope": "series",
        "bind": {"kind": "cell", "address": "Inputs!C1"},
    }
    return series


def test_schema_accepts_dimension_dtype() -> None:
    doc = _reference_period_document()
    doc["series"][0] = _reference_date_series()
    bindings = validate_bindings_document(doc)
    assert bindings["series"][0]["structure"]["dimensions"][1]["dtype"] == "datetime"


def test_schema_rejects_attribute_dtype() -> None:
    doc = _reference_period_document()
    doc["series"][0]["structure"]["attributes"] = [
        {
            "concept": "UNIT_MEASURE",
            "role": "attribute",
            "value": "PC_GDP",
            "dtype": "string",
        }
    ]
    with pytest.raises(SeriesBindingsSchemaError, match="dtype"):
        validate_bindings_document(doc)


def test_resolve_dimension_dtype_overrides_concept_dtype(tmp_path: Path) -> None:
    wb_path, graph = _reference_period_graph(tmp_path)
    doc = _reference_period_document()
    doc["series"][0] = _reference_date_series()
    bindings = validate_bindings_document(doc)
    series = bindings["series"][0]

    resolved = resolve_series_binding(
        graph,
        wb_path,
        series,
        concept_scheme=bindings.get("concept_scheme"),
        direction="input",
    )

    assert resolved["ok"] is True, resolved["issues"]
    leaf = resolved["leaves"][0]
    # TIME_PERIOD still inherits int from the concept scheme; the reference
    # dimension's declared dtype wins over the shared concept's int.
    assert leaf["key"]["TIME_PERIOD"] == 2020
    assert leaf["key"]["REFERENCE_TIME_PERIOD"] == datetime(2019, 1, 15)


def test_resolve_series_context_uses_dimension_dtype(tmp_path: Path) -> None:
    wb_path, graph = _reference_period_graph(tmp_path)
    doc = _reference_period_document()
    doc["series"][0] = _reference_date_series()
    dim = doc["series"][0]["structure"]["dimensions"][1]
    dim["include_in_record"] = False
    doc["series"][0]["key"] = ["TIME_PERIOD"]
    doc["series"][0]["series_context"] = {"REFERENCE_TIME_PERIOD": "2019-01-15"}
    bindings = validate_bindings_document(doc)
    series = bindings["series"][0]

    resolved = resolve_series_binding(
        graph,
        wb_path,
        series,
        concept_scheme=bindings.get("concept_scheme"),
        direction="input",
    )

    assert resolved["ok"] is True, resolved["issues"]
    record = resolved["leaves"][0]["record"]
    assert record["REFERENCE_TIME_PERIOD"] == datetime(2019, 1, 15)


def test_emit_setter_key_dtypes_include_dimension_dtype(tmp_path: Path) -> None:
    wb_path, graph = _reference_period_graph(tmp_path)
    doc = _reference_period_document()
    doc["series"][0] = _reference_date_series()
    bindings = validate_bindings_document(doc)
    series = bindings["series"][0]
    resolved = resolve_series_binding(
        graph,
        wb_path,
        series,
        concept_scheme=bindings.get("concept_scheme"),
        direction="input",
    )
    lines = emit_setter_function(series, resolved)
    code = "\n".join(lines)
    assert "'REFERENCE_TIME_PERIOD': 'datetime'" in code

    ns = _exec_setters(lines)
    setter = cast(
        Callable[[EvalContext, list[dict[str, object]]], None],
        ns["set_gdp_vs_reference"],
    )
    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    setter(
        ctx,
        [
            {
                "TIME_PERIOD": 2021,
                "REFERENCE_TIME_PERIOD": datetime(2019, 1, 15),
                "OBS_VALUE": 42.0,
            }
        ],
    )
    assert ctx.inputs["Inputs!G5"] == 42.0


def test_doc_contract_reports_dimension_dtype(tmp_path: Path) -> None:
    wb_path, graph = _reference_period_graph(tmp_path)
    doc = _reference_period_document()
    doc["series"][0] = _reference_date_series()
    bindings = validate_bindings_document(doc)
    series = bindings["series"][0]
    resolved = resolve_series_binding(
        graph,
        wb_path,
        series,
        concept_scheme=bindings.get("concept_scheme"),
        direction="input",
    )

    contract = derive_doc_contract(
        series,
        function_kind="setter",
        function_name="set_gdp_vs_reference",
        resolution=resolved,
        bindings=bindings,
    )

    # declared per-dimension dtype wins over the shared concept's int dtype
    assert contract.fields["REFERENCE_TIME_PERIOD"].dtype == "datetime"
    assert contract.fields["TIME_PERIOD"].dtype == "int"


def test_validate_warns_on_dimension_dtype_read_mismatch(tmp_path: Path) -> None:
    wb_path, graph = _reference_period_graph(tmp_path)
    doc = _reference_period_document()
    doc["series"][0] = _reference_date_series()
    dim = doc["series"][0]["structure"]["dimensions"][1]
    dim["bind"]["read"] = "int"
    bindings = validate_bindings_document(doc)
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    mismatches = [i for i in report["issues"] if i["code"] == "dtype_read_mismatch"]
    assert mismatches
    assert "REFERENCE_TIME_PERIOD" in mismatches[0]["message"]


def test_validate_read_diverging_from_concept_suggests_dimension_dtype(tmp_path: Path) -> None:
    """bind.read conflicting with inherited concept dtype nudges toward dtype or split."""
    wb_path, graph = _reference_period_graph(tmp_path)
    doc = _reference_period_document()
    dim = doc["series"][0]["structure"]["dimensions"][1]
    dim["bind"]["read"] = "datetime"
    bindings = validate_bindings_document(doc)
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    mismatches = [i for i in report["issues"] if i["code"] == "dtype_read_mismatch"]
    assert mismatches
    assert "per-dimension dtype" in mismatches[0]["message"]


def test_validate_dimension_dtype_without_read_is_clean(tmp_path: Path) -> None:
    wb_path, graph = _reference_period_graph(tmp_path)
    doc = _reference_period_document()
    doc["series"][0] = _reference_date_series()
    bindings = validate_bindings_document(doc)
    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    codes = {issue["code"] for issue in report["issues"]}
    assert "dtype_read_mismatch" not in codes
    assert report["ok"] is True
