"""Schema, coercion, and validation tests for `input.value_map` (schema 1.15.0)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.series_bindings.input_coerce import (
    apply_input_value_map,
    coerce_setter_input,
    input_value_map_from_series,
    measure_domain_from_series,
)
from excel_grapher.series_bindings.schema import (
    SeriesBindingsSchemaError,
    validate_bindings_document,
)
from excel_grapher.series_bindings.versions import SUPPORTED_SCHEMA_VERSIONS

_SELECTOR_MAP = {"High": "High ", "Medium": "Medium", "Low": "Low "}


def _scalar_string_doc(
    *,
    value_map: dict[str, Any] | None = None,
    domain: dict[str, Any] | None = None,
    layout: str = "scalar",
    extra_series: dict[str, Any] | None = None,
) -> dict[str, Any]:
    input_block: dict[str, Any] = {"setter": {"name": "set_selector"}}
    if domain is not None:
        input_block["domain"] = domain
    if value_map is not None:
        input_block["value_map"] = dict(value_map)
    series: dict[str, Any] = {
        "id": "selector",
        "sheet": "Inputs",
        "data_range": "Inputs!A1",
        "layout": layout,
        "input": input_block,
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
    if extra_series:
        series.update(extra_series)
    return {"schema_version": "1.15.0", "series": [series]}


def _series_layout_doc(*, value_map: dict[str, Any] | None = None) -> dict[str, Any]:
    input_block: dict[str, Any] = {"setter": {"name": "set_rates"}}
    if value_map is not None:
        input_block["value_map"] = dict(value_map)
    return {
        "schema_version": "1.15.0",
        "series": [
            {
                "id": "rates",
                "sheet": "Inputs",
                "data_range": "Inputs!B1:C1",
                "layout": "series",
                "input": input_block,
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
                                "header_row": 10,
                                "read": "int",
                            },
                        }
                    ],
                },
                "key": ["TIME_PERIOD"],
            }
        ],
    }


def _output_only_doc(*, extra: dict[str, Any] | None = None) -> dict[str, Any]:
    series: dict[str, Any] = {
        "id": "out",
        "sheet": "Outputs",
        "data_range": "Outputs!A1",
        "layout": "scalar",
        "output": {"compute": {"name": "compute_out"}},
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
    if extra:
        series.update(extra)
    return {"schema_version": "1.15.0", "series": [series]}


def test_schema_version_1_15_0_supported() -> None:
    assert "1.15.0" in SUPPORTED_SCHEMA_VERSIONS


def test_schema_accepts_value_map_on_scalar_input() -> None:
    bindings = validate_bindings_document(_scalar_string_doc(value_map=_SELECTOR_MAP))
    assert bindings["series"][0]["input"]["value_map"] == _SELECTOR_MAP


def test_schema_rejects_empty_value_map() -> None:
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(_scalar_string_doc(value_map={}))


def test_schema_rejects_value_map_on_non_scalar_input() -> None:
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(_series_layout_doc(value_map=_SELECTOR_MAP))


def test_schema_rejects_value_map_on_non_input_series() -> None:
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(_output_only_doc(extra={"value_map": _SELECTOR_MAP}))


def test_input_value_map_from_series_and_implied_domain() -> None:
    series = validate_bindings_document(_scalar_string_doc(value_map=_SELECTOR_MAP))["series"][0]
    assert input_value_map_from_series(series) == _SELECTOR_MAP
    domain = measure_domain_from_series(series)
    assert domain == {"enum": frozenset(_SELECTOR_MAP)}


def test_explicit_domain_is_not_replaced_by_value_map_keys() -> None:
    series = validate_bindings_document(
        _scalar_string_doc(
            value_map=_SELECTOR_MAP,
            domain={"enum": ["High", "Low"]},
        )
    )["series"][0]
    assert measure_domain_from_series(series) == {"enum": frozenset({"High", "Low"})}


def test_apply_input_value_map_rewrites_and_rejects() -> None:
    assert apply_input_value_map("High", _SELECTOR_MAP, series_id="selector") == "High "
    assert apply_input_value_map("Medium", _SELECTOR_MAP, series_id="selector") == "Medium"
    with pytest.raises(ValueError, match=r"selector value 'Nope' is not in value_map"):
        apply_input_value_map("Nope", _SELECTOR_MAP, series_id="selector")


def test_coerce_setter_input_maps_after_domain() -> None:
    domain = {"enum": frozenset(_SELECTOR_MAP)}
    assert coerce_setter_input(
        "High",
        layout="scalar",
        key_fields=(),
        measure_field="OBS_VALUE",
        key_order=None,
        strict=True,
        measure_dtype="string",
        measure_domain=domain,
        value_map=_SELECTOR_MAP,
    ) == [{"OBS_VALUE": "High "}]
    with pytest.raises(ValueError, match="out of domain"):
        coerce_setter_input(
            "Nope",
            layout="scalar",
            key_fields=(),
            measure_field="OBS_VALUE",
            key_order=None,
            strict=True,
            measure_dtype="string",
            measure_domain=domain,
            value_map=_SELECTOR_MAP,
        )


def test_validate_rejects_domain_enum_outside_value_map(tmp_path: Path) -> None:
    from excel_grapher.grapher import create_dependency_graph
    from excel_grapher.series_bindings import validate_series_bindings
    from tests.unit.exporter.inverted_tree.helpers import write_workbook

    workbook = write_workbook(tmp_path / "selector.xlsx", {"Inputs": {"A1": "High "}})
    doc = validate_bindings_document(
        _scalar_string_doc(
            value_map=_SELECTOR_MAP,
            domain={"enum": ["High", "Nope"]},
        )
    )
    graph = create_dependency_graph(workbook, ["Inputs!A1"], load_values=True)
    report = validate_series_bindings(graph, doc, workbook=workbook)
    messages = [
        issue["message"] for issue in report["issues"] if issue["code"] == "invalid_input_value_map"
    ]
    assert messages
    assert "Nope" in messages[0]
