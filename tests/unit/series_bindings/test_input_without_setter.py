"""Schema 1.16.0: `input: {}` is an input; leftover setter/reader blocks are stripped."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.series_bindings import (
    has_input_direction,
    normalize_series_entry,
    validate_bindings_document,
)
from excel_grapher.series_bindings.schema import SeriesBindingsSchemaError
from excel_grapher.series_bindings.versions import SUPPORTED_SCHEMA_VERSIONS


def _scalar_series(**overrides: Any) -> dict[str, Any]:
    series: dict[str, Any] = {
        "id": "interest_rate",
        "sheet": "Inputs",
        "data_range": "Inputs!B4",
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
    series.update(overrides)
    return series


def test_schema_1_16_0_is_supported() -> None:
    assert "1.16.0" in SUPPORTED_SCHEMA_VERSIONS


def test_empty_input_block_validates() -> None:
    bindings = validate_bindings_document(
        {"schema_version": "1.16.0", "series": [_scalar_series(input={})]}
    )
    series = bindings["series"][0]
    assert series["input"] == {}
    assert has_input_direction(series)
    assert "setter" not in series
    assert "setter" not in series["input"]


def test_input_domain_without_setter_validates() -> None:
    bindings = validate_bindings_document(
        {
            "schema_version": "1.16.0",
            "series": [
                _scalar_series(
                    input={"domain": {"real_between": {"min": 0, "max": 300}}},
                )
            ],
        }
    )
    assert bindings["series"][0]["input"] == {
        "domain": {"real_between": {"min": 0, "max": 300}}
    }
    assert has_input_direction(bindings["series"][0])


def test_legacy_top_level_setter_strips_to_empty_input() -> None:
    normalized = normalize_series_entry(
        _scalar_series(setter={"name": "set_interest_rate", "strict": True})
    )
    assert "setter" not in normalized
    assert normalized["input"] == {}
    assert has_input_direction(normalized)

    bindings = validate_bindings_document(
        {
            "schema_version": "1.0.0",
            "series": [_scalar_series(setter={"name": "set_interest_rate"})],
        }
    )
    assert bindings["series"][0]["input"] == {}
    assert has_input_direction(bindings["series"][0])


def test_input_setter_and_reader_are_stripped() -> None:
    bindings = validate_bindings_document(
        {
            "schema_version": "1.13.0",
            "series": [
                _scalar_series(
                    input={
                        "setter": {
                            "name": "set_interest_rate",
                            "record_contract": "records",
                            "strict": True,
                        },
                        "reader": {"name": "read_interest_rate"},
                        "domain": {"real_between": {"min": 0, "max": 1}},
                    }
                )
            ],
        }
    )
    assert bindings["series"][0]["input"] == {
        "domain": {"real_between": {"min": 0, "max": 1}}
    }


def test_invalid_setter_name_is_stripped_not_rejected() -> None:
    bindings = validate_bindings_document(
        {
            "schema_version": "1.0.0",
            "series": [_scalar_series(setter={"name": "not_a_setter"})],
        }
    )
    assert has_input_direction(bindings["series"][0])
    assert bindings["series"][0]["input"] == {}


def test_series_without_direction_still_rejected() -> None:
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(
            {"schema_version": "1.16.0", "series": [_scalar_series()]}
        )


def test_empty_input_exports_without_setters(tmp_path: Path) -> None:
    from excel_grapher.exporter import CodeGenerator
    from excel_grapher.grapher import create_dependency_graph
    from tests.unit.exporter.inverted_tree.helpers import write_workbook

    workbook = tmp_path / "rate.xlsx"
    write_workbook(
        workbook,
        {
            "Inputs": {"B4": 0.05},
            "Outputs": {"A1": "=Inputs!B4"},
        },
    )
    bindings = validate_bindings_document(
        {
            "schema_version": "1.16.0",
            "series": [
                _scalar_series(input={}),
                {
                    "id": "result",
                    "sheet": "Outputs",
                    "data_range": "Outputs!A1",
                    "layout": "scalar",
                    "output": {"compute": {"name": "compute_result"}},
                    "structure": {
                        "measure": {
                            "concept": "OBS_VALUE",
                            "dtype": "float",
                            "bind": {"kind": "data_cell", "read": "float"},
                        },
                        "dimensions": [],
                    },
                    "key": [],
                },
            ],
        }
    )
    graph = create_dependency_graph(workbook, ["Outputs!A1"], load_values=True)
    with CodeGenerator(graph) as gen:
        modules = gen.generate_modules(
            series_bindings=bindings,
            bindings_workbook=workbook,
        )
    api = modules["api.py"]
    assert "def set_" not in api
    assert "def compute_result(" in api
    assert "interest_rate" in api


def test_public_api_does_not_export_setter_emitters() -> None:
    import excel_grapher.series_bindings as series_bindings

    assert not hasattr(series_bindings, "emit_setter_function")
    assert not hasattr(series_bindings, "emit_setters_block")
    assert not hasattr(series_bindings, "generate_setters_module")
    assert not hasattr(series_bindings, "emit_series_bindings_block")
    assert "emit_setter_function" not in series_bindings.__all__
