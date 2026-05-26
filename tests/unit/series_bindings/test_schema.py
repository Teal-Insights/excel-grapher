"""Unit tests for series binding JSON Schema validation."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.series_bindings import load_series_bindings
from excel_grapher.series_bindings.schema import (
    SeriesBindingsSchemaError,
    format_schema_errors,
    validate_bindings_document,
)

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "series_bindings"


def test_example_fixture_passes_schema() -> None:
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    assert bindings["series"][0]["key"] == ["TIME_PERIOD"]


def test_schema_infers_sheet_from_sheet_qualified_data_range() -> None:
    doc = {
        "schema_version": "1.0.0",
        "series": [
            {
                "id": "borvelia_primary_balance",
                "data_range": "Sheet1!F5:J5",
                "layout": "row_series",
                "setter": {"name": "set_borvelia_primary_balance"},
                "structure": {
                    "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
                    "dimensions": [
                        {
                            "concept": "TIME_PERIOD",
                            "role": "key",
                            "scope": "cell",
                            "bind": {"kind": "column_header", "header_row": 1},
                        }
                    ],
                },
                "key": ["TIME_PERIOD"],
            }
        ],
    }
    bindings = validate_bindings_document(doc)
    assert bindings["series"][0]["sheet"] == "Sheet1"


def test_schema_rejects_missing_series() -> None:
    with pytest.raises(SeriesBindingsSchemaError, match="series"):
        validate_bindings_document({"schema_version": "1.0.0"})


def test_schema_rejects_bad_setter_name() -> None:
    doc = {
        "schema_version": "1.0.0",
        "series": [
            {
                "id": "bad",
                "sheet": "S",
                "data_range": "S!A1",
                "layout": "scalar",
                "setter": {"name": "not_a_setter"},
                "structure": {
                    "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
                    "dimensions": [
                        {
                            "concept": "X",
                            "role": "key",
                            "scope": "cell",
                            "bind": {"kind": "constant", "value": 1},
                        }
                    ],
                },
                "key": ["X"],
            }
        ],
    }
    errors = format_schema_errors(doc)
    assert any("setter" in e for e in errors)
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_accepts_1_1_0_matrix_fixture() -> None:
    bindings = load_series_bindings(FIXTURES / "matrix_country_block_1_1_0.yaml")
    assert bindings["schema_version"] == "1.1.0"
    assert bindings["series"][0]["layout"] == "matrix"
    assert bindings["series"][0]["structure"]["dimensions"][0]["bind"]["kind"] == "row_hierarchy"


def test_row_series_requires_cell_scoped_dimension() -> None:
    doc = {
        "schema_version": "1.0.0",
        "series": [
            {
                "id": "only_series_scope",
                "sheet": "S",
                "data_range": "S!B2:C2",
                "layout": "row_series",
                "setter": {"name": "set_only_series_scope"},
                "structure": {
                    "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
                    "dimensions": [
                        {
                            "concept": "COUNTRY",
                            "role": "key",
                            "scope": "series",
                            "bind": {"kind": "constant", "value": "X"},
                        }
                    ],
                },
                "key": ["COUNTRY"],
            }
        ],
    }
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)
