"""Unit tests for series binding JSON Schema validation."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.series_bindings import load_series_bindings
from excel_grapher.series_bindings.schema import (
    SeriesBindingsSchemaError,
    format_schema_errors,
    validate_bindings_document,
)

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "series_bindings"


def _scalar_series_doc(**series_overrides: Any) -> dict[str, Any]:
    """Minimal scalar binding document for schema edge-case tests."""
    schema_version = "1.0.0"
    if "schema_version" in series_overrides:
        schema_version = str(series_overrides.pop("schema_version"))
    series: dict[str, Any] = {
        "id": "bool_flag",
        "sheet": "Flags",
        "data_range": "Flags!B2",
        "layout": "scalar",
        "setter": {"name": "set_bool_flag"},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "concept": "IS_ACTIVE",
                    "role": "key",
                    "scope": "series",
                    "bind": {"kind": "constant", "value": True},
                }
            ],
        },
        "key": ["IS_ACTIVE"],
    }
    series.update(series_overrides)
    return {"schema_version": schema_version, "series": [series]}


def test_example_fixture_passes_schema() -> None:
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    assert bindings["series"][0]["key"] == ["TIME_PERIOD"]


def test_schema_infers_sheet_from_sheet_qualified_data_range() -> None:
    doc = {
        "schema_version": "1.3.0",
        "series": [
            {
                "id": "borvelia_primary_balance",
                "data_range": "Sheet1!F5:J5",
                "layout": "series",
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


def test_schema_accepts_bare_defined_name_data_range() -> None:
    doc = {
        "schema_version": "1.0.0",
        "series": [
            {
                "id": "defined_name_target",
                "sheet": "Inputs",
                "data_range": "growth_baseline",
                "layout": "scalar",
                "setter": {"name": "set_defined_name_target"},
                "structure": {
                    "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
                    "dimensions": [],
                },
                "key": [],
            }
        ],
    }

    bindings = validate_bindings_document(doc)
    assert bindings["series"][0]["data_range"] == "growth_baseline"


def test_schema_accepts_1_1_0_matrix_fixture() -> None:
    bindings = load_series_bindings(FIXTURES / "matrix_country_block_1_1_0.yaml")
    assert bindings["schema_version"] == "1.1.0"
    assert bindings["series"][0]["layout"] == "matrix"
    assert bindings["series"][0]["structure"]["dimensions"][0]["bind"]["kind"] == "row_hierarchy"


def test_schema_accepts_legacy_row_series_layout() -> None:
    doc = {
        "schema_version": "1.0.0",
        "series": [
            {
                "id": "legacy_row",
                "sheet": "S",
                "data_range": "S!B2:C2",
                "layout": "row_series",
                "setter": {"name": "set_legacy_row"},
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
    assert bindings["series"][0]["layout"] == "series"


def test_series_rejects_empty_key() -> None:
    doc = {
        "schema_version": "1.3.0",
        "series": [
            {
                "id": "empty_key_row",
                "sheet": "S",
                "data_range": "S!B2:C2",
                "layout": "series",
                "setter": {"name": "set_empty_key_row"},
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
                "key": [],
            }
        ],
    }
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_series_requires_cell_scoped_dimension() -> None:
    doc = {
        "schema_version": "1.3.0",
        "series": [
            {
                "id": "only_series_scope",
                "sheet": "S",
                "data_range": "S!B2:C2",
                "layout": "series",
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


def test_schema_accepts_read_bool_on_data_cell_bind() -> None:
    doc = _scalar_series_doc(
        structure={
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "bool",
                "bind": {"kind": "data_cell", "read": "bool"},
            },
            "dimensions": [],
        },
        key=[],
    )

    bindings = validate_bindings_document(doc)
    measure = bindings["series"][0]["structure"]["measure"]
    assert measure["dtype"] == "bool"
    assert measure["bind"]["read"] == "bool"


def test_schema_accepts_read_bool_on_column_header_bind() -> None:
    doc = {
        "schema_version": "1.0.0",
        "series": [
            {
                "id": "bool_columns",
                "sheet": "Flags",
                "data_range": "Flags!B2:C2",
                "layout": "series",
                "setter": {"name": "set_bool_columns"},
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": "float",
                        "bind": {"kind": "data_cell", "read": "float"},
                    },
                    "dimensions": [
                        {
                            "concept": "IS_ENABLED",
                            "role": "key",
                            "scope": "cell",
                            "bind": {
                                "kind": "column_header",
                                "header_row": 1,
                                "read": "bool",
                            },
                        }
                    ],
                },
                "key": ["IS_ENABLED"],
            }
        ],
    }

    bindings = validate_bindings_document(doc)
    bind = bindings["series"][0]["structure"]["dimensions"][0]["bind"]
    assert bind["read"] == "bool"


def test_schema_accepts_measure_dtype_bool() -> None:
    doc = _scalar_series_doc(
        structure={
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "bool",
                "bind": {"kind": "data_cell", "read": "bool"},
            },
            "dimensions": [],
        },
        key=[],
    )

    bindings = validate_bindings_document(doc)
    assert bindings["series"][0]["structure"]["measure"]["dtype"] == "bool"


def test_schema_accepts_boolean_constant_bind_value() -> None:
    doc = _scalar_series_doc(
        structure={
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "concept": "IS_ACTIVE",
                    "role": "key",
                    "scope": "series",
                    "bind": {"kind": "constant", "value": False},
                }
            ],
        },
        key=["IS_ACTIVE"],
    )

    bindings = validate_bindings_document(doc)
    bind = bindings["series"][0]["structure"]["dimensions"][0]["bind"]
    assert bind["value"] is False


def test_schema_accepts_boolean_series_context_value() -> None:
    doc = _scalar_series_doc(series_context={"IS_ACTIVE": True})

    bindings = validate_bindings_document(doc)
    assert bindings["series"][0]["series_context"]["IS_ACTIVE"] is True


def test_schema_accepts_boolean_attribute_value() -> None:
    doc = _scalar_series_doc(
        structure={
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [],
            "attributes": [
                {
                    "concept": "IS_ESTIMATE",
                    "role": "attribute",
                    "value": True,
                    "include_in_record": True,
                }
            ],
        },
        key=[],
    )

    bindings = validate_bindings_document(doc)
    attribute = bindings["series"][0]["structure"]["attributes"][0]
    assert attribute["value"] is True


def test_schema_accepts_concept_scheme_bool_dtype() -> None:
    doc = _scalar_series_doc()
    doc["concept_scheme"] = {
        "id": "flags",
        "concepts": [
            {
                "id": "IS_ACTIVE",
                "dtype": "bool",
                "description": "Whether the series row is active.",
            }
        ],
    }

    bindings = validate_bindings_document(doc)
    concept = bindings["concept_scheme"]["concepts"][0]
    assert concept["dtype"] == "bool"


@pytest.mark.parametrize(
    ("field_path", "invalid_value"),
    [
        ("measure.bind.read", "boolean"),
        ("measure.dtype", "boolean"),
        ("dimensions.0.bind.read", "boolean"),
    ],
)
def test_schema_rejects_non_enum_bool_tokens(
    field_path: str,
    invalid_value: str,
) -> None:
    doc = _scalar_series_doc(
        structure={
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "bool",
                "bind": {"kind": "data_cell", "read": "bool"},
            },
            "dimensions": [
                {
                    "concept": "IS_ACTIVE",
                    "role": "key",
                    "scope": "cell",
                    "bind": {
                        "kind": "column_header",
                        "header_row": 1,
                        "read": "bool",
                    },
                }
            ],
        },
        key=["IS_ACTIVE"],
    )
    series = doc["series"][0]
    structure = series["structure"]
    if field_path == "measure.bind.read":
        structure["measure"]["bind"]["read"] = invalid_value
    elif field_path == "measure.dtype":
        structure["measure"]["dtype"] = invalid_value
    elif field_path == "dimensions.0.bind.read":
        structure["dimensions"][0]["bind"]["read"] = invalid_value

    errors = format_schema_errors(doc)
    assert errors
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_accepts_read_datetime_on_column_header_bind() -> None:
    doc = {
        "schema_version": "1.4.0",
        "series": [
            {
                "id": "calendar_periods",
                "sheet": "Inputs",
                "data_range": "Inputs!B2:C2",
                "layout": "row_series",
                "setter": {"name": "set_calendar_periods"},
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
                                "header_row": 1,
                                "read": "datetime",
                            },
                        }
                    ],
                },
                "key": ["TIME_PERIOD"],
            }
        ],
    }

    bindings = validate_bindings_document(doc)
    assert bindings["schema_version"] == "1.4.0"
    bind = bindings["series"][0]["structure"]["dimensions"][0]["bind"]
    assert bind["read"] == "datetime"


def test_schema_accepts_measure_dtype_datetime() -> None:
    doc = _scalar_series_doc(
        schema_version="1.4.0",
        structure={
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "datetime",
                "bind": {"kind": "data_cell", "read": "datetime"},
            },
            "dimensions": [],
        },
        key=[],
    )

    bindings = validate_bindings_document(doc)
    assert bindings["series"][0]["structure"]["measure"]["dtype"] == "datetime"


def test_schema_accepts_iso_datetime_constant_bind_value() -> None:
    doc = _scalar_series_doc(
        schema_version="1.4.0",
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
                    "scope": "series",
                    "bind": {"kind": "constant", "value": "2024-01-15"},
                }
            ],
        },
        key=["TIME_PERIOD"],
    )

    bindings = validate_bindings_document(doc)
    bind = bindings["series"][0]["structure"]["dimensions"][0]["bind"]
    assert bind["value"] == "2024-01-15"


def test_schema_accepts_iso_datetime_series_context_value() -> None:
    doc = _scalar_series_doc(
        schema_version="1.4.0",
        series_context={"TIME_PERIOD": "2024-01-15T00:00:00"},
    )

    bindings = validate_bindings_document(doc)
    assert bindings["series"][0]["series_context"]["TIME_PERIOD"] == "2024-01-15T00:00:00"


def test_schema_accepts_concept_scheme_datetime_dtype() -> None:
    doc = _scalar_series_doc(schema_version="1.4.0")
    doc["concept_scheme"] = {
        "id": "calendar",
        "concepts": [
            {
                "id": "TIME_PERIOD",
                "dtype": "datetime",
                "description": "Observation reference period.",
            }
        ],
    }

    bindings = validate_bindings_document(doc)
    concept = bindings["concept_scheme"]["concepts"][0]
    assert concept["dtype"] == "datetime"


@pytest.mark.parametrize(
    ("field_path", "invalid_value"),
    [
        ("measure.bind.read", "date"),
        ("measure.dtype", "date"),
        ("dimensions.0.bind.read", "DateTime"),
    ],
)
def test_schema_rejects_non_enum_datetime_tokens(
    field_path: str,
    invalid_value: str,
) -> None:
    doc = _scalar_series_doc(
        schema_version="1.4.0",
        structure={
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "datetime",
                "bind": {"kind": "data_cell", "read": "datetime"},
            },
            "dimensions": [
                {
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": {
                        "kind": "column_header",
                        "header_row": 1,
                        "read": "datetime",
                    },
                }
            ],
        },
        key=["TIME_PERIOD"],
    )
    series = doc["series"][0]
    structure = series["structure"]
    if field_path == "measure.bind.read":
        structure["measure"]["bind"]["read"] = invalid_value
    elif field_path == "measure.dtype":
        structure["measure"]["dtype"] = invalid_value
    elif field_path == "dimensions.0.bind.read":
        structure["dimensions"][0]["bind"]["read"] = invalid_value

    errors = format_schema_errors(doc)
    assert errors
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)
