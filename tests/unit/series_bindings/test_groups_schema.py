"""Schema tests for series binding API groups (schema 1.6.0)."""

from __future__ import annotations

from typing import Any

import pytest

from excel_grapher.series_bindings.schema import (
    SeriesBindingsSchemaError,
    validate_bindings_document,
)
from excel_grapher.series_bindings.versions import SUPPORTED_SCHEMA_VERSIONS


def _scalar_series_doc(**series_overrides: Any) -> dict[str, Any]:
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
    return {"schema_version": "1.6.0", "series": [series]}


def test_schema_1_6_0_supported() -> None:
    assert "1.6.0" in SUPPORTED_SCHEMA_VERSIONS


def test_schema_accepts_groups_with_path_and_order() -> None:
    doc = _scalar_series_doc(
        groups=[{"path": ["Climate scenarios", "Paris"], "order": 13}],
    )
    bindings = validate_bindings_document(doc)
    assert bindings["series"][0]["groups"] == [
        {"path": ["Climate scenarios", "Paris"], "order": 13},
    ]


def test_schema_accepts_groups_without_order() -> None:
    doc = _scalar_series_doc(groups=[{"path": ["Baseline setup"]}])
    bindings = validate_bindings_document(doc)
    assert bindings["series"][0]["groups"] == [{"path": ["Baseline setup"]}]


def test_schema_rejects_empty_group_path() -> None:
    doc = _scalar_series_doc(groups=[{"path": []}])
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_rejects_group_without_path() -> None:
    doc = _scalar_series_doc(groups=[{"order": 1}])
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_1_5_0_ignores_groups_when_omitted() -> None:
    doc = _scalar_series_doc()
    doc["schema_version"] = "1.5.0"
    bindings = validate_bindings_document(doc)
    assert "groups" not in bindings["series"][0]
