"""Unit tests for series binding normalization and merge helpers."""

from __future__ import annotations

import pytest

from excel_grapher.series_bindings import (
    SeriesBindingsLoadError,
    has_input_direction,
    has_output_direction,
    merge_series_binding_documents,
    normalize_series_entry,
    parse_bindings_file,
)
from tests.paths import SERIES_BINDINGS_FIXTURES as FIXTURES


def test_normalize_renames_legacy_row_series_layout() -> None:
    series = {
        "id": "x",
        "layout": "row_series",
        "input": {},
    }
    normalized = normalize_series_entry(series)
    assert normalized["layout"] == "series"


def test_empty_input_block_is_input_direction() -> None:
    series = {
        "id": "x",
        "input": {},
    }
    normalized = normalize_series_entry(series)
    assert "setter" not in normalized
    assert normalized["input"] == {}
    assert has_input_direction(normalized)
    assert not has_output_direction(normalized)


def test_merge_input_and_output_shards() -> None:
    input_doc = parse_bindings_file(FIXTURES / "shard_borvelia_input.yaml")
    output_doc = parse_bindings_file(FIXTURES / "shard_borvelia_output.yaml")
    merged = merge_series_binding_documents([input_doc, output_doc])
    series = merged["series"][0]
    assert has_input_direction(series)
    assert has_output_direction(series)
    assert series["input"] == {}
    assert series["output"]["compute"]["name"] == "compute_borvelia_primary_balance"


def test_schema_requires_at_least_one_direction() -> None:
    from excel_grapher.series_bindings import SeriesBindingsSchemaError, validate_bindings_document

    doc = {
        "schema_version": "1.3.0",
        "series": [
            {
                "id": "no_direction",
                "sheet": "S",
                "data_range": "S!A1",
                "layout": "scalar",
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
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_schema_rejects_non_mapping_series_entry() -> None:
    from excel_grapher.series_bindings import SeriesBindingsSchemaError, validate_bindings_document

    doc = {
        "schema_version": "1.3.0",
        "series": [
            {
                "id": "ok",
                "sheet": "S",
                "data_range": "S!A1",
                "layout": "scalar",
                "input": {},
                "structure": {
                    "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
                    "dimensions": [
                        {
                            "concept": "X",
                            "role": "key",
                            "scope": "cell",
                            "bind": {"kind": "data_cell"},
                        }
                    ],
                },
                "key": ["X"],
            },
            123,
        ],
    }
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(doc)


def test_merge_rejects_conflicting_output_blocks() -> None:
    merged_base = merge_series_binding_documents(
        [
            parse_bindings_file(FIXTURES / "shard_borvelia_input.yaml"),
            parse_bindings_file(FIXTURES / "shard_borvelia_output.yaml"),
        ]
    )
    conflicting = parse_bindings_file(FIXTURES / "shard_borvelia_output.yaml")
    conflicting["series"][0]["output"]["compute"]["name"] = "compute_other_name"
    with pytest.raises(SeriesBindingsLoadError, match="conflicting output"):
        merge_series_binding_documents([merged_base, conflicting])
