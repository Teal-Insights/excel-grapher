"""Unit tests for series binding normalization and merge helpers."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.series_bindings import (
    SeriesBindingsLoadError,
    has_input_direction,
    has_output_direction,
    load_series_bindings,
    merge_series_binding_documents,
    normalize_series_entry,
    parse_bindings_file,
)

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "series_bindings"


def test_normalize_moves_legacy_setter_to_input() -> None:
    series = {
        "id": "x",
        "setter": {"name": "set_x"},
    }
    normalized = normalize_series_entry(series)
    assert "setter" not in normalized
    assert normalized["input"]["setter"]["name"] == "set_x"
    assert has_input_direction(normalized)
    assert not has_output_direction(normalized)


def test_merge_input_and_output_shards(tmp_path: Path) -> None:
    input_doc = parse_bindings_file(FIXTURES / "shard_borvelia_input.yaml")
    output_doc = parse_bindings_file(FIXTURES / "shard_borvelia_output.yaml")
    merged = merge_series_binding_documents([input_doc, output_doc])
    series = merged["series"][0]
    assert has_input_direction(series)
    assert has_output_direction(series)
    assert series["input"]["setter"]["name"] == "set_borvelia_primary_balance"
    assert series["output"]["compute"]["name"] == "compute_borvelia_primary_balance"


def test_schema_requires_at_least_one_direction() -> None:
    from excel_grapher.series_bindings import SeriesBindingsSchemaError, validate_bindings_document

    doc = {
        "schema_version": "1.2.0",
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


def test_load_merged_input_output_directory(tmp_path: Path) -> None:
    shard_dir = tmp_path / "shards"
    shard_dir.mkdir()
    (shard_dir / "input.bindings.yaml").write_text(
        (FIXTURES / "shard_borvelia_input.yaml").read_text(encoding="utf-8"),
        encoding="utf-8",
    )
    (shard_dir / "output.bindings.yaml").write_text(
        (FIXTURES / "shard_borvelia_output.yaml").read_text(encoding="utf-8"),
        encoding="utf-8",
    )
    bindings = load_series_bindings(shard_dir)
    series = bindings["series"][0]
    assert has_input_direction(series)
    assert has_output_direction(series)


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
