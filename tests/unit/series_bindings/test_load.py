"""Unit tests for series binding load and merge."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.series_bindings import (
    SeriesBindingsLoadError,
    bindings_canonical_sha256,
    load_series_bindings,
    merge_series_binding_documents,
    parse_bindings_file,
)
from tests.paths import SERIES_BINDINGS_FIXTURES as FIXTURES


def test_load_yaml_binding_file() -> None:
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    assert bindings["schema_version"] == "1.3.0"
    assert bindings["workbook"] == "lic_inputs.xlsx"
    assert len(bindings["series"]) == 1
    assert bindings["series"][0]["id"] == "borvelia_primary_balance"


def test_load_json_binding_file(tmp_path: Path) -> None:
    src = FIXTURES / "borvelia_primary_balance.yaml"
    doc = parse_bindings_file(src)
    path = tmp_path / "model.bindings.json"
    import json

    path.write_text(json.dumps(doc), encoding="utf-8")
    loaded = load_series_bindings(path)
    series = loaded["series"][0]
    assert series["input"]["setter"]["name"] == "set_borvelia_primary_balance"


def test_merge_directory_shards(tmp_path: Path) -> None:
    shard_dir = tmp_path / "shards"
    shard_dir.mkdir()
    (shard_dir / "Assumptions.bindings.yaml").write_text(
        (FIXTURES / "shard_assumptions.yaml").read_text(encoding="utf-8"),
        encoding="utf-8",
    )
    (shard_dir / "Inputs.bindings.yaml").write_text(
        (FIXTURES / "shard_inputs.yaml").read_text(encoding="utf-8"),
        encoding="utf-8",
    )

    merged = load_series_bindings(shard_dir)
    ids = [s["id"] for s in merged["series"]]
    assert ids == ["row_b", "row_a"]


def test_merge_composes_identical_series_id_across_shards() -> None:
    doc = parse_bindings_file(FIXTURES / "shard_inputs.yaml")
    merged = merge_series_binding_documents([doc, doc])
    assert len(merged["series"]) == 1


def test_merge_rejects_duplicate_series_id_within_single_document() -> None:
    doc = parse_bindings_file(FIXTURES / "shard_inputs.yaml")
    doc["series"].append(dict(doc["series"][0]))
    with pytest.raises(SeriesBindingsLoadError, match="Duplicate series id"):
        merge_series_binding_documents([doc])


def test_merge_rejects_conflicting_series_id() -> None:
    doc = parse_bindings_file(FIXTURES / "shard_inputs.yaml")
    conflicting = parse_bindings_file(FIXTURES / "shard_inputs.yaml")
    conflicting["series"][0]["structure"]["dimensions"][0]["bind"]["header_row"] = 2
    with pytest.raises(SeriesBindingsLoadError, match="structural fields differ"):
        merge_series_binding_documents([doc, conflicting])


def test_merge_rejects_workbook_mismatch() -> None:
    a = parse_bindings_file(FIXTURES / "shard_inputs.yaml")
    b = parse_bindings_file(FIXTURES / "shard_assumptions.yaml")
    b = {**b, "workbook": "other.xlsx"}
    with pytest.raises(SeriesBindingsLoadError, match="workbook mismatch"):
        merge_series_binding_documents([a, b])


def test_merge_skips_empty_series_shards() -> None:
    empty = {
        "schema_version": "1.10.0",
        "workbook": "workbook.xlsx",
        "concept_scheme": {"id": "inputs_placeholder", "concepts": []},
        "series": [],
    }
    populated = parse_bindings_file(FIXTURES / "shard_inputs.yaml")
    populated = {
        **populated,
        "schema_version": "1.10.0",
        "workbook": "workbook.xlsx",
    }
    merged = merge_series_binding_documents([empty, populated])
    assert [s["id"] for s in merged["series"]] == ["row_a"]


def test_load_directory_of_empty_series_placeholders(tmp_path: Path) -> None:
    root = tmp_path / "bindings"
    root.mkdir()
    for name, scheme_id in (
        ("inputs.bindings.yaml", "inputs_placeholder"),
        ("outputs.bindings.yaml", "outputs_placeholder"),
        ("internals.bindings.yaml", "internals_placeholder"),
    ):
        (root / name).write_text(
            f"""schema_version: 1.10.0
workbook: workbook.xlsx
concept_scheme:
  id: {scheme_id}
  concepts: []
series: []
""",
            encoding="utf-8",
        )

    loaded = load_series_bindings(root)
    assert loaded["schema_version"] == "1.10.0"
    assert loaded["workbook"] == "workbook.xlsx"
    assert loaded["series"] == []


def test_merge_unions_concept_scheme_across_shards() -> None:
    left = {
        "schema_version": "1.10.0",
        "workbook": "workbook.xlsx",
        "concept_scheme": {
            "id": "inputs_scheme",
            "concepts": [
                {"id": "PARAMETER", "dtype": "string"},
                {"id": "TIME_PERIOD", "dtype": "int"},
            ],
        },
        "series": [
            {
                "id": "param_series",
                "sheet": "Inputs",
                "data_range": "Inputs!B2",
                "layout": "scalar",
                "input": {"setter": {"name": "set_param_series"}},
                "structure": {
                    "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
                    "dimensions": [
                        {
                            "concept": "PARAMETER",
                            "role": "key",
                            "scope": "cell",
                            "bind": {"kind": "constant", "value": "x"},
                        }
                    ],
                },
                "key": ["PARAMETER"],
            }
        ],
    }
    right = {
        "schema_version": "1.10.0",
        "workbook": "workbook.xlsx",
        "concept_scheme": {
            "id": "outputs_scheme",
            "concepts": [
                {"id": "INDICATOR", "dtype": "string"},
                {"id": "TIME_PERIOD", "dtype": "int"},
            ],
        },
        "series": [
            {
                "id": "indicator_series",
                "sheet": "Outputs",
                "data_range": "Outputs!B2",
                "layout": "scalar",
                "output": {"compute": {"name": "compute_indicator_series"}},
                "structure": {
                    "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
                    "dimensions": [
                        {
                            "concept": "INDICATOR",
                            "role": "key",
                            "scope": "cell",
                            "bind": {"kind": "constant", "value": "y"},
                        }
                    ],
                },
                "key": ["INDICATOR"],
            }
        ],
    }

    merged = merge_series_binding_documents([left, right])
    concepts = merged["concept_scheme"]["concepts"]
    assert [c["id"] for c in concepts] == ["PARAMETER", "TIME_PERIOD", "INDICATOR"]
    assert merged["concept_scheme"]["id"] == "inputs_scheme"
    assert {s["id"] for s in merged["series"]} == {"param_series", "indicator_series"}


def test_merge_rejects_conflicting_concept_definitions() -> None:
    left = {
        "schema_version": "1.10.0",
        "concept_scheme": {
            "id": "scheme_a",
            "concepts": [{"id": "TIME_PERIOD", "dtype": "int"}],
        },
        "series": [
            {
                "id": "a",
                "sheet": "S",
                "data_range": "S!A1",
                "layout": "scalar",
                "input": {"setter": {"name": "set_a"}},
                "structure": {
                    "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
                    "dimensions": [
                        {
                            "concept": "TIME_PERIOD",
                            "role": "key",
                            "scope": "cell",
                            "bind": {"kind": "constant", "value": 1},
                        }
                    ],
                },
                "key": ["TIME_PERIOD"],
            }
        ],
    }
    right = {
        "schema_version": "1.10.0",
        "concept_scheme": {
            "id": "scheme_b",
            "concepts": [{"id": "TIME_PERIOD", "dtype": "datetime"}],
        },
        "series": [
            {
                "id": "b",
                "sheet": "S",
                "data_range": "S!B1",
                "layout": "scalar",
                "input": {"setter": {"name": "set_b"}},
                "structure": {
                    "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
                    "dimensions": [
                        {
                            "concept": "TIME_PERIOD",
                            "role": "key",
                            "scope": "cell",
                            "bind": {"kind": "constant", "value": 2},
                        }
                    ],
                },
                "key": ["TIME_PERIOD"],
            }
        ],
    }
    with pytest.raises(SeriesBindingsLoadError, match="concept_scheme"):
        merge_series_binding_documents([left, right])


def test_canonical_hash_stable_under_key_order() -> None:
    doc_a = {
        "schema_version": "1.0.0",
        "workbook": "w.xlsx",
        "series": [{"id": "a", "z": 1, "sheet": "S"}],
    }
    doc_b = {
        "series": [{"sheet": "S", "id": "a", "z": 1}],
        "workbook": "w.xlsx",
        "schema_version": "1.0.0",
    }
    assert bindings_canonical_sha256(doc_a) == bindings_canonical_sha256(doc_b)


def test_load_missing_path_raises() -> None:
    with pytest.raises(SeriesBindingsLoadError, match="does not exist"):
        load_series_bindings("/no/such/bindings.bindings.yaml")
