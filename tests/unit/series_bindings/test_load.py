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

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "series_bindings"


def test_load_yaml_binding_file() -> None:
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    assert bindings["schema_version"] == "1.0.0"
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


def test_merge_composes_identical_series_id() -> None:
    doc = parse_bindings_file(FIXTURES / "shard_inputs.yaml")
    merged = merge_series_binding_documents([doc, doc])
    assert len(merged["series"]) == 1


def test_merge_rejects_conflicting_series_id() -> None:
    doc = parse_bindings_file(FIXTURES / "shard_inputs.yaml")
    conflicting = parse_bindings_file(FIXTURES / "shard_inputs.yaml")
    conflicting["series"][0]["data_range"] = "Inputs!B3:C3"
    with pytest.raises(SeriesBindingsLoadError, match="structural fields differ"):
        merge_series_binding_documents([doc, conflicting])


def test_merge_rejects_workbook_mismatch() -> None:
    a = parse_bindings_file(FIXTURES / "shard_inputs.yaml")
    b = parse_bindings_file(FIXTURES / "shard_assumptions.yaml")
    b = {**b, "workbook": "other.xlsx"}
    with pytest.raises(SeriesBindingsLoadError, match="workbook mismatch"):
        merge_series_binding_documents([a, b])


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
