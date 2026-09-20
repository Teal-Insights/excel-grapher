"""Tests for surgical series-binding upsert."""

from __future__ import annotations

from pathlib import Path

import pytest
import yaml

from excel_grapher.series_bindings.upsert import (
    BindingUpsertError,
    bootstrap_binding_shards,
    upsert_series_binding,
)
from excel_grapher.series_bindings.versions import CURRENT_SCHEMA_VERSION
from tests.unit.series_bindings.authoring_helpers import (
    public_io_series,
    scalar_series,
    write_authoring_workbook,
    write_shards,
    years_internal_series,
)


def test_upsert_inserts_one_internal_series(tmp_path: Path) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx")
    inputs, result_a, result_b = public_io_series()
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
    )
    result = upsert_series_binding(
        workbook,
        bindings_dir,
        years_internal_series(fill=True),
    )
    assert result.series_id == "engine_years"
    assert result.replaced is False
    shard = yaml.safe_load(result.shard_path.read_text(encoding="utf-8"))
    assert [entry["id"] for entry in shard["series"]] == ["engine_years"]
    other = yaml.safe_load((bindings_dir / "inputs.bindings.yaml").read_text(encoding="utf-8"))
    assert [entry["id"] for entry in other["series"]] == ["input_rate"]


def test_upsert_rejects_duplicate_id_without_replace(tmp_path: Path) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx")
    inputs, result_a, result_b = public_io_series()
    series = years_internal_series(fill=True)
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
        internals=[series],
    )
    with pytest.raises(BindingUpsertError, match="already exists"):
        upsert_series_binding(workbook, bindings_dir, series)


def test_upsert_replace_updates_existing_id(tmp_path: Path) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx")
    inputs, result_a, result_b = public_io_series()
    original = years_internal_series(fill=True)
    original["notes"] = "original"
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
        internals=[original],
    )
    replacement = years_internal_series(fill=True)
    replacement["notes"] = "replaced"
    result = upsert_series_binding(
        workbook,
        bindings_dir,
        replacement,
        replace=True,
    )
    assert result.replaced is True
    shard = yaml.safe_load(result.shard_path.read_text(encoding="utf-8"))
    assert shard["series"][0]["notes"] == "replaced"
    assert len(shard["series"]) == 1


def test_upsert_rejects_formula_ownership_clash(tmp_path: Path) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx")
    inputs, result_a, result_b = public_io_series()
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
    )
    overlapping = scalar_series("engine_as_internal", "Outputs!B1", direction="internal")
    with pytest.raises(BindingUpsertError, match="ownership|audit|duplicate"):
        upsert_series_binding(workbook, bindings_dir, overlapping)


def test_upsert_rejects_unfilled_sparse_labels(tmp_path: Path) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx", sparse_year=True)
    inputs, result_a, result_b = public_io_series()
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
    )
    with pytest.raises(BindingUpsertError, match="sparse_label_without_fill|resolution"):
        upsert_series_binding(workbook, bindings_dir, years_internal_series(fill=False))


def test_upsert_places_series_by_groups_order(tmp_path: Path) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx", extra_engine_row=True)
    inputs, result_a, result_b = public_io_series()
    first = scalar_series("engine_b2", "Engine!B2", direction="internal")
    first["groups"] = [{"path": ["Engine"], "order": 1}]
    third = scalar_series("engine_c2", "Engine!C2", direction="internal")
    third["groups"] = [{"path": ["Engine"], "order": 3}]
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
        internals=[first, third],
    )
    middle = scalar_series("engine_b3", "Engine!B3", direction="internal")
    middle["groups"] = [{"path": ["Engine"], "order": 2}]
    result = upsert_series_binding(workbook, bindings_dir, middle)
    shard = yaml.safe_load(result.shard_path.read_text(encoding="utf-8"))
    assert [entry["id"] for entry in shard["series"]] == [
        "engine_b2",
        "engine_b3",
        "engine_c2",
    ]


def test_bootstrap_binding_shards_writes_empty_four_files(tmp_path: Path) -> None:
    directory = tmp_path / "model.bindings"
    written = bootstrap_binding_shards(directory, workbook="model.xlsx")
    assert {path.name for path in written} == {
        "inputs.bindings.yaml",
        "outputs.bindings.yaml",
        "internals.bindings.yaml",
        "constants.bindings.yaml",
    }
    for path in written:
        document = yaml.safe_load(path.read_text(encoding="utf-8"))
        assert document["schema_version"] == CURRENT_SCHEMA_VERSION
        assert document["series"] == []
        assert document["workbook"] == "model.xlsx"
