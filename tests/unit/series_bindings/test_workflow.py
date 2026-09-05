"""Tests for series binding workflow helpers."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.series_bindings import SeriesBindingsLoadError
from excel_grapher.series_bindings.workflow import (
    resolve_bindings_path,
    validate_bindings_workbook,
)
from tests.integration.user_flows.utils import write_ffv2_workbook
from tests.paths import SERIES_BINDINGS_FIXTURES as FIXTURES


def test_resolve_bindings_path_uses_colocated_yaml(tmp_path: Path) -> None:
    workbook = tmp_path / "model.xlsx"
    workbook.write_bytes(b"")
    bindings = tmp_path / "model.bindings.yaml"
    bindings.write_text("schema_version: '1.4.0'\nseries: []\n", encoding="utf-8")

    assert resolve_bindings_path(workbook) == bindings


def test_resolve_bindings_path_uses_shard_directory(tmp_path: Path) -> None:
    workbook = tmp_path / "model.xlsx"
    workbook.write_bytes(b"")
    shard_dir = tmp_path / "model.bindings"
    shard_dir.mkdir()
    shard = shard_dir / "Inputs.bindings.yaml"
    shard.write_text("schema_version: '1.4.0'\nseries: []\n", encoding="utf-8")

    assert resolve_bindings_path(workbook) == shard_dir


def test_resolve_bindings_path_raises_when_missing(tmp_path: Path) -> None:
    workbook = tmp_path / "model.xlsx"
    workbook.write_bytes(b"")

    with pytest.raises(SeriesBindingsLoadError, match="No binding sidecar found"):
        resolve_bindings_path(workbook)


def test_resolve_bindings_path_explicit_folder_relative_to_workbook(
    tmp_path: Path,
) -> None:
    project = tmp_path / "project"
    project.mkdir()
    workbook = project / "model.xlsx"
    workbook.write_bytes(b"")
    shard_dir = project / "model.bindings"
    shard_dir.mkdir()
    (shard_dir / "Inputs.bindings.yaml").write_text(
        "schema_version: '1.4.0'\nseries: []\n",
        encoding="utf-8",
    )

    resolved = resolve_bindings_path(workbook, Path("model.bindings"))

    assert resolved == shard_dir


def test_validate_bindings_workbook_ffv2_fixture(tmp_path: Path) -> None:
    workbook = tmp_path / "ffv2.xlsx"
    write_ffv2_workbook(workbook)
    bindings_path = FIXTURES / "ffv2.yaml"

    result = validate_bindings_workbook(workbook, bindings_path)

    assert result["report"]["ok"] is True
    assert len(result["setters"]) == 6
    assert len(result["computes"]) == 4
