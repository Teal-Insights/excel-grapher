"""Tests for ``excel-grapher bindings upsert``."""

from __future__ import annotations

from pathlib import Path

import pytest
import yaml

from excel_grapher.cli import main
from tests.unit.series_bindings.authoring_helpers import (
    public_io_series,
    write_authoring_workbook,
    write_shards,
    years_internal_series,
)


def test_bindings_upsert_writes_one_internal_series(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx")
    inputs, result_a, result_b = public_io_series()
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
    )
    series_path = tmp_path / "engine_years.yaml"
    series_path.write_text(
        yaml.safe_dump(years_internal_series(fill=True), sort_keys=False),
        encoding="utf-8",
    )
    exit_code = main(
        [
            "bindings",
            "upsert",
            str(workbook),
            "--bindings",
            str(bindings_dir),
            "--series",
            str(series_path),
        ]
    )
    captured = capsys.readouterr()
    assert exit_code == 0, captured.err
    shard = yaml.safe_load((bindings_dir / "internals.bindings.yaml").read_text(encoding="utf-8"))
    assert [entry["id"] for entry in shard["series"]] == ["engine_years"]
    assert "engine_years" in captured.out
