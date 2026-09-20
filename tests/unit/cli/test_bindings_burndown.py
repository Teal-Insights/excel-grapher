"""Tests for ``excel-grapher bindings burndown``."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.cli import main
from tests.unit.series_bindings.authoring_helpers import (
    public_io_series,
    write_authoring_workbook,
    write_shards,
    years_internal_series,
)


def test_bindings_burndown_prints_coverage_worklist(
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
    exit_code = main(
        [
            "bindings",
            "burndown",
            str(workbook),
            "--bindings",
            str(bindings_dir),
        ]
    )
    captured = capsys.readouterr()
    assert exit_code == 0, captured.err
    assert "coverage worklist" in captured.out.lower()
    assert "Engine!B2:C2" in captured.out
    assert "not a generator" in captured.out.lower()


def test_bindings_burndown_strict_exits_nonzero_when_unbound(
    tmp_path: Path,
) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx")
    inputs, result_a, result_b = public_io_series()
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
    )
    exit_code = main(
        [
            "bindings",
            "burndown",
            str(workbook),
            "--bindings",
            str(bindings_dir),
            "--strict",
        ]
    )
    assert exit_code == 1


def test_bindings_burndown_exempt_file_clears_worklist(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx")
    inputs, result_a, result_b = public_io_series()
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
        internals=[years_internal_series(data_range="Engine!B2", fill=True)],
    )
    exempt = tmp_path / "exempt.txt"
    exempt.write_text("Engine!C2\n", encoding="utf-8")
    exit_code = main(
        [
            "bindings",
            "burndown",
            str(workbook),
            "--bindings",
            str(bindings_dir),
            "--exempt",
            str(exempt),
            "--strict",
        ]
    )
    captured = capsys.readouterr()
    assert exit_code == 0, captured.err
    assert "Unbound internal formula cells: 0" in captured.out
