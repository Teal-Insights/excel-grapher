"""Tests for ``excel-grapher bindings audit``."""

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


def test_bindings_audit_exits_nonzero_on_sparse_labels(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx", sparse_year=True)
    inputs, result_a, result_b = public_io_series()
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
        internals=[years_internal_series(fill=False)],
    )
    exit_code = main(
        [
            "bindings",
            "audit",
            str(workbook),
            "--bindings",
            str(bindings_dir),
        ]
    )
    captured = capsys.readouterr()
    assert exit_code == 1
    assert "sparse_label_without_fill" in captured.out
    assert "engine_years" in captured.out


def test_bindings_audit_json_and_strict_warning(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from tests.unit.series_bindings.authoring_helpers import scalar_series

    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx")
    inputs, result_a, result_b = public_io_series()
    missing = scalar_series("missing_output", "Engine!B2", direction="output")
    missing["exclude_rows"] = [2]
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b, missing],
    )
    exit_code = main(
        [
            "bindings",
            "audit",
            str(workbook),
            "--bindings",
            str(bindings_dir),
            "--json",
        ]
    )
    captured = capsys.readouterr()
    assert exit_code == 0
    assert "empty_public_series" in captured.out

    exit_code = main(
        [
            "bindings",
            "audit",
            str(workbook),
            "--bindings",
            str(bindings_dir),
            "--strict",
        ]
    )
    assert exit_code == 1


def test_bindings_audit_clean_exit_zero(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx")
    inputs, result_a, result_b = public_io_series()
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
        internals=[years_internal_series(fill=True)],
    )
    exit_code = main(
        [
            "bindings",
            "audit",
            str(workbook),
            "--bindings",
            str(bindings_dir),
        ]
    )
    captured = capsys.readouterr()
    assert exit_code == 0, captured.err
    assert "0 error(s)" in captured.out
