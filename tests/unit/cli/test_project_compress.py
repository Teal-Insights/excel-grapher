"""Tests for ``excel-grapher project compress``."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from excel_grapher.cli import main
from tests.fixtures.tiny_dsa.workbook import TINY_DSA_TARGETS, build_tiny_dsa_workbook


def test_project_compress_missing_workbook(tmp_path: Path) -> None:
    missing = tmp_path / "missing.xlsx"
    exit_code = main(
        [
            "project",
            "compress",
            str(missing),
            "--targets",
            "Engine!C20",
        ]
    )
    assert exit_code == 1


def test_project_compress_tiny_dsa_similarity(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    workbook = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(workbook)
    exit_code = main(
        [
            "project",
            "compress",
            str(workbook),
            "--targets",
            *TINY_DSA_TARGETS,
            "--method",
            "similarity",
        ]
    )
    captured = capsys.readouterr()
    assert exit_code == 0, captured.err
    assert "Method: similarity" in captured.out
    assert "18 removed" in captured.out


def test_project_compress_json_output(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    workbook = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(workbook)
    manifest_path = tmp_path / "manifest.json"
    exit_code = main(
        [
            "project",
            "compress",
            str(workbook),
            "--targets",
            "Engine!C20",
            "Engine!H20",
            "--method",
            "optimal",
            "--json",
            "--manifest-out",
            str(manifest_path),
        ]
    )
    captured = capsys.readouterr()
    assert exit_code == 0, captured.err
    payload = json.loads(captured.out)
    assert payload["method"] == "optimal"
    assert payload["manifest_kind"] == "optimal_compression"
    assert manifest_path.is_file()
