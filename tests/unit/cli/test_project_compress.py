"""Tests for ``excel-grapher project compress``."""

from __future__ import annotations

import importlib
import json
from pathlib import Path
from types import ModuleType

import pytest

from excel_grapher.cli import main
from excel_grapher.exporter.compression_workflow import compress_workbook
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


def test_project_compress_preserve_nodes(tmp_path: Path) -> None:
    workbook = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(workbook)

    _, without_preserve_report = compress_workbook(
        workbook,
        list(TINY_DSA_TARGETS),
        method="similarity",
    )
    projection, with_preserve_report = compress_workbook(
        workbook,
        list(TINY_DSA_TARGETS),
        method="similarity",
        preserve={"Engine!C16"},
    )
    assert with_preserve_report.removed_count == without_preserve_report.removed_count - 1
    assert "Engine!C16" in projection.projected_graph
    assert "Engine!C16" not in with_preserve_report.removed_nodes


def test_project_compress_cli_preserve_flag(
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
            "--preserve",
            "Engine!C16",
            "--json",
        ]
    )
    captured = capsys.readouterr()
    assert exit_code == 0, captured.err
    payload = json.loads(captured.out)
    assert payload["removed_count"] == 17
    assert "Engine!C16" not in payload["removed_nodes"]


def test_project_compress_similarity_config_flags(
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
            "--embedding-provider",
            "mock",
            "--max-candidates",
            "200",
            "--top-n-packings",
            "50",
            "--score-flatness-epsilon",
            "0.01",
            "--json",
        ]
    )
    captured = capsys.readouterr()
    assert exit_code == 0, captured.err
    payload = json.loads(captured.out)
    assert payload["method"] == "similarity"
    assert payload["removed_count"] == 18
    assert "score" in payload


def test_project_compress_openai_provider_without_dependency_reports_error(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(workbook)

    real_import_module = importlib.import_module

    def _import_openai_only(name: str, package: str | None = None) -> ModuleType:
        if name == "openai":
            raise ImportError("no openai")
        return real_import_module(name, package)

    monkeypatch.setattr(importlib, "import_module", _import_openai_only)
    exit_code = main(
        [
            "project",
            "compress",
            str(workbook),
            "--targets",
            "Engine!C20",
            "--method",
            "similarity",
            "--embedding-provider",
            "openai",
        ]
    )
    captured = capsys.readouterr()
    assert exit_code == 1
    assert "Embedding provider error" in captured.err
