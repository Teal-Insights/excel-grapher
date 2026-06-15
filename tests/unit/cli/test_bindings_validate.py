"""Tests for ``excel-grapher bindings validate``."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest

from excel_grapher.cli import main
from tests.integration.user_flows.utils import write_ffv2_workbook

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "series_bindings"


def test_main_missing_workbook(tmp_path: Path) -> None:
    missing = tmp_path / "missing.xlsx"
    exit_code = main(["bindings", "validate", str(missing)])
    assert exit_code == 1


def test_main_validate_ffv2_fixture(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    workbook = tmp_path / "ffv2.xlsx"
    write_ffv2_workbook(workbook)
    bindings = FIXTURES / "ffv2.yaml"

    exit_code = main(
        [
            "bindings",
            "validate",
            str(workbook),
            "--bindings",
            str(bindings),
        ]
    )

    captured = capsys.readouterr()
    assert exit_code == 0
    assert "ok=True" in captured.out
    assert "set_puka_receptions" in captured.out


def test_main_validate_json_output(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    workbook = tmp_path / "ffv2.xlsx"
    write_ffv2_workbook(workbook)
    bindings = FIXTURES / "ffv2.yaml"

    exit_code = main(
        [
            "bindings",
            "validate",
            str(workbook),
            "--bindings",
            str(bindings),
            "--json",
        ]
    )

    captured = capsys.readouterr()
    assert exit_code == 0
    payload = json.loads(captured.out)
    assert payload["ok"] is True


def test_main_smoke_test_ffv2_fixture(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    workbook = tmp_path / "ffv2.xlsx"
    write_ffv2_workbook(workbook)
    bindings = FIXTURES / "ffv2.yaml"

    exit_code = main(
        [
            "bindings",
            "validate",
            str(workbook),
            "--bindings",
            str(bindings),
            "--smoke-test",
        ]
    )

    captured = capsys.readouterr()
    assert exit_code == 0
    assert "passed smoke checks" in captured.out


def test_console_script_is_registered() -> None:
    result = subprocess.run(
        [sys.executable, "-m", "excel_grapher.cli", "bindings", "validate", "--help"],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0
    assert "--smoke-test" in result.stdout
