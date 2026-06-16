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


def test_main_validate_verbose_prints_warnings(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    import yaml

    workbook = tmp_path / "ffv2.xlsx"
    write_ffv2_workbook(workbook)
    bindings = tmp_path / "dtype_mismatch.yaml"
    document = yaml.safe_load((FIXTURES / "ffv2.yaml").read_text(encoding="utf-8"))
    document["concept_scheme"]["concepts"][0]["dtype"] = "int"
    bindings.write_text(yaml.safe_dump(document, sort_keys=False), encoding="utf-8")

    exit_code = main(
        [
            "bindings",
            "validate",
            str(workbook),
            "--bindings",
            str(bindings),
            "-v",
        ]
    )

    captured = capsys.readouterr()
    assert exit_code == 0
    assert "ok=True" in captured.out
    assert "warning [dtype_read_mismatch]" in captured.out


def test_main_validate_prints_errors_without_verbose(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    import yaml

    workbook = tmp_path / "ffv2.xlsx"
    write_ffv2_workbook(workbook)
    bindings = tmp_path / "bad_read.yaml"
    document = yaml.safe_load((FIXTURES / "ffv2.yaml").read_text(encoding="utf-8"))
    document["series"][0]["structure"]["measure"]["bind"]["read"] = "bool"
    bindings.write_text(yaml.safe_dump(document, sort_keys=False), encoding="utf-8")

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
    assert exit_code == 1
    assert "ok=False" in captured.out
    assert "error [" in captured.out


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


def test_main_smoke_test_exits_nonzero_when_setters_collide(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    import yaml

    workbook = tmp_path / "ffv2.xlsx"
    write_ffv2_workbook(workbook)
    bindings = tmp_path / "dup_setter.yaml"
    document = yaml.safe_load((FIXTURES / "ffv2.yaml").read_text(encoding="utf-8"))
    document["series"][1]["input"]["setter"]["name"] = document["series"][0]["input"]["setter"][
        "name"
    ]
    bindings.write_text(yaml.safe_dump(document, sort_keys=False), encoding="utf-8")

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
    assert exit_code == 1
    assert "passed smoke checks" not in captured.out
    assert "Setter 'set_puka_receptions' did not update" in captured.err


def test_console_script_is_registered() -> None:
    result = subprocess.run(
        [sys.executable, "-m", "excel_grapher.cli", "bindings", "validate", "--help"],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0
    assert "--smoke-test" in result.stdout
    assert "--verbose" in result.stdout


def test_main_schema_error_is_human_readable(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
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
    assert exit_code == 1
    assert "Binding sidecar schema error:" in captured.err
    assert 'series[8] "puka_week_1_fantasy_score"' in captured.err
    assert "missing required field `key`" in captured.err
    assert captured.err.count("Traceback") == 0


def test_main_validate_bind_resolution_failed_for_row_label_in_measure_cell(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:

    import fastpyxl
    import yaml

    workbook = tmp_path / "ffv2.xlsx"
    write_ffv2_workbook(workbook)
    wb = fastpyxl.load_workbook(workbook)
    wb["Sheet1"]["A4"].value = "Tgts"
    wb.save(workbook)

    bindings = tmp_path / "puka_targets_row_label.yaml"
    document = yaml.safe_load((FIXTURES / "ffv2.yaml").read_text(encoding="utf-8"))
    document["series"] = [
        series for series in document["series"] if series["id"] != "puka_week_1_fantasy_score"
    ]
    for series in document["series"]:
        if series["id"] == "puka_targets":
            series["data_range"] = "Sheet1!A4:Q4"
            break
    bindings.write_text(yaml.safe_dump(document, sort_keys=False), encoding="utf-8")

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
    assert exit_code == 1
    assert (
        "error [bind_resolution_failed] puka_targets:Sheet1!A4: "
        "could not convert string to float: 'Tgts'"
    ) in captured.out
