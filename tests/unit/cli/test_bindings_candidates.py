"""Tests for ``excel-grapher bindings candidates``."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
import yaml

from excel_grapher.cli import main
from excel_grapher.series_bindings.versions import CURRENT_SCHEMA_VERSION
from tests.unit.exporter.inverted_tree.helpers import series_entry, write_workbook


def _offset_workbook(path: Path) -> Path:
    return write_workbook(
        path,
        {
            "Inputs": {"B1": 1, "G1": 2},
            "Lookups": {"A1": 100, "B1": 1, "B2": 2},
            "Engine": {
                "A1": "=OFFSET(Lookups!A1,Inputs!B1,0)",
                "D1": "=INDEX(Lookups!A1:B2,Inputs!G1,2)",
            },
            "Outputs": {"A1": "=Engine!A1+Engine!D1"},
        },
    )


def _write_output_sidecar(path: Path) -> None:
    document = {
        "schema_version": CURRENT_SCHEMA_VERSION,
        "series": [
            series_entry(
                "result",
                "Outputs!A1",
                layout="scalar",
                direction="output",
                compute_name="compute_result",
            )
        ],
    }
    path.write_text(yaml.safe_dump(document, sort_keys=False), encoding="utf-8")


def test_main_candidates_lists_dynamic_ref_leaves(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    workbook = _offset_workbook(tmp_path / "dyn.xlsx")
    bindings = tmp_path / "dyn.bindings.yaml"
    _write_output_sidecar(bindings)

    exit_code = main(
        ["bindings", "candidates", str(workbook), "--bindings", str(bindings), "--json"]
    )
    captured = capsys.readouterr()
    assert exit_code == 0, captured.err
    assert json.loads(captured.out) == ["Inputs!B1", "Inputs!G1"]


def test_main_candidates_strict_fails_while_leaves_remain(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    workbook = _offset_workbook(tmp_path / "dyn.xlsx")
    bindings = tmp_path / "dyn.bindings.yaml"
    _write_output_sidecar(bindings)

    exit_code = main(
        ["bindings", "candidates", str(workbook), "--bindings", str(bindings), "--strict"]
    )
    captured = capsys.readouterr()
    assert exit_code == 1
    assert "Inputs!B1" in captured.out
    assert "Inputs!G1" in captured.out


def test_main_candidates_target_without_sidecar(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    workbook = _offset_workbook(tmp_path / "dyn.xlsx")
    exit_code = main(["bindings", "candidates", str(workbook), "--target", "Outputs!A1"])
    captured = capsys.readouterr()
    assert exit_code == 0, captured.err
    assert "Inputs!B1" in captured.out
    assert "Inputs!G1" in captured.out


def test_main_candidates_requires_a_root(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    workbook = _offset_workbook(tmp_path / "dyn.xlsx")
    exit_code = main(["bindings", "candidates", str(workbook)])
    captured = capsys.readouterr()
    assert exit_code == 1
    assert "target" in captured.err.lower() or "series" in captured.err.lower()


def test_main_candidates_missing_workbook(tmp_path: Path) -> None:
    missing = tmp_path / "missing.xlsx"
    exit_code = main(["bindings", "candidates", str(missing), "--target", "Outputs!A1"])
    assert exit_code == 1
