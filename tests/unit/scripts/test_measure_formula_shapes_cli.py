"""CLI smoke tests for the formula-shape measurement harness (#517)."""

from __future__ import annotations

import json
from pathlib import Path

import fastpyxl
import pytest

from scripts.measure_formula_shapes import (
    DEFAULT_WORKBOOK,
    main,
    measure_parse_warm_times,
    summarize_scanned_formula_shapes,
)

pytestmark = pytest.mark.skipif(
    not DEFAULT_WORKBOOK.is_file(), reason="taco_patterns.xlsx fixture missing"
)


def test_main_renders_cardinality_report(capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["--no-parse-timing"]) == 0
    out = capsys.readouterr().out
    assert "distinct normalized formulas:" in out
    assert "distinct shapes:" in out
    assert "shapes / formula strings:" in out


def test_main_emits_json(capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["--json", "--no-parse-timing"]) == 0
    payload = json.loads(capsys.readouterr().out)
    summary = payload["summary"]
    assert summary["formula_nodes"] == 24
    assert summary["distinct_normalized_formulas"] == 24
    assert summary["distinct_shapes"] == 4
    assert summary["shapes_per_formula_string"] == pytest.approx(4 / 24)


def test_main_scan_workbook_emits_json(capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["--json", "--scan-workbook", "--no-parse-timing"]) == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["targets"] == "(scan-workbook: all formula cells)"
    summary = payload["summary"]
    assert summary["formula_nodes"] >= 24
    assert summary["distinct_shapes"] >= 4
    assert summary["shapes_per_formula_string"] < 1.0


def test_summarize_scanned_formula_shapes_normalizes_then_fingerprints(
    tmp_path: Path,
) -> None:
    path = tmp_path / "scan.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 1
    ws["B1"].value = "=A1+1"
    ws["B2"].value = "=A1+1"
    ws["C1"].value = "=1+"
    wb.save(path)
    wb.close()

    from scripts.measure_formula_shapes import _scan_workbook_formulas

    scanned = _scan_workbook_formulas(path)
    summary, parseable = summarize_scanned_formula_shapes(scanned)
    assert summary.formula_nodes == 2
    assert summary.distinct_shapes == 1
    assert summary.unparseable == 1
    assert summary.mean_instances_per_shape == 2.0
    assert len(parseable) == 2


def test_measure_parse_warm_times_reports_shape_collapse() -> None:
    formulas = [
        "=Sheet1!A1+Sheet1!B1",
        "=Sheet1!A2+Sheet1!B2",
        "=Sheet1!A3+Sheet1!B3",
        "=SUM(Sheet1!A1:A3)",
    ]
    times = measure_parse_warm_times(formulas, repeats=1)
    assert times["distinct_formulas"] == 4.0
    assert times["distinct_shapes"] == 2.0
    assert times["string_keyed_parse_s"] >= times["shape_keyed_parse_s"]
    assert times["repeats"] == 1.0


def test_missing_workbook_exits_with_usage_error() -> None:
    with pytest.raises(SystemExit) as excinfo:
        main(["--workbook", "does-not-exist.xlsx"])
    assert excinfo.value.code == 2
