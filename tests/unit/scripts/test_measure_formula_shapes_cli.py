"""CLI smoke tests for the formula-shape measurement harness (#517)."""

from __future__ import annotations

import json

import pytest

from scripts.measure_formula_shapes import DEFAULT_WORKBOOK, main

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


def test_missing_workbook_exits_with_usage_error() -> None:
    with pytest.raises(SystemExit) as excinfo:
        main(["--workbook", "does-not-exist.xlsx"])
    assert excinfo.value.code == 2
