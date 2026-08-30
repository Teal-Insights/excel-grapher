"""CLI smoke tests for the formula-AST intern measurement harness (#550)."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from scripts.measure_formula_ast_intern import DEFAULT_FFV2, main, measure_intern


def test_main_emits_json(capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["--json"]) == 0
    payload = json.loads(capsys.readouterr().out)
    assert isinstance(payload, list)
    assert payload
    names = {item["name"] for item in payload}
    assert "same-text different-offset =A1*2" in names
    for item in payload:
        intern = item["intern"]
        assert intern["ast_to_json_calls"] == 0
        assert intern["json_dumps_intern_key_calls"] == 0
        assert intern["identity_distinct_trees"] == intern["equality_distinct_trees"]
        assert item["cache"]["uses_formula_ast_id"] is True
        assert item["cache"]["uses_formula_ast_key"] is False


@pytest.mark.skipif(not DEFAULT_FFV2.is_file(), reason="ffv2.xlsx fixture missing")
def test_ffv2_autofill_collapses_to_one_tree() -> None:
    report = measure_intern(DEFAULT_FFV2, ["Sheet1!B18:Q18"], name="ffv2")
    assert report.intern.formula_nodes == 16
    assert report.intern.identity_distinct_trees == 1
    assert report.intern.intern_hits == 15
    assert report.intern.ast_to_json_calls == 0
    assert report.cache.pool_entries == 1
    assert report.cache.uses_formula_ast_id is True


def test_same_text_different_offset_keeps_two_trees(tmp_path: Path) -> None:
    from scripts.measure_formula_ast_intern import _write_offset_workbook

    path = tmp_path / "offset.xlsx"
    _write_offset_workbook(path)
    report = measure_intern(path, ["Sheet1!B1", "Sheet1!B2"], name="offset")
    assert report.intern.formula_nodes == 2
    assert report.intern.identity_distinct_trees == 2
    assert report.intern.intern_hits == 0
    assert report.intern.ast_to_json_calls == 0
    assert report.cache.pool_entries == 2
