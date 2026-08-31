"""CLI smoke tests for the leaf-store measurement harness (#579)."""

from __future__ import annotations

import json

import pytest

from scripts.measure_leaf_store import (
    DEFAULT_WORKBOOK,
    emit_coordinate_store_literal,
    emit_nodekey_dict_literal,
    main,
    measure_leaf_payload,
    synthetic_leaves,
)

pytestmark = pytest.mark.skipif(
    not DEFAULT_WORKBOOK.is_file(), reason="taco_patterns.xlsx fixture missing"
)


def test_coordinate_literal_is_smaller_than_nodekey_keys() -> None:
    leaves = synthetic_leaves(200, sheet="Sheet1")
    nodekey = emit_nodekey_dict_literal(leaves)
    coord = emit_coordinate_store_literal(leaves)
    assert "Sheet1!A1" in nodekey
    assert "'Sheet1': {" in coord
    assert "(1, 1):" in coord
    assert "Sheet1!A1" not in coord
    assert len(coord.encode("utf-8")) < len(nodekey.encode("utf-8"))


def test_measure_leaf_payload_reports_required_metrics() -> None:
    payload = measure_leaf_payload(synthetic_leaves(50), import_repeats=1, scan_repeats=1)
    assert payload["occupied_leaves"] == 50
    assert payload["distinct_sheets"] == 1
    assert payload["coordinate_bytes"] < payload["nodekey_bytes"]
    assert payload["bytes_ratio"] < 1.0
    assert payload["xl_range_coordinate_s"] >= 0.0
    assert payload["make_context_overlay_ok"] is True


def test_main_emits_json(capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["--json", "--synthetic-leaves", "80", "--import-repeats", "1"]) == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["occupied_leaves"] == 80
    assert payload["distinct_sheets"] == 1
    assert payload["coordinate_bytes"] < payload["nodekey_bytes"]
