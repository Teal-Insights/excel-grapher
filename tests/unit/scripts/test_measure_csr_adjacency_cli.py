"""CLI smoke tests for the CSR/CSC adjacency spike (#908)."""

from __future__ import annotations

import json

import pytest

from scripts.measure_csr_adjacency import DEFAULT_WORKBOOK, main

pytestmark = pytest.mark.skipif(
    not DEFAULT_WORKBOOK.is_file(), reason="taco_patterns.xlsx fixture missing"
)


def test_main_renders_a_comparison_table(capsys: pytest.CaptureFixture[str]) -> None:
    assert main([]) == 0
    out = capsys.readouterr().out
    assert "CSR/CSC:" in out
    assert "dict `_edges` exclusive" in out
    assert "node-index table exclusive" in out
    assert "csr_forward_ids" in out
    assert "second Python edge index" in out
    assert "landing API breaks" in out


def test_main_emits_json(capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["--json"]) == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["node_count"] > 0
    assert payload["csr_nnz"] > 0
    assert payload["csr_array_bytes"] > 0
    assert payload["adjacency_exclusive"] > payload["csr_exclusive_bytes"]
    assert payload["built_second_python_edge_index"] is False
    assert payload["held_input_adjacency"] is True
    names = {walk["name"] for walk in payload["walks"]}
    assert "csr_forward_ids" in names
    assert "dict_forward_sets" in names


def test_synthetic_preset_skips_workbook(capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["--preset", "synthetic", "--synthetic-nodes", "40", "--json"]) == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["node_count"] == 80
    assert payload["csr_n"] == 80
    assert payload["built_second_python_edge_index"] is False
