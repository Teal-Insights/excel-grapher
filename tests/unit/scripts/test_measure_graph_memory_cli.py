"""CLI smoke tests for the graph memory measurement harness (#490)."""

from __future__ import annotations

import json

import pytest

from scripts.measure_graph_memory import DEFAULT_WORKBOOK, main

pytestmark = pytest.mark.skipif(
    not DEFAULT_WORKBOOK.is_file(), reason="taco_patterns.xlsx fixture missing"
)


def test_main_renders_a_table(capsys: pytest.CaptureFixture[str]) -> None:
    assert main([]) == 0
    out = capsys.readouterr().out
    assert "DependencyGraph:" in out
    assert "provenance" in out
    assert "graph total" in out


def test_main_emits_json(capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["--json"]) == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["node_count"] > 0
    assert payload["edge_count"] > 0
    assert payload["total_bytes"] > 0


def test_no_provenance_drops_edge_metadata(capsys: pytest.CaptureFixture[str]) -> None:
    main(["--json"])
    with_provenance = json.loads(capsys.readouterr().out)
    main(["--json", "--no-provenance"])
    without = json.loads(capsys.readouterr().out)
    assert without["total_bytes"] < with_provenance["total_bytes"]


def test_missing_workbook_exits_with_usage_error() -> None:
    with pytest.raises(SystemExit) as excinfo:
        main(["--workbook", "does-not-exist.xlsx"])
    assert excinfo.value.code == 2
