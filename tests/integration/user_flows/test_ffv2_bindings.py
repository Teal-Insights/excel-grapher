"""Integration tests for ffv2 game-log series bindings (datetime TIME_PERIOD keys)."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.series_bindings import run_binding_checks, validate_bindings_workbook
from excel_grapher.series_bindings.workflow import setter_names
from tests.integration.user_flows.utils import write_ffv2_workbook

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "series_bindings"


@pytest.fixture
def workbook(tmp_path: Path) -> Path:
    path = tmp_path / "ffv2.xlsx"
    write_ffv2_workbook(path)
    return path


@pytest.fixture
def bindings_path() -> Path:
    return FIXTURES / "ffv2.yaml"


def test_ffv2_bindings_validate(workbook: Path, bindings_path: Path) -> None:
    from excel_grapher.series_bindings import load_series_bindings

    bindings = load_series_bindings(bindings_path)
    result = validate_bindings_workbook(workbook, bindings_path)
    assert result["report"]["ok"] is True
    assert not any(issue["level"] == "error" for issue in result["report"]["issues"])
    assert len(result["input_series"]) == len(setter_names(bindings))


def test_ffv2_all_setters_and_computes(
    workbook: Path,
    bindings_path: Path,
    tmp_path: Path,
) -> None:
    result = run_binding_checks(
        workbook,
        bindings_path,
        module_dir=tmp_path / "bindings_module",
        package_name="bindings_module",
    )
    assert result["setters"] == [
        "set_puka_longest_reception",
        "set_puka_receptions",
        "set_puka_targets",
        "set_puka_touchdowns",
        "set_puka_week_1_stats",
        "set_puka_yards",
    ]
    assert result["computes"] == [
        "compute_puka_avg_yards_per_reception",
        "compute_puka_fantasy_score",
        "compute_puka_week_1_fantasy_score",
        "compute_touchdowns",
    ]
