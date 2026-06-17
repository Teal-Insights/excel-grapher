"""Accuracy tests for ffv3_large player sheet bindings."""

from __future__ import annotations

from pathlib import Path
from typing import cast

import pytest

from excel_grapher.series_bindings import load_series_bindings
from tests.integration.user_flows.bindings_accuracy import (
    BindingsAccuracyCase,
    DownstreamUpdateCase,
    SeriesSpotCheck,
    assert_all_compute_functions_match_workbook,
    assert_bindings_validate,
    assert_downstream_update,
    assert_series_spot_check,
    assert_shared_game_log_columns,
    build_dependency_graph,
    generate_bindings_namespace,
)

EXAMPLES = Path(__file__).resolve().parents[3] / "examples" / "micro_workbooks"
WORKBOOK = EXAMPLES / "ffv3_large.xlsx"
BINDINGS_DIR = EXAMPLES / "ffv3_large.bindings"


def _player_case(
    slug: str,
    bindings_file: str,
    *,
    reg_stat_leaves: int,
    playoff_stat_leaves: int,
    reg_sample_key: dict[str, str],
    reg_sample_value: object,
    playoff_sample_key: dict[str, str],
    playoff_sample_value: object,
    setter_indicator: str,
    setter_bump_value: int,
    expected_fantasy_after_bump: float,
) -> BindingsAccuracyCase:
    return BindingsAccuracyCase(
        name=slug,
        workbook=WORKBOOK,
        bindings_path=BINDINGS_DIR / bindings_file,
        setter_name_prefix=slug,
        compute_name_prefix=slug,
        expected_setter_count=4,
        expected_compute_count=4,
        series_checks=(
            SeriesSpotCheck(
                series_id=f"{slug}_reg_season_stats",
                leaf_count=reg_stat_leaves,
                sample_key=reg_sample_key,
                sample_value=reg_sample_value,
                unique_key_fields=("GAME_DATE", "INDICATOR"),
            ),
            SeriesSpotCheck(
                series_id=f"{slug}_playoff_stats",
                leaf_count=playoff_stat_leaves,
                sample_key=playoff_sample_key,
                sample_value=playoff_sample_value,
            ),
        ),
        downstream_update=DownstreamUpdateCase(
            setter_name=f"set_{slug}_reg_season_stats",
            setter_records=(
                {
                    "GAME_DATE": "Sep 7",
                    "INDICATOR": setter_indicator,
                    "OBS_VALUE": setter_bump_value,
                },
            ),
            compute_name=f"compute_{slug}_reg_fantasy_score",
            record_key={"GAME_DATE": "Sep 7", "INDICATOR": "Fantasy Score"},
            expected_obs_value=expected_fantasy_after_bump,
        ),
    )


PLAYER_CASES = (
    _player_case(
        "stafford",
        "qb_stafford.bindings.yaml",
        reg_stat_leaves=136,
        playoff_stat_leaves=24,
        reg_sample_key={"GAME_DATE": "Sep 7", "INDICATOR": "Pass Att"},
        reg_sample_value=29,
        playoff_sample_key={"GAME_DATE": "Jan 10", "INDICATOR": "Pass Att"},
        playoff_sample_value=42,
        setter_indicator="Pass Yds",
        setter_bump_value=300,
        expected_fantasy_after_bump=13.8,
    ),
    _player_case(
        "k_williams",
        "rb_k_williams.bindings.yaml",
        reg_stat_leaves=51,
        playoff_stat_leaves=9,
        reg_sample_key={"GAME_DATE": "Sep 7", "INDICATOR": "Rush Att"},
        reg_sample_value=18,
        playoff_sample_key={"GAME_DATE": "Jan 10", "INDICATOR": "Rush Att"},
        playoff_sample_value=13,
        setter_indicator="Rush Yds",
        setter_bump_value=300,
        expected_fantasy_after_bump=36.0,
    ),
    _player_case(
        "b_corum",
        "rb_b_corum.bindings.yaml",
        reg_stat_leaves=51,
        playoff_stat_leaves=9,
        reg_sample_key={"GAME_DATE": "Sep 7", "INDICATOR": "Rush Att"},
        reg_sample_value=1,
        playoff_sample_key={"GAME_DATE": "Jan 10", "INDICATOR": "Rush Att"},
        playoff_sample_value=11,
        setter_indicator="Rush Yds",
        setter_bump_value=300,
        expected_fantasy_after_bump=30.0,
    ),
    _player_case(
        "p_nacua",
        "wr_p_nacua.bindings.yaml",
        reg_stat_leaves=68,
        playoff_stat_leaves=12,
        reg_sample_key={"GAME_DATE": "Sep 7", "INDICATOR": "Touchdowns"},
        reg_sample_value="N/A",
        playoff_sample_key={"GAME_DATE": "Jan 10", "INDICATOR": "Receptions"},
        playoff_sample_value=10,
        setter_indicator="Rec Yds",
        setter_bump_value=300,
        expected_fantasy_after_bump=40.0,
    ),
    _player_case(
        "d_adams",
        "wr_d_adams.bindings.yaml",
        reg_stat_leaves=68,
        playoff_stat_leaves=12,
        reg_sample_key={"GAME_DATE": "Sep 7", "INDICATOR": "Targets"},
        reg_sample_value="N/A",
        playoff_sample_key={"GAME_DATE": "Jan 10", "INDICATOR": "Receptions"},
        playoff_sample_value=5,
        setter_indicator="Rec Yds",
        setter_bump_value=300,
        expected_fantasy_after_bump=34.0,
    ),
)


@pytest.fixture(scope="module")
def workbook() -> Path:
    if not WORKBOOK.is_file():
        pytest.skip(f"Workbook fixture missing: {WORKBOOK}")
    return WORKBOOK


@pytest.fixture(params=PLAYER_CASES, ids=[case.name for case in PLAYER_CASES])
def player_case(request: pytest.FixtureRequest) -> BindingsAccuracyCase:
    return cast(BindingsAccuracyCase, request.param)


@pytest.fixture
def bindings(player_case: BindingsAccuracyCase):
    return load_series_bindings(player_case.bindings_path)


@pytest.fixture
def graph(workbook: Path, bindings):
    return build_dependency_graph(workbook, bindings)


@pytest.fixture
def generated_module(graph, workbook: Path, bindings):
    return generate_bindings_namespace(graph, workbook, bindings)


def test_player_bindings_validate(player_case: BindingsAccuracyCase) -> None:
    assert_bindings_validate(player_case)


def test_player_shared_game_log_columns(
    graph,
    workbook: Path,
    bindings,
    player_case: BindingsAccuracyCase,
) -> None:
    slug = player_case.setter_name_prefix or player_case.name
    assert_shared_game_log_columns(
        graph,
        workbook,
        bindings,
        date_series_id=f"{slug}_game_date",
        result_series_id=f"{slug}_game_result",
    )


@pytest.mark.parametrize("check_index", range(2), ids=["reg_season_stats", "playoff_stats"])
def test_player_stat_matrix_accuracy(
    graph,
    workbook: Path,
    bindings,
    player_case: BindingsAccuracyCase,
    check_index: int,
) -> None:
    assert_series_spot_check(graph, workbook, bindings, player_case.series_checks[check_index])


def test_player_compute_functions_match_workbook(
    graph,
    workbook: Path,
    bindings,
    generated_module: dict[str, object],
) -> None:
    assert_all_compute_functions_match_workbook(graph, workbook, bindings, generated_module)


def test_player_setter_updates_downstream_compute(
    generated_module: dict[str, object],
    player_case: BindingsAccuracyCase,
) -> None:
    if player_case.downstream_update is None:
        pytest.skip("No downstream update configured")
    assert_downstream_update(generated_module, player_case.downstream_update)
