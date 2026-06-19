"""Aggregate Stats bindings: parity tests and expected gaps (integration).

The ``aggregate_stats.bindings.yaml`` shard maps 325 scalar cells on the
``Aggregate Stats`` sheet. Several formula patterns are not fully supported yet;
tests marked ``xfail`` assert the desired behavior and should have the marker
removed once the underlying feature lands.

Known gaps exercised here:

- **Shard merge**: ``la_rams_2025`` vs ``player_game_log`` ``concept_scheme`` ids
  block merging the full ``ffv3_large.bindings/`` directory.
- **Dynamic INDEX/MATCH**: ``Aggregate Stats!D28`` needs cached or constraint-based
  dynamic-ref resolution; the generic ``build_dependency_graph`` helper does not
  enable ``use_cached_dynamic_refs``.
- **Whole-column references**: ``Aggregate Stats!E28:E32`` use ``'Sheet'!C:C`` and
  ``'Sheet'!A:A`` inside ``INDEX``/``MATCH``; the formula parser rejects them today.
- **Compute codegen**: exporting all 325 ``compute_*`` functions fails on the
  whole-column formulas above.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from fastpyxl import load_workbook

from excel_grapher.cli import main as cli_main
from excel_grapher.evaluator.errors import ParseError
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.evaluator.parser import parse
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.dynamic_refs import DynamicRefError
from excel_grapher.series_bindings import load_series_bindings
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.integration.user_flows.bindings_accuracy import (
    BindingsAccuracyCase,
    assert_bindings_validate,
    assert_compute_records_match_workbook,
    build_dependency_graph,
    compute_records,
    generate_bindings_namespace,
    read_workbook_cell,
    resolve_leaves,
)

EXAMPLES = Path(__file__).resolve().parents[3] / "examples" / "micro_workbooks"
WORKBOOK = EXAMPLES / "ffv3_large.xlsx"
BINDINGS_DIR = EXAMPLES / "ffv3_large.bindings"
AGGREGATE_BINDINGS = BINDINGS_DIR / "aggregate_stats.bindings.yaml"

BEST_WEEK_CELL = "Aggregate Stats!D28"
BEST_OPPONENT_CELL = "Aggregate Stats!E28"
BEST_OPPONENT_SERIES = "agg_perf_best_opp_stafford"
RANK_SERIES = "agg_rank_stafford_w1"


@pytest.fixture(scope="module")
def workbook() -> Path:
    if not WORKBOOK.is_file():
        pytest.skip(f"Workbook fixture missing: {WORKBOOK}")
    return WORKBOOK


@pytest.fixture(scope="module")
def aggregate_bindings(workbook: Path):
    if not AGGREGATE_BINDINGS.is_file():
        pytest.skip(f"Bindings fixture missing: {AGGREGATE_BINDINGS}")
    return load_series_bindings(AGGREGATE_BINDINGS)


@pytest.fixture(scope="module")
def aggregate_case(workbook: Path) -> BindingsAccuracyCase:
    return BindingsAccuracyCase(
        name="aggregate_stats",
        workbook=workbook,
        bindings_path=AGGREGATE_BINDINGS,
        expected_compute_count=325,
    )


@pytest.fixture(scope="module")
def aggregate_graph(workbook: Path, aggregate_bindings):
    targets = all_series_targets(aggregate_bindings, workbook=workbook)
    return create_dependency_graph(
        workbook,
        targets,
        load_values=True,
        use_cached_dynamic_refs=True,
    )


def test_aggregate_stats_bindings_shard_validates(aggregate_case: BindingsAccuracyCase) -> None:
    """The aggregate shard alone validates (uses cached dynamic refs internally)."""
    assert_bindings_validate(aggregate_case)


def test_aggregate_stats_rank_series_resolves_to_formula_cell(
    workbook: Path,
    aggregate_bindings,
    aggregate_graph,
) -> None:
    """Rank bindings resolve to the RANK formula cell on Aggregate Stats."""
    leaves = resolve_leaves(
        aggregate_graph,
        workbook,
        aggregate_bindings,
        RANK_SERIES,
        direction="output",
    )
    assert len(leaves) == 1
    assert leaves[0]["address"] == "'Aggregate Stats'!C12"
    assert leaves[0]["key"] == {"TIME_PERIOD": "W1"}

    with FormulaEvaluator(aggregate_graph) as ev:
        assert ev.evaluate("'Aggregate Stats'!C12") == read_workbook_cell(
            workbook, "Aggregate Stats!C12"
        )


@pytest.mark.xfail(
    strict=True,
    reason="Merge ffv3_large.bindings/ once aggregate and player shards share one concept_scheme.",
)
def test_ffv3_large_bindings_directory_merges(workbook: Path) -> None:
    """All ffv3_large binding shards should validate together."""
    exit_code = cli_main(
        [
            "bindings",
            "validate",
            str(workbook),
            "--bindings",
            "ffv3_large.bindings",
        ]
    )
    assert exit_code == 0


@pytest.mark.xfail(
    raises=DynamicRefError,
    strict=True,
    reason=(
        "Aggregate Stats best-week formulas use dynamic INDEX/MATCH; remove xfail when "
        "build_dependency_graph supports constraint-based dynamic refs by default."
    ),
)
def test_aggregate_stats_graph_builds_without_cached_dynamic_refs(
    workbook: Path,
    aggregate_bindings,
) -> None:
    """Dependency closure should build without cached dynamic-ref fallback."""
    build_dependency_graph(workbook, aggregate_bindings)


def test_aggregate_stats_best_week_evaluator_matches_excel(
    workbook: Path,
    aggregate_graph,
) -> None:
    """Dynamic INDEX/MATCH at D28 evaluates to the Excel cached week label."""
    with FormulaEvaluator(aggregate_graph) as ev:
        assert ev.evaluate(BEST_WEEK_CELL) == read_workbook_cell(workbook, BEST_WEEK_CELL)


@pytest.mark.xfail(
    raises=ParseError,
    strict=True,
    reason="Whole-column refs like 'QB - Stafford'!C:C are not parsed yet.",
)
def test_aggregate_stats_whole_column_formula_parses(workbook: Path) -> None:
    """Best-opponent formulas using !C:C and !A:A should parse."""
    wb = load_workbook(workbook, data_only=False, read_only=True)
    try:
        formula = wb["Aggregate Stats"]["E28"].value
    finally:
        wb.close()
    assert isinstance(formula, str) and formula.startswith("=")
    parse(formula)


def test_aggregate_stats_best_opponent_evaluator_matches_excel(
    workbook: Path,
    aggregate_graph,
) -> None:
    """Whole-column INDEX/MATCH at E28 should match the Excel cached opponent."""
    with FormulaEvaluator(aggregate_graph) as ev:
        assert ev.evaluate(BEST_OPPONENT_CELL) == read_workbook_cell(workbook, BEST_OPPONENT_CELL)


@pytest.mark.xfail(
    strict=True,
    reason="Codegen for agg_perf_best_opp_* fails until whole-column refs parse.",
)
def test_aggregate_stats_best_opponent_compute_matches_workbook(
    workbook: Path,
    aggregate_bindings,
    aggregate_graph,
) -> None:
    """Generated compute_agg_perf_best_opp_stafford should match the workbook."""
    namespace = generate_bindings_namespace(aggregate_graph, workbook, aggregate_bindings)
    records = compute_records(namespace, "compute_agg_perf_best_opp_stafford")
    assert_compute_records_match_workbook(
        aggregate_graph,
        workbook,
        aggregate_bindings,
        BEST_OPPONENT_SERIES,
        records,
    )


@pytest.mark.xfail(
    strict=True,
    reason="Exporting all 325 aggregate compute functions blocks on whole-column refs.",
)
def test_aggregate_stats_all_compute_functions_match_workbook(
    workbook: Path,
    aggregate_bindings,
    aggregate_graph,
) -> None:
    """Every declared aggregate compute_* function should match the workbook cache."""
    namespace = generate_bindings_namespace(aggregate_graph, workbook, aggregate_bindings)
    for compute_name in sorted(name for name in namespace if str(name).startswith("compute_agg_")):
        series_id = str(compute_name).removeprefix("compute_")
        records = compute_records(namespace, str(compute_name))
        assert_compute_records_match_workbook(
            aggregate_graph,
            workbook,
            aggregate_bindings,
            series_id,
            records,
        )
