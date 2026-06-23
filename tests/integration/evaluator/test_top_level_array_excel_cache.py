"""Excel cached-value checks for top-level array formulas (#284, Sprint 4)."""

from __future__ import annotations

from pathlib import Path

from excel_grapher import FormulaEvaluator, create_dependency_graph
from excel_grapher.core.types import XlError
from tests.fixtures.array_results.workbook import (
    blocked_spill_path,
    build_blocked_spill_workbook,
    ensure_committed_fixtures,
)
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator

ensure_committed_fixtures()

BLOCKED_SPILL_TARGET = "Data!D10"


def test_blocked_spill_evaluator_returns_spill_error() -> None:
    """Occupied spill slot yields ``#SPILL!`` at the anchor formula cell."""
    graph = create_dependency_graph(
        blocked_spill_path(),
        [BLOCKED_SPILL_TARGET, "Data!D11"],
        load_values=True,
    )
    with FormulaEvaluator(graph) as evaluator:
        result = evaluator.evaluate(BLOCKED_SPILL_TARGET)
    assert result == XlError.SPILL


def test_blocked_spill_eval_codegen_parity() -> None:
    """Export agrees with evaluator on ``#SPILL!`` for blocked spill footprints."""
    graph = create_dependency_graph(
        blocked_spill_path(),
        [BLOCKED_SPILL_TARGET, "Data!D11"],
        load_values=True,
    )
    result = assert_codegen_matches_evaluator(graph, [BLOCKED_SPILL_TARGET])
    assert result.evaluator_results[BLOCKED_SPILL_TARGET] == XlError.SPILL
    assert result.generated_results[BLOCKED_SPILL_TARGET] == XlError.SPILL


def test_blocked_spill_workbook_embeds_cached_error(tmp_path: Path) -> None:
    """Fixture builder records Excel-style ``#SPILL!`` as the cached anchor value."""
    path = build_blocked_spill_workbook(tmp_path / "blocked.xlsx")
    graph = create_dependency_graph(path, [BLOCKED_SPILL_TARGET, "Data!D11"], load_values=True)
    node = graph.get_node(BLOCKED_SPILL_TARGET)
    assert node is not None
    assert node.value in (XlError.SPILL, XlError.SPILL.value, "#SPILL!")
