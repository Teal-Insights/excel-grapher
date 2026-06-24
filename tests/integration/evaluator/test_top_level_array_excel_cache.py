"""Excel cached-value checks for top-level array formulas (#284, Sprint 4)."""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pytest

from excel_grapher import FormulaEvaluator, create_dependency_graph
from excel_grapher.core.types import XlError
from tests.fixtures.array_results.workbook import (
    blocked_spill_path,
    build_blocked_spill_workbook,
    column_compare_path,
    ensure_committed_fixtures,
)
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator

ensure_committed_fixtures()

BLOCKED_SPILL_TARGET = "Data!D10"
BLOCKED_SPILL_SLOT = "Data!D11"


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


@pytest.mark.xfail(
    strict=True,
    reason=(
        "Issue #284: graph extraction does not include spill-footprint slots in the "
        "anchor closure, so occupied neighbors omitted from the graph are invisible "
        "to #SPILL! blocking even when Excel cached the anchor as #SPILL!."
    ),
)
def test_blocked_spill_excel_cache_parity_when_footprint_not_in_closure() -> None:
    """Excel cached ``#SPILL!`` at the anchor when only the anchor is a graph target.

    ``Data!D10 = C5:C7="Software"`` spills into ``Data!D11``, which holds ``1`` in
    the fixture workbook. Excel records ``#SPILL!`` at ``Data!D10``. Extraction from
    the anchor alone does not pull ``Data!D11`` into the graph (nothing reads it),
    so spill blocking never sees the obstructing cell.
    """
    graph = create_dependency_graph(
        blocked_spill_path(),
        [BLOCKED_SPILL_TARGET],
        load_values=True,
    )
    assert graph.get_node(BLOCKED_SPILL_SLOT) is None

    anchor = graph.get_node(BLOCKED_SPILL_TARGET)
    assert anchor is not None
    assert anchor.value in (XlError.SPILL, XlError.SPILL.value, "#SPILL!")

    with FormulaEvaluator(graph) as evaluator:
        result = evaluator.evaluate(BLOCKED_SPILL_TARGET)

    assert result == XlError.SPILL
    assert not isinstance(result, np.ndarray)


@pytest.mark.xfail(
    strict=True,
    reason=(
        "Issue #284: graph extraction does not link spill slots to their anchor "
        "formulas, so reading an empty leaf spill cell returns the stored leaf "
        "value instead of projecting from the anchor array as Excel does."
    ),
)
def test_spill_slot_excel_parity_when_anchor_not_in_closure() -> None:
    """Excel spill projection at an empty slot when only the slot is a graph target.

    ``Data!D10 = C5:C7="Software"`` spills ``False`` into ``Data!D11``. Excel
    shows that value even though ``Data!D11`` is not a formula. Extraction from
    ``Data!D11`` alone keeps it as an empty leaf with no dependency on
    ``Data!D10``, so ``evaluate`` returns the leaf value rather than the spill.
    """
    graph = create_dependency_graph(
        column_compare_path(),
        [BLOCKED_SPILL_SLOT],
        load_values=True,
    )
    assert graph.get_node(BLOCKED_SPILL_TARGET) is None

    slot = graph.get_node(BLOCKED_SPILL_SLOT)
    assert slot is not None
    assert slot.is_leaf
    assert slot.value is None
    assert list(graph.get_dependencies(BLOCKED_SPILL_SLOT)) == []

    with FormulaEvaluator(graph) as evaluator:
        result = evaluator.evaluate(BLOCKED_SPILL_SLOT)

    assert result is False
