"""Evaluator helpers for function tests.

Address-keyed `CodeGenerator.generate` was removed (#764). Evaluator ↔ export
parity for packages uses inverted-tree `generate_modules` with series bindings.
`FormulaEvaluator` remains the in-process Excel engine for function tests.
"""

from __future__ import annotations

from typing import cast

from excel_grapher import DependencyGraph, FormulaEvaluator


def evaluate_targets(
    graph: DependencyGraph,
    targets: list[str],
    *,
    blank_ranges: tuple[str, ...] | list[str] | None = None,
) -> dict[str, object]:
    """Evaluate `targets` with `FormulaEvaluator` and return the result map."""
    with FormulaEvaluator(graph, blank_ranges=blank_ranges) as ev:
        return cast(dict[str, object], ev.evaluate(targets))
