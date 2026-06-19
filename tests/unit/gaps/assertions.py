"""Shared assertions for gap reproduction tests."""

from __future__ import annotations

from excel_grapher import FormulaEvaluator
from excel_grapher.grapher import DependencyGraph
from tests.integration.utils.parity_harness import exec_generated_code_with_cache


def assert_evaluator_and_codegen_disagree(graph: DependencyGraph, address: str) -> None:
    """Assert standalone codegen cache differs from ``FormulaEvaluator`` for one cell."""
    with FormulaEvaluator(graph) as evaluator:
        eval_value = evaluator.evaluate(address)
    generated_cache, _, _ = exec_generated_code_with_cache(graph, [address])
    generated_value = generated_cache[address]
    assert eval_value != generated_value, (
        f"expected evaluator/codegen gap at {address}, both returned {eval_value!r}"
    )
