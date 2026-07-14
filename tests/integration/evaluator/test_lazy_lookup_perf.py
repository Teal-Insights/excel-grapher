"""Loose time budgets for large exact MATCH/VLOOKUP on FormulaEvaluator.

These ceilings catch catastrophic regressions (e.g. re-eagerizing whole
ranges on the lookup path). They are not micro-benchmarks.
"""

from __future__ import annotations

import time

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.evaluator import FormulaEvaluator

LARGE_LOOKUP_ROWS = 10_000
# Calibrated ~0.1s for a full 10K leaf-key scan on CI-class hardware; 2s leaves
# comfortable headroom without masking multi-second regressions.
LARGE_LOOKUP_FULL_SCAN_BUDGET_SEC = 2.0
# Early exact hits must remain selective even when the rectangle has many
# trailing formula cells that would be expensive if eagerly resolved.
LARGE_LOOKUP_EARLY_HIT_BUDGET_SEC = 1.0


def _make_node(address: str, formula: str | None, value: object) -> Node:
    sheet, coord = parse_address(address)
    col = "".join(c for c in coord if c.isalpha())
    row = int("".join(c for c in coord if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=value,
        is_leaf=formula is None,
    )


def _lookup_table_graph(
    *,
    rows: int,
    formula: str,
    fill_trailing_formulas: bool = False,
) -> DependencyGraph:
    graph = DependencyGraph()
    for i in range(1, rows + 1):
        if fill_trailing_formulas and i > 1:
            # Cheap but non-trivial formula work if ever evaluated.
            graph.add_node(_make_node(f"S!A{i}", f"=S!A1+{i}", None))
            graph.add_node(_make_node(f"S!B{i}", f"=S!A{i}*2", None))
        else:
            graph.add_node(_make_node(f"S!A{i}", None, f"k{i}"))
            graph.add_node(_make_node(f"S!B{i}", None, i))
    graph.add_node(_make_node("S!Z1", formula, None))
    return graph


def test_large_exact_match_last_key_under_time_budget() -> None:
    """10K-row exact MATCH (last key) via FormulaEvaluator stays within budget."""
    graph = _lookup_table_graph(
        rows=LARGE_LOOKUP_ROWS,
        formula=f'=MATCH("k{LARGE_LOOKUP_ROWS}", S!A1:S!A{LARGE_LOOKUP_ROWS}, 0)',
    )
    started = time.perf_counter()
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!Z1"])
    elapsed = time.perf_counter() - started
    assert result == {"S!Z1": LARGE_LOOKUP_ROWS}
    assert elapsed < LARGE_LOOKUP_FULL_SCAN_BUDGET_SEC, (
        f"10K exact MATCH (last key) took {elapsed:.2f}s, "
        f"expected < {LARGE_LOOKUP_FULL_SCAN_BUDGET_SEC}s"
    )


def test_large_exact_vlookup_last_key_under_time_budget() -> None:
    """10K-row exact VLOOKUP (last key) via FormulaEvaluator stays within budget."""
    graph = _lookup_table_graph(
        rows=LARGE_LOOKUP_ROWS,
        formula=(f'=VLOOKUP("k{LARGE_LOOKUP_ROWS}", S!A1:S!B{LARGE_LOOKUP_ROWS}, 2, FALSE)'),
    )
    started = time.perf_counter()
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!Z1"])
    elapsed = time.perf_counter() - started
    assert result == {"S!Z1": LARGE_LOOKUP_ROWS}
    assert elapsed < LARGE_LOOKUP_FULL_SCAN_BUDGET_SEC, (
        f"10K exact VLOOKUP (last key) took {elapsed:.2f}s, "
        f"expected < {LARGE_LOOKUP_FULL_SCAN_BUDGET_SEC}s"
    )


def test_large_exact_match_early_hit_skips_trailing_formulas_under_budget() -> None:
    """Early exact MATCH must not evaluate trailing formula rows."""
    rows = LARGE_LOOKUP_ROWS
    graph = _lookup_table_graph(
        rows=rows,
        formula=f'=MATCH("k1", S!A1:S!A{rows}, 0)',
        fill_trailing_formulas=True,
    )
    started = time.perf_counter()
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!Z1"])
        assert result == {"S!Z1": 1}
        assert f"S!A{rows}" not in ev._cache
        assert f"S!B{rows}" not in ev._cache
    elapsed = time.perf_counter() - started
    assert elapsed < LARGE_LOOKUP_EARLY_HIT_BUDGET_SEC, (
        f"10K exact MATCH (early hit) took {elapsed:.2f}s, "
        f"expected < {LARGE_LOOKUP_EARLY_HIT_BUDGET_SEC}s"
    )


def test_large_exact_vlookup_early_hit_skips_trailing_formulas_under_budget() -> None:
    """Early exact VLOOKUP must not evaluate trailing formula rows."""
    rows = LARGE_LOOKUP_ROWS
    graph = _lookup_table_graph(
        rows=rows,
        formula=f'=VLOOKUP("k1", S!A1:S!B{rows}, 2, FALSE)',
        fill_trailing_formulas=True,
    )
    started = time.perf_counter()
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!Z1"])
        assert result == {"S!Z1": 1}
        assert f"S!A{rows}" not in ev._cache
        assert f"S!B{rows}" not in ev._cache
    elapsed = time.perf_counter() - started
    assert elapsed < LARGE_LOOKUP_EARLY_HIT_BUDGET_SEC, (
        f"10K exact VLOOKUP (early hit) took {elapsed:.2f}s, "
        f"expected < {LARGE_LOOKUP_EARLY_HIT_BUDGET_SEC}s"
    )
