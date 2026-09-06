"""Operation budgets for large exact MATCH/VLOOKUP on FormulaEvaluator.

The number of evaluated cells catches eager-range regressions without flaky
wall-clock thresholds that vary with host load.
"""

from __future__ import annotations

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.evaluator import FormulaEvaluator

LARGE_LOOKUP_ROWS = 10_000


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


def test_large_exact_match_last_key_evaluates_each_lookup_cell_once() -> None:
    """A last-key exact MATCH evaluates the lookup column and target once."""
    graph = _lookup_table_graph(
        rows=LARGE_LOOKUP_ROWS,
        formula=f'=MATCH("k{LARGE_LOOKUP_ROWS}", S!A1:S!A{LARGE_LOOKUP_ROWS}, 0)',
    )
    evaluated: list[str] = []
    with FormulaEvaluator(
        graph, on_cell_evaluated=lambda address, _value: evaluated.append(address)
    ) as ev:
        result = ev.evaluate(["S!Z1"])
    assert result == {"S!Z1": LARGE_LOOKUP_ROWS}
    assert len(evaluated) == LARGE_LOOKUP_ROWS + 1
    assert set(evaluated) == {"S!Z1", *(f"S!A{row}" for row in range(1, LARGE_LOOKUP_ROWS + 1))}


def test_large_exact_vlookup_last_key_evaluates_lookup_column_and_result() -> None:
    """A last-key exact VLOOKUP scans column A and reads one result cell."""
    graph = _lookup_table_graph(
        rows=LARGE_LOOKUP_ROWS,
        formula=(f'=VLOOKUP("k{LARGE_LOOKUP_ROWS}", S!A1:S!B{LARGE_LOOKUP_ROWS}, 2, FALSE)'),
    )
    evaluated: list[str] = []
    with FormulaEvaluator(
        graph, on_cell_evaluated=lambda address, _value: evaluated.append(address)
    ) as ev:
        result = ev.evaluate(["S!Z1"])
    assert result == {"S!Z1": LARGE_LOOKUP_ROWS}
    assert len(evaluated) == LARGE_LOOKUP_ROWS + 2
    assert set(evaluated) == {
        "S!Z1",
        f"S!B{LARGE_LOOKUP_ROWS}",
        *(f"S!A{row}" for row in range(1, LARGE_LOOKUP_ROWS + 1)),
    }


def test_large_exact_match_early_hit_evaluates_only_first_lookup_cell() -> None:
    """Early exact MATCH must not evaluate trailing formula rows."""
    rows = LARGE_LOOKUP_ROWS
    graph = _lookup_table_graph(
        rows=rows,
        formula=f'=MATCH("k1", S!A1:S!A{rows}, 0)',
        fill_trailing_formulas=True,
    )
    evaluated: list[str] = []
    with FormulaEvaluator(
        graph, on_cell_evaluated=lambda address, _value: evaluated.append(address)
    ) as ev:
        result = ev.evaluate(["S!Z1"])
        assert result == {"S!Z1": 1}
        assert f"S!A{rows}" not in ev._cache
        assert f"S!B{rows}" not in ev._cache
    assert evaluated == ["S!A1", "S!Z1"]


def test_large_exact_vlookup_early_hit_evaluates_only_first_lookup_row() -> None:
    """Early exact VLOOKUP must not evaluate trailing formula rows."""
    rows = LARGE_LOOKUP_ROWS
    graph = _lookup_table_graph(
        rows=rows,
        formula=f'=VLOOKUP("k1", S!A1:S!B{rows}, 2, FALSE)',
        fill_trailing_formulas=True,
    )
    evaluated: list[str] = []
    with FormulaEvaluator(
        graph, on_cell_evaluated=lambda address, _value: evaluated.append(address)
    ) as ev:
        result = ev.evaluate(["S!Z1"])
        assert result == {"S!Z1": 1}
        assert f"S!A{rows}" not in ev._cache
        assert f"S!B{rows}" not in ev._cache
    assert evaluated == ["S!A1", "S!B1", "S!Z1"]
