"""Integration tests for formula patterns in formula_test_cases.xlsx.

Verifies that ``FormulaEvaluator`` and generated export code compute G3:G10 to
the workbook cached values (Excel reference).
Run:
    uv run pytest tests/integration/user_flows/formula_test.py -v
"""

from __future__ import annotations

import pytest

from excel_grapher import FormulaEvaluator, create_dependency_graph
from excel_grapher.grapher import DependencyGraph
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator
from tests.paths import TEST_SHEETS_FIXTURES

WORKBOOK_PATH = TEST_SHEETS_FIXTURES / "formula_test_cases.xlsx"
SHEET = "Sheet1"
FORMULA_TARGETS = [f"{SHEET}!G{row}" for row in range(3, 10)]
TARGETS = FORMULA_TARGETS + [f"{SHEET}!G10"]

EXPECTED: dict[str, object] = {
    f"{SHEET}!G3": 10,
    f"{SHEET}!G4": 15,
    f"{SHEET}!G5": "C13 is large",
    f"{SHEET}!G6": "value at C14 is 100",
    f"{SHEET}!G7": '50 says "C15"',
    f"{SHEET}!G8": "see Sheet1!C16 then 30",
    f"{SHEET}!G9": "C17 has data 5",
    f"{SHEET}!G10": None,
}


@pytest.fixture(scope="module")
def formula_graph() -> DependencyGraph:
    return create_dependency_graph(WORKBOOK_PATH, TARGETS, load_values=True)


def test_g3_g10_match_workbook_cached_values(formula_graph: DependencyGraph) -> None:
    """G3:G10 evaluate to the values saved in formula_test_cases.xlsx."""
    with FormulaEvaluator(formula_graph) as ev:
        for address, expected in EXPECTED.items():
            computed = ev._evaluate_cell(address)
            if isinstance(expected, (int, float)) and isinstance(computed, (int, float)):
                assert computed == pytest.approx(expected), address
            else:
                assert computed == expected, address


def test_g3_g10_codegen_matches_evaluator(formula_graph: DependencyGraph) -> None:
    """Generated export code agrees with FormulaEvaluator on formula cells G3:G9."""
    result = assert_codegen_matches_evaluator(formula_graph, FORMULA_TARGETS)
    for address, expected in EXPECTED.items():
        if address not in FORMULA_TARGETS:
            continue
        computed = result.generated_results[address]
        if isinstance(expected, (int, float)) and isinstance(computed, (int, float)):
            assert computed == pytest.approx(expected), address
        else:
            assert computed == expected, address
