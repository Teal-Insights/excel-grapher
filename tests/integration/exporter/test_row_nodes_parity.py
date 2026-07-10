"""Sprint 5: Option B evaluator ↔ codegen parity and hardening."""

from __future__ import annotations

from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.evaluator.types import XlError
from excel_grapher.exporter.codegen import CodeGenerator
from tests.fixtures.row_nodes.option_b_stripe import (
    assert_unique_occupancy_for_row,
    build_cell_only_regression_graph,
    build_option_b_div_error_graph,
    build_option_b_quoted_sheet_graph,
    build_option_b_stripe_fixture,
)
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator


def test_option_b_codegen_matches_evaluator_values() -> None:
    fixture = build_option_b_stripe_fixture()
    assert_unique_occupancy_for_row(fixture.option_b, fixture.row_key)
    result = assert_codegen_matches_evaluator(fixture.option_b, list(fixture.member_keys))
    assert result.evaluator_results["Sheet1!D63"] == 6
    assert result.evaluator_results["Sheet1!E63"] == 10
    assert "_row_sheet1_d63_e63" in result.generated_code


def test_option_b_matches_cell_only_twin_via_parity() -> None:
    fixture = build_option_b_stripe_fixture()
    row_result = assert_codegen_matches_evaluator(fixture.option_b, list(fixture.member_keys))
    twin_result = assert_codegen_matches_evaluator(fixture.cell_only, list(fixture.member_keys))
    for member in fixture.member_keys:
        assert row_result.evaluator_results[member] == twin_result.evaluator_results[member]


def test_option_b_error_code_parity_div0() -> None:
    g = build_option_b_div_error_graph()
    result = assert_codegen_matches_evaluator(g, ["Sheet1!D63", "Sheet1!E63"])
    assert result.evaluator_results["Sheet1!D63"] == XlError.DIV
    assert result.evaluator_results["Sheet1!E63"] == XlError.DIV
    assert result.generated_results["Sheet1!D63"] == XlError.DIV
    assert result.generated_results["Sheet1!E63"] == XlError.DIV


def test_option_b_quoted_sheet_and_static_range_parity() -> None:
    g = build_option_b_quoted_sheet_graph()
    members = ["'My Sheet'!D2", "'My Sheet'!E2"]
    result = assert_codegen_matches_evaluator(g, members)
    # D2: 3 + SUM(10,20) = 33; E2: 4 + SUM(10,20) = 34
    assert result.evaluator_results["'My Sheet'!D2"] == 33
    assert result.evaluator_results["'My Sheet'!E2"] == 34
    assert "SUM" in result.generated_code.upper() or "xl_sum" in result.generated_code
    assert "f\"'My Sheet'!{column}1\"" in result.generated_code or (
        "f\"'My Sheet'!{column}1\"" in result.generated_code
    )


def test_cell_only_graph_codegen_unaffected() -> None:
    g = build_cell_only_regression_graph()
    result = assert_codegen_matches_evaluator(g, ["Sheet1!B1"])
    assert result.evaluator_results["Sheet1!B1"] == 6
    assert "_row_" not in result.generated_code
    assert "def cell_sheet1_b1(ctx):" in result.generated_code


def test_option_b_eval_lazy_and_rejects_row_key() -> None:
    """Keep Sprint 3 success criteria covered alongside parity."""
    import pytest

    fixture = build_option_b_stripe_fixture()
    with FormulaEvaluator(fixture.option_b) as ev:
        assert ev.evaluate("Sheet1!E63") == 10
        assert "Sheet1!D63" not in ev._cache
        with pytest.raises(ValueError, match="member"):
            ev.evaluate(fixture.row_key)


def test_codegen_without_edges_still_exports_sibling_varying_leaves() -> None:
    """Varying-slot leaves for non-anchor columns are exported even without edges."""
    from excel_grapher.grapher.graph import DependencyGraph
    from excel_grapher.grapher.node import Node, make_row_node

    g = DependencyGraph()
    g.add_node(Node("Sheet1", "D", 35, None, None, 3, True))
    g.add_node(Node("Sheet1", "E", 35, None, None, 5, True))
    g.add_node(
        make_row_node(
            "Sheet1",
            63,
            "D",
            "E",
            formula="=Sheet1!D35*2",
            normalized_formula="=Sheet1!D35*2",
            varying_ref_slots=(0,),
            is_leaf=False,
        )
    )
    # No edges — AST only sees D35; expansion must still pull E35 into inputs.
    code = CodeGenerator(g).generate(["Sheet1!E63"])
    assert "'Sheet1!E35'" in code or '"Sheet1!E35"' in code
    result = assert_codegen_matches_evaluator(g, ["Sheet1!E63"])
    assert result.evaluator_results["Sheet1!E63"] == 10
