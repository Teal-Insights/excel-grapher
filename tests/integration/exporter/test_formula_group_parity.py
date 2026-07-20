"""Sprint 5: formula-group parity (evaluator ↔ export ↔ cell-only twin)."""

from __future__ import annotations

from typing import Any, cast

import pytest

from excel_grapher.evaluator.errors import FormulaGroupKeyError
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.evaluator.types import XlError
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher.node import NodeKind, locate_cell
from tests.fixtures.formula_groups.hand_built import (
    assert_formula_group_occupancy,
    build_cross_sheet_cell_only_twin,
    build_cross_sheet_union_group,
    build_div_zero_cell_only_twin,
    build_div_zero_group,
    build_row_stripe_cell_only_twin,
    build_row_stripe_group,
)
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator


def test_row_stripe_member_eval_matches_twin() -> None:
    fx = build_row_stripe_group()
    twin = build_row_stripe_cell_only_twin()
    group = fx.graph.get_node(fx.group_key)
    assert group is not None
    assert_formula_group_occupancy(group)
    with FormulaEvaluator(fx.graph) as ev_g, FormulaEvaluator(twin) as ev_t:
        for member in fx.members:
            assert fx.graph.get_node(member) is None
            loc = locate_cell(fx.graph, member)
            assert loc is not None and loc.kind is not NodeKind.cell
            assert ev_g.evaluate(member) == ev_t.evaluate(member)


def test_row_stripe_codegen_matches_evaluator() -> None:
    fx = build_row_stripe_group()
    result = assert_codegen_matches_evaluator(fx.graph, list(fx.members))
    assert result.evaluator_results == {
        "Sheet1!D63": 10.0,
        "Sheet1!E63": 11.0,
        "Sheet1!F63": 12.0,
    }
    assert "_group_" in result.generated_code
    assert result.generated_code.count("def _group_") == 1


def test_cross_sheet_codegen_matches_evaluator_and_twin() -> None:
    fx = build_cross_sheet_union_group()
    twin = build_cross_sheet_cell_only_twin()
    result = assert_codegen_matches_evaluator(fx.graph, list(fx.members))
    with FormulaEvaluator(twin) as ev_t:
        for member in fx.members:
            assert result.evaluator_results[member] == ev_t.evaluate(member)
            assert result.generated_results[member] == result.evaluator_results[member]


def test_div_zero_error_codes_match_across_channels() -> None:
    """Evaluator XlError sentinels match export XlErrorException codes."""
    fx = build_div_zero_group()
    twin = build_div_zero_cell_only_twin()
    with FormulaEvaluator(fx.graph) as ev_g, FormulaEvaluator(twin) as ev_t:
        assert ev_g.evaluate("Sheet1!A1") == ev_t.evaluate("Sheet1!A1") == XlError.DIV
        assert ev_g.evaluate("Sheet1!B1") == ev_t.evaluate("Sheet1!B1") == 0.5

    result = assert_codegen_matches_evaluator(fx.graph, list(fx.members))
    assert result.evaluator_results["Sheet1!A1"] == XlError.DIV
    assert result.generated_results["Sheet1!A1"] == XlError.DIV
    assert result.generated_results["Sheet1!B1"] == 0.5

    code = result.generated_code
    ns: dict[str, object] = {}
    exec(code, ns)
    ctx = ns["make_context"]()  # type: ignore[operator]
    xl_error_exception = ns["XlErrorException"]
    assert isinstance(xl_error_exception, type)
    with pytest.raises(xl_error_exception) as exc_info:
        ns["cell_sheet1_a1"](ctx)  # type: ignore[operator]
    assert cast(Any, exc_info.value).code == XlError.DIV


def test_codegen_rejects_formula_group_key_target() -> None:
    fx = build_row_stripe_group()
    with pytest.raises(FormulaGroupKeyError, match="multi-cell group key"):
        CodeGenerator(fx.graph).generate(targets=[fx.group_key])

    union = build_cross_sheet_union_group()
    with pytest.raises(FormulaGroupKeyError, match="multi-cell group key"):
        CodeGenerator(union.graph).generate(targets=[union.group_key])


def test_cell_only_twin_codegen_has_no_group_helpers() -> None:
    twin = build_row_stripe_cell_only_twin()
    result = assert_codegen_matches_evaluator(twin, ["Sheet1!D63", "Sheet1!E63", "Sheet1!F63"])
    assert "def _group_" not in result.generated_code
    assert "Formula-group member wrappers" not in result.generated_code
