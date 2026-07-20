"""Unit tests for formula-group ProjectedAddress mapping and codegen helpers."""

from __future__ import annotations

from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.exporter.projection import ProjectedAddress
from tests.fixtures.formula_groups.option_b import (
    build_cross_sheet_union_option_b,
    build_row_stripe_cell_only_twin,
    build_row_stripe_option_b,
)


def test_map_address_to_projected_attaches_group_bindings() -> None:
    fx = build_row_stripe_option_b()
    with CodeGenerator(fx.graph) as gen:
        projected = gen._map_address_to_projected("Sheet1!E63")
    assert isinstance(projected, ProjectedAddress)
    assert projected.address == fx.group_key
    assert projected.parameters is not None
    assert projected.parameters["member"] == "Sheet1!E63"
    assert projected.parameters["bindings"] == ("Sheet1!E35",)


def test_codegen_emits_one_group_helper_and_member_wrappers() -> None:
    fx = build_row_stripe_option_b()
    with CodeGenerator(fx.graph) as gen:
        code = gen.generate(targets=list(fx.members))
    helper = gen._group_helper_name(fx.group_key)
    assert f"def {helper}(ctx, b0):" in code
    assert "Formula-group member wrappers" in code
    assert "def cell_sheet1_e63(ctx):" in code
    assert f"return {helper}(ctx, 'Sheet1!E35')" in code
    # One helper, not one specialized body per member
    assert code.count(f"def {helper}(") == 1


def test_codegen_group_member_matches_evaluator() -> None:
    fx = build_row_stripe_option_b()
    twin = build_row_stripe_cell_only_twin()
    with FormulaEvaluator(fx.graph) as ev, CodeGenerator(fx.graph) as gen:
        code = gen.generate(targets=["Sheet1!E63"])
        ns: dict[str, object] = {}
        exec(code, ns)
        make_context = ns["make_context"]
        compute_all = ns["compute_all"]
        assert callable(make_context) and callable(compute_all)
        ctx = make_context()
        exported = compute_all(ctx=ctx)
        assert isinstance(exported, dict)
        assert exported["Sheet1!E63"] == ev.evaluate("Sheet1!E63")
        with FormulaEvaluator(twin) as ev_t:
            assert exported["Sheet1!E63"] == ev_t.evaluate("Sheet1!E63")


def test_codegen_cross_sheet_member_wrapper() -> None:
    fx = build_cross_sheet_union_option_b()
    with CodeGenerator(fx.graph) as gen:
        code = gen.generate(targets=list(fx.members))
    assert "def cell_sheet2_b10(ctx):" in code
    helper = gen._group_helper_name(fx.group_key)
    assert helper in code
