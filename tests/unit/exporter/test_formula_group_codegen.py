"""Unit tests for formula-group ProjectedAddress mapping and codegen helpers."""

from __future__ import annotations

import re

from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.exporter.projection import ProjectedAddress
from tests.fixtures.formula_groups.hand_built import (
    build_column_self_group,
    build_cross_sheet_union_group,
    build_row_self_group,
    build_row_stripe_cell_only_twin,
    build_row_stripe_group,
    build_sum_over_constant_group,
)


def test_map_address_to_projected_attaches_group_bindings() -> None:
    fx = build_row_stripe_group()
    with CodeGenerator(fx.graph) as gen:
        projected = gen._map_address_to_projected("Sheet1!E63")
    assert isinstance(projected, ProjectedAddress)
    assert projected.address == fx.group_key
    assert projected.parameters is not None
    assert projected.parameters["member"] == "Sheet1!E63"
    assert projected.parameters["bindings"] == ("Sheet1!E35",)


def test_codegen_emits_one_group_helper_and_member_wrappers() -> None:
    fx = build_row_stripe_group()
    with CodeGenerator(fx.graph) as gen:
        code = gen.generate(targets=list(fx.members))
        helper = gen._group_helper_name(fx.group_key)
    assert helper.startswith("_group_")
    assert not re.fullmatch(r"_group_[0-9a-f]{12}", helper)
    assert f"def {helper}(ctx, member, b0):" in code
    assert "Formula-group member wrappers" in code
    assert "def cell_sheet1_e63(ctx):" in code
    assert f"return {helper}(ctx, 'Sheet1!E63', 'Sheet1!E35')" in code
    # One helper, not one specialized body per member
    assert code.count(f"def {helper}(") == 1


def test_codegen_hashed_group_helper_names_are_short_and_deterministic() -> None:
    fx = build_row_stripe_group()
    with CodeGenerator(fx.graph, hash_group_helper_names=True) as gen:
        code = gen.generate(targets=list(fx.members))
        helper = gen._group_helper_name(fx.group_key)
    assert re.fullmatch(r"_group_[0-9a-f]{12}", helper)
    assert len(helper) < 32
    with CodeGenerator(fx.graph, hash_group_helper_names=True) as gen2:
        assert gen2._group_helper_name(fx.group_key) == helper
    assert f"def {helper}(ctx, member, b0):" in code
    assert f"return {helper}(ctx, 'Sheet1!E63', 'Sheet1!E35')" in code
    assert code.count(f"def {helper}(") == 1


def test_codegen_group_member_matches_evaluator() -> None:
    fx = build_row_stripe_group()
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
    fx = build_cross_sheet_union_group()
    with CodeGenerator(fx.graph) as gen:
        code = gen.generate(targets=list(fx.members))
    assert "def cell_sheet2_b10(ctx):" in code
    helper = gen._group_helper_name(fx.group_key)
    assert helper in code


def test_codegen_group_no_arg_row_matches_evaluator() -> None:
    """Bare ROW() in a shared group helper must use the member address, not the group key."""
    from excel_grapher.evaluator.name_utils import address_to_python_name

    fx = build_row_self_group()
    with FormulaEvaluator(fx.graph) as ev:
        expected = {m: ev.evaluate(m) for m in fx.members}
    assert expected == {
        "Sheet1!B10": 10,
        "Sheet1!B11": 11,
        "Sheet1!B12": 12,
    }

    with CodeGenerator(fx.graph) as gen:
        code = gen.generate(targets=list(fx.members))
    helper = gen._group_helper_name(fx.group_key)
    assert f"def {helper}(ctx, member):" in code
    assert "xl_formula_row(member)" in code
    assert f"return {helper}(ctx, 'Sheet1!B11')" in code

    ns: dict[str, object] = {}
    exec(code, ns)
    make_context = ns["make_context"]
    assert callable(make_context)
    ctx = make_context()
    for member, want in expected.items():
        wrapper = ns[address_to_python_name(member)]
        assert callable(wrapper)
        assert wrapper(ctx) == want


def test_codegen_group_no_arg_column_matches_evaluator() -> None:
    from excel_grapher.evaluator.name_utils import address_to_python_name

    fx = build_column_self_group()
    with FormulaEvaluator(fx.graph) as ev:
        expected = {m: ev.evaluate(m) for m in fx.members}
    assert expected == {
        "Sheet1!D5": 4,
        "Sheet1!E5": 5,
        "Sheet1!F5": 6,
    }

    with CodeGenerator(fx.graph) as gen:
        code = gen.generate(targets=list(fx.members))
    helper = gen._group_helper_name(fx.group_key)
    assert f"def {helper}(ctx, member):" in code
    assert "xl_formula_column(member)" in code

    ns: dict[str, object] = {}
    exec(code, ns)
    make_context = ns["make_context"]
    assert callable(make_context)
    ctx = make_context()
    for member, want in expected.items():
        wrapper = ns[address_to_python_name(member)]
        assert callable(wrapper)
        assert wrapper(ctx) == want


def test_codegen_emits_wrappers_for_non_target_group_members() -> None:
    """LIC-DSF: range walk needs wrappers for members that are not generate targets."""
    fx = build_sum_over_constant_group()
    assert fx.dependent == "Sheet1!B1"
    with CodeGenerator(fx.graph) as gen:
        code = gen.generate(targets=[fx.dependent])
    assert "def cell_sheet1_b1(ctx):" in code
    assert "def cell_sheet1_a1(ctx):" in code
    assert "def cell_sheet1_a2(ctx):" in code


def test_codegen_sum_over_non_target_group_members_matches_evaluator() -> None:
    """Exported SUM over a coalesced group must match evaluator (LIC-DSF KeyError repro)."""
    fx = build_sum_over_constant_group()
    assert fx.dependent is not None
    with FormulaEvaluator(fx.graph) as ev:
        expected = ev.evaluate(fx.dependent)
    assert expected == 20.0

    with CodeGenerator(fx.graph) as gen:
        code = gen.generate(targets=[fx.dependent])
    ns: dict[str, object] = {}
    exec(code, ns)
    make_context = ns["make_context"]
    compute_all = ns["compute_all"]
    assert callable(make_context) and callable(compute_all)
    exported = compute_all(ctx=make_context())
    assert isinstance(exported, dict)
    assert exported[fx.dependent] == expected
