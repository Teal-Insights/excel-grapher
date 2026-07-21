"""Minimal regressions for ``verify_lic_dsf_modules`` failure modes.

Issues from a LIC-DSF Chart Data verify run:

1. **formula_groups KeyError** — ``'Chart Data'!D67 not found in graph`` when
   ``D74 = SUM(D67:N67)`` walks coalesced group members that are not generate
   targets.
2. **cell_only numeric_drift** — Chart Data ``E250``/``E292`` (and ~225 peers)
   drift because ``AVERAGE`` treated ``=""`` blanks as ``0``, halving
   ``Input 6!G36`` and cascading through commodity / PV residual financing.
"""

from __future__ import annotations

from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import make_cell_node
from tests.fixtures.formula_groups.hand_built import (
    build_chart_data_threshold_sum_group,
    build_offset_address_hole_group,
    build_sum_over_constant_group,
)


# --- Issue 1: formula_groups KeyError on Chart Data!D67 ---


def test_verify_issue_sum_over_non_target_group_members_emits_wrappers() -> None:
    fx = build_sum_over_constant_group()
    assert fx.dependent is not None
    with CodeGenerator(fx.graph) as gen:
        code = gen.generate(targets=[fx.dependent])
    assert "def cell_sheet1_a1(ctx):" in code
    assert "def cell_sheet1_a2(ctx):" in code


def test_verify_issue_chart_data_d74_sum_over_d67_group_matches_evaluator() -> None:
    """Exact shape: ``D74 = SUM(D67:E67)``; only ``D74`` is a generate target."""
    fx = build_chart_data_threshold_sum_group()
    assert fx.dependent == "Sheet1!D74"
    with FormulaEvaluator(fx.graph) as ev:
        assert ev.evaluate("Sheet1!D67") == 1.0
        expected = ev.evaluate(fx.dependent)
    assert expected == 2.0

    with CodeGenerator(fx.graph) as gen:
        code = gen.generate(targets=[fx.dependent])
    assert "def cell_sheet1_d67(ctx):" in code
    assert "def cell_sheet1_e67(ctx):" in code

    ns: dict[str, object] = {}
    exec(code, ns)
    compute_all = ns["compute_all"]
    make_context = ns["make_context"]
    assert callable(compute_all) and callable(make_context)
    exported = compute_all(ctx=make_context())
    assert isinstance(exported, dict)
    assert exported[fx.dependent] == expected


def test_verify_issue_chart_data_d74_generate_modules_resolves_d67() -> None:
    """Modular export path used by ``codegen_lic_dsf.py`` must also wrap D67."""
    fx = build_chart_data_threshold_sum_group()
    assert fx.dependent is not None
    with CodeGenerator(fx.graph) as gen:
        modules = gen.generate_modules(targets=[fx.dependent])
    internals = modules["internals.py"]
    assert "def cell_sheet1_d67(ctx):" in internals
    assert "def cell_sheet1_e67(ctx):" in internals
    assert "def cell_sheet1_d74(ctx):" in internals


# --- Issue 2: cell_only ↔ excel numeric_drift (AVERAGE blanks) ---


def _average_blank_cascade_graph() -> DependencyGraph:
    """Input-6 ``G36`` AVERAGE blanks → commodity-style ratio → Chart Data display.

    Excel: ``AVERAGE("", shock, "", shock) == shock``.
    Bug: blanks coerced to 0 → average ``shock/2`` → display drifts like E250.
    """
    shock = -0.2168778144226342
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(
        make_cell_node(
            "Sheet1", "G", 38, formula='=""', normalized_formula='=""', is_leaf=False
        )
    )
    g.add_node(make_cell_node("Sheet1", "G", 39, value=shock, is_leaf=True))
    g.add_node(
        make_cell_node(
            "Sheet1", "G", 40, formula='=""', normalized_formula='=""', is_leaf=False
        )
    )
    g.add_node(make_cell_node("Sheet1", "G", 41, value=shock, is_leaf=True))
    g.add_node(
        make_cell_node(
            "Sheet1",
            "G",
            36,
            formula="=AVERAGE(G38:G41)",
            normalized_formula="=AVERAGE(Sheet1!G38:G41)",
            is_leaf=False,
        )
    )
    # Tailored-test scale (K27 ≈ G36 * 0.75 / -10%) then pass through to Chart Data.
    g.add_node(
        make_cell_node(
            "Sheet1",
            "K",
            27,
            formula="=G36*0.75/-10%",
            normalized_formula="=Sheet1!G36*0.75/-10%",
            is_leaf=False,
        )
    )
    g.add_node(
        make_cell_node(
            "Sheet1",
            "E",
            250,
            formula="=K27",
            normalized_formula="=Sheet1!K27",
            is_leaf=False,
        )
    )
    for m in ("Sheet1!G38", "Sheet1!G39", "Sheet1!G40", "Sheet1!G41"):
        g.add_edge("Sheet1!G36", m)
    g.add_edge("Sheet1!K27", "Sheet1!G36")
    g.add_edge("Sheet1!E250", "Sheet1!K27")
    return g


def test_verify_issue_average_blank_cascade_matches_excel_scale() -> None:
    shock = -0.2168778144226342
    # K27 = shock * 0.75 / -0.10 = shock * -7.5
    expect_k27 = shock * 0.75 / -0.10
    graph = _average_blank_cascade_graph()

    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate("Sheet1!G36") == shock
        assert ev.evaluate("Sheet1!K27") == expect_k27
        assert ev.evaluate("Sheet1!E250") == expect_k27

    with CodeGenerator(graph) as gen:
        code = gen.generate(targets=["Sheet1!E250"])
    ns: dict[str, object] = {}
    exec(code, ns)
    compute_all = ns["compute_all"]
    assert callable(compute_all)
    exported = compute_all()
    assert isinstance(exported, dict)
    assert exported["Sheet1!E250"] == expect_k27
    # Guard against the historical half-average bug (shock/2 * -7.5).
    assert exported["Sheet1!E250"] != (shock / 2.0) * 0.75 / -0.10


def test_verify_issue_offset_hole_group_matches_evaluator() -> None:
    """OFFSET base address holes must not compile to unconditional #REF!."""
    from excel_grapher.evaluator.name_utils import address_to_python_name

    fx = build_offset_address_hole_group()
    with FormulaEvaluator(fx.graph) as ev:
        expected = {m: ev.evaluate(m) for m in fx.members}
    assert expected == {"Sheet1!A12": "alpha", "Sheet1!A13": "beta"}

    with CodeGenerator(fx.graph) as gen:
        code = gen.generate(targets=list(fx.members))
    assert "xl_address_ref_info(" in code
    assert "xl_offset(ctx, xl_address_ref_info(b0)" in code

    ns: dict[str, object] = {}
    exec(code, ns)
    make_context = ns["make_context"]
    assert callable(make_context)
    ctx = make_context()
    for member, want in expected.items():
        wrapper = ns[address_to_python_name(member)]
        assert callable(wrapper)
        assert wrapper(ctx) == want
