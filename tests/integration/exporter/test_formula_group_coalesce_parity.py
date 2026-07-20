"""Sprint 4: coalesced formula-group parity, skips, and cell-only regression."""

from __future__ import annotations

from pathlib import Path

from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.exporter.projection import OptimalCompression
from excel_grapher.grapher.builder import create_dependency_graph
from excel_grapher.grapher.formula_groups import coalesce_formula_groups
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import NodeKind, locate_cell, make_cell_node
from excel_grapher.grapher.range_compression.grouping import column_adjacent_groups
from tests.fixtures.formula_groups.cell_only import (
    build_cross_sheet_scale_cell_only,
    write_cross_sheet_scale_workbook,
)
from tests.fixtures.formula_groups.hand_built import (
    build_cross_sheet_cell_only_twin,
    build_cross_sheet_union_group,
)
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator


def _add_formula(
    g: DependencyGraph,
    address: str,
    formula: str,
    *,
    is_target: bool = False,
) -> None:
    sheet, cell = address.split("!", 1)
    col = "".join(ch for ch in cell if ch.isalpha())
    row = int("".join(ch for ch in cell if ch.isdigit()))
    g.add_node(
        make_cell_node(
            sheet,
            col,
            row,
            formula=formula,
            normalized_formula=formula,
            is_leaf=False,
            is_target=is_target,
        )
    )


def test_coalesced_eval_matches_pre_coalesce_cell_graph() -> None:
    pre = build_cross_sheet_scale_cell_only()
    post = build_cross_sheet_scale_cell_only()
    report = coalesce_formula_groups(post.graph)
    assert len(report.created_groups) == 1

    with FormulaEvaluator(pre.graph) as ev_pre, FormulaEvaluator(post.graph) as ev_post:
        for member in pre.members:
            assert ev_post.evaluate(member) == ev_pre.evaluate(member)


def test_coalesced_codegen_matches_evaluator() -> None:
    fx = build_cross_sheet_scale_cell_only()
    coalesce_formula_groups(fx.graph)
    result = assert_codegen_matches_evaluator(fx.graph, list(fx.members))
    assert result.evaluator_results == {
        "Sheet1!B1": 10.0,
        "Sheet2!B1": 20.0,
    }
    assert result.generated_code.count("def _group_") == 1


def test_coalesced_matches_hand_built_group_for_same_family() -> None:
    cell_only = build_cross_sheet_cell_only_twin()
    hand = build_cross_sheet_union_group()
    report = coalesce_formula_groups(cell_only)
    assert len(report.created_groups) == 1
    assert report.created_groups[0] == hand.group_key

    with FormulaEvaluator(cell_only) as ev_c, FormulaEvaluator(hand.graph) as ev_h:
        for member in hand.members:
            assert ev_c.evaluate(member) == ev_h.evaluate(member)


def test_intra_family_edge_remains_cells_with_skip_reason() -> None:
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(make_cell_node("Sheet1", "A", 1, value=1.0, is_leaf=True))
    _add_formula(g, "Sheet1!B1", "=Sheet1!A1+1")
    _add_formula(g, "Sheet1!B2", "=Sheet1!B1+1")
    g.add_edge("Sheet1!B1", "Sheet1!A1")
    g.add_edge("Sheet1!B2", "Sheet1!B1")

    report = coalesce_formula_groups(g)
    assert report.created_groups == ()
    assert len(report.skipped_families) == 1
    assert report.skipped_families[0].reason == "intra_family_edge"
    assert g.get_node("Sheet1!B1") is not None
    assert g.get_node("Sheet1!B2") is not None
    assert g.get_node("Sheet1!B1").kind is NodeKind.cell


def test_lone_formula_cell_never_becomes_one_member_group() -> None:
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    g.add_node(make_cell_node("Sheet1", "A", 1, value=1.0, is_leaf=True))
    _add_formula(g, "Sheet1!B1", "=Sheet1!A1*10", is_target=True)
    g.add_edge("Sheet1!B1", "Sheet1!A1")

    report = coalesce_formula_groups(g)
    assert report.created_groups == ()
    assert any(s.reason == "below_min_size" for s in report.skipped_families)
    assert g.get_node("Sheet1!B1") is not None
    assert g.get_node("Sheet1!B1").kind is NodeKind.cell


def test_builder_formula_groups_false_matches_default(tmp_path: Path) -> None:
    wb_path = write_cross_sheet_scale_workbook(tmp_path / "scale.xlsx")
    kwargs = dict(
        workbook=wb_path,
        targets=["Sheet1!B1", "Sheet2!B1"],
        load_values=True,
        use_cached_dynamic_refs=True,
    )
    default = create_dependency_graph(**kwargs)
    explicit_off = create_dependency_graph(**kwargs, formula_groups=False)
    assert default.keys(order="workbook") == explicit_off.keys(order="workbook")
    for key in default.keys():
        a = default.get_node(key)
        b = explicit_off.get_node(key)
        assert a is not None and b is not None
        assert a.kind is b.kind
        assert a.normalized_formula == b.normalized_formula
        assert a.value == b.value


def test_dependent_formula_text_unchanged_after_coalesce() -> None:
    fx = build_cross_sheet_scale_cell_only()
    _add_formula(fx.graph, "Sheet1!Z9", "=Sheet1!B1+1")
    fx.graph.add_edge("Sheet1!Z9", "Sheet1!B1")
    formula_before = fx.graph.get_node("Sheet1!Z9")
    assert formula_before is not None
    text_before = formula_before.normalized_formula

    report = coalesce_formula_groups(fx.graph)
    group_key = report.created_groups[0]
    z9 = fx.graph.get_node("Sheet1!Z9")
    assert z9 is not None
    assert z9.normalized_formula == text_before == "=Sheet1!B1+1"
    assert fx.graph.get_dependencies("Sheet1!Z9") == frozenset({group_key})


def test_optimal_compression_and_taco_grouping_skip_coalesced_group() -> None:
    fx = build_cross_sheet_scale_cell_only()
    report = coalesce_formula_groups(fx.graph)
    group_key = report.created_groups[0]

    # TACO column-adjacent grouping must not absorb the multi-cell group key.
    groups = column_adjacent_groups(fx.graph, min_len=2)
    flat = {key for group in groups for key in group}
    assert group_key not in flat

    # OptimalCompression must not crash; group remains addressable by occupancy.
    projection = OptimalCompression().project(fx.graph)
    assert group_key in projection.projected_graph
    assert projection.projected_graph.get_node(group_key) is not None
    assert projection.projected_graph.cell_owner("Sheet1!B1") == group_key
    assert locate_cell(projection.projected_graph, "Sheet2!B1") is not None


def test_evaluation_order_includes_group_without_crash() -> None:
    fx = build_cross_sheet_scale_cell_only()
    report = coalesce_formula_groups(fx.graph)
    group_key = report.created_groups[0]
    order = fx.graph.evaluation_order()
    assert group_key in order
    # Leaves before group.
    assert order.index("Sheet1!A1") < order.index(group_key)
    assert order.index("Sheet2!A1") < order.index(group_key)
