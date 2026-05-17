"""
Tests that ``grapher`` correctly and reliably extracts cells, dependencies,
and guards for different cases, including nested conditionals and cycles.

These tests are intended to map to the micro-workbook examples in
examples/micro_workbooks/extraction_basics.qmd, but without dependency on
or logical/semantic coupling to the xlsx file. Unimplemented roadmap
behaviors are marked ``xfail``.
"""

from __future__ import annotations

from pathlib import Path
from typing import Literal as TypingLiteral

import pytest

from excel_grapher import DynamicRefConfig
from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher import CycleReport, DependencyGraph, create_dependency_graph
from excel_grapher.grapher.guard import And, CellRef, Compare, GuardExpr, Literal, Not
from excel_grapher.grapher.node import NodeKey, NodeView
from tests.integration.user_flows.utils import (
    WorkbookFactory,
    build_workbook_factory,
    write_single_row,
)


@pytest.fixture
def workbook_factory(tmp_path: Path) -> WorkbookFactory:
    return build_workbook_factory(tmp_path, prefix="extraction_basics")


def test_formula_with_no_dependencies_is_extracted_as_single_formula_leaf_node(
    workbook_factory: WorkbookFactory,
) -> None:
    path = workbook_factory(
        lambda ws, _wb: write_single_row(ws, ("Formula with no dependencies", "=1+1"))
    )
    graph: DependencyGraph = create_dependency_graph(path, ["Sheet1!B1"], load_values=True)
    assert len(graph._nodes) == 1

    node: NodeView | None = graph.get_node("Sheet1!B1")
    assert node is not None
    assert node.sheet == "Sheet1"
    assert node.column == "B"
    assert node.row == 1
    assert node.formula == "=1+1"
    assert node.normalized_formula == "=1+1"
    assert node.value == 0
    # Formula with no dependencies must be a leaf node
    assert node.is_leaf is True
    assert dict(node.metadata) == {}

    dependencies: frozenset[NodeKey] = graph.get_dependencies("Sheet1!B1")
    assert dependencies == frozenset()
    dependents: frozenset[NodeKey] = graph.get_dependents("Sheet1!B1")
    assert dependents == frozenset()
    assert not graph._guards
    assert not graph._edge_extra
    assert not graph._hooks
    assert graph.leaf_classification is None


def test_linear_dependency_is_extracted_as_two_nodes_with_one_edge(
    workbook_factory: WorkbookFactory,
) -> None:
    path = workbook_factory(lambda ws, _wb: write_single_row(ws, ("Linear dependency", 1, "=B1+1")))
    graph: DependencyGraph = create_dependency_graph(path, ["Sheet1!C1"], load_values=True)
    assert len(graph._nodes) == 2

    target_node: NodeView | None = graph.get_node("Sheet1!C1")
    assert target_node is not None
    assert target_node.sheet == "Sheet1"
    assert target_node.column == "C"
    assert target_node.row == 1
    assert target_node.formula == "=B1+1"
    assert target_node.normalized_formula == "=Sheet1!B1+1"
    assert target_node.value == 0
    assert target_node.is_leaf is False
    assert dict(target_node.metadata) == {}

    leaf_node: NodeView | None = graph.get_node("Sheet1!B1")
    assert leaf_node is not None
    assert leaf_node.sheet == "Sheet1"
    assert leaf_node.column == "B"
    assert leaf_node.row == 1
    assert leaf_node.formula is None
    assert leaf_node.normalized_formula is None
    assert leaf_node.value == 1
    assert leaf_node.is_leaf is True
    assert dict(leaf_node.metadata) == {}

    dependencies: frozenset[NodeKey] = graph.get_dependencies("Sheet1!C1")
    assert dependencies == frozenset(["Sheet1!B1"])
    dependents: frozenset[NodeKey] = graph.get_dependents("Sheet1!C1")
    assert dependents == frozenset()
    assert not graph._guards
    assert not graph._edge_extra
    assert not graph._hooks
    assert graph.leaf_classification is None


def test_conditions_are_extracted_as_unguarded_but_conditional_branches_as_guarded(
    workbook_factory: WorkbookFactory,
) -> None:
    path = workbook_factory(
        lambda ws, _wb: write_single_row(ws, ("Conditional branches", 1, 10, 20, "=IF(B1=1,C1,D1)"))
    )
    graph: DependencyGraph = create_dependency_graph(path, ["Sheet1!E1"], load_values=True)
    assert len(graph._nodes) == 4

    target_node: NodeView | None = graph.get_node("Sheet1!E1")
    assert target_node is not None
    assert target_node.sheet == "Sheet1"
    assert target_node.column == "E"
    assert target_node.row == 1
    assert target_node.formula == "=IF(B1=1,C1,D1)"
    assert target_node.normalized_formula == "=IF(Sheet1!B1=1,Sheet1!C1,Sheet1!D1)"
    assert target_node.value == 0
    assert target_node.is_leaf is False
    assert dict(target_node.metadata) == {}

    dependencies: frozenset[NodeKey] = graph.get_dependencies("Sheet1!E1")
    assert dependencies == frozenset(["Sheet1!B1", "Sheet1!C1", "Sheet1!D1"])
    dependents: frozenset[NodeKey] = graph.get_dependents("Sheet1!E1")
    assert dependents == frozenset()
    guard: GuardExpr | None = graph.get_edge_guard("Sheet1!E1", "Sheet1!B1")
    guard = graph.get_edge_guard("Sheet1!E1", "Sheet1!C1")
    assert guard == Compare(left=CellRef(key="Sheet1!B1"), op="=", right=Literal(value=1))
    guard = graph.get_edge_guard("Sheet1!E1", "Sheet1!D1")
    assert guard == Not(
        operand=Compare(left=CellRef(key="Sheet1!B1"), op="=", right=Literal(value=1))
    )


@pytest.mark.xfail(
    reason="Nested conditionals are not currently being correctly consolidated into a single AND guard"
)
def test_nested_conditional_in_a_cell_is_extracted_as_an_AND_guard(
    workbook_factory: WorkbookFactory,
) -> None:
    path = workbook_factory(
        lambda ws, _wb: write_single_row(
            ws, ("Nested conditional in a cell", 0, 10, "=IF(NOT(B1=1),IF(B1=0,C1,1),0)")
        )
    )
    graph: DependencyGraph = create_dependency_graph(path, ["Sheet1!D1"], load_values=True)
    assert len(graph._nodes) == 3
    guard: GuardExpr | None = graph.get_edge_guard("Sheet1!D1", "Sheet1!B1")
    assert guard == And(
        operands=(
            Not(operand=Compare(left=CellRef(key="Sheet1!B1"), op="=", right=Literal(value=1))),
            Compare(left=CellRef(key="Sheet1!B1"), op="=", right=Literal(value=0)),
        )
    )


def test_nested_conditional_across_cells_preserves_guards_on_both_edges(
    workbook_factory: WorkbookFactory,
) -> None:
    path = workbook_factory(
        lambda ws, _wb: write_single_row(
            ws,
            ("Nested conditional across cells", 1, 1, "=IF(B1=0,C1,2)", "=IF(B1=1,C1,D1)"),
        )
    )
    graph: DependencyGraph = create_dependency_graph(path, ["Sheet1!E1"], load_values=True)
    assert len(graph._nodes) == 4

    dependencies: frozenset[NodeKey] = graph.get_dependencies("Sheet1!E1")
    assert dependencies == frozenset(["Sheet1!B1", "Sheet1!C1", "Sheet1!D1"])

    e1_to_c1: GuardExpr | None = graph.get_edge_guard("Sheet1!E1", "Sheet1!C1")
    assert e1_to_c1 == Compare(left=CellRef(key="Sheet1!B1"), op="=", right=Literal(value=1))
    e1_to_d1: GuardExpr | None = graph.get_edge_guard("Sheet1!E1", "Sheet1!D1")
    assert e1_to_d1 == Not(
        operand=Compare(left=CellRef(key="Sheet1!B1"), op="=", right=Literal(value=1))
    )
    d1_to_c1: GuardExpr | None = graph.get_edge_guard("Sheet1!D1", "Sheet1!C1")
    assert d1_to_c1 == Compare(left=CellRef(key="Sheet1!B1"), op="=", right=Literal(value=0))


def test_must_cycle_is_reported_as_unconditional_cycle(
    workbook_factory: WorkbookFactory,
) -> None:
    path = workbook_factory(lambda ws, _wb: write_single_row(ws, ("Must cycle", "=C1+1", "=B1+1")))
    graph: DependencyGraph = create_dependency_graph(path, ["Sheet1!C1"], load_values=False)
    assert len(graph._nodes) == 2

    # The cycle is unconditional (no guards on either edge).
    assert not graph._guards
    assert graph.get_dependencies("Sheet1!C1") == frozenset(["Sheet1!B1"])
    assert graph.get_dependencies("Sheet1!B1") == frozenset(["Sheet1!C1"])

    report: CycleReport = graph.cycle_report()
    assert report.has_must_cycles is True
    assert report.has_may_cycles is False
    assert report.must_cycles == [{"Sheet1!B1", "Sheet1!C1"}]
    assert report.may_cycles == []
    # Example path should be a closed traversal of the 2-cycle.
    assert report.example_must_cycle_path is not None
    assert len(report.example_must_cycle_path) == 3
    assert report.example_must_cycle_path[0] == report.example_must_cycle_path[-1]
    assert set(report.example_must_cycle_path) == {"Sheet1!B1", "Sheet1!C1"}
    assert report.example_may_cycle_path is None


def test_wont_cycle_is_not_reported_when_guards_are_mutually_exclusive(
    workbook_factory: WorkbookFactory,
) -> None:
    path = workbook_factory(
        lambda ws, _wb: write_single_row(
            ws, ("Won't cycle", 0, "=IF(B1=0,1,D1)", "=IF(NOT(B1=0),2,C1)")
        )
    )
    graph: DependencyGraph = create_dependency_graph(path, ["Sheet1!C1"], load_values=False)
    assert len(graph._nodes) == 3

    report: CycleReport = graph.cycle_report()
    assert report.has_must_cycles is False
    assert report.has_may_cycles is False
    assert report.must_cycles == []
    assert report.may_cycles == []
    assert report.example_must_cycle_path is None
    assert report.example_may_cycle_path is None


def test_may_cycle_is_reported_when_guards_are_jointly_feasible(
    workbook_factory: WorkbookFactory,
) -> None:
    path = workbook_factory(
        lambda ws, _wb: write_single_row(ws, ("May cycle", 0, "=IF(B1=0,1,D1)", "=IF(B1=1,2,C1)"))
    )
    graph: DependencyGraph = create_dependency_graph(path, ["Sheet1!C1"], load_values=False)
    assert len(graph._nodes) == 3

    report: CycleReport = graph.cycle_report()
    assert report.has_must_cycles is False
    assert report.has_may_cycles is True
    assert report.must_cycles == []
    assert report.may_cycles == [{"Sheet1!C1", "Sheet1!D1"}]
    assert report.example_must_cycle_path is None
    # Example may-cycle path should be a closed traversal of the C1 <-> D1 cycle.
    assert report.example_may_cycle_path is not None
    assert len(report.example_may_cycle_path) == 3
    assert report.example_may_cycle_path[0] == report.example_may_cycle_path[-1]
    assert set(report.example_may_cycle_path) == {"Sheet1!C1", "Sheet1!D1"}


def test_offset_with_scalar_arguments_resolves_to_static_dependency(
    workbook_factory: WorkbookFactory,
) -> None:
    path = workbook_factory(
        lambda ws, _wb: write_single_row(ws, ("OFFSET scalar arguments", 5, 10, "=OFFSET(B1,0,1)"))
    )
    graph: DependencyGraph = create_dependency_graph(path, ["Sheet1!D1"], load_values=False)

    assert len(graph._nodes) == 2
    dependencies: frozenset[NodeKey] = graph.get_dependencies("Sheet1!D1")
    assert dependencies == frozenset(["Sheet1!C1"])


def test_offset_with_dynamic_arguments_cached_resolution_can_break_after_input_change(
    workbook_factory: WorkbookFactory,
) -> None:
    path = workbook_factory(
        lambda ws, _wb: write_single_row(
            ws, ("OFFSET dynamic arguments", 0, 10, 20, "=OFFSET(C1,0,B1)")
        )
    )
    graph: DependencyGraph = create_dependency_graph(
        path,
        ["Sheet1!E1"],
        load_values=False,
        use_cached_dynamic_refs=True,
    )

    graph.set_node_value("Sheet1!B1", 1)
    with FormulaEvaluator(graph) as evaluator, pytest.raises(KeyError, match="Sheet1!D1"):
        evaluator.evaluate("Sheet1!E1")


def test_offset_with_dynamic_arguments_can_expand_with_constraints() -> None:
    path = (
        Path(__file__).resolve().parents[3]
        / "examples"
        / "micro_workbooks"
        / "extraction_basics.xlsx"
    )
    config = DynamicRefConfig.from_constraints({"Sheet1!B10": TypingLiteral[0, 1]}, {})
    graph: DependencyGraph = create_dependency_graph(
        path,
        ["Sheet1!E10"],
        load_values=False,
        dynamic_refs=config,
    )

    dependencies: frozenset[NodeKey] = graph.get_dependencies("Sheet1!E10")
    assert dependencies == frozenset(["Sheet1!B10", "Sheet1!C10", "Sheet1!D10"])


def test_multiple_targets_are_equivalent_for_cells_range_and_named_range(
    workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, wb) -> None:
        ws.write_number(0, 1, 1)  # B1
        ws.write_formula(0, 2, "=B1+1", None, 2)  # C1
        ws.write_formula(0, 3, "=B1+2", None, 3)  # D1
        wb.define_name("MyNamedRange", "='Sheet1'!$C$1:$D$1")

    path = workbook_factory(_populate)

    graph_from_cells = create_dependency_graph(path, ["Sheet1!C1", "Sheet1!D1"], load_values=True)
    graph_from_range = create_dependency_graph(path, ["Sheet1!C1:D1"], load_values=True)
    graph_from_name = create_dependency_graph(path, ["MyNamedRange"], load_values=True)

    assert graph_from_cells == graph_from_range
    assert graph_from_cells == graph_from_name
