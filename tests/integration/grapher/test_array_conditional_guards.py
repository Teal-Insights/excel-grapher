"""Array-context `IF` yields per-element edge guards (issue #483, integration).

`SUM(IF(A1:A10>0,B1:B10,0))` is evaluated element-wise by Excel, so `B3` is read
only when `A3>0`. Each expanded range dependency therefore carries the condition
instantiated at its own element rather than one guard for the whole range.
"""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.guard import And, CellRef, Compare, Literal, Not


def _write(path: Path, formula: str, *, target: str = "D1", array: bool = False) -> Path:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    for row in range(5):
        ws.write_number(row, 0, row - 1)  # A1:A5
        ws.write_number(row, 1, 10 * (row + 1))  # B1:B5
        ws.write_number(row, 2, 100 * (row + 1))  # C1:C5
    ws.write_number(0, 4, 1)  # E1
    if array:
        ws.write_array_formula(f"{target}:{target}", formula, None, 0)
    else:
        ws.write_formula(target, formula, None, 0)
    wb.close()
    return path


def _positive(row: int) -> Compare:
    return Compare(CellRef(f"Sheet1!A{row}"), ">", Literal(0))


def test_sum_if_over_ranges_guards_each_value_element(tmp_path: Path) -> None:
    path = _write(tmp_path / "sum_if.xlsx", "=SUM(IF(A1:A3>0,B1:B3,0))")
    graph = create_dependency_graph(path, ["Sheet1!D1"], load_values=False)

    for row in (1, 2, 3):
        # The whole condition range is read to build the boolean array.
        assert graph.get_edge_guard("Sheet1!D1", f"Sheet1!A{row}") is None
        assert graph.get_edge_guard("Sheet1!D1", f"Sheet1!B{row}") == _positive(row)


def test_sum_if_without_else_still_guards_elements(tmp_path: Path) -> None:
    path = _write(tmp_path / "sum_if_no_else.xlsx", "=SUM(IF(A1:A3>0,B1:B3))")
    graph = create_dependency_graph(path, ["Sheet1!D1"], load_values=False)

    for row in (1, 2, 3):
        assert graph.get_edge_guard("Sheet1!D1", f"Sheet1!B{row}") == _positive(row)


def test_else_range_elements_are_guarded_by_the_negated_condition(tmp_path: Path) -> None:
    path = _write(tmp_path / "sum_if_else.xlsx", "=SUM(IF(A1:A3>0,B1:B3,C1:C3))")
    graph = create_dependency_graph(path, ["Sheet1!D1"], load_values=False)

    for row in (1, 2, 3):
        assert graph.get_edge_guard("Sheet1!D1", f"Sheet1!B{row}") == _positive(row)
        assert graph.get_edge_guard("Sheet1!D1", f"Sheet1!C{row}") == Not(_positive(row))


def test_scalar_operands_in_the_condition_broadcast_across_elements(tmp_path: Path) -> None:
    path = _write(tmp_path / "sum_if_scalar.xlsx", "=SUM(IF(A1:A3=E1,B1:B3,0))")
    graph = create_dependency_graph(path, ["Sheet1!D1"], load_values=False)

    assert graph.get_edge_guard("Sheet1!D1", "Sheet1!E1") is None
    for row in (1, 2, 3):
        assert graph.get_edge_guard("Sheet1!D1", f"Sheet1!B{row}") == Compare(
            CellRef(f"Sheet1!A{row}"), "=", CellRef("Sheet1!E1")
        )


def test_two_dimensional_ranges_are_guarded_element_wise(tmp_path: Path) -> None:
    path = _write(tmp_path / "sum_if_2d.xlsx", "=SUM(IF(A1:B2>0,B1:C2,0))")
    graph = create_dependency_graph(path, ["Sheet1!D1"], load_values=False)

    assert graph.get_edge_guard("Sheet1!D1", "Sheet1!C1") == Compare(
        CellRef("Sheet1!B1"), ">", Literal(0)
    )
    assert graph.get_edge_guard("Sheet1!D1", "Sheet1!C2") == Compare(
        CellRef("Sheet1!B2"), ">", Literal(0)
    )


def test_nested_array_conditionals_conjoin_element_guards(tmp_path: Path) -> None:
    path = _write(tmp_path / "nested.xlsx", "=SUM(IF(A1:A3>0,IF(B1:B3>0,C1:C3,0),0))")
    graph = create_dependency_graph(path, ["Sheet1!D1"], load_values=False)

    for row in (1, 2, 3):
        inner = Compare(CellRef(f"Sheet1!B{row}"), ">", Literal(0))
        assert graph.get_edge_guard("Sheet1!D1", f"Sheet1!C{row}") == And((_positive(row), inner))
        # The inner condition is only evaluated where the outer branch is taken,
        # matching how nested scalar conditionals inherit their branch guard.
        assert graph.get_edge_guard("Sheet1!D1", f"Sheet1!B{row}") == _positive(row)


def test_other_aggregates_also_establish_array_context(tmp_path: Path) -> None:
    path = _write(tmp_path / "average_if.xlsx", "=AVERAGE(IF(A1:A3>0,B1:B3,0))")
    graph = create_dependency_graph(path, ["Sheet1!D1"], load_values=False)

    assert graph.get_edge_guard("Sheet1!D1", "Sheet1!B2") == _positive(2)


def test_cse_array_formula_establishes_array_context_without_an_aggregate(
    tmp_path: Path,
) -> None:
    path = _write(tmp_path / "cse.xlsx", "=IF(A1:A3>0,B1:B3,0)", array=True)
    graph = create_dependency_graph(path, ["Sheet1!D1"], load_values=False)

    for row in (1, 2, 3):
        assert graph.get_edge_guard("Sheet1!D1", f"Sheet1!B{row}") == _positive(row)


def test_non_array_context_range_conditional_stays_unguarded(tmp_path: Path) -> None:
    # Without CSE or an enclosing aggregate, Excel applies implicit intersection;
    # element alignment would not describe that, so stay conservative.
    path = _write(tmp_path / "plain.xlsx", "=IF(A1:A3>0,B1:B3,0)")
    graph = create_dependency_graph(path, ["Sheet1!D1"], load_values=False)

    for row in (1, 2, 3):
        assert graph.get_edge_guard("Sheet1!D1", f"Sheet1!B{row}") is None


def test_explicit_implicit_intersection_stays_unguarded(tmp_path: Path) -> None:
    # `@` marks legacy implicit intersection, which picks one element by the
    # formula's own position rather than evaluating element-wise.
    path = _write(tmp_path / "at_operator.xlsx", "=SUM(IF(@A1:A3>0,@B1:B3,0))")
    graph = create_dependency_graph(path, ["Sheet1!D1"], load_values=False)

    for row in (1, 2, 3):
        assert graph.get_edge_guard("Sheet1!D1", f"Sheet1!B{row}") is None


def test_aggregated_branch_ranges_are_not_element_aligned(tmp_path: Path) -> None:
    # Each active element pulls the whole `SUM(B1:B3)`, so no element guard holds.
    path = _write(tmp_path / "inner_sum.xlsx", "=SUM(IF(A1:A3>0,SUM(B1:B3),0))")
    graph = create_dependency_graph(path, ["Sheet1!D1"], load_values=False)

    for row in (1, 2, 3):
        assert graph.get_edge_guard("Sheet1!D1", f"Sheet1!B{row}") is None


def test_shape_mismatch_falls_back_to_unguarded_edges(tmp_path: Path) -> None:
    path = _write(tmp_path / "mismatch.xlsx", "=SUM(IF(A1:A3>0,B1:B5,0))")
    graph = create_dependency_graph(path, ["Sheet1!D1"], load_values=False)

    for row in (1, 2, 3, 4, 5):
        assert graph.get_edge_guard("Sheet1!D1", f"Sheet1!B{row}") is None


def test_aggregating_logical_condition_falls_back_to_unguarded_edges(tmp_path: Path) -> None:
    # `AND` collapses the array to one boolean; element guards would misdescribe it.
    path = _write(tmp_path / "and_cond.xlsx", "=SUM(IF(AND(A1:A3>0,E1>0),B1:B3,0))")
    graph = create_dependency_graph(path, ["Sheet1!D1"], load_values=False)

    for row in (1, 2, 3):
        assert graph.get_edge_guard("Sheet1!D1", f"Sheet1!B{row}") is None


def test_array_context_ifs_stays_conservative(tmp_path: Path) -> None:
    # Only `IF` grows element guards for now; IFS/CHOOSE/SWITCH over ranges keep
    # the conservative unconditional treatment.
    path = _write(tmp_path / "ifs.xlsx", "=SUM(IFS(A1:A3>0,B1:B3,TRUE,C1:C3))")
    graph = create_dependency_graph(path, ["Sheet1!D1"], load_values=False)

    for row in (1, 2, 3):
        assert graph.get_edge_guard("Sheet1!D1", f"Sheet1!B{row}") is None
        # The `TRUE` catch-all keeps its (unconditional) literal guard.
        assert graph.get_edge_guard("Sheet1!D1", f"Sheet1!C{row}") == Literal(True)


def test_scalar_conditional_inside_an_aggregate_keeps_scalar_guards(tmp_path: Path) -> None:
    # Regression guard for #481: scalar embedded IF is unchanged by array handling.
    path = _write(tmp_path / "scalar.xlsx", "=SUM(IF(E1=1,B1,C1),B2)")
    graph = create_dependency_graph(path, ["Sheet1!D1"], load_values=False)

    scalar_guard = Compare(CellRef("Sheet1!E1"), "=", Literal(1))
    assert graph.get_edge_guard("Sheet1!D1", "Sheet1!B1") == scalar_guard
    assert graph.get_edge_guard("Sheet1!D1", "Sheet1!C1") == Not(scalar_guard)
    assert graph.get_edge_guard("Sheet1!D1", "Sheet1!B2") is None
    assert graph.get_edge_guard("Sheet1!D1", "Sheet1!E1") is None
