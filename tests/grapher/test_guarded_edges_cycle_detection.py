from __future__ import annotations

from pathlib import Path

import pytest
import xlsxwriter

from excel_grapher import CycleError, create_dependency_graph
from excel_grapher.evaluator.codegen import CodeGenerator
from tests.utils.workbook_xml import patch_workbook_calcpr


def _make_may_cycle_if_workbook(path: Path) -> None:
    """
    Create a workbook with a *may-cycle* that is broken by mutually exclusive IF guards:

    A1 = IF($C$1=0, B1, 1)
    B1 = IF($C$1=1, A1, 2)
    C1 = 0
    """
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")

    ws.write_number(0, 2, 0)  # C1
    ws.write_formula(0, 0, "=IF($C$1=0,B1,1)", None, 1)  # A1 cached
    ws.write_formula(0, 1, "=IF($C$1=1,A1,2)", None, 2)  # B1 cached

    wb.close()


def _make_feasible_may_cycle_if_workbook(path: Path) -> None:
    """
    Create a workbook with a *feasible may-cycle* (guarded edges only):

    A1 = IF($C$1=0, B1, 1)
    B1 = IF($C$1=0, A1, 2)
    C1 = 0
    """
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")

    ws.write_number(0, 2, 0)  # C1
    ws.write_formula(0, 0, "=IF($C$1=0,B1,1)", None, 1)  # A1 cached
    ws.write_formula(0, 1, "=IF($C$1=0,A1,2)", None, 2)  # B1 cached

    wb.close()


def _make_must_cycle_workbook(path: Path) -> None:
    """
    Create a workbook with a *must-cycle* (unconditional):

    A1 = B1
    B1 = A1
    """
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_formula(0, 0, "=B1", None, 0)  # A1 cached
    ws.write_formula(0, 1, "=A1", None, 0)  # B1 cached
    wb.close()


def test_guarded_edges_are_created_for_top_level_if(tmp_path: Path) -> None:
    excel_path = tmp_path / "if_may_cycle.xlsx"
    _make_may_cycle_if_workbook(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)

    # Condition cell is always depended on (unconditional)
    assert "Sheet1!C1" in graph.dependencies("Sheet1!A1")
    assert "Sheet1!C1" in graph.dependencies("Sheet1!B1")

    # Branch-only deps are guarded
    a1_to_b1 = graph.edge_attrs("Sheet1!A1", "Sheet1!B1")
    b1_to_a1 = graph.edge_attrs("Sheet1!B1", "Sheet1!A1")
    assert a1_to_b1.get("guard") is not None
    assert b1_to_a1.get("guard") is not None

    a1_to_c1 = graph.edge_attrs("Sheet1!A1", "Sheet1!C1")
    b1_to_c1 = graph.edge_attrs("Sheet1!B1", "Sheet1!C1")
    assert a1_to_c1.get("guard") is None
    assert b1_to_c1.get("guard") is None


def test_cycle_report_distinguishes_must_vs_may_cycles(tmp_path: Path) -> None:
    excel_path = tmp_path / "if_feasible_may_cycle.xlsx"
    _make_feasible_may_cycle_if_workbook(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)
    report = graph.cycle_report()

    assert report.has_must_cycles is False
    assert report.has_may_cycles is True
    assert report.must_cycles == []
    assert any({"Sheet1!A1", "Sheet1!B1"} == s for s in report.may_cycles)


def test_infeasible_may_cycle_is_not_reported(tmp_path: Path) -> None:
    """
    This workbook contains a syntactic SCC {A1, B1} only when guarded edges are
    included, but the only cycle requires mutually contradictory guards:

      A1 -> B1 requires C1=0
      B1 -> A1 requires C1=1

    Since those cannot both be true, the cycle is infeasible and should not be
    reported as a may-cycle.
    """
    excel_path = tmp_path / "if_may_cycle.xlsx"
    _make_may_cycle_if_workbook(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)
    report = graph.cycle_report()
    assert report.has_must_cycles is False
    assert report.has_may_cycles is False


def test_evaluation_order_strict_true_raises_on_may_cycle(tmp_path: Path) -> None:
    excel_path = tmp_path / "if_feasible_may_cycle.xlsx"
    _make_feasible_may_cycle_if_workbook(excel_path)
    graph = create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)

    with pytest.raises(CycleError) as e:
        graph.evaluation_order(strict=True)
    assert e.value.is_must_cycle is False


def test_evaluation_order_strict_false_warns_and_excludes_may_cycle_nodes(tmp_path: Path) -> None:
    excel_path = tmp_path / "if_feasible_may_cycle.xlsx"
    _make_feasible_may_cycle_if_workbook(excel_path)
    graph = create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)

    with pytest.warns(UserWarning):
        order = graph.evaluation_order(strict=False)

    # May-cycle nodes are excluded; unconditional leaf remains.
    assert "Sheet1!C1" in order
    assert "Sheet1!A1" not in order
    assert "Sheet1!B1" not in order


def test_evaluation_order_iterate_true_raises_on_may_cycle(tmp_path: Path) -> None:
    """Workbook iterative calc + may-cycle: we cannot emulate convergence; fail fast."""
    base = tmp_path / "may_cycle_base.xlsx"
    _make_feasible_may_cycle_if_workbook(base)
    excel_path = tmp_path / "may_cycle_iterate.xlsx"
    patch_workbook_calcpr(base, excel_path, iterate=True, iterate_count=100, iterate_delta=0.001)

    graph = create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)

    with pytest.raises(CycleError) as e:
        graph.evaluation_order(strict=False, iterate_enabled=True)
    assert e.value.is_must_cycle is False
    assert "guarded" in str(e.value).lower()
    assert "iterate" in str(e.value).lower()


def test_codegen_iterate_true_emits_iterative_runtime_on_may_cycle(tmp_path: Path) -> None:
    base = tmp_path / "may_cycle_codegen_base.xlsx"
    _make_feasible_may_cycle_if_workbook(base)
    excel_path = tmp_path / "may_cycle_codegen_iterate.xlsx"
    patch_workbook_calcpr(base, excel_path, iterate=True, iterate_count=100, iterate_delta=0.001)

    graph = create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)

    code = CodeGenerator(graph, iterate_enabled=True).generate(["Sheet1!A1"])
    assert "xl_iterative_compute" in code


def test_evaluation_order_iterate_true_raises_on_must_cycle(tmp_path: Path) -> None:
    base = tmp_path / "must_cycle_base.xlsx"
    _make_must_cycle_workbook(base)
    excel_path = tmp_path / "must_cycle_iterate.xlsx"
    patch_workbook_calcpr(base, excel_path, iterate=True, iterate_count=100, iterate_delta=0.001)

    graph = create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)

    with pytest.raises(CycleError) as e:
        graph.evaluation_order(strict=False, iterate_enabled=True)
    assert e.value.is_must_cycle is True
    assert "unconditional" in str(e.value).lower()
    assert "iterate" in str(e.value).lower()


def test_must_cycle_is_reported_and_always_raises(tmp_path: Path) -> None:
    excel_path = tmp_path / "must_cycle.xlsx"
    _make_must_cycle_workbook(excel_path)
    graph = create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)

    report = graph.cycle_report()
    assert report.has_must_cycles is True
    assert any({"Sheet1!A1", "Sheet1!B1"} == s for s in report.must_cycles)

    with pytest.raises(CycleError) as e1:
        graph.evaluation_order(strict=True)
    assert e1.value.is_must_cycle is True

    with pytest.raises(CycleError) as e2:
        graph.evaluation_order(strict=False)
    assert e2.value.is_must_cycle is True
