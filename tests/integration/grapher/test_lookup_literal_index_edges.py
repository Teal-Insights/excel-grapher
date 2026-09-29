"""HLOOKUP/VLOOKUP with a literal index read only the key line and result line (#1027)."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.graph_consistency import collect_graph_consistency_issues
from excel_grapher.grapher.parser import lookup_table_arg_spans
from tests.integration.utils.parity_harness import evaluate_targets


def _hlookup_workbook(path: Path, formula: str) -> None:
    """Table `B1:D3`: headers row 1, spare row 2, results row 3; A1 hosts `formula`."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    for col in range(1, 4):
        ws.write_number(0, col, col)  # B1:D1 headers
        ws.write_number(2, col, col * 10)  # B3:D3 results
    ws.write_formula(1, 2, "=A1+1")  # C2 closes a cycle only via row 2
    ws.write_number(4, 0, 2)  # A5 key
    ws.write_formula(0, 0, formula)
    wb.close()


def _vlookup_workbook(path: Path, formula: str) -> None:
    """Table `A2:C4`: keys col A, spare col B, results col C; E1 hosts `formula`."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    for row in range(1, 4):
        ws.write_number(row, 0, row)  # A2:A4 keys
        ws.write_number(row, 2, row * 10)  # C2:C4 results
    ws.write_formula(2, 1, "=E1+1")  # B3 closes a cycle only via col B
    ws.write_number(0, 5, 2)  # F1 key
    ws.write_formula(0, 4, formula)
    wb.close()


def test_hlookup_literal_index_skips_non_result_rows(tmp_path: Path) -> None:
    path = tmp_path / "hlookup.xlsx"
    _hlookup_workbook(path, "=HLOOKUP(A5,B1:D3,3,FALSE)")

    graph = create_dependency_graph(path, ["Sheet1!A1"], load_values=False)

    assert graph.get_dependencies("Sheet1!A1") == frozenset(
        {"Sheet1!A5", "Sheet1!B1", "Sheet1!C1", "Sheet1!D1", "Sheet1!B3", "Sheet1!C3", "Sheet1!D3"}
    )
    assert graph.cycle_report().has_may_cycles is False
    assert graph.cycle_report().has_must_cycles is False
    assert collect_graph_consistency_issues(graph) == ()


def test_vlookup_literal_index_skips_non_result_columns(tmp_path: Path) -> None:
    path = tmp_path / "vlookup.xlsx"
    _vlookup_workbook(path, "=VLOOKUP(F1,A2:C4,3,FALSE)")

    graph = create_dependency_graph(path, ["Sheet1!E1"], load_values=False)

    assert graph.get_dependencies("Sheet1!E1") == frozenset(
        {"Sheet1!F1", "Sheet1!A2", "Sheet1!A3", "Sheet1!A4", "Sheet1!C2", "Sheet1!C3", "Sheet1!C4"}
    )
    report = graph.cycle_report()
    assert report.has_may_cycles is False
    assert report.has_must_cycles is False
    assert collect_graph_consistency_issues(graph) == ()


def test_lookup_result_line_cycle_still_reported(tmp_path: Path) -> None:
    path = tmp_path / "hlookup_row2.xlsx"
    _hlookup_workbook(path, "=HLOOKUP(A5,B1:D3,2,FALSE)")

    graph = create_dependency_graph(path, ["Sheet1!A1"], load_values=False)

    assert "Sheet1!C2" in graph.get_dependencies("Sheet1!A1")
    assert "Sheet1!B3" not in graph.get_dependencies("Sheet1!A1")
    assert graph.cycle_report().has_must_cycles is True


def test_non_literal_index_keeps_whole_table(tmp_path: Path) -> None:
    path = tmp_path / "hlookup_dynamic.xlsx"
    _hlookup_workbook(path, "=HLOOKUP(A5,B1:D3,A5+1,FALSE)")

    graph = create_dependency_graph(path, ["Sheet1!A1"], load_values=False)

    deps = graph.get_dependencies("Sheet1!A1")
    assert {"Sheet1!B2", "Sheet1!C2", "Sheet1!D2"} <= deps
    assert graph.cycle_report().has_must_cycles is True
    assert collect_graph_consistency_issues(graph) == ()


def test_out_of_range_literal_index_keeps_key_line_only(tmp_path: Path) -> None:
    path = tmp_path / "hlookup_ref.xlsx"
    _hlookup_workbook(path, "=HLOOKUP(A5,B1:D3,7,FALSE)")

    graph = create_dependency_graph(path, ["Sheet1!A1"], load_values=False)

    assert graph.get_dependencies("Sheet1!A1") == frozenset(
        {"Sheet1!A5", "Sheet1!B1", "Sheet1!C1", "Sheet1!D1"}
    )
    assert collect_graph_consistency_issues(graph) == ()


def test_lookup_table_arg_spans_only_literal_indices() -> None:
    formula = "=HLOOKUP(A5,Sheet1!B1:D3,3)+VLOOKUP(F1, A2:C4 ,2.9,0)+VLOOKUP(F1,A2:C4,G1)"
    spans = lookup_table_arg_spans(formula)
    assert len(spans) == 2
    got = {formula[a:b]: sel for (a, b), sel in spans.items()}
    assert got == {"Sheet1!B1:D3": ("HLOOKUP", 3), "A2:C4": ("VLOOKUP", 2)}


def test_lookup_table_arg_spans_rejects_non_positive_index() -> None:
    assert lookup_table_arg_spans("=VLOOKUP(F1,A2:C4,0)") == {}
    assert lookup_table_arg_spans("=VLOOKUP(F1,A2:C4,-1)") == {}


def test_narrowed_lookup_graph_evaluates_and_exports(tmp_path: Path) -> None:
    path = tmp_path / "lookup_parity.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    for col in range(1, 4):
        ws.write_number(0, col, col)  # B1:D1 headers
        ws.write_formula(1, col, "=1/0")  # B2:D2 unread errors
        ws.write_number(2, col, col * 10)  # B3:D3 results
    ws.write_number(4, 0, 2)  # A5 key
    ws.write_formula(0, 0, "=HLOOKUP($A$5,$B$1:$D$3,3,FALSE)", None, 20)
    ws.write_formula(5, 0, "=VLOOKUP(10,'Sheet1'!$B$3:$D$3,2.9,TRUE)", None, 20)
    ws.write_formula(6, 0, "=HLOOKUP(3,B:D,3,FALSE)", None, 30)
    wb.close()

    targets = ["Sheet1!A1", "Sheet1!A6", "Sheet1!A7"]
    graph = create_dependency_graph(path, targets, load_values=True)

    assert "Sheet1!C2" not in graph
    results = evaluate_targets(graph, targets)
    assert results == {"Sheet1!A1": 20, "Sheet1!A6": 20, "Sheet1!A7": 30}
    assert collect_graph_consistency_issues(graph) == ()
