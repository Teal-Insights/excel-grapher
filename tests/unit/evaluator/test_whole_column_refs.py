from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import FormulaEvaluator, create_dependency_graph
from tests.integration.utils.parity_harness import evaluate_targets


def test_match_whole_column_finds_interior_value(tmp_path: Path) -> None:
    excel_path = tmp_path / "index_match.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    data = wb.add_worksheet("Data")
    sheet1 = wb.add_worksheet("Sheet1")
    data.write_string(9, 0, "lookup")
    data.write_string(9, 2, "result")
    sheet1.write_formula(
        0,
        1,
        '=INDEX(Data!C:C,MATCH("lookup",Data!A:A,0))',
        None,
        "result",
    )
    wb.close()

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!B1"],
        load_values=True,
        max_range_cells=2,
        use_cached_dynamic_refs=True,
    )
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate("Sheet1!B1") == "result"


def test_match_whole_column_interior_row_25(tmp_path: Path) -> None:
    excel_path = tmp_path / "match_row.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    data = wb.add_worksheet("Data")
    sheet1 = wb.add_worksheet("Sheet1")
    data.write_string(24, 0, "SEA")
    sheet1.write_formula(0, 1, '=MATCH("SEA",Data!A:A,0)', None, 25)
    wb.close()

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!B1"],
        load_values=True,
        max_range_cells=2,
        use_cached_dynamic_refs=True,
    )
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate("Sheet1!B1") == 25


def test_index_match_cross_sheet_quoted_sheet_name(tmp_path: Path) -> None:
    excel_path = tmp_path / "quoted_sheet.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    qb = wb.add_worksheet("QB - Stafford")
    sheet1 = wb.add_worksheet("Sheet1")
    qb.write_string(4, 0, "target")
    qb.write_string(4, 2, "A SEA")
    sheet1.write_formula(
        0,
        0,
        "=INDEX('QB - Stafford'!C:C,MATCH(\"target\",'QB - Stafford'!A:A,0))",
        None,
        "A SEA",
    )
    wb.close()

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!A1"],
        load_values=True,
        max_range_cells=2,
        use_cached_dynamic_refs=True,
    )
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate("Sheet1!A1") == "A SEA"


def test_sum_whole_column_span_uses_used_range(tmp_path: Path) -> None:
    excel_path = tmp_path / "sum_span.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    data = wb.add_worksheet("Data")
    sheet1 = wb.add_worksheet("Sheet1")
    data.write_number(0, 0, 1)
    data.write_number(0, 1, 2)
    data.write_number(0, 2, 3)
    data.write_number(1, 0, 4)
    data.write_number(1, 1, 5)
    data.write_number(1, 2, 6)
    sheet1.write_formula(0, 0, "=SUM(Data!A:C)", None, 21)
    wb.close()

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!A1"],
        load_values=True,
        max_range_cells=2,
        use_cached_dynamic_refs=True,
    )
    results = evaluate_targets(graph, ["Sheet1!A1"])
    assert results["Sheet1!A1"] == 21


def test_index_match_over_whole_column_span(tmp_path: Path) -> None:
    excel_path = tmp_path / "index_span.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    data = wb.add_worksheet("Data")
    sheet1 = wb.add_worksheet("Sheet1")
    data.write_string(0, 0, "name")
    data.write_number(0, 1, 10)
    data.write_number(0, 2, 20)
    data.write_string(1, 0, "lookup")
    data.write_number(1, 1, 3)
    data.write_number(1, 2, 4)
    sheet1.write_formula(
        0,
        0,
        '=INDEX(Data!A:C,MATCH("lookup",Data!A:A,0),3)',
        None,
        4,
    )
    wb.close()

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!A1"],
        load_values=True,
        max_range_cells=2,
        use_cached_dynamic_refs=True,
    )
    results = evaluate_targets(graph, ["Sheet1!A1"])
    assert results["Sheet1!A1"] == 4


def test_match_whole_row_span_on_single_used_column(tmp_path: Path) -> None:
    excel_path = tmp_path / "match_row_span.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    data = wb.add_worksheet("Data")
    sheet1 = wb.add_worksheet("Sheet1")
    data.write_string(4, 0, "a")
    data.write_string(9, 0, "hit")
    sheet1.write_formula(0, 0, '=MATCH("hit",Data!5:10,0)', None, 6)
    wb.close()

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!A1"],
        load_values=True,
        max_range_cells=2,
        use_cached_dynamic_refs=True,
    )
    results = evaluate_targets(graph, ["Sheet1!A1"])
    assert results["Sheet1!A1"] == 6
