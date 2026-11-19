"""Live Excel recalc after `write_workbook` (run-if-available, #564)."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher import create_dependency_graph, write_workbook
from tests.utils.excel_live_parity import compare_cached_to_evaluator
from tests.utils.modify_and_recalculate import (
    ExcelRecalculationError,
    modify_and_recalculate_workbook,
)


@pytest.mark.slow
def test_write_workbook_live_excel_matches_evaluator(tmp_path: Path) -> None:
    import fastpyxl

    source = tmp_path / "src.xlsx"
    dest = tmp_path / "written.xlsx"
    recalc = tmp_path / "recalc.xlsx"

    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 2
    ws["B1"] = "=A1*3"
    wb.save(source)
    wb.close()

    graph = create_dependency_graph(source, ["Sheet1!B1"], load_values=True)
    write_workbook(graph, dest)

    try:
        modify_and_recalculate_workbook(dest, recalc, {})
    except (ExcelRecalculationError, RuntimeError, ImportError) as exc:
        pytest.skip(f"Excel recalculation not available: {exc}")

    live = create_dependency_graph(recalc, ["Sheet1!B1"], load_values=True)
    with FormulaEvaluator(graph) as ev:
        computed = ev.evaluate(["Sheet1!B1"])["Sheet1!B1"]
    node = live.get_node("Sheet1!B1")
    assert node is not None
    mismatch = compare_cached_to_evaluator(node.value, computed)
    assert mismatch is None, f"Excel {node.value!r} vs evaluator {computed!r}"


@pytest.mark.slow
def test_write_workbook_in_place_live_excel_recalculates_downstream_and_clears_vacated(
    tmp_path: Path,
) -> None:
    import fastpyxl

    workbook = tmp_path / "in_place_model.xlsx"
    recalc = tmp_path / "in_place_recalc.xlsx"

    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 4
    ws["B1"] = "=A1*3"
    ws["C1"] = "=B1+5"  # Off-closure downstream formula
    ws["D1"] = 10
    ws["E1"] = "=D1*2"  # Moved in-graph node
    ws["Z1"] = "Preserved Header"
    wb.save(workbook)
    wb.close()

    graph = create_dependency_graph(workbook, ["Sheet1!B1", "Sheet1!E1"], load_values=True)
    graph.set_node_value("Sheet1!A1", 7)
    graph.replace_node_formula("Sheet1!B1", "=A1*4", None)
    graph.move_node("Sheet1!E1", "Sheet1!E3")
    write_workbook(graph, workbook, in_place=True)

    try:
        modify_and_recalculate_workbook(workbook, recalc, {})
    except (ExcelRecalculationError, RuntimeError, ImportError) as exc:
        pytest.skip(f"Excel recalculation not available: {exc}")

    live = create_dependency_graph(
        recalc, ["Sheet1!B1", "Sheet1!C1", "Sheet1!E3"], load_values=True
    )
    with FormulaEvaluator(graph) as ev:
        computed = ev.evaluate(["Sheet1!B1", "Sheet1!E3"])

    for target in ("Sheet1!B1", "Sheet1!E3"):
        node = live.get_node(target)
        assert node is not None
        mismatch = compare_cached_to_evaluator(node.value, computed[target])
        assert mismatch is None, f"{target}: Excel {node.value!r} vs evaluator {computed[target]!r}"

    # Off-closure downstream formula Sheet1!C1 (=B1+5 -> 28+5 = 33) recalculated in Excel,
    # vacated Sheet1!E1 is empty, and off-closure header Sheet1!Z1 survived.
    c1_node = live.get_node("Sheet1!C1")
    assert c1_node is not None
    assert compare_cached_to_evaluator(c1_node.value, 33) is None

    wb_recalc = fastpyxl.load_workbook(recalc, data_only=True)
    try:
        assert wb_recalc["Sheet1"]["E1"].value is None
        assert wb_recalc["Sheet1"]["Z1"].value == "Preserved Header"
    finally:
        wb_recalc.close()
