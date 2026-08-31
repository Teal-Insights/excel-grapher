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
