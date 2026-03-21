from __future__ import annotations

from pathlib import Path
from typing import cast

import pytest
import xlsxwriter

from excel_grapher import CycleError, FormulaEvaluator, create_dependency_graph, get_calc_settings
from excel_grapher.evaluator.codegen import CodeGenerator
from excel_grapher.evaluator.export_runtime.cache import CircularReferenceWarning
from tests.utils.workbook_xml import patch_workbook_calcpr


def _make_self_cycle_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_formula(0, 0, "=A1", None, 0)
    wb.close()


def _make_iterative_convergence_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_formula(0, 0, "=(A1+1)/2", None, 0)
    wb.close()


def _make_oscillation_workbook(path: Path) -> None:
    """A1 = 1 - A1 oscillates between 0 and 1; never converges for positive iterate_delta."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_formula(0, 0, "=1-A1", None, 0)
    wb.close()


def test_workbook_iterate_disabled_keeps_default_circular_behavior(tmp_path: Path) -> None:
    base = tmp_path / "cycle_base.xlsx"
    _make_self_cycle_workbook(base)
    workbook = tmp_path / "cycle_iterate_off.xlsx"
    patch_workbook_calcpr(
        base, workbook, iterate=False, iterate_count=100, iterate_delta=0.001
    )

    settings = get_calc_settings(workbook)
    graph = create_dependency_graph(workbook, ["Sheet1!A1"], load_values=False)

    with FormulaEvaluator(
        graph,
        iterate_enabled=settings.iterate_enabled,
        iterate_count=settings.iterate_count,
        iterate_delta=settings.iterate_delta,
    ) as ev, pytest.warns(CircularReferenceWarning):
        evaluator_result = ev.evaluate(["Sheet1!A1"])
    assert evaluator_result["Sheet1!A1"] == 0

    with pytest.raises(CycleError, match="Must-cycle"):
        CodeGenerator(
            graph,
            iterate_enabled=settings.iterate_enabled,
            iterate_count=settings.iterate_count,
            iterate_delta=settings.iterate_delta,
        ).generate(["Sheet1!A1"])


def test_workbook_iterate_enabled_drives_iterative_convergence(tmp_path: Path) -> None:
    base = tmp_path / "iterative_base.xlsx"
    _make_iterative_convergence_workbook(base)
    workbook = tmp_path / "iterative_on.xlsx"
    patch_workbook_calcpr(
        base, workbook, iterate=True, iterate_count=75, iterate_delta=1e-6
    )

    settings = get_calc_settings(workbook)
    graph = create_dependency_graph(workbook, ["Sheet1!A1"], load_values=False)

    with FormulaEvaluator(
        graph,
        iterate_enabled=settings.iterate_enabled,
        iterate_count=settings.iterate_count,
        iterate_delta=settings.iterate_delta,
    ) as ev:
        evaluator_result = ev.evaluate(["Sheet1!A1"])

    generated_code = CodeGenerator(
        graph,
        iterate_enabled=settings.iterate_enabled,
        iterate_count=settings.iterate_count,
        iterate_delta=settings.iterate_delta,
    ).generate(["Sheet1!A1"])
    assert "iterate_count=75" in generated_code
    ns: dict[str, object] = {}
    exec(generated_code, ns)
    generated_result = cast("dict[str, object]", ns["compute_all"]())

    assert abs(float(evaluator_result["Sheet1!A1"]) - 1.0) <= 1e-4
    assert abs(float(generated_result["Sheet1!A1"]) - 1.0) <= 1e-4


def test_workbook_iterate_max_iterations_without_convergence_parity(tmp_path: Path) -> None:
    """Oscillation never meets iterateDelta; both paths must agree after iterateCount sweeps."""
    base = tmp_path / "oscillation_base.xlsx"
    _make_oscillation_workbook(base)
    workbook = tmp_path / "oscillation_iterate_on.xlsx"
    patch_workbook_calcpr(
        base, workbook, iterate=True, iterate_count=3, iterate_delta=1e-12
    )

    settings = get_calc_settings(workbook)
    assert settings.iterate_count == 3
    assert settings.iterate_delta == 1e-12

    graph = create_dependency_graph(workbook, ["Sheet1!A1"], load_values=False)

    with FormulaEvaluator(
        graph,
        iterate_enabled=settings.iterate_enabled,
        iterate_count=settings.iterate_count,
        iterate_delta=settings.iterate_delta,
    ) as ev:
        evaluator_result = ev.evaluate(["Sheet1!A1"])

    generated_code = CodeGenerator(
        graph,
        iterate_enabled=settings.iterate_enabled,
        iterate_count=settings.iterate_count,
        iterate_delta=settings.iterate_delta,
    ).generate(["Sheet1!A1"])
    assert "iterate_count=3" in generated_code
    assert "iterate_delta=" in generated_code

    ns: dict[str, object] = {}
    exec(generated_code, ns)
    generated_result = cast("dict[str, object]", ns["compute_all"]())

    assert evaluator_result["Sheet1!A1"] == generated_result["Sheet1!A1"]
