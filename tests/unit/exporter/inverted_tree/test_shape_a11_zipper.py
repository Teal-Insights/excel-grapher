"""Layer A11 — period-lag zippers are a cell DAG and a series SCC (#603)."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.ast_emit import emit_rung2_scc, emit_rung3_scc
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.runtime import (
    XlError,
    as_measure,
    demand_instance,
    eval_instance,
    is_error,
    live_measure,
)
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)


def _zipper_workbook(tmp_path: Path) -> Path:
    """Cell-acyclic zipper from `plans/inverted-tree-scheduling.md`.

    `debt_t = debt_{t-1} + adj_t` and `adj_t = debt_{t-1} * r` (lag only).
    """
    return write_workbook(
        tmp_path / "a11_zipper.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=100",
                "B2": "=A2+B3",
                "C2": "=B2+C3",
                "B3": "=A2*0.02",
                "C3": "=B2*0.02",
            },
        },
    )


def _zipper_bindings() -> dict:
    return bindings_document(
        series_entry(
            "debt",
            "Engine!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
        series_entry(
            "adjustment",
            "Engine!B3:C3",
            layout="series",
            direction="internal",
            header_row=1,
        ),
    )


def _simultaneous_workbook(tmp_path: Path) -> Path:
    """Same-year `debt_t ↔ adj_t` (Excel circular ref), not a lag zipper."""
    return write_workbook(
        tmp_path / "a11_simultaneous.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=100",
                "B2": "=A2+B3",
                "C2": "=B2+C3",
                "B3": "=B2-A2",
                "C3": "=C2-B2",
            },
        },
    )


def test_lag_zipper_emits_fused_union_loop(tmp_path: Path) -> None:
    workbook = _zipper_workbook(tmp_path)
    modules = generate_inverted(workbook, _zipper_bindings())
    internals = modules["internals.py"]
    api = modules["api.py"]
    assert "cyclic formula-series" not in internals
    assert "eval_instance" not in internals
    assert "for t in range(" in internals
    assert internals.count("for t in range(") == 1
    pkg = load_package(modules, tmp_path, name="a11_zip")
    got = pkg.compute_debt()
    assert got == pytest.approx((100.0, 102.0, 104.04))
    assert "scan_debt_adjustment" in internals
    assert "internals.scan_debt_adjustment" in api


def test_lag_zipper_matches_formula_evaluator(tmp_path: Path) -> None:
    workbook = _zipper_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _zipper_bindings()), tmp_path, name="a11_num")
    graph = create_dependency_graph(
        workbook,
        ["Engine!A2", "Engine!B2", "Engine!C2", "Engine!B3", "Engine!C3"],
        load_values=True,
    )
    assert graph.cycle_report().has_must_cycles is False
    expected = FormulaEvaluator(graph).evaluate(
        ["Engine!A2", "Engine!B2", "Engine!C2", "Engine!B3", "Engine!C3"]
    )
    got = pkg.compute_debt()
    assert got == pytest.approx(
        (expected["Engine!A2"], expected["Engine!B2"], expected["Engine!C2"])
    )
    adj = pkg.internals.adjustment()
    assert adj == pytest.approx((expected["Engine!B3"], expected["Engine!C3"]))


def test_same_year_cell_cycle_still_fail_closed(tmp_path: Path) -> None:
    workbook = _simultaneous_workbook(tmp_path)
    with pytest.raises(InvertedTreeExportError, match="distance-zero residual"):
        generate_inverted(workbook, _zipper_bindings())


def _exec_scan(body: list[str], names: set[str]) -> tuple[tuple[object, ...], ...]:
    runtime = {
        "XlError": XlError,
        "as_measure": as_measure,
        "demand_instance": demand_instance,
        "eval_instance": eval_instance,
        "is_error": is_error,
        "live_measure": live_measure,
    }
    ns = {name: runtime[name] for name in names if name in runtime}
    exec("def scan():\n" + "\n".join(body), ns)
    return ns["scan"]()


def test_fused_loop_agrees_with_rung3_oracle(tmp_path: Path) -> None:
    catalog, deps, graph = inverted_graph_parts(_zipper_workbook(tmp_path), _zipper_bindings())
    scc = ("debt", "adjustment")
    fused, fused_used = emit_rung2_scc(scc, catalog=catalog, deps=deps, graph=graph)
    demand, demand_used = emit_rung3_scc(scc, catalog=catalog, deps=deps, graph=graph)
    assert _exec_scan(fused, fused_used) == _exec_scan(demand, demand_used)
