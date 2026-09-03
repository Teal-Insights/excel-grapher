"""Rung 3 is the floor: backward recursion must emit, not fail closed."""

from __future__ import annotations

from pathlib import Path

import pytest
from fastpyxl.utils.cell import get_column_letter

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.schedule import plan_fused_scc
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)


def _horizontal_terminal_workbook(tmp_path: Path) -> Path:
    """`value_t = value_{t+1} * 0.9` with a terminal seed (common TV models)."""
    return write_workbook(
        tmp_path / "a19_terminal.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=B2*0.9",
                "B2": "=C2*0.9",
                "C2": "=100",
            },
        },
    )


def _horizontal_terminal_bindings() -> dict:
    return bindings_document(
        series_entry(
            "value",
            "Engine!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )


def _vertical_terminal_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a19_terminal_vertical.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "A2": 2010,
                "A3": 2011,
                "B1": "=B2*0.9",
                "B2": "=B3*0.9",
                "B3": "=100",
            },
        },
    )


def _vertical_terminal_bindings() -> dict:
    return bindings_document(
        series_entry(
            "value",
            "Engine!B1:B3",
            layout="series",
            direction="output",
            label_column="A",
        ),
    )


def _stride2_terminal_workbook(tmp_path: Path) -> Path:
    """Non-unit stride (distance -2) cannot use a scan; falls through to rung 3."""
    return write_workbook(
        tmp_path / "a19_stride2.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "D1": 2012,
                "A2": "=C2*0.5",
                "B2": "=D2*0.5",
                "C2": "=100",
                "D2": "=200",
            },
        },
    )


def _stride2_terminal_bindings() -> dict:
    return bindings_document(
        series_entry(
            "value",
            "Engine!A2:D2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )


@pytest.mark.parametrize(
    ("workbook_fn", "bindings_fn", "cells", "pkg_name"),
    [
        (
            _horizontal_terminal_workbook,
            _horizontal_terminal_bindings,
            ["Engine!A2", "Engine!B2", "Engine!C2"],
            "a19_h",
        ),
        (
            _vertical_terminal_workbook,
            _vertical_terminal_bindings,
            ["Engine!B1", "Engine!B2", "Engine!B3"],
            "a19_v",
        ),
    ],
    ids=["horizontal", "vertical"],
)
def test_backward_recursion_emits_reversed_scan_and_matches_evaluator(
    tmp_path: Path,
    workbook_fn,
    bindings_fn,
    cells: list[str],
    pkg_name: str,
) -> None:
    workbook = workbook_fn(tmp_path)
    modules = generate_inverted(workbook, bindings_fn())
    catalog, _deps, graph_bound = inverted_graph_parts(workbook, bindings_fn())
    plan = plan_fused_scc(("value",), catalog=catalog, graph=graph_bound)
    assert plan is not None
    assert plan.direction == "reversed"
    pkg = load_package(modules, tmp_path, name=pkg_name)
    graph = create_dependency_graph(workbook, cells, load_values=True)
    assert graph.cycle_report().has_must_cycles is False
    expected = FormulaEvaluator(graph).evaluate(cells)
    got = pkg.compute_value()
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx((81.0, 90.0, 100.0))


def test_irregular_recurrence_emits_rung3_and_matches_evaluator(tmp_path: Path) -> None:
    """Stride-2 look-ahead cannot use Rung 1; falls through to demand floor."""
    workbook = _stride2_terminal_workbook(tmp_path)
    doc = _stride2_terminal_bindings()
    cells = ["Engine!A2", "Engine!B2", "Engine!C2", "Engine!D2"]
    modules = generate_inverted(workbook, doc)
    catalog, _deps, graph_bound = inverted_graph_parts(workbook, doc)
    assert plan_fused_scc(("value",), catalog=catalog, graph=graph_bound) is None
    pkg = load_package(modules, tmp_path, name="a19_stride2")
    graph = create_dependency_graph(workbook, cells, load_values=True)
    expected = FormulaEvaluator(graph).evaluate(cells)
    got = pkg.compute_value()
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx((50.0, 100.0, 100.0, 200.0))


def test_backward_chain_large_n_matches_closed_form(tmp_path: Path) -> None:
    """Backward chain must not hit RecursionError at large N (gh #614, #615)."""
    n = 5000
    cells: dict[str, object] = {}
    for c in range(1, n + 1):
        cells[f"{get_column_letter(c)}1"] = 2000 + c
        if c < n:
            cells[f"{get_column_letter(c)}2"] = f"={get_column_letter(c + 1)}2*0.99"
        else:
            cells[f"{get_column_letter(c)}2"] = "=100"

    workbook = write_workbook(tmp_path / "a19_back_5000.xlsx", {"Engine": cells})
    last_col = get_column_letter(n)
    doc = bindings_document(
        series_entry(
            "value",
            f"Engine!A2:{last_col}2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )
    modules = generate_inverted(workbook, doc)
    pkg = load_package(modules, tmp_path, name="a19_back_5000")
    got = pkg.compute_value()
    assert len(got) == n
    expected = tuple(100.0 * (0.99 ** (n - 1 - i)) for i in range(n))
    assert got == pytest.approx(expected)
