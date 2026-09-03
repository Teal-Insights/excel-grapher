"""Layer A22 — rung-1 scans for shift-k / multi-lag self references (#617).

Rung 1 used to accept only a unit predecessor lag. `t-2`, `t-4`, and
`{t-1, t-2}` self-edges must stay on a fused forward scan (rung 2 body for a
single-statement SCC), match `FormulaEvaluator`, and keep code size independent
of series length.
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any, cast

import pytest
from fastpyxl.utils.cell import get_column_letter

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.ast_emit import emit_rung2_scc, emit_rung3_scc
from excel_grapher.exporter.inverted_tree.deps import (
    collect_series_edges,
    requires_demand_driven,
)
from excel_grapher.exporter.inverted_tree.runtime import (
    XlError,
    as_measure,
    demand_instance,
    eval_instance,
    is_error,
    live_measure,
)
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


def _exec_scan(body: list[str], names: set[str]) -> tuple[object, ...]:
    runtime = {
        "XlError": XlError,
        "as_measure": as_measure,
        "demand_instance": demand_instance,
        "eval_instance": eval_instance,
        "is_error": is_error,
        "live_measure": live_measure,
        "require_aligned": lambda *args: len(args[0]),
    }
    ns: dict[str, Any] = {name: runtime[name] for name in names if name in runtime}
    exec("def scan():\n" + "\n".join(body), ns)
    scan = cast(Callable[[], tuple[object, ...]], ns["scan"])
    return scan()


def _self_distances(series, catalog, graph) -> frozenset[int]:
    return frozenset(
        edge.distance
        for edge in collect_series_edges(series, catalog=catalog, graph=graph)
        if edge.producer_id == series.series_id
    )


def _stride_k_workbook(tmp_path: Path, n: int, lag: int, *, stem: str) -> Path:
    """`x_t = x_{t-lag} * 1.1` with `lag` leading seed members equal to 1."""
    cells: dict[str, object] = {}
    for index in range(n):
        col = get_column_letter(index + 1)
        cells[f"{col}1"] = 2009 + index
        if index < lag:
            cells[f"{col}2"] = "=1"
        else:
            pred = get_column_letter(index + 1 - lag)
            cells[f"{col}2"] = f"={pred}2*1.1"
    return write_workbook(tmp_path / f"{stem}.xlsx", {"Engine": cells})


def _stride_k_bindings(n: int) -> dict:
    last = get_column_letter(n)
    return bindings_document(
        series_entry(
            "path",
            f"Engine!A2:{last}2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )


def _dual_lag_workbook(tmp_path: Path, n: int, *, stem: str) -> Path:
    """`x_t = 0.4 * x_{t-2} + 0.6 * x_{t-1}` with two leading seeds."""
    cells: dict[str, object] = {}
    for index in range(n):
        col = get_column_letter(index + 1)
        cells[f"{col}1"] = 2009 + index
        if index < 2:
            cells[f"{col}2"] = "=1"
        else:
            lag2 = get_column_letter(index - 1)
            lag1 = get_column_letter(index)
            cells[f"{col}2"] = f"={lag2}2*0.4+{lag1}2*0.6"
    return write_workbook(tmp_path / f"{stem}.xlsx", {"Engine": cells})


def _expected_stride(n: int, lag: int, factor: float = 1.1) -> tuple[float, ...]:
    out = [1.0] * n
    for index in range(lag, n):
        out[index] = out[index - lag] * factor
    return tuple(out)


def _expected_dual(n: int) -> tuple[float, ...]:
    out = [1.0] * n
    for index in range(2, n):
        out[index] = 0.4 * out[index - 2] + 0.6 * out[index - 1]
    return tuple(out)


def _assert_fused_scan(internals: str, series_id: str = "path") -> None:
    assert "eval_instance" not in internals
    assert f"{series_id}_compute" not in internals
    assert "for t in range(" in internals
    assert "prior:" not in internals


@pytest.mark.parametrize(
    ("lag", "n", "stem"),
    [
        (2, 5, "a22_t2"),
        (4, 8, "a22_t4"),
    ],
    ids=["t-2", "t-4"],
)
def test_stride_k_self_lag_emits_fused_scan_and_matches_evaluator(
    tmp_path: Path, lag: int, n: int, stem: str
) -> None:
    workbook = _stride_k_workbook(tmp_path, n, lag, stem=stem)
    doc = _stride_k_bindings(n)
    catalog, deps, graph = inverted_graph_parts(workbook, doc)
    series = catalog.get("path")
    distances = _self_distances(series, catalog, graph)
    assert distances == frozenset({lag})
    assert requires_demand_driven(series, catalog=catalog, graph=graph) is False
    assert deps["path"].is_scan is True
    assert deps["path"].scan_direction == "forward"

    plan = plan_fused_scc(("path",), catalog=catalog, graph=graph)
    assert plan is not None
    assert plan.direction == "forward"
    assert plan.regions[-1].start == lag
    assert plan.schedule == tuple(range(n))

    modules = generate_inverted(workbook, doc)
    internals = modules["internals.py"]
    _assert_fused_scan(internals)
    assert f"path[t - {lag}]" in internals or f"path[t-{lag}]" in internals

    cells = [f"Engine!{get_column_letter(i + 1)}2" for i in range(n)]
    pkg = load_package(modules, tmp_path, name=stem)
    graph_full = create_dependency_graph(workbook, cells, load_values=True)
    expected = FormulaEvaluator(graph_full).evaluate(cells)
    got = pkg.compute_path()
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx(_expected_stride(n, lag))


def test_multi_lag_t1_t2_emits_fused_scan_and_matches_evaluator(tmp_path: Path) -> None:
    n = 5
    workbook = _dual_lag_workbook(tmp_path, n, stem="a22_dual")
    doc = _stride_k_bindings(n)
    catalog, deps, graph = inverted_graph_parts(workbook, doc)
    series = catalog.get("path")
    distances = _self_distances(series, catalog, graph)
    assert distances == frozenset({1, 2})
    assert requires_demand_driven(series, catalog=catalog, graph=graph) is False
    assert deps["path"].is_scan is True

    plan = plan_fused_scc(("path",), catalog=catalog, graph=graph)
    assert plan is not None
    assert plan.direction == "forward"
    assert plan.regions[-1].start == 2

    modules = generate_inverted(workbook, doc)
    _assert_fused_scan(modules["internals.py"])
    assert "path[t - 1]" in modules["internals.py"]
    assert "path[t - 2]" in modules["internals.py"]

    cells = [f"Engine!{get_column_letter(i + 1)}2" for i in range(n)]
    pkg = load_package(modules, tmp_path, name="a22_dual")
    graph_full = create_dependency_graph(workbook, cells, load_values=True)
    expected = FormulaEvaluator(graph_full).evaluate(cells)
    got = pkg.compute_path()
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx(_expected_dual(n))


def test_stride_k_fused_loop_agrees_with_rung3_oracle(tmp_path: Path) -> None:
    n, lag = 5, 2
    workbook = _stride_k_workbook(tmp_path, n, lag, stem="a22_oracle")
    catalog, deps, graph = inverted_graph_parts(workbook, _stride_k_bindings(n))
    scc = ("path",)
    fused, fused_used = emit_rung2_scc(scc, catalog=catalog, deps=deps, graph=graph)
    demand, demand_used = emit_rung3_scc(scc, catalog=catalog, deps=deps, graph=graph)
    fused_got = _exec_scan(fused, fused_used)
    demand_got = _exec_scan(demand, demand_used)
    if isinstance(fused_got[0], tuple):
        fused_got = fused_got[0]
    if isinstance(demand_got[0], tuple):
        demand_got = demand_got[0]
    assert fused_got == pytest.approx(demand_got)


@pytest.mark.parametrize("lag", [2, 4], ids=["t-2", "t-4"])
def test_stride_k_code_size_independent_of_series_length(tmp_path: Path, lag: int) -> None:
    small_n, large_n = 8, 24
    small_wb = _stride_k_workbook(tmp_path, small_n, lag, stem=f"a22_sz_s{lag}")
    large_wb = _stride_k_workbook(tmp_path, large_n, lag, stem=f"a22_sz_l{lag}")
    small = generate_inverted(small_wb, _stride_k_bindings(small_n))["internals.py"]
    large = generate_inverted(large_wb, _stride_k_bindings(large_n))["internals.py"]
    _assert_fused_scan(small)
    _assert_fused_scan(large)
    assert abs(small.count("\n") - large.count("\n")) <= 2
    assert "eval_instance" not in small
    assert "eval_instance" not in large


def test_dual_lag_code_size_independent_of_series_length(tmp_path: Path) -> None:
    small_n, large_n = 8, 24
    small_wb = _dual_lag_workbook(tmp_path, small_n, stem="a22_dual_s")
    large_wb = _dual_lag_workbook(tmp_path, large_n, stem="a22_dual_l")
    small = generate_inverted(small_wb, _stride_k_bindings(small_n))["internals.py"]
    large = generate_inverted(large_wb, _stride_k_bindings(large_n))["internals.py"]
    _assert_fused_scan(small)
    _assert_fused_scan(large)
    assert abs(small.count("\n") - large.count("\n")) <= 2
