"""Layer A17 — overlapping TIME_PERIOD catalogs get a per-call `take` (#607)."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)


def _time_series(
    series_id: str, data_range: str, *, output: bool = False, internal: bool = False
) -> dict:
    if output:
        direction = "output"
    elif internal:
        direction = "internal"
    else:
        direction = "input"
    return series_entry(
        series_id,
        data_range,
        layout="series",
        direction=direction,
        header_row=1,
    )


def _overlap_workbook(tmp_path: Path) -> Path:
    """GDP is `B:D` (3); revenue / ratio / output are `C:D` (2)."""
    return write_workbook(
        tmp_path / "a17_overlap.xlsx",
        {
            "Engine": {
                "B1": 2009,
                "C1": 2010,
                "D1": 2011,
                "B2": 100,
                "C2": 110,
                "D2": 121,
                "C3": 10,
                "D3": 12,
                "C4": "=C2/B2",
                "D4": "=D2/C2",
                "C5": "=C3/C2",
                "D5": "=D3/D2",
                "C6": "=C4+C5",
                "D6": "=D4+D5",
            },
        },
    )


def _overlap_bindings() -> dict:
    return bindings_document(
        _time_series("gdp", "Engine!B2:D2"),
        _time_series("revenue", "Engine!C3:D3"),
        _time_series("gdp_growth", "Engine!C4:D4", internal=True),
        _time_series("revenue_pct_gdp", "Engine!C5:D5", internal=True),
        _time_series("result", "Engine!C6:D6", output=True),
    )


def test_revenue_pct_gdp_index_map_is_overlapping_window(tmp_path: Path) -> None:
    catalog, deps, _graph = inverted_graph_parts(_overlap_workbook(tmp_path), _overlap_bindings())
    assert deps["revenue_pct_gdp"].index_maps["gdp"] == (1, 2)
    assert deps["revenue_pct_gdp"].index_maps["revenue"] == (0, 1)


def test_overlap_call_site_takes_gdp_window(tmp_path: Path) -> None:
    workbook = _overlap_workbook(tmp_path)
    modules = generate_inverted(workbook, _overlap_bindings())
    api = modules["api.py"]
    assert "take(gdp, range(1, 3))" in api
    assert "internals.gdp_growth(gdp)" in api
    assert "internals.revenue_pct_gdp(take(gdp, range(1, 3)), revenue)" in api
    internals = modules["internals.py"]
    assert "require_length(gdp, 2)" in internals
    assert "require_aligned(revenue)" in internals
    pkg = load_package(modules, tmp_path, name="a17_overlap")
    got = pkg.compute_result(gdp=(100.0, 110.0, 121.0), revenue=(10.0, 12.0))
    assert got == pytest.approx((110 / 100 + 10 / 110, 121 / 110 + 12 / 121))


def test_overlap_matches_formula_evaluator(tmp_path: Path) -> None:
    workbook = _overlap_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _overlap_bindings()), tmp_path, name="a17_eval")
    graph = create_dependency_graph(workbook, ["Engine!C6", "Engine!D6"], load_values=True)
    expected = FormulaEvaluator(graph).evaluate(["Engine!C6", "Engine!D6"])
    got = pkg.compute_result(gdp=(100.0, 110.0, 121.0), revenue=(10.0, 12.0))
    assert got == pytest.approx((expected["Engine!C6"], expected["Engine!D6"]))


def _nested_compute_body(source: str, helper: str) -> str:
    start = source.index(f"def {helper}_compute(")
    rest = source[start:]
    nxt = rest.find("\n    def ", len(f"def {helper}_compute("))
    return rest if nxt < 0 else rest[:nxt]


def test_overlap_rung3_indexes_taken_window(tmp_path: Path) -> None:
    """Rung-3 helpers subscript the taken gdp window, not the catalog (#633)."""
    workbook = _overlap_workbook(tmp_path)
    modules = generate_inverted(workbook, _overlap_bindings(), force_rung=3)
    body = _nested_compute_body(modules["internals.py"], "revenue_pct_gdp")
    assert "gdp[i + 1]" not in body
    assert "gdp[i]" in body
    pkg = load_package(modules, tmp_path, name="a17_r3")
    graph = create_dependency_graph(workbook, ["Engine!C6", "Engine!D6"], load_values=True)
    expected = FormulaEvaluator(graph).evaluate(["Engine!C6", "Engine!D6"])
    got = pkg.compute_result(gdp=(100.0, 110.0, 121.0), revenue=(10.0, 12.0))
    assert got == pytest.approx((expected["Engine!C6"], expected["Engine!D6"]))
