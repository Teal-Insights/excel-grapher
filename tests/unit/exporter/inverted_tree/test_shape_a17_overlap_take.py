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
    internals = modules["internals.py"]
    assert "gdp[time_period]" in internals
    assert "gdp[time_period - 1]" in internals
    assert "revenue[time_period]" in internals
    assert "GDP_SCHEMA.validate(gdp)" in internals
    pkg = load_package(modules, tmp_path, name="a17_overlap")
    got = pkg.compute_result(
        gdp=pkg.data.Gdp.from_nested(domain=pkg.data.GDP_DOMAIN, values=(100.0, 110.0, 121.0)),
        revenue=pkg.data.Revenue.from_nested(domain=pkg.data.REVENUE_DOMAIN, values=(10.0, 12.0)),
    )
    assert [got[year] for year in (2010, 2011)] == pytest.approx(
        (110 / 100 + 10 / 110, 121 / 110 + 12 / 121)
    )


def test_overlap_matches_formula_evaluator(tmp_path: Path) -> None:
    workbook = _overlap_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _overlap_bindings()), tmp_path, name="a17_eval")
    graph = create_dependency_graph(workbook, ["Engine!C6", "Engine!D6"], load_values=True)
    expected = FormulaEvaluator(graph).evaluate(["Engine!C6", "Engine!D6"])
    got = pkg.compute_result(
        gdp=pkg.data.Gdp.from_nested(domain=pkg.data.GDP_DOMAIN, values=(100.0, 110.0, 121.0)),
        revenue=pkg.data.Revenue.from_nested(domain=pkg.data.REVENUE_DOMAIN, values=(10.0, 12.0)),
    )
    assert [got[year] for year in (2010, 2011)] == pytest.approx(
        (expected["Engine!C6"], expected["Engine!D6"])
    )


def _nested_compute_body(source: str, helper: str) -> str:
    start = source.index(f"def {helper}_compute(")
    rest = source[start:]
    nxt = rest.find("\n    def ", len(f"def {helper}_compute("))
    return rest if nxt < 0 else rest[:nxt]


def test_overlap_rung3_indexes_taken_window(tmp_path: Path) -> None:
    """Rung-3 helpers subscript the taken gdp window, not the catalog (#633)."""
    workbook = _overlap_workbook(tmp_path)
    modules = generate_inverted(workbook, _overlap_bindings(), force_rung=3)
    body = modules["internals.py"]
    assert "gdp[time_period]" in body
    assert "revenue[time_period]" in body
    pkg = load_package(modules, tmp_path, name="a17_r3")
    graph = create_dependency_graph(workbook, ["Engine!C6", "Engine!D6"], load_values=True)
    expected = FormulaEvaluator(graph).evaluate(["Engine!C6", "Engine!D6"])
    got = pkg.compute_result(
        gdp=pkg.data.Gdp.from_nested(domain=pkg.data.GDP_DOMAIN, values=(100.0, 110.0, 121.0)),
        revenue=pkg.data.Revenue.from_nested(domain=pkg.data.REVENUE_DOMAIN, values=(10.0, 12.0)),
    )
    assert [got[year] for year in (2010, 2011)] == pytest.approx(
        (expected["Engine!C6"], expected["Engine!D6"])
    )
