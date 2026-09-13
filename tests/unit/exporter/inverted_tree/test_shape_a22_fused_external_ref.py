"""Layer A22 — fused refs to cells off the SCC union schedule (#623).

`coord_to_t` only maps SCC-member coordinates. An external seed (or any
producer whose join-key is not on that union) must emit `live_measure` in
the producer's catalog index space — never `KeyError`. Reversed #614
schedules must use the opposite step so subscripts stay in catalog order.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.schedule import plan_scc
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_forced_rung_packages,
    load_package,
    series_entry,
    write_workbook,
)


def _forward_seed_workbook(tmp_path: Path) -> Path:
    """Issue #623 MCVE: zipper (2010–2011) reads seed at 2009 (coord 0)."""
    return write_workbook(
        tmp_path / "a22_forward_seed.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": 100,
                "B2": "=A2+B3",
                "C2": "=B2+C3",
                "B3": "=A2*0.02",
                "C3": "=B2*0.02",
            },
        },
    )


def _forward_seed_bindings() -> dict:
    return bindings_document(
        series_entry("seed", "Engine!A2:A2", layout="series", direction="input", header_row=1),
        series_entry("debt", "Engine!B2:C2", layout="series", direction="output", header_row=1),
        series_entry("adj", "Engine!B3:C3", layout="series", direction="internal", header_row=1),
    )


def _reversed_seed_workbook(tmp_path: Path) -> Path:
    """Look-ahead zipper (2009–2010) reads a terminal seed at 2011 (off union)."""
    return write_workbook(
        tmp_path / "a22_reversed_seed.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=B2*0.9+A3",
                "B2": "=C2*0.9+B3",
                "C2": 100,
                "A3": "=B2*0.01",
                "B3": "=C2*0.01",
            },
        },
    )


def _reversed_seed_bindings() -> dict:
    return bindings_document(
        series_entry("value", "Engine!A2:B2", layout="series", direction="output", header_row=1),
        series_entry("flow", "Engine!A3:B3", layout="series", direction="internal", header_row=1),
        series_entry("seed", "Engine!C2:C2", layout="series", direction="input", header_row=1),
    )


def _reversed_rate_workbook(tmp_path: Path) -> Path:
    """Reversed zipper plus a multi-period rate series in catalog order."""
    return write_workbook(
        tmp_path / "a22_reversed_rate.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=B2*0.9+A3",
                "B2": "=C2*0.9+B3",
                "C2": "=100",
                "A3": "=B2*A4",
                "B3": "=C2*B4",
                "A4": 0.01,
                "B4": 0.02,
                "C4": 0.03,
            },
        },
    )


def _reversed_rate_bindings() -> dict:
    return bindings_document(
        series_entry("value", "Engine!A2:C2", layout="series", direction="output", header_row=1),
        series_entry("flow", "Engine!A3:B3", layout="series", direction="internal", header_row=1),
        series_entry("rate", "Engine!A4:C4", layout="series", direction="input", header_row=1),
    )


def test_forward_off_union_seed_does_not_raise_keyerror(tmp_path: Path) -> None:
    workbook = _forward_seed_workbook(tmp_path)
    doc = _forward_seed_bindings()
    catalog, _deps, graph = inverted_graph_parts(workbook, doc)
    scc = ("debt", "adj")
    choice = plan_scc(scc, catalog=catalog, graph=graph)
    assert choice.rung == 2
    assert choice.plan is not None
    assert choice.plan.schedule == (1, 2)
    assert 0 not in choice.plan.coord_to_t

    fused, demand = load_forced_rung_packages(workbook, doc, tmp_path, "a22_fwd_seed_or")
    assert dict(
        fused.compute_debt(
            seed=fused.data.Seed.from_nested(domain=fused.data.SEED_DOMAIN, values=(100.0,))
        ).items()
    ) == pytest.approx(
        dict(
            demand.compute_debt(
                seed=demand.data.Seed.from_nested(domain=demand.data.SEED_DOMAIN, values=(100.0,))
            ).items()
        )
    )


def test_forward_off_union_seed_matches_evaluator(tmp_path: Path) -> None:
    workbook = _forward_seed_workbook(tmp_path)
    doc = _forward_seed_bindings()
    catalog, _deps, graph_bound = inverted_graph_parts(workbook, doc)
    assert plan_scc(("debt", "adj"), catalog=catalog, graph=graph_bound).rung == 2
    modules = generate_inverted(workbook, doc)
    internals = modules["internals.py"]
    assert "debt[time_period - 1]" in internals
    assert "KeyError" not in internals

    pkg = load_package(modules, tmp_path, name="a22_fwd_seed")
    cells = ["Engine!B2", "Engine!C2", "Engine!B3", "Engine!C3"]
    graph = create_dependency_graph(workbook, cells, load_values=True)
    expected = FormulaEvaluator(graph).evaluate(cells)
    got = pkg.compute_debt(
        seed=pkg.data.Seed.from_nested(domain=pkg.data.SEED_DOMAIN, values=(100.0,))
    )
    assert [value for _, value in got.items()] == pytest.approx(
        tuple(expected[cell] for cell in cells[:2])
    )
    assert [value for _, value in got.items()] == pytest.approx((102.0, 104.04))
    adj = pkg.internals.scan_debt(
        seed=pkg.data.Seed.from_nested(domain=pkg.data.SEED_DOMAIN, values=(100.0,))
    ).adj
    assert [value for _, value in adj.items()] == pytest.approx(
        tuple(expected[cell] for cell in cells[2:])
    )


def test_reversed_off_union_seed_matches_evaluator(tmp_path: Path) -> None:
    workbook = _reversed_seed_workbook(tmp_path)
    doc = _reversed_seed_bindings()
    catalog, _deps, graph = inverted_graph_parts(workbook, doc)
    scc = ("value", "flow")
    choice = plan_scc(scc, catalog=catalog, graph=graph)
    assert choice.rung == 2
    assert choice.plan is not None
    assert choice.plan.direction == "reversed"
    assert 2 not in choice.plan.coord_to_t

    fused, demand = load_forced_rung_packages(workbook, doc, tmp_path, "a22_rev_seed")
    assert dict(
        fused.compute_value(
            seed=fused.data.Seed.from_nested(domain=fused.data.SEED_DOMAIN, values=(100.0,))
        ).items()
    ) == pytest.approx(
        dict(
            demand.compute_value(
                seed=demand.data.Seed.from_nested(domain=demand.data.SEED_DOMAIN, values=(100.0,))
            ).items()
        )
    )
    pkg = fused
    cells = ["Engine!A2", "Engine!B2", "Engine!A3", "Engine!B3"]
    graph_full = create_dependency_graph(workbook, cells, load_values=True)
    expected = FormulaEvaluator(graph_full).evaluate(cells)
    got = pkg.compute_value(
        seed=pkg.data.Seed.from_nested(domain=pkg.data.SEED_DOMAIN, values=(100.0,))
    )
    assert [value for _, value in got.items()] == pytest.approx(
        tuple(expected[cell] for cell in cells[:2])
    )


def test_reversed_aligned_external_rate_uses_catalog_index(tmp_path: Path) -> None:
    """`idx - host_union` is wrong on a reversed schedule; catalog step must flip."""
    workbook = _reversed_rate_workbook(tmp_path)
    doc = _reversed_rate_bindings()
    catalog, _deps, graph = inverted_graph_parts(workbook, doc)
    scc = ("value", "flow")
    choice = plan_scc(scc, catalog=catalog, graph=graph)
    assert choice.rung == 2
    assert choice.plan is not None
    assert choice.plan.direction == "reversed"

    fused, demand = load_forced_rung_packages(workbook, doc, tmp_path, "a22_rev_rate")
    rate = fused.data.Rate.from_nested(domain=fused.data.RATE_DOMAIN, values=(0.01, 0.02, 0.03))
    demand_rate = demand.data.Rate.from_nested(
        domain=demand.data.RATE_DOMAIN, values=(0.01, 0.02, 0.03)
    )
    assert dict(fused.compute_value(rate=rate).items()) == pytest.approx(
        dict(demand.compute_value(rate=demand_rate).items())
    )
    pkg = fused
    cells = ["Engine!A2", "Engine!B2", "Engine!C2", "Engine!A3", "Engine!B3"]
    graph_full = create_dependency_graph(workbook, cells, load_values=True)
    expected = FormulaEvaluator(graph_full).evaluate(cells)
    got = pkg.compute_value(rate=rate)
    assert [value for _, value in got.items()] == pytest.approx(
        tuple(expected[cell] for cell in cells[:3])
    )
    flow = pkg.internals.scan_value(rate=rate).flow
    assert [value for _, value in flow.items()] == pytest.approx(
        tuple(expected[cell] for cell in cells[3:])
    )
    assert [value for _, value in got.items()] == pytest.approx((83.72, 92.0, 100.0))
