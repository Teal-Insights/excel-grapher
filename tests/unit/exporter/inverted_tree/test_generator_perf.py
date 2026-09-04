"""Generator-side inverted-tree hot spots (#618)."""

from __future__ import annotations

import time
from collections.abc import Sequence
from pathlib import Path

import pytest

from excel_grapher.exporter.inverted_tree import deps as deps_mod
from excel_grapher.exporter.inverted_tree import schedule as schedule_mod
from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    KeyPoint,
    SeriesCatalog,
    Statement,
)
from excel_grapher.exporter.inverted_tree.deps import DependenceEdge
from excel_grapher.exporter.inverted_tree.schedule import plan_fused_scc
from tests.unit.exporter.inverted_tree.helpers import generate_inverted, make_catalog
from tests.unit.exporter.inverted_tree.test_shape_a1_leaf_closure import (
    _a1_bindings,
    _a1_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a11_zipper import (
    _zipper_bindings,
    _zipper_workbook,
)


def test_generate_modules_walks_each_series_ast_once(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    walks: dict[str, int] = {}
    original = deps_mod.collect_series_edges

    def counting(
        series: BoundSeries,
        *,
        catalog: SeriesCatalog,
        graph: object,
    ) -> list[DependenceEdge]:
        walks[series.series_id] = walks.get(series.series_id, 0) + 1
        return original(series, catalog=catalog, graph=graph)

    monkeypatch.setattr(deps_mod, "collect_series_edges", counting)
    monkeypatch.setattr(schedule_mod, "collect_series_edges", counting)
    generate_inverted(_a1_workbook(tmp_path), _a1_bindings())
    generate_inverted(_zipper_workbook(tmp_path), _zipper_bindings())
    assert walks, "expected collect_series_edges to run during generate_modules"
    assert all(count == 1 for count in walks.values()), walks


def _synthetic_series(
    series_id: str,
    cells: tuple[str, ...],
    years: Sequence[int],
    *,
    direction: str = "internal",
    peeled_seed: bool = False,
) -> BoundSeries:
    domain = tuple(KeyPoint((("TIME_PERIOD", year),)) for year in years)
    n = len(cells)
    if peeled_seed and n >= 2:
        statements = (
            Statement(f"{series_id}__0", series_id, None, 0, 1, cells[:1], domain[:1]),
            Statement(f"{series_id}__1", series_id, None, 1, n, cells[1:], domain[1:]),
        )
    else:
        statements = (Statement(series_id, series_id, None, 0, n, cells, domain),)
    return BoundSeries(
        series_id=series_id,
        layout="series",
        direction=direction,
        cells=cells,
        key_fields=("TIME_PERIOD",),
        dtype="float",
        compute_name=None,
        raw={},
        domain=domain,
        statements=statements,
    )


def _synthetic_zipper(n: int) -> tuple[SeriesCatalog, tuple[DependenceEdge, ...]]:
    """2-series zipper: `debt_t = debt_{t-1} + adj_t`, `adj_t = debt_{t-1} * r`.

    Debt is a peeled-seed two-statement series so fused planning hits
    `_statement_at_union` instead of the single-statement short-circuit.
    """
    debt_cells = tuple(f"Engine!A{i}" for i in range(1, n + 1))
    adj_cells = tuple(f"Engine!B{i}" for i in range(2, n + 1))
    debt = _synthetic_series("debt", debt_cells, range(n), direction="output", peeled_seed=True)
    adj = _synthetic_series("adjustment", adj_cells, range(1, n), direction="internal")
    catalog = make_catalog(
        series={"debt": debt, "adjustment": adj},
        order=("debt", "adjustment"),
        address_to_id={
            **{cell: "debt" for cell in debt_cells},
            **{cell: "adjustment" for cell in adj_cells},
        },
    )
    edges: list[DependenceEdge] = []
    for t in range(1, n):
        debt_cell = debt_cells[t]
        prev_debt = debt_cells[t - 1]
        adj_cell = adj_cells[t - 1]
        edges.append(
            DependenceEdge(
                consumer_id="debt",
                producer_id="debt",
                consumer_cell=debt_cell,
                producer_cell=prev_debt,
                distance=1,
                access="shift",
            )
        )
        edges.append(
            DependenceEdge(
                consumer_id="debt",
                producer_id="adjustment",
                consumer_cell=debt_cell,
                producer_cell=adj_cell,
                distance=0,
                access="identity",
            )
        )
        edges.append(
            DependenceEdge(
                consumer_id="adjustment",
                producer_id="debt",
                consumer_cell=adj_cell,
                producer_cell=prev_debt,
                distance=1,
                access="shift",
            )
        )
    return catalog, tuple(edges)


def test_plan_fused_scc_5k_periods_under_a_second() -> None:
    catalog, edges = _synthetic_zipper(5_000)
    assert len(catalog.get("debt").statements) == 2
    assert len(edges) == 14_997
    start = time.perf_counter()
    plan = plan_fused_scc(("debt", "adjustment"), catalog=catalog, edges=edges)
    elapsed = time.perf_counter() - start
    assert plan is not None
    assert plan.regions[-1].body_order == ("adjustment", "debt")
    assert len(plan.schedule) == 5_000
    assert elapsed < 1.0, f"planning took {elapsed:.3f}s"


def test_plan_fused_scc_two_statement_scales_linearly() -> None:
    catalog_5k, edges_5k = _synthetic_zipper(5_000)
    catalog_20k, edges_20k = _synthetic_zipper(20_000)
    start = time.perf_counter()
    plan_5k = plan_fused_scc(("debt", "adjustment"), catalog=catalog_5k, edges=edges_5k)
    elapsed_5k = time.perf_counter() - start
    start = time.perf_counter()
    plan_20k = plan_fused_scc(("debt", "adjustment"), catalog=catalog_20k, edges=edges_20k)
    elapsed_20k = time.perf_counter() - start
    assert plan_5k is not None and plan_20k is not None
    assert len(plan_20k.schedule) == 20_000
    assert elapsed_20k <= elapsed_5k * 5 + 0.25, (
        f"planning 20k took {elapsed_20k:.3f}s vs 5k {elapsed_5k:.3f}s"
    )
