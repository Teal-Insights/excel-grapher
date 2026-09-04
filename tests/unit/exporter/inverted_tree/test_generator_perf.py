"""Generator-side inverted-tree hot spots (#618, #636, #653)."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from pathlib import Path

import pytest

from excel_grapher.core import address_keys as address_keys_mod
from excel_grapher.exporter.inverted_tree import deps as deps_mod
from excel_grapher.exporter.inverted_tree import schedule as schedule_mod
from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    KeyPoint,
    SeriesCatalog,
    Statement,
    build_catalog,
    schedule_coord,
)
from excel_grapher.exporter.inverted_tree.deps import DependenceEdge
from excel_grapher.exporter.inverted_tree.schedule import plan_fused_scc
from tests.unit.exporter.inverted_tree.helpers import (
    generate_inverted,
    make_catalog,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a1_leaf_closure import (
    _a1_bindings,
    _a1_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a11_zipper import (
    _zipper_bindings,
    _zipper_workbook,
)

_CATALOG_CONSTANT_CELLS = 45_000


def _constant_series_bindings(n: int = _CATALOG_CONSTANT_CELLS) -> dict:
    """Minimal catalog bindings: one constant series, no declared key.

    Schema validation requires a non-empty `key` on `layout: series`.
    `build_catalog` accepts `key: []` and treats expansion order as the
    schedule, which is the 45k-cell case in #636.
    """
    return {
        "series": [
            {
                "id": "consts",
                "data_range": f"Sheet1!A1:A{n}",
                "layout": "series",
                "constant": {},
                "key": [],
                "structure": {"measure": {"dtype": "float"}},
            }
        ]
    }


def _count_normalize_key_calls(monkeypatch: pytest.MonkeyPatch) -> dict[str, int]:
    calls = {"n": 0}
    original = address_keys_mod.normalize_key

    def counting(key: str) -> str:
        calls["n"] += 1
        return original(key)

    monkeypatch.setattr(address_keys_mod, "normalize_key", counting)
    return calls


def _count_fused_plan_ops(monkeypatch: pytest.MonkeyPatch) -> dict[str, int]:
    """Count bucketed edge walks during `plan_fused_scc` (#618, #653)."""
    counts = {"index_keys": 0, "index_edges": 0, "buckets": 0, "zero_distance": 0}
    original_key = schedule_mod._index_region_key
    original_bucket = schedule_mod._bucket_edges_by_consumer_coord
    original_zero = schedule_mod._zero_distance_edges

    def counting_key(
        scc: tuple[str, ...],
        *,
        catalog: SeriesCatalog,
        domain: Mapping[str, tuple[int, int]],
        index_edges: Sequence[DependenceEdge],
        union_t: int,
        index: int,
    ) -> object:
        counts["index_keys"] += 1
        counts["index_edges"] += len(index_edges)
        return original_key(
            scc,
            catalog=catalog,
            domain=domain,
            index_edges=index_edges,
            union_t=union_t,
            index=index,
        )

    def counting_bucket(
        edges: Sequence[DependenceEdge],
        catalog: SeriesCatalog,
        *,
        partition: tuple[object, ...] | None = None,
    ) -> dict[int, list[DependenceEdge]]:
        counts["buckets"] += 1
        return original_bucket(edges, catalog, partition=partition)

    def counting_zero(
        scc: tuple[str, ...],
        edges: Sequence[DependenceEdge],
    ) -> list[DependenceEdge]:
        counts["zero_distance"] += 1
        return original_zero(scc, edges)

    monkeypatch.setattr(schedule_mod, "_index_region_key", counting_key)
    monkeypatch.setattr(schedule_mod, "_bucket_edges_by_consumer_coord", counting_bucket)
    monkeypatch.setattr(schedule_mod, "_zero_distance_edges", counting_zero)
    return counts


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


def test_build_catalog_normalizes_each_cell_once(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    workbook = write_workbook(tmp_path / "const.xlsx", {"Sheet1": {"A1": 1}})
    calls = _count_normalize_key_calls(monkeypatch)
    catalog = build_catalog(_constant_series_bindings(), workbook=workbook)
    assert len(catalog.get("consts").cells) == _CATALOG_CONSTANT_CELLS
    assert calls["n"] == _CATALOG_CONSTANT_CELLS, calls["n"]


def test_build_catalog_normalize_key_scales_linearly(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    workbook = write_workbook(tmp_path / "const.xlsx", {"Sheet1": {"A1": 1}})
    n_small, n_large = 1_000, 4_000
    calls = _count_normalize_key_calls(monkeypatch)
    build_catalog(_constant_series_bindings(n_small), workbook=workbook)
    small = calls["n"]
    calls["n"] = 0
    build_catalog(_constant_series_bindings(n_large), workbook=workbook)
    large = calls["n"]
    assert small == n_small, small
    assert large == n_large, large
    assert large == small * (n_large // n_small)


def test_schedule_coord_does_not_renormalize_catalog_cells(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    workbook = write_workbook(tmp_path / "const.xlsx", {"Sheet1": {"A1": 1}})
    catalog = build_catalog(_constant_series_bindings(n=8), workbook=workbook)
    calls = _count_normalize_key_calls(monkeypatch)
    for cell in catalog.get("consts").cells:
        schedule_coord(cell, catalog)
        assert catalog.series_for(cell) is catalog.get("consts")
        assert catalog.get("consts").index_of(cell) is not None
    assert calls["n"] == 0, calls["n"]


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


def test_plan_fused_scc_examines_each_edge_once_per_bucket(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    n = 5_000
    catalog, edges = _synthetic_zipper(n)
    assert len(catalog.get("debt").statements) == 2
    assert len(edges) == 3 * (n - 1)
    counts = _count_fused_plan_ops(monkeypatch)
    plan = plan_fused_scc(("debt", "adjustment"), catalog=catalog, edges=edges)
    assert plan is not None
    assert plan.regions[-1].body_order == ("adjustment", "debt")
    assert len(plan.schedule) == n
    # One bucket pass; each same-partition edge is examined once across
    # union indices. Scanning the full list per index is T × |E|.
    assert counts["buckets"] == 1, counts
    assert counts["index_keys"] == n, counts
    assert counts["index_edges"] == len(edges), counts
    assert counts["zero_distance"] == 2, counts


def test_plan_fused_scc_edge_exams_scale_with_edges(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    n_small, n_large = 1_000, 4_000
    counts = _count_fused_plan_ops(monkeypatch)
    catalog_small, edges_small = _synthetic_zipper(n_small)
    plan_small = plan_fused_scc(("debt", "adjustment"), catalog=catalog_small, edges=edges_small)
    small = dict(counts)
    counts["index_keys"] = counts["index_edges"] = counts["buckets"] = counts["zero_distance"] = 0
    catalog_large, edges_large = _synthetic_zipper(n_large)
    plan_large = plan_fused_scc(("debt", "adjustment"), catalog=catalog_large, edges=edges_large)
    assert plan_small is not None and plan_large is not None
    assert len(plan_large.schedule) == n_large
    assert small["index_keys"] == n_small, small
    assert counts["index_keys"] == n_large, counts
    assert small["index_edges"] == len(edges_small), small
    assert counts["index_edges"] == len(edges_large), counts
    assert small["zero_distance"] == counts["zero_distance"] == 2
    assert small["buckets"] == counts["buckets"] == 1
