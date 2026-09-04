"""Statement-graph legality and fused-SCC classification."""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path

import pytest

from excel_grapher.core.address_keys import normalize_key as normalize_address
from excel_grapher.core.address_keys import parse_cell_coords
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.exporter.inverted_tree import catalog as catalog_mod
from excel_grapher.exporter.inverted_tree import deps as deps_mod
from excel_grapher.exporter.inverted_tree import schedule as schedule_mod
from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    KeyPoint,
    ScheduleIndex,
    SeriesCatalog,
    Statement,
    build_catalog,
    schedule_coord,
)
from excel_grapher.exporter.inverted_tree.deps import (
    DependenceEdge,
    collect_all_dependence_edges,
    identity_join_indices,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.schedule import (
    FusedRegion,
    plan_fused_scc,
    plan_scc,
    residual_body_order,
)
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    inverted_graph_parts,
    load_package,
    make_catalog,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a11_zipper import (
    _cross_sheet_zipper_bindings,
    _cross_sheet_zipper_workbook,
    _offset_zipper_bindings,
    _offset_zipper_workbook,
    _simultaneous_workbook,
    _vertical_zipper_bindings,
    _vertical_zipper_workbook,
    _zipper_bindings,
    _zipper_workbook,
)


def _catalog_from_edges(edges: Sequence[DependenceEdge]) -> SeriesCatalog:
    """Catalog whose schedule coords are spreadsheet columns (legacy synthetic tests)."""
    coord_of: dict[str, int] = {}
    for edge in edges:
        for addr in (edge.consumer_cell, edge.producer_cell):
            coord_of[normalize_address(addr)] = parse_cell_coords(addr)[2]
    return SeriesCatalog(
        series={},
        order=(),
        address_to_id={},
        schedule=ScheduleIndex(preferred={}, coord_of=coord_of, index_by_coord={}),
    )


def test_lagged_zipper_residual_orders_adjustment_before_debt() -> None:
    edges = (
        DependenceEdge("adjustment", "debt", "Engine!B3", "Engine!A2", 1),
        DependenceEdge("debt", "adjustment", "Engine!B2", "Engine!B3", 0),
        DependenceEdge("debt", "debt", "Engine!B2", "Engine!A2", 1),
    )
    catalog = _catalog_from_edges(edges)
    assert residual_body_order(("debt", "adjustment"), edges, catalog) == ("adjustment", "debt")


def test_distance_zero_cycle_is_illegal() -> None:
    edges = (
        DependenceEdge("debt", "adjustment", "Engine!B2", "Engine!B3", 0),
        DependenceEdge("adjustment", "debt", "Engine!B3", "Engine!B2", 0),
    )
    catalog = _catalog_from_edges(edges)
    with pytest.raises(InvertedTreeExportError, match="distance-zero residual"):
        residual_body_order(("debt", "adjustment"), edges, catalog)


def test_distance_zero_guarded_cycle_is_legal_and_demotes() -> None:
    edges = (
        DependenceEdge("debt", "adjustment", "Engine!B2", "Engine!B3", 0, guarded=True),
        DependenceEdge("adjustment", "debt", "Engine!B3", "Engine!B2", 0, guarded=False),
    )
    catalog = _catalog_from_edges(edges)
    schedule_mod.assert_distance_zero_legal(("debt", "adjustment"), edges, catalog)
    assert residual_body_order(("debt", "adjustment"), edges, catalog) is None


def test_identity_flip_has_no_single_body_order() -> None:
    edges = (
        DependenceEdge("x", "y", "Engine!A2", "Engine!A4", 0),
        DependenceEdge("y", "x", "Engine!B4", "Engine!B2", 0),
    )
    catalog = _catalog_from_edges(edges)
    assert residual_body_order(("x", "y"), edges, catalog) is None


def test_plan_fused_scc_for_corrected_zipper(tmp_path: Path) -> None:
    catalog, _deps, graph = inverted_graph_parts(_zipper_workbook(tmp_path), _zipper_bindings())
    plan = plan_fused_scc(("debt", "adjustment"), catalog=catalog, graph=graph)
    assert plan is not None
    assert plan.regions[-1].body_order == ("adjustment", "debt")
    assert plan.schedule == (0, 1, 2)
    assert plan.domain["debt"] == (0, 3)
    assert plan.domain["adjustment"] == (1, 3)
    assert plan.regions[-1].start == 1
    assert plan.regions == (
        FusedRegion(start=0, stop=1, body_order=("debt",)),
        FusedRegion(start=1, stop=3, body_order=("adjustment", "debt")),
    )


def test_plan_fused_scc_rejects_same_year_cycle(tmp_path: Path) -> None:
    catalog, _deps, graph = inverted_graph_parts(
        _simultaneous_workbook(tmp_path), _zipper_bindings()
    )
    with pytest.raises(InvertedTreeExportError, match="distance-zero residual"):
        plan_fused_scc(("debt", "adjustment"), catalog=catalog, graph=graph)


def test_schedule_coord_joins_resolved_time_period(tmp_path: Path) -> None:
    catalog, _deps, _graph = inverted_graph_parts(_zipper_workbook(tmp_path), _zipper_bindings())
    assert [point["TIME_PERIOD"] for point in catalog.get("debt").domain] == [2009, 2010, 2011]
    assert [point["TIME_PERIOD"] for point in catalog.get("adjustment").domain] == [2010, 2011]
    assert schedule_coord("Engine!A2", catalog) == 0
    assert schedule_coord("Engine!B2", catalog) == schedule_coord("Engine!B3", catalog) == 1
    assert schedule_coord("Engine!C2", catalog) == schedule_coord("Engine!C3", catalog) == 2


def test_vertical_zipper_lag_is_not_a_same_index_cycle(tmp_path: Path) -> None:
    catalog, _deps, graph = inverted_graph_parts(
        _vertical_zipper_workbook(tmp_path), _vertical_zipper_bindings()
    )
    assert [point["TIME_PERIOD"] for point in catalog.get("debt").domain] == [2009, 2010, 2011]
    assert schedule_coord("Engine!B2", catalog) == 1
    assert schedule_coord("Engine!B1", catalog) == 0
    edges = collect_all_dependence_edges(catalog, graph)
    lag = next(
        edge
        for edge in edges
        if edge.consumer_id == "debt"
        and edge.producer_id == "debt"
        and edge.consumer_cell == "Engine!B2"
        and edge.producer_cell == "Engine!B1"
    )
    assert lag.distance == 1
    plan = plan_fused_scc(("debt", "adjustment"), catalog=catalog, graph=graph)
    assert plan is not None
    assert plan.regions[-1].body_order == ("adjustment", "debt")
    assert plan.schedule == (0, 1, 2)
    assert plan.domain["debt"] == (0, 3)
    assert plan.domain["adjustment"] == (1, 3)


def test_offset_helper_block_is_same_index_not_look_ahead(tmp_path: Path) -> None:
    catalog, _deps, graph = inverted_graph_parts(
        _offset_zipper_workbook(tmp_path), _offset_zipper_bindings()
    )
    assert schedule_coord("Engine!B2", catalog) == schedule_coord("Engine!E3", catalog)
    edges = collect_all_dependence_edges(catalog, graph)
    same_year = next(
        edge
        for edge in edges
        if edge.consumer_cell == "Engine!B2" and edge.producer_cell == "Engine!E3"
    )
    assert same_year.distance == 0
    plan = plan_fused_scc(("debt", "adjustment"), catalog=catalog, graph=graph)
    assert plan is not None
    assert plan.regions[-1].body_order == ("adjustment", "debt")
    assert plan.domain["adjustment"] == (1, 3)


def test_cross_sheet_coords_are_not_column_subtraction(tmp_path: Path) -> None:
    catalog, _deps, graph = inverted_graph_parts(
        _cross_sheet_zipper_workbook(tmp_path), _cross_sheet_zipper_bindings()
    )
    assert schedule_coord("Engine!B2", catalog) == schedule_coord("Helper!C2", catalog) == 1
    edges = collect_all_dependence_edges(catalog, graph)
    same_year = next(
        edge
        for edge in edges
        if edge.consumer_cell == "Engine!B2" and edge.producer_cell == "Helper!C2"
    )
    assert same_year.distance == 0
    plan = plan_fused_scc(("debt", "adjustment"), catalog=catalog, graph=graph)
    assert plan is not None
    assert "eval_instance" not in str(plan)


def test_schedule_coord_does_not_rebuild_join_domain(tmp_path: Path, monkeypatch) -> None:
    catalog, _deps, _graph = inverted_graph_parts(_zipper_workbook(tmp_path), _zipper_bindings())
    assert schedule_coord("Engine!B2", catalog) == 1
    calls = {"n": 0}
    original = catalog_mod._ordered_domain

    def counting(catalog_arg, fields):
        calls["n"] += 1
        return original(catalog_arg, fields)

    monkeypatch.setattr(catalog_mod, "_ordered_domain", counting)
    for address in ("Engine!A2", "Engine!B2", "Engine!C2", "Engine!B3", "Engine!C3"):
        schedule_coord(address, catalog)
    assert calls["n"] == 0


def test_identity_join_does_not_rescan_join_domain(tmp_path: Path, monkeypatch) -> None:
    catalog, _deps, _graph = inverted_graph_parts(_zipper_workbook(tmp_path), _zipper_bindings())
    host = catalog.get("adjustment")
    producer = catalog.get("debt")
    assert identity_join_indices(host, producer, catalog) == (1, 2)
    domain_calls = {"n": 0}
    distance_calls = {"n": 0}
    original_domain = catalog_mod._ordered_domain
    original_distance = deps_mod._layout_distance

    def counting_domain(catalog_arg, fields):
        domain_calls["n"] += 1
        return original_domain(catalog_arg, fields)

    def counting_distance(*args, **kwargs):
        distance_calls["n"] += 1
        return original_distance(*args, **kwargs)

    monkeypatch.setattr(catalog_mod, "_ordered_domain", counting_domain)
    monkeypatch.setattr(deps_mod, "_layout_distance", counting_distance)
    assert identity_join_indices(host, producer, catalog) == (1, 2)
    assert domain_calls["n"] == 0
    assert distance_calls["n"] == 0


def test_plan_fused_scc_for_lookahead_reversed_direction(tmp_path: Path) -> None:
    wb = write_workbook(
        tmp_path / "lookahead_plan.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=B2*0.9+A3",
                "B2": "=C2*0.9+B3",
                "C2": "=100",
                "A3": "=B2*0.01",
                "B3": "=C2*0.01",
            },
        },
    )
    doc = bindings_document(
        series_entry("value", "Engine!A2:C2", layout="series", direction="output", header_row=1),
        series_entry("flow", "Engine!A3:B3", layout="series", direction="internal", header_row=1),
    )
    catalog, _deps, graph = inverted_graph_parts(wb, doc)
    plan = plan_fused_scc(("value", "flow"), catalog=catalog, graph=graph)
    assert plan is not None
    assert plan.direction == "reversed"
    assert plan.schedule == (2, 1, 0)
    assert plan.domain["value"] == (0, 3)
    assert plan.domain["flow"] == (1, 3)
    assert plan.regions == (
        FusedRegion(start=0, stop=1, body_order=("value",)),
        FusedRegion(start=1, stop=3, body_order=("flow", "value")),
    )


def test_empty_key_emits_on_expansion_order(tmp_path: Path) -> None:
    """`key: []` is positional; headers must not silently become the schedule.

    Schema `SeriesBindingLayoutKeyRules` still requires a non-empty key on
    non-scalar layouts, so this document is not schema-validated. The catalog
    and emit contract is expansion order once `key` is empty.
    """
    workbook = write_workbook(
        tmp_path / "keyless_expansion.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "D1": 2012,
                "A2": 1.0,
                "B2": 2.0,
                "C2": 3.0,
                "B3": "=A2",
                "C3": "=B2",
                "D3": "=C2",
            },
        },
    )
    document = bindings_document(
        series_entry(
            "values",
            "Engine!A2:C2",
            layout="series",
            direction="input",
            header_row=1,
            key=[],
        ),
        series_entry(
            "path",
            "Engine!B3:D3",
            layout="series",
            direction="output",
            header_row=1,
            key=[],
        ),
    )
    catalog = build_catalog(document, workbook=workbook)
    assert [point.as_mapping() for point in catalog.get("values").domain] == [{}, {}, {}]
    assert catalog.get("values").domain == (
        KeyPoint(()),
        KeyPoint(()),
        KeyPoint(()),
    )
    assert schedule_coord("Engine!A2", catalog) == 0
    assert schedule_coord("Engine!B3", catalog) == 0
    targets = all_series_targets(document, workbook=workbook)
    graph = create_dependency_graph(workbook, targets, load_values=True)
    with CodeGenerator(graph) as gen:
        modules = gen.generate_modules(
            targets,
            series_bindings=document,
            bindings_workbook=workbook,
            paradigm="inverted_tree",
        )
    pkg = load_package(modules, tmp_path, "keyless_expansion")
    assert pkg.compute_path(values=(1.0, 2.0, 3.0)) == pytest.approx((1.0, 2.0, 3.0))
    assert pkg.compute_path(values=(10.0, 20.0, 30.0)) == pytest.approx((10.0, 20.0, 30.0))


def test_partial_key_domain_fails_closed_in_schedule() -> None:
    cells = ("Engine!A2", "Engine!B2")
    domain = (KeyPoint((("TIME_PERIOD", 2009),)), KeyPoint(()))
    series = BoundSeries(
        series_id="values",
        layout="series",
        direction="input",
        cells=cells,
        key_fields=("TIME_PERIOD",),
        dtype="float",
        compute_name=None,
        raw={},
        domain=domain,
        statements=(Statement("values", "values", None, 0, 2, cells, domain),),
    )
    with pytest.raises(InvertedTreeExportError, match="Engine!B2"):
        make_catalog(
            series={"values": series},
            order=("values",),
            address_to_id={cell: "values" for cell in cells},
        )


def test_plan_fused_scc_rejects_mixed_signs(tmp_path: Path) -> None:
    wb = write_workbook(
        tmp_path / "mixed_plan.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=100",
                "B2": "=A2+C3",
                "C2": "=B2+10",
                "A3": "=B2*0.1",
                "B3": "=A2*0.1",
                "C3": "=10",
            },
        },
    )
    doc = bindings_document(
        series_entry("s1", "Engine!A2:C2", layout="series", direction="output", header_row=1),
        series_entry("s2", "Engine!A3:C3", layout="series", direction="internal", header_row=1),
    )
    catalog, _deps, graph = inverted_graph_parts(wb, doc)
    assert plan_fused_scc(("s1", "s2"), catalog=catalog, graph=graph) is None
    assert plan_scc(("s1", "s2"), catalog=catalog, graph=graph).rung == 3


def test_schedule_index_is_eager_catalog_field(tmp_path: Path) -> None:
    catalog, _deps, _graph = inverted_graph_parts(_zipper_workbook(tmp_path), _zipper_bindings())
    assert isinstance(catalog.schedule, ScheduleIndex)
    assert catalog.schedule.coord_of["Engine!B2"] == 1
    with pytest.raises(InvertedTreeExportError, match="no schedule coordinate"):
        schedule_coord("Engine!Z99", catalog)


def test_fused_plan_has_no_last_region_aliases(tmp_path: Path) -> None:
    catalog, _deps, graph = inverted_graph_parts(_zipper_workbook(tmp_path), _zipper_bindings())
    plan = plan_fused_scc(("debt", "adjustment"), catalog=catalog, graph=graph)
    assert plan is not None
    assert not hasattr(type(plan), "body_order")
    assert not hasattr(type(plan), "peel_stop")


def test_plan_scc_classifies_zipper_as_rung_2(tmp_path: Path) -> None:
    catalog, _deps, graph = inverted_graph_parts(_zipper_workbook(tmp_path), _zipper_bindings())
    choice = plan_scc(("debt", "adjustment"), catalog=catalog, graph=graph)
    assert choice.rung == 2
    assert choice.plan is not None
    assert choice.plan.regions[-1].body_order == ("adjustment", "debt")


def test_plan_scc_classifies_elementwise_as_rung_0(tmp_path: Path) -> None:
    wb = write_workbook(
        tmp_path / "rung0.xlsx",
        {"Engine": {"A1": 2009, "B1": 2010, "A2": "=1", "B2": "=A2+1"}},
    )
    doc = bindings_document(
        series_entry("src", "Engine!A2", layout="scalar", direction="input"),
        series_entry("out", "Engine!B2", layout="scalar", direction="output"),
    )
    catalog, _deps, graph = inverted_graph_parts(wb, doc)
    choice = plan_scc(("out",), catalog=catalog, graph=graph)
    assert choice.rung == 0
    assert choice.plan is None


def test_plan_scc_classifies_self_lag_as_rung_1(tmp_path: Path) -> None:
    wb = write_workbook(
        tmp_path / "rung1.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=1",
                "B2": "=A2*1.1",
                "C2": "=B2*1.1",
            },
        },
    )
    doc = bindings_document(
        series_entry("path", "Engine!A2:C2", layout="series", direction="output", header_row=1),
    )
    catalog, _deps, graph = inverted_graph_parts(wb, doc)
    choice = plan_scc(("path",), catalog=catalog, graph=graph)
    assert choice.rung == 1
    assert choice.plan is not None
    assert choice.plan.direction == "forward"


def test_statement_at_union_is_precomputed_coord_map_hit(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    cells = tuple(f"Engine!A{i}" for i in range(1, 6))
    domain = tuple(KeyPoint((("TIME_PERIOD", year),)) for year in range(5))
    series = BoundSeries(
        series_id="value",
        layout="series",
        direction="output",
        cells=cells,
        key_fields=("TIME_PERIOD",),
        dtype="float",
        compute_name=None,
        raw={},
        domain=domain,
        statements=(
            Statement("value__0", "value", None, 0, 1, cells[:1], domain[:1]),
            Statement("value__1", "value", None, 1, 5, cells[1:], domain[1:]),
        ),
    )
    catalog = make_catalog(
        series={"value": series},
        order=("value",),
        address_to_id={cell: "value" for cell in cells},
    )

    def boom(address: str, _catalog: SeriesCatalog) -> int:
        raise AssertionError(f"schedule coordinate walk should not run during lookup: {address}")

    monkeypatch.setattr(schedule_mod, "schedule_axis_coord", boom)
    assert schedule_mod._statement_at_union(catalog, "value", 0, 0) == "value__0"
    assert schedule_mod._statement_at_union(catalog, "value", 1, 3) == "value__1"
