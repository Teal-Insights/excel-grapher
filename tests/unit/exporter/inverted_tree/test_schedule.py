"""Statement-graph legality and schedule coordinates."""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path

import pytest

from excel_grapher.core.address_keys import normalize_key as normalize_address
from excel_grapher.core.address_keys import parse_cell_coords
from excel_grapher.exporter.inverted_tree import catalog as catalog_mod
from excel_grapher.exporter.inverted_tree import deps as deps_mod
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
    collect_catalog_edges,
    identity_join_indices,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.schedule import assert_distance_zero_legal
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    inverted_graph_parts,
    make_catalog,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a11_zipper import (
    _cross_sheet_zipper_bindings,
    _cross_sheet_zipper_workbook,
    _offset_zipper_bindings,
    _offset_zipper_workbook,
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


def test_distance_zero_cycle_is_illegal() -> None:
    edges = (
        DependenceEdge("debt", "adjustment", "Engine!B2", "Engine!B3", 0),
        DependenceEdge("adjustment", "debt", "Engine!B3", "Engine!B2", 0),
    )
    catalog = _catalog_from_edges(edges)
    with pytest.raises(InvertedTreeExportError, match="distance-zero residual"):
        assert_distance_zero_legal(("debt", "adjustment"), edges, catalog)


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
    edges = collect_catalog_edges(catalog, graph).edges
    lag = next(
        edge
        for edge in edges
        if edge.consumer_id == "debt"
        and edge.producer_id == "debt"
        and edge.consumer_cell == "Engine!B2"
        and edge.producer_cell == "Engine!B1"
    )
    assert lag.distance == 1


def test_offset_helper_block_is_same_index_not_look_ahead(tmp_path: Path) -> None:
    catalog, _deps, graph = inverted_graph_parts(
        _offset_zipper_workbook(tmp_path), _offset_zipper_bindings()
    )
    assert schedule_coord("Engine!B2", catalog) == schedule_coord("Engine!E3", catalog)
    edges = collect_catalog_edges(catalog, graph).edges
    same_year = next(
        edge
        for edge in edges
        if edge.consumer_cell == "Engine!B2" and edge.producer_cell == "Engine!E3"
    )
    assert same_year.distance == 0


def test_cross_sheet_coords_are_not_column_subtraction(tmp_path: Path) -> None:
    catalog, _deps, graph = inverted_graph_parts(
        _cross_sheet_zipper_workbook(tmp_path), _cross_sheet_zipper_bindings()
    )
    assert schedule_coord("Engine!B2", catalog) == schedule_coord("Helper!C2", catalog) == 1
    edges = collect_catalog_edges(catalog, graph).edges
    same_year = next(
        edge
        for edge in edges
        if edge.consumer_cell == "Engine!B2" and edge.producer_cell == "Helper!C2"
    )
    assert same_year.distance == 0


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


def test_empty_key_rejects_ambiguous_public_coordinates(tmp_path: Path) -> None:
    """Public non-scalar series require unambiguous authored coordinates."""
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
    with pytest.raises(InvertedTreeExportError, match="authored semantic keys"):
        _ = catalog.get("values").tensor_domain


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


def test_schedule_index_is_eager_catalog_field(tmp_path: Path) -> None:
    catalog, _deps, _graph = inverted_graph_parts(_zipper_workbook(tmp_path), _zipper_bindings())
    assert isinstance(catalog.schedule, ScheduleIndex)
    assert catalog.schedule.coord_of["Engine!B2"] == 1
    with pytest.raises(InvertedTreeExportError, match="no schedule coordinate"):
        schedule_coord("Engine!Z99", catalog)
