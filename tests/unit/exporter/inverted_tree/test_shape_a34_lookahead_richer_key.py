"""T+1 look-ahead into a richer-keyed producer is a schedule peer (#745, #747).

A host keyed by `TIME_PERIOD` that reads `producer[holder, t+1]` is a lagged
/ cross-partition take of the holder nest, not a seed. After #745 export
succeeds; #747 classifies those edges as `shift` / `cross_partition` so the
last host cell is not an ambiguous-seed warning.
"""

from __future__ import annotations

import warnings
from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.access import overlapping_schedule_peer
from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, KeyPoint, Statement
from excel_grapher.exporter.inverted_tree.deps import (
    collect_all_dependence_edges,
    successor_address,
)
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    call_compute,
    generate_inverted,
    input_kwargs,
    inverted_graph_parts,
    load_package,
    make_catalog,
    write_workbook,
)

_TIME_DIM = {
    "id": "TIME_PERIOD",
    "concept": "TIME_PERIOD",
    "role": "key",
    "scope": "cell",
    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
}


def _measure() -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }


def _holder_dim(values: dict[str, int]) -> dict[str, Any]:
    return {
        "id": "HOLDER",
        "concept": "HOLDER",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "value_map", "values": values, "read": "string"},
    }


def _mcve_sheets() -> dict[str, dict[str, object]]:
    return {
        "Engine": {
            "B1": 2010,
            "C1": 2011,
            "D1": 2012,
            "A2": "residents",
            "B2": 10,
            "C2": 11,
            "D2": 12,
            "A3": "non-residents",
            "B3": 20,
            "C3": 21,
            "D3": 22,
            "A4": "non-residents",
            "B4": 30,
            "C4": 31,
            "D4": 32,
            "G1": 1,
            "B5": "=IF($G$1=1,C3+C4,C2)",
            "C5": "=IF($G$1=1,D3+D4,D2)",
        },
    }


def _mcve_bindings() -> dict[str, Any]:
    document = bindings_document(
        {
            "id": "fx_st",
            "sheet": "Engine",
            "data_range": "Engine!B2:D4",
            "layout": "series",
            "exclude_rows": ["3"],
            "input": {"setter": {"name": "set_fx_st"}},
            "structure": {
                "measure": _measure(),
                "dimensions": [
                    _holder_dim({"residents": 2, "non-residents": 4}),
                    _TIME_DIM,
                ],
            },
            "key": ["HOLDER", "TIME_PERIOD"],
        },
        {
            "id": "lc_st",
            "sheet": "Engine",
            "data_range": "Engine!B3:D3",
            "layout": "series",
            "input": {"setter": {"name": "set_lc_st"}},
            "structure": {
                "measure": _measure(),
                "dimensions": [_holder_dim({"non-residents": 3}), _TIME_DIM],
            },
            "key": ["HOLDER", "TIME_PERIOD"],
        },
        {
            "id": "flag",
            "sheet": "Engine",
            "data_range": "Engine!G1",
            "layout": "scalar",
            "input": {"setter": {"name": "set_flag"}},
            "structure": {"measure": _measure(), "dimensions": []},
            "key": [],
        },
        {
            "id": "stock",
            "sheet": "Engine",
            "data_range": "Engine!B5:C5",
            "layout": "series",
            "output": {"compute": {"name": "compute_stock"}},
            "structure": {"measure": _measure(), "dimensions": [_TIME_DIM]},
            "key": ["TIME_PERIOD"],
        },
        schema_version="1.14.0",
    )
    document["concept_scheme"]["concepts"].append({"id": "HOLDER", "dtype": "string"})
    return document


def _unique_lookahead_sheets() -> dict[str, dict[str, object]]:
    return {
        "Engine": {
            "B1": 2010,
            "C1": 2011,
            "D1": 2012,
            "A3": "non-residents",
            "B3": 20,
            "C3": 21,
            "D3": 22,
            "B5": "=C3",
            "C5": "=D3",
        },
    }


def _unique_lookahead_bindings() -> dict[str, Any]:
    document = bindings_document(
        {
            "id": "lc_st",
            "sheet": "Engine",
            "data_range": "Engine!B3:D3",
            "layout": "series",
            "input": {"setter": {"name": "set_lc_st"}},
            "structure": {
                "measure": _measure(),
                "dimensions": [_holder_dim({"non-residents": 3}), _TIME_DIM],
            },
            "key": ["HOLDER", "TIME_PERIOD"],
        },
        {
            "id": "stock",
            "sheet": "Engine",
            "data_range": "Engine!B5:C5",
            "layout": "series",
            "output": {"compute": {"name": "compute_stock"}},
            "structure": {"measure": _measure(), "dimensions": [_TIME_DIM]},
            "key": ["TIME_PERIOD"],
        },
        schema_version="1.14.0",
    )
    document["concept_scheme"]["concepts"].append({"id": "HOLDER", "dtype": "string"})
    return document


def _ambiguous_seed_warnings(caught: list[warnings.WarningMessage]) -> list[str]:
    return [
        str(item.message)
        for item in caught
        if issubclass(item.category, UserWarning) and "ambiguous seed" in str(item.message)
    ]


def _accesses(
    edges: tuple[Any, ...],
    consumer_id: str,
    producer_id: str,
) -> set[str]:
    return {
        edge.access
        for edge in edges
        if edge.consumer_id == consumer_id and edge.producer_id == producer_id
    }


def test_if_lookahead_into_richer_key_is_not_an_ambiguous_seed(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a34.xlsx", _mcve_sheets())
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        catalog, deps, graph = inverted_graph_parts(workbook, _mcve_bindings())
    assert _ambiguous_seed_warnings(caught) == []
    host = catalog.get("stock")
    stock = deps["stock"]
    assert stock.seed_id is None
    assert stock.is_scan is False
    assert successor_address(host, 0, catalog, graph) == host.cells[1]
    assert successor_address(host, 1, catalog, graph) is None
    edges = collect_all_dependence_edges(catalog, graph)
    for producer_id in ("fx_st", "lc_st"):
        classes = _accesses(edges, "stock", producer_id)
        assert classes
        assert classes <= {"shift", "cross_partition"}


def test_if_lookahead_into_richer_key_emits_and_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a34_eval.xlsx", _mcve_sheets())
    document = _mcve_bindings()
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        catalog, deps, graph = inverted_graph_parts(workbook, document)
        modules = generate_inverted(workbook, document)
    assert _ambiguous_seed_warnings(caught) == []
    assert deps["stock"].seed_id is None
    pkg = load_package(modules, tmp_path, name="a34_eval")
    cells = ["Engine!B5", "Engine!C5"]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = call_compute(pkg, "stock", input_kwargs(catalog, graph))
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx((52.0, 54.0))


def test_unique_lookahead_into_richer_key_is_not_a_seed(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a34_unique.xlsx", _unique_lookahead_sheets())
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        catalog, deps, graph = inverted_graph_parts(workbook, _unique_lookahead_bindings())
    assert _ambiguous_seed_warnings(caught) == []
    host = catalog.get("stock")
    stock = deps["stock"]
    assert stock.seed_id is None
    assert stock.is_scan is False
    assert successor_address(host, 1, catalog, graph) is None
    edges = collect_all_dependence_edges(catalog, graph)
    classes = _accesses(edges, "stock", "lc_st")
    assert classes
    assert classes <= {"shift", "cross_partition"}


def _series(
    series_id: str,
    cells: tuple[str, ...],
    key_fields: tuple[str, ...],
    domain: tuple[KeyPoint, ...],
    *,
    layout: str = "series",
) -> BoundSeries:
    n = len(cells)
    return BoundSeries(
        series_id=series_id,
        layout=layout,
        direction="internal",
        cells=cells,
        key_fields=key_fields,
        dtype="float",
        compute_name=None,
        raw={},
        domain=domain,
        statements=(Statement(series_id, series_id, None, 0, n, cells, domain),),
    )


def _peer_catalog(
    host: BoundSeries,
    producer: BoundSeries,
) -> tuple[Any, BoundSeries, BoundSeries]:
    catalog = make_catalog(
        series={host.series_id: host, producer.series_id: producer},
        order=(host.series_id, producer.series_id),
        address_to_id={
            **{cell: host.series_id for cell in host.cells},
            **{cell: producer.series_id for cell in producer.cells},
        },
    )
    return catalog, host, producer


def test_overlapping_schedule_peer_accepts_richer_keyed_producer() -> None:
    host = _series(
        "stock",
        ("Engine!B5", "Engine!C5"),
        ("TIME_PERIOD",),
        (KeyPoint((("TIME_PERIOD", 2010),)), KeyPoint((("TIME_PERIOD", 2011),))),
    )
    producer = _series(
        "fx_st",
        ("Engine!B2", "Engine!C2", "Engine!D2"),
        ("HOLDER", "TIME_PERIOD"),
        (
            KeyPoint((("HOLDER", "residents"), ("TIME_PERIOD", 2010))),
            KeyPoint((("HOLDER", "residents"), ("TIME_PERIOD", 2011))),
            KeyPoint((("HOLDER", "residents"), ("TIME_PERIOD", 2012))),
        ),
    )
    catalog, host, producer = _peer_catalog(host, producer)
    assert overlapping_schedule_peer(host, producer, catalog)
    assert overlapping_schedule_peer(host, host, catalog)


def test_overlapping_schedule_peer_rejects_unrelated_or_unkeyed_producers() -> None:
    host = _series(
        "stock",
        ("Engine!B5", "Engine!C5"),
        ("TIME_PERIOD",),
        (KeyPoint((("TIME_PERIOD", 2010),)), KeyPoint((("TIME_PERIOD", 2011),))),
    )
    unkeyed = _series(
        "flag",
        ("Engine!G1",),
        (),
        (KeyPoint(()),),
        layout="scalar",
    )
    no_time = _series(
        "label",
        ("Engine!A2", "Engine!A3"),
        ("HOLDER",),
        (KeyPoint((("HOLDER", "residents"),)), KeyPoint((("HOLDER", "non-residents"),))),
    )
    other_outer = _series(
        "gdp",
        ("Engine!B8", "Engine!C8"),
        ("COUNTRY", "TIME_PERIOD"),
        (
            KeyPoint((("COUNTRY", "US"), ("TIME_PERIOD", 2010))),
            KeyPoint((("COUNTRY", "US"), ("TIME_PERIOD", 2011))),
        ),
    )
    disjoint = _series(
        "fx_st",
        ("Engine!B12", "Engine!C12"),
        ("HOLDER", "TIME_PERIOD"),
        (
            KeyPoint((("HOLDER", "residents"), ("TIME_PERIOD", 2008))),
            KeyPoint((("HOLDER", "residents"), ("TIME_PERIOD", 2009))),
        ),
    )
    richer_host = _series(
        "fx_st_host",
        ("Engine!B22", "Engine!C22"),
        ("HOLDER", "TIME_PERIOD"),
        (
            KeyPoint((("HOLDER", "residents"), ("TIME_PERIOD", 2010))),
            KeyPoint((("HOLDER", "residents"), ("TIME_PERIOD", 2011))),
        ),
    )
    catalog = make_catalog(
        series={
            "stock": host,
            "flag": unkeyed,
            "label": no_time,
            "gdp": other_outer,
            "fx_st": disjoint,
            "fx_st_host": richer_host,
        },
        order=("stock", "flag", "label", "gdp", "fx_st", "fx_st_host"),
        address_to_id={
            **{cell: "stock" for cell in host.cells},
            **{cell: "flag" for cell in unkeyed.cells},
            **{cell: "label" for cell in no_time.cells},
            **{cell: "gdp" for cell in other_outer.cells},
            **{cell: "fx_st" for cell in disjoint.cells},
            **{cell: "fx_st_host" for cell in richer_host.cells},
        },
    )
    assert not overlapping_schedule_peer(host, unkeyed, catalog)
    assert not overlapping_schedule_peer(host, no_time, catalog)
    assert not overlapping_schedule_peer(richer_host, host, catalog)
    assert not overlapping_schedule_peer(other_outer, disjoint, catalog)
    assert not overlapping_schedule_peer(host, disjoint, catalog)
