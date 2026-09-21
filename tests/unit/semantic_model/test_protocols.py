"""Inverted-tree catalog types satisfy the shared semantic-model contract."""

from __future__ import annotations

from excel_grapher.core.address_keys import CanonicalAddress, as_canonical
from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    KeyPoint,
    ScheduleIndex,
    SeriesCatalog,
    SeriesHole,
    Statement,
)
from excel_grapher.exporter.inverted_tree.deps import CatalogEdges, DependenceEdge
from excel_grapher.exporter.semantic_catalog import (
    SemanticCatalogView,
    load_semantic_catalog,
)
from excel_grapher.grapher.series_graph import SeriesGraphOccupancy
from excel_grapher.semantic_model import (
    SCHEDULE_AXIS,
    AccessClass,
    CatalogOccupancy,
    Direction,
    HoleKind,
    Layout,
    SemanticCatalogLoader,
)
from excel_grapher.semantic_model import BoundSeries as BoundSeriesContract
from excel_grapher.semantic_model import CatalogEdges as CatalogEdgesContract
from excel_grapher.semantic_model import DependenceEdge as DependenceEdgeContract
from excel_grapher.semantic_model import KeyPoint as KeyPointContract
from excel_grapher.semantic_model import ScheduleIndex as ScheduleIndexContract
from excel_grapher.semantic_model import SemanticCatalogView as SemanticCatalogViewContract
from excel_grapher.semantic_model import SeriesCatalog as SeriesCatalogContract
from excel_grapher.semantic_model import SeriesHole as SeriesHoleContract
from excel_grapher.semantic_model import Statement as StatementContract
from excel_grapher.semantic_model.ownership import SHARED_CATALOG_TYPE_NAMES
from excel_grapher.semantic_model.types import (
    AccessClass as TypesAccessClass,
)
from excel_grapher.semantic_model.types import (
    Direction as TypesDirection,
)
from excel_grapher.semantic_model.types import (
    HoleKind as TypesHoleKind,
)
from excel_grapher.semantic_model.types import (
    Layout as TypesLayout,
)


def _protocol_members(protocol: type) -> set[str]:
    members = getattr(protocol, "__protocol_attrs__", None)
    if members:
        return set(members)
    names: set[str] = set(getattr(protocol, "__annotations__", {}))
    for name, value in vars(protocol).items():
        if name.startswith("_") and name not in {"__getitem__", "__call__"}:
            continue
        if callable(value) or isinstance(value, property):
            names.add(name)
    return names


def _assert_provides_protocol(obj: object, protocol: type) -> None:
    missing = sorted(name for name in _protocol_members(protocol) if not hasattr(obj, name))
    assert not missing, f"{type(obj).__name__} missing {protocol.__name__} members: {missing}"


def _minimal_catalog() -> tuple[BoundSeries, SeriesCatalog, Statement, SeriesHole, KeyPoint]:
    cell = as_canonical("Sheet1!A1")
    point = KeyPoint((("TIME_PERIOD", 2020),))
    hole = SeriesHole(index=0, address=cell, kind="blank")
    statement = Statement(
        statement_id="gdp",
        series_id="gdp",
        shape_key=None,
        start=0,
        stop=1,
        cells=(cell,),
        domain=(point,),
    )
    series = BoundSeries(
        series_id="gdp",
        layout="series",
        direction="output",
        cells=(cell,),
        key_fields=("TIME_PERIOD",),
        dtype="float",
        compute_name="compute_gdp",
        raw={"id": "gdp"},
        domain=(point,),
        statements=(statement,),
        holes=(hole,),
    )
    catalog = SeriesCatalog(
        series={"gdp": series},
        order=("gdp",),
        address_to_id={cell: "gdp"},
        schedule=ScheduleIndex(preferred={"gdp": ("TIME_PERIOD",)}, coord_of={}, index_by_coord={}),
    )
    return series, catalog, statement, hole, point


def test_shared_type_aliases_are_owned_by_semantic_model() -> None:
    from excel_grapher.exporter.inverted_tree import catalog as catalog_mod
    from excel_grapher.exporter.inverted_tree import deps as deps_mod

    assert catalog_mod.Direction is Direction is TypesDirection
    assert catalog_mod.Layout is Layout is TypesLayout
    assert catalog_mod.HoleKind is HoleKind is TypesHoleKind
    assert deps_mod.AccessClass is AccessClass is TypesAccessClass
    assert catalog_mod._SCHEDULE_AXIS == SCHEDULE_AXIS == "TIME_PERIOD"


def test_series_graph_occupancy_is_the_shared_occupancy_contract() -> None:
    assert SeriesGraphOccupancy is CatalogOccupancy


def test_inverted_tree_catalog_types_provide_contract_members() -> None:
    series, catalog, statement, hole, point = _minimal_catalog()
    edge = DependenceEdge(
        consumer_id="gdp",
        producer_id="gdp",
        consumer_cell=as_canonical("Sheet1!A1"),
        producer_cell=as_canonical("Sheet1!A1"),
        distance=0,
        access="identity",
    )
    edges = CatalogEdges(edges=(edge,), by_consumer={"gdp": (edge,)}, by_producer={"gdp": (edge,)})
    view = SemanticCatalogView(catalog=catalog, edges=edges, concepts={})

    _assert_provides_protocol(point, KeyPointContract)
    _assert_provides_protocol(statement, StatementContract)
    _assert_provides_protocol(hole, SeriesHoleContract)
    _assert_provides_protocol(series, BoundSeriesContract)
    _assert_provides_protocol(catalog, SeriesCatalogContract)
    _assert_provides_protocol(catalog.schedule, ScheduleIndexContract)
    _assert_provides_protocol(catalog, CatalogOccupancy)
    _assert_provides_protocol(edge, DependenceEdgeContract)
    _assert_provides_protocol(edges, CatalogEdgesContract)
    _assert_provides_protocol(view, SemanticCatalogViewContract)


def test_load_semantic_catalog_matches_loader_contract() -> None:
    assert callable(load_semantic_catalog)
    required = {"graph", "bindings", "workbook"}
    assert required <= set(load_semantic_catalog.__code__.co_varnames)
    _assert_provides_protocol(load_semantic_catalog, SemanticCatalogLoader)


def test_shared_catalog_type_names_match_issue_inventory() -> None:
    assert SHARED_CATALOG_TYPE_NAMES == (
        "KeyPoint",
        "BoundSeries",
        "Statement",
        "SeriesCatalog",
        "ScheduleIndex",
        "SeriesHole",
        "DependenceEdge",
        "CatalogEdges",
    )


def test_canonical_address_round_trip_on_key_point_protocol() -> None:
    _series, catalog, _statement, _hole, point = _minimal_catalog()
    address: CanonicalAddress = as_canonical("Sheet1!A1")
    found = catalog.key_point_for(address)
    assert found is not None
    assert found.as_mapping() == {"TIME_PERIOD": 2020}
    assert found["TIME_PERIOD"] == 2020
    assert catalog.series_id_for(address) == "gdp"
