"""Public catalog lookup from a graph node to key points and dimension binds."""

from __future__ import annotations

import ast
from pathlib import Path
from types import MappingProxyType

import pytest

from excel_grapher.core.address_keys import as_canonical
from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, KeyPoint, Statement
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher.node import EMPTY_METADATA, make_cell_node, node_to_view
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    make_catalog,
    series_entry,
    write_workbook,
)

_A2 = as_canonical("Engine!A2")
_B2 = as_canonical("Engine!B2")
_UNBOUND = as_canonical("Engine!Z99")


def _series(
    *,
    series_id: str = "path",
    cells: tuple[str, ...] = (_A2, _B2),
    key_fields: tuple[str, ...] = ("TIME_PERIOD",),
    domain: tuple[KeyPoint, ...] | None = None,
    raw: dict | None = None,
) -> BoundSeries:
    cell_tuple = tuple(as_canonical(cell) for cell in cells)
    if domain is None:
        domain = tuple(
            KeyPoint((("TIME_PERIOD", year),)) for year in range(2009, 2009 + len(cell_tuple))
        )
    return BoundSeries(
        series_id=series_id,
        layout="series",
        direction="input",
        cells=cell_tuple,
        key_fields=key_fields,
        dtype="float",
        compute_name=None,
        raw=raw or {},
        domain=domain,
        statements=(Statement(series_id, series_id, None, 0, len(cell_tuple), cell_tuple, domain),),
    )


def _catalog(series: BoundSeries):
    return make_catalog(
        series={series.series_id: series},
        order=(series.series_id,),
        address_to_id={cell: series.series_id for cell in series.cells},
    )


def test_key_point_for_returns_cell_coordinates() -> None:
    catalog = _catalog(_series())
    point = catalog.key_point_for(_A2)
    assert point is not None
    assert point.as_mapping() == {"TIME_PERIOD": 2009}
    assert catalog.key_point_for(_B2) == KeyPoint((("TIME_PERIOD", 2010),))


def test_key_point_for_unbound_cell_is_none() -> None:
    catalog = _catalog(_series())
    assert catalog.key_point_for(_UNBOUND) is None


def test_require_key_point_for_raises_when_unbound() -> None:
    catalog = _catalog(_series())
    with pytest.raises(InvertedTreeExportError, match="Engine!Z99"):
        catalog.require_key_point_for(_UNBOUND)


def test_key_point_for_empty_key_is_empty_point_not_none() -> None:
    series = _series(key_fields=(), domain=(KeyPoint(()), KeyPoint(())))
    catalog = _catalog(series)
    found = catalog.key_point_for(_A2)
    assert found is not None
    assert found.as_mapping() == {}


def test_binds_for_returns_series_level_dimension_binds() -> None:
    time_bind = {"kind": "column_header", "header_row": 1, "read": "int"}
    scenario_bind = {"kind": "constant", "value": "baseline"}
    series = _series(
        raw={
            "structure": {
                "measure": {"bind": {"kind": "data_cell", "read": "float"}},
                "dimensions": [
                    {"id": "TIME_PERIOD", "bind": time_bind},
                    {"id": "SCENARIO", "bind": scenario_bind},
                ],
                "attributes": [
                    {"concept": "UNIT_MEASURE", "value": "PC_GDP"},
                ],
            }
        }
    )
    catalog = _catalog(series)
    binds = catalog.binds_for(_A2)
    assert binds == catalog.binds_for(_B2)
    assert binds == {
        "TIME_PERIOD": time_bind,
        "SCENARIO": scenario_bind,
    }
    assert "OBS_VALUE" not in binds
    assert "UNIT_MEASURE" not in binds


def test_binds_for_unbound_cell_is_none() -> None:
    catalog = _catalog(_series())
    assert catalog.binds_for(_UNBOUND) is None


def test_binds_for_bound_series_without_dimensions_is_empty_mapping() -> None:
    catalog = _catalog(_series(raw={"structure": {"dimensions": []}}))
    binds = catalog.binds_for(_A2)
    assert binds is not None
    assert dict(binds) == {}
    assert isinstance(binds, MappingProxyType)


def test_require_binds_for_raises_when_unbound() -> None:
    catalog = _catalog(_series())
    with pytest.raises(InvertedTreeExportError, match="Engine!Z99"):
        catalog.require_binds_for(_UNBOUND)


def test_bound_series_dimension_bind_matches_field() -> None:
    time_bind = {"kind": "row_label", "label_column": "A", "read": "string"}
    series = _series(
        raw={"structure": {"dimensions": [{"concept": "TIME_PERIOD", "bind": time_bind}]}}
    )
    assert series.dimension_bind("TIME_PERIOD") == time_bind
    assert series.dimension_bind("COUNTRY") is None


def test_node_adapters_delegate_to_catalog() -> None:
    time_bind = {"kind": "column_header", "header_row": 1, "read": "int"}
    series = _series(raw={"structure": {"dimensions": [{"id": "TIME_PERIOD", "bind": time_bind}]}})
    catalog = _catalog(series)
    node = make_cell_node("Engine", "A", 2, value=1.0)
    view = node_to_view(node)

    assert node.key_point(catalog) == KeyPoint((("TIME_PERIOD", 2009),))
    assert view.key_point(catalog) == node.key_point(catalog)
    assert node.dimension_binds(catalog) == {"TIME_PERIOD": time_bind}
    assert view.dimension_binds(catalog) is node.dimension_binds(catalog)


def test_node_adapters_return_none_when_unbound() -> None:
    catalog = _catalog(_series())
    node = make_cell_node("Engine", "Z", 99, value=0)
    view = node_to_view(node)
    assert node.key_point(catalog) is None
    assert node.dimension_binds(catalog) is None
    assert view.key_point(catalog) is None
    assert view.dimension_binds(catalog) is None


def test_lookup_does_not_copy_binds_onto_node_metadata() -> None:
    series = _series(
        raw={
            "structure": {"dimensions": [{"id": "TIME_PERIOD", "bind": {"kind": "column_header"}}]}
        }
    )
    catalog = _catalog(series)
    node = make_cell_node("Engine", "A", 2)
    assert node.metadata is EMPTY_METADATA
    node.key_point(catalog)
    node.dimension_binds(catalog)
    assert node.metadata is EMPTY_METADATA
    assert list(node.metadata) == []


def test_node_module_does_not_import_inverted_tree() -> None:
    tree = ast.parse(Path("excel_grapher/grapher/node.py").read_text(encoding="utf-8"))
    found: list[str] = []
    for child in ast.walk(tree):
        if isinstance(child, ast.ImportFrom) and child.module and "inverted_tree" in child.module:
            found.append(child.module)
        if isinstance(child, ast.Import):
            found.extend(alias.name for alias in child.names if "inverted_tree" in alias.name)
    assert found == []


def test_build_catalog_lookup_resolves_header_binds(tmp_path: Path) -> None:
    from excel_grapher.exporter.inverted_tree.catalog import build_catalog
    from excel_grapher.grapher import create_dependency_graph
    from excel_grapher.series_bindings import validate_bindings_document
    from excel_grapher.series_bindings.workflow import all_series_targets

    workbook = write_workbook(
        tmp_path / "lookup.xlsx",
        {"Engine": {"A1": 2009, "B1": 2010, "A2": 1.0, "B2": 2.0}},
    )
    bindings = validate_bindings_document(
        bindings_document(
            series_entry(
                "path",
                "Engine!A2:B2",
                layout="series",
                header_row=1,
            )
        )
    )
    graph = create_dependency_graph(
        workbook,
        all_series_targets(bindings, workbook=workbook),
        load_values=True,
    )
    catalog = build_catalog(bindings, workbook=workbook, graph=graph)
    node = graph.get_node("Engine!A2")
    assert node is not None
    assert catalog.key_point_for(as_canonical(node.key)) == KeyPoint((("TIME_PERIOD", 2009),))
    binds = catalog.binds_for(as_canonical(node.key))
    assert binds is not None
    assert binds["TIME_PERIOD"]["kind"] == "column_header"
    assert node.key_point(catalog) == catalog.key_point_for(as_canonical(node.key))
    assert node.dimension_binds(catalog) == binds
