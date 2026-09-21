"""Extract-only series compile CellTypeEnv without inverted-tree tensors (#951)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.exporter.inverted_tree.catalog import build_catalog
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from excel_grapher.series_bindings import (
    validate_bindings_document,
    validate_series_bindings,
)
from excel_grapher.series_bindings.domains import cell_type_env_from_bindings
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    series_entry,
    write_workbook,
)

_SCHEMA = "1.21.0"
_SKIPPED_OWNER = "validation.catalog: false uniquely owns on-graph formula cell"


def _engine_pin_series(**overrides: Any) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "id": "engine_stock_pin",
        "sheet": "Engine",
        "data_range": "Engine!B2:C2",
        "layout": "series",
        "constant": {},
        "validation": {"intersect_graph_leaves": False, "catalog": False},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "id": "ROW",
                    "concept": "ROW",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "value_map", "values": {2: 2}, "read": "int"},
                },
                {
                    "id": "COL",
                    "concept": "COL",
                    "role": "key",
                    "scope": "cell",
                    "bind": {
                        "kind": "value_map",
                        "values": {"B": "B", "C": "C"},
                        "read": "string",
                    },
                },
            ],
        },
        "key": ["ROW", "COL"],
        "notes": "Extract-time domain pin for OFFSET/INDEX; not a public series.",
    }
    entry.update(overrides)
    return entry


def _engine_pin_bindings(**overrides: Any) -> dict[str, Any]:
    return validate_bindings_document(
        {
            "schema_version": _SCHEMA,
            "workbook": "mcve.xlsx",
            "series": [_engine_pin_series(**overrides)],
        }
    )


def _engine_pin_workbook(tmp_path: Path) -> Path:
    return write_workbook(tmp_path / "mcve.xlsx", {"Engine": {"B2": 1.5, "C2": 2.5}})


def test_schema_accepts_validation_catalog_false() -> None:
    bindings = _engine_pin_bindings()
    assert bindings["series"][0]["validation"]["catalog"] is False


def test_extract_only_pin_compiles_cell_type_env_and_is_omitted_from_catalog(
    tmp_path: Path,
) -> None:
    path = _engine_pin_workbook(tmp_path)
    bindings = _engine_pin_bindings()

    env = cell_type_env_from_bindings(bindings, workbook=path)
    assert "Engine!B2" in env
    assert "Engine!C2" in env
    config = DynamicRefConfig.from_bindings(bindings, path)
    assert "Engine!B2" in config.cell_type_env

    catalog = build_catalog(bindings, workbook=path)
    assert "engine_stock_pin" not in catalog.series


def test_extract_only_pin_may_opt_out_of_unique_keys(tmp_path: Path) -> None:
    path = write_workbook(tmp_path / "collapsed.xlsx", {"Engine": {"B2": 1, "C2": 2}})
    bindings = validate_bindings_document(
        {
            "schema_version": _SCHEMA,
            "series": [
                {
                    "id": "engine_collapsed_years",
                    "sheet": "Engine",
                    "data_range": "Engine!B2:C2",
                    "layout": "series",
                    "constant": {},
                    "validation": {"catalog": False, "require_unique_key": False},
                    "structure": {
                        "measure": {
                            "concept": "OBS_VALUE",
                            "dtype": "float",
                            "bind": {"kind": "data_cell", "read": "float"},
                        },
                        "dimensions": [
                            {
                                "id": "TIME_PERIOD",
                                "concept": "TIME_PERIOD",
                                "role": "key",
                                "scope": "cell",
                                "bind": {
                                    "kind": "value_map",
                                    "values": {2020: "B:C"},
                                    "read": "int",
                                },
                            }
                        ],
                    },
                    "key": ["TIME_PERIOD"],
                }
            ],
        }
    )
    graph = create_dependency_graph(path, ["Engine!B2", "Engine!C2"], load_values=True)
    report = validate_series_bindings(graph, bindings, workbook=path)
    assert report["ok"] is True
    catalog = build_catalog(bindings, workbook=path)
    assert "engine_collapsed_years" not in catalog.series
    env = cell_type_env_from_bindings(bindings, workbook=path)
    assert "Engine!B2" in env
    assert "Engine!C2" in env


def test_cataloged_series_remain_in_catalog_by_default(tmp_path: Path) -> None:
    path = _engine_pin_workbook(tmp_path)
    bindings = _engine_pin_bindings(validation={"intersect_graph_leaves": False})
    catalog = build_catalog(bindings, workbook=path)
    assert "engine_stock_pin" in catalog.series
    _ = catalog.get("engine_stock_pin").tensor_domain


def test_validate_reports_catalog_skipped_formula_owner(tmp_path: Path) -> None:
    path = write_workbook(
        tmp_path / "formula_owner.xlsx",
        {
            "Inputs": {"A1": 1},
            "Engine": {"B2": "=Inputs!A1+1"},
            "Outputs": {"A1": "=Engine!B2"},
        },
    )
    document = bindings_document(
        series_entry("value", "Inputs!A1", layout="scalar", direction="input", dtype="int"),
        series_entry("path", "Engine!B2", layout="scalar", direction="internal"),
        series_entry("out", "Outputs!A1", layout="scalar", direction="output"),
        schema_version=_SCHEMA,
    )
    document["series"][1]["validation"] = {"catalog": False}
    bindings = validate_bindings_document(document)
    graph = create_dependency_graph(
        path, all_series_targets(bindings, workbook=path), load_values=True
    )
    report = validate_series_bindings(graph, bindings, workbook=path)
    assert report["ok"] is False
    assert any(
        issue["code"] == "catalog_skipped_formula_owner"
        and issue["series_id"] == "path"
        and issue["address"] == "Engine!B2"
        and _SKIPPED_OWNER in issue["message"]
        for issue in report["issues"]
    )


def test_build_catalog_fails_closed_when_skipped_series_owns_formula(tmp_path: Path) -> None:
    path = write_workbook(
        tmp_path / "formula_catalog.xlsx",
        {
            "Inputs": {"A1": 1},
            "Engine": {"B2": "=Inputs!A1+1"},
            "Outputs": {"A1": "=Engine!B2"},
        },
    )
    document = bindings_document(
        series_entry("value", "Inputs!A1", layout="scalar", direction="input", dtype="int"),
        series_entry("path", "Engine!B2", layout="scalar", direction="internal"),
        series_entry("out", "Outputs!A1", layout="scalar", direction="output"),
        schema_version=_SCHEMA,
    )
    document["series"][1]["validation"] = {"catalog": False}
    bindings = validate_bindings_document(document)
    graph = create_dependency_graph(
        path, all_series_targets(bindings, workbook=path), load_values=True
    )
    with pytest.raises(InvertedTreeExportError, match=_SKIPPED_OWNER):
        build_catalog(bindings, workbook=path, graph=graph)


def test_extract_only_pin_is_omitted_from_exported_data_module(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "emit.xlsx",
        {
            "Inputs": {"A1": 1},
            "Engine": {"B2": 1.5, "C2": 2.5},
            "Outputs": {"A1": "=Inputs!A1"},
        },
    )
    pin = _engine_pin_series()
    document = bindings_document(
        pin,
        series_entry("value", "Inputs!A1", layout="scalar", direction="input", dtype="int"),
        series_entry("out", "Outputs!A1", layout="scalar", direction="output"),
        schema_version=_SCHEMA,
    )
    modules = generate_inverted(workbook, document)
    assert "engine_stock_pin" not in modules["data.py"]
    assert "ENGINE_STOCK_PIN" not in modules["data.py"]
    assert "VALUE_DEFAULT" in modules["data.py"]
