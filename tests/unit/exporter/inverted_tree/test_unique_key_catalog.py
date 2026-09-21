"""Keyed catalog series cannot opt out of unique tensor coordinates (#947)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.exporter.inverted_tree.catalog import build_catalog
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    validate_bindings_document,
    validate_series_bindings,
)
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    series_entry,
    write_workbook,
)

_INCOMPATIBLE = "require_unique_key: false is incompatible with a keyed catalog series"


def _collapsed_year_series(*, unique_key: bool | None) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "id": "engine_collapsed_years",
        "sheet": "Engine",
        "data_range": "Engine!B2:C2",
        "layout": "series",
        "constant": {},
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
    if unique_key is not None:
        entry["validation"] = {"require_unique_key": unique_key}
    return entry


def _collapsed_year_bindings(*, unique_key: bool | None) -> dict[str, Any]:
    return validate_bindings_document(
        {
            "schema_version": "1.19.0",
            "workbook": "mcve.xlsx",
            "series": [_collapsed_year_series(unique_key=unique_key)],
        }
    )


def _collapsed_year_workbook(tmp_path: Path) -> Path:
    return write_workbook(tmp_path / "mcve.xlsx", {"Engine": {"B2": 1, "C2": 2}})


def test_catalog_refuses_require_unique_key_false_on_keyed_series(tmp_path: Path) -> None:
    path = _collapsed_year_workbook(tmp_path)
    with pytest.raises(InvertedTreeExportError, match=_INCOMPATIBLE):
        build_catalog(_collapsed_year_bindings(unique_key=False), workbook=path)


def test_catalog_refuses_require_unique_key_false_when_keys_are_unique(tmp_path: Path) -> None:
    path = write_workbook(tmp_path / "unique.xlsx", {"Engine": {"B2": 1}})
    bindings = validate_bindings_document(
        {
            "schema_version": "1.19.0",
            "series": [
                {
                    "id": "engine_year",
                    "sheet": "Engine",
                    "data_range": "Engine!B2",
                    "layout": "scalar",
                    "constant": {},
                    "validation": {"require_unique_key": False},
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
                                    "values": {2020: "B"},
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
    with pytest.raises(InvertedTreeExportError, match=_INCOMPATIBLE):
        build_catalog(bindings, workbook=path)


def test_tensor_domain_stays_fail_closed_on_duplicate_coordinates(tmp_path: Path) -> None:
    path = _collapsed_year_workbook(tmp_path)
    catalog = build_catalog(_collapsed_year_bindings(unique_key=True), workbook=path)
    with pytest.raises(
        InvertedTreeExportError,
        match="domain contains duplicate coordinates; correct the authored keys",
    ):
        _ = catalog.get("engine_collapsed_years").tensor_domain


def test_keyless_series_may_opt_out_of_unique_key(tmp_path: Path) -> None:
    path = write_workbook(tmp_path / "scalar.xlsx", {"Engine": {"B2": 1}})
    bindings = validate_bindings_document(
        {
            "schema_version": "1.19.0",
            "series": [
                {
                    "id": "engine_pin",
                    "sheet": "Engine",
                    "data_range": "Engine!B2",
                    "layout": "scalar",
                    "constant": {},
                    "validation": {"require_unique_key": False},
                    "structure": {
                        "measure": {
                            "concept": "OBS_VALUE",
                            "dtype": "float",
                            "bind": {"kind": "data_cell", "read": "float"},
                        },
                        "dimensions": [],
                    },
                    "key": [],
                }
            ],
        }
    )
    series = build_catalog(bindings, workbook=path).get("engine_pin")
    assert tuple(series.tensor_domain) == ((),)


def test_validate_reports_require_unique_key_incompatible(tmp_path: Path) -> None:
    path = _collapsed_year_workbook(tmp_path)
    graph = create_dependency_graph(path, ["Engine!B2", "Engine!C2"], load_values=True)
    report = validate_series_bindings(
        graph, _collapsed_year_bindings(unique_key=False), workbook=path
    )
    assert report["ok"] is False
    assert any(
        issue["code"] == "require_unique_key_incompatible"
        and issue["series_id"] == "engine_collapsed_years"
        and _INCOMPATIBLE in issue["message"]
        for issue in report["issues"]
    )


def test_emit_refuses_require_unique_key_false_on_keyed_series(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "emit.xlsx",
        {
            "Engine": {"B1": 2020, "C1": 2021, "B2": 1, "C2": 2},
            "Outputs": {"A1": "=Engine!B2"},
        },
    )
    engine = series_entry(
        "engine_years",
        "Engine!B2:C2",
        layout="series",
        direction="constant",
        header_row=1,
    )
    engine["validation"] = {"require_unique_key": False}
    document = bindings_document(
        engine,
        series_entry("out", "Outputs!A1", layout="scalar", direction="output"),
    )
    with pytest.raises(InvertedTreeExportError, match=_INCOMPATIBLE):
        generate_inverted(workbook, document)
