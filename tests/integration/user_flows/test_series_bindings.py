"""Integration tests for declarative series bindings against micro-workbook graphs."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

import pytest

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    bindings_canonical_sha256,
    expand_data_range,
    resolve_series_bindings,
    validate_bindings_document,
    validate_series_bindings,
)
from excel_grapher.series_bindings.types import WorkbookSeriesBindings

MICRO_WORKBOOKS = Path(__file__).resolve().parents[3] / "examples" / "micro_workbooks"
WORKBOOK = MICRO_WORKBOOKS / "series_bindings.xlsx"

BINDINGS_DOCUMENT: dict[str, Any] = {
    "schema_version": "1.0.0",
    "workbook": "series_bindings.xlsx",
    "series": [
        {
            "id": "borvelia_primary_balance",
            "sheet": "Sheet1",
            "data_range": "Sheet1!F5:J5",
            "layout": "row_series",
            "editable": True,
            "setter": {"name": "set_borvelia_primary_balance"},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [
                    {
                        "concept": "REF_AREA",
                        "role": "key",
                        "scope": "series",
                        "bind": {"kind": "cell", "address": "Sheet1!A2", "read": "string"},
                        "include_in_record": False,
                    },
                    {
                        "concept": "INDICATOR",
                        "role": "key",
                        "scope": "series",
                        "bind": {
                            "kind": "row_label",
                            "label_column": "A",
                            "read": "string",
                            "normalize": "strip_trailing_unit",
                        },
                        "include_in_record": False,
                    },
                    {
                        "concept": "TIME_PERIOD",
                        "role": "key",
                        "scope": "cell",
                        "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                    },
                ],
                "attributes": [
                    {
                        "concept": "UNIT_MEASURE",
                        "role": "attribute",
                        "value": "PC_GDP",
                        "include_in_record": True,
                    }
                ],
            },
            "key": ["TIME_PERIOD"],
            "series_context": {
                "REF_AREA": "Borvelia",
                "INDICATOR": "Primary balance (% of GDP)",
            },
        }
    ],
}


@pytest.fixture
def bindings() -> WorkbookSeriesBindings:
    return validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))


@pytest.fixture
def graph(bindings: WorkbookSeriesBindings):
    targets: list[str] = []
    for series in bindings["series"]:
        targets.extend(expand_data_range(series["data_range"], workbook=WORKBOOK))
    return create_dependency_graph(WORKBOOK, targets, load_values=True)


def test_micro_workbook_bindings_validate_against_graph(
    bindings: WorkbookSeriesBindings,
    graph,
) -> None:
    report = validate_series_bindings(graph, bindings, workbook=WORKBOOK)
    assert report["ok"] is True
    assert not any(issue["level"] == "error" for issue in report["issues"])
    assert not any(issue["code"] == "unique_key_deferred" for issue in report["issues"])


def test_micro_workbook_resolves_unique_keys(bindings: WorkbookSeriesBindings, graph) -> None:
    report = resolve_series_bindings(graph, bindings, workbook=WORKBOOK)
    assert report["ok"] is True
    borvelia = next(s for s in report["series"] if s["series_id"] == "borvelia_primary_balance")
    assert borvelia["requires_address"] is False
    assert len(borvelia["leaves"]) == 5
    periods = {leaf["key"]["TIME_PERIOD"] for leaf in borvelia["leaves"]}
    assert periods == {1, 2, 3, 4, 5}


def test_micro_workbook_covers_mvp_series_layouts(bindings: WorkbookSeriesBindings) -> None:
    by_id = {series["id"]: series for series in bindings["series"]}
    assert set(by_id) == {"borvelia_primary_balance"}
    assert by_id["borvelia_primary_balance"]["layout"] == "row_series"


def test_bindings_canonical_hash_is_stable(bindings: WorkbookSeriesBindings) -> None:
    first = bindings_canonical_sha256(bindings)
    second = bindings_canonical_sha256(validate_bindings_document(deepcopy(BINDINGS_DOCUMENT)))
    assert first == second
    assert len(first) == 64
