"""Grouped bindings still export via `generate_modules()`; packages omit `list_groups`."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

import pytest

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import validate_bindings_document
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.integration.user_flows.utils import (
    add_formula_mirror_row,
    load_generated_package,
    write_series_bindings_workbook,
)


def _row_series(
    series_id: str,
    row: int,
    groups: list[dict[str, Any]] | None,
    *,
    direction: str = "input",
) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "id": series_id,
        "sheet": "Sheet1",
        "data_range": f"Sheet1!F{row}:J{row}",
        "layout": "series",
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                }
            ],
        },
        "key": ["TIME_PERIOD"],
    }
    if direction == "input":
        entry["input"] = {"setter": {"name": f"set_{series_id}"}}
    else:
        entry["output"] = {"compute": {"name": f"compute_{series_id}"}}
    if groups is not None:
        entry["groups"] = groups
    return entry


GROUPED_DOCUMENT: dict[str, Any] = {
    "schema_version": "1.5.0",
    "workbook": "series_bindings.xlsx",
    "series": [
        _row_series("primary_balance", 5, [{"path": ["Fiscal"], "order": 1}]),
        _row_series("gdp_growth", 3, [{"path": ["Macro", "Growth"]}]),
        _row_series("interest_rate", 4, None),
        _row_series("primary_balance_out", 6, None, direction="output"),
    ],
}


@pytest.fixture
def workbook(tmp_path: Path) -> Path:
    path = tmp_path / "series_bindings.xlsx"
    write_series_bindings_workbook(path)
    add_formula_mirror_row(path, source_row=5, dest_row=6)
    return path


def _export(workbook: Path, tmp_path: Path, document: dict[str, Any], name: str):
    bindings = validate_bindings_document(deepcopy(document))
    targets = all_series_targets(bindings, workbook=workbook)
    graph = create_dependency_graph(workbook, targets, load_values=True)
    return load_generated_package(graph, bindings, workbook, tmp_path, name=name)


def test_grouped_export_omits_list_groups(workbook: Path, tmp_path: Path) -> None:
    pkg, modules = _export(workbook, tmp_path, GROUPED_DOCUMENT, "grouped_export")
    joined = "\n".join(modules.values())
    assert "def list_groups(" not in joined
    assert "def set_primary_balance(" not in joined
    result = pkg.compute_primary_balance_out(primary_balance=(-1.0, -0.5, 0.0, 7.5, 1.0))
    assert result == pytest.approx((-1.0, -0.5, 0.0, 7.5, 1.0))


def test_ungrouped_bindings_export_omits_list_groups(workbook: Path, tmp_path: Path) -> None:
    document = deepcopy(GROUPED_DOCUMENT)
    for series in document["series"]:
        series.pop("groups", None)
    _pkg, modules = _export(workbook, tmp_path, document, "ungrouped_export")
    assert "def list_groups(" not in "\n".join(modules.values())
