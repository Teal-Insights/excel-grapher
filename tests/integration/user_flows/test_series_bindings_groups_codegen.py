"""Grouped bindings still export via `generate_modules()`."""

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
        entry["input"] = {}
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


def test_grouped_bindings_export_computes(workbook: Path, tmp_path: Path) -> None:
    bindings = validate_bindings_document(deepcopy(GROUPED_DOCUMENT))
    targets = all_series_targets(bindings, workbook=workbook)
    graph = create_dependency_graph(workbook, targets, load_values=True)
    pkg, modules = load_generated_package(
        graph, bindings, workbook, tmp_path, name="grouped_export"
    )

    assert "def set_primary_balance(" not in "\n".join(modules.values())
    result = pkg.compute_primary_balance_out(
        primary_balance=pkg.data.PRIMARY_BALANCE.with_records(
            zip(((1,), (2,), (3,), (4,), (5,)), (-1.0, -0.5, 0.0, 7.5, 1.0), strict=True),
        )
    )
    assert tuple(result[period] for period in (1, 2, 3, 4, 5)) == pytest.approx(
        (-1.0, -0.5, 0.0, 7.5, 1.0)
    )
