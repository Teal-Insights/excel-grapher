"""Integration: series I/O exports through `generate_modules()`, not `generate()`."""

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

_STRUCTURE: dict[str, Any] = {
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
}

BINDINGS_DOCUMENT: dict[str, Any] = {
    "schema_version": "1.3.0",
    "workbook": "series_bindings.xlsx",
    "series": [
        {
            "id": "borvelia_primary_balance",
            "sheet": "Sheet1",
            "data_range": "Sheet1!F5:J5",
            "layout": "series",
            "input": {"setter": {"name": "set_borvelia_primary_balance"}},
            "structure": deepcopy(_STRUCTURE),
            "key": ["TIME_PERIOD"],
            "series_context": {
                "REF_AREA": "Borvelia",
                "INDICATOR": "Primary balance (% of GDP)",
            },
        },
        {
            "id": "borvelia_primary_balance_out",
            "sheet": "Sheet1",
            "data_range": "Sheet1!F6:J6",
            "layout": "series",
            "output": {"compute": {"name": "compute_borvelia_primary_balance_out"}},
            "structure": deepcopy(_STRUCTURE),
            "key": ["TIME_PERIOD"],
            "series_context": {
                "REF_AREA": "Borvelia",
                "INDICATOR": "Primary balance (% of GDP)",
            },
        },
    ],
}


@pytest.fixture
def workbook(tmp_path: Path) -> Path:
    path = tmp_path / "series_bindings.xlsx"
    write_series_bindings_workbook(path)
    add_formula_mirror_row(path, source_row=5, dest_row=6)
    return path


def test_generate_modules_computes_input_series(workbook: Path, tmp_path: Path) -> None:
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    targets = all_series_targets(bindings, workbook=workbook)
    graph = create_dependency_graph(workbook, targets, load_values=True)
    pkg, modules = load_generated_package(
        graph, bindings, workbook, tmp_path, name="series_input_export"
    )

    joined = "\n".join(modules.values())
    assert "def set_borvelia_primary_balance(" not in joined
    assert "def list_setters(" not in joined
    assert "def list_readers(" not in joined
    assert not hasattr(pkg, "set_borvelia_primary_balance")
    assert not hasattr(pkg, "make_context")

    result = pkg.compute_borvelia_primary_balance_out(
        borvelia_primary_balance=(-1.0, -0.5, 0.0, 7.5, 1.0)
    )
    assert result == pytest.approx((-1.0, -0.5, 0.0, 7.5, 1.0))
    records = pkg.as_records(pkg.compute_borvelia_primary_balance_out, result)
    by_period = {record["TIME_PERIOD"]: record["OBS_VALUE"] for record in records}
    assert by_period[4] == pytest.approx(7.5)
