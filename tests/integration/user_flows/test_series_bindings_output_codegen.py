"""Integration: output series export through `generate_modules()`.

`BINDINGS_DOCUMENT` and `_write_output_workbook` remain the mixed G5 fixture used by
`emit_computes_block` unit tests. Package export uses a split input/output workbook.
"""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

import pytest
import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import validate_bindings_document
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.integration.user_flows.utils import load_generated_package


def _write_output_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write("A2", "Borvelia")
    ws.write("A5", "Primary balance (% of GDP)")
    for col, year in enumerate([1, 2, 3, 4, 5], start=5):
        ws.write(0, col, year)
        ws.write_number(4, col, float(year))
    ws.write_formula("G5", "=F5+1")
    wb.close()


def _write_inverted_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write("A2", "Borvelia")
    ws.write("A5", "Primary balance (% of GDP)")
    for col, year in enumerate([1, 2, 3, 4, 5], start=5):
        ws.write(0, col, year)
        ws.write_number(4, col, float(year))
        letter = "FGHIJ"[col - 5]
        ws.write_formula(5, col, f"={letter}5")
    wb.close()


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
    "workbook": "series_bindings_output.xlsx",
    "series": [
        {
            "id": "borvelia_primary_balance",
            "sheet": "Sheet1",
            "data_range": "Sheet1!F5:J5",
            "layout": "series",
            "input": {"setter": {"name": "set_borvelia_primary_balance"}},
            "output": {"compute": {"name": "compute_borvelia_primary_balance"}},
            "structure": deepcopy(_STRUCTURE),
            "key": ["TIME_PERIOD"],
            "series_context": {
                "REF_AREA": "Borvelia",
                "INDICATOR": "Primary balance (% of GDP)",
            },
        }
    ],
}

INVERTED_DOCUMENT: dict[str, Any] = {
    "schema_version": "1.3.0",
    "workbook": "series_bindings_output.xlsx",
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
    path = tmp_path / "series_bindings_output.xlsx"
    _write_inverted_workbook(path)
    return path


def test_generate_modules_emits_output_compute(workbook: Path, tmp_path: Path) -> None:
    bindings = validate_bindings_document(deepcopy(INVERTED_DOCUMENT))
    targets = all_series_targets(bindings, workbook=workbook)
    graph = create_dependency_graph(workbook, targets, load_values=True)
    pkg, modules = load_generated_package(
        graph, bindings, workbook, tmp_path, name="series_output_export"
    )

    joined = "\n".join(modules.values())
    assert "def compute_borvelia_primary_balance_out(" in joined
    assert "def set_borvelia_primary_balance(" not in joined
    assert "def list_setters(" not in joined
    assert "def list_computes(" not in joined
    assert "-> Records:" not in joined
    assert not hasattr(pkg, "make_context")

    result = pkg.compute_borvelia_primary_balance_out(
        borvelia_primary_balance=pkg.data.BorveliaPrimaryBalance.from_records(
            domain=pkg.data.BORVELIA_PRIMARY_BALANCE_DOMAIN,
            records=zip(((1,), (2,), (3,), (4,), (5,)), (1.0, 2.0, 3.0, 7.5, 5.0), strict=True),
        )
    )
    assert tuple(result[period] for period in (1, 2, 3, 4, 5)) == pytest.approx(
        (1.0, 2.0, 3.0, 7.5, 5.0)
    )
    records = pkg.as_records(pkg.compute_borvelia_primary_balance_out, result)
    by_period = {record["TIME_PERIOD"]: record["OBS_VALUE"] for record in records}
    assert by_period[4] == pytest.approx(7.5)
    assert by_period[5] == pytest.approx(5.0)
