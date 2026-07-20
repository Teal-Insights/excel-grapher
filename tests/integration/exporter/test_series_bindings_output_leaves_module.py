"""Modular export routes `_OUTPUT_LEAVES_*` tables into `_output_leaves.py`."""

from __future__ import annotations

import importlib
import subprocess
import sys
from copy import deepcopy
from pathlib import Path
from typing import Any

import pytest
import xlsxwriter

from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import expand_data_range, validate_bindings_document


def _write_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write("A2", "Borvelia")
    ws.write("A5", "Primary balance (% of GDP)")
    for col, year in enumerate([1, 2, 3, 4, 5], start=5):
        ws.write(0, col, year)
        ws.write_number(4, col, float(year))
    ws.write_formula("G5", "=F5+1")
    wb.close()


BINDINGS_DOCUMENT: dict[str, Any] = {
    "schema_version": "1.10.0",
    "workbook": "series_bindings_output_leaves.xlsx",
    "series": [
        {
            "id": "borvelia_primary_balance",
            "sheet": "Sheet1",
            "data_range": "Sheet1!F5:J5",
            "layout": "series",
            "input": {"setter": {"name": "set_borvelia_primary_balance"}},
            "output": {
                "compute": {
                    "name": "compute_borvelia_primary_balance",
                    "helper": {"name": "borvelia_primary_balance_hot", "dims": ["TIME_PERIOD"]},
                }
            },
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
def workbook(tmp_path: Path) -> Path:
    path = tmp_path / "series_bindings_output_leaves.xlsx"
    _write_workbook(path)
    return path


def _generate(workbook: Path) -> dict[str, str]:
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    targets = expand_data_range("Sheet1!F5:J5", workbook=workbook) + ["Sheet1!G5"]
    graph = create_dependency_graph(workbook, targets, load_values=True)
    with CodeGenerator(graph) as gen:
        return gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )


def test_modular_export_emits_private_output_leaves_module(workbook: Path) -> None:
    files = _generate(workbook)

    assert "_output_leaves.py" in files
    leaves = files["_output_leaves.py"]
    assert "_OUTPUT_LEAVES_BORVELIA_PRIMARY_BALANCE" in leaves
    assert "def compute_borvelia_primary_balance(" not in leaves

    api = files["api.py"]
    assert "from ._output_leaves import" in api
    assert "_OUTPUT_LEAVES_BORVELIA_PRIMARY_BALANCE =" not in api
    assert "def compute_borvelia_primary_balance(" in api
    assert 'borvelia_primary_balance_hot(ctx, time_period=static_record["TIME_PERIOD"])' not in api
    assert "borvelia_primary_balance_hot(ctx, time_period=static_record['TIME_PERIOD'])" in api
    assert "from .internals import" in api
    assert "borvelia_primary_balance_hot" in api


def test_output_leaves_module_round_trips_with_helper(workbook: Path, tmp_path: Path) -> None:
    files = _generate(workbook)
    # Consumer-style parameterized helper covering the published span.
    files["internals.py"] += (
        "\n\ndef borvelia_primary_balance_hot(ctx, *, time_period):\n"
        "    from .runtime import xl_cell as _xl_cell\n"
        "    address = {\n"
        "        1: 'Sheet1!F5',\n"
        "        2: 'Sheet1!G5',\n"
        "        3: 'Sheet1!H5',\n"
        "        4: 'Sheet1!I5',\n"
        "        5: 'Sheet1!J5',\n"
        "    }[time_period]\n"
        "    return _xl_cell(ctx, address)\n"
    )
    pkg_dir = tmp_path / "output_leaves_pkg"
    pkg_dir.mkdir()
    for filename, content in files.items():
        (pkg_dir / filename).write_text(content, encoding="utf-8")

    sys.path.insert(0, str(tmp_path))
    try:
        pkg = importlib.import_module("output_leaves_pkg")
        ctx = pkg.make_context()
        pkg.set_borvelia_primary_balance(ctx, [{"TIME_PERIOD": 4, "OBS_VALUE": 7.5}])
        records = pkg.compute_borvelia_primary_balance(ctx)
        by_year = {r["TIME_PERIOD"]: r["OBS_VALUE"] for r in records}
        assert by_year[4] == 7.5
    finally:
        sys.path.remove(str(tmp_path))
        for name in list(sys.modules):
            if name == "output_leaves_pkg" or name.startswith("output_leaves_pkg."):
                del sys.modules[name]


def test_output_leaves_module_is_ruff_clean(workbook: Path, tmp_path: Path) -> None:
    files = _generate(workbook)
    pkg_root = tmp_path / "ruff_output_leaves_pkg"
    pkg_root.mkdir()
    for filename, content in files.items():
        (pkg_root / filename).write_text(content, encoding="utf-8")
    result = subprocess.run(
        ["uv", "run", "--no-sync", "ruff", "check", str(pkg_root / "_output_leaves.py")],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, (
        f"_output_leaves.py is not Ruff-clean:\n{result.stdout}\n{result.stderr}"
    )
