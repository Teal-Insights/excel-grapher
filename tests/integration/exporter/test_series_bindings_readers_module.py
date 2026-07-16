"""Modular export routes readers and leaf maps into `_readers.py`.

Public `read_*` names stay re-exported through `api.py` / `__init__.py` so the
package surface is unchanged, while formula bodies in `internals.py` can import
readers without forming an `api` ↔ `internals` cycle.
"""

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
    "schema_version": "1.3.0",
    "workbook": "series_bindings_readers.xlsx",
    "series": [
        {
            "id": "borvelia_primary_balance",
            "sheet": "Sheet1",
            "data_range": "Sheet1!F5:J5",
            "layout": "series",
            "input": {"setter": {"name": "set_borvelia_primary_balance"}},
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
        }
    ],
}


@pytest.fixture
def workbook(tmp_path: Path) -> Path:
    path = tmp_path / "series_bindings_readers.xlsx"
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


def test_modular_export_emits_private_readers_module(workbook: Path) -> None:
    files = _generate(workbook)

    assert "_readers.py" in files
    readers = files["_readers.py"]
    assert "def read_borvelia_primary_balance(" in readers
    assert "def read_borvelia_primary_balance_range(" in readers
    assert "_LEAF_INDEX_BORVELIA_PRIMARY_BALANCE" in readers

    api = files["api.py"]
    assert "from ._readers import" in api
    assert "_LEAF_INDEX_BORVELIA_PRIMARY_BALANCE" in api
    assert "def read_borvelia_primary_balance(" not in api
    assert "def set_borvelia_primary_balance(" in api
    # Public readers are re-exported from `_readers` (not via unused api imports).
    assert "read_borvelia_primary_balance" not in api.split("def set_borvelia_primary_balance")[0]

    init = files["__init__.py"]
    assert "from ._readers import" in init
    assert "read_borvelia_primary_balance" in init


def test_modular_readers_module_is_importable_and_usable(
    workbook: Path,
    tmp_path: Path,
) -> None:
    files = _generate(workbook)
    pkg_dir = tmp_path / "readers_pkg"
    pkg_dir.mkdir()
    for filename, content in files.items():
        (pkg_dir / filename).write_text(content, encoding="utf-8")

    sys.path.insert(0, str(tmp_path))
    try:
        pkg = importlib.import_module("readers_pkg")
        ctx = pkg.make_context()
        pkg.set_borvelia_primary_balance(ctx, [{"TIME_PERIOD": 1, "OBS_VALUE": 9.5}])
        assert pkg.read_borvelia_primary_balance(ctx, time_period=1) == 9.5
        assert pkg.list_readers() == ["read_borvelia_primary_balance"]
    finally:
        sys.path.remove(str(tmp_path))
        for name in list(sys.modules):
            if name == "readers_pkg" or name.startswith("readers_pkg."):
                del sys.modules[name]


def test_readers_module_is_ruff_clean(workbook: Path, tmp_path: Path) -> None:
    files = _generate(workbook)
    pkg_root = tmp_path / "ruff_readers_pkg"
    pkg_root.mkdir()
    for filename, content in files.items():
        (pkg_root / filename).write_text(content, encoding="utf-8")

    # Scope to modules this PR owns. `internals.py` still uses legacy `'''Formula'''`
    # docstrings (D300) outside this change.
    targets = [
        pkg_root / "_readers.py",
        pkg_root / "api.py",
        pkg_root / "__init__.py",
    ]
    ruff = subprocess.run(
        ["uv", "run", "--no-sync", "ruff", "check", *[str(p) for p in targets]],
        check=False,
        capture_output=True,
        text=True,
    )
    assert ruff.returncode == 0, f"generated package is not Ruff-clean:\n{ruff.stdout}\n{ruff.stderr}"
