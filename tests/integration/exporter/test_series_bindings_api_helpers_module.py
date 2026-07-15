"""Modular export keeps `api.py` lean by inlining coercion into `_api_helpers.py`.

The series-binding setter machinery (input coercion, record-apply helpers, and the
type aliases they need) is verbose. In the multi-module export it is routed to a
dedicated private `_api_helpers` module so that `api.py` contains only the public
surface (``make_context``, ``compute_all``, and generated ``set_*`` / ``compute_*``).
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
    "workbook": "series_bindings_helpers.xlsx",
    "series": [
        {
            "id": "borvelia_primary_balance",
            "sheet": "Sheet1",
            "data_range": "Sheet1!F5:J5",
            "layout": "series",
            "input": {"setter": {"name": "set_borvelia_primary_balance"}},
            "output": {"compute": {"name": "compute_borvelia_primary_balance"}},
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
    path = tmp_path / "series_bindings_helpers.xlsx"
    _write_workbook(path)
    return path


def _generate(workbook: Path) -> dict[str, str]:
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    targets = expand_data_range("Sheet1!F5:J5", workbook=workbook) + ["Sheet1!G5"]
    graph = create_dependency_graph(workbook, targets, load_values=True)
    return CodeGenerator(graph).generate_modules(
        targets,
        series_bindings=bindings,
        bindings_workbook=workbook,
    )


def test_coercion_helpers_live_in_dedicated_module(workbook: Path) -> None:
    files = _generate(workbook)

    assert "_api_helpers.py" in files
    helpers = files["_api_helpers.py"]
    api = files["api.py"]

    # The heavy coercion machinery moved out of the public api module.
    assert "def coerce_setter_input(" in helpers
    assert "def _apply_series_records(" in helpers
    assert "def coerce_scalar(" in helpers
    assert "def coerce_setter_input(" not in api
    assert "def _apply_series_records(" not in api
    assert "def coerce_scalar(" not in api

    # api.py keeps only the public surface plus an import from the helper module.
    assert "from ._api_helpers import" in api
    assert "def make_context(" in api
    assert "def set_borvelia_primary_balance(" in api
    assert "def compute_borvelia_primary_balance(" in api
    assert "def compute_all(" in api


def test_api_module_is_smaller_than_helper_module(workbook: Path) -> None:
    files = _generate(workbook)
    api_lines = files["api.py"].count("\n")
    helper_lines = files["_api_helpers.py"].count("\n")
    # The bulk of the generated lines should now be in the helper module.
    assert helper_lines > api_lines


def test_modules_execute_and_round_trip(workbook: Path, tmp_path: Path) -> None:
    files = _generate(workbook)
    pkg_dir = tmp_path / "exported_helpers"
    pkg_dir.mkdir(parents=True, exist_ok=True)
    for filename, content in files.items():
        (pkg_dir / filename).write_text(content, encoding="utf-8")

    sys.path.insert(0, str(tmp_path))
    try:
        pkg = importlib.import_module("exported_helpers")
        ctx = pkg.make_context()
        pkg.set_borvelia_primary_balance(ctx, [{"TIME_PERIOD": 4, "OBS_VALUE": 7.5}])
        assert ctx.inputs["Sheet1!I5"] == 7.5
        records = pkg.compute_borvelia_primary_balance(ctx)
        by_year = {r["TIME_PERIOD"]: r["OBS_VALUE"] for r in records}
        assert by_year[4] == 7.5
    finally:
        sys.path.remove(str(tmp_path))
        for name in list(sys.modules):
            if name == "exported_helpers" or name.startswith("exported_helpers."):
                sys.modules.pop(name, None)


def test_helpers_module_is_ruff_clean_and_package_type_checks(
    workbook: Path,
    tmp_path: Path,
) -> None:
    """The emitted `_api_helpers.py` is Ruff-clean and the whole package type-checks.

    Ruff is run without `--fix`: this asserts what codegen actually emits, not what an
    auto-fixer could repair. The check is scoped to `_api_helpers.py` (the module this
    split introduces); whole-package raw-Ruff hygiene (e.g. `api.py` import ordering) is
    tracked separately by issues #252/#253 and is intentionally not asserted here.
    """
    # Generated helpers TYPE_CHECKING-import pandas/polars; ty needs them installed.
    pytest.importorskip("pandas")
    pytest.importorskip("polars")
    repo_root = Path(__file__).resolve().parents[3]
    files = _generate(workbook)
    pkg_root = tmp_path / "exported"
    pkg_root.mkdir(parents=True, exist_ok=True)
    for filename, content in files.items():
        (pkg_root / filename).write_text(content, encoding="utf-8")

    def _run(cmd: list[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.run(cmd, cwd=str(repo_root), capture_output=True, text=True, check=False)

    ruff = _run(["uv", "run", "--no-sync", "ruff", "check", str(pkg_root / "_api_helpers.py")])
    assert ruff.returncode == 0, f"_api_helpers.py is not Ruff-clean:\n{ruff.stdout}\n{ruff.stderr}"

    ty = _run(
        [
            "uv",
            "run",
            "--no-sync",
            "ty",
            "check",
            "--project",
            str(repo_root),
            "--extra-search-path",
            str(tmp_path),
            str(pkg_root),
        ]
    )
    assert ty.returncode == 0, f"ty failed:\n{ty.stdout}\n{ty.stderr}"
