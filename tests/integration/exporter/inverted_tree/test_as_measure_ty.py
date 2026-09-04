"""Inverted-tree `as_measure` default path type-checks as `float | str` (issue 687).

Writes helpers and a generated package outside `tests/` so `ty` still reports
`invalid-argument-type` (that rule is ignored under `tests/**`).
"""

from __future__ import annotations

import subprocess
from pathlib import Path

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    series_entry,
    write_workbook,
)

_REPO_ROOT = Path(__file__).resolve().parents[4]

_HELPER_SNIPPET = """\
from datetime import datetime

from excel_grapher.exporter.inverted_tree.runtime import as_measure


def float_default() -> None:
    out: list[float | str] = []
    out.append(as_measure(1.0))
    out.append(as_measure(1.0, "float"))


def int_measure() -> None:
    out: list[int | str] = []
    out.append(as_measure(1, "int"))


def str_measure() -> None:
    out: list[str] = []
    out.append(as_measure("n/a", "str"))


def bool_measure() -> None:
    out: list[bool | str] = []
    out.append(as_measure(True, "bool"))


def datetime_measure() -> None:
    out: list[datetime | str] = []
    out.append(as_measure(datetime(2020, 1, 1), "datetime"))
"""


def _run(cmd: list[str], *, cwd: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        cmd,
        cwd=str(cwd),
        capture_output=True,
        text=True,
        check=False,
    )


def _ty_check(*paths: Path) -> subprocess.CompletedProcess[str]:
    return _run(
        [
            "uv",
            "run",
            "--no-sync",
            "ty",
            "check",
            "--project",
            str(_REPO_ROOT),
            "--extra-search-path",
            str(paths[0].parent),
            *[str(path) for path in paths],
        ],
        cwd=_REPO_ROOT,
    )


def _add_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "as_measure_add.xlsx",
        {
            "S": {
                "A1": 1,
                "B1": 2,
                "C1": 3,
                "A2": 10.0,
                "B2": 20.0,
                "C2": 30.0,
                "A3": 1.0,
                "B3": 2.0,
                "C3": 3.0,
                "A4": "=A2+A3",
                "B4": "=B2+B3",
                "C4": "=C2+C3",
            }
        },
    )


def _add_bindings() -> dict:
    return bindings_document(
        series_entry("xs", "S!A2:C2", layout="series", direction="input", header_row=1),
        series_entry("ys", "S!A3:C3", layout="series", direction="input", header_row=1),
        series_entry("out", "S!A4:C4", layout="series", direction="output", header_row=1),
    )


def test_as_measure_overloads_type_check_against_measure_lists(tmp_path: Path) -> None:
    snippet = tmp_path / "as_measure_helpers.py"
    snippet.write_text(_HELPER_SNIPPET, encoding="utf-8")

    ty = _ty_check(snippet)
    assert ty.returncode == 0, f"ty failed:\n{ty.stdout}\n{ty.stderr}"
    assert ty.stderr.strip() == ""


def test_inverted_tree_float_helper_append_type_checks(tmp_path: Path) -> None:
    modules = generate_inverted(_add_workbook(tmp_path), _add_bindings())
    internals = modules["internals.py"]
    assert "out: list[float | str] = []" in internals
    assert "out.append(as_measure(" in internals

    pkg = tmp_path / "inv_as_measure"
    pkg.mkdir()
    for name in ("__init__.py", "runtime.py", "internals.py", "data.py"):
        (pkg / name).write_text(modules[name], encoding="utf-8")

    ty = _run(
        [
            "uv",
            "run",
            "--no-sync",
            "ty",
            "check",
            "--project",
            str(_REPO_ROOT),
            "--extra-search-path",
            str(tmp_path),
            "--ignore",
            "unresolved-attribute",
            str(pkg / "runtime.py"),
            str(pkg / "internals.py"),
        ],
        cwd=_REPO_ROOT,
    )
    assert ty.returncode == 0, f"ty failed:\n{ty.stdout}\n{ty.stderr}"
    assert ty.stderr.strip() == ""
