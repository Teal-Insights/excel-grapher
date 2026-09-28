"""Generated inverted-tree packages pass ruff and ty."""

from __future__ import annotations

import json
import subprocess
from pathlib import Path

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    series_entry,
    write_workbook,
)

_REPO_ROOT = Path(__file__).resolve().parents[4]


def _run(cmd: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        cmd,
        cwd=str(_REPO_ROOT),
        capture_output=True,
        text=True,
        check=False,
    )


def test_generate_modules_package_has_no_ty_or_ruff_diagnostics(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "diag.xlsx",
        {"S": {"A1": 2, "B1": "=A1*2"}},
    )
    document = bindings_document(
        series_entry("src", "S!A1", layout="scalar", direction="input"),
        series_entry("out", "S!B1", layout="scalar", direction="output"),
    )
    _assert_clean_package(tmp_path, generate_inverted(workbook, document))


def test_formula_series_domains_type_check_as_bound(tmp_path: Path) -> None:
    """Formula series with a concrete domain expose `Domain`, not a template union."""
    cells: dict[str, object] = {"A2": "rate", "A3": "growth", "A4": "level", "F1": 2.0}
    for col, year in zip("BCD", (2024, 2025, 2026), strict=True):
        cells[f"{col}1"] = year
        cells[f"{col}2"] = 0.05
        cells[f"{col}3"] = f"={col}2*$F$1"
        cells[f"{col}4"] = f"=1+{col}3" if col == "B" else f"={chr(ord(col) - 1)}4*(1+{col}3)"
    # A sparse input and formula over the same years (2025 is a structural blank).
    cells.update({"B5": 0.1, "D5": 0.2, "B6": "=B5*2", "D6": "=D5*2"})
    workbook = write_workbook(tmp_path / "rows.xlsx", {"S": cells})
    row = {"layout": "row_series", "header_row": 1, "key": ["TIME_PERIOD"]}
    document = bindings_document(
        series_entry("level", "S!B4:D4", direction="output", **row),
        series_entry("rate", "S!B2:D2", direction="input", **row),
        series_entry("growth", "S!B3:D3", direction="internal", **row),
        series_entry("shift", "S!B5:D5", direction="input", **row),
        series_entry("shift_x2", "S!B6:D6", direction="output", **row),
        series_entry("multiplier", "S!F1", direction="input"),
    )
    _assert_clean_package(tmp_path, generate_inverted(workbook, document))


def _assert_clean_package(tmp_path: Path, modules: dict[str, str]) -> None:
    pkg = tmp_path / "inv_diag"
    pkg.mkdir()
    for name, content in modules.items():
        (pkg / name).write_text(content, encoding="utf-8")

    ruff_fix = _run(["uv", "run", "--no-sync", "ruff", "check", "--fix", str(pkg)])
    assert ruff_fix.returncode == 0, f"ruff --fix failed:\n{ruff_fix.stdout}\n{ruff_fix.stderr}"

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
            str(pkg / "data.py"),
        ]
    )
    assert ty.returncode == 0, f"ty failed:\n{ty.stdout}\n{ty.stderr}"
    assert ty.stderr.strip() == ""

    ruff = _run(["uv", "run", "--no-sync", "ruff", "check", str(pkg)])
    assert ruff.returncode == 0, f"ruff failed after --fix:\n{ruff.stdout}\n{ruff.stderr}"
    assert ruff.stderr.strip() == ""


def test_generated_validation_rebinds_coerced_inputs_without_ty_errors(tmp_path: Path) -> None:
    """`coerce_input_measure` keeps the input's type, so rebinding typed parameters checks."""
    workbook = write_workbook(
        tmp_path / "coerce.xlsx",
        {
            "S": {
                "B1": 2024,
                "C1": 2025,
                "B2": 0.1,
                "C2": 0.2,
                "B3": "=B2*$F$1",
                "C3": "=C2*$F$1",
                "F1": 2.0,
                "G1": "On",
                "G2": '=IF(G1="On",F1,0)',
            }
        },
    )
    document = bindings_document(
        series_entry(
            "rate",
            "S!B2:C2",
            layout="row_series",
            header_row=1,
            domain={"real_between": {"min": 0, "max": 1}},
        ),
        series_entry(
            "level",
            "S!B3:C3",
            layout="row_series",
            header_row=1,
            direction="output",
            compute_name="compute_level",
        ),
        series_entry("multiplier", "S!F1", domain={"real_between": {"min": 0, "max": 10}}),
        series_entry("switch", "S!G1", dtype="string", domain={"enum": ["Off", "On"]}),
        series_entry("flag", "S!G2", direction="output", compute_name="compute_flag"),
    )
    modules = generate_inverted(workbook, document)
    pkg = tmp_path / "inv_coerce"
    pkg.mkdir()
    for name, content in modules.items():
        (pkg / name).write_text(content, encoding="utf-8")
    assert "coerce_input_measure(" in modules["validation.py"]

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
            str(pkg / "validation.py"),
        ]
    )
    assert ty.returncode == 0, f"ty failed:\n{ty.stdout}\n{ty.stderr}"


_SHIPPED_SUPPORT_MODULES = ("excel.py", "runtime.py", "tensor.py", "provenance.py")


def test_shipped_support_modules_pass_ty_and_pyright(tmp_path: Path) -> None:
    """Support modules copied into every package type-check under ty and pyright."""
    # `excel.py` imports NumPy for its fast paths; resolving it needs the package.
    pytest.importorskip("numpy")
    workbook = write_workbook(
        tmp_path / "support.xlsx",
        {"S": {"A1": 2, "B1": "=A1*2"}},
    )
    document = bindings_document(
        series_entry("src", "S!A1", layout="scalar", direction="input"),
        series_entry("out", "S!B1", layout="scalar", direction="output"),
    )
    modules = generate_inverted(workbook, document)
    pkg = tmp_path / "inv_support"
    pkg.mkdir()
    for name, content in modules.items():
        (pkg / name).write_text(content, encoding="utf-8")
    files = [str(pkg / name) for name in _SHIPPED_SUPPORT_MODULES]

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
            *files,
        ]
    )
    assert ty.returncode == 0, f"ty failed:\n{ty.stdout}\n{ty.stderr}"

    pyright = _run(["uv", "run", "--no-sync", "pyright", "--outputjson", *files])
    report = json.loads(pyright.stdout)
    errors = [
        f"{Path(d['file']).name}:{d['range']['start']['line'] + 1}: {d['message']}"
        for d in report["generalDiagnostics"]
        if d["severity"] == "error"
    ]
    assert errors == []
