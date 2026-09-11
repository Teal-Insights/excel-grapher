"""Generated input checks live in `validation.py`, not the public `api.py`."""

from __future__ import annotations

import subprocess
from pathlib import Path

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_input_domain import (
    _enum_flag_bindings,
    _enum_flag_workbook,
)

_REPO_ROOT = Path(__file__).resolve().parents[4]


def _two_flag_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "two_flags.xlsx",
        {
            "Inputs": {"A1": 0, "A2": 1},
            "Outputs": {"A1": "=Inputs!A1", "A2": "=Inputs!A2"},
        },
    )


def _two_flag_bindings() -> dict:
    return bindings_document(
        series_entry(
            "flag",
            "Inputs!A1",
            layout="scalar",
            direction="input",
            dtype="int",
            domain={"enum": [0, 1]},
        ),
        series_entry(
            "other",
            "Inputs!A2",
            layout="scalar",
            direction="input",
            dtype="int",
            domain={"enum": [0, 1]},
        ),
        series_entry("out_a", "Outputs!A1", layout="scalar", direction="output", dtype="int"),
        series_entry("out_b", "Outputs!A2", layout="scalar", direction="output", dtype="int"),
    )


def _run(cmd: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        cmd,
        cwd=str(_REPO_ROOT),
        capture_output=True,
        text=True,
        check=False,
    )


def test_input_checks_are_emitted_in_validation_not_api(tmp_path: Path) -> None:
    modules = generate_inverted(_enum_flag_workbook(tmp_path), _enum_flag_bindings())
    assert "validation.py" in modules
    assert "def _check_flag(" in modules["validation.py"]
    assert "require_input_domain(flag" in modules["validation.py"]
    assert "CHECKS = {" in modules["validation.py"]
    assert "def _check_" not in modules["api.py"]
    assert "require_input_domain" not in modules["api.py"]
    assert "validation.CHECKS.get(name)" in modules["api.py"]
    assert "require_input_domain" not in modules["internals.py"]


def test_check_functions_are_separated_by_two_blank_lines(tmp_path: Path) -> None:
    modules = generate_inverted(_two_flag_workbook(tmp_path), _two_flag_bindings())
    validation = modules["validation.py"]
    assert "def _check_flag(" in validation
    assert "def _check_other(" in validation
    assert "\n\n\ndef _check_other(" in validation
    assert "\n\n\nCHECKS = {" in validation


def test_validation_module_passes_ruff_check_and_format(tmp_path: Path) -> None:
    modules = generate_inverted(_two_flag_workbook(tmp_path), _two_flag_bindings())
    pkg = tmp_path / "inv_validation"
    pkg.mkdir()
    for name, content in modules.items():
        (pkg / name).write_text(content, encoding="utf-8")
    target = str(pkg / "validation.py")
    check = _run(["uv", "run", "--no-sync", "ruff", "check", target])
    assert check.returncode == 0, f"ruff check failed:\n{check.stdout}\n{check.stderr}"
    fmt = _run(["uv", "run", "--no-sync", "ruff", "format", "--check", target])
    assert fmt.returncode == 0, f"ruff format --check failed:\n{fmt.stdout}\n{fmt.stderr}"


def test_model_still_validates_inputs_through_validation_module(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_enum_flag_workbook(tmp_path), _enum_flag_bindings()),
        tmp_path,
        name="validation_runtime",
    )
    assert pkg.compute_out(flag=0) == 0
    with pytest.raises(ValueError, match=r"flag out of domain"):
        pkg.compute_out(flag=2)
