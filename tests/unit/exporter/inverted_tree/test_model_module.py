"""Generated `Model` lives in `model.py`; `compute_*` stays in `api.py`."""

from __future__ import annotations

import inspect
import subprocess
from pathlib import Path

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)

_REPO_ROOT = Path(__file__).resolve().parents[4]


def _two_cell_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "two_cell.xlsx",
        {"Inputs": {"A1": 2.0, "B1": "=A1*3"}},
    )


def _two_cell_bindings() -> dict:
    return bindings_document(
        series_entry("seed", "Inputs!A1", layout="scalar", direction="input"),
        series_entry(
            "result",
            "Inputs!B1",
            layout="scalar",
            direction="output",
            compute_name="compute_result",
        ),
    )


def _write_package(modules: dict[str, str], tmp_path: Path, name: str) -> Path:
    package = tmp_path / name
    package.mkdir()
    for filename, content in modules.items():
        (package / filename).write_text(content, encoding="utf-8")
    return package


def _run(cmd: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        cmd,
        cwd=str(_REPO_ROOT),
        capture_output=True,
        text=True,
        check=False,
    )


def test_model_is_emitted_in_model_module(tmp_path: Path) -> None:
    modules = generate_inverted(_two_cell_workbook(tmp_path), _two_cell_bindings())
    assert "model.py" in modules
    assert "class Model" in modules["model.py"]
    assert "class Model" not in modules["api.py"]
    assert "from .model import Model" not in modules["api.py"]
    assert "def compute_result" in modules["api.py"]
    assert "def compute_result" not in modules["model.py"]
    assert "model.Model(**locals())" in modules["api.py"]

    pkg = load_package(modules, tmp_path, name="model_module")
    assert pkg.compute_result(seed=2.0) == 6.0
    assert pkg.compute_result(seed=3.0) == 9.0
    assert pkg.model.__all__ == ["Model"]
    assert "Model" not in pkg.api.__all__
    assert not hasattr(pkg.api, "Model")
    assert "Model" in pkg.__all__
    assert pkg.Model is pkg.model.Model
    assert pkg.Model.__module__ == "model_module.model"
    assert not hasattr(pkg.Model, "compute_result")
    source = inspect.getsource(pkg.compute_result)
    assert "model.Model(**locals())" in source
    first = pkg.Model(seed=2.0)
    second = pkg.Model(seed=2.0)
    assert first is not second
    assert first.result == 6.0
    assert second.result == 6.0


def test_generated_model_and_api_pass_ruff_and_ty(tmp_path: Path) -> None:
    modules = generate_inverted(_two_cell_workbook(tmp_path), _two_cell_bindings())
    package = _write_package(modules, tmp_path, "inv_model_lint")
    for filename in ("api.py", "model.py"):
        target = str(package / filename)
        check = _run(["uv", "run", "--no-sync", "ruff", "check", target])
        assert check.returncode == 0, f"ruff check {filename}:\n{check.stdout}\n{check.stderr}"
        fmt = _run(["uv", "run", "--no-sync", "ruff", "format", "--check", target])
        assert fmt.returncode == 0, f"ruff format {filename}:\n{fmt.stdout}\n{fmt.stderr}"
    ty = _run(
        [
            "uv",
            "run",
            "--no-sync",
            "ty",
            "check",
            "--extra-search-path",
            str(package.parent),
            "--project",
            str(_REPO_ROOT),
            "--ignore",
            "unresolved-attribute",
            str(package / "api.py"),
            str(package / "model.py"),
        ]
    )
    assert ty.returncode == 0, f"ty failed:\n{ty.stdout}\n{ty.stderr}"
