"""Generated `Model` lives in `model.py`; `compute_*` stays in `api.py`."""

from __future__ import annotations

import inspect
from pathlib import Path

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)


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


def test_model_is_emitted_in_model_module(tmp_path: Path) -> None:
    modules = generate_inverted(_two_cell_workbook(tmp_path), _two_cell_bindings())
    assert "model.py" in modules
    assert "class Model" in modules["model.py"]
    assert "class Model" not in modules["api.py"]
    assert "from .model import Model" in modules["api.py"]
    assert "def compute_result" in modules["api.py"]
    assert "def compute_result" not in modules["model.py"]
    assert "'Model'" in modules["model.py"]
    assert "    'Model'," not in modules["api.py"]

    pkg = load_package(modules, tmp_path, name="model_module")
    assert hasattr(pkg, "compute_result")
    assert pkg.compute_result(seed=2.0) == 6.0
    assert hasattr(pkg.model, "Model")
    assert pkg.model.__all__ == ["Model"]
    assert "Model" not in pkg.api.__all__
    assert "Model" not in pkg.__all__
    assert not hasattr(pkg.model.Model, "compute_result")
    assert pkg.model.Model.__module__ == "model_module.model"
    assert inspect.getmodule(pkg.compute_result).Model is pkg.model.Model
    assert pkg.model.Model(seed=2.0).result == 6.0
