"""Issue 673 / 689 — data.py defaults match compute_* parameter inner types."""

from __future__ import annotations

import subprocess
from pathlib import Path

from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, KeyPoint, Statement
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    series_entry,
    write_workbook,
)


def _make(direction: str, dtype: str, *, layout: str = "series", n: int = 2) -> BoundSeries:
    cells = tuple(f"Inputs!B{i}" for i in range(10, 10 + n))
    domain = tuple(KeyPoint((("idx", i),)) for i in range(n))
    return BoundSeries(
        series_id="demo",
        layout=layout,
        direction=direction,
        cells=cells,
        key_fields=("idx",),
        dtype=dtype,
        compute_name=None,
        raw={},
        domain=domain,
        statements=(Statement("demo", "demo", None, 0, n, cells, domain),),
    )


def _annotation_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "data_annos.xlsx",
        {
            "Inputs": {
                "B1": 1.5,
                "C1": 2.5,
                "A2": 3,
                "B10": 1,
                "C10": 2,
            },
            "Outputs": {
                "A1": "=Inputs!B1",
                "B1": "=Inputs!C1",
                "C1": "=Inputs!A2",
                "A10": 1,
                "B10": 2,
            },
        },
    )


def _annotation_bindings() -> dict:
    return bindings_document(
        series_entry(
            "growth",
            "Inputs!B1:C1",
            layout="series",
            direction="input",
            header_row=10,
        ),
        series_entry("count", "Inputs!A2", layout="scalar", direction="input", dtype="int"),
        series_entry(
            "labels",
            "Inputs!B10:C10",
            layout="series",
            direction="constant",
            dtype="int",
            header_row=10,
        ),
        series_entry(
            "out",
            "Outputs!A1:B1",
            layout="series",
            direction="output",
            header_row=10,
        ),
        series_entry("total", "Outputs!C1", layout="scalar", direction="output", dtype="int"),
    )


def test_emit_data_module_uses_param_inner_types(tmp_path: Path) -> None:
    modules = generate_inverted(_annotation_workbook(tmp_path), _annotation_bindings())
    data = modules["data.py"]
    model = modules["model.py"]
    assert "GROWTH: Series[" in data
    assert "COUNT_DEFAULT = 3" in data
    assert "LABELS: Series[" in data
    assert "class Labels" not in data
    assert "growth: data.Growth" in model
    assert "class Growth" not in data
    assert "Growth = Series[float | str | None]" in data
    assert "count: int | str" in model


def _cached_text_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "cached_text_const.xlsx",
        {
            "Store": {
                "A1": 1,
                "B1": 2,
                "C1": 3,
                "A2": 1.0,
                "B2": "n/a",
                "C2": 3.0,
                "A3": "=A2",
            },
        },
    )


def _cached_text_bindings() -> dict:
    return bindings_document(
        series_entry(
            "store",
            "Store!A2:C2",
            layout="series",
            direction="constant",
            header_row=1,
        ),
        series_entry("out", "Store!A3", layout="scalar", direction="output"),
    )


def test_cached_text_constant_emits_measure_tensor(tmp_path: Path) -> None:
    modules = generate_inverted(_cached_text_workbook(tmp_path), _cached_text_bindings())
    data = modules["data.py"]
    internals = modules["internals.py"]
    assert "STORE: Series[" in data
    assert "'n/a'" in data
    assert "store: data.Series[" in internals
    assert "Sequence[" not in internals


def _run_ty(package: Path, *ignore: str) -> subprocess.CompletedProcess[str]:
    repo_root = Path(__file__).resolve().parents[4]
    return subprocess.run(
        [
            "uv",
            "run",
            "--no-sync",
            "ty",
            "check",
            "--extra-search-path",
            str(package.parent),
            "--project",
            str(repo_root),
            "--ignore",
            "unresolved-attribute",
            *(arg for rule in ignore for arg in ("--ignore", rule)),
            str(package / "data.py"),
            str(package / "internals.py"),
            str(package / "api.py"),
            str(package / "model.py"),
            str(package / "validation.py"),
        ],
        cwd=str(repo_root),
        capture_output=True,
        text=True,
        check=False,
    )


def test_cached_text_constant_and_helper_type_check_together(tmp_path: Path) -> None:
    modules = generate_inverted(_cached_text_workbook(tmp_path), _cached_text_bindings())
    package = tmp_path / "cached_text_package"
    package.mkdir()
    for filename, source in modules.items():
        (package / filename).write_text(source, encoding="utf-8")
    # `invalid-assignment` in validation.py is issue 1011, unrelated to blank defaults.
    ty = _run_ty(package, "invalid-assignment")
    assert ty.returncode == 0, f"ty failed:\n{ty.stdout}\n{ty.stderr}"


def _blank_scalar_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "blank_scalar.xlsx",
        {"S": {"F1": 2.0, "G1": "=F1+F2"}},
    )


def _blank_scalar_bindings() -> dict:
    return bindings_document(
        series_entry("multiplier", "S!F1", layout="scalar", direction="input"),
        series_entry("adjustment", "S!F2", layout="scalar", direction="input"),
        series_entry("total", "S!G1", layout="scalar", direction="output"),
    )


def test_blank_scalar_input_annotation_admits_none(tmp_path: Path) -> None:
    """Issue 1012 — a blank scalar default is `None`, so its annotation must admit it."""
    modules = generate_inverted(_blank_scalar_workbook(tmp_path), _blank_scalar_bindings())
    assert "ADJUSTMENT_DEFAULT = None" in modules["data.py"]
    assert "adjustment: float | str | None = data.ADJUSTMENT_DEFAULT" in modules["model.py"]
    assert "multiplier: float | str = data.MULTIPLIER_DEFAULT" in modules["model.py"]
    assert "def _check_adjustment(adjustment: float | str | None)" in modules["validation.py"]
    package = tmp_path / "blank_scalar_package"
    package.mkdir()
    for filename, source in modules.items():
        (package / filename).write_text(source, encoding="utf-8")
    # `invalid-assignment` in validation.py is issue 1011, unrelated to blank defaults.
    ty = _run_ty(package, "invalid-assignment")
    assert ty.returncode == 0, f"ty failed:\n{ty.stdout}\n{ty.stderr}"
