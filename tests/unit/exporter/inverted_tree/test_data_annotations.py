"""Issue 673 / 689 — data.py defaults match compute_* parameter inner types."""

from __future__ import annotations

import subprocess
from pathlib import Path

from excel_grapher.exporter.inverted_tree.ast_emit import (
    _python_param_inner,
    python_annotation,
    python_data_annotation,
    python_measure_type,
)
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


def test_python_param_inner_uses_measure_type_for_numeric_leaves() -> None:
    series = _make("input", "float")
    constant = _make("constant", "float")
    scalar = _make("input", "int", layout="scalar", n=1)
    labels = _make("constant", "int")
    text = _make("input", "str")
    assert _python_param_inner(series) == python_measure_type(series) == "float | str | None"
    assert _python_param_inner(constant) == "float | str | None"
    assert _python_param_inner(scalar) == "int | str"
    assert _python_param_inner(labels) == "int | str | None"
    assert _python_param_inner(text) == "str | None"


def test_python_annotation_uses_measure_type_for_non_formula_series() -> None:
    series = _make("input", "float")
    scalar = _make("input", "int", layout="scalar", n=1)
    assert python_annotation(series) == "Sequence[float | str | None]"
    assert python_measure_type(series) == "float | str | None"
    assert python_annotation(scalar) == "int | str"
    assert python_measure_type(scalar) == "int | str"


def test_python_data_annotation_matches_compute_param_inner_type() -> None:
    series = _make("input", "float")
    scalar = _make("input", "int", layout="scalar", n=1)
    constant = _make("constant", "float")
    formula = _make("output", "float")
    assert python_data_annotation(series) == "tuple[float | str | None, ...]"
    assert python_data_annotation(scalar) == "int | str"
    assert python_data_annotation(constant) == "tuple[float | str | None, ...]"
    assert python_data_annotation(formula) == "tuple[float | str | None, ...]"
    assert python_annotation(formula) == "Sequence[float | str | None]"
    assert python_annotation(series) == "Sequence[float | str | None]"
    assert python_annotation(constant) == "Sequence[float | str | None]"


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
    api = modules["api.py"]
    assert "GROWTH_DEFAULT = Growth[float | str | None].from_legacy(" in data
    assert "COUNT_DEFAULT = 3" in data
    assert "LABELS = Labels[int | str | None].from_legacy(" in data
    assert "growth: data.Growth[float | str | None]" in api
    assert "count: int | str" in api


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
    assert "STORE = Store[float | str | None].from_legacy(" in data
    assert "'n/a'" in data
    assert "store: data.Store[float | str | None]" in internals
    assert "Sequence[" not in internals


def _run_ty(target: Path) -> subprocess.CompletedProcess[str]:
    repo_root = Path(__file__).resolve().parents[4]
    return subprocess.run(
        [
            "uv",
            "run",
            "--no-sync",
            "ty",
            "check",
            "--extra-search-path",
            str(target.parent),
            "--project",
            str(repo_root),
            str(target),
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
    ty = _run_ty(package)
    assert ty.returncode == 0, f"ty failed:\n{ty.stdout}\n{ty.stderr}"
