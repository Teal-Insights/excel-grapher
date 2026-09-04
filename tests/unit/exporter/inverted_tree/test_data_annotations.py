"""Issue 673 — data.py defaults match compute_* parameter inner types."""

from __future__ import annotations

from pathlib import Path

from excel_grapher.exporter.inverted_tree.ast_emit import (
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


def test_python_annotation_uses_dtype_for_non_formula_series() -> None:
    series = _make("input", "float")
    scalar = _make("input", "int", layout="scalar", n=1)
    assert python_annotation(series) == "Sequence[float]"
    assert python_measure_type(series) == "float | str"
    assert python_annotation(scalar) == "int"
    assert python_measure_type(scalar) == "int | str"


def test_python_data_annotation_matches_compute_param_inner_type() -> None:
    series = _make("input", "float")
    scalar = _make("input", "int", layout="scalar", n=1)
    constant = _make("constant", "float")
    formula = _make("output", "float")
    assert python_data_annotation(series) == "tuple[float, ...]"
    assert python_data_annotation(scalar) == "int"
    assert python_data_annotation(constant) == "tuple[float, ...]"
    assert python_data_annotation(formula) == "tuple[float | str, ...]"
    assert python_annotation(formula) == "Sequence[float | str]"


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
    assert "GROWTH_DEFAULT: tuple[float, ...] =" in data
    assert "COUNT_DEFAULT: int =" in data
    assert "LABELS: tuple[int, ...] =" in data
    assert "float | str" not in data
    assert "int | str" not in data
    assert "growth: Sequence[float]" in api
    assert "count: int" in api
