"""Excel `NORMDIST` must lower through inverted-tree `excel.py` (#995)."""

from __future__ import annotations

import types
from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    invoke_public_compute,
    load_package,
    series_entry,
    write_workbook,
)


def _normdist_workbook(path: Path, formula: str, *, z: float = 0) -> Path:
    return write_workbook(path, {"out": {"A1": z, "B1": formula}})


def _normdist_bindings() -> dict:
    return bindings_document(
        series_entry("z", "out!A1", layout="scalar", direction="input"),
        series_entry("distress_probability", "out!B1", layout="scalar", direction="output"),
    )


def _scalar_workbook(path: Path, formula: str) -> Path:
    return write_workbook(path, {"out": {"A1": formula}})


def _scalar_bindings() -> dict:
    return bindings_document(
        series_entry("distress_probability", "out!A1", layout="scalar", direction="output"),
    )


def _sum_workbook(path: Path, *, left: float, right: float) -> Path:
    return write_workbook(
        path,
        {"out": {"A1": left, "B1": right, "C1": "=NORMDIST(A1+B1,0,1,1)*100"}},
    )


def _sum_bindings() -> dict:
    return bindings_document(
        series_entry("left", "out!A1", layout="scalar", direction="input"),
        series_entry("right", "out!B1", layout="scalar", direction="input"),
        series_entry("distress_probability", "out!C1", layout="scalar", direction="output"),
    )


def _evaluator_value(workbook: Path, cell: str) -> object:
    return FormulaEvaluator(create_dependency_graph(workbook, [cell], load_values=True)).evaluate(
        [cell]
    )[cell]


def _assert_matches_evaluator(
    pkg: types.ModuleType, workbook: Path, cell: str, **kwargs: object
) -> object:
    got = invoke_public_compute(pkg, pkg.compute_distress_probability, kwargs)
    expected = _evaluator_value(workbook, cell)
    if isinstance(expected, float) and isinstance(got, float):
        assert got == pytest.approx(expected)
    else:
        assert got == expected
    return got


def test_normdist_lowers_true_cumulative_to_xl_normdist(tmp_path: Path) -> None:
    workbook = _scalar_workbook(tmp_path / "normdist_scalar.xlsx", "=NORMDIST(0,0,1,TRUE)")
    modules = generate_inverted(workbook, _scalar_bindings())
    assert "xl_normdist(0, 0, 1, True)" in modules["internals.py"]
    assert "def xl_normdist" in modules["excel.py"]
    pkg = load_package(modules, tmp_path, name="normdist_true")
    assert invoke_public_compute(pkg, pkg.compute_distress_probability, {}) == pytest.approx(0.5)


def test_normdist_cumulative_one_matches_evaluator(tmp_path: Path) -> None:
    formula = "=NORMDIST(A1,0,1,1)*100"
    workbook = _normdist_workbook(tmp_path / "normdist.xlsx", formula, z=0)
    modules = generate_inverted(workbook, _normdist_bindings())
    assert "xl_mul(xl_normdist(z, 0, 1, 1), 100)" in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="normdist_cdf")
    assert _assert_matches_evaluator(pkg, workbook, "out!B1", z=0) == pytest.approx(50.0)
    shifted = _normdist_workbook(tmp_path / "normdist_z.xlsx", formula, z=1.96)
    _assert_matches_evaluator(pkg, shifted, "out!B1", z=1.96)


def test_normdist_density_matches_evaluator(tmp_path: Path) -> None:
    formula = "=NORMDIST(A1,0,1,FALSE)"
    workbook = _normdist_workbook(tmp_path / "normdist_pdf.xlsx", formula, z=0)
    pkg = load_package(
        generate_inverted(workbook, _normdist_bindings()),
        tmp_path,
        name="normdist_pdf",
    )
    _assert_matches_evaluator(pkg, workbook, "out!B1", z=0)
    shifted = _normdist_workbook(tmp_path / "normdist_pdf_z.xlsx", formula, z=1.96)
    _assert_matches_evaluator(pkg, shifted, "out!B1", z=1.96)


@pytest.mark.parametrize(
    "formula",
    ["=NORMDIST(A1,0,0,TRUE)", "=NORMDIST(A1,0,-1,1)"],
)
def test_normdist_nonpositive_stdev_matches_evaluator(tmp_path: Path, formula: str) -> None:
    workbook = _normdist_workbook(tmp_path / "normdist_num.xlsx", formula)
    pkg = load_package(
        generate_inverted(workbook, _normdist_bindings()),
        tmp_path,
        name="normdist_num",
    )
    assert _assert_matches_evaluator(pkg, workbook, "out!B1", z=0) == "#NUM!"


def test_normdist_of_sum_matches_evaluator(tmp_path: Path) -> None:
    workbook = _sum_workbook(tmp_path / "normdist_sum.xlsx", left=1.2, right=0.76)
    modules = generate_inverted(workbook, _sum_bindings())
    assert "xl_mul(xl_normdist(xl_add(left, right), 0, 1, 1), 100)" in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="normdist_sum")
    _assert_matches_evaluator(pkg, workbook, "out!C1", left=1.2, right=0.76)
