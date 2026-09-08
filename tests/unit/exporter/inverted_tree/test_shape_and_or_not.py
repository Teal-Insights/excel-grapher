"""Inverted-tree AND/OR/NOT runtime helpers and range arguments (#765)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    call_compute,
    generate_inverted,
    input_kwargs,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)


def _scalar(value: object) -> object:
    if isinstance(value, tuple):
        assert len(value) == 1
        return value[0]
    return value


def _scalar_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("out", "Engine!A1", layout="scalar", direction="output"),
    )


def _package_matches_output(
    tmp_path: Path,
    workbook: Path,
    document: dict[str, Any],
    name: str,
    cell: str,
) -> None:
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name=name)
    expected = FormulaEvaluator(graph).evaluate([cell])[cell]
    series_id = catalog.series_id_for(cell)
    assert series_id is not None
    got = call_compute(pkg, series_id, input_kwargs(catalog, graph))
    assert _scalar(got) == expected


def test_if_and_true_false_matches_mcve(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "and_mcve.xlsx",
        {"Engine": {"A1": "=IF(AND(TRUE(),FALSE()),10,20)"}},
    )
    modules = generate_inverted(workbook, _scalar_bindings())
    assert "xl_and(" in modules["internals.py"]
    assert "def xl_and" in modules["runtime.py"]
    pkg = load_package(modules, tmp_path, name="and_mcve")
    assert pkg.compute_out() == pytest.approx((20.0,))
    _package_matches_output(tmp_path, workbook, _scalar_bindings(), "and_mcve_eval", "Engine!A1")


def test_if_or_and_not_scalars_match_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "or_not.xlsx",
        {
            "Engine": {
                "A1": "=IF(OR(FALSE(),TRUE()),10,20)",
                "A2": "=IF(NOT(TRUE()),10,20)",
                "A3": "=IF(NOT(FALSE()),10,20)",
            }
        },
    )
    document = bindings_document(
        series_entry("or_out", "Engine!A1", layout="scalar", direction="output"),
        series_entry("not_true", "Engine!A2", layout="scalar", direction="output"),
        series_entry("not_false", "Engine!A3", layout="scalar", direction="output"),
    )
    modules = generate_inverted(workbook, document)
    assert "xl_or(" in modules["internals.py"]
    assert "xl_not(" in modules["internals.py"]
    assert "def xl_or" in modules["runtime.py"]
    assert "def xl_not" in modules["runtime.py"]
    pkg = load_package(modules, tmp_path, name="or_not")
    assert pkg.compute_or_out() == pytest.approx((10.0,))
    assert pkg.compute_not_true() == pytest.approx((20.0,))
    assert pkg.compute_not_false() == pytest.approx((10.0,))
    for cell in ("Engine!A1", "Engine!A2", "Engine!A3"):
        _package_matches_output(tmp_path, workbook, document, f"or_not_{cell[-2:]}", cell)


def test_and_or_over_bound_series_match_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "and_or_range.xlsx",
        {
            "Inputs": {"A1": 2024, "B1": 2025, "A2": True, "B2": False},
            "Outputs": {
                "Z1": "=AND(Inputs!A2:B2)",
                "Z2": "=OR(Inputs!A2:B2)",
            },
        },
    )
    document = bindings_document(
        series_entry(
            "flags",
            "Inputs!A2:B2",
            layout="series",
            direction="input",
            dtype="bool",
            header_row=1,
        ),
        series_entry("and_out", "Outputs!Z1", layout="scalar", direction="output", dtype="bool"),
        series_entry("or_out", "Outputs!Z2", layout="scalar", direction="output", dtype="bool"),
    )
    modules = generate_inverted(workbook, document)
    assert "xl_and(" in modules["internals.py"]
    assert "xl_or(" in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="and_or_range")
    assert pkg.compute_and_out(flags=(True, False)) == (False,)
    assert pkg.compute_or_out(flags=(True, False)) == (True,)
    _package_matches_output(tmp_path, workbook, document, "and_range_eval", "Outputs!Z1")
    _package_matches_output(tmp_path, workbook, document, "or_range_eval", "Outputs!Z2")


def test_and_range_window_takes_only_the_range(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "and_window.xlsx",
        {
            "Inputs": {
                "A1": 2024,
                "B1": 2025,
                "C1": 2026,
                "A2": True,
                "B2": True,
                "C2": False,
            },
            "Outputs": {"Z1": "=AND(Inputs!A2:B2)"},
        },
    )
    document = bindings_document(
        series_entry(
            "flags",
            "Inputs!A2:C2",
            layout="series",
            direction="input",
            dtype="bool",
            header_row=1,
        ),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output", dtype="bool"),
    )
    modules = generate_inverted(workbook, document)
    assert "take(" in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="and_window")
    assert pkg.compute_out(flags=(True, True, False)) == (True,)
    _package_matches_output(tmp_path, workbook, document, "and_window_eval", "Outputs!Z1")


def test_and_mixed_range_and_scalar_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "and_mixed.xlsx",
        {
            "Inputs": {"A1": 2024, "B1": 2025, "A2": True, "B2": True},
            "Outputs": {"Z1": "=AND(Inputs!A2:B2,TRUE())"},
        },
    )
    document = bindings_document(
        series_entry(
            "flags",
            "Inputs!A2:B2",
            layout="series",
            direction="input",
            dtype="bool",
            header_row=1,
        ),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output", dtype="bool"),
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="and_mixed")
    assert pkg.compute_out(flags=(True, True)) == (True,)
    _package_matches_output(tmp_path, workbook, document, "and_mixed_eval", "Outputs!Z1")
