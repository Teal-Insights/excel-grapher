"""Absolute IF selector in the previous column is not a year-0 scan seed (#631).

A row-series that starts in column B and reads `$A$2` in every member is an
identity zip over a scalar selector, not a recursive path seeded by that cell.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)


def _selector_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a24_selector.xlsx",
        {
            "Engine": {
                "B1": 2009,
                "C1": 2010,
                "D1": 2011,
                "A2": "Nominal",
                "B2": '=IF($A$2=$A$5,B5,IF($A$2=$A$6,B6,""""))',
                "C2": '=IF($A$2=$A$5,C5,IF($A$2=$A$6,C6,""""))',
                "D2": '=IF($A$2=$A$5,D5,IF($A$2=$A$6,D6,""""))',
                "A5": "Nominal",
                "B5": 4.0,
                "C5": 5.0,
                "D5": 6.0,
                "A6": "Other",
                "B6": 1.0,
                "C6": 2.0,
                "D6": 3.0,
            },
        },
    )


def _selector_bindings() -> dict:
    return bindings_document(
        series_entry("mode", "Engine!A2", layout="scalar", direction="input", dtype="string"),
        series_entry(
            "label_nominal",
            "Engine!A5",
            layout="scalar",
            direction="constant",
            dtype="string",
        ),
        series_entry(
            "label_other",
            "Engine!A6",
            layout="scalar",
            direction="constant",
            dtype="string",
        ),
        series_entry(
            "nominal",
            "Engine!B5:D5",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry(
            "other",
            "Engine!B6:D6",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry(
            "selected",
            "Engine!B2:D2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )


def _recursive_seed_workbook(tmp_path: Path) -> Path:
    """True year-0 seed: only the first member reads the previous-column scalar."""
    return write_workbook(
        tmp_path / "a24_recursive.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": 10.0,
                "B2": "=A2+1",
                "C2": "=B2+1",
            },
        },
    )


def _recursive_seed_bindings() -> dict:
    return bindings_document(
        series_entry("year0", "Engine!A2", layout="scalar", direction="input"),
        series_entry(
            "path",
            "Engine!B2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )


def test_absolute_selector_is_not_a_scan_seed(tmp_path: Path) -> None:
    catalog, deps, _graph = inverted_graph_parts(_selector_workbook(tmp_path), _selector_bindings())
    selected = deps["selected"]
    assert selected.is_scan is False
    assert selected.seed_id is None
    assert catalog.get("mode").is_scalar
    assert "mode" in selected.param_ids


def test_absolute_selector_emits_identity_loop_reading_mode(tmp_path: Path) -> None:
    modules = generate_inverted(_selector_workbook(tmp_path), _selector_bindings())
    internals = modules["internals.py"]
    assert "prior: float | str = mode" not in internals
    assert "prior == label_nominal" not in internals
    assert "xl_eq(mode, label_nominal)" in internals
    assert "for i in range(n):" in internals
    assert "out.append(" in internals


def test_absolute_selector_matches_evaluator(tmp_path: Path) -> None:
    workbook = _selector_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _selector_bindings()), tmp_path, name="a24_sel")
    cells = ["Engine!B2", "Engine!C2", "Engine!D2"]
    graph = create_dependency_graph(workbook, cells, load_values=True)
    expected = FormulaEvaluator(graph).evaluate(cells)
    got = pkg.compute_selected(
        mode="Nominal",
        nominal=(4.0, 5.0, 6.0),
        other=(1.0, 2.0, 3.0),
    )
    assert got == pytest.approx(
        (expected["Engine!B2"], expected["Engine!C2"], expected["Engine!D2"])
    )
    assert got == pytest.approx((4.0, 5.0, 6.0))


def test_absolute_selector_other_branch_and_fallback(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_selector_workbook(tmp_path), _selector_bindings()),
        tmp_path,
        name="a24_branches",
    )
    assert pkg.compute_selected(
        mode="Other",
        nominal=(4.0, 5.0, 6.0),
        other=(1.0, 2.0, 3.0),
    ) == pytest.approx((1.0, 2.0, 3.0))
    assert pkg.compute_selected(
        mode="Neither",
        nominal=(4.0, 5.0, 6.0),
        other=(1.0, 2.0, 3.0),
    ) == ('"', '"', '"')


def test_year0_seed_read_only_by_first_member_is_still_a_scan(tmp_path: Path) -> None:
    _catalog, deps, _graph = inverted_graph_parts(
        _recursive_seed_workbook(tmp_path),
        _recursive_seed_bindings(),
    )
    path = deps["path"]
    assert path.is_scan is True
    assert path.seed_id == "year0"
    modules = generate_inverted(_recursive_seed_workbook(tmp_path), _recursive_seed_bindings())
    pkg = load_package(modules, tmp_path, name="a24_seed")
    assert pkg.compute_path(year0=10.0) == pytest.approx((11.0, 12.0))
