"""Copying a TIME_PERIOD row out of a matrix is an aligned take, not a scan (#649).

A `layout: series` keyed by `TIME_PERIOD` that reads one row of a
`layout: matrix` keyed by `(SCENARIO, TIME_PERIOD)` with relative refs
(`B4=B2`, `C4=C2`) is an identity join on the inner axis inside one
scenario partition. `_non_peer_seed_ref` must not treat the first host
cell's matrix read as a year-0 seed of the whole matrix sequence.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.deps import predecessor_address
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    input_kwargs,
    inverted_graph_parts,
    load_package,
    oriented_addresses,
    oriented_document,
    series_entry,
    write_oriented_workbook,
)


def _measure(dtype: str = "float") -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": dtype,
        "bind": {"kind": "data_cell", "read": dtype},
    }


def _time_dim(*, header_row: int = 1) -> dict[str, Any]:
    return {
        "id": "TIME_PERIOD",
        "concept": "TIME_PERIOD",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "column_header", "header_row": header_row, "read": "int"},
    }


def _scenario_dim(*, label_column: str = "A") -> dict[str, Any]:
    return {
        "id": "SCENARIO",
        "concept": "SCENARIO",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "row_label", "label_column": label_column, "read": "string"},
    }


def _slice_sheets(
    *, selected_row: int, formulas: tuple[str, str, str]
) -> dict[str, dict[str, object]]:
    cells: dict[str, object] = {
        "B1": 2009,
        "C1": 2010,
        "D1": 2011,
        "A2": "Paris",
        "B2": 4.0,
        "C2": 5.0,
        "D2": 6.0,
        "A3": "Other",
        "B3": 1.0,
        "C3": 2.0,
        "D3": 3.0,
        f"B{selected_row}": formulas[0],
        f"C{selected_row}": formulas[1],
        f"D{selected_row}": formulas[2],
    }
    return {"Engine": cells}


def _slice_bindings(*, selected_range: str) -> dict[str, Any]:
    return bindings_document(
        {
            "id": "shocks",
            "sheet": "Engine",
            "data_range": "Engine!B2:D3",
            "layout": "matrix",
            "input": {"setter": {"name": "set_shocks"}},
            "structure": {
                "measure": _measure(),
                "dimensions": [_scenario_dim(), _time_dim()],
            },
            "key": ["SCENARIO", "TIME_PERIOD"],
        },
        {
            "id": "selected",
            "sheet": "Engine",
            "data_range": selected_range,
            "layout": "series",
            "output": {"compute": {"name": "compute_selected"}},
            "structure": {"measure": _measure(), "dimensions": [_time_dim()]},
            "key": ["TIME_PERIOD"],
        },
    )


def _paris_sheets() -> dict[str, dict[str, object]]:
    return _slice_sheets(selected_row=4, formulas=("=B2", "=C2", "=D2"))


def _other_sheets() -> dict[str, dict[str, object]]:
    return _slice_sheets(selected_row=4, formulas=("=B3", "=C3", "=D3"))


def _paris_bindings() -> dict[str, Any]:
    return _slice_bindings(selected_range="Engine!B4:D4")


def test_matrix_row_slice_is_not_a_scan(tmp_path: Path, orientation: str) -> None:
    workbook = write_oriented_workbook(
        tmp_path / f"a25_paris_{orientation[0]}.xlsx",
        _paris_sheets(),
        orientation=orientation,
    )
    document = oriented_document(_paris_bindings(), orientation)
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    selected = deps["selected"]
    assert selected.is_scan is False
    assert selected.seed_id is None
    assert "shocks" in selected.aligned_ids
    assert selected.index_maps["shocks"] == (0, 1, 2) or selected.affine_maps.get("shocks") == (
        2,
        0,
    )
    assert predecessor_address(catalog.get("selected"), 0, catalog, graph) is None


def test_matrix_row_slice_emits_aligned_take(tmp_path: Path, orientation: str) -> None:
    workbook = write_oriented_workbook(
        tmp_path / f"a25_emit_{orientation[0]}.xlsx",
        _paris_sheets(),
        orientation=orientation,
    )
    internals = generate_inverted(workbook, oriented_document(_paris_bindings(), orientation))[
        "internals.py"
    ]
    assert "prior: float | str = shocks" not in internals
    assert "prior = as_measure(prior)" not in internals
    assert "require_aligned" in internals or "require_length(shocks, 3)" in internals
    assert "shocks[i]" in internals


def test_matrix_row_slice_matches_evaluator(tmp_path: Path, orientation: str) -> None:
    workbook = write_oriented_workbook(
        tmp_path / f"a25_eval_{orientation[0]}.xlsx",
        _paris_sheets(),
        orientation=orientation,
    )
    document = oriented_document(_paris_bindings(), orientation)
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    pkg = load_package(
        generate_inverted(workbook, document), tmp_path, name=f"a25_p_{orientation[0]}"
    )
    cells = oriented_addresses(("Engine!B4", "Engine!C4", "Engine!D4"), orientation)
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, list(cells), load_values=True)
    ).evaluate(list(cells))
    got = pkg.compute_selected(**input_kwargs(catalog, graph))
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx((4.0, 5.0, 6.0))


def test_matrix_other_row_slice_matches_evaluator(tmp_path: Path, orientation: str) -> None:
    workbook = write_oriented_workbook(
        tmp_path / f"a25_other_{orientation[0]}.xlsx",
        _other_sheets(),
        orientation=orientation,
    )
    document = oriented_document(_paris_bindings(), orientation)
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    selected = deps["selected"]
    assert selected.is_scan is False
    assert selected.seed_id is None
    assert "shocks" in selected.aligned_ids
    pkg = load_package(
        generate_inverted(workbook, document), tmp_path, name=f"a25_o_{orientation[0]}"
    )
    cells = oriented_addresses(("Engine!B4", "Engine!C4", "Engine!D4"), orientation)
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, list(cells), load_values=True)
    ).evaluate(list(cells))
    got = pkg.compute_selected(**input_kwargs(catalog, graph))
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx((1.0, 2.0, 3.0))


def test_year0_scalar_seed_is_still_a_scan(tmp_path: Path) -> None:
    """#631 / #649: a relative previous-column scalar remains a year-0 seed."""
    workbook = write_oriented_workbook(
        tmp_path / "a25_seed.xlsx",
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
        orientation="horizontal",
    )
    document = bindings_document(
        series_entry("year0", "Engine!A2", layout="scalar", direction="input"),
        series_entry(
            "path",
            "Engine!B2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )
    _catalog, deps, _graph = inverted_graph_parts(workbook, document)
    path = deps["path"]
    assert path.is_scan is True
    assert path.seed_id == "year0"
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a25_seed")
    assert pkg.compute_path(year0=10.0) == pytest.approx((11.0, 12.0))
