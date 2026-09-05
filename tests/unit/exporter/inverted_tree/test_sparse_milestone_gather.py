"""Issue 695 — sparse year picks must gather catalog slots, not `i + offset`.

Consecutive host cells (`Out!D6:F6`) can read non-colinear members of a longer
`(SCENARIO, TIME_PERIOD)` matrix (2050 / 2075 / 2099 → catalog 22, 47, 71).
`fit_affine_map` returns `None`; the helper must still take those slots.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from fastpyxl.utils.cell import get_column_letter

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

YEARS = tuple(range(2028, 2100))
assert YEARS[22] == 2050 and YEARS[47] == 2075 and YEARS[71] == 2099
_LAST = get_column_letter(len(YEARS) + 1)
_C2050 = get_column_letter(YEARS.index(2050) + 2)
_C2075 = get_column_letter(YEARS.index(2075) + 2)
_C2099 = get_column_letter(YEARS.index(2099) + 2)


def _measure() -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }


def _time_dim(*, header_row: int) -> dict[str, Any]:
    return {
        "id": "TIME_PERIOD",
        "concept": "TIME_PERIOD",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "column_header", "header_row": header_row, "read": "int"},
    }


def _scenario_dim() -> dict[str, Any]:
    return {
        "id": "SCENARIO",
        "concept": "SCENARIO",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "row_label", "label_column": "A", "read": "string"},
    }


def _milestone_workbook(tmp_path: Path) -> Path:
    """Engine row 2028–2099; Out!D6:F6 copies 2050 / 2075 / 2099."""
    engine: dict[str, object] = {"A2": "Paris"}
    for col, year in enumerate(YEARS, start=2):
        engine[f"{get_column_letter(col)}1"] = year
        engine[f"{get_column_letter(col)}2"] = f"=Inputs!$A$1+{year}"
    return write_workbook(
        tmp_path / "sparse_milestones.xlsx",
        {
            "Inputs": {"A1": 0},
            "Engine": engine,
            "Out": {
                "D4": 2050,
                "E4": 2075,
                "F4": 2099,
                "D6": f"=Engine!{_C2050}2",
                "E6": f"=Engine!{_C2075}2",
                "F6": f"=Engine!{_C2099}2",
            },
        },
    )


def _milestone_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("rate", "Inputs!A1", layout="scalar", direction="input"),
        {
            "id": "engine_pb",
            "sheet": "Engine",
            "data_range": f"Engine!B2:{_LAST}2",
            "layout": "matrix",
            "internal": {},
            "structure": {
                "measure": _measure(),
                "dimensions": [_scenario_dim(), _time_dim(header_row=1)],
            },
            "key": ["SCENARIO", "TIME_PERIOD"],
        },
        {
            "id": "milestones",
            "sheet": "Out",
            "data_range": "Out!D6:F6",
            "layout": "series",
            "output": {"compute": {"name": "compute_milestones"}},
            "structure": {
                "measure": _measure(),
                "dimensions": [_time_dim(header_row=4)],
            },
            "key": ["TIME_PERIOD"],
        },
    )


def test_sparse_picks_record_literal_index_map(tmp_path: Path) -> None:
    _catalog, deps, _graph = inverted_graph_parts(
        _milestone_workbook(tmp_path), _milestone_bindings()
    )
    picked = deps["milestones"]
    assert picked.index_maps["engine_pb"] == (22, 47, 71)
    assert "engine_pb" in picked.aligned_ids
    assert "engine_pb" not in picked.affine_maps


def test_sparse_picks_emit_take_not_consecutive_offset(tmp_path: Path) -> None:
    workbook = _milestone_workbook(tmp_path)
    modules = generate_inverted(workbook, _milestone_bindings())
    api = modules["api.py"]
    internals = modules["internals.py"]
    assert "take(engine_pb, (22, 47, 71))" in api
    assert "i + 22" not in internals
    assert "engine_pb[i]" in internals


def test_sparse_picks_match_formula_evaluator(tmp_path: Path) -> None:
    workbook = _milestone_workbook(tmp_path)
    cells = ["Out!D6", "Out!E6", "Out!F6"]
    pkg = load_package(
        generate_inverted(workbook, _milestone_bindings()),
        tmp_path,
        name="sparse_milestones",
    )
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = pkg.compute_milestones(rate=0)
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx((2050.0, 2075.0, 2099.0))
