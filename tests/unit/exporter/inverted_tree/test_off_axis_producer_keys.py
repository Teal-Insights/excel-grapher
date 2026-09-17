"""Sampled integer affines must not index keys the producer axis does not own.

A host calendar (or other integer) axis whose formulas copy a shorter producer
life axis emits `producer[host - origin]`. First/middle/last sampling plus the
largest-family unconditional return used to replicate that lookup onto years
whose affine image is missing from the producer axis, raising `CoordinateError`.
On-axis sparse holes still fold as named indexes; off-axis images return `None`.
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
    load_package,
    write_workbook,
)

_N_HOST = 13
_MATRIX_LIFE = 11
_VALUED_LIFE = 4
_ORIGIN = 2024
_BUCKET_ORIGIN = 100
_ALPHA, _BETA = "Alpha", "Beta"


def _col(life: int) -> int:
    return 2 + life


def _letter(life: int) -> str:
    return get_column_letter(_col(life))


def _matrix_copy_workbook(tmp_path: Path) -> Path:
    """Host years copy Input life columns; the matrix stops before the host tail."""
    inp: dict[str, object] = {_letter(life) + "1": life for life in range(_N_HOST)}
    inp["A2"] = _ALPHA
    inp["A3"] = _BETA
    for life in range(_MATRIX_LIFE):
        inp[f"{_letter(life)}2"] = 0.1 * (life + 1)
        inp[f"{_letter(life)}3"] = 0.2 * (life + 1)
    host: dict[str, object] = {_letter(life) + "1": _ORIGIN + life for life in range(_N_HOST)}
    host["A2"] = _ALPHA
    host["A3"] = _BETA
    for life in range(_N_HOST):
        src = _letter(life)
        host[f"{src}2"] = f"=Input!{src}2"
        host[f"{src}3"] = f"=Input!{src}3"
    return write_workbook(tmp_path / "off_axis_matrix.xlsx", {"Input": inp, "Host": host})


def _matrix_copy_bindings() -> dict[str, Any]:
    last_matrix = _letter(_MATRIX_LIFE - 1)
    last_host = _letter(_N_HOST - 1)
    instruments = {_ALPHA: 2, _BETA: 3}
    document = bindings_document(
        {
            "id": "principal",
            "sheet": "Input",
            "data_range": f"Input!B2:{last_matrix}3",
            "layout": "matrix",
            "input": {},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [
                    {
                        "id": "INSTRUMENT",
                        "concept": "INSTRUMENT",
                        "role": "key",
                        "scope": "cell",
                        "bind": {"kind": "value_map", "values": instruments},
                    },
                    {
                        "id": "TIME_PERIOD",
                        "concept": "TIME_PERIOD",
                        "role": "key",
                        "scope": "cell",
                        "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                    },
                ],
            },
            "key": ["INSTRUMENT", "TIME_PERIOD"],
        },
        {
            "id": "schedule",
            "sheet": "Host",
            "data_range": f"Host!B2:{last_host}3",
            "layout": "series",
            "output": {"compute": {"name": "compute_schedule"}},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [
                    {
                        "id": "INSTRUMENT",
                        "concept": "INSTRUMENT",
                        "role": "key",
                        "scope": "cell",
                        "bind": {
                            "kind": "value_map",
                            "values": instruments,
                            "read": "string",
                        },
                    },
                    {
                        "id": "TIME_PERIOD",
                        "concept": "TIME_PERIOD",
                        "role": "key",
                        "scope": "cell",
                        "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                    },
                ],
            },
            "key": ["INSTRUMENT", "TIME_PERIOD"],
        },
        schema_version="1.16.0",
    )
    document["concept_scheme"]["concepts"].append({"id": "INSTRUMENT", "dtype": "string"})
    return document


def _mcve_blank_ranges() -> tuple[str, ...]:
    return (
        f"Input!{_letter(_VALUED_LIFE)}2:{_letter(_N_HOST - 1)}2",
        f"Input!{_letter(0)}3:{_letter(1)}3",
        f"Input!{_letter(_VALUED_LIFE)}3:{_letter(_N_HOST - 1)}3",
    )


def _series_copy_workbook(
    tmp_path: Path, *, n_host: int = _N_HOST, matrix_life: int = _MATRIX_LIFE
) -> Path:
    """1-D host buckets copy a shorter STEP producer; field names differ."""
    cells: dict[str, object] = {}
    for life in range(n_host):
        letter = _letter(life)
        cells[f"{letter}1"] = life
        cells[f"{letter}3"] = _BUCKET_ORIGIN + life
        cells[f"{letter}4"] = f"=Sheet!{letter}2"
        if life < matrix_life:
            cells[f"{letter}2"] = float(life + 1)
    return write_workbook(tmp_path / "off_axis_series.xlsx", {"Sheet": cells})


def _series_copy_bindings(
    *, n_host: int = _N_HOST, matrix_life: int = _MATRIX_LIFE
) -> dict[str, Any]:
    last_matrix = _letter(matrix_life - 1)
    last_host = _letter(n_host - 1)
    document = bindings_document(
        {
            "id": "steps",
            "sheet": "Sheet",
            "data_range": f"Sheet!B2:{last_matrix}2",
            "layout": "series",
            "input": {},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [
                    {
                        "id": "STEP",
                        "concept": "STEP",
                        "role": "key",
                        "scope": "cell",
                        "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                    }
                ],
            },
            "key": ["STEP"],
        },
        {
            "id": "copied",
            "sheet": "Sheet",
            "data_range": f"Sheet!B4:{last_host}4",
            "layout": "series",
            "output": {"compute": {"name": "compute_copied"}},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [
                    {
                        "id": "TIME_PERIOD",
                        "concept": "TIME_PERIOD",
                        "role": "key",
                        "scope": "cell",
                        "bind": {"kind": "column_header", "header_row": 3, "read": "int"},
                    }
                ],
            },
            "key": ["TIME_PERIOD"],
        },
        schema_version="1.16.0",
    )
    document["concept_scheme"]["concepts"].append({"id": "STEP", "dtype": "int"})
    return document


def _series_blank_ranges(
    *, n_host: int = _N_HOST, valued_life: int = _VALUED_LIFE
) -> tuple[str, ...]:
    return (f"Sheet!{_letter(valued_life)}2:{_letter(n_host - 1)}2",)


def test_matrix_copy_does_not_index_off_axis_life_keys(tmp_path: Path) -> None:
    """LIC-DSF MCVE: Alpha year whose life image is 4 must not look up principal."""
    workbook = _matrix_copy_workbook(tmp_path)
    document = _matrix_copy_bindings()
    blanks = _mcve_blank_ranges()
    modules = generate_inverted(workbook, document, blank_ranges=blanks)
    pkg = load_package(modules, tmp_path, name="off_axis_matrix")
    owned_life = {coord[1] for coord in pkg.data.PRINCIPAL.domain}
    assert 4 not in owned_life
    got = pkg.compute_schedule(principal=pkg.data.PRINCIPAL_DEFAULT)
    for life in range(_N_HOST):
        year = _ORIGIN + life
        if life < _VALUED_LIFE:
            assert got[_ALPHA, year] == pytest.approx(0.1 * (life + 1))
        else:
            assert got[_ALPHA, year] is None
        if 2 <= life < _VALUED_LIFE:
            assert got[_BETA, year] == pytest.approx(0.2 * (life + 1))
        else:
            assert got[_BETA, year] is None


def test_cross_field_integer_affine_does_not_index_missing_step(tmp_path: Path) -> None:
    """Host TIME_PERIOD origin 100 copying producer STEP 0..n must blank off-axis years."""
    workbook = _series_copy_workbook(tmp_path)
    document = _series_copy_bindings()
    blanks = _series_blank_ranges()
    modules = generate_inverted(workbook, document, blank_ranges=blanks)
    pkg = load_package(modules, tmp_path, name="off_axis_series")
    producer_steps = pkg.data.STEPS.domain.axes[0].keys
    assert 4 not in producer_steps
    got = pkg.compute_copied(steps=pkg.data.STEPS_DEFAULT)
    for life in range(_N_HOST):
        bucket = _BUCKET_ORIGIN + life
        if life < _VALUED_LIFE:
            assert got[bucket] == pytest.approx(float(life + 1))
        else:
            assert got[bucket] is None
    targets = [f"Sheet!{_letter(life)}4" for life in range(_VALUED_LIFE)]
    graph = create_dependency_graph(workbook, targets, load_values=True, blank_ranges=blanks)
    expected = FormulaEvaluator(graph, blank_ranges=blanks).evaluate(targets)
    for life, cell in enumerate(targets):
        assert got[_BUCKET_ORIGIN + life] == pytest.approx(expected[cell])


def test_largest_lookup_family_does_not_cover_off_axis_tail(tmp_path: Path) -> None:
    """A lookup that is the unconditional default must still blank off-axis years."""
    n_host, matrix_life, valued_life = 13, 12, 10
    workbook = _series_copy_workbook(tmp_path, n_host=n_host, matrix_life=matrix_life)
    document = _series_copy_bindings(n_host=n_host, matrix_life=matrix_life)
    blanks = _series_blank_ranges(n_host=n_host, valued_life=valued_life)
    modules = generate_inverted(workbook, document, blank_ranges=blanks)
    internals = modules["internals.py"]
    assert "return as_measure(steps[time_period - 100])" in internals
    assert "as_measure(None)" in internals
    pkg = load_package(modules, tmp_path, name="off_axis_lookup_default")
    got = pkg.compute_copied(steps=pkg.data.STEPS_DEFAULT)
    for life in range(n_host):
        bucket = _BUCKET_ORIGIN + life
        if life < valued_life:
            assert got[bucket] == pytest.approx(float(life + 1))
        else:
            assert got[bucket] is None
