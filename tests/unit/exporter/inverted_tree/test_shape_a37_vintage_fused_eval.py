"""Layer A37 — fused named-axis zipper evaluation at vintage-triangle scale (#828).

#762 made the distance-zero residual a DAG per `ISSUANCE_YEAR`. The two-vintage
MCVE then special-cases four cells (literals plus one amortizing walk). LIC-DSF
Input 5 is the same zipper with many `INSTRUMENT × ISSUANCE_YEAR` partitions:
issuance-year stock is this year's issuance, opening principal/interest must
not compile missing coordinates as `None`, and `compute_*` must match
`FormulaEvaluator`.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from fastpyxl.utils.cell import get_column_letter

from excel_grapher.evaluator import FormulaEvaluator
from tests.unit.exporter.inverted_tree.helpers import (
    generate_inverted,
    inverted_graph_parts,
    load_package,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a36_vintage_residual import (
    vintage_residual_bindings,
)

_YEARS = (2024, 2025, 2026, 2027, 2028)
_ISSUANCE = (100.0, 80.0, 60.0, 40.0, 20.0)


def _measure() -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }


def _time(*, header_row: int = 1) -> dict[str, Any]:
    return {
        "id": "TIME_PERIOD",
        "concept": "TIME_PERIOD",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "column_header", "header_row": header_row, "read": "int"},
    }


def _issuance_dim(*, label_column: str = "A") -> dict[str, Any]:
    return {
        "id": "ISSUANCE_YEAR",
        "concept": "ISSUANCE_YEAR",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "row_label", "label_column": label_column, "read": "int"},
    }


def _instrument_dim(*, label_column: str = "B") -> dict[str, Any]:
    return {
        "id": "INSTRUMENT",
        "concept": "INSTRUMENT",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "row_label", "label_column": label_column, "read": "string"},
    }


def issuance_formula_workbook(tmp_path: Path) -> Path:
    """#762 topology with issuance-year stock as a formula, not a literal."""
    return write_workbook(
        tmp_path / "a37_issuance_formula.xlsx",
        {
            "V": {
                "D1": 2024,
                "E1": 2025,
                "D3": 100,
                "E3": 50,
                "A5": 2024,
                "A6": 2025,
                "A7": 2024,
                "A8": 2025,
                "D5": "=D3",
                "E5": "=D5-E7",
                "D6": 0,
                "E6": "=E3",
                "D7": 0,
                "E7": "=IF(E$1>2024,$D5/2,0)",
                "D8": 0,
                "E8": "=E6",
                "A10": "=E5",
            }
        },
    )


def issuance_formula_bindings() -> dict[str, Any]:
    document = vintage_residual_bindings()
    document["series"].insert(
        1,
        {
            "id": "issuance",
            "sheet": "V",
            "data_range": "V!D3:E3",
            "layout": "series",
            "constant": {},
            "structure": {
                "measure": _measure(),
                "dimensions": [_time()],
            },
            "key": ["TIME_PERIOD"],
        },
    )
    return document


def _triangle_cells(*, instruments: tuple[str, ...] = ("Bond",)) -> dict[str, object]:
    cells: dict[str, object] = {}
    for index, year in enumerate(_YEARS):
        col = get_column_letter(index + 3)
        cells[f"{col}1"] = year
        cells[f"{col}2"] = _ISSUANCE[index]
    n = len(instruments) * len(_YEARS)
    stock_start = 5
    principal_start = stock_start + n + 2
    interest_start = principal_start + n + 2
    for inst_i, instrument in enumerate(instruments):
        for vintage_i, issued in enumerate(_YEARS):
            stock_row = stock_start + inst_i * len(_YEARS) + vintage_i
            prin_row = principal_start + inst_i * len(_YEARS) + vintage_i
            int_row = interest_start + inst_i * len(_YEARS) + vintage_i
            cells[f"A{stock_row}"] = issued
            cells[f"B{stock_row}"] = instrument
            cells[f"A{prin_row}"] = issued
            cells[f"B{prin_row}"] = instrument
            cells[f"A{int_row}"] = issued
            cells[f"B{int_row}"] = instrument
            opening_col = get_column_letter(vintage_i + 3)
            for time_i, year in enumerate(_YEARS):
                col = get_column_letter(time_i + 3)
                if year < issued:
                    continue
                if time_i == 0:
                    cells[f"{col}{int_row}"] = 0
                else:
                    prev = get_column_letter(time_i + 2)
                    # Previous-column stock at issuance year is a triangle hole.
                    cells[f"{col}{int_row}"] = f"={prev}{stock_row}*0.05"
                if year == issued:
                    cells[f"{col}{stock_row}"] = f"={col}2"
                    if instrument == "T-bills":
                        cells[f"{col}{prin_row}"] = f"={col}{stock_row}"
                    else:
                        cells[f"{col}{prin_row}"] = (
                            f"=IF({col}$1>{issued},${opening_col}{stock_row}/2,0)"
                        )
                    continue
                prev = get_column_letter(time_i + 2)
                cells[f"{col}{stock_row}"] = f"={prev}{stock_row}-{col}{prin_row}"
                if instrument == "T-bills":
                    cells[f"{col}{prin_row}"] = f"={prev}{stock_row}"
                else:
                    cells[f"{col}{prin_row}"] = (
                        f"=IF({col}$1>{issued},${opening_col}{stock_row}/2,0)"
                    )
    last_col = get_column_letter(len(_YEARS) + 2)
    last_stock = stock_start + n - 1
    cells[f"A{interest_start + n + 2}"] = f"=SUM(C{stock_start}:{last_col}{last_stock})"
    return cells


def triangle_workbook(tmp_path: Path, *, instruments: tuple[str, ...] = ("Bond",)) -> Path:
    return write_workbook(
        tmp_path / "a37_triangle.xlsx",
        {"V": _triangle_cells(instruments=instruments)},
    )


def _triangle_meta(*, instruments: tuple[str, ...] = ("Bond",)) -> dict[str, Any]:
    n = len(instruments) * len(_YEARS)
    stock_start = 5
    principal_start = stock_start + n + 2
    interest_start = principal_start + n + 2
    last_col = get_column_letter(len(_YEARS) + 2)
    return {
        "stock": f"V!C{stock_start}:{last_col}{stock_start + n - 1}",
        "principal": f"V!C{principal_start}:{last_col}{principal_start + n - 1}",
        "interest": f"V!C{interest_start}:{last_col}{interest_start + n - 1}",
        "result": f"V!A{interest_start + n + 2}",
        "instruments": instruments,
    }


def triangle_bindings(*, instruments: tuple[str, ...] = ("Bond",)) -> dict[str, Any]:
    meta = _triangle_meta(instruments=instruments)
    extra = [_instrument_dim()] if len(instruments) > 1 else ()
    key = (["INSTRUMENT"] if extra else []) + ["ISSUANCE_YEAR", "TIME_PERIOD"]
    dims = [*(extra), _issuance_dim(), _time()]
    concepts = [
        {"id": "TIME_PERIOD", "dtype": "int"},
        {"id": "OBS_VALUE", "dtype": "number"},
        {"id": "ISSUANCE_YEAR", "dtype": "int"},
    ]
    if extra:
        concepts.append({"id": "INSTRUMENT", "dtype": "string"})

    def keyed(series_id: str, data_range: str, *, output: bool = False) -> dict[str, Any]:
        entry: dict[str, Any] = {
            "id": series_id,
            "sheet": "V",
            "data_range": data_range,
            "layout": "matrix",
            "structure": {"measure": _measure(), "dimensions": dims},
            "key": key,
        }
        if output:
            entry["output"] = {"compute": {"name": f"compute_{series_id}"}}
        else:
            entry["internal"] = {}
        return entry

    last_col = get_column_letter(len(_YEARS) + 2)
    return {
        "schema_version": "1.15.0",
        "concept_scheme": {"id": "a37_triangle", "concepts": concepts},
        "series": [
            {
                "id": "years",
                "sheet": "V",
                "data_range": f"V!C1:{last_col}1",
                "layout": "series",
                "constant": {},
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": "int",
                        "bind": {"kind": "data_cell", "read": "int"},
                    },
                    "dimensions": [_time()],
                },
                "key": ["TIME_PERIOD"],
            },
            {
                "id": "issuance",
                "sheet": "V",
                "data_range": f"V!C2:{last_col}2",
                "layout": "series",
                "constant": {},
                "structure": {"measure": _measure(), "dimensions": [_time()]},
                "key": ["TIME_PERIOD"],
            },
            keyed("stock", meta["stock"], output=True),
            keyed("principal", meta["principal"]),
            keyed("interest", meta["interest"], output=True),
            {
                "id": "result",
                "sheet": "V",
                "data_range": meta["result"],
                "layout": "scalar",
                "output": {"compute": {"name": "compute_result"}},
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": "float",
                        "bind": {"kind": "data_cell", "read": "float"},
                    },
                    "dimensions": [],
                },
                "key": [],
            },
        ],
    }


def _triangle_blank_ranges(*, instruments: tuple[str, ...] = ("Bond",)) -> tuple[str, ...]:
    blanks: list[str] = []
    n = len(instruments) * len(_YEARS)
    stock_start = 5
    principal_start = stock_start + n + 2
    interest_start = principal_start + n + 2
    for inst_i, _instrument in enumerate(instruments):
        for vintage_i, issued in enumerate(_YEARS):
            stock_row = stock_start + inst_i * len(_YEARS) + vintage_i
            prin_row = principal_start + inst_i * len(_YEARS) + vintage_i
            int_row = interest_start + inst_i * len(_YEARS) + vintage_i
            for time_i, year in enumerate(_YEARS):
                if year >= issued:
                    continue
                col = get_column_letter(time_i + 3)
                blanks.append(f"V!{col}{stock_row}")
                blanks.append(f"V!{col}{prin_row}")
                blanks.append(f"V!{col}{int_row}")
    return tuple(blanks)


def test_issuance_year_stock_formula_matches_evaluator(tmp_path: Path) -> None:
    workbook = issuance_formula_workbook(tmp_path)
    document = issuance_formula_bindings()
    modules = generate_inverted(workbook, document)
    internals = modules["internals.py"]
    assert "eval_instance" not in internals
    assert "issuance[time_period]" in internals or "issuance[issuance_year]" in internals
    pkg = load_package(modules, tmp_path, name="a37_issuance")
    expected = FormulaEvaluator(inverted_graph_parts(workbook, document)[2]).evaluate(
        ["V!D5", "V!E5", "V!D6", "V!E6", "V!A10"]
    )
    stock = dict(pkg.compute_stock().items())
    assert stock[2024, 2024] == pytest.approx(expected["V!D5"])
    assert stock[2024, 2025] == pytest.approx(expected["V!E5"])
    assert stock[2025, 2025] == pytest.approx(expected["V!E6"])
    assert pkg.compute_result() == pytest.approx(expected["V!A10"])


def test_opening_stock_is_named_as_issuance_year(tmp_path: Path) -> None:
    workbook = triangle_workbook(tmp_path)
    document = triangle_bindings()
    internals = generate_inverted(workbook, document)["internals.py"]
    assert "eval_instance" not in internals
    assert "stock[issuance_year, issuance_year]" in internals
    assert "stock[issuance_year, 2024]" not in internals
    assert "xl_div(None" not in internals
    assert "xl_mul(None" not in internals
    assert "xl_mul(stock[" in internals


def test_vintage_triangle_export_matches_evaluator(tmp_path: Path) -> None:
    instruments = ("Bond", "T-bills")
    workbook = triangle_workbook(tmp_path, instruments=instruments)
    document = triangle_bindings(instruments=instruments)
    modules = generate_inverted(workbook, document)
    internals = modules["internals.py"]
    assert "eval_instance" not in internals
    assert "issuance[time_period]" in internals
    pkg = load_package(modules, tmp_path, name="a37_triangle")
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    stock = catalog.get("stock")
    interest = catalog.get("interest")
    expected = FormulaEvaluator(graph).evaluate(
        [*stock.cells, *interest.cells, document["series"][-1]["data_range"]]
    )
    got = pkg.compute_stock()
    for coord, cell in stock.coordinate_cells.items():
        if cell not in expected:
            continue
        assert got[coord] == pytest.approx(expected[cell]), (coord, cell)
    got_interest = pkg.compute_interest()
    for coord, cell in interest.coordinate_cells.items():
        if cell not in expected:
            continue
        assert got_interest[coord] == pytest.approx(expected[cell]), (coord, cell)
    result = pkg.compute_result()
    if isinstance(result, tuple):
        result = result[0]
    assert result == pytest.approx(expected[document["series"][-1]["data_range"]])


def test_triangle_blanks_are_named_indexes_not_none(tmp_path: Path) -> None:
    instruments = ("Bond", "T-bills")
    workbook = triangle_workbook(tmp_path, instruments=instruments)
    document = triangle_bindings(instruments=instruments)
    blanks = _triangle_blank_ranges(instruments=instruments)
    modules = generate_inverted(workbook, document, blank_ranges=blanks)
    internals = modules["internals.py"]
    assert "xl_div(None" not in internals
    assert "xl_mul(None" not in internals
    assert "xl_mul(stock[" in internals
    pkg = load_package(modules, tmp_path, name="a37_blanks")
    catalog, _deps, graph = inverted_graph_parts(workbook, document, blank_ranges=blanks)
    stock = catalog.get("stock")
    interest = catalog.get("interest")
    expected = FormulaEvaluator(graph, blank_ranges=blanks).evaluate(
        [cell for cell in (*stock.cells, *interest.cells) if graph.get_node(cell) is not None]
    )
    got = pkg.compute_stock()
    for coord, cell in stock.coordinate_cells.items():
        if cell not in expected:
            continue
        assert got[coord] == pytest.approx(expected[cell]), (coord, cell)
    got_interest = pkg.compute_interest()
    for coord, cell in interest.coordinate_cells.items():
        if cell not in expected:
            continue
        assert got_interest[coord] == pytest.approx(expected[cell]), (coord, cell)
