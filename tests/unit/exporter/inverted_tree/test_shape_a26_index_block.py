"""INDEX/OFFSET into a bound 2-D block keep column and range offset (#654).

Q-CRAFT historical rows are `INDEX(window, MATCH(country), 1)` where the
window slides one column per year inside a bound matrix. Lowering that to
`xl_at(flat_block, row - 1)` drops the column and the window offset, so
every year returns the same wrong element. OFFSET used `len(block)` as the
row stride. Both must fail closed when they cannot be expressed.
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from fastpyxl.utils.cell import get_column_letter

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
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

_MEASURE = {
    "concept": "OBS_VALUE",
    "dtype": "float",
    "bind": {"kind": "data_cell", "read": "float"},
}
_TIME = {
    "id": "TIME_PERIOD",
    "concept": "TIME_PERIOD",
    "role": "key",
    "scope": "cell",
    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
}


def _const(
    series_id: str,
    data_range: str,
    *,
    dtype: str = "float",
    layout: str = "series",
) -> dict[str, Any]:
    read = {"string": "string", "int": "int"}.get(dtype, "float")
    return {
        "id": series_id,
        "sheet": data_range.split("!", 1)[0],
        "data_range": data_range,
        "layout": layout,
        "constant": {},
        "structure": {
            "measure": {
                "concept": series_id.upper(),
                "dtype": dtype,
                "bind": {"kind": "data_cell", "read": read},
            },
            "dimensions": [],
        },
        "key": [],
    }


def country_table_sheets(n: int) -> dict[str, dict[str, object]]:
    """Country×year block with a sliding `INDEX(..., MATCH(country), 1)` window."""
    cells: dict[str, object] = {"A1": "Kenya"}
    for index in range(n):
        col = get_column_letter(index + 2)
        cells[f"{col}1"] = 2008 + index
    cells["A5"] = "France"
    cells["A6"] = "Kenya"
    cells["A7"] = "Peru"
    # Window is 3 columns wide, so the bound block needs two extra columns.
    for col_offset in range(n + 2):
        col = get_column_letter(col_offset + 2)
        cells[f"{col}5"] = 10.0 + col_offset
        cells[f"{col}6"] = 20.0 + col_offset
        cells[f"{col}7"] = 30.0 + col_offset
    for index in range(n):
        start = get_column_letter(index + 2)
        end = get_column_letter(index + 4)
        host = get_column_letter(index + 2)
        cells[f"{host}2"] = f"=INDEX(${start}$5:${end}$7,MATCH($A$1,$A$5:$A$7,0),1)"
    return {"Engine": cells}


def country_table_bindings(n: int) -> dict[str, Any]:
    last_year = get_column_letter(n + 1)
    last_block = get_column_letter(n + 3)
    return bindings_document(
        series_entry("country", "Engine!A1", layout="scalar", direction="input", dtype="string"),
        _const("names", "Engine!A5:A7", dtype="string"),
        _const("block", f"Engine!B5:{last_block}7", layout="matrix"),
        {
            "id": "picked",
            "sheet": "Engine",
            "data_range": f"Engine!B2:{last_year}2",
            "layout": "series",
            "output": {"compute": {"name": "compute_picked"}},
            "structure": {"measure": _MEASURE, "dimensions": [_TIME]},
            "key": ["TIME_PERIOD"],
        },
    )


def _country_table_workbook(tmp_path: Path) -> Path:
    return write_oriented_workbook(
        tmp_path / "country_table.xlsx",
        country_table_sheets(3),
        orientation="horizontal",
    )


def _country_table_bindings() -> dict[str, Any]:
    return country_table_bindings(3)


def _case_sheets(
    formula: Callable[[str, int], str],
    *,
    extra: dict[str, object] | None = None,
) -> dict[str, dict[str, object]]:
    cells: dict[str, object] = {
        "A1": "Kenya",
        "B1": 2008,
        "C1": 2009,
        "D1": 2010,
        "A5": "France",
        "B5": 10,
        "C5": 11,
        "D5": 12,
        "E5": 13,
        "F5": 14,
        "A6": "Kenya",
        "B6": 20,
        "C6": 21,
        "D6": 22,
        "E6": 23,
        "F6": 24,
        "A7": "Peru",
        "B7": 30,
        "C7": 31,
        "D7": 32,
        "E7": 33,
        "F7": 34,
    }
    cells.update(extra or {})
    for index, col in enumerate("BCD"):
        cells[f"{col}2"] = formula(col, index)
    return {"Engine": cells}


def _case_bindings(block_range: str, *, extra_series: tuple[dict[str, Any], ...] = ()) -> dict:
    return bindings_document(
        series_entry("country", "Engine!A1", layout="scalar", direction="input", dtype="string"),
        _const("names", "Engine!A5:A7", dtype="string"),
        _const("block", block_range, layout="matrix"),
        *extra_series,
        {
            "id": "picked",
            "sheet": "Engine",
            "data_range": "Engine!B2:D2",
            "layout": "series",
            "output": {"compute": {"name": "compute_picked"}},
            "structure": {"measure": _MEASURE, "dimensions": [_TIME]},
            "key": ["TIME_PERIOD"],
        },
    )


def _slide_sheets() -> dict[str, dict[str, object]]:
    return _case_sheets(
        lambda _col, index: (
            f"=INDEX(${get_column_letter(2 + index)}$5:"
            f"${get_column_letter(4 + index)}$7,MATCH($A$1,$A$5:$A$7,0),1)"
        )
    )


def _literal_sheets() -> dict[str, dict[str, object]]:
    return _case_sheets(
        lambda _col, index: f"=INDEX($B$5:$D$7,MATCH($A$1,$A$5:$A$7,0),{index + 1})"
    )


def _offset_sheets() -> dict[str, dict[str, object]]:
    return _case_sheets(
        lambda _col, index: f"=OFFSET($B$5,$A$9,{index})",
        extra={"A9": 1},
    )


def _slide_bindings() -> dict[str, Any]:
    return _case_bindings("Engine!B5:F7")


def _literal_bindings() -> dict[str, Any]:
    return _case_bindings("Engine!B5:D7")


def _offset_bindings() -> dict[str, Any]:
    return _case_bindings(
        "Engine!B5:D7",
        extra_series=(
            series_entry("rowsel", "Engine!A9", layout="scalar", direction="input", dtype="int"),
        ),
    )


def _export_matches_evaluator(
    tmp_path: Path,
    sheets: dict[str, dict[str, object]],
    document: dict[str, Any],
    orientation: str,
    stem: str,
) -> tuple[tuple[object, ...], tuple[object, ...]]:
    workbook = write_oriented_workbook(
        tmp_path / f"{stem}_{orientation[0]}.xlsx",
        sheets,
        orientation=orientation,
    )
    bound = oriented_document(document, orientation)
    catalog, _deps, graph = inverted_graph_parts(workbook, bound)
    pkg = load_package(
        generate_inverted(workbook, bound), tmp_path, name=f"{stem}_{orientation[0]}"
    )
    cells = oriented_addresses(("Engine!B2", "Engine!C2", "Engine!D2"), orientation)
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, list(cells), load_values=True)
    ).evaluate(list(cells))
    got = pkg.compute_picked(**input_kwargs(catalog, graph))
    want = tuple(expected[cell] for cell in cells)
    assert got == pytest.approx(want)
    return got, want


def test_sliding_index_window_is_one_affine_statement(tmp_path: Path, orientation: str) -> None:
    workbook = write_oriented_workbook(
        tmp_path / f"a26_slide_stmt_{orientation[0]}.xlsx",
        _slide_sheets(),
        orientation=orientation,
    )
    catalog, _deps, _graph = inverted_graph_parts(
        workbook, oriented_document(_slide_bindings(), orientation)
    )
    picked = catalog.get("picked")
    assert len(picked.statements) == 1
    assert picked.statements[0].start == 0
    assert picked.statements[0].stop == 3


def test_sliding_index_matches_evaluator(tmp_path: Path, orientation: str) -> None:
    got, _want = _export_matches_evaluator(
        tmp_path, _slide_sheets(), _slide_bindings(), orientation, "a26_slide"
    )
    if orientation == "horizontal":
        assert got == pytest.approx((20.0, 21.0, 22.0))


def test_literal_index_column_matches_evaluator(tmp_path: Path, orientation: str) -> None:
    got, _want = _export_matches_evaluator(
        tmp_path, _literal_sheets(), _literal_bindings(), orientation, "a26_lit"
    )
    if orientation == "horizontal":
        assert got == pytest.approx((20.0, 21.0, 22.0))


def test_offset_into_matrix_matches_evaluator(tmp_path: Path, orientation: str) -> None:
    got, _want = _export_matches_evaluator(
        tmp_path, _offset_sheets(), _offset_bindings(), orientation, "a26_off"
    )
    if orientation == "horizontal":
        assert got == pytest.approx((20.0, 21.0, 22.0))


def test_unbound_index_column_fails_closed_naming_host(tmp_path: Path) -> None:
    sheets = _case_sheets(
        lambda _col, _index: "=INDEX($B$5:$D$7,MATCH($A$1,$A$5:$A$7,0),$Z$9)",
        extra={"Z9": 2},
    )
    workbook = write_oriented_workbook(
        tmp_path / "a26_unbound_col.xlsx", sheets, orientation="horizontal"
    )
    with pytest.raises(InvertedTreeExportError, match="Engine!B2") as exc:
        generate_inverted(workbook, _literal_bindings())
    assert "INDEX column" in str(exc.value)


def test_offset_nonzero_row_into_series_fails_closed(tmp_path: Path) -> None:
    workbook = write_oriented_workbook(
        tmp_path / "a26_offset_1d.xlsx",
        {
            "Inputs": {"A1": 1, "B1": -2.0, "C1": 2.0, "D1": -1.0},
            "Engine": {"A1": "=OFFSET(Inputs!B1,1,0)"},
            "Outputs": {"A1": "=Engine!A1"},
        },
        orientation="horizontal",
    )
    document = bindings_document(
        series_entry("shock_type", "Inputs!A1", layout="scalar", direction="input", dtype="int"),
        series_entry(
            "shock_magnitudes",
            "Inputs!B1:D1",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry(
            "shock_magnitude_resolved",
            "Engine!A1",
            layout="scalar",
            direction="internal",
        ),
        series_entry("output_mag", "Outputs!A1", layout="scalar", direction="output"),
    )
    with pytest.raises(InvertedTreeExportError, match="Engine!A1") as exc:
        generate_inverted(workbook, document)
    assert "non-matrix" in str(exc.value)
