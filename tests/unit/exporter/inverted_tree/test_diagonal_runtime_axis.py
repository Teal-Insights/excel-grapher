"""Diagonal bands on a runtime-labelled axis compare keys in one space.

A header-inferred labeller reads cached formula results, so these workbooks
plant `<v>` values on the header formulas. The diagonal test is
`issuance_year == _ax_time_period.keys[_p_time_period]`.
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Any

import pytest
from fastpyxl.utils.cell import column_index_from_string, get_column_letter

from tests.unit.exporter.inverted_tree.helpers import (
    assert_package_matches_evaluator,
    bindings_document,
    generate_inverted,
    write_workbook,
)

_MEASURE = {
    "concept": "OBS_VALUE",
    "dtype": "float",
    "bind": {"kind": "data_cell", "read": "float"},
}


def _plant_cached_numbers(path: Path, cached: dict[str, int]) -> None:
    """Keep each formula and set its saved numeric value."""
    sheet = "xl/worksheets/sheet1.xml"
    with zipfile.ZipFile(path) as archive:
        xml = archive.read(sheet).decode()
        others = {name: archive.read(name) for name in archive.namelist() if name != sheet}
    for coordinate, value in cached.items():
        marker = f'<c r="{coordinate}"'
        start = xml.find(marker)
        end = xml.find("</c>", start)
        if start < 0 or end < 0:
            raise AssertionError(f"{coordinate} formula cell not found")
        cell = xml[start:end]
        if "<f>" not in cell and "<f " not in cell:
            raise AssertionError(f"{coordinate} is not a formula cell: {cell}")
        if "<v />" in cell:
            cell = cell.replace("<v />", f"<v>{value}</v>", 1)
        elif "<v/>" in cell:
            cell = cell.replace("<v/>", f"<v>{value}</v>", 1)
        elif "<v>" in cell:
            close = cell.find("</v>")
            open_end = cell.find(">", cell.find("<v"))
            cell = cell[: open_end + 1] + str(value) + cell[close:]
        else:
            cell += f"<v>{value}</v>"
        xml = xml[:start] + cell + xml[end:]
    with zipfile.ZipFile(path, "w") as archive:
        for name, payload in others.items():
            archive.writestr(name, payload)
        archive.writestr(sheet, xml)


def _header_formulas(years: tuple[int, ...], *, row: int = 1) -> dict[str, str]:
    cells: dict[str, str] = {}
    for index, year in enumerate(years):
        column = get_column_letter(index + 2)
        if index == 0:
            cells[f"{column}{row}"] = f"={year}"
        else:
            previous = get_column_letter(index + 1)
            cells[f"{column}{row}"] = f"={previous}{row}+1"
    return cells


def _header_caches(years: tuple[int, ...], *, row: int = 1) -> dict[str, int]:
    return {f"{get_column_letter(index + 2)}{row}": year for index, year in enumerate(years)}


def _row_formulas(
    years: tuple[int, ...], *, column: str = "A", start_row: int = 2
) -> dict[str, str]:
    cells: dict[str, str] = {}
    for index, year in enumerate(years):
        row = start_row + index
        if index == 0:
            cells[f"{column}{row}"] = f"={year}"
        else:
            cells[f"{column}{row}"] = f"={column}{row - 1}+1"
    return cells


def _row_label(concept: str, *, column: str = "A") -> dict[str, Any]:
    return {
        "id": concept,
        "concept": concept,
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "row_label", "label_column": column, "read": "int"},
    }


def _column_header(concept: str, *, row: int = 1) -> dict[str, Any]:
    return {
        "id": concept,
        "concept": concept,
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "column_header", "header_row": row, "read": "int"},
    }


def _series(
    series_id: str,
    data_range: str,
    dimensions: list[dict[str, Any]],
    key: list[str],
    *,
    output: bool,
    dtype: str = "float",
) -> dict[str, Any]:
    measure = dict(_MEASURE)
    if dtype != "float":
        measure = {
            "concept": "OBS_VALUE",
            "dtype": dtype,
            "bind": {"kind": "data_cell", "read": dtype},
        }
    entry: dict[str, Any] = {
        "id": series_id,
        "sheet": data_range.split("!", 1)[0],
        "data_range": data_range,
        "layout": "scalar" if not key else "series",
        "structure": {"measure": measure, "dimensions": dimensions},
        "key": key,
    }
    if output:
        entry["output"] = {"compute": {"name": f"compute_{series_id}"}}
    else:
        entry["input"] = {}
    return entry


def _document(*series: dict[str, Any]) -> dict[str, Any]:
    document = bindings_document(*series, schema_version="1.22.0")
    document["concept_scheme"]["concepts"].append({"id": "ISSUANCE_YEAR", "dtype": "int"})
    return document


def _body(source: str, series_id: str) -> str:
    marker = f"def {series_id}("
    start = source.find(marker)
    if start < 0:
        raise AssertionError(f"{marker} missing from internals")
    rest = source[start:]
    nxt = rest.find("\ndef ", 1)
    return rest if nxt < 0 else rest[:nxt]


def _vintage_grid(
    years: tuple[int, ...],
    issued: tuple[int, ...],
    *,
    formula_for: Any,
    label_row: int = 2,
    header_formulas: bool = True,
    row_formulas: bool = False,
) -> dict[str, object]:
    """Build one vintage block: headers, row labels, and per-cell formulas."""
    cells: dict[str, object] = {}
    if header_formulas:
        cells.update(_header_formulas(years))
    else:
        for index, year in enumerate(years):
            cells[f"{get_column_letter(index + 2)}1"] = year
    if row_formulas:
        cells.update(_row_formulas(issued, start_row=label_row))
    else:
        for index, year in enumerate(issued):
            cells[f"A{label_row + index}"] = year
    for row_offset, issuance in enumerate(issued):
        row = label_row + row_offset
        for index, year in enumerate(years):
            column = get_column_letter(index + 2)
            formula = formula_for(year, issuance, column, row)
            if formula is not None:
                cells[f"{column}{row}"] = formula
    return cells


def test_labelled_column_diagonal_compares_year_keys(tmp_path: Path) -> None:
    """A year on the diagonal matches the header key at that position."""
    years = (2024, 2025, 2026, 2027, 2028, 2029)
    issued = (2024, 2025, 2026)

    def formula(year: int, issuance: int, column: str, row: int) -> str | None:
        if year == issuance:
            return "=$Z$1"
        if year > issuance:
            previous = get_column_letter(column_index_from_string(column) - 1)
            return f"={previous}{row}"
        if column == "B" and row == 4:
            return "=99"
        return None

    cells = _vintage_grid(years, issued, formula_for=formula)
    cells["Z1"] = 7
    for index, amount in enumerate((10, 20, 30, 40, 50, 60)):
        cells[f"{get_column_letter(index + 2)}6"] = amount
    path = write_workbook(tmp_path / "years.xlsx", {"Sheet": cells})
    _plant_cached_numbers(path, _header_caches(years))
    period = _column_header("TIME_PERIOD")
    issuance = _row_label("ISSUANCE_YEAR")
    document = _document(
        _series("seed", "Sheet!Z1", [], [], output=False),
        _series("opening", "Sheet!B6:G6", [period], ["TIME_PERIOD"], output=False),
        _series(
            "stock",
            "Sheet!B2:G4",
            [issuance, period],
            ["ISSUANCE_YEAR", "TIME_PERIOD"],
            output=True,
        ),
    )
    blanks = ["Sheet!B3", "Sheet!C4"]
    modules = generate_inverted(path, document, blank_ranges=blanks)
    body = _body(modules["internals.py"], "stock")
    assert "issuance_year == _ax_time_period.keys[_p_time_period]" in body
    assert "issuance_year == _p_time_period" not in body
    assert "_p_time_period - issuance_year" not in body
    assert "issuance_year == 2026 and _p_time_period == 0" in body
    pkg = assert_package_matches_evaluator(
        path, document, tmp_path, "diagonal_year_keys", blank_ranges=blanks
    )
    result = pkg.compute_stock(pkg.StockInputs.from_defaults())
    assert result[2024, 2024] == pytest.approx(7)
    assert result[2024, 2025] == pytest.approx(7)


def test_misaligned_labelled_column_uses_key_difference(tmp_path: Path) -> None:
    """A calendar diagonal is key difference 0 even when positions differ by 4."""
    years = tuple(range(2020, 2028))
    issued = (2024, 2025, 2026)

    def formula(year: int, issuance: int, column: str, row: int) -> str:
        del column, row
        if year == issuance:
            return "=$Z$1"
        if year == issuance + 1:
            return "=$Z$1*2"
        return "=0"

    cells = _vintage_grid(years, issued, formula_for=formula)
    cells["Z1"] = 7
    for index, _year in enumerate(years):
        cells[f"{get_column_letter(index + 2)}12"] = float(index)
    path = write_workbook(tmp_path / "misaligned.xlsx", {"Sheet": cells})
    _plant_cached_numbers(path, _header_caches(years))
    period = _column_header("TIME_PERIOD")
    issuance = _row_label("ISSUANCE_YEAR")
    last = get_column_letter(len(years) + 1)
    document = _document(
        _series("seed", "Sheet!Z1", [], [], output=False),
        _series("opening", f"Sheet!B12:{last}12", [period], ["TIME_PERIOD"], output=False),
        _series(
            "stock",
            f"Sheet!B2:{last}4",
            [issuance, period],
            ["ISSUANCE_YEAR", "TIME_PERIOD"],
            output=True,
        ),
    )
    modules = generate_inverted(path, document)
    body = _body(modules["internals.py"], "stock")
    assert "issuance_year == _ax_time_period.keys[_p_time_period]" in body
    assert "_ax_time_period.keys[_p_time_period] - issuance_year == 1" in body
    assert "_p_time_period - issuance_year" not in body
    assert "== 4" not in body
    assert "== 5" not in body
    pkg = assert_package_matches_evaluator(path, document, tmp_path, "diagonal_misaligned")
    result = pkg.compute_stock(pkg.StockInputs.from_defaults())
    assert result[2024, 2024] == pytest.approx(7)
    assert result[2024, 2025] == pytest.approx(14)


def test_labelled_column_open_band_compares_keys(tmp_path: Path) -> None:
    """An open band after the diagonal subtracts year keys, not positions."""
    years = (2024, 2025, 2026, 2027, 2028)
    issued = (2026, 2027, 2028)

    def formula(year: int, issuance: int, column: str, row: int) -> str:
        source = f"{column}{row - 6}"
        if year == issuance:
            return f"={source}"
        if year > issuance:
            return f"={source}*2"
        return f"={source}*0"

    cells = _vintage_grid(years, issued, formula_for=formula, label_row=8)
    for row_offset, issuance in enumerate(issued):
        cells[f"A{2 + row_offset}"] = issuance
        for index, _year in enumerate(years):
            cells[f"{get_column_letter(index + 2)}{2 + row_offset}"] = float(index + 1)
    path = write_workbook(tmp_path / "open_band.xlsx", {"Sheet": cells})
    _plant_cached_numbers(path, _header_caches(years))
    period = _column_header("TIME_PERIOD")
    issuance = _row_label("ISSUANCE_YEAR")
    document = _document(
        _series(
            "vintage",
            "Sheet!B2:F4",
            [issuance, period],
            ["ISSUANCE_YEAR", "TIME_PERIOD"],
            output=False,
        ),
        _series(
            "aged",
            "Sheet!B8:F10",
            [issuance, period],
            ["ISSUANCE_YEAR", "TIME_PERIOD"],
            output=True,
        ),
    )
    modules = generate_inverted(path, document)
    body = _body(modules["internals.py"], "aged")
    assert "issuance_year == _ax_time_period.keys[_p_time_period]" in body
    assert (
        "_ax_time_period.keys[_p_time_period] - issuance_year >= 1" in body
        or "_ax_time_period.keys[_p_time_period] - issuance_year <= -1" in body
    )
    assert "issuance_year == _p_time_period" not in body
    assert "_p_time_period - issuance_year" not in body
    assert_package_matches_evaluator(path, document, tmp_path, "diagonal_open_band")


def test_labelled_row_diagonal_compares_year_keys(tmp_path: Path) -> None:
    """A runtime row axis is read back through its labeller keys."""
    years = (2024, 2025, 2026)
    issued = years

    def formula(year: int, issuance: int, column: str, row: int) -> str:
        if year == issuance:
            return f"={column}{row - 6}"
        return "=0"

    cells = _vintage_grid(years, issued, formula_for=formula, label_row=8, header_formulas=False)
    cells.update(_row_formulas(issued, start_row=2))
    for row_offset, _issuance in enumerate(issued):
        for index in range(len(years)):
            cells[f"{get_column_letter(index + 2)}{2 + row_offset}"] = float(index + 1)
    path = write_workbook(tmp_path / "row_axis.xlsx", {"Sheet": cells})
    _plant_cached_numbers(path, {f"A{2 + index}": year for index, year in enumerate(issued)})
    period = _column_header("TIME_PERIOD")
    issuance = _row_label("ISSUANCE_YEAR")
    document = _document(
        _series(
            "vintage",
            "Sheet!B2:D4",
            [issuance, period],
            ["ISSUANCE_YEAR", "TIME_PERIOD"],
            output=False,
        ),
        _series(
            "stock",
            "Sheet!B8:D10",
            [issuance, period],
            ["ISSUANCE_YEAR", "TIME_PERIOD"],
            output=True,
        ),
    )
    modules = generate_inverted(path, document)
    body = _body(modules["internals.py"], "stock")
    assert "_ax_issuance_year.keys[_p_issuance_year] == time_period" in body
    assert "_p_issuance_year == time_period" not in body
    assert_package_matches_evaluator(path, document, tmp_path, "diagonal_row_axis")


def test_two_labelled_axes_compare_positions(tmp_path: Path) -> None:
    """Two runtime axes stay in position space."""
    years = (2024, 2025, 2026)

    def formula(year: int, issuance: int, column: str, row: int) -> str:
        if year == issuance:
            return f"={column}{row - 6}"
        return "=0"

    cells = _vintage_grid(
        years, years, formula_for=formula, label_row=8, header_formulas=False, row_formulas=False
    )
    cells.update(_header_formulas(years))
    cells.update(_row_formulas(years, start_row=2))
    for row_offset in range(len(years)):
        for index in range(len(years)):
            cells[f"{get_column_letter(index + 2)}{2 + row_offset}"] = float(
                row_offset * 10 + index
            )
    path = write_workbook(tmp_path / "both_axes.xlsx", {"Sheet": cells})
    cached = _header_caches(years)
    cached.update({f"A{2 + index}": year for index, year in enumerate(years)})
    _plant_cached_numbers(path, cached)
    period = _column_header("TIME_PERIOD")
    issuance = _row_label("ISSUANCE_YEAR")
    document = _document(
        _series(
            "vintage",
            "Sheet!B2:D4",
            [issuance, period],
            ["ISSUANCE_YEAR", "TIME_PERIOD"],
            output=False,
        ),
        _series(
            "stock",
            "Sheet!B8:D10",
            [issuance, period],
            ["ISSUANCE_YEAR", "TIME_PERIOD"],
            output=True,
        ),
    )
    modules = generate_inverted(path, document)
    body = _body(modules["internals.py"], "stock")
    assert "_p_issuance_year == _p_time_period" in body
    assert "_ax_issuance_year.keys[_p_issuance_year] ==" not in body
    assert "== _ax_time_period.keys[_p_time_period]" not in body
    assert_package_matches_evaluator(path, document, tmp_path, "diagonal_both_positions")
