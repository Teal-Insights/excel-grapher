from __future__ import annotations

from dataclasses import dataclass

import pytest

from excel_grapher.core import CellValue, ExcelRange, XlError


@dataclass(frozen=True)
class WorkbookBounds:
    sheet: str
    min_row: int
    max_row: int
    min_col: int
    max_col: int

    def contains(self, rng: ExcelRange) -> bool:
        if rng.sheet != self.sheet:
            return False
        return (
            self.min_row <= rng.start_row <= rng.end_row <= self.max_row
            and self.min_col <= rng.start_col <= rng.end_col <= self.max_col
        )


def _make_bounds(sheet: str = "Sheet1") -> WorkbookBounds:
    # Simple 10x10 sheet for tests (rows 1-10, cols 1-10).
    return WorkbookBounds(sheet=sheet, min_row=1, max_row=10, min_col=1, max_col=10)


def _offset_range(
    base: ExcelRange,
    rows: CellValue,
    cols: CellValue,
    height: CellValue | None = None,
    width: CellValue | None = None,
    *,
    bounds: WorkbookBounds,
):
    # Local import to avoid hard dependency in core __all__ during initial RED phase.
    from excel_grapher.core.addressing import offset_range

    return offset_range(base, rows, cols, height, width, bounds=bounds)


def test_offset_range_single_cell_in_bounds() -> None:
    base = ExcelRange(sheet="Sheet1", start_row=1, start_col=1, end_row=1, end_col=1)
    bounds = _make_bounds()

    result = _offset_range(base, rows=1, cols=2, bounds=bounds)
    assert isinstance(result, ExcelRange)
    assert result.sheet == "Sheet1"
    assert (result.start_row, result.start_col, result.end_row, result.end_col) == (
        2,
        3,
        2,
        3,
    )
    assert bounds.contains(result)


def test_offset_range_single_cell_out_of_bounds_returns_ref() -> None:
    base = ExcelRange(sheet="Sheet1", start_row=1, start_col=1, end_row=1, end_col=1)
    bounds = _make_bounds()

    # Move above top row.
    result = _offset_range(base, rows=-2, cols=0, bounds=bounds)
    assert result == XlError.REF

    # Move left of first column.
    result = _offset_range(base, rows=0, cols=-2, bounds=bounds)
    assert result == XlError.REF


def test_offset_range_propagates_row_col_coercion_errors() -> None:
    base = ExcelRange(sheet="Sheet1", start_row=1, start_col=1, end_row=1, end_col=1)
    bounds = _make_bounds()

    # to_number should propagate VALUE error for non-numeric text.
    result = _offset_range(base, rows="foo", cols=0, bounds=bounds)
    assert result == XlError.VALUE

    result = _offset_range(base, rows=0, cols="bar", bounds=bounds)
    assert result == XlError.VALUE


def test_offset_range_uses_base_height_and_width_by_default() -> None:
    base = ExcelRange(sheet="Sheet1", start_row=2, start_col=3, end_row=4, end_col=5)
    bounds = _make_bounds()

    # Base shape is 3 rows x 3 cols.
    result = _offset_range(base, rows=1, cols=1, bounds=bounds)
    assert isinstance(result, ExcelRange)
    assert result.shape == (3, 3)


def test_offset_range_allows_positive_and_negative_offsets_within_bounds() -> None:
    base = ExcelRange(sheet="Sheet1", start_row=5, start_col=5, end_row=5, end_col=5)
    bounds = _make_bounds()

    result = _offset_range(base, rows=-1, cols=-2, bounds=bounds)
    assert isinstance(result, ExcelRange)
    assert (result.start_row, result.start_col) == (4, 3)
    assert bounds.contains(result)


def test_offset_range_height_and_width_semantics() -> None:
    base = ExcelRange(sheet="Sheet1", start_row=3, start_col=3, end_row=3, end_col=3)
    bounds = _make_bounds()

    # Explicit height/width override base shape.
    result = _offset_range(base, rows=0, cols=0, height=2, width=4, bounds=bounds)
    assert isinstance(result, ExcelRange)
    assert result.shape == (2, 4)

    # Zero or negative height/width should return VALUE.
    assert _offset_range(base, rows=0, cols=0, height=0, width=1, bounds=bounds) == XlError.VALUE
    assert _offset_range(base, rows=0, cols=0, height=-1, width=1, bounds=bounds) == XlError.VALUE
    assert _offset_range(base, rows=0, cols=0, height=1, width=0, bounds=bounds) == XlError.VALUE
    assert _offset_range(base, rows=0, cols=0, height=1, width=-1, bounds=bounds) == XlError.VALUE


def test_offset_range_out_of_bounds_target_returns_ref() -> None:
    base = ExcelRange(sheet="Sheet1", start_row=9, start_col=9, end_row=9, end_col=9)
    bounds = _make_bounds()

    # This would try to move the target row/col outside the 10x10 grid.
    result = _offset_range(base, rows=2, cols=0, bounds=bounds)
    assert result == XlError.REF

    result = _offset_range(base, rows=0, cols=2, bounds=bounds)
    assert result == XlError.REF


def _indirect_text_to_range(text: str, a1: bool, *, bounds: WorkbookBounds):
    from excel_grapher.core.addressing import indirect_text_to_range

    return indirect_text_to_range(text, a1, bounds=bounds)


@pytest.mark.parametrize(
    "text, expected",
    [
        ("Sheet1!A1", (1, 1, 1, 1)),
        ("Sheet1!A1:B2", (1, 1, 2, 2)),
        ("A1", (1, 1, 1, 1)),
        ("A1:B2", (1, 1, 2, 2)),
    ],
)
def test_indirect_text_to_range_a1_basic(text: str, expected: tuple[int, int, int, int]) -> None:
    bounds = _make_bounds()
    result = _indirect_text_to_range(text, a1=True, bounds=bounds)
    assert isinstance(result, ExcelRange)
    assert (result.start_row, result.start_col, result.end_row, result.end_col) == expected
    assert bounds.contains(result)


def test_indirect_text_to_range_out_of_bounds_returns_ref() -> None:
    bounds = _make_bounds()

    # Row beyond bounds.
    result = _indirect_text_to_range("Sheet1!A11", a1=True, bounds=bounds)
    assert result == XlError.REF

    # Column beyond bounds.
    result = _indirect_text_to_range("Sheet1!K1", a1=True, bounds=bounds)
    assert result == XlError.REF


@pytest.mark.parametrize("text", ["", "foo", "Sheet1!R1C1", "Sheet1!A1:", "Sheet1!A:B"])
def test_indirect_text_to_range_malformed_or_unsupported_returns_name_error(text: str) -> None:
    bounds = _make_bounds()
    result = _indirect_text_to_range(text, a1=True, bounds=bounds)
    assert result == XlError.NAME


def test_indirect_text_to_range_respects_a1_flag_for_unsupported_r1c1() -> None:
    bounds = _make_bounds()
    # When a1 is False (R1C1 mode), we currently treat all text as unsupported.
    result = _indirect_text_to_range("R1C1", a1=False, bounds=bounds)
    assert result == XlError.NAME


def test_index_excel_range_single_column_matches_lookup_style_index() -> None:
    from excel_grapher.core.addressing import index_excel_range

    base = ExcelRange(sheet="lookup", start_row=4, start_col=3, end_row=73, end_col=3)
    r = index_excel_range(base, 21, 1)
    assert isinstance(r, ExcelRange)
    assert (r.start_row, r.start_col, r.end_row, r.end_col) == (24, 3, 24, 3)


def test_index_excel_range_two_dimensional_cell() -> None:
    from excel_grapher.core.addressing import index_excel_range

    base = ExcelRange(sheet="S", start_row=1, start_col=1, end_row=3, end_col=3)
    r = index_excel_range(base, 2, 2)
    assert isinstance(r, ExcelRange)
    assert (r.start_row, r.start_col, r.end_row, r.end_col) == (2, 2, 2, 2)
