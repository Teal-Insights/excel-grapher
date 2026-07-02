"""Core types for TACO-style range-pattern compression."""

from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import format_cell_key, parse_address


class PatternKind(StrEnum):
    """TACO range-reference pattern kinds."""

    rr = "RR"
    rf = "RF"
    fr = "FR"
    ff = "FF"
    rr_chain = "RR-Chain"
    single = "Single"


class Orientation(StrEnum):
    """Axis along which a compressed formula run advances."""

    column = "column"
    row = "row"


@dataclass(frozen=True, slots=True)
class RangeRef:
    """Sheet-qualified rectangular cell range."""

    sheet: str
    min_col: str
    min_row: int
    max_col: str
    max_row: int

    def __post_init__(self) -> None:
        c1 = fastpyxl.utils.cell.column_index_from_string(self.min_col)
        c2 = fastpyxl.utils.cell.column_index_from_string(self.max_col)
        if c1 > c2 or self.min_row > self.max_row:
            raise ValueError(f"Invalid range bounds: {self!r}")

    @classmethod
    def single_cell(cls, sheet: str, column: str, row: int) -> RangeRef:
        """Return a 1x1 range."""
        return cls(
            sheet=sheet,
            min_col=column,
            min_row=row,
            max_col=column,
            max_row=row,
        )

    @classmethod
    def column_span(cls, sheet: str, column: str, first_row: int, last_row: int) -> RangeRef:
        """Return a single-column vertical range."""
        lo, hi = (first_row, last_row) if first_row <= last_row else (last_row, first_row)
        return cls(sheet=sheet, min_col=column, min_row=lo, max_col=column, max_row=hi)

    @classmethod
    def row_span(cls, sheet: str, row: int, first_col: str, last_col: str) -> RangeRef:
        """Return a single-row horizontal range."""
        c1 = fastpyxl.utils.cell.column_index_from_string(first_col)
        c2 = fastpyxl.utils.cell.column_index_from_string(last_col)
        clo, chi = (c1, c2) if c1 <= c2 else (c2, c1)
        return cls(
            sheet=sheet,
            min_col=fastpyxl.utils.cell.get_column_letter(clo),
            min_row=row,
            max_col=fastpyxl.utils.cell.get_column_letter(chi),
            max_row=row,
        )

    @classmethod
    def rectangle(
        cls,
        sheet: str,
        min_col: str,
        min_row: int,
        max_col: str,
        max_row: int,
    ) -> RangeRef:
        """Return a normalized rectangular range."""
        c1 = fastpyxl.utils.cell.column_index_from_string(min_col)
        c2 = fastpyxl.utils.cell.column_index_from_string(max_col)
        r1, r2 = (min_row, max_row) if min_row <= max_row else (max_row, min_row)
        clo, chi = (c1, c2) if c1 <= c2 else (c2, c1)
        return cls(
            sheet=sheet,
            min_col=fastpyxl.utils.cell.get_column_letter(clo),
            min_row=r1,
            max_col=fastpyxl.utils.cell.get_column_letter(chi),
            max_row=r2,
        )

    def contains(self, key: str) -> bool:
        """Return True when `key` lies inside this range."""
        sheet, coord = parse_address(key)
        if sheet != self.sheet:
            return False
        col, row = fastpyxl.utils.cell.coordinate_from_string(coord)
        col_i = fastpyxl.utils.cell.column_index_from_string(col)
        c_lo = fastpyxl.utils.cell.column_index_from_string(self.min_col)
        c_hi = fastpyxl.utils.cell.column_index_from_string(self.max_col)
        return c_lo <= col_i <= c_hi and self.min_row <= row <= self.max_row

    def cell_keys(self) -> list[str]:
        """Expand this range to sheet-qualified cell keys."""
        c_lo = fastpyxl.utils.cell.column_index_from_string(self.min_col)
        c_hi = fastpyxl.utils.cell.column_index_from_string(self.max_col)
        out: list[str] = []
        for row in range(self.min_row, self.max_row + 1):
            for col_i in range(c_lo, c_hi + 1):
                col = fastpyxl.utils.cell.get_column_letter(col_i)
                out.append(format_cell_key(self.sheet, col, row))
        return out


@dataclass(frozen=True, slots=True)
class PatternMeta:
    """Metadata describing how precedent and dependent ranges relate."""

    kind: PatternKind
    col_offset: int = 0
    row_offset: int = 0
    fixed_head_col: str | None = None
    fixed_head_row: int | None = None
    fixed_tail_col: str | None = None
    fixed_tail_row: int | None = None
    orientation: Orientation = Orientation.column


@dataclass(frozen=True, slots=True)
class CompressedEdge:
    """One compressed dependency edge between rectangular ranges."""

    precedent: RangeRef
    dependent: RangeRef
    meta: PatternMeta


@dataclass(frozen=True, slots=True)
class SingleEdge:
    """Uncompressed cell-level fallback edge."""

    precedent: str
    dependent: str
