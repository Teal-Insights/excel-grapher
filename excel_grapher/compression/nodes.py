"""Compression artifact AST node types."""

from __future__ import annotations

from dataclasses import dataclass

import fastpyxl.utils.cell

from excel_grapher.core.formula_ast import (
    AstNode,
    ColumnVarCellRefNode,
    SubexpressionRefNode,
)
from excel_grapher.grapher.range_compression.types import Orientation, PatternKind, RangeRef

__all__ = [
    "ColumnVarCellRefNode",
    "ParallelFormulaNode",
    "SubexpressionRefNode",
    "TacoPatternNode",
    "TemplateAstNode",
]


def _column_index(column: str) -> int:
    return fastpyxl.utils.cell.column_index_from_string(column)


@dataclass(frozen=True, slots=True)
class ParallelFormulaNode:
    """Compressed parallel row formulas sharing a template."""

    sheet: str
    template: AstNode
    start_col: str
    end_col: str
    output_row: int
    column_variable: str = "COL"
    condition: AstNode | None = None
    if_true: AstNode | None = None
    if_false: AstNode | None = None

    def __post_init__(self) -> None:
        if not self.sheet:
            raise ValueError("sheet must be a non-empty string")
        if not self.column_variable:
            raise ValueError("column_variable must be a non-empty string")
        if self.output_row < 1:
            raise ValueError(f"output_row must be >= 1 (got {self.output_row})")
        if _column_index(self.start_col) > _column_index(self.end_col):
            raise ValueError(
                f"start_col must be <= end_col (got {self.start_col!r} > {self.end_col!r})"
            )

    @property
    def column_count(self) -> int:
        """Return the number of output columns covered by this parallel group."""
        return _column_index(self.end_col) - _column_index(self.start_col) + 1


@dataclass(frozen=True, slots=True)
class TacoPatternNode:
    """Compressed TACO autofill formula range."""

    kind: PatternKind
    sheet: str
    min_col: str
    min_row: int
    max_col: str
    max_row: int
    template: AstNode
    orientation: Orientation

    def __post_init__(self) -> None:
        if self.kind is PatternKind.single:
            raise ValueError("TacoPatternNode cannot use PatternKind.single")
        if not self.sheet:
            raise ValueError("sheet must be a non-empty string")
        RangeRef(
            sheet=self.sheet,
            min_col=self.min_col,
            min_row=self.min_row,
            max_col=self.max_col,
            max_row=self.max_row,
        )

    @property
    def cell_count(self) -> int:
        """Return the number of formula cells represented by this artifact."""
        cols = _column_index(self.max_col) - _column_index(self.min_col) + 1
        rows = self.max_row - self.min_row + 1
        return cols * rows


# Templates may embed compression placeholders; alias documents that intent.
TemplateAstNode = AstNode
