"""Compression artifact AST node types."""

from __future__ import annotations

from dataclasses import dataclass

from excel_grapher.core.formula_ast import AstNode
from excel_grapher.grapher.range_compression.types import Orientation, PatternKind


@dataclass(frozen=True, slots=True)
class ColumnVarCellRefNode:
    """Column placeholder in a parallel row template.

    Refs in the output column that change across the row normalize to this node;
    expansion substitutes the concrete column letter for `column_variable`.
    """

    column_variable: str = "COL"
    sheet: str | None = None
    row: int | None = None


@dataclass(frozen=True, slots=True)
class SubexpressionRefNode:
    """Reference to a hoisted `_cse!` binding in the compressed map."""

    ref_key: str


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
