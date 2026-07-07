"""Expand compressed artifacts back to per-cell ASTs."""

from __future__ import annotations

from collections.abc import Callable, Mapping

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import format_cell_key, parse_address
from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    CellRefNode,
    ColumnVarCellRefNode,
    FunctionCallNode,
    RangeNode,
    SubexpressionRefNode,
    UnaryOpNode,
)

from .nodes import ParallelFormulaNode, TacoPatternNode
from .types import CompressedNode

_CSE_KEY_PREFIX = "_cse!"


def expand_compressed_to_cells(
    compressed: Mapping[str, CompressedNode],
) -> dict[str, AstNode]:
    """Materialize compressed artifacts to one AST per formula cell.

    Args:
        compressed: Mixed map of per-cell ASTs, `_cse!` bindings, and artifact
            nodes (`ParallelFormulaNode`, `TacoPatternNode`).

    Returns:
        Sheet-qualified cell keys mapped to expanded per-cell ASTs.
    """
    cse_bindings = {
        key: node
        for key, node in compressed.items()
        if _is_cse_key(key) and isinstance(node, AstNode)
    }
    expanded: dict[str, AstNode] = {}
    for key, node in compressed.items():
        if _is_cse_key(key):
            continue
        if isinstance(node, ParallelFormulaNode):
            expanded.update(_expand_parallel_node(node))
        elif isinstance(node, TacoPatternNode):
            expanded.update(_expand_taco_node(node))
        else:
            expanded[key] = node

    if not cse_bindings:
        return expanded
    return {
        cell_key: inline_subexpression_refs(ast_node, cse_bindings)
        for cell_key, ast_node in expanded.items()
    }


def substitute_column_var(
    node: AstNode,
    *,
    col: str,
    output_sheet: str,
    output_row: int,
) -> AstNode:
    """Replace `ColumnVarCellRefNode` placeholders with a concrete column."""
    return _transform_ast(
        node,
        lambda current: _substitute_column_var_node(
            current,
            col=col,
            output_sheet=output_sheet,
            output_row=output_row,
        ),
    )


def inline_subexpression_refs(
    node: AstNode,
    cse_bindings: Mapping[str, AstNode],
) -> AstNode:
    """Inline `_cse!` references using hoisted binding ASTs."""
    return _transform_ast(
        node,
        lambda current: (
            cse_bindings[current.ref_key] if isinstance(current, SubexpressionRefNode) else current
        ),
    )


def shift_ast_to_cell(
    node: AstNode,
    *,
    anchor_col: str,
    anchor_row: int,
    target_col: str,
    target_row: int,
) -> AstNode:
    """Shift relative cell and range references from an anchor cell to a target."""
    dcol = _column_index(target_col) - _column_index(anchor_col)
    drow = target_row - anchor_row
    if dcol == 0 and drow == 0:
        return node
    return _transform_ast(
        node,
        lambda current: _shift_ref_node(current, dcol=dcol, drow=drow),
    )


def _is_cse_key(key: str) -> bool:
    return key.startswith(_CSE_KEY_PREFIX)


def _column_index(column: str) -> int:
    return fastpyxl.utils.cell.column_index_from_string(column)


def _column_letters(start_col: str, end_col: str) -> list[str]:
    start_i = _column_index(start_col)
    end_i = _column_index(end_col)
    return [fastpyxl.utils.cell.get_column_letter(col_i) for col_i in range(start_i, end_i + 1)]


def _iter_range_cell_keys(
    *,
    sheet: str,
    min_col: str,
    min_row: int,
    max_col: str,
    max_row: int,
) -> list[str]:
    keys: list[str] = []
    start_col_i = _column_index(min_col)
    end_col_i = _column_index(max_col)
    for row in range(min_row, max_row + 1):
        for col_i in range(start_col_i, end_col_i + 1):
            col = fastpyxl.utils.cell.get_column_letter(col_i)
            keys.append(format_cell_key(sheet, col, row))
    return keys


def _expand_parallel_node(node: ParallelFormulaNode) -> dict[str, AstNode]:
    expanded: dict[str, AstNode] = {}
    for col in _column_letters(node.start_col, node.end_col):
        cell_key = format_cell_key(node.sheet, col, node.output_row)
        expanded[cell_key] = substitute_column_var(
            node.template,
            col=col,
            output_sheet=node.sheet,
            output_row=node.output_row,
        )
    return expanded


def _expand_taco_node(node: TacoPatternNode) -> dict[str, AstNode]:
    expanded: dict[str, AstNode] = {}
    for cell_key in _iter_range_cell_keys(
        sheet=node.sheet,
        min_col=node.min_col,
        min_row=node.min_row,
        max_col=node.max_col,
        max_row=node.max_row,
    ):
        _, coord = parse_address(cell_key)
        target_col, target_row = fastpyxl.utils.cell.coordinate_from_string(coord)
        expanded[cell_key] = shift_ast_to_cell(
            node.template,
            anchor_col=node.min_col,
            anchor_row=node.min_row,
            target_col=target_col,
            target_row=target_row,
        )
    return expanded


def _substitute_column_var_node(
    node: AstNode,
    *,
    col: str,
    output_sheet: str,
    output_row: int,
) -> AstNode:
    if not isinstance(node, ColumnVarCellRefNode):
        return node
    sheet = node.sheet if node.sheet is not None else output_sheet
    row = node.row if node.row is not None else output_row
    return CellRefNode(format_cell_key(sheet, col, row))


def _shift_ref_node(node: AstNode, *, dcol: int, drow: int) -> AstNode:
    if isinstance(node, CellRefNode):
        return CellRefNode(_shift_address(node.address, dcol=dcol, drow=drow))
    if isinstance(node, RangeNode):
        return RangeNode(
            start=_shift_address(node.start, dcol=dcol, drow=drow),
            end=_shift_address(node.end, dcol=dcol, drow=drow),
        )
    return node


def _shift_address(address: str, *, dcol: int, drow: int) -> str:
    sheet, coord = parse_address(address)
    coord = coord.replace("$", "")
    col, row = fastpyxl.utils.cell.coordinate_from_string(coord)
    col_i = _column_index(col) + dcol
    row_i = row + drow
    return format_cell_key(
        sheet,
        fastpyxl.utils.cell.get_column_letter(col_i),
        row_i,
    )


def _transform_ast(node: AstNode, transform: Callable[[AstNode], AstNode]) -> AstNode:
    replacement = transform(node)
    if replacement is not node:
        return replacement

    if isinstance(node, FunctionCallNode):
        return FunctionCallNode(node.name, [_transform_ast(arg, transform) for arg in node.args])
    if isinstance(node, BinaryOpNode):
        return BinaryOpNode(
            node.op,
            _transform_ast(node.left, transform),
            _transform_ast(node.right, transform),
        )
    if isinstance(node, UnaryOpNode):
        return UnaryOpNode(node.op, _transform_ast(node.operand, transform))
    return node
