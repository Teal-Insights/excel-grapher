"""Template signatures and column-variable normalization for parallel rows."""

from __future__ import annotations

from collections.abc import Callable, Sequence

from excel_grapher.core.address_keys import normalize_key, parse_address
from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    BoolNode,
    CellRefNode,
    ColumnVarCellRefNode,
    EmptyArgNode,
    ErrorNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    StringNode,
    SubexpressionRefNode,
    UnaryOpNode,
    WholeColumnNode,
    WholeRowNode,
)

from .ast_utils import map_ast

TemplateSignature = tuple[object, ...]

__all__ = [
    "TemplateSignature",
    "collect_cell_ref_addresses",
    "fixed_cell_refs_in_group",
    "template_signature",
    "with_column_variable",
]


def collect_cell_ref_addresses(ast: AstNode) -> frozenset[str]:
    """Return normalized cell-reference addresses appearing in `ast`."""
    refs: set[str] = set()

    def _record(node: AstNode) -> None:
        if isinstance(node, CellRefNode):
            refs.add(normalize_key(node.address))

    _walk_ast(ast, _record)
    return frozenset(refs)


def fixed_cell_refs_in_group(asts: Sequence[AstNode]) -> frozenset[str]:
    """Return addresses that appear identically in every formula in `asts`."""
    if not asts:
        return frozenset()
    shared = collect_cell_ref_addresses(asts[0])
    for ast in asts[1:]:
        shared &= collect_cell_ref_addresses(ast)
    return frozenset(shared)


def with_column_variable(
    ast: AstNode,
    *,
    output_sheet: str,
    output_col: str,
    output_row: int,
    peer_asts: Sequence[AstNode],
) -> AstNode:
    """Normalize column-varying refs in `ast` to `ColumnVarCellRefNode` placeholders.

    Args:
        ast: Formula AST for one cell in a candidate parallel row group.
        output_sheet: Sheet containing the formula cell (reserved for future
            same-sheet relative refs).
        output_col: Output column letter for `ast`.
        output_row: Output row number for `ast`.
        peer_asts: ASTs for every formula in the candidate group, including `ast`.

    Returns:
        A copy of `ast` with non-fixed cell refs replaced by column placeholders.
    """
    _ = (output_sheet, output_col, output_row)
    fixed_refs = fixed_cell_refs_in_group(peer_asts)

    def _replace(node: AstNode) -> AstNode:
        if not isinstance(node, CellRefNode):
            return node
        address = normalize_key(node.address)
        if address in fixed_refs:
            return node
        sheet, coord = parse_address(address)
        row = int("".join(character for character in coord if character.isdigit()))
        return ColumnVarCellRefNode(column_variable="COL", sheet=sheet, row=row)

    return map_ast(ast, _replace)


def template_signature(ast: AstNode) -> TemplateSignature:
    """Return a hashable structural signature for comparing parallel templates."""
    return _signature_node(ast)


def _signature_node(node: AstNode) -> TemplateSignature:
    if isinstance(node, ColumnVarCellRefNode):
        return ("COL", node.column_variable, node.sheet, node.row)
    if isinstance(node, CellRefNode):
        return ("REF", normalize_key(node.address))
    if isinstance(node, RangeNode):
        return ("RNG", normalize_key(node.start), normalize_key(node.end))
    if isinstance(node, WholeColumnNode):
        return ("WCOL", node.sheet, node.column)
    if isinstance(node, WholeRowNode):
        return ("WROW", node.sheet, node.row)
    if isinstance(node, SubexpressionRefNode):
        return ("CSE", node.ref_key)
    if isinstance(node, NumberNode):
        return ("NUM", node.value)
    if isinstance(node, StringNode):
        return ("STR", node.value)
    if isinstance(node, BoolNode):
        return ("BOOL", node.value)
    if isinstance(node, ErrorNode):
        return ("ERR", node.error)
    if isinstance(node, EmptyArgNode):
        return ("EMPTY",)
    if isinstance(node, UnaryOpNode):
        return ("UNARY", node.op, _signature_node(node.operand))
    if isinstance(node, BinaryOpNode):
        return (
            "BIN",
            node.op,
            _signature_node(node.left),
            _signature_node(node.right),
        )
    if isinstance(node, FunctionCallNode):
        return ("FN", node.name.upper(), tuple(_signature_node(arg) for arg in node.args))
    return ("UNKNOWN", type(node).__name__)


def _walk_ast(node: AstNode, visit: Callable[[AstNode], None]) -> None:
    visit(node)
    if isinstance(node, FunctionCallNode):
        for arg in node.args:
            _walk_ast(arg, visit)
    elif isinstance(node, BinaryOpNode):
        _walk_ast(node.left, visit)
        _walk_ast(node.right, visit)
    elif isinstance(node, UnaryOpNode):
        _walk_ast(node.operand, visit)
