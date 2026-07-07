"""Rule 1: direct cell-reference (pass-through) elimination."""

from __future__ import annotations

from collections.abc import Mapping

from excel_grapher.core.address_keys import normalize_key
from excel_grapher.core.formula_ast import AstNode, CellRefNode, UnaryOpNode

from .ast_utils import map_ast
from .stats import CompressionStats


def singleton_cell_ref_target(ast: AstNode) -> str | None:
    """Return the referenced address when `ast` is a singleton cell reference."""
    node = ast
    while isinstance(node, UnaryOpNode) and node.op == "+":
        node = node.operand
    if isinstance(node, CellRefNode):
        return normalize_key(node.address)
    return None


def identify_pass_through_cells(ast_map: Mapping[str, AstNode]) -> dict[str, str]:
    """Return transit cell keys mapped to their direct reference targets."""
    mapping: dict[str, str] = {}
    for cell_key, ast in ast_map.items():
        target = singleton_cell_ref_target(ast)
        if target is None:
            continue
        transit_key = normalize_key(cell_key)
        if transit_key == target:
            continue
        mapping[transit_key] = target
    return mapping


def resolve_pass_through_chains(mapping: Mapping[str, str]) -> dict[str, str]:
    """Resolve transitive pass-through chains to ultimate targets."""
    resolved = dict(mapping)
    for transit_key in mapping:
        target = resolved[transit_key]
        seen = {transit_key}
        while target in resolved and target not in seen:
            seen.add(target)
            target = resolved[target]
        resolved[transit_key] = target
    return resolved


def replace_pass_through_refs(
    ast: AstNode,
    pass_through: Mapping[str, str],
) -> AstNode:
    """Rewrite references to pass-through cells to their ultimate targets."""

    def _replace(node: AstNode) -> AstNode:
        if isinstance(node, CellRefNode):
            ref_key = normalize_key(node.address)
            target = pass_through.get(ref_key)
            if target is not None:
                return CellRefNode(target)
        return node

    return map_ast(ast, _replace)


def apply_pass_through(
    ast_map: Mapping[str, AstNode],
    stats: CompressionStats | None = None,
) -> dict[str, AstNode]:
    """Apply pass-through elimination to a per-cell AST map."""
    pass_through = resolve_pass_through_chains(identify_pass_through_cells(ast_map))
    result: dict[str, AstNode] = {}
    transforms = 0

    for cell_key, ast in ast_map.items():
        normalized_key = normalize_key(cell_key)
        if normalized_key in pass_through:
            result[normalized_key] = ast
            continue
        rewritten = replace_pass_through_refs(ast, pass_through)
        if rewritten != ast:
            transforms += 1
        result[normalized_key] = rewritten

    if stats is not None:
        stats.contribution_for("pass_through").record(
            in_place_transforms=transforms,
            cells_affected=transforms,
        )
    return result
