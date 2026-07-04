"""Pre-parse formula ASTs during graph extraction for evaluator reuse."""

from __future__ import annotations

from excel_grapher.core.formula_ast import AstNode, parse

from .graph import DependencyGraph


def warm_preparsed_formulas(graph: DependencyGraph) -> dict[str, AstNode]:
    """Parse each distinct `normalized_formula` in `graph`.

    Returns a mapping from stripped normalized formula strings to AST roots.
    Duplicate formulas across cells share one entry.
    """
    warmed: dict[str, AstNode] = {}
    for _, node in graph.formula_nodes():
        nf = node.normalized_formula
        if nf is None or not isinstance(nf, str):
            continue
        stripped = nf.strip()
        if not stripped or stripped in warmed:
            continue
        warmed[stripped] = parse(stripped)
    return warmed
