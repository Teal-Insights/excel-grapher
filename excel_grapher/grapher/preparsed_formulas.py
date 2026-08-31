"""Pre-parse formula ASTs during graph extraction for evaluator reuse."""

from __future__ import annotations

from excel_grapher.core.formula_ast import AstNode, bind_axes, parse

from .graph import DependencyGraph


def warm_preparsed_formulas(graph: DependencyGraph) -> dict[str, AstNode]:
    """Parse each distinct `normalized_formula` in `graph`.

    Returns a mapping from stripped absolute A1 formula strings to fully
    absolute-bound AST roots (`bind_axes`). Duplicate formulas across cells
    share one entry. When a node already has `formula_ast`, relatives are
    resolved against that node's `NodeKey` instead of re-parsing, so a
    relative tree cannot poison another host that shares the same spelling.
    Nodes without a stored AST are parsed from the string (`parse` always
    yields `AbsoluteAxis`).

    This overlay is the evaluator's string-keyed `AstCache` fallback, not a
    `NodeKey` map. Per-node `formula_ast` remains the primary evaluation path.
    `move_node` that preserves resolved targets can leave entries valid;
    re-warm when resolved targets change.

    Re-call after loading a graph from JSON cache or mutating node formulas
    post-extraction so `DependencyGraph.preparsed_formulas` stays aligned.

    Raises:
        FormulaParseError: If any distinct normalized formula is syntactically
            invalid. Warming fails fast on the first bad formula (unlike
            `FormulaEvaluator`, which raises `ParseError` per cell at evaluate
            time).
    """
    warmed: dict[str, AstNode] = {}
    for key, node in graph.formula_nodes():
        nf = node.normalized_formula
        if nf is None or not isinstance(nf, str):
            continue
        stripped = nf.strip()
        if not stripped or stripped in warmed:
            continue
        if node.formula_ast is not None:
            host = node.address if node.address is not None else key
            warmed[stripped] = bind_axes(node.formula_ast, host)
        else:
            warmed[stripped] = parse(stripped)
    return warmed
