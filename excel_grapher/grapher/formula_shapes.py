"""Warm interned parameterized formula AST shapes on a dependency graph."""

from __future__ import annotations

from collections.abc import Mapping

from excel_grapher.core.formula_ast import AstNode
from excel_grapher.core.formula_shape import FormulaShapeTable, intern_formula_shapes

from .graph import DependencyGraph


def warm_formula_shapes(
    graph: DependencyGraph,
    *,
    parsed: Mapping[str, AstNode] | None = None,
) -> FormulaShapeTable:
    """Intern punched AST shapes keyed by each formula node's `NodeKey`.

    `Node.formula_ast` (when present) is the source of truth for formula
    meaning. The returned table is an optional acceleration overlay for eval
    and codegen; assign it to `graph.formula_shapes` to install it. Missing
    or dropped shapes fall back to AST. Correctness does not require the
    overlay to be warm.

    Prefer `Node.formula_ast` when present. Duplicate formula strings still
    share a skeleton; each cell gets its own binding. Pass `parsed` from
    `warm_preparsed_formulas` to avoid a second parse for nodes that have no
    stored AST. That overlay is absolute-bound; shapes for those nodes intern
    the bound tree.

    Shapes are opt-in. `create_dependency_graph` defaults to
    `warm_formula_shapes=False`; pass `True` there or call this function and
    assign the result.

    Compression, formula rewrite, and JSON/pickle load drop or omit
    `graph.formula_shapes` (`None`, not serialized). After those events, call
    this function again and assign the result if you want the overlay. The
    graph does not auto-rewarm.

    A live `FormulaEvaluator` compiles shape helpers from
    `graph.formula_shapes` at construction. Rewarming the overlay on the
    graph does not refresh `_shape_fns` on an existing evaluator; construct
    a new `FormulaEvaluator` after rewarming.

    Returns:
        Interned `FormulaShapeTable`. Assign to `graph.formula_shapes` to
        install the overlay.

    Raises:
        FormulaParseError: If any distinct normalized formula is syntactically
            invalid (fail-fast, same contract as `warm_preparsed_formulas`).
    """
    items: list[tuple[str, str | AstNode]] = []
    for key, node in graph.formula_nodes():
        if node.formula_ast is not None:
            items.append((key, node.formula_ast))
            continue
        nf = node.normalized_formula
        if nf is None or not isinstance(nf, str):
            continue
        stripped = nf.strip()
        if not stripped:
            continue
        ast = parsed.get(stripped) if parsed is not None else None
        items.append((key, ast if ast is not None else stripped))
    return intern_formula_shapes(items)
