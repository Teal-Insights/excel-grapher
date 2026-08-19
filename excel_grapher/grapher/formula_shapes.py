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
    """Intern punched AST shapes for each distinct `normalized_formula`.

    Duplicate formula strings share one parameter tuple. Formulas that differ
    only in cell/range addresses share one skeleton. Pass `parsed` from
    `warm_preparsed_formulas` to avoid a second parse.

    Re-call after loading a graph from JSON/pickle cache or mutating node
    formulas so `DependencyGraph.formula_shapes` stays aligned.

    Raises:
        FormulaParseError: If any distinct normalized formula is syntactically
            invalid (fail-fast, same contract as `warm_preparsed_formulas`).
    """
    formulas: list[str] = []
    for _, node in graph.formula_nodes():
        nf = node.normalized_formula
        if nf is None or not isinstance(nf, str):
            continue
        stripped = nf.strip()
        if stripped:
            formulas.append(stripped)
    return intern_formula_shapes(formulas, parsed=parsed)
