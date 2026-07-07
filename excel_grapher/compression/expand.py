"""Expand compressed artifacts back to per-cell ASTs."""

from __future__ import annotations

from collections.abc import Mapping

from excel_grapher.core.formula_ast import AstNode

from .types import CompressedNode


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
    raise NotImplementedError
