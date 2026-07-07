"""Type aliases for the compression package."""

from __future__ import annotations

from typing import TypeAlias

from excel_grapher.core.formula_ast import AstNode

from .nodes import ParallelFormulaNode, TacoPatternNode, TemplateAstNode

CompressedNode: TypeAlias = AstNode | ParallelFormulaNode | TacoPatternNode

__all__ = ["CompressedNode", "TemplateAstNode"]
