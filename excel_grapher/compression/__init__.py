"""Excel formula AST compression: structural rules, expansion, and parity.

This package applies semantics-preserving compression to per-cell formula ASTs,
producing mixed maps of plain ASTs, `_cse!` bindings, and artifact nodes
(`ParallelFormulaNode`, `TacoPatternNode`). Use `expand_compressed_to_cells`
to materialize artifacts back to per-cell ASTs for evaluation parity checks.
"""

from __future__ import annotations

from .expand import (
    expand_compressed_to_cells,
    inline_subexpression_refs,
    shift_ast_to_cell,
    substitute_column_var,
)
from .nodes import (
    ColumnVarCellRefNode,
    ParallelFormulaNode,
    SubexpressionRefNode,
    TacoPatternNode,
)
from .parity import (
    CompressionParityMismatch,
    assert_compression_parity,
    compare_compression_parity,
    compression_values_equal,
)
from .rules import (
    COMPRESSION_RULES,
    CompressionStats,
    RuleContribution,
    RuleSpec,
    compression_rule_ids,
    empty_compression_stats,
)
from .types import CompressedNode, TemplateAstNode

__all__ = [
    "COMPRESSION_RULES",
    "ColumnVarCellRefNode",
    "CompressedNode",
    "CompressionParityMismatch",
    "TemplateAstNode",
    "CompressionStats",
    "ParallelFormulaNode",
    "RuleContribution",
    "RuleSpec",
    "SubexpressionRefNode",
    "TacoPatternNode",
    "assert_compression_parity",
    "compare_compression_parity",
    "compression_rule_ids",
    "compression_values_equal",
    "empty_compression_stats",
    "expand_compressed_to_cells",
    "inline_subexpression_refs",
    "shift_ast_to_cell",
    "substitute_column_var",
]
