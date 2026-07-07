"""Excel formula AST compression: structural rules, expansion, and parity.

This package applies semantics-preserving compression to per-cell formula ASTs,
producing mixed maps of plain ASTs, `_cse!` bindings, and artifact nodes
(`ParallelFormulaNode`, `TacoPatternNode`). Use `expand_compressed_to_cells`
to materialize artifacts back to per-cell ASTs for evaluation parity checks.
"""

from __future__ import annotations

from .ast_utils import ast_contains_refs, is_literal_ast, map_ast
from .constant_folding import (
    apply_constant_folding,
    fold_literals_in_ast,
    try_fold_ast,
)
from .engine import (
    apply_compression_rules,
    compression_rules_with_apply,
    get_rule_apply,
)
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
from .pass_through import (
    apply_pass_through,
    identify_pass_through_cells,
    replace_pass_through_refs,
    resolve_pass_through_chains,
    singleton_cell_ref_target,
)
from .rules import COMPRESSION_RULES, RuleApplyFn, RuleSpec, compression_rule_ids
from .stats import CompressionStats, RuleContribution, empty_compression_stats
from .types import CompressedNode, TemplateAstNode

__all__ = [
    "COMPRESSION_RULES",
    "ColumnVarCellRefNode",
    "CompressedNode",
    "CompressionParityMismatch",
    "CompressionStats",
    "RuleApplyFn",
    "RuleContribution",
    "RuleSpec",
    "SubexpressionRefNode",
    "TacoPatternNode",
    "ParallelFormulaNode",
    "TemplateAstNode",
    "apply_compression_rules",
    "apply_constant_folding",
    "apply_pass_through",
    "assert_compression_parity",
    "ast_contains_refs",
    "compare_compression_parity",
    "compression_rule_ids",
    "compression_rules_with_apply",
    "compression_values_equal",
    "empty_compression_stats",
    "expand_compressed_to_cells",
    "fold_literals_in_ast",
    "get_rule_apply",
    "identify_pass_through_cells",
    "inline_subexpression_refs",
    "is_literal_ast",
    "map_ast",
    "replace_pass_through_refs",
    "resolve_pass_through_chains",
    "shift_ast_to_cell",
    "singleton_cell_ref_target",
    "substitute_column_var",
    "try_fold_ast",
]
