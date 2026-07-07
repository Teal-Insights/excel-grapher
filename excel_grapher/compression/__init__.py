"""Excel formula AST compression: structural rules, expansion, and parity.

This package applies semantics-preserving compression to per-cell formula ASTs,
producing mixed maps of plain ASTs, `_cse!` bindings, and artifact nodes
(`ParallelFormulaNode`, `TacoPatternNode`). Use `expand_compressed_to_cells`
to materialize artifacts back to per-cell ASTs for evaluation parity checks.
"""

from __future__ import annotations

from .ast_utils import (
    ast_contains_refs,
    is_cell_ast,
    is_literal_ast,
    map_ast,
    partition_compressed_map,
)
from .constant_folding import (
    apply_constant_folding,
    fold_literals_in_ast,
    try_fold_ast,
)
from .cse import (
    CseCandidate,
    CseConfig,
    CseGateRejection,
    CseGateResult,
    CseResult,
    SubtreeOccurrence,
    allocate_cse_key,
    apply_cell_cse,
    apply_hoist,
    enumerate_subtrees,
    find_shared_subtrees,
    hoist_common_subexpressions_to_fixpoint,
    hoist_one_subexpression,
    net_ast_savings,
    passes_cse_gates,
    subtree_node_count,
    subtree_signature,
)
from .engine import (
    apply_compression_rules,
    compression_rules_with_apply,
    get_rule_apply,
)
from .expand import (
    expand_compressed_to_cells,
    inline_subexpression_refs,
    materialize_parallel_node,
    shift_ast_to_cell,
    substitute_column_var,
)
from .nodes import (
    ColumnVarCellRefNode,
    ParallelFormulaNode,
    SubexpressionRefNode,
    TacoPatternNode,
)
from .parallel_row import (
    ParallelRun,
    RowCell,
    apply_parallel_row,
    build_parallel_node,
    find_parallel_runs,
    find_parallel_runs_in_map,
    group_row_cells,
    merge_adjacent_runs,
    parallel_artifact_key,
    split_contiguous_row_segments,
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
from .template_signature import (
    TemplateSignature,
    collect_cell_ref_addresses,
    fixed_cell_refs_in_group,
    template_signature,
    with_column_variable,
)
from .types import (
    CompressedNode,
    TemplateAstNode,
    is_synthetic_compressed_key,
    normalize_compressed_key,
)

__all__ = [
    "COMPRESSION_RULES",
    "ColumnVarCellRefNode",
    "CompressedNode",
    "CompressionParityMismatch",
    "CompressionStats",
    "CseCandidate",
    "CseConfig",
    "CseGateRejection",
    "CseGateResult",
    "CseResult",
    "RuleApplyFn",
    "RuleContribution",
    "RuleSpec",
    "SubexpressionRefNode",
    "SubtreeOccurrence",
    "TacoPatternNode",
    "ParallelFormulaNode",
    "ParallelRun",
    "RowCell",
    "TemplateAstNode",
    "TemplateSignature",
    "apply_cell_cse",
    "apply_compression_rules",
    "apply_constant_folding",
    "apply_hoist",
    "apply_parallel_row",
    "apply_pass_through",
    "assert_compression_parity",
    "allocate_cse_key",
    "ast_contains_refs",
    "build_parallel_node",
    "collect_cell_ref_addresses",
    "compare_compression_parity",
    "compression_rule_ids",
    "compression_rules_with_apply",
    "compression_values_equal",
    "empty_compression_stats",
    "enumerate_subtrees",
    "expand_compressed_to_cells",
    "find_parallel_runs",
    "find_parallel_runs_in_map",
    "find_shared_subtrees",
    "fixed_cell_refs_in_group",
    "fold_literals_in_ast",
    "get_rule_apply",
    "group_row_cells",
    "hoist_common_subexpressions_to_fixpoint",
    "hoist_one_subexpression",
    "identify_pass_through_cells",
    "inline_subexpression_refs",
    "is_cell_ast",
    "is_literal_ast",
    "is_synthetic_compressed_key",
    "map_ast",
    "materialize_parallel_node",
    "merge_adjacent_runs",
    "net_ast_savings",
    "normalize_compressed_key",
    "parallel_artifact_key",
    "partition_compressed_map",
    "passes_cse_gates",
    "replace_pass_through_refs",
    "resolve_pass_through_chains",
    "shift_ast_to_cell",
    "singleton_cell_ref_target",
    "split_contiguous_row_segments",
    "substitute_column_var",
    "subtree_node_count",
    "subtree_signature",
    "template_signature",
    "try_fold_ast",
    "with_column_variable",
]
