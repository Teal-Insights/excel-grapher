"""Common subexpression elimination: subtree analysis and cost gates."""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Collection, Mapping
from dataclasses import dataclass
from enum import Enum
from typing import TypeAlias

from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    FunctionCallNode,
    SubexpressionRefNode,
    UnaryOpNode,
)

from .ast_utils import merge_compressed_map, partition_compressed_map
from .constant_folding import fold_literals_in_ast
from .stats import CompressionStats
from .template_signature import TemplateSignature, template_signature
from .types import CompressedNode

SubtreePath: TypeAlias = tuple[str | int, ...]
SubtreeSignature = TemplateSignature

_CSE_KEY_PREFIX = "_cse!"

__all__ = [
    "CseCandidate",
    "CseConfig",
    "CseGateRejection",
    "CseGateResult",
    "CseResult",
    "SubtreeOccurrence",
    "SubtreePath",
    "SubtreeSignature",
    "allocate_cse_key",
    "apply_cell_cse",
    "apply_hoist",
    "enumerate_subtrees",
    "find_shared_subtrees",
    "hoist_common_subexpressions_to_fixpoint",
    "hoist_one_subexpression",
    "net_ast_savings",
    "passes_cse_gates",
    "subtree_node_count",
    "subtree_signature",
]


@dataclass(frozen=True, slots=True)
class CseConfig:
    """Cost gates for cell-level common subexpression elimination."""

    min_occurrences: int = 3
    min_subtree_nodes: int = 3
    min_net_ast_savings: int = 4


class CseGateRejection(Enum):
    """Why a shared subtree failed CSE cost gates."""

    INSUFFICIENT_OCCURRENCES = "insufficient_occurrences"
    SUBTREE_TOO_SMALL = "subtree_too_small"
    INSUFFICIENT_NET_SAVINGS = "insufficient_net_savings"


@dataclass(frozen=True, slots=True)
class CseGateResult:
    """Outcome of evaluating CSE cost gates for one candidate."""

    passes: bool
    rejection: CseGateRejection | None = None


@dataclass(frozen=True, slots=True)
class CseCandidate:
    """A shared subtree eligible for hoisting evaluation."""

    signature: SubtreeSignature
    ast: AstNode
    occurrences: tuple[tuple[str, SubtreePath], ...]

    @property
    def occurrence_count(self) -> int:
        """Return how many times this subtree appears across the cell map."""
        return len(self.occurrences)

    @classmethod
    def from_cell_map(cls, cell_map: Mapping[str, AstNode]) -> tuple[CseCandidate, ...]:
        """Build hoist candidates from shared-subtree groups in formula cells."""
        candidates: list[CseCandidate] = []
        formula_cells = _formula_cells(cell_map)
        for signature, occurrences in _group_occurrences(
            find_shared_subtrees(formula_cells)
        ).items():
            candidates.append(
                cls(
                    signature=signature,
                    ast=occurrences[0].ast,
                    occurrences=tuple((item.cell_key, item.path) for item in occurrences),
                )
            )
        return tuple(candidates)


@dataclass(frozen=True, slots=True)
class SubtreeOccurrence:
    cell_key: str
    path: SubtreePath
    ast: AstNode


def subtree_signature(ast: AstNode) -> SubtreeSignature:
    """Return a hashable structural signature for a subtree root."""
    return template_signature(ast)


def subtree_node_count(ast: AstNode) -> int:
    """Count AST nodes in the subtree rooted at `ast`."""
    if isinstance(ast, FunctionCallNode):
        return 1 + sum(subtree_node_count(arg) for arg in ast.args)
    if isinstance(ast, BinaryOpNode):
        return 1 + subtree_node_count(ast.left) + subtree_node_count(ast.right)
    if isinstance(ast, UnaryOpNode):
        return 1 + subtree_node_count(ast.operand)
    return 1


def enumerate_subtrees(ast: AstNode) -> tuple[tuple[SubtreePath, AstNode], ...]:
    """Return `(path, subtree)` pairs for every compound subtree in `ast`."""
    results: list[tuple[SubtreePath, AstNode]] = []

    def _visit(node: AstNode, path: SubtreePath) -> None:
        if _is_compound_subtree_root(node):
            results.append((path, node))
        if isinstance(node, FunctionCallNode):
            for index, arg in enumerate(node.args):
                _visit(arg, path + ("args", index))
        elif isinstance(node, BinaryOpNode):
            _visit(node.left, path + ("left",))
            _visit(node.right, path + ("right",))
        elif isinstance(node, UnaryOpNode):
            _visit(node.operand, path + ("operand",))

    _visit(ast, ())
    return tuple(results)


def find_shared_subtrees(
    cell_map: Mapping[str, AstNode],
) -> dict[SubtreeSignature, tuple[SubtreeOccurrence, ...]]:
    """Group compound subtrees that appear in multiple places across `cell_map`."""
    grouped: dict[SubtreeSignature, list[SubtreeOccurrence]] = defaultdict(list)
    for cell_key, ast in cell_map.items():
        for path, subtree in enumerate_subtrees(ast):
            grouped[subtree_signature(subtree)].append(
                SubtreeOccurrence(cell_key=cell_key, path=path, ast=subtree)
            )
    return {signature: tuple(items) for signature, items in grouped.items()}


def net_ast_savings(candidate: CseCandidate) -> int:
    """Return `(occurrences - 1) * (subtree_nodes - 1)` for a candidate."""
    nodes = subtree_node_count(candidate.ast)
    return (candidate.occurrence_count - 1) * (nodes - 1)


def passes_cse_gates(candidate: CseCandidate, config: CseConfig) -> CseGateResult:
    """Return whether `candidate` passes all configured CSE cost gates."""
    if candidate.occurrence_count < config.min_occurrences:
        return CseGateResult(False, CseGateRejection.INSUFFICIENT_OCCURRENCES)
    nodes = subtree_node_count(candidate.ast)
    if nodes < config.min_subtree_nodes:
        return CseGateResult(False, CseGateRejection.SUBTREE_TOO_SMALL)
    if net_ast_savings(candidate) < config.min_net_ast_savings:
        return CseGateResult(False, CseGateRejection.INSUFFICIENT_NET_SAVINGS)
    return CseGateResult(True)


@dataclass
class CseResult:
    """Counters from one or more CSE hoist rounds."""

    binding_key: str | None = None
    hoisted: bool = False
    binding_sites: int = 0
    redundant_evaluations_eliminated: int = 0
    ast_subnodes_saved: int = 0
    candidates_rejected: int = 0
    cse_fixpoint_rounds: int = 0
    post_cse_formula_folds: int = 0


def allocate_cse_key(existing_keys: Collection[str]) -> str:
    """Return the lowest unused `_cse!N` key not present in `existing_keys`."""
    index = 0
    while True:
        key = f"{_CSE_KEY_PREFIX}{index}"
        if key not in existing_keys:
            return key
        index += 1


def apply_hoist(
    cell_map: Mapping[str, AstNode],
    candidate: CseCandidate,
    key: str,
) -> dict[str, AstNode]:
    """Add a `_cse!` binding and replace all candidate occurrences with refs."""
    ref = SubexpressionRefNode(key)
    paths_by_cell: dict[str, list[SubtreePath]] = defaultdict(list)
    for cell_key, path in candidate.occurrences:
        paths_by_cell[cell_key].append(path)

    updated = dict(cell_map)
    for cell_key, paths in paths_by_cell.items():
        ast = updated[cell_key]
        for path in sorted(paths, key=len, reverse=True):
            ast = _replace_subtree_at_path(ast, path, ref)
        updated[cell_key] = ast
    updated[key] = candidate.ast
    return updated


def hoist_one_subexpression(
    cell_map: Mapping[str, AstNode],
    *,
    config: CseConfig | None = None,
) -> tuple[dict[str, AstNode], CseResult]:
    """Hoist the best gated shared subtree once, or return the input unchanged."""
    config = config or CseConfig()
    bindings = _cse_bindings(cell_map)
    formula_cells = _formula_cells(cell_map)
    candidates = CseCandidate.from_cell_map(cell_map)
    passing: list[CseCandidate] = []
    rejected = 0
    for candidate in candidates:
        if passes_cse_gates(candidate, config).passes:
            passing.append(candidate)
        else:
            rejected += 1
    if not passing:
        return dict(cell_map), CseResult(candidates_rejected=rejected)

    best = max(passing, key=net_ast_savings)
    key = allocate_cse_key(cell_map)
    updated_formulas = apply_hoist(formula_cells, best, key)
    return {**bindings, **updated_formulas}, CseResult(
        binding_key=key,
        hoisted=True,
        binding_sites=1,
        redundant_evaluations_eliminated=best.occurrence_count - 1,
        ast_subnodes_saved=net_ast_savings(best),
        candidates_rejected=rejected,
    )


def hoist_common_subexpressions_to_fixpoint(
    cell_map: Mapping[str, AstNode],
    *,
    config: CseConfig | None = None,
) -> tuple[dict[str, AstNode], CseResult]:
    """Repeatedly hoist shared subtrees until no gated candidate remains."""
    config = config or CseConfig()
    working = dict(cell_map)
    total = CseResult()

    while True:
        working, round_result = hoist_one_subexpression(working, config=config)
        total.candidates_rejected += round_result.candidates_rejected
        if not round_result.hoisted:
            break

        total.hoisted = True
        total.binding_key = round_result.binding_key
        total.binding_sites += round_result.binding_sites
        total.redundant_evaluations_eliminated += round_result.redundant_evaluations_eliminated
        total.ast_subnodes_saved += round_result.ast_subnodes_saved
        total.cse_fixpoint_rounds += 1

        working, folds = _fold_cell_map(working)
        total.post_cse_formula_folds += folds

    return working, total


def apply_cell_cse(
    compressed_map: Mapping[str, CompressedNode],
    stats: CompressionStats | None = None,
) -> dict[str, CompressedNode]:
    """Apply cell-level CSE to formula cells while preserving compressed artifacts."""
    cell_map, artifacts = partition_compressed_map(compressed_map)
    hoisted, cse_result = hoist_common_subexpressions_to_fixpoint(cell_map)

    if stats is not None:
        stats.cse_fixpoint_rounds = cse_result.cse_fixpoint_rounds
        stats.binding_sites += cse_result.binding_sites
        stats.redundant_evaluations_eliminated += cse_result.redundant_evaluations_eliminated
        stats.ast_subnodes_saved += cse_result.ast_subnodes_saved
        stats.post_cse_formula_folds += cse_result.post_cse_formula_folds
        stats.contribution_for("common_subexpression").record(
            binding_sites=cse_result.binding_sites,
            redundant_evaluations_eliminated=cse_result.redundant_evaluations_eliminated,
            ast_subnodes_saved=cse_result.ast_subnodes_saved,
            candidates_rejected=cse_result.candidates_rejected,
            cells_affected=_cells_changed(cell_map, hoisted),
        )

    return merge_compressed_map(artifacts, hoisted)


def _replace_subtree_at_path(
    ast: AstNode,
    path: SubtreePath,
    replacement: AstNode,
) -> AstNode:
    if not path:
        return replacement
    step, *rest = path
    tail = tuple(rest)
    if step == "left" and isinstance(ast, BinaryOpNode):
        return BinaryOpNode(
            ast.op, _replace_subtree_at_path(ast.left, tail, replacement), ast.right
        )
    if step == "right" and isinstance(ast, BinaryOpNode):
        return BinaryOpNode(
            ast.op, ast.left, _replace_subtree_at_path(ast.right, tail, replacement)
        )
    if step == "operand" and isinstance(ast, UnaryOpNode):
        return UnaryOpNode(ast.op, _replace_subtree_at_path(ast.operand, tail, replacement))
    if step == "args" and isinstance(ast, FunctionCallNode) and rest:
        index = rest[0]
        if not isinstance(index, int):
            raise ValueError(f"expected integer function-arg index in path {path!r}")
        inner_path = tuple(rest[1:])
        new_args = list(ast.args)
        new_args[index] = _replace_subtree_at_path(new_args[index], inner_path, replacement)
        return FunctionCallNode(ast.name, new_args)
    raise ValueError(f"cannot follow path {path!r} in {type(ast).__name__}")


def _fold_cell_map(cell_map: Mapping[str, AstNode]) -> tuple[dict[str, AstNode], int]:
    folded_map: dict[str, AstNode] = {}
    transforms = 0
    for key, ast in cell_map.items():
        folded = fold_literals_in_ast(ast)
        folded_map[key] = folded
        if folded != ast:
            transforms += 1
    return folded_map, transforms


def _cells_changed(
    before: Mapping[str, AstNode],
    after: Mapping[str, AstNode],
) -> int:
    keys = set(before) | set(after)
    return sum(1 for key in keys if before.get(key) != after.get(key))


def _formula_cells(cell_map: Mapping[str, AstNode]) -> dict[str, AstNode]:
    return {key: ast for key, ast in cell_map.items() if not _is_cse_binding_key(key)}


def _cse_bindings(cell_map: Mapping[str, AstNode]) -> dict[str, AstNode]:
    return {key: ast for key, ast in cell_map.items() if _is_cse_binding_key(key)}


def _is_cse_binding_key(key: str) -> bool:
    return key.startswith(_CSE_KEY_PREFIX)


def _group_occurrences(
    shared: Mapping[SubtreeSignature, tuple[SubtreeOccurrence, ...]],
) -> dict[SubtreeSignature, tuple[SubtreeOccurrence, ...]]:
    return {
        signature: occurrences for signature, occurrences in shared.items() if len(occurrences) >= 2
    }


def _is_compound_subtree_root(node: AstNode) -> bool:
    return isinstance(node, (FunctionCallNode, BinaryOpNode, UnaryOpNode))
