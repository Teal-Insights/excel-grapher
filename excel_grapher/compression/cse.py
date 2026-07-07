"""Common subexpression elimination: subtree analysis and cost gates."""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Mapping
from dataclasses import dataclass
from enum import Enum
from typing import TypeAlias

from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    FunctionCallNode,
    UnaryOpNode,
)

from .template_signature import TemplateSignature, template_signature

SubtreePath: TypeAlias = tuple[str | int, ...]
SubtreeSignature = TemplateSignature

__all__ = [
    "CseCandidate",
    "CseConfig",
    "CseGateRejection",
    "CseGateResult",
    "SubtreeOccurrence",
    "SubtreePath",
    "SubtreeSignature",
    "enumerate_subtrees",
    "find_shared_subtrees",
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
        """Build hoist candidates from shared-subtree groups in `cell_map`."""
        candidates: list[CseCandidate] = []
        for signature, occurrences in _group_occurrences(find_shared_subtrees(cell_map)).items():
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


def _group_occurrences(
    shared: Mapping[SubtreeSignature, tuple[SubtreeOccurrence, ...]],
) -> dict[SubtreeSignature, tuple[SubtreeOccurrence, ...]]:
    return {
        signature: occurrences for signature, occurrences in shared.items() if len(occurrences) >= 2
    }


def _is_compound_subtree_root(node: AstNode) -> bool:
    return isinstance(node, (FunctionCallNode, BinaryOpNode, UnaryOpNode))
