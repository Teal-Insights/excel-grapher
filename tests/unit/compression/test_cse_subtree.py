"""Unit tests for CSE subtree analysis and cost gates."""

from __future__ import annotations

from excel_grapher.compression.cse import (
    CseCandidate,
    CseConfig,
    CseGateRejection,
    find_shared_subtrees,
    passes_cse_gates,
    subtree_node_count,
    subtree_signature,
)
from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    CellRefNode,
    NumberNode,
    SubexpressionRefNode,
)

from .conftest import parse_formula


def _shared_sum_times_three() -> dict[str, AstNode]:
    shared = parse_formula("=Sheet1!B1+Sheet1!C1")
    return {
        "Sheet1!A1": BinaryOpNode("*", shared, NumberNode(2.0)),
        "Sheet1!A2": BinaryOpNode("*", shared, NumberNode(3.0)),
        "Sheet1!A3": BinaryOpNode("+", shared, NumberNode(10.0)),
    }


def _shared_sum_times_two() -> dict[str, AstNode]:
    shared = parse_formula("=Sheet1!B1+Sheet1!C1")
    return {
        "Sheet1!D1": BinaryOpNode("*", shared, NumberNode(2.0)),
        "Sheet1!D2": BinaryOpNode("*", shared, NumberNode(3.0)),
    }


def test_subtree_node_count_for_binary_sum() -> None:
    ast = parse_formula("=Sheet1!B1+Sheet1!C1")
    assert subtree_node_count(ast) == 3


def test_subtree_node_count_for_cell_ref_is_one() -> None:
    assert subtree_node_count(CellRefNode("Sheet1!B1")) == 1


def test_subtree_signature_matches_structurally() -> None:
    left = parse_formula("=Sheet1!B1+Sheet1!C1")
    right = parse_formula("=Sheet1!C1+Sheet1!B1")
    assert subtree_signature(left) != subtree_signature(right)


def test_subtree_signature_includes_cse_ref() -> None:
    ref = SubexpressionRefNode("_cse!0")
    assert subtree_signature(ref) == ("CSE", "_cse!0")


def test_find_shared_subtrees_three_cell_pattern() -> None:
    cell_map = _shared_sum_times_three()
    groups = find_shared_subtrees(cell_map)
    shared = parse_formula("=Sheet1!B1+Sheet1!C1")
    signature = subtree_signature(shared)
    assert signature in groups
    assert len(groups[signature]) == 3


def test_find_shared_subtrees_two_copy_not_grouped_for_gates() -> None:
    cell_map = _shared_sum_times_two()
    candidate = CseCandidate.from_cell_map(cell_map)[0]
    result = passes_cse_gates(candidate, CseConfig())
    assert not result.passes
    assert result.rejection is CseGateRejection.INSUFFICIENT_OCCURRENCES


def test_cse_gates_require_min_subtree_nodes() -> None:
    ref = CellRefNode("Sheet1!B1")
    candidate = CseCandidate(
        signature=subtree_signature(ref),
        ast=ref,
        occurrences=(
            ("Sheet1!A1", ()),
            ("Sheet1!A2", ()),
            ("Sheet1!A3", ()),
        ),
    )
    result = passes_cse_gates(candidate, CseConfig())
    assert not result.passes
    assert result.rejection is CseGateRejection.SUBTREE_TOO_SMALL


def test_cse_gates_require_net_ast_savings() -> None:
    tiny = BinaryOpNode("+", NumberNode(1.0), NumberNode(2.0))
    candidate = CseCandidate(
        signature=subtree_signature(tiny),
        ast=tiny,
        occurrences=(
            ("Sheet1!A1", ()),
            ("Sheet1!A2", ()),
            ("Sheet1!A3", ()),
        ),
    )
    result = passes_cse_gates(candidate, CseConfig(min_net_ast_savings=5))
    assert not result.passes
    assert result.rejection is CseGateRejection.INSUFFICIENT_NET_SAVINGS


def test_cse_gates_pass_three_cell_sum_pattern() -> None:
    candidate = CseCandidate.from_cell_map(_shared_sum_times_three())[0]
    result = passes_cse_gates(candidate, CseConfig())
    assert result.passes
    assert result.rejection is None


def test_cse_candidate_from_cell_map_picks_shared_sum() -> None:
    candidates = CseCandidate.from_cell_map(_shared_sum_times_three())
    matching = [item for item in candidates if passes_cse_gates(item, CseConfig()).passes]
    assert len(matching) == 1
    assert subtree_node_count(matching[0].ast) == 3
