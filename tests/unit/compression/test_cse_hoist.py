"""Unit tests for single-round CSE hoisting."""

from __future__ import annotations

from excel_grapher.compression.cse import (
    CseCandidate,
    allocate_cse_key,
    apply_hoist,
    hoist_one_subexpression,
)
from excel_grapher.compression.expand import expand_compressed_to_cells
from excel_grapher.compression.parity import assert_compression_parity
from excel_grapher.core.formula_ast import (
    BinaryOpNode,
    NumberNode,
    SubexpressionRefNode,
)

from .conftest import parse_formula
from .test_cse_subtree import _shared_sum_times_three, _shared_sum_times_two


def test_allocate_cse_key_starts_at_zero() -> None:
    assert allocate_cse_key(()) == "_cse!0"


def test_allocate_cse_key_skips_existing() -> None:
    assert allocate_cse_key(("_cse!0", "_cse!1", "Sheet1!A1")) == "_cse!2"


def test_apply_hoist_replaces_all_occurrences() -> None:
    cell_map = _shared_sum_times_three()
    candidate = CseCandidate.from_cell_map(cell_map)[0]
    hoisted = apply_hoist(cell_map, candidate, "_cse!0")
    shared = parse_formula("=Sheet1!B1+Sheet1!C1")
    assert hoisted["_cse!0"] == shared
    assert hoisted["Sheet1!A1"] == BinaryOpNode(
        "*", SubexpressionRefNode("_cse!0"), NumberNode(2.0)
    )
    assert hoisted["Sheet1!A2"] == BinaryOpNode(
        "*", SubexpressionRefNode("_cse!0"), NumberNode(3.0)
    )
    assert hoisted["Sheet1!A3"] == BinaryOpNode(
        "+", SubexpressionRefNode("_cse!0"), NumberNode(10.0)
    )


def test_hoist_one_subexpression_returns_none_for_two_copy() -> None:
    cell_map = _shared_sum_times_two()
    compressed, result = hoist_one_subexpression(cell_map)
    assert compressed == cell_map
    assert not result.hoisted
    assert result.binding_key is None
    assert result.candidates_rejected >= 1


def test_hoist_one_subexpression_hoists_three_cell_pattern() -> None:
    cell_map = _shared_sum_times_three()
    compressed, result = hoist_one_subexpression(cell_map)
    assert result.hoisted
    assert result.binding_key == "_cse!0"
    assert result.binding_sites == 1
    assert result.redundant_evaluations_eliminated == 2
    assert result.ast_subnodes_saved == 4
    assert "_cse!0" in compressed
    assert isinstance(compressed["Sheet1!A1"], BinaryOpNode)
    assert isinstance(compressed["Sheet1!A1"].left, SubexpressionRefNode)


def test_expand_round_trip_after_hoist() -> None:
    original = _shared_sum_times_three()
    compressed, _ = hoist_one_subexpression(original)
    assert expand_compressed_to_cells(compressed) == original


def test_hoist_expand_parity_three_cell_pattern() -> None:
    original = _shared_sum_times_three()
    compressed, result = hoist_one_subexpression(original)
    assert result.hoisted
    assert_compression_parity(
        original,
        compressed,
        input_values={"Sheet1!B1": 2, "Sheet1!C1": 3},
    )


def test_hoist_one_subexpression_allocates_next_key_when_bindings_exist() -> None:
    cell_map = _shared_sum_times_three()
    first_pass, first_result = hoist_one_subexpression(cell_map)
    assert first_result.binding_key == "_cse!0"
    # Second hoist should not run on same pattern (refs, not shared AST), but allocator
    # must skip occupied keys when bindings are already present.
    assert allocate_cse_key(first_pass) == "_cse!1"
