"""Unit tests for pass-through compression rule."""

from __future__ import annotations

from excel_grapher.compression.parity import assert_compression_parity
from excel_grapher.compression.pass_through import (
    apply_pass_through,
    identify_pass_through_cells,
    replace_pass_through_refs,
    resolve_pass_through_chains,
    singleton_cell_ref_target,
)
from excel_grapher.compression.rules import empty_compression_stats
from excel_grapher.core.address_keys import normalize_key
from excel_grapher.core.formula_ast import CellRefNode

from .conftest import parse_formula


def test_singleton_cell_ref_target_accepts_unary_plus() -> None:
    ast = parse_formula("=+Sheet1!B1")
    assert singleton_cell_ref_target(ast) == "Sheet1!B1"


def test_singleton_cell_ref_target_rejects_non_refs() -> None:
    assert singleton_cell_ref_target(parse_formula("=Sheet1!B1+1")) is None
    assert singleton_cell_ref_target(parse_formula("=2+3")) is None


def test_identify_pass_through_cells_single_hop() -> None:
    ast_map = {
        "Sheet1!A1": CellRefNode("Sheet1!B1"),
        "Sheet1!C1": parse_formula("=Sheet1!A1+10"),
    }
    assert identify_pass_through_cells(ast_map) == {"Sheet1!A1": "Sheet1!B1"}


def test_resolve_pass_through_chains_transitive() -> None:
    mapping = {
        "Sheet1!A1": "Sheet1!B1",
        "Sheet1!B1": "Sheet1!C1",
    }
    assert resolve_pass_through_chains(mapping) == {
        "Sheet1!A1": "Sheet1!C1",
        "Sheet1!B1": "Sheet1!C1",
    }


def test_resolve_pass_through_chains_skips_cycles() -> None:
    mapping = {
        "Sheet1!A1": "Sheet1!B1",
        "Sheet1!B1": "Sheet1!A1",
    }
    resolved = resolve_pass_through_chains(mapping)
    assert resolved["Sheet1!A1"] in {"Sheet1!A1", "Sheet1!B1"}
    assert resolved["Sheet1!B1"] in {"Sheet1!A1", "Sheet1!B1"}


def test_replace_pass_through_refs_rewrites_dependents() -> None:
    ast = parse_formula("=Sheet1!A1+10")
    rewritten = replace_pass_through_refs(ast, {"Sheet1!A1": "Sheet1!B1"})
    assert rewritten == parse_formula("=Sheet1!B1+10")


def test_apply_pass_through_rewrites_multiple_dependents() -> None:
    ast_map = {
        "Sheet1!A1": CellRefNode("Sheet1!B1"),
        "Sheet1!C1": parse_formula("=Sheet1!A1+10"),
        "Sheet1!D1": parse_formula("=Sheet1!A1*2"),
    }
    result = apply_pass_through(ast_map)
    assert result["Sheet1!A1"] == CellRefNode("Sheet1!B1")
    assert result["Sheet1!C1"] == parse_formula("=Sheet1!B1+10")
    assert result["Sheet1!D1"] == parse_formula("=Sheet1!B1*2")


def test_apply_pass_through_records_stats() -> None:
    ast_map = {
        "Sheet1!A1": CellRefNode("Sheet1!B1"),
        "Sheet1!C1": parse_formula("=Sheet1!A1+10"),
    }
    stats = empty_compression_stats()
    apply_pass_through(ast_map, stats=stats)
    contrib = stats.contribution_for("pass_through")
    assert contrib.in_place_transforms == 1
    assert contrib.cells_affected == 1


def test_apply_pass_through_chain_parity() -> None:
    input_values = {"Sheet1!C1": 100}
    original = {
        "Sheet1!B1": CellRefNode("Sheet1!C1"),
        "Sheet1!A1": CellRefNode("Sheet1!B1"),
        "Sheet1!D1": parse_formula("=Sheet1!A1*2"),
    }
    compressed = apply_pass_through(original)
    assert_compression_parity(original, compressed, input_values=input_values)


def test_apply_pass_through_parity_multi_dependent() -> None:
    input_values = {"Sheet1!B1": 42}
    original = {
        "Sheet1!A1": CellRefNode("Sheet1!B1"),
        "Sheet1!C1": parse_formula("=Sheet1!A1+10"),
        "Sheet1!D1": parse_formula("=Sheet1!A1*2"),
    }
    compressed = apply_pass_through(original)
    assert normalize_key("Sheet1!C1") in compressed
    assert_compression_parity(original, compressed, input_values=input_values)
