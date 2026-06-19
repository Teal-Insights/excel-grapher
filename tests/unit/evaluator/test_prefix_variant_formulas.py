"""Unit tests for evaluator parity across bare, ``_xlfn.``, and ``_xludf.`` spellings."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher import DependencyGraph, Node, create_dependency_graph
from excel_grapher.core.address_keys import parse_address
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator
from tests.unit.fixtures.prefix_variant_workbook import (
    PrefixVariant,
    write_prefix_variant_workbook,
)


def _make_node(address: str, formula: str | None, value: object) -> Node:
    sheet, coord = parse_address(address)
    column = "".join(character for character in coord if character.isalpha())
    row = int("".join(character for character in coord if character.isdigit()))
    return Node(
        sheet=sheet,
        column=column,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=value,
        is_leaf=formula is None,
    )


def _make_graph(*nodes: Node) -> DependencyGraph:
    graph = DependencyGraph()
    for node in nodes:
        graph.add_node(node)
    return graph


@pytest.mark.parametrize(
    "formula",
    [
        '=IFNA(XLOOKUP(1,Lookups!A1:Lookups!A1,Lookups!B1:Lookups!B1),"x")',
        '=_xlfn.IFNA(_xlfn.XLOOKUP(1,Lookups!A1:Lookups!A1,Lookups!B1:Lookups!B1),"x")',
        '=_xludf.IFNA(_xludf.XLOOKUP(1,Lookups!A1:Lookups!A1,Lookups!B1:Lookups!B1),"x")',
    ],
)
def test_lookup_prefix_spellings_evaluate_equivalently(formula: str) -> None:
    graph = _make_graph(
        _make_node("Lookups!A1", None, 1),
        _make_node("Lookups!B1", None, "hit"),
        _make_node("Lookups!C1", formula, None),
    )
    with FormulaEvaluator(graph) as evaluator:
        assert evaluator.evaluate("Lookups!C1") == "hit"


@pytest.mark.parametrize("variant", ["xlfn", "xludf"])
def test_prefix_variant_workbook_graph_evaluates_lookup_cell(
    tmp_path: Path,
    variant: PrefixVariant,
) -> None:
    workbook = write_prefix_variant_workbook(tmp_path / f"{variant}.xlsx", variant=variant)
    graph = create_dependency_graph(
        workbook,
        ["Lookups!C1"],
        load_values=True,
        use_cached_dynamic_refs=True,
    )
    with FormulaEvaluator(graph) as evaluator:
        assert evaluator.evaluate("Lookups!C1") == "hit"


@pytest.mark.parametrize("variant", ["xlfn", "xludf"])
def test_prefix_variant_workbook_eval_matches_codegen(
    tmp_path: Path,
    variant: PrefixVariant,
) -> None:
    workbook = write_prefix_variant_workbook(tmp_path / f"{variant}.xlsx", variant=variant)
    graph = create_dependency_graph(
        workbook,
        ["Lookups!C1"],
        load_values=True,
        use_cached_dynamic_refs=True,
    )
    assert_codegen_matches_evaluator(graph, ["Lookups!C1"])
