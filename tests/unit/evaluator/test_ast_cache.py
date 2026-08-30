"""Tests for FormulaEvaluator AST parse cache."""

from __future__ import annotations

from unittest.mock import patch

import pytest

import excel_grapher.evaluator.evaluator as evaluator_module
from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.core.formula_ast import (
    BinaryOpNode,
    CellRefNode,
    RelativeAxis,
    bind_axes,
    parse,
    parse_preserving_axes,
)
from excel_grapher.evaluator import parser as evaluator_parser
from excel_grapher.evaluator.ast_cache import DEFAULT_AST_CACHE_MAXSIZE, AstCache
from excel_grapher.evaluator.errors import ParseError
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.grapher.node import make_cell_node
from excel_grapher.grapher.preparsed_formulas import warm_preparsed_formulas


def _make_node(
    address: str,
    formula: str | None,
    value: object,
    *,
    normalized_formula: str | None = None,
) -> Node:
    sheet, coord = parse_address(address)
    col = "".join(c for c in coord if c.isalpha())
    row = int("".join(c for c in coord if c.isdigit()))
    node = Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=None,
        value=value,
        is_leaf=formula is None,
    )
    # Leave `formula_ast` unset so these tests cover the evaluator's
    # string-keyed parse cache rather than stored-AST evaluation.
    text = normalized_formula if normalized_formula is not None else formula
    if text is not None:
        node._unparseable_formula = text
    return node


def _make_graph(*nodes: Node) -> DependencyGraph:
    graph = DependencyGraph()
    for node in nodes:
        graph.add_node(node)
    return graph


def test_second_eval_after_invalidation_reuses_ast() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
    )
    graph.add_edge("S!B1", "S!A1")

    parse_calls = 0
    original_parse = evaluator_module.parse

    def counting_parse(formula: str):
        nonlocal parse_calls
        parse_calls += 1
        return original_parse(formula)

    with (
        FormulaEvaluator(graph, auto_detect_changes=True) as ev,
        patch.object(evaluator_module, "parse", counting_parse),
    ):
        ev.evaluate(["S!B1"])
        assert parse_calls == 1

        graph.set_node_value("S!A1", 5)

        ev.evaluate(["S!B1"])
        assert parse_calls == 1
        assert ev.evaluate(["S!B1"])["S!B1"] == 10.0


def test_distinct_cells_same_normalized_formula_share_ast() -> None:
    shared = "=S!A1*2"
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None, normalized_formula=shared),
        _make_node("S!B2", "=S!A2*2", None, normalized_formula=shared),
    )
    graph.add_edge("S!B1", "S!A1")
    graph.add_edge("S!B2", "S!A1")

    parse_calls = 0
    original_parse = evaluator_module.parse

    def counting_parse(formula: str):
        nonlocal parse_calls
        parse_calls += 1
        return original_parse(formula)

    with FormulaEvaluator(graph) as ev, patch.object(evaluator_module, "parse", counting_parse):
        ev.evaluate(["S!B1", "S!B2"])
        assert parse_calls == 1


def test_different_formulas_distinct_ast_entries() -> None:
    graph = _make_graph(
        _make_node("S!A1", "=1+1", None),
        _make_node("S!A2", "=2+2", None),
    )

    parse_calls = 0
    original_parse = evaluator_module.parse

    def counting_parse(formula: str):
        nonlocal parse_calls
        parse_calls += 1
        return original_parse(formula)

    with FormulaEvaluator(graph) as ev, patch.object(evaluator_module, "parse", counting_parse):
        ev.evaluate(["S!A1", "S!A2"])
        assert parse_calls == 2


def test_ast_cache_is_bounded() -> None:
    graph = _make_graph(
        _make_node("S!A1", "=1", None),
        _make_node("S!A2", "=2", None),
        _make_node("S!A3", "=3", None),
    )

    with FormulaEvaluator(graph, ast_cache_maxsize=2) as ev:
        ev.evaluate(["S!A1", "S!A2", "S!A3"])
        assert len(ev._ast_cache) == 2

        info_before = ev.ast_cache_info()
        assert info_before.currsize == 2
        assert info_before.maxsize == 2

        ev._cache.pop("S!A1")
        ev.evaluate(["S!A1"])
        info_after = ev.ast_cache_info()
        assert info_after.misses == info_before.misses + 1


def test_parse_error_not_cached() -> None:
    graph = _make_graph(_make_node("S!A1", "=1+", None))

    with FormulaEvaluator(graph) as ev:
        with pytest.raises(ParseError):
            ev.evaluate(["S!A1"])
        assert len(ev._ast_cache) == 0

        with pytest.raises(ParseError):
            ev.evaluate(["S!A1"])
        info = ev.ast_cache_info()
        assert info.misses == 2
        assert info.hits == 0


def test_clear_caches_clears_ast_and_value_caches() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!B1", "=S!A1+1", None),
    )

    with FormulaEvaluator(graph) as ev:
        ev.evaluate(["S!B1"])
        assert ev._cache
        assert len(ev._ast_cache) == 1

        ev.clear_caches()
        assert ev._cache == {}
        assert len(ev._ast_cache) == 0


def test_ast_cache_default_maxsize() -> None:
    cache = AstCache()
    assert cache.maxsize == DEFAULT_AST_CACHE_MAXSIZE


def test_ast_cache_info_tracks_hits_and_misses() -> None:
    cache = AstCache(maxsize=8)
    cache.get("=1", parse_fn=evaluator_parser.parse)
    cache.get("=1", parse_fn=evaluator_parser.parse)
    cache.get("=2", parse_fn=evaluator_parser.parse)

    info = cache.cache_info()
    assert info.hits == 1
    assert info.misses == 2
    assert info.currsize == 2


def test_ast_cache_seed_does_not_affect_hit_miss_stats() -> None:
    cache = AstCache(maxsize=8)
    cache.seed({"=1": evaluator_parser.parse("=1")})

    info = cache.cache_info()
    assert info.hits == 0
    assert info.misses == 0
    assert info.currsize == 1


def test_ast_cache_seed_skips_existing_keys() -> None:
    cache = AstCache(maxsize=8)
    original = cache.get("=1", parse_fn=evaluator_parser.parse)
    info_before = cache.cache_info()

    replacement = evaluator_parser.parse("=2")
    cache.seed({"=1": replacement, "=2": replacement})

    assert cache.get("=1", parse_fn=evaluator_parser.parse) is original
    info_after = cache.cache_info()
    assert info_after.hits == info_before.hits + 1
    assert info_after.misses == info_before.misses
    assert info_after.currsize == 2


def test_ast_cache_seed_respects_maxsize() -> None:
    cache = AstCache(maxsize=2)
    cache.seed(
        {
            "=1": evaluator_parser.parse("=1"),
            "=2": evaluator_parser.parse("=2"),
            "=3": evaluator_parser.parse("=3"),
        }
    )

    assert len(cache) == 2
    assert "=1" not in cache._cache
    assert "=2" in cache._cache
    assert "=3" in cache._cache


def _drop_formula_ast_keep_normalized(graph: DependencyGraph, key: str) -> None:
    """Force the string-keyed fallback while keeping the absolute A1 spelling."""
    node = graph._get_internal_node(key)
    assert node is not None
    nf = node.normalized_formula
    assert isinstance(nf, str) and nf.strip()
    node.formula_ast = None
    node._unparseable_formula = nf


def test_relative_and_absolute_same_a1_spelling_do_not_poison_fallback() -> None:
    """B1 `=A1*2` and C1 `=$A$1*2` share `normalized_formula`; C1 fallback stays 20."""
    graph = DependencyGraph()
    graph.add_node(make_cell_node("S", "A", 1, value=10, is_leaf=True))
    graph.add_node(
        make_cell_node(
            "S",
            "B",
            1,
            is_leaf=False,
            formula_ast=parse_preserving_axes("=A1*2", anchor="S!B1"),
        )
    )
    graph.add_node(
        make_cell_node(
            "S",
            "C",
            1,
            is_leaf=False,
            formula_ast=parse_preserving_axes("=$A$1*2", anchor="S!C1"),
        )
    )
    b1 = graph.get_node("S!B1")
    c1 = graph.get_node("S!C1")
    assert b1 is not None and c1 is not None
    assert b1.normalized_formula == c1.normalized_formula == "=S!A1*2"

    _drop_formula_ast_keep_normalized(graph, "S!C1")

    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate("S!C1") == 20.0
        assert ev.evaluate("S!B1") == 20.0


def test_same_relative_excel_text_at_different_hosts_do_not_poison_fallback() -> None:
    """`=A1*2` at B1 vs C5 share absolute A1 text; C5 fallback still means A1."""
    graph = DependencyGraph()
    graph.add_node(make_cell_node("S", "A", 1, value=10, is_leaf=True))
    graph.add_node(make_cell_node("S", "B", 5, value=99, is_leaf=True))
    graph.add_node(
        make_cell_node(
            "S",
            "B",
            1,
            is_leaf=False,
            formula_ast=parse_preserving_axes("=A1*2", anchor="S!B1"),
        )
    )
    graph.add_node(
        make_cell_node(
            "S",
            "C",
            5,
            is_leaf=False,
            formula_ast=parse_preserving_axes("=A1*2", anchor="S!C5"),
        )
    )
    b1 = graph.get_node("S!B1")
    c5 = graph.get_node("S!C5")
    assert b1 is not None and c5 is not None
    assert b1.normalized_formula == c5.normalized_formula == "=S!A1*2"

    _drop_formula_ast_keep_normalized(graph, "S!C5")

    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate("S!C5") == 20.0
        assert ev.evaluate("S!B1") == 20.0


def test_preparsed_formulas_overlay_does_not_poison_absolute_fallback() -> None:
    """`warm_preparsed_formulas` first-wins must bind relatives before the overlay."""
    graph = DependencyGraph()
    graph.add_node(make_cell_node("S", "A", 1, value=10, is_leaf=True))
    graph.add_node(
        make_cell_node(
            "S",
            "B",
            1,
            is_leaf=False,
            formula_ast=parse_preserving_axes("=A1*2", anchor="S!B1"),
        )
    )
    graph.add_node(
        make_cell_node(
            "S",
            "C",
            1,
            is_leaf=False,
            formula_ast=parse_preserving_axes("=$A$1*2", anchor="S!C1"),
        )
    )
    overlay = warm_preparsed_formulas(graph)
    graph.preparsed_formulas = overlay
    shared = graph.get_node("S!B1")
    assert shared is not None
    nf = shared.normalized_formula
    assert nf is not None
    cached = overlay[nf]
    assert cached == bind_axes(parse_preserving_axes("=A1*2", anchor="S!B1"), "S!B1")
    assert cached == parse(nf)

    _drop_formula_ast_keep_normalized(graph, "S!B1")
    _drop_formula_ast_keep_normalized(graph, "S!C1")

    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate("S!C1") == 20.0
        assert ev.evaluate("S!B1") == 20.0


def test_seeded_string_cache_stores_absolute_bound_trees() -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("S", "A", 1, value=10, is_leaf=True))
    rel = parse_preserving_axes("=A1*2", anchor="S!B1")
    assert isinstance(rel, BinaryOpNode)
    assert isinstance(rel.left, CellRefNode)
    assert isinstance(rel.left.ref.col, RelativeAxis)
    graph.add_node(make_cell_node("S", "B", 1, is_leaf=False, formula_ast=rel))

    with FormulaEvaluator(graph) as ev:
        cached = ev._ast_cache._cache["=S!A1*2"]
        assert cached == bind_axes(rel, "S!B1")
        assert cached == parse("=S!A1*2")
