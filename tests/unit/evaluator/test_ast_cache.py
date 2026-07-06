"""Tests for FormulaEvaluator AST parse cache."""

from __future__ import annotations

from unittest.mock import patch

import pytest

import excel_grapher.evaluator.evaluator as evaluator_module
from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.evaluator import parser as evaluator_parser
from excel_grapher.evaluator.ast_cache import DEFAULT_AST_CACHE_MAXSIZE, AstCache
from excel_grapher.evaluator.errors import ParseError
from excel_grapher.evaluator.evaluator import FormulaEvaluator


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
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=normalized_formula if normalized_formula is not None else formula,
        value=value,
        is_leaf=formula is None,
    )


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
        assert len(ev._ast_cache) == 2  # noqa: SLF001

        info_before = ev.ast_cache_info()
        assert info_before.currsize == 2
        assert info_before.maxsize == 2

        ev._cache.pop("S!A1")  # noqa: SLF001 — force re-evaluation past value cache
        ev.evaluate(["S!A1"])
        info_after = ev.ast_cache_info()
        assert info_after.misses == info_before.misses + 1


def test_parse_error_not_cached() -> None:
    graph = _make_graph(_make_node("S!A1", "=1+", None))

    with FormulaEvaluator(graph) as ev:
        with pytest.raises(ParseError):
            ev.evaluate(["S!A1"])
        assert len(ev._ast_cache) == 0  # noqa: SLF001

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
        assert ev._cache  # noqa: SLF001
        assert len(ev._ast_cache) == 1  # noqa: SLF001

        ev.clear_caches()
        assert ev._cache == {}  # noqa: SLF001
        assert len(ev._ast_cache) == 0  # noqa: SLF001


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
    assert "=1" not in cache._cache  # noqa: SLF001
    assert "=2" in cache._cache  # noqa: SLF001
    assert "=3" in cache._cache  # noqa: SLF001
