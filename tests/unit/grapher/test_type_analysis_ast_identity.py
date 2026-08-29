"""Type-analysis cache identity uses stored `formula_ast` (#544 / PR 552)."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import excel_grapher.grapher.dynamic_refs as dynamic_refs_mod
from excel_grapher.core.cell_types import CellKind, CellType, EnumDomain
from excel_grapher.core.formula_ast import parse, parse_preserving_axes
from excel_grapher.core.formula_ast_json import formula_identity_digest
from excel_grapher.grapher.dynamic_refs import (
    DynamicRefLimits,
    expand_leaf_env_to_argument_env,
)
from excel_grapher.grapher.type_analysis_cache import TypeAnalysisCache

_FORMULAS = {
    "Sheet1!C1": "=Sheet1!A1+Sheet1!B1",
}
_REFS = {
    "=Sheet1!A1+Sheet1!B1": {"Sheet1!A1", "Sheet1!B1"},
}
_LEAF_ENV = {
    "Sheet1!A1": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1, 2}))),
    "Sheet1!B1": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({10, 20}))),
}
_LIMITS = DynamicRefLimits()
_WB_SHA = "ast_identity_workbook_sha"


def _get_cell_formula(addr: str) -> str | None:
    return _FORMULAS.get(addr)


def _get_refs(formula: str, sheet: str) -> set[str]:
    return _REFS.get(formula, set())


def test_relative_and_absolute_asts_sharing_a1_do_not_collide(tmp_path: Path) -> None:
    relative = parse_preserving_axes("=A1+B1", anchor="Sheet1!C1")
    absolute = parse("=Sheet1!A1+Sheet1!B1")
    formula = "=Sheet1!A1+Sheet1!B1"
    assert formula_identity_digest(formula=formula, formula_ast=relative) != (
        formula_identity_digest(formula=formula, formula_ast=absolute)
    )

    cache = TypeAnalysisCache.open(tmp_path / "ast-identity.sqlite3")
    try:
        env_rel = expand_leaf_env_to_argument_env(
            {"Sheet1!C1"},
            _get_cell_formula,
            _get_refs,
            _LEAF_ENV,
            _LIMITS,
            type_analysis_cache=cache,
            workbook_sha256=_WB_SHA,
            get_cell_ast=lambda addr: relative if addr == "Sheet1!C1" else None,
        )
        cache.flush()
        hits_after_relative = cache.stats.hits

        env_abs = expand_leaf_env_to_argument_env(
            {"Sheet1!C1"},
            _get_cell_formula,
            _get_refs,
            _LEAF_ENV,
            _LIMITS,
            type_analysis_cache=cache,
            workbook_sha256=_WB_SHA,
            get_cell_ast=lambda addr: absolute if addr == "Sheet1!C1" else None,
        )
        assert env_rel["Sheet1!C1"] == env_abs["Sheet1!C1"]
        assert cache.stats.hits == hits_after_relative
    finally:
        cache.close()


def test_stored_ast_cache_hit_does_not_reparse(tmp_path: Path) -> None:
    stored = parse_preserving_axes("=A1+B1", anchor="Sheet1!C1")
    cache = TypeAnalysisCache.open(tmp_path / "ast-hit.sqlite3")
    try:
        expand_leaf_env_to_argument_env(
            {"Sheet1!C1"},
            _get_cell_formula,
            _get_refs,
            _LEAF_ENV,
            _LIMITS,
            type_analysis_cache=cache,
            workbook_sha256=_WB_SHA,
            get_cell_ast=lambda addr: stored if addr == "Sheet1!C1" else None,
        )
        cache.flush()

        parse_calls = 0
        original_parse = dynamic_refs_mod.parse_ast

        def counting_parse(formula: str):
            nonlocal parse_calls
            parse_calls += 1
            return original_parse(formula)

        with patch.object(dynamic_refs_mod, "parse_ast", counting_parse):
            env = expand_leaf_env_to_argument_env(
                {"Sheet1!C1"},
                _get_cell_formula,
                _get_refs,
                _LEAF_ENV,
                _LIMITS,
                type_analysis_cache=cache,
                workbook_sha256=_WB_SHA,
                get_cell_ast=lambda addr: stored if addr == "Sheet1!C1" else None,
            )
        assert env["Sheet1!C1"].kind is CellKind.NUMBER
        assert cache.stats.hits >= 1
        assert parse_calls == 0
    finally:
        cache.close()


def test_string_fallback_parses_only_after_cache_miss(tmp_path: Path) -> None:
    cache = TypeAnalysisCache.open(tmp_path / "string-fallback.sqlite3")
    original_parse = dynamic_refs_mod.parse_ast
    parse_calls = 0

    def counting_parse(formula: str):
        nonlocal parse_calls
        parse_calls += 1
        return original_parse(formula)

    try:
        with patch.object(dynamic_refs_mod, "parse_ast", counting_parse):
            expand_leaf_env_to_argument_env(
                {"Sheet1!C1"},
                _get_cell_formula,
                _get_refs,
                _LEAF_ENV,
                _LIMITS,
                type_analysis_cache=cache,
                workbook_sha256=_WB_SHA,
            )
        cache.flush()
        first_parses = parse_calls
        assert first_parses >= 1

        with patch.object(dynamic_refs_mod, "parse_ast", counting_parse):
            env = expand_leaf_env_to_argument_env(
                {"Sheet1!C1"},
                _get_cell_formula,
                _get_refs,
                _LEAF_ENV,
                _LIMITS,
                type_analysis_cache=cache,
                workbook_sha256=_WB_SHA,
            )
        assert env["Sheet1!C1"].kind is CellKind.NUMBER
        assert cache.stats.hits >= 1
        assert parse_calls == first_parses
    finally:
        cache.close()
