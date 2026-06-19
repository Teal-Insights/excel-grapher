"""Integration tests for ``_xlfn`` / ``_xludf`` workbook fixtures."""

from __future__ import annotations

import re
from collections.abc import Iterator
from pathlib import Path

import pytest
from fastpyxl import load_workbook

from excel_grapher import FormulaEvaluator
from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    FunctionCallNode,
    UnaryOpNode,
    parse,
)
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher import DependencyGraph, Node
from excel_grapher.series_bindings.workflow import validate_bindings_workbook
from tests.integration.user_flows.bindings_accuracy import (
    BindingsAccuracyCase,
    assert_bindings_validate,
)
from tests.integration.utils.rewrite_prefixed_workbook import (
    rewrite_formula_to_xlfn,
    rewrite_formula_to_xludf,
)

EXAMPLES = Path(__file__).resolve().parents[3] / "examples" / "micro_workbooks"

XLFN_WORKBOOK = EXAMPLES / "advanced_formula_workbook_xlfn.xlsx"
XLFN_BINDINGS = EXAMPLES / "advanced_formula_workbook_xlfn.bindings"

_LEGACY_RUNTIME_PATTERN = re.compile(r"xl__xl(fn|udf)_[a-z0-9_]+")
_XLFN_IN_FORMULA = re.compile(r"_xlfn\.", re.IGNORECASE)


def _skip_if_xlfn_fixture_missing() -> None:
    if not XLFN_WORKBOOK.is_file():
        pytest.skip(f"Workbook fixture missing: {XLFN_WORKBOOK}")
    if not XLFN_BINDINGS.is_dir():
        pytest.skip(f"Bindings directory missing: {XLFN_BINDINGS}")


def _iter_function_calls(node: AstNode) -> Iterator[FunctionCallNode]:
    if isinstance(node, FunctionCallNode):
        yield node
        for arg in node.args:
            yield from _iter_function_calls(arg)
    elif isinstance(node, BinaryOpNode):
        yield from _iter_function_calls(node.left)
        yield from _iter_function_calls(node.right)
    elif isinstance(node, UnaryOpNode):
        yield from _iter_function_calls(node.operand)


def test_xlfn_fixture_bindings_validate() -> None:
    """Committed ``_xlfn`` workbook fixture validates with its binding shards."""
    _skip_if_xlfn_fixture_missing()
    assert_bindings_validate(
        BindingsAccuracyCase(
            name="advanced_formula_workbook_xlfn",
            workbook=XLFN_WORKBOOK,
            bindings_path=XLFN_BINDINGS,
        )
    )


def test_xlfn_fixture_workbook_contains_xlfn_formulas() -> None:
    """Fixture on disk uses ``_xlfn.`` spelling for allowlisted built-ins."""
    _skip_if_xlfn_fixture_missing()
    wb = load_workbook(XLFN_WORKBOOK, data_only=False)
    found = 0
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and _XLFN_IN_FORMULA.search(cell.value):
                        found += 1
    finally:
        wb.close()
    assert found >= 4, "expected multiple _xlfn formulas in xlfn fixture"


def test_xlfn_fixture_formulas_parse_to_canonical_function_names() -> None:
    """Graph formulas with ``_xlfn.`` parse to canonical AST function names."""
    _skip_if_xlfn_fixture_missing()
    graph = validate_bindings_workbook(XLFN_WORKBOOK, XLFN_BINDINGS)["graph"]
    checked = 0
    for key in graph:
        node = graph.get_node(key)
        if node is None:
            continue
        formula = node.normalized_formula
        if not formula or not _XLFN_IN_FORMULA.search(formula):
            continue
        ast = parse(formula)
        for call in _iter_function_calls(ast):
            assert "." not in call.name
            assert not call.name.startswith("_XL")
        checked += 1
    assert checked >= 4, "expected multiple _xlfn formulas in dependency graph"


def test_xlfn_fixture_modular_codegen_has_no_legacy_runtime_symbols() -> None:
    """Modular export of ``_xlfn`` formulas must not emit ``xl__xlfn_*`` / ``xl__xludf_*``."""
    graph = DependencyGraph()
    graph.add_node(
        Node(
            sheet="S",
            column="A",
            row=1,
            formula=None,
            normalized_formula=None,
            value=1,
            is_leaf=True,
        )
    )
    graph.add_node(
        Node(
            sheet="S",
            column="B",
            row=1,
            formula='=_xlfn.IFNA(_xlfn.XLOOKUP(1,S!A1:S!A1,S!A1:S!A1),"x")',
            normalized_formula='=_xlfn.IFNA(_xlfn.XLOOKUP(1,S!A1:S!A1,S!A1:S!A1),"x")',
            value=None,
            is_leaf=False,
        )
    )
    combined = "\n".join(CodeGenerator(graph).generate_modules(["S!B1"]).values())
    assert _LEGACY_RUNTIME_PATTERN.search(combined) is None
    assert "xl_xlookup" in combined or "xl_ifna" in combined


@pytest.mark.parametrize(
    ("bare", "prefixed"),
    [
        ('=NUMBERVALUE("1,234.56", ".", ",")', '=_xlfn.NUMBERVALUE("1,234.56", ".", ",")'),
        ("=IFNA(S!A1, 9)", "=_xlfn.IFNA(S!A1, 9)"),
    ],
)
def test_xlfn_rewrite_helper_matches_bare_and_prefixed_evaluation(bare: str, prefixed: str) -> None:
    """Prefix normalization is not the root cause when bare and ``_xlfn`` disagree."""
    from excel_grapher.core.address_keys import parse_address
    from excel_grapher.evaluator.types import XlError

    def _node(address: str, formula: str | None) -> Node:
        sheet, coord = parse_address(address)
        col = "".join(c for c in coord if c.isalpha())
        row = int("".join(c for c in coord if c.isdigit()))
        return Node(
            sheet=sheet,
            column=col,
            row=row,
            formula=formula,
            normalized_formula=formula,
            value=XlError.NA if address == "S!A1" else None,
            is_leaf=formula is None,
        )

    graph = DependencyGraph()
    graph.add_node(_node("S!A1", None))
    graph.add_node(_node("S!B1", bare))
    graph.add_node(_node("S!B2", prefixed))
    with FormulaEvaluator(graph) as ev:
        bare_result = ev.evaluate(["S!B1"])["S!B1"]
        prefixed_result = ev.evaluate(["S!B2"])["S!B2"]
    assert bare_result == prefixed_result


def test_rewrite_helpers_produce_distinct_prefix_spellings() -> None:
    formula = "=IFNA(XLOOKUP(1,A1:A3,B1:B3),0)"
    assert "_xlfn." in rewrite_formula_to_xlfn(formula).lower()
    assert "_xludf." in rewrite_formula_to_xludf(formula).lower()
