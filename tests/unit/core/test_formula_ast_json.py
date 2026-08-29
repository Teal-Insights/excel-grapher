"""JSON codec round-trip for formula AST nodes."""

from __future__ import annotations

import pytest

from excel_grapher.core.formula_ast import (
    EmptyArgNode,
    ErrorNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    WholeColumnNode,
    WholeRowNode,
    parse,
)
from excel_grapher.core.formula_ast_json import (
    ast_from_json,
    ast_identity_key,
    ast_to_json,
    formula_identity_digest,
)
from excel_grapher.core.types import XlError


def test_ast_json_round_trips_parse_trees() -> None:
    formulas = (
        "=1.5",
        '="hi"',
        "=TRUE",
        "=#DIV/0!",
        "=Sheet1!A1",
        "=Sheet1!A1:B2",
        "=Data!A:A",
        "=Data!3:3",
        "=-Sheet1!A1",
        "=1%",
        "=Sheet1!A1+Sheet1!B1",
        "=SUM(Sheet1!A1,)",
        '=IF(Sheet1!A1,"x",Sheet1!B1)',
    )
    for formula in formulas:
        original = parse(formula)
        restored = ast_from_json(ast_to_json(original))
        assert restored == original


def test_ast_json_round_trips_empty_arg_and_error_nodes() -> None:
    tree = FunctionCallNode(
        "INDEX",
        [
            RangeNode("Sheet1!A1", "Sheet1!B2"),
            EmptyArgNode(),
            NumberNode(1.0),
        ],
    )
    assert ast_from_json(ast_to_json(tree)) == tree
    assert ast_from_json(ast_to_json(ErrorNode(XlError.NA))) == ErrorNode(XlError.NA)


def test_ast_json_round_trips_whole_column_and_row() -> None:
    col = WholeColumnNode(sheet="Data", column="C")
    row = WholeRowNode(sheet="Data", row=4)
    assert ast_from_json(ast_to_json(col)) == col
    assert ast_from_json(ast_to_json(row)) == row


def test_ast_from_json_rejects_unknown_tag() -> None:
    with pytest.raises(TypeError, match="unknown"):
        ast_from_json({"t": "not-a-node"})


def test_ast_from_json_rejects_non_object() -> None:
    with pytest.raises(TypeError):
        ast_from_json(["cell"])
    with pytest.raises(TypeError):
        ast_from_json(None)


def test_formula_identity_digest_uses_ast_when_present() -> None:
    ast = parse("=Sheet1!A1+1")
    stale = "=Sheet1!A1+999"
    digest = formula_identity_digest(formula=stale, formula_ast=ast)
    assert digest == formula_identity_digest(formula="ignored", formula_ast=ast)
    assert digest != formula_identity_digest(formula=stale, formula_ast=None)
    assert ast_identity_key(ast) != stale
