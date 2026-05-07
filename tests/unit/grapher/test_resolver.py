from __future__ import annotations

import pytest

from excel_grapher.core.formula_ast import FunctionCallNode
from excel_grapher.core.formula_ast import (
    parse as parse_formula_ast,
)
from excel_grapher.grapher.resolver import _eval_indirect_formula_to_range


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ('=INDIRECT("A1:A3")', ("Sheet1", "A1", "A3")),
        ('=INDIRECT("Sheet1!A1:A3")', ("Sheet1", "A1", "A3")),
        ('=INDIRECT("Sheet1!A1:Sheet1!A3")', ("Sheet1", "A1", "A3")),
        ("=INDIRECT(\"'My Sheet'!A1:'My Sheet'!A3\")", ("My Sheet", "A1", "A3")),
    ],
)
def test_eval_indirect_formula_to_range_supports_single_and_qualified_endpoints(
    formula: str, expected: tuple[str, str, str]
) -> None:
    node = parse_formula_ast(formula)
    assert isinstance(node, FunctionCallNode)

    result = _eval_indirect_formula_to_range(
        node,
        get_cell_value=lambda _addr: None,
        bounds={"Sheet1": (100, 26), "My Sheet": (100, 26)},
    )

    assert result == expected


def test_eval_indirect_formula_to_range_rejects_mixed_sheet_endpoints() -> None:
    node = parse_formula_ast('=INDIRECT("Sheet1!A1:Sheet2!A3")')
    assert isinstance(node, FunctionCallNode)

    result = _eval_indirect_formula_to_range(
        node,
        get_cell_value=lambda _addr: None,
        bounds={"Sheet1": (100, 26), "Sheet2": (100, 26)},
    )

    assert result is None


def test_eval_indirect_formula_to_range_normalizes_reversed_range_order() -> None:
    node = parse_formula_ast('=INDIRECT("A3:A1")')
    assert isinstance(node, FunctionCallNode)

    result = _eval_indirect_formula_to_range(
        node,
        get_cell_value=lambda _addr: None,
        bounds={"Sheet1": (100, 26)},
    )

    assert result == ("Sheet1", "A1", "A3")
