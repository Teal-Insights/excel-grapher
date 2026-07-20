"""Unit tests for formula-group shape fingerprints."""

from __future__ import annotations

from excel_grapher.core.formula_ast import (
    AddressHoleNode,
    AddressLeafKind,
    CellRefNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    WholeColumnNode,
    WholeRowNode,
)
from excel_grapher.evaluator.parser import parse
from excel_grapher.grapher.formula_groups import shape_fingerprint


def _index_match(*, lookup_cell: str, match_type: float = 0.0) -> FunctionCallNode:
    """Build the INDEX/MATCH shape used in the Issue 2 design example."""
    return FunctionCallNode(
        name="INDEX",
        args=[
            RangeNode(start="Sheet1!D40", end="Sheet1!AJ50"),
            FunctionCallNode(
                name="MATCH",
                args=[
                    NumberNode(value=1.0),
                    RangeNode(start="Sheet1!AJ40", end="Sheet1!AJ50"),
                    NumberNode(value=0.0),
                ],
            ),
            FunctionCallNode(
                name="MATCH",
                args=[
                    CellRefNode(address=lookup_cell),
                    RangeNode(start="Sheet1!D35", end="Sheet1!Y35"),
                    NumberNode(value=match_type),
                ],
            ),
        ],
    )


def test_fingerprint_ignores_concrete_addresses() -> None:
    a = _index_match(lookup_cell="Sheet1!D35")
    b = _index_match(lookup_cell="Sheet2!Z9")
    assert shape_fingerprint(a) == shape_fingerprint(b)


def test_fingerprint_distinguishes_literals() -> None:
    a = _index_match(lookup_cell="Sheet1!D35", match_type=0.0)
    b = _index_match(lookup_cell="Sheet1!D35", match_type=1.0)
    assert shape_fingerprint(a) != shape_fingerprint(b)


def test_fingerprint_distinguishes_ops_and_function_names() -> None:
    plus = parse("=Sheet1!A1+Sheet1!B1")
    minus = parse("=Sheet1!A1-Sheet1!B1")
    assert shape_fingerprint(plus) != shape_fingerprint(minus)

    suma = parse("=SUM(Sheet1!A1:B2)")
    avg = parse("=AVERAGE(Sheet1!A1:B2)")
    assert shape_fingerprint(suma) != shape_fingerprint(avg)


def test_fingerprint_distinguishes_leaf_kinds() -> None:
    cell = CellRefNode(address="Sheet1!A1")
    rng = RangeNode(start="Sheet1!A1", end="Sheet1!B2")
    assert shape_fingerprint(cell) != shape_fingerprint(rng)

    col = WholeColumnNode(sheet="Sheet1", column="A")
    row = WholeRowNode(sheet="Sheet1", row=1)
    assert shape_fingerprint(col) != shape_fingerprint(row)
    assert shape_fingerprint(col) != shape_fingerprint(cell)
    assert shape_fingerprint(row) != shape_fingerprint(rng)


def test_fingerprint_of_skeleton_matches_concrete() -> None:
    concrete = _index_match(lookup_cell="Sheet1!D35")
    # Same shape with typed holes in place of every address leaf (walk order).
    skeleton = FunctionCallNode(
        name="INDEX",
        args=[
            AddressHoleNode(kind=AddressLeafKind.range, slot=0),
            FunctionCallNode(
                name="MATCH",
                args=[
                    NumberNode(value=1.0),
                    AddressHoleNode(kind=AddressLeafKind.range, slot=1),
                    NumberNode(value=0.0),
                ],
            ),
            FunctionCallNode(
                name="MATCH",
                args=[
                    AddressHoleNode(kind=AddressLeafKind.cell, slot=2),
                    AddressHoleNode(kind=AddressLeafKind.range, slot=3),
                    NumberNode(value=0.0),
                ],
            ),
        ],
    )
    assert shape_fingerprint(concrete) == shape_fingerprint(skeleton)


def test_fingerprint_stable_string() -> None:
    fp = shape_fingerprint(_index_match(lookup_cell="Sheet1!D35"))
    assert isinstance(fp, str)
    assert fp  # non-empty
    assert "Sheet1" not in fp  # addresses must not appear
