"""Unit tests for `specialize_group`."""

from __future__ import annotations

import pytest

from excel_grapher.core.formula_ast import (
    AddressHoleNode,
    AddressLeafKind,
    CellRefNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
)
from excel_grapher.grapher.formula_groups import SpecializeError, specialize_group


def test_specialize_fills_single_cell_hole() -> None:
    skeleton = AddressHoleNode(kind=AddressLeafKind.cell, slot=0)
    out = specialize_group(skeleton, (CellRefNode(address="Sheet1!D35"),))
    assert out == CellRefNode(address="Sheet1!D35")


def test_specialize_fills_holes_in_walk_order() -> None:
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
    bindings = (
        RangeNode(start="Sheet1!D40", end="Sheet1!AJ50"),
        RangeNode(start="Sheet1!AJ40", end="Sheet1!AJ50"),
        CellRefNode(address="Sheet1!D35"),
        RangeNode(start="Sheet1!D35", end="Sheet1!Y35"),
    )
    out = specialize_group(skeleton, bindings)
    assert out == FunctionCallNode(
        name="INDEX",
        args=[
            bindings[0],
            FunctionCallNode(
                name="MATCH",
                args=[NumberNode(value=1.0), bindings[1], NumberNode(value=0.0)],
            ),
            FunctionCallNode(
                name="MATCH",
                args=[bindings[2], bindings[3], NumberNode(value=0.0)],
            ),
        ],
    )


def test_specialize_rejects_kind_mismatch() -> None:
    skeleton = AddressHoleNode(kind=AddressLeafKind.cell, slot=0)
    with pytest.raises(SpecializeError, match="kind"):
        specialize_group(
            skeleton,
            (RangeNode(start="Sheet1!A1", end="Sheet1!B2"),),
        )


def test_specialize_rejects_arity_mismatch_too_few() -> None:
    skeleton = FunctionCallNode(
        name="SUM",
        args=[
            AddressHoleNode(kind=AddressLeafKind.cell, slot=0),
            AddressHoleNode(kind=AddressLeafKind.cell, slot=1),
        ],
    )
    with pytest.raises(SpecializeError, match="binding"):
        specialize_group(skeleton, (CellRefNode(address="Sheet1!A1"),))


def test_specialize_rejects_arity_mismatch_too_many() -> None:
    skeleton = AddressHoleNode(kind=AddressLeafKind.cell, slot=0)
    with pytest.raises(SpecializeError, match="binding"):
        specialize_group(
            skeleton,
            (
                CellRefNode(address="Sheet1!A1"),
                CellRefNode(address="Sheet1!B1"),
            ),
        )


def test_specialize_preserves_non_address_structure() -> None:
    skeleton = FunctionCallNode(
        name="IF",
        args=[
            NumberNode(value=1.0),
            AddressHoleNode(kind=AddressLeafKind.cell, slot=0),
            NumberNode(value=0.0),
        ],
    )
    out = specialize_group(skeleton, (CellRefNode(address="Sheet1!Z9"),))
    assert out == FunctionCallNode(
        name="IF",
        args=[
            NumberNode(value=1.0),
            CellRefNode(address="Sheet1!Z9"),
            NumberNode(value=0.0),
        ],
    )
