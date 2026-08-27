"""JSON encode/decode for `excel_grapher.core.formula_ast.AstNode` trees."""

from __future__ import annotations

from typing import Any, cast

from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    BoolNode,
    CellRefNode,
    EmptyArgNode,
    ErrorNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    StringNode,
    UnaryOpNode,
    WholeColumnNode,
    WholeRowNode,
)
from excel_grapher.core.types import XlError

_JsonObject = dict[str, Any]


def ast_to_json(node: AstNode) -> _JsonObject:
    """Encode `node` as a JSON-serializable object."""
    match node:
        case NumberNode(value):
            return {"t": "num", "v": value}
        case StringNode(value):
            return {"t": "str", "v": value}
        case BoolNode(value):
            return {"t": "bool", "v": value}
        case ErrorNode(error):
            return {"t": "err", "v": error.value}
        case CellRefNode(address):
            return {"t": "cell", "v": address}
        case RangeNode(start, end):
            return {"t": "range", "s": start, "e": end}
        case WholeColumnNode(sheet, column):
            return {"t": "whole_col", "sheet": sheet, "column": column}
        case WholeRowNode(sheet, row):
            return {"t": "whole_row", "sheet": sheet, "row": row}
        case FunctionCallNode(name, args):
            return {"t": "fn", "n": name, "a": [ast_to_json(arg) for arg in args]}
        case BinaryOpNode(op, left, right):
            return {"t": "bin", "op": op, "l": ast_to_json(left), "r": ast_to_json(right)}
        case UnaryOpNode(op, operand):
            return {"t": "un", "op": op, "x": ast_to_json(operand)}
        case EmptyArgNode():
            return {"t": "empty"}
    raise TypeError(f"unsupported AST node: {type(node).__name__}")


def ast_from_json(payload: object) -> AstNode:
    """Decode an `AstNode` previously encoded by `ast_to_json`.

    Raises:
        TypeError: If `payload` is not a recognized AST encoding.
    """
    if not isinstance(payload, dict):
        raise TypeError("AST JSON must be an object")
    d = cast(_JsonObject, payload)
    tag = d.get("t")
    if tag == "num":
        value = d.get("v")
        if not isinstance(value, (int, float)) or isinstance(value, bool):
            raise TypeError("num payload must be a number")
        return NumberNode(float(value))
    if tag == "str":
        value = d.get("v")
        if not isinstance(value, str):
            raise TypeError("str payload must be a string")
        return StringNode(value)
    if tag == "bool":
        value = d.get("v")
        if not isinstance(value, bool):
            raise TypeError("bool payload must be a bool")
        return BoolNode(value)
    if tag == "err":
        value = d.get("v")
        if not isinstance(value, str):
            raise TypeError("err payload must be a string")
        error = XlError.from_text(value)
        if error is None:
            raise TypeError(f"unknown Excel error literal: {value!r}")
        return ErrorNode(error)
    if tag == "cell":
        value = d.get("v")
        if not isinstance(value, str):
            raise TypeError("cell payload must be a string")
        return CellRefNode(value)
    if tag == "range":
        start = d.get("s")
        end = d.get("e")
        if not isinstance(start, str) or not isinstance(end, str):
            raise TypeError("range payload must have string start and end")
        return RangeNode(start, end)
    if tag == "whole_col":
        sheet = d.get("sheet")
        column = d.get("column")
        if not isinstance(sheet, str) or not isinstance(column, str):
            raise TypeError("whole_col payload must have string sheet and column")
        return WholeColumnNode(sheet=sheet, column=column)
    if tag == "whole_row":
        sheet = d.get("sheet")
        row = d.get("row")
        if not isinstance(sheet, str) or not isinstance(row, int) or isinstance(row, bool):
            raise TypeError("whole_row payload must have string sheet and int row")
        return WholeRowNode(sheet=sheet, row=row)
    if tag == "fn":
        name = d.get("n")
        args = d.get("a")
        if not isinstance(name, str) or not isinstance(args, list):
            raise TypeError("fn payload must have string name and list args")
        return FunctionCallNode(name, [ast_from_json(arg) for arg in args])
    if tag == "bin":
        op = d.get("op")
        if not isinstance(op, str):
            raise TypeError("bin payload must have string op")
        return BinaryOpNode(op, ast_from_json(d.get("l")), ast_from_json(d.get("r")))
    if tag == "un":
        op = d.get("op")
        if not isinstance(op, str):
            raise TypeError("un payload must have string op")
        return UnaryOpNode(op, ast_from_json(d.get("x")))
    if tag == "empty":
        return EmptyArgNode()
    raise TypeError(f"unknown AST JSON tag: {tag!r}")
