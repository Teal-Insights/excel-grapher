"""JSON encode/decode for `excel_grapher.core.formula_ast.AstNode` trees."""

from __future__ import annotations

import hashlib
import json
from typing import Any, cast

from fastpyxl.utils.cell import column_index_from_string

from excel_grapher.core.formula_ast import (
    AbsoluteAxis,
    AstNode,
    AxisRef,
    BinaryOpNode,
    BoolNode,
    CellRef,
    CellRefNode,
    EmptyArgNode,
    ErrorNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    RelativeAxis,
    StringNode,
    UnaryOpNode,
    WholeColumnNode,
    WholeRowNode,
    cell_ref_from_a1,
)
from excel_grapher.core.types import XlError

_JsonObject = dict[str, Any]


def _axis_to_json(axis: AbsoluteAxis | RelativeAxis) -> _JsonObject:
    if isinstance(axis, AbsoluteAxis):
        return {"k": "abs", "n": axis.index}
    return {"k": "rel", "n": axis.offset}


def _cell_ref_to_json(ref: CellRef) -> _JsonObject:
    return {
        "sheet": ref.sheet,
        "col": _axis_to_json(ref.col),
        "row": _axis_to_json(ref.row),
    }


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
        case CellRefNode(ref):
            return {"t": "cell", **_cell_ref_to_json(ref)}
        case RangeNode(start_ref, end_ref):
            return {
                "t": "range",
                "s": _cell_ref_to_json(start_ref),
                "e": _cell_ref_to_json(end_ref),
            }
        case WholeColumnNode(sheet, col):
            return {"t": "whole_col", "sheet": sheet, "col": _axis_to_json(col)}
        case WholeRowNode(sheet, row):
            return {"t": "whole_row", "sheet": sheet, "row": _axis_to_json(row)}
        case FunctionCallNode(name, args):
            return {"t": "fn", "n": name, "a": [ast_to_json(arg) for arg in args]}
        case BinaryOpNode(op, left, right):
            return {"t": "bin", "op": op, "l": ast_to_json(left), "r": ast_to_json(right)}
        case UnaryOpNode(op, operand):
            return {"t": "un", "op": op, "x": ast_to_json(operand)}
        case EmptyArgNode():
            return {"t": "empty"}
    raise TypeError(f"unsupported AST node: {type(node).__name__}")


def ast_identity_key(node: AstNode) -> str:
    """Canonical JSON document used as a formula-AST identity key."""
    return json.dumps(ast_to_json(node), sort_keys=True, separators=(",", ":"))


def formula_identity_digest(*, formula: str, formula_ast: AstNode | None) -> str:
    """SHA-256 of `formula_ast` when present, otherwise of `formula` text.

    Type-analysis and similar caches key formula identity by AST so relative
    vs absolute trees that share A1 text do not collide. Unparseable cells
    fall back to the stored formula string.
    """
    payload = ast_identity_key(formula_ast) if formula_ast is not None else formula
    return hashlib.sha256(payload.encode()).hexdigest()


def _axis_from_json(payload: object) -> AxisRef:
    if not isinstance(payload, dict):
        raise TypeError("axis payload must be an object")
    d = cast(_JsonObject, payload)
    kind = d.get("k")
    n = d.get("n")
    if not isinstance(n, int) or isinstance(n, bool):
        raise TypeError("axis payload must have int n")
    if kind == "abs":
        return AbsoluteAxis(n)
    if kind == "rel":
        return RelativeAxis(n)
    raise TypeError(f"unknown axis kind: {kind!r}")


def _cell_ref_from_json(payload: object, *, legacy_a1: object = None) -> CellRef:
    if isinstance(legacy_a1, str):
        return cell_ref_from_a1(legacy_a1)
    if not isinstance(payload, dict):
        raise TypeError("cell ref payload must be an object")
    d = cast(_JsonObject, payload)
    sheet = d.get("sheet")
    if not isinstance(sheet, str):
        raise TypeError("cell ref payload must have string sheet")
    return CellRef(
        sheet=sheet, col=_axis_from_json(d.get("col")), row=_axis_from_json(d.get("row"))
    )


def _axis_or_column_letter_from_json(payload: object) -> AxisRef:
    if isinstance(payload, str):
        return AbsoluteAxis(int(column_index_from_string(payload)))
    return _axis_from_json(payload)


def _axis_or_int_from_json(payload: object) -> AxisRef:
    if isinstance(payload, int) and not isinstance(payload, bool):
        return AbsoluteAxis(payload)
    return _axis_from_json(payload)


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
        return CellRefNode(_cell_ref_from_json(d, legacy_a1=d.get("v")))
    if tag == "range":
        start = d.get("s")
        end = d.get("e")
        return RangeNode(
            start_ref=_cell_ref_from_json(
                start, legacy_a1=start if isinstance(start, str) else None
            ),
            end_ref=_cell_ref_from_json(end, legacy_a1=end if isinstance(end, str) else None),
        )
    if tag == "whole_col":
        sheet = d.get("sheet")
        if not isinstance(sheet, str):
            raise TypeError("whole_col payload must have string sheet")
        col = d.get("col", d.get("column"))
        return WholeColumnNode(sheet=sheet, col=_axis_or_column_letter_from_json(col))
    if tag == "whole_row":
        sheet = d.get("sheet")
        row = d.get("row")
        if not isinstance(sheet, str):
            raise TypeError("whole_row payload must have string sheet")
        return WholeRowNode(sheet=sheet, row=_axis_or_int_from_json(row))
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
