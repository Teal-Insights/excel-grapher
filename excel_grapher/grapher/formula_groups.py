"""Formula-group shape fingerprinting and skeleton specialization (Issue 2).

Sprint 1: fingerprint + `specialize_group` only. Evaluator and codegen wire-up
come in later sprints.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import TypeAlias

from excel_grapher.core.formula_ast import (
    AddressHoleNode,
    AddressLeafKind,
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

AddressLeaf: TypeAlias = CellRefNode | RangeNode | WholeColumnNode | WholeRowNode


class SpecializeError(ValueError):
    """Raised when skeleton specialization fails (arity or kind mismatch)."""


def address_leaf_kind(leaf: AddressLeaf | AddressHoleNode) -> AddressLeafKind:
    """Return the address-leaf kind for a concrete leaf or hole."""
    if isinstance(leaf, AddressHoleNode):
        return leaf.kind
    if isinstance(leaf, CellRefNode):
        return AddressLeafKind.cell
    if isinstance(leaf, RangeNode):
        return AddressLeafKind.range
    if isinstance(leaf, WholeColumnNode):
        return AddressLeafKind.whole_column
    if isinstance(leaf, WholeRowNode):
        return AddressLeafKind.whole_row
    raise TypeError(f"Not an address leaf: {type(leaf)!r}")


def shape_fingerprint(ast: AstNode) -> str:
    """Return a shape fingerprint for `ast`.

    Address leaves (`CellRefNode` / `RangeNode` / `WholeColumnNode` /
    `WholeRowNode`) and `AddressHoleNode` contribute only their leaf kind.
    Concrete addresses are ignored so formulas that differ only in refs share a
    fingerprint. Ops, function names, arity, and literals remain concrete.

    Args:
        ast: Formula AST (concrete member formula or skeleton with holes).

    Returns:
        Stable fingerprint string suitable for equality comparisons.
    """
    parts: list[str] = []
    _append_fingerprint(ast, parts)
    return "".join(parts)


def specialize_group(skeleton: AstNode, bindings: Sequence[AddressLeaf]) -> AstNode:
    """Fill `AddressHoleNode` slots in `skeleton` with `bindings` in walk order.

    Args:
        skeleton: Group template AST containing zero or more address holes.
        bindings: Concrete address leaves, one per hole in preorder walk order.

    Returns:
        A new AST with holes replaced by the corresponding bindings.

    Raises:
        SpecializeError: If binding count or leaf kinds do not match the holes.
    """
    holes = _collect_holes(skeleton)
    if len(bindings) != len(holes):
        raise SpecializeError(
            f"binding arity mismatch: skeleton has {len(holes)} hole(s), "
            f"got {len(bindings)} binding(s)"
        )
    for hole, binding in zip(holes, bindings, strict=True):
        got = address_leaf_kind(binding)
        if got is not hole.kind:
            raise SpecializeError(
                f"kind mismatch at slot {hole.slot}: hole is {hole.kind.value}, "
                f"binding is {got.value}"
            )
    index = 0

    def _fill(node: AstNode) -> AstNode:
        nonlocal index
        if isinstance(node, AddressHoleNode):
            leaf = bindings[index]
            index += 1
            return leaf
        if isinstance(node, FunctionCallNode):
            return FunctionCallNode(name=node.name, args=[_fill(a) for a in node.args])
        if isinstance(node, BinaryOpNode):
            return BinaryOpNode(op=node.op, left=_fill(node.left), right=_fill(node.right))
        if isinstance(node, UnaryOpNode):
            return UnaryOpNode(op=node.op, operand=_fill(node.operand))
        return node

    return _fill(skeleton)


def _collect_holes(node: AstNode) -> list[AddressHoleNode]:
    holes: list[AddressHoleNode] = []

    def _walk(n: AstNode) -> None:
        if isinstance(n, AddressHoleNode):
            holes.append(n)
            return
        if isinstance(n, FunctionCallNode):
            for arg in n.args:
                _walk(arg)
            return
        if isinstance(n, BinaryOpNode):
            _walk(n.left)
            _walk(n.right)
            return
        if isinstance(n, UnaryOpNode):
            _walk(n.operand)

    _walk(node)
    return holes


def _append_fingerprint(node: AstNode, parts: list[str]) -> None:
    if isinstance(node, NumberNode):
        parts.append(f"N:{node.value!r}")
        return
    if isinstance(node, StringNode):
        parts.append(f"S:{node.value!r}")
        return
    if isinstance(node, BoolNode):
        parts.append(f"B:{node.value!r}")
        return
    if isinstance(node, ErrorNode):
        parts.append(f"E:{node.error!r}")
        return
    if isinstance(node, EmptyArgNode):
        parts.append("EMPTY")
        return
    if isinstance(node, (CellRefNode, RangeNode, WholeColumnNode, WholeRowNode, AddressHoleNode)):
        parts.append(f"A:{address_leaf_kind(node).value}")
        return
    if isinstance(node, FunctionCallNode):
        parts.append(f"F:{node.name.upper()}(")
        for i, arg in enumerate(node.args):
            if i:
                parts.append(",")
            _append_fingerprint(arg, parts)
        parts.append(")")
        return
    if isinstance(node, BinaryOpNode):
        parts.append(f"O:{node.op}(")
        _append_fingerprint(node.left, parts)
        parts.append(",")
        _append_fingerprint(node.right, parts)
        parts.append(")")
        return
    if isinstance(node, UnaryOpNode):
        parts.append(f"U:{node.op}(")
        _append_fingerprint(node.operand, parts)
        parts.append(")")
        return
    raise TypeError(f"Unknown AST node: {type(node)!r}")
