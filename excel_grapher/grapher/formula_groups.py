"""Formula-group shape fingerprinting and skeleton specialization (Issue 2).

Provides `shape_fingerprint`, `specialize_group`, and template-field validation
used by hand-built Option B group nodes (evaluator/codegen wire-up later).
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from typing import TypeAlias

from excel_grapher.core.address_keys import CellKey, format_range_key, parse_address, parse_node_key
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


def serialize_address_leaf(leaf: AddressLeaf) -> str:
    """Serialize an address leaf to a canonical address string for export wrappers.

    Cell refs become `Sheet!A1`. Ranges become a single-prefix range key
    (`Sheet!A1:B2`). Whole-column / whole-row refs use `Sheet!A:A` / `Sheet!1:1`.
    """
    if isinstance(leaf, CellRefNode):
        return leaf.address
    if isinstance(leaf, RangeNode):
        start_sheet, start_cell = parse_address(leaf.start)
        end_sheet, end_cell = parse_address(leaf.end)
        if start_sheet != end_sheet:
            return f"{leaf.start}:{leaf.end}"
        return format_range_key(start_sheet, start_cell, end_cell)
    if isinstance(leaf, WholeColumnNode):
        return f"{leaf.sheet}!{leaf.column}:{leaf.column}"
    return f"{leaf.sheet}!{leaf.row}:{leaf.row}"


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


def collect_holes(skeleton: AstNode) -> list[AddressHoleNode]:
    """Return `AddressHoleNode`s in `skeleton` in preorder walk order."""
    return _collect_holes(skeleton)


def validate_group_template(
    *,
    members: Sequence[str],
    skeleton: AstNode,
    member_bindings: Mapping[str, Sequence[AddressLeaf]],
    shape_fingerprint_value: str,
) -> dict[str, tuple[AddressLeaf, ...]]:
    """Validate and canonicalize formula-group template fields.

    Args:
        members: Canonical member cell keys owned by the group.
        skeleton: Template AST with typed address holes.
        member_bindings: Per-member binding sequences (hole walk order).
        shape_fingerprint_value: Expected fingerprint for `skeleton`.

    Returns:
        Normalized `member_bindings` with canonical cell keys and tuples.

    Raises:
        ValueError: If fingerprint, membership, arity, or kinds are invalid.
    """
    expected_fp = shape_fingerprint(skeleton)
    if shape_fingerprint_value != expected_fp:
        raise ValueError(
            "shape_fingerprint does not match skeleton: "
            f"got {shape_fingerprint_value!r}, expected {expected_fp!r}"
        )

    member_set: set[str] = set()
    ordered_members: list[str] = []
    for raw in members:
        parsed = parse_node_key(raw)
        if not isinstance(parsed, CellKey):
            raise ValueError(f"Group members must be single cells; got {raw!r}")
        key = str(parsed)
        if key not in member_set:
            member_set.add(key)
            ordered_members.append(key)

    normalized: dict[str, tuple[AddressLeaf, ...]] = {}
    for raw_key, bindings in member_bindings.items():
        parsed = parse_node_key(raw_key)
        if not isinstance(parsed, CellKey):
            raise ValueError(f"Binding keys must be single cells; got {raw_key!r}")
        key = str(parsed)
        if key not in member_set:
            raise ValueError(f"Binding for non-member cell {key!r}")
        binding_tuple = tuple(bindings)
        try:
            specialize_group(skeleton, binding_tuple)
        except SpecializeError as exc:
            raise ValueError(f"Invalid bindings for {key}: {exc}") from exc
        normalized[key] = binding_tuple

    missing = [m for m in ordered_members if m not in normalized]
    if missing:
        raise ValueError(f"Missing member_bindings for: {', '.join(missing)}")

    return normalized


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
