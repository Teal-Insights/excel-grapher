"""Parameterized formula AST shapes (address holes + parameter bindings).

Fingerprint a formula AST by punching cell/range/whole-column/whole-row leaves
into typed holes. Formulas that differ only in those addresses share a
`shape_key` and skeleton; each instance carries its own parameter tuple.

See GitHub #517.
"""

from __future__ import annotations

from collections import Counter
from collections.abc import Iterable, Iterator
from dataclasses import dataclass
from typing import Literal, Protocol, TypeAlias, cast

from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    BoolNode,
    CellRefNode,
    EmptyArgNode,
    ErrorNode,
    FormulaParseError,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    StringNode,
    UnaryOpNode,
    WholeColumnNode,
    WholeRowNode,
    parse,
)

AddressKind: TypeAlias = Literal["CELL", "RANGE", "WHOLE_COL", "WHOLE_ROW"]
AddressLeaf: TypeAlias = CellRefNode | RangeNode | WholeColumnNode | WholeRowNode


@dataclass(frozen=True, slots=True)
class AddressHoleNode:
    """Typed hole left after punching an address leaf out of a formula AST."""

    kind: AddressKind
    index: int


# Skeletons reuse concrete AST node classes for non-address structure; address
# leaves are replaced by `AddressHoleNode`. Nested `FunctionCallNode.args` /
# binary/unary children may therefore contain holes at runtime.
SkeletonNode: TypeAlias = AstNode | AddressHoleNode


@dataclass(frozen=True, slots=True)
class FormulaShape:
    """Shared formula shape plus the address parameters for one instance."""

    shape_key: str
    skeleton: SkeletonNode
    params: tuple[AddressLeaf, ...]


def _address_kind(leaf: AddressLeaf) -> AddressKind:
    match leaf:
        case CellRefNode():
            return "CELL"
        case RangeNode():
            return "RANGE"
        case WholeColumnNode():
            return "WHOLE_COL"
        case WholeRowNode():
            return "WHOLE_ROW"
    raise TypeError(f"not an address leaf: {type(leaf).__name__}")


def _stable_float(value: float) -> str:
    """Return a stable, parse-friendly float token for shape keys."""
    return format(value, ".17g")


def _punch(node: AstNode, params: list[AddressLeaf]) -> tuple[str, SkeletonNode]:
    """Return `(shape_key_fragment, skeleton_subtree)` for `node`."""
    match node:
        case NumberNode(value):
            return f"N({_stable_float(value)})", node
        case StringNode(value):
            return f"S({value!r})", node
        case BoolNode(value):
            return f"B({value})", node
        case ErrorNode(error):
            return f"E({error.value})", node
        case EmptyArgNode():
            return "$EMPTY", node
        case CellRefNode() | RangeNode() | WholeColumnNode() | WholeRowNode() as leaf:
            kind = _address_kind(leaf)
            index = len(params)
            params.append(leaf)
            hole = AddressHoleNode(kind=kind, index=index)
            return f"${kind}", hole
        case FunctionCallNode(name, args):
            parts: list[str] = []
            skel_args: list[SkeletonNode] = []
            for arg in args:
                key_part, skel_arg = _punch(arg, params)
                parts.append(key_part)
                skel_args.append(skel_arg)
            key = f"F({name},[{','.join(parts)}])"
            skeleton = FunctionCallNode(name, cast(list[AstNode], skel_args))
            return key, skeleton
        case BinaryOpNode(op, left, right):
            left_key, left_skel = _punch(left, params)
            right_key, right_skel = _punch(right, params)
            key = f"B({op},{left_key},{right_key})"
            skeleton = BinaryOpNode(op, cast(AstNode, left_skel), cast(AstNode, right_skel))
            return key, skeleton
        case UnaryOpNode(op, operand):
            operand_key, operand_skel = _punch(operand, params)
            key = f"U({op},{operand_key})"
            skeleton = UnaryOpNode(op, cast(AstNode, operand_skel))
            return key, skeleton
    raise TypeError(f"unsupported AST node: {type(node).__name__}")


def fingerprint_formula_shape(ast_or_formula: AstNode | str) -> FormulaShape:
    """Punch address leaves out of a formula AST.

    Args:
        ast_or_formula: Parsed AST, or a normalized formula string (`=` optional).

    Returns:
        A `FormulaShape` whose `shape_key` is a stable structural token (ops,
        function names/arity, and literals fixed; address sites as typed holes),
        `skeleton` is the punched tree, and `params` is the ordered tuple of
        original address leaves in preorder walk order.
    """
    if isinstance(ast_or_formula, str):
        text = ast_or_formula.strip()
        if text and not text.startswith("="):
            text = "=" + text
        ast: AstNode = parse(text)
    else:
        ast = ast_or_formula

    params: list[AddressLeaf] = []
    shape_key, skeleton = _punch(ast, params)
    return FormulaShape(shape_key=shape_key, skeleton=skeleton, params=tuple(params))


def _fill(node: SkeletonNode, params: tuple[AddressLeaf, ...], seen: list[int]) -> AstNode:
    match node:
        case AddressHoleNode(kind, index):
            if index < 0 or index >= len(params):
                raise ValueError(
                    f"hole index {index} out of range for {len(params)} params (arity mismatch)"
                )
            leaf = params[index]
            actual = _address_kind(leaf)
            if actual != kind:
                raise ValueError(f"kind mismatch at hole {index}: expected {kind}, got {actual}")
            seen.append(index)
            return leaf
        case (
            NumberNode()
            | StringNode()
            | BoolNode()
            | ErrorNode()
            | EmptyArgNode()
            | CellRefNode()
            | RangeNode()
            | WholeColumnNode()
            | WholeRowNode()
        ):
            return node
        case FunctionCallNode(name, args):
            return FunctionCallNode(name, [_fill(arg, params, seen) for arg in args])
        case BinaryOpNode(op, left, right):
            return BinaryOpNode(op, _fill(left, params, seen), _fill(right, params, seen))
        case UnaryOpNode(op, operand):
            return UnaryOpNode(op, _fill(operand, params, seen))
    raise TypeError(f"unsupported skeleton node: {type(node).__name__}")


def specialize_formula_shape(
    skeleton: SkeletonNode,
    params: tuple[AddressLeaf, ...] | list[AddressLeaf],
) -> AstNode:
    """Fill address holes in `skeleton` with `params` in hole-index order.

    Args:
        skeleton: Tree produced by `fingerprint_formula_shape`.
        params: Address leaves whose length and kinds must match the holes.

    Returns:
        A concrete `AstNode` with holes replaced.

    Raises:
        ValueError: On hole/param arity or kind mismatch.
    """
    param_tuple = tuple(params)
    seen: list[int] = []
    result = _fill(skeleton, param_tuple, seen)
    if len(seen) != len(param_tuple):
        raise ValueError(
            f"param arity mismatch: skeleton has {len(seen)} holes, got {len(param_tuple)} params"
        )
    if seen != list(range(len(param_tuple))):
        raise ValueError(f"hole indices must be a dense 0..n-1 sequence in preorder; got {seen!r}")
    return result


@dataclass(frozen=True, slots=True)
class FormulaShapeSummary:
    """Cardinality of exact formulas vs punched AST shapes.

    `formula_nodes` counts successfully fingerprinted instances only.
    `unparseable` is the count of fingerprint/`parse` failures excluded from
    that total (and from shape counts).
    """

    formula_nodes: int
    distinct_normalized_formulas: int
    distinct_shapes: int
    unparseable: int
    shape_counts: tuple[tuple[str, int], ...]

    @property
    def shapes_per_formula_string(self) -> float:
        """`distinct_shapes / distinct_normalized_formulas` (1.0 means no collapse)."""
        if self.distinct_normalized_formulas == 0:
            return 0.0
        return self.distinct_shapes / self.distinct_normalized_formulas

    @property
    def mean_instances_per_shape(self) -> float:
        """Average successfully fingerprinted formula nodes per distinct shape."""
        if self.distinct_shapes == 0:
            return 0.0
        return self.formula_nodes / self.distinct_shapes

    def to_dict(self) -> dict[str, object]:
        """JSON-serializable report dict."""
        return {
            "formula_nodes": self.formula_nodes,
            "distinct_normalized_formulas": self.distinct_normalized_formulas,
            "distinct_shapes": self.distinct_shapes,
            "unparseable": self.unparseable,
            "shapes_per_formula_string": self.shapes_per_formula_string,
            "mean_instances_per_shape": self.mean_instances_per_shape,
            "shape_counts": [
                {"shape_key": key, "count": count} for key, count in self.shape_counts
            ],
        }


class _FormulaNodeView(Protocol):
    """Minimal node surface needed by `summarize_formula_shapes`."""

    @property
    def normalized_formula(self) -> str | None: ...


class _FormulaGraphView(Protocol):
    """Minimal graph surface needed by `summarize_formula_shapes`.

    Kept as a `Protocol` so `core` does not import `grapher` (package boundary).
    """

    def formula_nodes(self) -> Iterator[tuple[object, _FormulaNodeView]]: ...


def summarize_normalized_formulas(
    formulas: Iterable[str],
) -> tuple[FormulaShapeSummary, list[str]]:
    """Fingerprint an iterable of normalized formula strings.

    Args:
        formulas: Already-normalized formula texts (leading `=` optional).

    Returns:
        `(summary, parseable_formulas)` where `parseable_formulas` are the
        successfully fingerprinted inputs (same order), suitable for parse-warm
        timing. Failed parses increment `summary.unparseable` and are omitted
        from `formula_nodes` / shape counts.
    """
    parseable: list[str] = []
    shape_counter: Counter[str] = Counter()
    unparseable = 0

    for formula in formulas:
        stripped = formula.strip()
        if not stripped:
            continue
        try:
            shape = fingerprint_formula_shape(stripped)
        except FormulaParseError:
            unparseable += 1
            continue
        parseable.append(stripped)
        shape_counter[shape.shape_key] += 1

    summary = FormulaShapeSummary(
        formula_nodes=len(parseable),
        distinct_normalized_formulas=len(set(parseable)),
        distinct_shapes=len(shape_counter),
        unparseable=unparseable,
        shape_counts=tuple(shape_counter.most_common()),
    )
    return summary, parseable


def summarize_formula_shapes(graph: _FormulaGraphView) -> FormulaShapeSummary:
    """Count distinct normalized formulas vs punched AST shapes in `graph`.

    Walks `graph.formula_nodes()`, fingerprints each `normalized_formula`, and
    reports whether shapes collapse the string-keyed set (#517 go/no-go metric).
    """
    formulas: list[str] = []
    for _, node in graph.formula_nodes():
        nf = node.normalized_formula
        if isinstance(nf, str) and nf.strip():
            formulas.append(nf.strip())
    summary, _ = summarize_normalized_formulas(formulas)
    return summary
