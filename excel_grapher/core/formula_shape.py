"""Parameterized formula AST shapes (address holes + parameter bindings).

Fingerprint a formula AST by punching cell/range/whole-column/whole-row leaves
into typed holes. Formulas that differ only in those addresses share a
`shape_key` and skeleton; each instance carries its own parameter tuple.

See GitHub #517. Overlay lifetime (opt-in, caller rewarm, evaluator
construction snapshot) is documented under GitHub #560 and
`excel_grapher.grapher.formula_shapes.warm_formula_shapes`.

`parse_preserving_axes` reuses this module's typed holes as a parse accelerator:
a private regex punch produces a skeleton the existing parser accepts, the
cached tree is projected into `AddressHoleNode`s, and `fill_address_holes`
rebinds per cell. The punched skeleton string is not a second public shape
language; identity for copies is `FormulaShape.shape_key`.
"""

from __future__ import annotations

import re
from collections import OrderedDict
from collections.abc import Iterable
from dataclasses import dataclass
from typing import Literal, TypeAlias, cast

from fastpyxl.utils.cell import get_column_letter

from excel_grapher.core.address_keys import CellKey
from excel_grapher.core.formula_ast import (
    AbsoluteAxis,
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


def fill_address_holes(
    skeleton: SkeletonNode,
    params: tuple[AddressLeaf, ...] | list[AddressLeaf],
) -> AstNode:
    """Replace holes in `skeleton` (or a subtree) using `params` by hole index.

    Unlike `specialize_formula_shape`, this does not require the subtree to
    mention every param; nested INDEX/OFFSET args can fill a subset of holes.
    """
    param_tuple = tuple(params)
    seen: list[int] = []
    return _fill(skeleton, param_tuple, seen)


_PARSE_HOLE_SHEET = "_EG_SHAPE_HOLE"
_SHAPE_SHEET_PREFIX = r"(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!"
_SHAPE_COL = r"\$?[A-Za-z]{1,3}"
_SHAPE_ROW = r"\$?\d+"
_SHAPE_RANGE_RE = re.compile(
    rf"(?:{_SHAPE_SHEET_PREFIX})?{_SHAPE_COL}{_SHAPE_ROW}\s*:\s*"
    rf"(?:{_SHAPE_SHEET_PREFIX})?{_SHAPE_COL}{_SHAPE_ROW}"
)
_SHAPE_CELL_RE = re.compile(rf"(?:{_SHAPE_SHEET_PREFIX})?{_SHAPE_COL}{_SHAPE_ROW}")
_SHAPE_WHOLE_COL_RE = re.compile(
    rf"(?:{_SHAPE_SHEET_PREFIX})?{_SHAPE_COL}\s*:\s*(?:{_SHAPE_SHEET_PREFIX})?{_SHAPE_COL}(?!\d)"
)
_SHAPE_WHOLE_ROW_RE = re.compile(
    rf"(?:{_SHAPE_SHEET_PREFIX})?{_SHAPE_ROW}\s*:\s*(?:{_SHAPE_SHEET_PREFIX})?{_SHAPE_ROW}(?!\d)"
)
_SHAPE_PARSE_CACHE_MAXSIZE = 8192
_SHAPE_PARSE_CACHE: OrderedDict[str, SkeletonNode] = OrderedDict()
_SHAPE_PARSE_HITS = 0
_SHAPE_PARSE_MISSES = 0
_SHAPE_PARSE_FALLBACKS = 0


def clear_shape_parse_cache() -> None:
    """Drop the bounded skeleton parse cache (tests)."""
    global _SHAPE_PARSE_HITS, _SHAPE_PARSE_MISSES, _SHAPE_PARSE_FALLBACKS
    _SHAPE_PARSE_CACHE.clear()
    _SHAPE_PARSE_HITS = 0
    _SHAPE_PARSE_MISSES = 0
    _SHAPE_PARSE_FALLBACKS = 0


def shape_parse_cache_info() -> tuple[int, int, int]:
    """Return `(hits, misses, fallbacks)` for punched-skeleton formula parses."""
    return _SHAPE_PARSE_HITS, _SHAPE_PARSE_MISSES, _SHAPE_PARSE_FALLBACKS


def try_parse_preserving_axes_from_shape(formula: str, *, anchor: CellKey | str) -> AstNode | None:
    """Parse `formula` by filling a cached `FormulaShape` skeleton.

    Returns None when `formula` has no address holes, or when the accelerator
    cannot reconstruct the tree (caller should parse directly). The punched
    skeleton string is a private parse-cache key, not a public shape language.
    """
    global _SHAPE_PARSE_FALLBACKS
    skeleton_text, hole_texts = _skeletonize_formula_addresses(formula)
    if not hole_texts:
        return None
    try:
        skeleton = _cached_parse_skeleton(skeleton_text)
        params = tuple(
            _require_address_leaf(parse("=" + hole, anchor=anchor, preserve_axes=True))
            for hole in hole_texts
        )
        return fill_address_holes(skeleton, params)
    except (FormulaParseError, ValueError, IndexError, TypeError):
        _SHAPE_PARSE_FALLBACKS += 1
        return None


def _require_address_leaf(node: AstNode) -> AddressLeaf:
    if isinstance(node, (CellRefNode, RangeNode, WholeColumnNode, WholeRowNode)):
        return node
    raise ValueError(f"hole parse did not yield an address leaf: {type(node).__name__}")


def _mask_formula_strings_as_spaces(text: str) -> str:
    """Replace Excel `"..."` literals with spaces so address regexes skip them."""
    chars = list(text)
    i = 0
    n = len(text)
    while i < n:
        if text[i] != '"':
            i += 1
            continue
        j = i + 1
        while j < n:
            if text[j] != '"':
                j += 1
                continue
            if j + 1 < n and text[j + 1] == '"':
                j += 2
                continue
            j += 1
            break
        for k in range(i, min(j, n)):
            chars[k] = " "
        i = j
    return "".join(chars)


def _is_function_call_span(text: str, end: int) -> bool:
    rest = text[end:]
    i = 0
    while i < len(rest) and rest[i].isspace():
        i += 1
    return i < len(rest) and rest[i] == "("


def _formula_address_spans(formula: str) -> list[tuple[int, int, AddressKind]]:
    """Return non-overlapping `(start, end, kind)` address spans."""
    masked = _mask_formula_strings_as_spaces(formula)
    found: list[tuple[int, int, AddressKind]] = []
    for regex, kind in (
        (_SHAPE_RANGE_RE, "RANGE"),
        (_SHAPE_WHOLE_COL_RE, "WHOLE_COL"),
        (_SHAPE_WHOLE_ROW_RE, "WHOLE_ROW"),
        (_SHAPE_CELL_RE, "CELL"),
    ):
        for match in regex.finditer(masked):
            start, end = match.span()
            if _is_function_call_span(masked, end):
                continue
            found.append((start, end, kind))
    found.sort(key=lambda item: (item[0], -(item[1] - item[0])))
    spans: list[tuple[int, int, AddressKind]] = []
    last_end = -1
    for start, end, kind in found:
        if start < last_end:
            continue
        spans.append((start, end, kind))
        last_end = end
    return spans


def _hole_placeholder(kind: AddressKind, index: int) -> str:
    n = index + 1
    if kind == "CELL":
        return f"{_PARSE_HOLE_SHEET}!A{n}"
    if kind == "RANGE":
        return f"{_PARSE_HOLE_SHEET}!A{n}:B{n}"
    if kind == "WHOLE_COL":
        letter = get_column_letter(n)
        return f"{_PARSE_HOLE_SHEET}!{letter}:{letter}"
    return f"{_PARSE_HOLE_SHEET}!{n}:{n}"


def _skeletonize_formula_addresses(formula: str) -> tuple[str, tuple[str, ...]]:
    """Replace addresses with parseable hole placeholders.

    Returns the punched skeleton and the original address texts in left-to-right
    order, matching AST preorder of those leaves.
    """
    spans = _formula_address_spans(formula)
    if not spans:
        return formula, ()
    holes: list[str] = []
    parts: list[str] = []
    cursor = 0
    for index, (start, end, kind) in enumerate(spans):
        parts.append(formula[cursor:start])
        holes.append(formula[start:end])
        parts.append(_hole_placeholder(kind, index))
        cursor = end
    parts.append(formula[cursor:])
    return "".join(parts), tuple(holes)


def _project_parse_holes(node: AstNode) -> SkeletonNode:
    """Replace `_EG_SHAPE_HOLE` leaves with typed `AddressHoleNode`s."""
    match node:
        case CellRefNode(ref) if ref.sheet == _PARSE_HOLE_SHEET:
            if not isinstance(ref.row, AbsoluteAxis) or ref.row.index < 1:
                raise ValueError("parse hole cell index out of range")
            return AddressHoleNode(kind="CELL", index=ref.row.index - 1)
        case RangeNode(start_ref=start, end_ref=_end) if start.sheet == _PARSE_HOLE_SHEET:
            if not isinstance(start.row, AbsoluteAxis) or start.row.index < 1:
                raise ValueError("parse hole range index out of range")
            return AddressHoleNode(kind="RANGE", index=start.row.index - 1)
        case WholeColumnNode() if node.sheet == _PARSE_HOLE_SHEET:
            if not isinstance(node.col, AbsoluteAxis) or node.col.index < 1:
                raise ValueError("parse hole whole-column index out of range")
            return AddressHoleNode(kind="WHOLE_COL", index=node.col.index - 1)
        case WholeRowNode() if node.sheet == _PARSE_HOLE_SHEET:
            if not isinstance(node.row, AbsoluteAxis) or node.row.index < 1:
                raise ValueError("parse hole whole-row index out of range")
            return AddressHoleNode(kind="WHOLE_ROW", index=node.row.index - 1)
        case FunctionCallNode(name, args):
            return FunctionCallNode(
                name, [cast(AstNode, _project_parse_holes(arg)) for arg in args]
            )
        case BinaryOpNode(op, left, right):
            return BinaryOpNode(
                op,
                cast(AstNode, _project_parse_holes(left)),
                cast(AstNode, _project_parse_holes(right)),
            )
        case UnaryOpNode(op, operand):
            return UnaryOpNode(op, cast(AstNode, _project_parse_holes(operand)))
        case _:
            return node


def _cached_parse_skeleton(skeleton: str) -> SkeletonNode:
    """Parse a punched skeleton and intern it as a `FormulaShape` hole tree."""
    global _SHAPE_PARSE_HITS, _SHAPE_PARSE_MISSES
    cached = _SHAPE_PARSE_CACHE.get(skeleton)
    if cached is not None:
        _SHAPE_PARSE_CACHE.move_to_end(skeleton)
        _SHAPE_PARSE_HITS += 1
        return cached
    parsed = _project_parse_holes(parse(skeleton))
    if len(_SHAPE_PARSE_CACHE) >= _SHAPE_PARSE_CACHE_MAXSIZE:
        _SHAPE_PARSE_CACHE.popitem(last=False)
    _SHAPE_PARSE_CACHE[skeleton] = parsed
    _SHAPE_PARSE_MISSES += 1
    return parsed


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
class FormulaShapeTable:
    """Interned skeletons plus per-node parameter bindings.

    `shapes` maps `shape_key` to one shared skeleton. `bindings` maps a
    `NodeKey` to `(shape_key, params)`. Excel-facing formula text is the derived
    `normalized_formula` view (`render_formula` / `A1_ABSOLUTE`). This table is
    an optional overlay: `Node.formula_ast` is authoritative, and missing
    shapes fall back to AST. The evaluator compiles helpers from it at
    construction; codegen reads it at generate time.
    """

    shapes: dict[str, SkeletonNode]
    bindings: dict[str, tuple[str, tuple[AddressLeaf, ...]]]

    def copy(self) -> FormulaShapeTable:
        """Shallow-copy maps; skeletons and param tuples are shared."""
        return FormulaShapeTable(shapes=dict(self.shapes), bindings=dict(self.bindings))

    def lookup(self, node_key: str) -> tuple[str, SkeletonNode, tuple[AddressLeaf, ...]] | None:
        """Return `(shape_key, skeleton, params)` for `node_key`.

        `node_key` is a graph node address, not formula text.
        """
        binding = self.bindings.get(node_key)
        if binding is None:
            return None
        shape_key, params = binding
        skeleton = self.shapes.get(shape_key)
        if skeleton is None:
            return None
        return shape_key, skeleton, params

    def drop_binding(self, node_key: str) -> None:
        """Remove the parameter binding for `node_key` if present."""
        self.bindings.pop(node_key, None)

    def rebind(self, node_key: str, formula_or_ast: str | AstNode) -> None:
        """Intern `formula_or_ast` as the binding for `node_key`."""
        shape = fingerprint_formula_shape(formula_or_ast)
        self.shapes.setdefault(shape.shape_key, shape.skeleton)
        self.bindings[node_key] = (shape.shape_key, shape.params)

    def prune_unused_shapes(self) -> None:
        """Drop skeletons that no remaining binding references."""
        used = {shape_key for shape_key, _params in self.bindings.values()}
        for shape_key in list(self.shapes):
            if shape_key not in used:
                del self.shapes[shape_key]


def intern_formula_shapes(
    items: Iterable[tuple[str, str | AstNode]],
) -> FormulaShapeTable:
    """Build a shape table from `(node_key, formula_or_ast)` pairs.

    Bindings are keyed by `node_key` (a cell address), not by formula text.
    Each node gets its own binding even when formula text is shared.
    Formulas that share a punched `shape_key` share one skeleton.
    """
    shapes: dict[str, SkeletonNode] = {}
    bindings: dict[str, tuple[str, tuple[AddressLeaf, ...]]] = {}
    for item in items:
        if not isinstance(item, tuple) or len(item) != 2:
            raise TypeError("intern_formula_shapes expects (node_key, formula_or_ast) pairs")
        node_key, formula_or_ast = item
        if node_key in bindings:
            continue
        source: str | AstNode
        if isinstance(formula_or_ast, str):
            stripped = formula_or_ast.strip()
            if not stripped:
                continue
            source = stripped
        else:
            source = formula_or_ast
        shape = fingerprint_formula_shape(source)
        shapes.setdefault(shape.shape_key, shape.skeleton)
        bindings[node_key] = (shape.shape_key, shape.params)
    return FormulaShapeTable(shapes=shapes, bindings=bindings)
