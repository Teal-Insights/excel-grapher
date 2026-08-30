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

from fastpyxl.utils.cell import get_column_letter

from excel_grapher.core.address_keys import format_range_key, parse_address, quote_sheet_if_needed
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
    RelativeAxis,
    StringNode,
    UnaryOpNode,
    WholeColumnNode,
    WholeRowNode,
    parse,
    resolve_cell_ref,
    resolve_whole_column_ref,
    resolve_whole_row_ref,
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


def iter_address_holes(skeleton: SkeletonNode) -> Iterator[AddressHoleNode]:
    """Yield address holes in preorder."""
    match skeleton:
        case AddressHoleNode():
            yield skeleton
        case FunctionCallNode(_, args):
            for arg in args:
                yield from iter_address_holes(cast(SkeletonNode, arg))
        case BinaryOpNode(_, left, right):
            yield from iter_address_holes(cast(SkeletonNode, left))
            yield from iter_address_holes(cast(SkeletonNode, right))
        case UnaryOpNode(_, operand):
            yield from iter_address_holes(cast(SkeletonNode, operand))
        case _:
            return


def _encode_col_axis(axis: AbsoluteAxis | RelativeAxis) -> str:
    if isinstance(axis, AbsoluteAxis):
        return f"${get_column_letter(axis.index)}"
    return f"C[{axis.offset}]"


def _encode_row_axis(axis: AbsoluteAxis | RelativeAxis) -> str:
    if isinstance(axis, AbsoluteAxis):
        return f"${axis.index}"
    return f"R[{axis.offset}]"


def resolve_address_leaf(leaf: AddressLeaf, anchor: str | None) -> str:
    """Resolve an address leaf to canonical sheet-qualified A1 / range text.

    Relative axes add their offset to `anchor` (the formula host cell).
    Fully-absolute leaves ignore `anchor`.

    Args:
        leaf: Cell, range, whole-column, or whole-row AST leaf.
        anchor: Host cell used to bind relative axes. Ignored when every axis
            on `leaf` is absolute.

    Returns:
        Canonical `Sheet!A1`, `Sheet!A1:B2`, `Sheet!A:A`, or `Sheet!1:1`.

    Raises:
        ValueError: If a relative axis is present and `anchor` is missing.
    """
    match leaf:
        case CellRefNode(ref):
            return resolve_cell_ref(ref, anchor)
        case RangeNode(start_ref, end_ref):
            start = resolve_cell_ref(start_ref, anchor)
            end = resolve_cell_ref(end_ref, anchor)
            sheet, start_cell = parse_address(start)
            _, end_cell = parse_address(end)
            return format_range_key(sheet, start_cell, end_cell)
        case WholeColumnNode():
            sheet, letter = resolve_whole_column_ref(leaf, anchor)
            return format_range_key(sheet, letter, letter)
        case WholeRowNode():
            sheet, row = resolve_whole_row_ref(leaf, anchor)
            return format_range_key(sheet, str(row), str(row))
    raise TypeError(f"not an address leaf: {type(leaf).__name__}")


def encode_address_leaf(leaf: AddressLeaf) -> str:
    """Encode an address leaf as a sheet-qualified A1 / range string.

    Fully-absolute cell refs stay canonical `Sheet!A1`. Relative/mixed refs use
    `$` on absolute axes and `C[n]`/`R[n]` on relative axes. Ranges become
    `Sheet!A1:B2` when both endpoints are fully absolute.
    """
    match leaf:
        case CellRefNode(ref):
            if isinstance(ref.col, AbsoluteAxis) and isinstance(ref.row, AbsoluteAxis):
                return resolve_cell_ref(ref, None)
            return (
                f"{quote_sheet_if_needed(ref.sheet)}!"
                f"{_encode_col_axis(ref.col)}{_encode_row_axis(ref.row)}"
            )
        case RangeNode(start_ref, end_ref):
            start_abs = isinstance(start_ref.col, AbsoluteAxis) and isinstance(
                start_ref.row, AbsoluteAxis
            )
            end_abs = isinstance(end_ref.col, AbsoluteAxis) and isinstance(
                end_ref.row, AbsoluteAxis
            )
            if start_abs and end_abs:
                start = resolve_cell_ref(start_ref, None)
                end = resolve_cell_ref(end_ref, None)
                sheet, start_cell = parse_address(start)
                _, end_cell = parse_address(end)
                return format_range_key(sheet, start_cell, end_cell)
            start_enc = encode_address_leaf(CellRefNode(start_ref))
            end_enc = encode_address_leaf(CellRefNode(end_ref))
            return f"{start_enc}:{end_enc}"
        case WholeColumnNode():
            if isinstance(leaf.col, AbsoluteAxis):
                sheet, letter = resolve_whole_column_ref(leaf, None)
                return format_range_key(sheet, letter, letter)
            return f"{quote_sheet_if_needed(leaf.sheet)}!{_encode_col_axis(leaf.col)}:{_encode_col_axis(leaf.col)}"
        case WholeRowNode():
            if isinstance(leaf.row, AbsoluteAxis):
                sheet, row = resolve_whole_row_ref(leaf, None)
                return format_range_key(sheet, str(row), str(row))
            return (
                f"{quote_sheet_if_needed(leaf.sheet)}!"
                f"{_encode_row_axis(leaf.row)}:{_encode_row_axis(leaf.row)}"
            )
    raise TypeError(f"not an address leaf: {type(leaf).__name__}")


@dataclass(frozen=True, slots=True)
class FormulaShapeTable:
    """Interned skeletons plus per-node parameter bindings.

    `shapes` maps `shape_key` to one shared skeleton. `bindings` maps a
    `NodeKey` to `(shape_key, params)`. Excel-facing formula text is the derived
    `normalized_formula` view (`render_formula` / `A1_ABSOLUTE`); this table is
    an overlay the evaluator and codegen both read.
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

    @property
    def formula_ast(self) -> AstNode | None: ...


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
    """Count distinct formula identities vs punched AST shapes in `graph`.

    Walks `graph.formula_nodes()`, fingerprints each `formula_ast` (falling
    back to `normalized_formula`), and reports whether shapes collapse the
    distinct A1 formula-string set (#517 go/no-go metric).
    """
    from excel_grapher.core.formula_ast_json import ast_identity_key

    shape_counter: Counter[str] = Counter()
    unparseable = 0
    string_idents: set[str] = set()
    parseable = 0

    for _, node in graph.formula_nodes():
        nf = node.normalized_formula
        source: AstNode | str | None = node.formula_ast
        if source is None:
            if isinstance(nf, str) and nf.strip():
                source = nf.strip()
            else:
                continue
        try:
            shape = fingerprint_formula_shape(source)
        except FormulaParseError:
            unparseable += 1
            continue
        parseable += 1
        if isinstance(nf, str) and nf.strip():
            string_idents.add(nf.strip())
        elif not isinstance(source, str):
            string_idents.add(ast_identity_key(source))
        else:
            string_idents.add(source)
        shape_counter[shape.shape_key] += 1

    return FormulaShapeSummary(
        formula_nodes=parseable,
        distinct_normalized_formulas=len(string_idents),
        distinct_shapes=len(shape_counter),
        unparseable=unparseable,
        shape_counts=tuple(shape_counter.most_common()),
    )
