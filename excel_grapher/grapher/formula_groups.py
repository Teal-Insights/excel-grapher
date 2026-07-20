"""Formula-group shape fingerprinting, specialization, and family discovery.

Issue 2: `shape_fingerprint`, `specialize_group`, template-field validation.
Issue 3: `build_group_template`, `iter_formula_families` (detect only; no mutate).
"""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Iterator, Mapping, Sequence
from dataclasses import dataclass
from typing import Literal, Protocol, TypeAlias

from excel_grapher.core.address_keys import (
    CellKey,
    format_range_key,
    parse_address,
    parse_node_key,
    sort_node_keys,
)
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
from excel_grapher.evaluator.errors import ParseError
from excel_grapher.evaluator.parser import parse

AddressLeaf: TypeAlias = CellRefNode | RangeNode | WholeColumnNode | WholeRowNode
SkipReason: TypeAlias = Literal[
    "below_min_size",
    "intra_family_edge",
    "unparseable_formula",
    "kind_mismatch",
]


class _FamilyGraph(Protocol):
    """Minimal graph surface for formula-family discovery."""

    sheet_order: list[str]

    def keys(
        self,
        *,
        order: Literal["insertion", "lexical", "workbook"] = "insertion",
        source: object | None = None,
    ) -> list[str]: ...

    def get_node(self, address: str) -> object | None: ...

    def get_dependencies(self, address: str) -> frozenset[str]: ...


@dataclass(frozen=True, slots=True)
class GroupTemplate:
    """Skeleton + bindings produced from a same-shape formula family."""

    skeleton: AstNode
    member_bindings: Mapping[str, tuple[AddressLeaf, ...]]
    shape_fingerprint: str


@dataclass(frozen=True, slots=True)
class ReadyFamily:
    """A same-shape family ready to coalesce into a formula-group node."""

    fingerprint: str
    members: tuple[str, ...]
    skeleton: AstNode
    member_bindings: Mapping[str, tuple[AddressLeaf, ...]]


@dataclass(frozen=True, slots=True)
class SkippedFamily:
    """A fingerprint bucket that will not be coalesced."""

    fingerprint: str
    members: tuple[str, ...]
    reason: SkipReason


class SpecializeError(ValueError):
    """Raised when skeleton specialization fails (arity or kind mismatch)."""


class TemplateBuildError(ValueError):
    """Raised when a family cannot form a consistent group template."""


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


def collect_address_leaves(ast: AstNode) -> list[AddressLeaf]:
    """Return concrete address leaves in `ast` in fingerprint / specialize order."""
    leaves: list[AddressLeaf] = []

    def _walk(n: AstNode) -> None:
        if isinstance(n, (CellRefNode, RangeNode, WholeColumnNode, WholeRowNode)):
            leaves.append(n)
            return
        if isinstance(n, AddressHoleNode):
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

    _walk(ast)
    return leaves


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


def build_group_template(
    members: Sequence[str],
    formulas: Mapping[str, str],
) -> GroupTemplate:
    """Build a skeleton + bindings from same-shape member formulas.

    Address leaves that are identical across all members are baked into the
    skeleton. Differing leaves become typed `AddressHoleNode`s with per-member
    bindings in walk order.

    Args:
        members: Member cell keys (order preserved for binding keys; prefer
            workbook order from the caller).
        formulas: Mapping of member key → `normalized_formula` string.

    Returns:
        `GroupTemplate` with skeleton, bindings, and shape fingerprint.

    Raises:
        TemplateBuildError: If members/formulas are empty, fingerprints diverge,
            leaf counts diverge, or hole kinds disagree across members.
        KeyError: If a member is missing from `formulas`.
        ParseError: If a formula cannot be parsed.
    """
    if not members:
        raise TemplateBuildError("build_group_template requires at least one member")

    canonical_members: list[str] = []
    for raw in members:
        parsed = parse_node_key(raw)
        if not isinstance(parsed, CellKey):
            raise TemplateBuildError(f"Members must be single cells; got {raw!r}")
        canonical_members.append(str(parsed))

    asts = [parse(formulas[m].strip()) for m in canonical_members]
    fingerprints = [shape_fingerprint(ast) for ast in asts]
    if len(set(fingerprints)) != 1:
        raise TemplateBuildError(
            "members do not share a shape fingerprint: "
            + ", ".join(
                f"{m}={fp!r}" for m, fp in zip(canonical_members, fingerprints, strict=True)
            )
        )
    fp = fingerprints[0]

    leaves_by_member = [collect_address_leaves(ast) for ast in asts]
    leaf_counts = {len(leaves) for leaves in leaves_by_member}
    if len(leaf_counts) != 1:
        raise TemplateBuildError(
            f"address-leaf arity mismatch across members: {sorted(leaf_counts)}"
        )
    leaf_count = next(iter(leaf_counts))

    bake_flags: list[bool] = []
    hole_kinds: list[AddressLeafKind] = []
    for i in range(leaf_count):
        column = [leaves[i] for leaves in leaves_by_member]
        kinds = {address_leaf_kind(leaf) for leaf in column}
        if len(kinds) != 1:
            raise TemplateBuildError(
                f"kind mismatch at address-leaf index {i}: {sorted(k.value for k in kinds)}"
            )
        kind = next(iter(kinds))
        if all(leaf == column[0] for leaf in column[1:]):
            bake_flags.append(True)
            hole_kinds.append(kind)
        else:
            bake_flags.append(False)
            hole_kinds.append(kind)

    hole_slot = 0

    def _rebuild(node: AstNode, leaf_index: list[int]) -> AstNode:
        nonlocal hole_slot
        if isinstance(node, (CellRefNode, RangeNode, WholeColumnNode, WholeRowNode)):
            i = leaf_index[0]
            leaf_index[0] = i + 1
            if bake_flags[i]:
                return node
            hole = AddressHoleNode(kind=hole_kinds[i], slot=hole_slot)
            hole_slot += 1
            return hole
        if isinstance(node, FunctionCallNode):
            return FunctionCallNode(
                name=node.name,
                args=[_rebuild(a, leaf_index) for a in node.args],
            )
        if isinstance(node, BinaryOpNode):
            return BinaryOpNode(
                op=node.op,
                left=_rebuild(node.left, leaf_index),
                right=_rebuild(node.right, leaf_index),
            )
        if isinstance(node, UnaryOpNode):
            return UnaryOpNode(op=node.op, operand=_rebuild(node.operand, leaf_index))
        return node

    # Rebuild from the first member's AST (structure matches all by fingerprint).
    skeleton = _rebuild(asts[0], [0])
    if shape_fingerprint(skeleton) != fp:
        raise TemplateBuildError(
            "internal error: skeleton fingerprint diverged from member fingerprint"
        )

    bindings: dict[str, tuple[AddressLeaf, ...]] = {}
    for member, leaves in zip(canonical_members, leaves_by_member, strict=True):
        member_bindings = tuple(leaf for i, leaf in enumerate(leaves) if not bake_flags[i])
        bindings[member] = member_bindings
        try:
            specialize_group(skeleton, member_bindings)
        except SpecializeError as exc:
            raise TemplateBuildError(f"kind_mismatch for {member}: {exc}") from exc

    return GroupTemplate(
        skeleton=skeleton,
        member_bindings=bindings,
        shape_fingerprint=fp,
    )


def iter_formula_families(
    graph: _FamilyGraph,
    *,
    min_family_size: int = 2,
) -> Iterator[ReadyFamily | SkippedFamily]:
    """Discover same-shape formula families on a cell-only (or mixed) graph.

    Existing multi-cell group nodes are ignored. Unparseable formula cells are
    omitted from clusters. Families below `min_family_size`, with intra-family
    edges, or with template kind mismatches are yielded as `SkippedFamily`.

    Yields families in lexicographic fingerprint order; members within a family
    are workbook-ordered via `graph.sheet_order`.
    """
    if min_family_size < 1:
        raise ValueError("min_family_size must be >= 1")

    buckets: dict[str, list[str]] = defaultdict(list)
    formulas: dict[str, str] = {}

    for key in graph.keys(order="workbook"):
        node = graph.get_node(key)
        if node is None:
            continue
        kind = getattr(node, "kind", None)
        # Compare by value to avoid importing NodeKind (circular with node.py).
        if getattr(kind, "value", kind) != "cell":
            continue
        try:
            parsed = parse_node_key(key)
        except ValueError:
            continue
        if not isinstance(parsed, CellKey):
            continue
        nf = getattr(node, "normalized_formula", None)
        if not isinstance(nf, str) or not nf.strip():
            continue
        try:
            ast = parse(nf.strip())
        except ParseError:
            continue
        fp = shape_fingerprint(ast)
        member = str(parsed)
        buckets[fp].append(member)
        formulas[member] = nf.strip()

    sheet_order = list(getattr(graph, "sheet_order", None) or [])
    for fp in sorted(buckets):
        members = tuple(sort_node_keys(buckets[fp], sheet_order=sheet_order))
        if len(members) < min_family_size:
            yield SkippedFamily(fingerprint=fp, members=members, reason="below_min_size")
            continue
        if _has_intra_family_edge(graph, members):
            yield SkippedFamily(fingerprint=fp, members=members, reason="intra_family_edge")
            continue
        try:
            template = build_group_template(members, formulas)
        except TemplateBuildError:
            yield SkippedFamily(fingerprint=fp, members=members, reason="kind_mismatch")
            continue
        yield ReadyFamily(
            fingerprint=template.shape_fingerprint,
            members=members,
            skeleton=template.skeleton,
            member_bindings=template.member_bindings,
        )


def _has_intra_family_edge(graph: _FamilyGraph, members: Sequence[str]) -> bool:
    member_set = set(members)
    for member in members:
        for dep in graph.get_dependencies(member):
            if dep in member_set:
                return True
    return False


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
