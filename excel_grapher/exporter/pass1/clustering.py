from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path
from typing import TypeAlias

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
from excel_grapher.exporter import ProjectionResult
from excel_grapher.exporter.pass1.addresses import parse_workbook_address
from excel_grapher.exporter.pass1.bindings import BindingKeyValue, expected_keys_for_address
from excel_grapher.exporter.pass1.modes import ClusteringMode, VariationMode
from excel_grapher.grapher.graph import DependencyGraph

ClusterableGraph: TypeAlias = DependencyGraph | ProjectionResult

StructuralFingerprint: TypeAlias = tuple[tuple, tuple[str, ...]]

BoundAddressKeys: TypeAlias = Mapping[str, Mapping[str, BindingKeyValue]]
AddressToSeriesId: TypeAlias = Mapping[str, str]


def _require_bound_address_keys(
    bound_address_keys: BoundAddressKeys | None,
) -> BoundAddressKeys:
    if bound_address_keys is None:
        raise ValueError(
            "bound_address_keys is required for formula clustering; "
            "configure series bindings in bindings/*.bindings.yaml, then build "
            "keys with build_bound_address_keys() from the loaded input, "
            "output, internal, and constant series"
        )
    return bound_address_keys


def _require_address_to_series_id(
    address_to_series_id: AddressToSeriesId | None,
    *,
    clustering_mode: ClusteringMode,
) -> AddressToSeriesId:
    if clustering_mode == "ast":
        return address_to_series_id or {}
    if address_to_series_id is None:
        raise ValueError(
            "address_to_series_id is required for formula clustering when "
            f"clustering_mode={clustering_mode!r}; build it with "
            "build_address_to_series_id() from derived internal series, "
            "falling back to constant then public output/input series"
        )
    return address_to_series_id


def _partition_members_by_series(
    members: tuple[str, ...],
    address_to_series_id: AddressToSeriesId,
) -> tuple[tuple[str, ...], ...]:
    if len(members) < 2:
        return (members,)

    grouped: dict[str, list[str]] = {}
    unowned: list[str] = []
    for address in members:
        series_id = address_to_series_id.get(address)
        if series_id is None:
            unowned.append(address)
        else:
            grouped.setdefault(series_id, []).append(address)

    partitions: list[tuple[str, ...]] = []
    for series_members in grouped.values():
        partitions.append(tuple(sorted(series_members)))
    partitions.extend((address,) for address in sorted(unowned))
    return tuple(partitions)


def _cluster_members_by_series_only(
    formula_nodes: Mapping[str, str],
    address_to_series_id: AddressToSeriesId,
) -> list[tuple[str, ...]]:
    grouped: dict[str, list[str]] = {}
    unowned: list[str] = []
    for address in sorted(formula_nodes):
        series_id = address_to_series_id.get(address)
        if series_id is None:
            unowned.append(address)
        else:
            grouped.setdefault(series_id, []).append(address)

    raw_clusters: list[tuple[str, ...]] = [tuple(sorted(members)) for members in grouped.values()]
    raw_clusters.extend((address,) for address in sorted(unowned))
    return raw_clusters


def _apply_variation_mode_splits(
    members: tuple[str, ...],
    formula_nodes: Mapping[str, str],
    bound_address_keys: BoundAddressKeys,
    *,
    variation_mode: VariationMode,
    workbook_path: Path | None,
    key_cache: _ClusteringKeyCache | None,
) -> list[tuple[str, ...]]:
    ordered_members = tuple(sorted(members))
    if variation_mode == "dominant_key_only" and len(ordered_members) >= 2:
        return list(
            _split_cluster_by_dominant_keys(
                ordered_members,
                formula_nodes,
                bound_address_keys,
                workbook_path=workbook_path,
                key_cache=key_cache,
            )
        )
    return [ordered_members]


@dataclass
class _ClusteringKeyCache:
    bound_address_keys: BoundAddressKeys
    workbook_path: Path | None = None
    _concept_cache: dict[str, tuple[str, ...] | None] = field(default_factory=dict, repr=False)
    _value_cache: dict[str, dict[str, BindingKeyValue] | None] = field(
        default_factory=dict, repr=False
    )
    _fingerprint_cache: dict[str, StructuralFingerprint | None] = field(
        default_factory=dict, repr=False
    )

    def warm_from_formula_nodes(self, formula_nodes: Mapping[str, str]) -> None:
        for address, formula in formula_nodes.items():
            self.fingerprint_for_formula(address, formula)

    def fingerprint_for_formula(self, address: str, formula: str) -> StructuralFingerprint | None:
        if address not in self._fingerprint_cache:
            self._fingerprint_cache[address] = _structural_fingerprint(
                formula,
                bound_address_keys=self.bound_address_keys,
                workbook_path=self.workbook_path,
                key_cache=self,
            )
        return self._fingerprint_cache[address]

    def concepts_for_address(self, address: str) -> tuple[str, ...] | None:
        if address not in self._concept_cache:
            self._concept_cache[address] = self._resolve_concepts(address)
        return self._concept_cache[address]

    def values_for_address(self, address: str) -> dict[str, BindingKeyValue] | None:
        if address not in self._value_cache:
            self._value_cache[address] = self._resolve_values(address)
        return self._value_cache[address]

    def _resolve_concepts(self, address: str) -> tuple[str, ...] | None:
        if self.workbook_path is not None:
            keys = expected_keys_for_address(
                address,
                bound_address_keys=self.bound_address_keys,
            )
            return tuple(sorted(keys))
        raw_keys = self.bound_address_keys.get(address)
        if raw_keys is None:
            return None
        return tuple(sorted(raw_keys))

    def _resolve_values(self, address: str) -> dict[str, BindingKeyValue] | None:
        if self.workbook_path is not None:
            return expected_keys_for_address(
                address,
                bound_address_keys=self.bound_address_keys,
            )
        raw_keys = self.bound_address_keys.get(address)
        if raw_keys is None:
            return None
        return dict(raw_keys)


@dataclass(frozen=True)
class FormulaCluster:
    """A group of workbook cells whose normalized formulas share one AST shape."""

    cluster_id: int
    members: tuple[str, ...]
    canonical_template: str
    row: int | None


def _projected_graph(graph: ClusterableGraph) -> DependencyGraph:
    if isinstance(graph, ProjectionResult):
        return graph.projected_graph
    return graph


def _formula_nodes(graph: ClusterableGraph) -> dict[str, str]:
    projected = _projected_graph(graph)
    nodes: dict[str, str] = {}
    for address in projected:
        node = projected.get_node(address)
        # Include literal-only formulas (``=1+1``); those are ``is_leaf`` in the
        # graph sense but still Pass-1 helper codegen units when bound.
        if node is None or not node.has_formula or node.normalized_formula is None:
            continue
        nodes[address] = node.normalized_formula
    return nodes


def _formula_body(normalized_formula: str) -> str:
    return normalized_formula.removeprefix("=")


def _binding_key_concepts_for_address(
    address: str,
    bound_address_keys: BoundAddressKeys,
    *,
    workbook_path: Path | None = None,
    key_cache: _ClusteringKeyCache | None = None,
) -> tuple[str, ...] | None:
    if key_cache is not None:
        return key_cache.concepts_for_address(address)
    if workbook_path is not None:
        keys = expected_keys_for_address(
            address,
            bound_address_keys=bound_address_keys,
        )
        return tuple(sorted(keys))
    keys = bound_address_keys.get(address)
    if keys is None:
        return None
    return tuple(sorted(keys))


def _address_only_structural_tuple(node: AstNode, refs: list[str]) -> tuple:
    """Build a binding-agnostic skeleton for tests and diagnostics."""
    if isinstance(node, NumberNode):
        return ("num", node.value)
    if isinstance(node, StringNode):
        return ("str", node.value)
    if isinstance(node, BoolNode):
        return ("bool", node.value)
    if isinstance(node, ErrorNode):
        return ("err", str(node.error))
    if isinstance(node, CellRefNode):
        if node.address not in refs:
            refs.append(node.address)
        ref_index = refs.index(node.address)
        return ("ref", ref_index)
    if isinstance(node, RangeNode):
        for address in (node.start, node.end):
            if address not in refs:
                refs.append(address)
        start_index = refs.index(node.start)
        end_index = refs.index(node.end)
        return ("range", start_index, end_index)
    if isinstance(node, WholeColumnNode):
        return ("wcol", node.sheet, node.column)
    if isinstance(node, WholeRowNode):
        return ("wrow", node.sheet, node.row)
    if isinstance(node, EmptyArgNode):
        return ("empty",)
    if isinstance(node, UnaryOpNode):
        return (
            "unary",
            node.op,
            _address_only_structural_tuple(node.operand, refs),
        )
    if isinstance(node, BinaryOpNode):
        return (
            "bin",
            node.op,
            _address_only_structural_tuple(node.left, refs),
            _address_only_structural_tuple(node.right, refs),
        )
    if isinstance(node, FunctionCallNode):
        return (
            "fn",
            node.name,
            tuple(_address_only_structural_tuple(arg, refs) for arg in node.args),
        )
    raise TypeError(type(node))


def _structural_tuple(
    node: AstNode,
    refs: list[str],
    *,
    bound_address_keys: BoundAddressKeys,
    workbook_path: Path | None = None,
    key_cache: _ClusteringKeyCache | None = None,
) -> tuple:
    if isinstance(node, NumberNode):
        return ("num", node.value)
    if isinstance(node, StringNode):
        return ("str", node.value)
    if isinstance(node, BoolNode):
        return ("bool", node.value)
    if isinstance(node, ErrorNode):
        return ("err", str(node.error))
    if isinstance(node, CellRefNode):
        if node.address not in refs:
            refs.append(node.address)
        ref_index = refs.index(node.address)
        key_concepts = _binding_key_concepts_for_address(
            node.address,
            bound_address_keys,
            workbook_path=workbook_path,
            key_cache=key_cache,
        )
        if key_concepts is None:
            return ("ref", ref_index, None)
        return ("ref", ref_index, key_concepts)
    if isinstance(node, RangeNode):
        for address in (node.start, node.end):
            if address not in refs:
                refs.append(address)
        start_index = refs.index(node.start)
        end_index = refs.index(node.end)
        start_keys = _binding_key_concepts_for_address(
            node.start,
            bound_address_keys,
            workbook_path=workbook_path,
            key_cache=key_cache,
        )
        end_keys = _binding_key_concepts_for_address(
            node.end,
            bound_address_keys,
            workbook_path=workbook_path,
            key_cache=key_cache,
        )
        if start_keys is None or end_keys is None:
            return ("range", start_index, end_index, None, None)
        return ("range", start_index, end_index, start_keys, end_keys)
    if isinstance(node, WholeColumnNode):
        return ("wcol", node.sheet, node.column)
    if isinstance(node, WholeRowNode):
        return ("wrow", node.sheet, node.row)
    if isinstance(node, EmptyArgNode):
        return ("empty",)
    if isinstance(node, UnaryOpNode):
        return (
            "unary",
            node.op,
            _structural_tuple(
                node.operand,
                refs,
                bound_address_keys=bound_address_keys,
                workbook_path=workbook_path,
                key_cache=key_cache,
            ),
        )
    if isinstance(node, BinaryOpNode):
        return (
            "bin",
            node.op,
            _structural_tuple(
                node.left,
                refs,
                bound_address_keys=bound_address_keys,
                workbook_path=workbook_path,
                key_cache=key_cache,
            ),
            _structural_tuple(
                node.right,
                refs,
                bound_address_keys=bound_address_keys,
                workbook_path=workbook_path,
                key_cache=key_cache,
            ),
        )
    if isinstance(node, FunctionCallNode):
        return (
            "fn",
            node.name,
            tuple(
                _structural_tuple(
                    arg,
                    refs,
                    bound_address_keys=bound_address_keys,
                    workbook_path=workbook_path,
                    key_cache=key_cache,
                )
                for arg in node.args
            ),
        )
    raise TypeError(type(node))


def _structural_fingerprint(
    normalized_formula: str,
    *,
    bound_address_keys: BoundAddressKeys,
    workbook_path: Path | None = None,
    key_cache: _ClusteringKeyCache | None = None,
) -> StructuralFingerprint | None:
    try:
        ast = parse(_formula_body(normalized_formula))
    except FormulaParseError:
        return None
    refs: list[str] = []
    return (
        _structural_tuple(
            ast,
            refs,
            bound_address_keys=bound_address_keys,
            workbook_path=workbook_path,
            key_cache=key_cache,
        ),
        tuple(refs),
    )


def address_only_structural_fingerprint(
    normalized_formula: str,
) -> StructuralFingerprint | None:
    """Return a binding-agnostic ``(skeleton, refs)`` for tests and diagnostics."""
    try:
        ast = parse(_formula_body(normalized_formula))
    except FormulaParseError:
        return None
    refs: list[str] = []
    return (_address_only_structural_tuple(ast, refs), tuple(refs))


def structural_fingerprint(
    normalized_formula: str,
    *,
    bound_address_keys: BoundAddressKeys | None,
    workbook_path: Path | None = None,
    key_cache: _ClusteringKeyCache | None = None,
) -> StructuralFingerprint | None:
    """Return ``(skeleton, refs)`` with binding-aware ref placeholders."""
    resolved_bound_keys = _require_bound_address_keys(bound_address_keys)
    return _structural_fingerprint(
        normalized_formula,
        bound_address_keys=resolved_bound_keys,
        workbook_path=workbook_path,
        key_cache=key_cache,
    )


_BINARY_PRECEDENCE: dict[str, int] = {
    "^": 6,
    "*": 5,
    "/": 5,
    "+": 4,
    "-": 4,
    "&": 3,
    "=": 2,
    "<>": 2,
    "<": 2,
    ">": 2,
    "<=": 2,
    ">=": 2,
}


def _format_excel_number(value: object) -> str:
    if isinstance(value, bool):
        raise TypeError("bool is not a number literal")
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return format(value, "g")
    raise TypeError(f"expected int or float number literal, got {type(value)!r}")


def _format_excel_string(value: object) -> str:
    if not isinstance(value, str):
        raise TypeError(f"expected str string literal, got {type(value)!r}")
    text = value.replace('"', '""')
    return f'"{text}"'


def _format_excel_bool(value: object) -> str:
    if value is True:
        return "TRUE"
    if value is False:
        return "FALSE"
    raise TypeError(f"expected bool literal, got {type(value)!r}")


def _skeleton_literal_value(node: tuple, *, kind: str) -> object:
    if len(node) < 2:
        raise TypeError(f"{kind} skeleton node requires a value")
    return node[1]


def _format_ref_placeholder(index: object, concepts: object | None = None) -> str:
    label = f"ref_{index}"
    if concepts is None:
        return label
    if not isinstance(concepts, tuple):
        return label
    return f"{label}[{','.join(str(concept) for concept in concepts)}]"


def _format_skeleton_node(node: tuple, *, min_prec: int) -> str:
    kind = node[0]
    if kind == "num":
        return _format_excel_number(_skeleton_literal_value(node, kind=kind))
    if kind == "str":
        return _format_excel_string(_skeleton_literal_value(node, kind=kind))
    if kind == "bool":
        return _format_excel_bool(_skeleton_literal_value(node, kind=kind))
    if kind == "empty":
        return ""
    if kind == "err":
        return str(_skeleton_literal_value(node, kind=kind))
    if kind == "ref":
        if len(node) == 2:
            return _format_ref_placeholder(node[1])
        return _format_ref_placeholder(node[1], node[2])
    if kind == "range":
        if len(node) == 3:
            start = _format_ref_placeholder(node[1])
            end = _format_ref_placeholder(node[2])
        else:
            start = _format_ref_placeholder(node[1], node[3])
            end = _format_ref_placeholder(node[2], node[4])
        return f"{start}:{end}"
    if kind == "wcol":
        sheet, column = node[1], node[2]
        return f"{sheet}!{column}:{column}"
    if kind == "wrow":
        sheet, row = node[1], node[2]
        return f"{sheet}!{row}:{row}"
    if kind == "unary":
        op = str(node[1])
        operand = _format_skeleton_node(node[2], min_prec=7)
        return f"{op}{operand}"
    if kind == "bin":
        op = str(node[1])
        prec = _BINARY_PRECEDENCE.get(op, 1)
        left = _format_skeleton_node(node[2], min_prec=prec)
        # Left-associative: parenthesize right side on equal precedence for
        # non-associative-looking ops (-, /, comparisons).
        right_prec = prec + (0 if op in {"*", "+", "&", "^"} else 1)
        right = _format_skeleton_node(node[3], min_prec=right_prec)
        rendered = f"{left}{op}{right}"
        if prec < min_prec:
            return f"({rendered})"
        return rendered
    if kind == "fn":
        name = str(node[1])
        args = ",".join(_format_skeleton_node(arg, min_prec=0) for arg in node[2])
        return f"{name}({args})"
    raise TypeError(f"unknown skeleton node kind: {kind!r}")


def format_structural_skeleton(skeleton: tuple) -> str:
    """Render a structural skeleton as an Excel-like formula with ref placeholders.

    Cell/range slots become ``ref_N`` or ``ref_N[DIM,...]`` using the fingerprint's
    sorted binding dimension ids. Scalar literals render as Excel number, string
    (``"..."``), or boolean (``TRUE``/``FALSE``) tokens; empty args render blank.
    """
    return f"={_format_skeleton_node(skeleton, min_prec=0)}"


def _binding_aware_fingerprint_complete(
    fingerprint: StructuralFingerprint,
) -> bool:
    skeleton, _refs = fingerprint

    def walk(node: tuple) -> bool:
        if not node:
            return True
        kind = node[0]
        if kind == "ref":
            return len(node) == 3 and node[2] is not None
        if kind == "range":
            return len(node) == 5 and node[3] is not None and node[4] is not None
        if kind in {"num", "str", "bool", "empty"}:
            return True
        if kind == "err":
            return True
        if kind in {"wcol", "wrow"}:
            return True
        if kind == "unary":
            return walk(node[2])
        if kind == "bin":
            return walk(node[2]) and walk(node[3])
        if kind == "fn":
            return all(walk(arg) for arg in node[2])
        return False

    return walk(skeleton)


def _formulas_are_parameterizable(
    left_formula: str,
    right_formula: str,
    *,
    bound_address_keys: BoundAddressKeys,
    workbook_path: Path | None = None,
    key_cache: _ClusteringKeyCache | None = None,
) -> bool:
    """Return whether two normalized formulas belong in the same parameterizable bucket."""
    left_fingerprint = _structural_fingerprint(
        left_formula,
        bound_address_keys=bound_address_keys,
        workbook_path=workbook_path,
        key_cache=key_cache,
    )
    right_fingerprint = _structural_fingerprint(
        right_formula,
        bound_address_keys=bound_address_keys,
        workbook_path=workbook_path,
        key_cache=key_cache,
    )
    if left_fingerprint is None or right_fingerprint is None:
        return False
    left_skeleton, left_refs = left_fingerprint
    right_skeleton, right_refs = right_fingerprint
    if left_skeleton != right_skeleton:
        return False

    if not _binding_aware_fingerprint_complete(
        left_fingerprint
    ) or not _binding_aware_fingerprint_complete(right_fingerprint):
        return False
    return len(left_refs) == len(right_refs)


def formulas_are_parameterizable(
    left_formula: str,
    right_formula: str,
    *,
    bound_address_keys: BoundAddressKeys | None,
    workbook_path: Path | None = None,
    key_cache: _ClusteringKeyCache | None = None,
) -> bool:
    """Return whether two normalized formulas belong in the same parameterizable bucket."""
    resolved_bound_keys = _require_bound_address_keys(bound_address_keys)
    return _formulas_are_parameterizable(
        left_formula,
        right_formula,
        bound_address_keys=resolved_bound_keys,
        workbook_path=workbook_path,
        key_cache=key_cache,
    )


def _clustering_bucket_key(
    fingerprint: StructuralFingerprint | None,
    *,
    address: str,
) -> tuple[object, ...]:
    """Return the bucket key for one formula's structural fingerprint."""
    if fingerprint is None:
        return ("singleton", address)
    if not _binding_aware_fingerprint_complete(fingerprint):
        return ("singleton", address)
    skeleton, refs = fingerprint
    return ("cluster", skeleton, len(refs))


def _ref_position_key_values(
    member_address: str,
    formula: str,
    bound_address_keys: BoundAddressKeys,
    *,
    workbook_path: Path | None = None,
    key_cache: _ClusteringKeyCache | None = None,
) -> list[dict[str, BindingKeyValue]] | None:
    if key_cache is not None:
        fingerprint = key_cache.fingerprint_for_formula(member_address, formula)
    else:
        fingerprint = _structural_fingerprint(
            formula,
            bound_address_keys=bound_address_keys,
            workbook_path=workbook_path,
            key_cache=None,
        )
    if fingerprint is None or not _binding_aware_fingerprint_complete(fingerprint):
        return None
    _skeleton, refs = fingerprint
    values_by_ref: list[dict[str, BindingKeyValue]] = []
    for ref_address in refs:
        if key_cache is not None:
            keys = key_cache.values_for_address(ref_address)
            if keys is None:
                return None
            values_by_ref.append(keys)
            continue
        if workbook_path is not None:
            keys = expected_keys_for_address(
                ref_address,
                bound_address_keys=bound_address_keys,
            )
        else:
            raw_keys = bound_address_keys.get(ref_address)
            if raw_keys is None:
                return None
            keys = dict(raw_keys)
        values_by_ref.append(keys)
    return values_by_ref


def _varying_concepts_at_ref_from_matrix(
    member_matrices: Mapping[str, list[dict[str, BindingKeyValue]] | None],
    members: tuple[str, ...],
    ref_index: int,
) -> frozenset[str]:
    concept_values: dict[str, set[BindingKeyValue]] = {}
    for member in members:
        ref_values = member_matrices[member]
        if ref_values is None or ref_index >= len(ref_values):
            return frozenset()
        for concept, value in ref_values[ref_index].items():
            concept_values.setdefault(concept, set()).add(value)
    return frozenset(concept for concept, values in concept_values.items() if len(values) > 1)


def _dominant_key_at_ref_from_matrix(
    member_matrices: Mapping[str, list[dict[str, BindingKeyValue]] | None],
    members: tuple[str, ...],
    ref_index: int,
    varying_concepts: frozenset[str],
) -> str | None:
    if not varying_concepts:
        return None
    counts: dict[str, int] = {}
    for concept in varying_concepts:
        values: set[BindingKeyValue] = set()
        for member in members:
            ref_values = member_matrices[member]
            if ref_values is None or ref_index >= len(ref_values):
                continue
            if concept in ref_values[ref_index]:
                values.add(ref_values[ref_index][concept])
        counts[concept] = len(values)
    return max(counts, key=lambda concept: (counts[concept], concept))


def _split_signature_from_ref_keys(
    keys_at_ref: dict[str, BindingKeyValue],
    dominant_key: str,
    varying_concepts: frozenset[str],
) -> tuple[BindingKeyValue, ...] | None:
    signature: list[BindingKeyValue] = []
    for concept in sorted(varying_concepts):
        if concept == dominant_key:
            continue
        if concept not in keys_at_ref:
            return None
        signature.append(keys_at_ref[concept])
    return tuple(signature)


def _split_cluster_by_dominant_keys(
    members: tuple[str, ...],
    formula_nodes: Mapping[str, str],
    bound_address_keys: BoundAddressKeys,
    *,
    workbook_path: Path | None = None,
    key_cache: _ClusteringKeyCache | None = None,
) -> tuple[tuple[str, ...], ...]:
    if len(members) < 2:
        return (members,)

    canonical_formula = formula_nodes[members[0]]
    if key_cache is not None:
        fingerprint = key_cache.fingerprint_for_formula(members[0], canonical_formula)
    else:
        fingerprint = _structural_fingerprint(
            canonical_formula,
            bound_address_keys=bound_address_keys,
            workbook_path=workbook_path,
            key_cache=None,
        )
    if fingerprint is None:
        return (members,)
    _skeleton, refs = fingerprint
    if not refs:
        return (members,)

    member_matrices = {
        member: _ref_position_key_values(
            member,
            formula_nodes[member],
            bound_address_keys,
            workbook_path=workbook_path,
            key_cache=key_cache,
        )
        for member in members
    }

    ref_dominance: list[tuple[frozenset[str], str | None]] = []
    for ref_index in range(len(refs)):
        varying = _varying_concepts_at_ref_from_matrix(member_matrices, members, ref_index)
        dominant = _dominant_key_at_ref_from_matrix(member_matrices, members, ref_index, varying)
        ref_dominance.append((varying, dominant))

    groups: dict[tuple[tuple[BindingKeyValue, ...], ...], list[str]] = {}
    for member in members:
        ref_values = member_matrices[member]
        signatures: list[tuple[BindingKeyValue, ...]] = []
        for ref_index, (varying, dominant) in enumerate(ref_dominance):
            if dominant is None or ref_values is None or ref_index >= len(ref_values):
                signatures.append(())
                continue
            signature = _split_signature_from_ref_keys(
                ref_values[ref_index],
                dominant,
                varying,
            )
            if signature is None:
                signatures.append(())
            else:
                signatures.append(signature)
        groups.setdefault(tuple(signatures), []).append(member)

    return tuple(
        tuple(sorted(group_members))
        for _signature, group_members in sorted(groups.items(), key=lambda item: item[1])
    )


def formula_nodes_for_clustering(graph: ClusterableGraph) -> dict[str, str]:
    """Return non-leaf formula addresses and normalized formulas from a clusterable graph."""
    return _formula_nodes(graph)


def cluster_has_independent_operand_variation(
    cluster: FormulaCluster,
    formula_nodes: Mapping[str, str],
    bound_address_keys: BoundAddressKeys,
    varying_concepts: frozenset[str],
    *,
    workbook_path: Path | None = None,
    key_cache: _ClusteringKeyCache | None = None,
) -> bool:
    """Return whether a cluster needs unsupported per-operand parameterization."""
    if not varying_concepts:
        return False

    for concept in varying_concepts:
        operand_patterns: list[tuple[BindingKeyValue, ...]] = []
        for member in cluster.members:
            ref_values = _ref_position_key_values(
                member,
                formula_nodes[member],
                bound_address_keys,
                workbook_path=workbook_path,
                key_cache=key_cache,
            )
            if ref_values is None:
                continue
            pattern = tuple(ref_keys[concept] for ref_keys in ref_values if concept in ref_keys)
            if pattern:
                operand_patterns.append(pattern)

        if not operand_patterns:
            continue

        if (
            len(set(operand_patterns)) == 1
            and len(operand_patterns[0]) > 1
            and len(set(operand_patterns[0])) > 1
        ):
            return True

        if len(set(operand_patterns)) > 1 and all(
            len(pattern) > 1 and len(set(pattern)) == len(pattern) for pattern in operand_patterns
        ):
            return True
    return False


def cluster_graph_formulas(
    graph: ClusterableGraph,
    *,
    bound_address_keys: BoundAddressKeys | None,
    variation_mode: VariationMode = "independent",
    clustering_mode: ClusteringMode = "series_ast",
    address_to_series_id: AddressToSeriesId | None = None,
    workbook_path: Path | None = None,
) -> tuple[FormulaCluster, ...]:
    """Cluster non-leaf formula nodes for internals refactor.

    ``clustering_mode`` selects the base grouping strategy:

    - ``series``: one unit per partition series id (internal first, else public
      output/input binding series; no cross-series merging).
    - ``series_ast``: AST-cluster, partition each cluster by owning series id
      (internal first, else public output/input), then apply ``variation_mode``
      within each series partition.
    - ``ast``: AST (+ keys) only; series-blind (legacy behavior).

    Missing internal ownership is not the same as an intended singleton: public
    output (and input override) series ids keep multi-member time sweeps together.
    """
    resolved_bound_keys = _require_bound_address_keys(bound_address_keys)
    resolved_series_ids = _require_address_to_series_id(
        address_to_series_id,
        clustering_mode=clustering_mode,
    )
    formula_nodes = _formula_nodes(graph)

    key_cache: _ClusteringKeyCache | None = None
    if clustering_mode != "series":
        key_cache = _ClusteringKeyCache(
            bound_address_keys=resolved_bound_keys,
            workbook_path=workbook_path,
        )
        key_cache.warm_from_formula_nodes(formula_nodes)

    raw_clusters: list[tuple[str, ...]] = []
    if clustering_mode == "series":
        raw_clusters.extend(_cluster_members_by_series_only(formula_nodes, resolved_series_ids))
    else:
        assert key_cache is not None
        grouped: dict[tuple[object, ...], list[str]] = {}
        for address in sorted(formula_nodes):
            fingerprint = key_cache.fingerprint_for_formula(address, formula_nodes[address])
            bucket_key = _clustering_bucket_key(fingerprint, address=address)
            grouped.setdefault(bucket_key, []).append(address)

        for _bucket_key, members in sorted(grouped.items(), key=lambda item: item[1]):
            ordered_members = tuple(sorted(members))
            member_groups = (
                _partition_members_by_series(ordered_members, resolved_series_ids)
                if clustering_mode == "series_ast"
                else (ordered_members,)
            )
            for member_group in member_groups:
                raw_clusters.extend(
                    _apply_variation_mode_splits(
                        member_group,
                        formula_nodes,
                        resolved_bound_keys,
                        variation_mode=variation_mode,
                        workbook_path=workbook_path,
                        key_cache=key_cache,
                    )
                )

    clusters: list[FormulaCluster] = []
    for cluster_id, ordered_members in enumerate(sorted(raw_clusters, key=lambda members: members)):
        if not ordered_members:
            continue
        template = formula_nodes[ordered_members[0]]
        rows = {parse_workbook_address(address)[2] for address in ordered_members}
        row = next(iter(rows)) if len(rows) == 1 else None
        clusters.append(
            FormulaCluster(
                cluster_id=cluster_id,
                members=ordered_members,
                canonical_template=template,
                row=row,
            )
        )
    return tuple(clusters)
