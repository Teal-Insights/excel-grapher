from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field, replace
from typing import Any

from excel_grapher.core.address_keys import NodeShape
from excel_grapher.core.formula_ast import (
    CellRefNode,
    FormulaParseError,
    UnaryOpNode,
    parse,
)

from .dependency_provenance import DependencyCause, EdgeProvenance
from .graph import DependencyGraph
from .node import NodeKey

_singleton_ref_address: dict[str, str] = {}
_singleton_ref_negative: set[str] = set()


def clear_identity_singleton_ref_cache() -> None:
    """Drop parse cache used by `is_identity_transit` (call around compression passes)."""
    _singleton_ref_address.clear()
    _singleton_ref_negative.clear()


def _singleton_cell_ref_address(normalized_formula: str) -> str | None:
    """Return the normalized address for a singleton cell-reference formula.

    When `normalized_formula` is a single unary-plus-stripped cell reference,
    return its normalized address; otherwise return None. Results are memoized
    for the current process until `clear_identity_singleton_ref_cache`.
    """
    if normalized_formula in _singleton_ref_negative:
        return None
    hit = _singleton_ref_address.get(normalized_formula)
    if hit is not None:
        return hit
    try:
        ast = parse(normalized_formula)
    except FormulaParseError:
        _singleton_ref_negative.add(normalized_formula)
        return None
    while isinstance(ast, UnaryOpNode) and ast.op == "+":
        ast = ast.operand
    if not isinstance(ast, CellRefNode):
        _singleton_ref_negative.add(normalized_formula)
        return None
    _singleton_ref_address[normalized_formula] = ast.address
    return ast.address


def is_identity_transit(graph: DependencyGraph, transit_key: NodeKey) -> NodeKey | None:
    """Return the sole dependency key for an identity transit node.

    When `transit_key` is a pure identity reference to exactly one dependency,
    return that dependency's key; otherwise return None.
    """
    node = graph.get_node(transit_key)
    if node is None or node.is_leaf or not node.normalized_formula:
        return None
    deps = graph.get_dependencies(transit_key)
    if len(deps) != 1:
        return None
    r_key = next(iter(deps))
    if graph.get_edge_guard(transit_key, r_key) is not None:
        return None
    addr = _singleton_cell_ref_address(node.normalized_formula)
    if addr is None:
        return None
    r_node = graph.get_node(r_key)
    if r_node is None:
        return None
    if addr != r_node.key:
        return None
    return r_key


def replace_substrings_at_spans(
    formula: str, spans: tuple[tuple[int, int], ...], replacement: str
) -> str:
    """Replace each `[a,b)` span in `formula` with `replacement` (right-to-left)."""
    out = formula
    for a, b in sorted(spans, reverse=True):
        if 0 <= a <= b <= len(out):
            out = out[:a] + replacement + out[b:]
    return out


def direct_provenance_for_key_in_strings(
    formula: str | None,
    normalized: str | None,
    dep_key: str,
) -> EdgeProvenance:
    """Build minimal direct-ref provenance by locating `dep_key` substrings."""
    sites_f: list[tuple[int, int]] = []
    sites_n: list[tuple[int, int]] = []
    if formula:
        sites_f.extend(_find_literal_spans(formula, dep_key))
    if normalized:
        sites_n.extend(_find_literal_spans(normalized, dep_key))
    return EdgeProvenance(
        causes=frozenset({DependencyCause.direct_ref}),
        direct_sites_formula=tuple(sites_f),
        direct_sites_normalized=tuple(sites_n),
    )


def _find_literal_spans(s: str, needle: str) -> list[tuple[int, int]]:
    if not needle:
        return []
    out: list[tuple[int, int]] = []
    i = 0
    while True:
        j = s.find(needle, i)
        if j < 0:
            break
        out.append((j, j + len(needle)))
        i = j + len(needle)
    return out


class CompressionProvenanceRequiredError(RuntimeError):
    """Compression cannot run safely without captured dependency provenance."""


def refresh_direct_sites(
    provenance: EdgeProvenance,
    *,
    old_formula: str | None,
    new_formula: str | None,
    old_normalized: str | None,
    new_normalized: str | None,
    precedent_key: NodeKey,
) -> EdgeProvenance:
    """Re-locate direct-ref spans after a dependent formula rewrite."""
    sites_n: list[tuple[int, int]] = []
    if new_normalized:
        sites_n = _find_literal_spans(new_normalized, precedent_key)

    sites_f: list[tuple[int, int]] = []
    if new_formula:
        if provenance.direct_sites_formula and old_formula:
            for a, b in provenance.direct_sites_formula:
                if 0 <= a <= b <= len(old_formula):
                    needle = old_formula[a:b]
                    if needle:
                        sites_f.extend(_find_literal_spans(new_formula, needle))
        else:
            sites_f = _find_literal_spans(new_formula, precedent_key)

    return replace(
        provenance,
        direct_sites_formula=tuple(sites_f),
        direct_sites_normalized=tuple(sites_n),
    )


def _structural_inline_candidate(
    graph: DependencyGraph,
    transit_key: NodeKey,
) -> NodeKey | None:
    """Return the sole dependent when `transit_key` is structurally inlinable."""
    if is_identity_transit(graph, transit_key) is not None:
        return None
    t_node = graph.get_node(transit_key)
    if t_node is None or t_node.is_leaf or t_node.formula is None:
        return None
    dependents = graph.get_dependents(transit_key)
    if len(dependents) != 1:
        return None
    dependent = next(iter(dependents))
    if graph._is_dependency_reachable(transit_key, dependent):
        return None
    if graph.get_edge_guard(dependent, transit_key) is not None:
        return None
    if not node_body_substitutable(graph, transit_key):
        return None
    if not dependent_context_substitutable(graph, dependent, replacing=transit_key):
        return None
    prov = graph.get_edge_attrs(dependent, transit_key).provenance
    if prov is not None:
        if len(prov.direct_sites_normalized) != 1:
            return None
        dep_node = graph.get_node(dependent)
        if dep_node is None:
            return None
        if dep_node.formula is not None and len(prov.direct_sites_formula) != 1:
            return None
    return dependent


def require_compression_provenance(graph: DependencyGraph) -> None:
    """Raise when compression candidates exist but edge provenance is missing."""
    missing: list[tuple[str, str]] = []
    for transit_key in graph:
        replacement = is_identity_transit(graph, transit_key)
        if replacement is not None:
            for dependent in graph.get_dependents(transit_key):
                if graph.get_edge_attrs(dependent, transit_key).provenance is None:
                    missing.append((dependent, transit_key))
            continue
        dependent = _structural_inline_candidate(graph, transit_key)
        if (
            dependent is not None
            and graph.get_edge_attrs(dependent, transit_key).provenance is None
        ):
            missing.append((dependent, transit_key))
    if missing:
        dependent, transit = missing[0]
        raise CompressionProvenanceRequiredError(
            "Dependency provenance is required for compression but is missing on edge "
            f"{dependent} -> {transit}. Build the graph with "
            "capture_dependency_provenance=True."
        )


def compression_safe_provenance(prov: EdgeProvenance | None) -> bool:
    if prov is None:
        return False
    if DependencyCause.direct_ref not in prov.causes:
        return False
    unsafe = {
        DependencyCause.static_range,
        DependencyCause.dynamic_offset,
        DependencyCause.dynamic_indirect,
    }
    return not (prov.causes & unsafe)


@dataclass(frozen=True)
class FormulaRewrite:
    """Dependent formula rewrite performed during projection."""

    dependent: str
    before_formula: str | None
    after_formula: str | None
    before_normalized: str | None
    after_normalized: str | None

    def to_dict(self) -> dict[str, Any]:
        return {
            "dependent": self.dependent,
            "before_formula": self.before_formula,
            "after_formula": self.after_formula,
            "before_normalized": self.before_normalized,
            "after_normalized": self.after_normalized,
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> FormulaRewrite:
        return cls(
            dependent=str(data["dependent"]),
            before_formula=data.get("before_formula"),
            after_formula=data.get("after_formula"),
            before_normalized=data.get("before_normalized"),
            after_normalized=data.get("after_normalized"),
        )


@dataclass(frozen=True)
class ProjectedNodeSnapshot:
    """Workbook node state captured before projection removes or rewrites it."""

    address: str
    sheet: str
    column: str
    row: int
    formula: str | None
    normalized_formula: str | None
    value: Any
    is_target: bool
    is_leaf: bool
    metadata: dict[str, Any]

    def to_dict(self) -> dict[str, Any]:
        return {
            "address": self.address,
            "sheet": self.sheet,
            "column": self.column,
            "row": self.row,
            "formula": self.formula,
            "normalized_formula": self.normalized_formula,
            "value": self.value,
            "is_target": self.is_target,
            "is_leaf": self.is_leaf,
            "metadata": dict(self.metadata),
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> ProjectedNodeSnapshot:
        return cls(
            address=str(data["address"]),
            sheet=str(data["sheet"]),
            column=str(data["column"]),
            row=int(data["row"]),
            formula=data.get("formula"),
            normalized_formula=data.get("normalized_formula"),
            value=data.get("value"),
            is_target=bool(data.get("is_target", False)),
            is_leaf=bool(data.get("is_leaf", False)),
            metadata=dict(data.get("metadata") or {}),
        )


@dataclass
class IdentityTransitCompressionRecord:
    """Lineage collected while compressing identity transit nodes."""

    immediate_removed: dict[str, str] = field(default_factory=dict)
    removal_order: list[str] = field(default_factory=list)
    formula_rewrites: list[FormulaRewrite] = field(default_factory=list)
    snapshots_by_removed: dict[str, ProjectedNodeSnapshot] = field(default_factory=dict)

    def note_removal(
        self,
        t_key: NodeKey,
        r_key: NodeKey,
        snapshot: ProjectedNodeSnapshot,
    ) -> None:
        self.immediate_removed[t_key] = r_key
        self.removal_order.append(t_key)
        self.snapshots_by_removed[t_key] = snapshot


def snapshot_transit_node(graph: DependencyGraph, key: NodeKey) -> ProjectedNodeSnapshot:
    """Capture workbook node state for `key` before identity transit removal."""
    node = graph.get_node(key)
    if node is None:
        raise KeyError(key)
    if (
        node.shape is not NodeShape.cell
        or node.sheet is None
        or node.column is None
        or node.row is None
    ):
        raise ValueError(f"ProjectedNodeSnapshot requires a cell node: {key}")
    return ProjectedNodeSnapshot(
        address=key,
        sheet=node.sheet,
        column=node.column,
        row=node.row,
        formula=node.formula,
        normalized_formula=node.normalized_formula,
        value=node.value,
        is_target=node.is_target,
        is_leaf=node.is_leaf,
        metadata=dict(node.metadata),
    )


def _formula_body(formula: str) -> str:
    return formula[1:] if formula.startswith("=") else formula


def node_body_substitutable(graph: DependencyGraph, key: NodeKey) -> bool:
    """Return whether `key`'s formula body may be pasted into a sole dependent."""
    node = graph.get_node(key)
    if node is None or node.is_leaf or not node.normalized_formula:
        return False
    for dep in graph.get_dependencies(key):
        attrs = graph.get_edge_attrs(key, dep)
        if graph.get_edge_guard(key, dep) is not None:
            return False
        if not compression_safe_provenance(attrs.provenance):
            return False
    return True


def dependent_context_substitutable(
    graph: DependencyGraph,
    dependent: NodeKey,
    *,
    replacing: NodeKey,
) -> bool:
    """Return whether `dependent` can be rewritten after inlining `replacing`."""
    for dep in graph.get_dependencies(dependent):
        if dep == replacing:
            continue
        if graph.get_edge_guard(dependent, dep) is not None:
            return False
        attrs = graph.get_edge_attrs(dependent, dep)
        if not compression_safe_provenance(attrs.provenance):
            return False
    return True


def _incoming_edge_substitutable(
    graph: DependencyGraph,
    dependent: NodeKey,
    precedent: NodeKey,
) -> bool:
    if graph.get_edge_guard(dependent, precedent) is not None:
        return False
    prov = graph.get_edge_attrs(dependent, precedent).provenance
    if not compression_safe_provenance(prov):
        return False
    if prov is None:
        return False
    if len(prov.direct_sites_normalized) != 1:
        return False
    dep_node = graph.get_node(dependent)
    if dep_node is None:
        return False
    return dep_node.formula is None or len(prov.direct_sites_formula) == 1


def substitute_body_at_spans(
    formula: str,
    spans: tuple[tuple[int, int], ...],
    body_formula: str,
) -> str:
    """Replace provenance spans with a parenthesized formula body."""
    return replace_substrings_at_spans(formula, spans, f"({_formula_body(body_formula)})")


@dataclass
class OptimalCompressionRecord:
    """Lineage collected while performing optimal graph compression."""

    forwarded_removed: dict[str, str] = field(default_factory=dict)
    inlined_to: dict[str, str] = field(default_factory=dict)
    removal_order: list[str] = field(default_factory=list)
    formula_rewrites: list[FormulaRewrite] = field(default_factory=list)
    snapshots_by_removed: dict[str, ProjectedNodeSnapshot] = field(default_factory=dict)

    def note_forwarding(
        self,
        removed: NodeKey,
        replacement: NodeKey,
        snapshot: ProjectedNodeSnapshot,
    ) -> None:
        self.forwarded_removed[removed] = replacement
        self.removal_order.append(removed)
        self.snapshots_by_removed[removed] = snapshot

    def note_inline(
        self,
        removed: NodeKey,
        retained: NodeKey,
        snapshot: ProjectedNodeSnapshot,
    ) -> None:
        self.inlined_to[removed] = retained
        self.removal_order.append(removed)
        self.snapshots_by_removed[removed] = snapshot

    def ensure_snapshot(self, key: NodeKey, snapshot: ProjectedNodeSnapshot) -> None:
        if key not in self.snapshots_by_removed:
            self.snapshots_by_removed[key] = snapshot
