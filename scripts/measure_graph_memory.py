#!/usr/bin/env python3
"""Measure `DependencyGraph` in-memory size, broken down by component.

Every figure is produced by a recursive walk that counts each distinct object
**once**, keyed on `id()`. That matters here: canonical key strings, `GuardExpr`
trees and `IntFlag` members are shared between maps (and some are process-wide
singletons), so naive per-object summing double-counts them badly enough to
invert conclusions about where graph memory actually goes.

Each component therefore reports both:

- `exclusive_bytes` — objects reachable only from that component (owned), and
- `shared_bytes` — objects also reachable from another component.

A drop in `total_bytes` that only moves bytes from `exclusive` to `shared`
somewhere else is a re-attribution, not a saving.

Empty adjacency sets are **split out** of the forward/reverse maps (not walked
as a second root) so the empty-`set()` tax stays exclusive to those rows when
those sets exist. After #910, `add_node` omits empty neighbor sets, so those
rows are absent on a current graph.
`formula_ast` intern-pool figures are recovered from unique `Node.formula_ast`
trees; they are not a second walk component, so interned trees stay exclusive
to `nodes` unless another stored field also reaches them.

Usage:
    uv run python scripts/measure_graph_memory.py
    uv run python scripts/measure_graph_memory.py --workbook book.xlsx --targets 'Sheet1!A1'
    uv run python scripts/measure_graph_memory.py --workbook examples/micro_workbooks/if_guards.xlsx --targets 'Sheet1!D1:D200'
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from enum import Enum
from functools import cache
from pathlib import Path
from types import BuiltinFunctionType, FunctionType, MethodType, ModuleType
from typing import Any

from excel_grapher.core.formula_ast import AstNode
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.guard import guard_intern_pool_size

_DESCRIPTION = "Measure DependencyGraph in-memory size, broken down by component."

_REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_WORKBOOK = _REPO_ROOT / "examples" / "micro_workbooks" / "taco_patterns.xlsx"
DEFAULT_TARGETS = (
    "Patterns!D3:D7",
    "Patterns!F3:F7",
    "Patterns!H3:H7",
    "Patterns!K3:K7",
    "Patterns!P3:P7",
)
IF_HEAVY_WORKBOOK = _REPO_ROOT / "examples" / "micro_workbooks" / "if_guards.xlsx"
IF_HEAVY_TARGETS = ("Sheet1!D1:D200",)

_EMPTY_TUPLE: tuple[object, ...] = ()
_SMALL_INT_MIN = -5
_SMALL_INT_MAX = 256

_CONTAINER_TYPES = (dict, list, tuple, set, frozenset)
_LEAF_TYPES = (str, bytes, bytearray, int, float, complex, type(None), range)
_OPAQUE_TYPES = (type, ModuleType, FunctionType, MethodType, BuiltinFunctionType)


# ---- object walk ----------------------------------------------------------


@dataclass(frozen=True, slots=True)
class _ObjectInfo:
    """Shallow measurement of one distinct object seen by the walk."""

    size: int
    is_container: bool
    is_singleton: bool


@cache
def _slot_names(cls: type) -> tuple[str, ...]:
    """Return every `__slots__` name declared across `cls`'s MRO."""
    names: list[str] = []
    for klass in cls.__mro__:
        slots = klass.__dict__.get("__slots__")
        if slots is None:
            continue
        if isinstance(slots, str):
            slots = (slots,)
        names.extend(name for name in slots if name not in ("__dict__", "__weakref__"))
    return tuple(dict.fromkeys(names))


def _is_process_singleton(obj: object) -> bool:
    """Return True for objects CPython shares process-wide, graph or no graph.

    Enum members (including `DependencyCause` flags), `None`/`True`/`False`,
    small ints, single-character strings and the empty tuple all exist whether
    or not a graph does, so charging them to a graph component overstates it.
    """
    if obj is None or obj is True or obj is False or obj is Ellipsis:
        return True
    if obj is _EMPTY_TUPLE:
        return True
    if isinstance(obj, Enum):
        return True
    if isinstance(obj, int) and _SMALL_INT_MIN <= obj <= _SMALL_INT_MAX:
        return True
    return isinstance(obj, str) and len(obj) <= 1


def _referents(obj: object) -> list[object]:
    """Return the objects directly held by `obj`, excluding classes and modules."""
    if isinstance(obj, dict):
        out: list[object] = []
        for key, value in obj.items():
            out.append(key)
            out.append(value)
        return out
    if isinstance(obj, (tuple, list, set, frozenset)):
        return list(obj)
    if isinstance(obj, (*_LEAF_TYPES, *_OPAQUE_TYPES)):
        return []
    if isinstance(obj, Mapping):
        out = []
        for key, value in obj.items():
            out.append(key)
            out.append(value)
        return out
    slots = _slot_names(type(obj))
    instance_dict = getattr(obj, "__dict__", None)
    if not slots and not isinstance(instance_dict, dict):
        # Opaque C-level object (iterator, memoryview, ndarray, ...): getsizeof
        # already covers whatever buffer it owns, and there is nothing safe to walk.
        return []

    out = []
    if isinstance(instance_dict, dict):
        # Charge the per-instance __dict__ itself; dropping it is the slots win.
        out.append(instance_dict)
    for name in slots:
        try:
            out.append(getattr(obj, name))
        except AttributeError:
            continue
    return out


def _walk(roots: Iterable[object]) -> dict[int, _ObjectInfo]:
    """Return `{id(obj): info}` for every distinct object reachable from `roots`."""
    seen: dict[int, _ObjectInfo] = {}
    stack = list(roots)
    while stack:
        obj = stack.pop()
        obj_id = id(obj)
        if obj_id in seen:
            continue
        singleton = _is_process_singleton(obj)
        seen[obj_id] = _ObjectInfo(
            size=sys.getsizeof(obj),
            is_container=isinstance(obj, _CONTAINER_TYPES),
            is_singleton=singleton,
        )
        if singleton:
            continue
        stack.extend(_referents(obj))
    return seen


def deep_size(*roots: object, include_singletons: bool = False) -> int:
    """Return the byte size of every distinct object reachable from `roots`.

    Objects reached more than once are counted once. Process-wide singletons
    (see `_is_process_singleton`) are excluded unless `include_singletons` is
    set, because they are not attributable to the structure being measured.

    Args:
        *roots: Objects to walk.
        include_singletons: Count process-wide shared objects too.

    Returns:
        Total bytes, as reported by `sys.getsizeof` per distinct object.
    """
    return sum(
        info.size for info in _walk(roots).values() if include_singletons or not info.is_singleton
    )


# ---- report ---------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class ComponentSize:
    """Measured size of one graph component."""

    name: str
    note: str
    node_count: int
    edge_count: int
    total_bytes: int
    exclusive_bytes: int
    shared_bytes: int
    scaffolding_bytes: int
    object_count: int

    @property
    def bytes_per_node(self) -> float:
        """Return component bytes per graph node (0.0 for an empty graph)."""
        return self.total_bytes / self.node_count if self.node_count else 0.0

    @property
    def bytes_per_edge(self) -> float:
        """Return component bytes per graph edge (0.0 when there are no edges)."""
        return self.total_bytes / self.edge_count if self.edge_count else 0.0

    def to_dict(self) -> dict[str, Any]:
        """Return a JSON-serializable view of this component."""
        return {
            "name": self.name,
            "note": self.note,
            "total_bytes": self.total_bytes,
            "exclusive_bytes": self.exclusive_bytes,
            "shared_bytes": self.shared_bytes,
            "scaffolding_bytes": self.scaffolding_bytes,
            "object_count": self.object_count,
            "bytes_per_node": self.bytes_per_node,
            "bytes_per_edge": self.bytes_per_edge,
        }


@dataclass(frozen=True, slots=True)
class GraphMemoryReport:
    """Component-level memory breakdown for one `DependencyGraph`."""

    node_count: int
    edge_count: int
    components: tuple[ComponentSize, ...]
    total_bytes: int
    shared_bytes: int
    singleton_bytes: int
    empty_forward_sets: int
    empty_reverse_sets: int
    forward_sets: int
    reverse_sets: int
    empty_forward_bytes: int
    empty_reverse_bytes: int
    guarded_edge_count: int
    guard_intern_pool_size: int
    identity_distinct_guards: int
    formula_ast_intern_count: int
    formula_ast_intern_bytes: int
    formula_nodes_with_ast: int

    @property
    def bytes_per_node(self) -> float:
        """Return graph bytes per node (0.0 for an empty graph)."""
        return self.total_bytes / self.node_count if self.node_count else 0.0

    @property
    def bytes_per_edge(self) -> float:
        """Return graph bytes per edge (0.0 when there are no edges)."""
        return self.total_bytes / self.edge_count if self.edge_count else 0.0

    def component(self, name: str) -> ComponentSize:
        """Return the component named `name`.

        Raises:
            KeyError: If no component with that name was measured.
        """
        for component in self.components:
            if component.name == name:
                return component
        raise KeyError(name)

    def to_dict(self) -> dict[str, Any]:
        """Return a JSON-serializable view of the whole report."""
        return {
            "node_count": self.node_count,
            "edge_count": self.edge_count,
            "total_bytes": self.total_bytes,
            "shared_bytes": self.shared_bytes,
            "singleton_bytes": self.singleton_bytes,
            "bytes_per_node": self.bytes_per_node,
            "bytes_per_edge": self.bytes_per_edge,
            "empty_forward_sets": self.empty_forward_sets,
            "empty_reverse_sets": self.empty_reverse_sets,
            "forward_sets": self.forward_sets,
            "reverse_sets": self.reverse_sets,
            "empty_forward_bytes": self.empty_forward_bytes,
            "empty_reverse_bytes": self.empty_reverse_bytes,
            "guarded_edge_count": self.guarded_edge_count,
            "guard_intern_pool_size": self.guard_intern_pool_size,
            "identity_distinct_guards": self.identity_distinct_guards,
            "formula_ast_intern_count": self.formula_ast_intern_count,
            "formula_ast_intern_bytes": self.formula_ast_intern_bytes,
            "formula_nodes_with_ast": self.formula_nodes_with_ast,
            "components": [component.to_dict() for component in self.components],
        }

    def render(self) -> str:
        """Return a human-readable table of the breakdown."""
        return _render(self)


@dataclass(frozen=True, slots=True)
class _ComponentSpec:
    """A named set of walk roots covering one part of the graph."""

    name: str
    note: str
    roots: tuple[object, ...]
    empty_adj: Mapping[Any, set[Any]] | None = None


def _unique_formula_asts(graph: DependencyGraph) -> tuple[AstNode, ...]:
    """Return identity-distinct `formula_ast` trees stored on `graph` nodes."""
    seen: dict[int, AstNode] = {}
    for node in graph._nodes.values():
        ast = node.formula_ast
        if ast is not None and id(ast) not in seen:
            seen[id(ast)] = ast
    return tuple(seen.values())


def _empty_adj_stats(adj: Mapping[Any, set[Any]] | None) -> tuple[int, int, int]:
    """Return `(set_count, empty_count, empty_bytes)` for one adjacency map."""
    if not adj:
        return 0, 0, 0
    n_sets = n_empty = empty_bytes = 0
    for neighbors in adj.values():
        n_sets += 1
        if not neighbors:
            n_empty += 1
            empty_bytes += sys.getsizeof(neighbors)
    return n_sets, n_empty, empty_bytes


def _component_specs(graph: DependencyGraph) -> list[_ComponentSpec]:
    """Return the component roots measured for `graph`, in report order."""
    specs = [
        _ComponentSpec(
            "nodes",
            "_nodes: Node instances, formula_ast trees, addresses, metadata",
            (graph._nodes,),
        ),
    ]
    if graph._staging:
        specs.extend(
            [
                _ComponentSpec(
                    "edges_forward",
                    "_edges: staging node -> dependency set (empty sets split into edges_forward_empty)",
                    (graph._edges,),
                    empty_adj=graph._edges,
                ),
                _ComponentSpec(
                    "edges_reverse",
                    "_reverse_edges: staging node -> dependent set (empty sets split into edges_reverse_empty)",
                    (graph._reverse_edges,),
                    empty_adj=graph._reverse_edges,
                ),
                _ComponentSpec(
                    "guards",
                    "_guards: staging edge key -> interned GuardExpr tree (trees are often shared)",
                    (graph._guards,),
                ),
                _ComponentSpec(
                    "provenance",
                    "_edge_provenance: staging edge key -> EdgeProvenance (causes + normalized site offsets)",
                    (graph._edge_provenance,),
                ),
            ]
        )
    else:
        specs.extend(
            [
                _ComponentSpec(
                    "adjacency_csr",
                    "uint32 CSR+CSC arrays (row_ptr/col_idx + col_ptr/row_idx); no values array",
                    (graph._row_ptr, graph._col_idx, graph._col_ptr, graph._row_idx),
                ),
                _ComponentSpec(
                    "node_index",
                    "NodeKey -> row-id table plus CSR key list (keys shared with _nodes)",
                    (graph._node_index, graph._csr_keys),
                ),
                _ComponentSpec(
                    "guards",
                    "uint32 guard_id[nnz] plus interned GuardExpr table (0 = unguarded)",
                    (graph._guard_id, graph._guard_exprs),
                ),
                _ComponentSpec(
                    "provenance",
                    "uint32 prov_id[nnz] plus interned EdgeProvenance table (0 = none)",
                    (graph._prov_id, graph._provenances),
                ),
            ]
        )
    metadata_roots = tuple(
        value
        for value in (
            graph.leaf_classification,
            graph.sheet_order,
            graph.sheet_bounds,
            graph.named_ranges,
            graph.named_range_ranges,
            graph.cell_type_env,
        )
        if value is not None
    )
    if metadata_roots:
        specs.append(
            _ComponentSpec(
                "workbook_metadata",
                "leaf_classification, sheet_order, sheet_bounds, named ranges, cell_type_env",
                metadata_roots,
            )
        )
    if graph.preparsed_formulas is not None:
        specs.append(
            _ComponentSpec(
                "preparsed_formulas",
                "preparsed_formulas: opt-in AST cache from warm_ast_cache",
                (graph.preparsed_formulas,),
            )
        )
    if graph.formula_shapes is not None:
        specs.append(
            _ComponentSpec(
                "formula_shapes",
                "formula_shapes: opt-in interned skeleton overlay from warm_formula_shapes",
                (graph.formula_shapes,),
            )
        )
    return specs


def edge_count(graph: DependencyGraph) -> int:
    """Return the number of stored dependency edges in `graph`."""
    return graph.edge_count()


def _guarded_edge_count(graph: DependencyGraph) -> int:
    """Return the number of edges that carry a guard."""
    if graph._staging:
        return len(graph._guards)
    return sum(1 for gid in graph._guard_id if gid)


def _identity_distinct_guards(graph: DependencyGraph) -> int:
    """Return the number of distinct interned `GuardExpr` trees stored on `graph`."""
    if graph._staging:
        return len({id(expr) for expr in graph._guards.values()})
    return sum(expr is not None for expr in graph._guard_exprs)


def _accumulate_component(
    *,
    name: str,
    note: str,
    node_count: int,
    edge_count_value: int,
    walk: Mapping[int, _ObjectInfo],
    reach_counts: Counter[int],
    include_ids: set[int] | None = None,
    exclude_ids: set[int] | None = None,
) -> ComponentSize:
    """Build one `ComponentSize` from a (possibly filtered) walk."""
    total = exclusive = shared = scaffolding = count = 0
    for obj_id, info in walk.items():
        if include_ids is not None and obj_id not in include_ids:
            continue
        if exclude_ids is not None and obj_id in exclude_ids:
            continue
        if info.is_singleton:
            continue
        count += 1
        total += info.size
        if reach_counts[obj_id] > 1:
            shared += info.size
        else:
            exclusive += info.size
        if info.is_container:
            scaffolding += info.size
    return ComponentSize(
        name=name,
        note=note,
        node_count=node_count,
        edge_count=edge_count_value,
        total_bytes=total,
        exclusive_bytes=exclusive,
        shared_bytes=shared,
        scaffolding_bytes=scaffolding,
        object_count=count,
    )


def measure_graph_memory(graph: DependencyGraph) -> GraphMemoryReport:
    """Return a component-level memory breakdown for `graph`.

    Each component is walked independently; an object reachable from more than
    one component is counted once per component in `total_bytes` but only once
    in the report's `total_bytes`, and is reported as shared rather than owned.

    Empty adjacency sets are partitioned out of the parent map walk (same
    object ids, not a second walk) so their exclusive bytes stay visible.

    Args:
        graph: The graph to measure.

    Returns:
        A `GraphMemoryReport` with per-component and whole-graph figures.
    """
    nodes = len(graph)
    edges = edge_count(graph)
    specs = _component_specs(graph)
    walks = {spec.name: _walk(spec.roots) for spec in specs}

    reach_counts: Counter[int] = Counter()
    for objects in walks.values():
        reach_counts.update(objects.keys())

    components: list[ComponentSize] = []
    for spec in specs:
        walk = walks[spec.name]
        empty_ids: set[int] | None = None
        if spec.empty_adj is not None:
            empty_ids = {id(neighbors) for neighbors in spec.empty_adj.values() if not neighbors}
        components.append(
            _accumulate_component(
                name=spec.name,
                note=spec.note,
                node_count=nodes,
                edge_count_value=edges,
                walk=walk,
                reach_counts=reach_counts,
                exclude_ids=empty_ids,
            )
        )
        if empty_ids:
            components.append(
                _accumulate_component(
                    name=f"{spec.name}_empty",
                    note=f"empty neighbor sets split out of {spec.name}",
                    node_count=nodes,
                    edge_count_value=edges,
                    walk=walk,
                    reach_counts=reach_counts,
                    include_ids=empty_ids,
                )
            )

    distinct: dict[int, _ObjectInfo] = {}
    for objects in walks.values():
        distinct.update(objects)
    total_bytes = sum(info.size for info in distinct.values() if not info.is_singleton)
    shared_bytes = sum(
        info.size
        for obj_id, info in distinct.items()
        if not info.is_singleton and reach_counts[obj_id] > 1
    )
    singleton_bytes = sum(info.size for info in distinct.values() if info.is_singleton)

    forward_sets, empty_forward_sets, empty_forward_bytes = _empty_adj_stats(
        graph._edges if graph._staging else None
    )
    reverse_sets, empty_reverse_sets, empty_reverse_bytes = _empty_adj_stats(
        graph._reverse_edges if graph._staging else None
    )

    unique_asts = _unique_formula_asts(graph)
    formula_nodes_with_ast = sum(
        1 for node in graph._nodes.values() if node.formula_ast is not None
    )
    formula_ast_intern_bytes = deep_size(*unique_asts) if unique_asts else 0

    return GraphMemoryReport(
        node_count=nodes,
        edge_count=edges,
        components=tuple(components),
        total_bytes=total_bytes,
        shared_bytes=shared_bytes,
        singleton_bytes=singleton_bytes,
        empty_forward_sets=empty_forward_sets,
        empty_reverse_sets=empty_reverse_sets,
        forward_sets=forward_sets,
        reverse_sets=reverse_sets,
        empty_forward_bytes=empty_forward_bytes,
        empty_reverse_bytes=empty_reverse_bytes,
        guarded_edge_count=_guarded_edge_count(graph),
        guard_intern_pool_size=guard_intern_pool_size(),
        identity_distinct_guards=_identity_distinct_guards(graph),
        formula_ast_intern_count=len(unique_asts),
        formula_ast_intern_bytes=formula_ast_intern_bytes,
        formula_nodes_with_ast=formula_nodes_with_ast,
    )


# ---- rendering ------------------------------------------------------------

_HEADER = (
    f"{'component':<22}{'total':>14}{'B/node':>10}{'B/edge':>10}"
    f"{'exclusive':>14}{'shared':>14}{'scaffold':>14}{'objects':>12}"
)

_LEGEND = (
    "exclusive = bytes reachable only from this component (owned by it)\n"
    "shared    = bytes also reachable from another component; a total that falls\n"
    "            because bytes became shared with another component is a\n"
    "            re-attribution, not a saving. Shared bytes are counted once in\n"
    "            the graph total below, and once per component in the rows above.\n"
    "scaffold  = the dict/set/list/tuple containers themselves, excluding contents\n"
    "empty-set rows (staging maps only) are a partition of the parent adjacency walk,\n"
    "            not a second walk, so they stay exclusive rather than shared.\n"
    "CSR adjacency is counted as adjacency_csr (arrays) plus node_index (row-id table).\n"
    "Compact edge metadata is counted as guards/provenance intern-id arrays plus intern tables.\n"
    "Process-wide singletons (enum members, small ints, single-char strings) are\n"
    "excluded from every figure: they exist whether or not the graph does."
)


def _human(value: int) -> str:
    """Format `value` bytes as KiB, or MiB when the figure is at least 1 MiB."""
    if value >= 1024 * 1024:
        return f"{value / (1024 * 1024):,.1f} MiB"
    return f"{value / 1024:,.1f} KiB"


def _render(report: GraphMemoryReport) -> str:
    lines = [
        f"DependencyGraph: {report.node_count:,} nodes, {report.edge_count:,} edges",
        "",
        _HEADER,
        "-" * len(_HEADER),
    ]
    for component in report.components:
        lines.append(
            f"{component.name:<22}{component.total_bytes:>14,}"
            f"{component.bytes_per_node:>10.1f}{component.bytes_per_edge:>10.1f}"
            f"{component.exclusive_bytes:>14,}{component.shared_bytes:>14,}"
            f"{component.scaffolding_bytes:>14,}{component.object_count:>12,}"
        )
    naive = sum(component.total_bytes for component in report.components)
    empty_tax = report.empty_forward_bytes + report.empty_reverse_bytes
    empty_tax_per_node = empty_tax / report.node_count if report.node_count else 0.0
    lines += [
        "-" * len(_HEADER),
        f"{'graph total':<22}{report.total_bytes:>14,}"
        f"{report.bytes_per_node:>10.1f}{report.bytes_per_edge:>10.1f}",
        "",
        f"graph total          {report.total_bytes:>14,} B ({_human(report.total_bytes)})",
        f"  of which shared    {report.shared_bytes:>14,} B ({_human(report.shared_bytes)})",
        f"sum of component rows{naive:>14,} B "
        f"(over-counts shared objects by {naive - report.total_bytes:,} B)",
        f"process singletons   {report.singleton_bytes:>14,} B (excluded from the total)",
        "",
        "adjacency empty-set tax:",
        f"  forward  {report.empty_forward_sets:,} empty / {report.forward_sets:,} sets"
        f"  {report.empty_forward_bytes:,} B",
        f"  reverse  {report.empty_reverse_sets:,} empty / {report.reverse_sets:,} sets"
        f"  {report.empty_reverse_bytes:,} B",
        f"  combined {empty_tax:,} B ({empty_tax_per_node:.1f} B/node)",
        "",
        "guards:",
        f"  guarded edges          {report.guarded_edge_count:,}",
        f"  identity-distinct trees{report.identity_distinct_guards:,}",
        f"  guard intern pool      {report.guard_intern_pool_size:,} live entries"
        " (process-wide WeakValueDictionary)",
        "",
        "formula_ast intern pool (unique Node.formula_ast trees, counted once):",
        f"  nodes with formula_ast {report.formula_nodes_with_ast:,}",
        f"  interned trees         {report.formula_ast_intern_count:,}",
        f"  interned tree bytes    {report.formula_ast_intern_bytes:,} B"
        f" ({_human(report.formula_ast_intern_bytes)})",
        "",
        "notes:",
    ]
    lines += [f"  {component.name:<22}{component.note}" for component in report.components]
    lines += ["", _LEGEND]
    return "\n".join(lines)


# ---- CLI ------------------------------------------------------------------


def _build_graph(args: argparse.Namespace) -> DependencyGraph:
    from excel_grapher import create_dependency_graph

    return create_dependency_graph(
        args.workbook,
        args.targets,
        load_values=args.load_values,
        capture_dependency_provenance=not args.no_provenance,
        warm_formula_shapes=args.warm_formula_shapes,
        warm_ast_cache=args.warm_ast_cache,
        use_cached_dynamic_refs=args.use_cached_dynamic_refs,
        store_raw_formula=args.store_raw_formula,
    )


def main(argv: list[str] | None = None) -> int:
    """Measure a workbook's dependency graph and print the breakdown."""
    parser = argparse.ArgumentParser(description=_DESCRIPTION)
    parser.add_argument(
        "--workbook",
        type=Path,
        default=DEFAULT_WORKBOOK,
        help=f"Workbook to build the graph from (default: {DEFAULT_WORKBOOK.name})",
    )
    parser.add_argument(
        "--targets",
        nargs="+",
        default=list(DEFAULT_TARGETS),
        help="Sheet-qualified target cells or ranges",
    )
    parser.add_argument(
        "--load-values",
        action="store_true",
        help="Load cached Excel values while building the graph",
    )
    parser.add_argument(
        "--no-provenance",
        action="store_true",
        help="Skip dependency-provenance capture (measures the graph without it)",
    )
    parser.add_argument(
        "--warm-formula-shapes",
        action="store_true",
        help="Install the optional formula_shapes overlay before measuring",
    )
    parser.add_argument(
        "--warm-ast-cache",
        action="store_true",
        help="Install the optional preparsed_formulas overlay before measuring",
    )
    parser.add_argument(
        "--use-cached-dynamic-refs",
        action="store_true",
        help="Resolve OFFSET/INDIRECT/INDEX from cached workbook values",
    )
    parser.add_argument(
        "--store-raw-formula",
        action="store_true",
        help="Keep Node.formula audit strings (off by default; matches extract)",
    )
    parser.add_argument("--json", action="store_true", help="Emit JSON instead of a table")
    args = parser.parse_args(argv)

    if not args.workbook.is_file():
        parser.error(f"workbook not found: {args.workbook}")

    graph = _build_graph(args)
    report = measure_graph_memory(graph)
    if args.json:
        print(json.dumps(report.to_dict(), indent=2))
    else:
        print(f"workbook: {args.workbook}")
        print(report.render())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
