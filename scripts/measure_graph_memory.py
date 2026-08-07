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

Usage:
    uv run python scripts/measure_graph_memory.py
    uv run python scripts/measure_graph_memory.py --workbook book.xlsx --targets 'Sheet1!A1'
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

from excel_grapher.grapher.graph import DependencyGraph

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


def _component_specs(graph: DependencyGraph) -> list[_ComponentSpec]:
    """Return the component roots measured for `graph`, in report order."""
    specs = [
        _ComponentSpec(
            "nodes",
            "_nodes: Node instances, formula text, addresses, metadata dicts",
            (graph._nodes,),
        ),
        _ComponentSpec(
            "edges_forward",
            "_edges: node -> dependency set (one set object per node with deps)",
            (graph._edges,),
        ),
        _ComponentSpec(
            "edges_reverse",
            "_reverse_edges: node -> dependent set",
            (graph._reverse_edges,),
        ),
        _ComponentSpec(
            "guards",
            "_guards: edge key -> GuardExpr tree (trees are often shared)",
            (graph._guards,),
        ),
        _ComponentSpec(
            "provenance",
            "_edge_provenance: edge key -> EdgeProvenance (causes + site offsets)",
            (graph._edge_provenance,),
        ),
        _ComponentSpec(
            "occupancy",
            "_occupancy: member cell -> owning node key",
            (graph._occupancy,),
        ),
    ]
    metadata_roots = tuple(
        value
        for value in (
            graph.leaf_classification,
            graph.sheet_order,
            graph.sheet_bounds,
            graph.named_ranges,
            graph.named_range_ranges,
        )
        if value is not None
    )
    if metadata_roots:
        specs.append(
            _ComponentSpec(
                "workbook_metadata",
                "leaf_classification, sheet_order, sheet_bounds, named ranges",
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
    return specs


def edge_count(graph: DependencyGraph) -> int:
    """Return the number of stored dependency edges in `graph`."""
    return sum(len(deps) for deps in graph._edges.values())


def measure_graph_memory(graph: DependencyGraph) -> GraphMemoryReport:
    """Return a component-level memory breakdown for `graph`.

    Each component is walked independently; an object reachable from more than
    one component is counted once per component in `total_bytes` but only once
    in the report's `total_bytes`, and is reported as shared rather than owned.

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
        total = exclusive = shared = scaffolding = count = 0
        for obj_id, info in walks[spec.name].items():
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
        components.append(
            ComponentSize(
                name=spec.name,
                note=spec.note,
                node_count=nodes,
                edge_count=edges,
                total_bytes=total,
                exclusive_bytes=exclusive,
                shared_bytes=shared,
                scaffolding_bytes=scaffolding,
                object_count=count,
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

    return GraphMemoryReport(
        node_count=nodes,
        edge_count=edges,
        components=tuple(components),
        total_bytes=total_bytes,
        shared_bytes=shared_bytes,
        singleton_bytes=singleton_bytes,
    )


# ---- rendering ------------------------------------------------------------

_HEADER = (
    f"{'component':<20}{'total':>12}{'B/node':>10}{'B/edge':>10}"
    f"{'exclusive':>12}{'shared':>12}{'scaffold':>11}{'objects':>10}"
)

_LEGEND = (
    "exclusive = bytes reachable only from this component (owned by it)\n"
    "shared    = bytes also reachable from another component; a total that falls\n"
    "            because bytes became shared with another component is a\n"
    "            re-attribution, not a saving. Shared bytes are counted once in\n"
    "            the graph total below, and once per component in the rows above.\n"
    "scaffold  = the dict/set/list/tuple containers themselves, excluding contents\n"
    "Process-wide singletons (enum members, small ints, single-char strings) are\n"
    "excluded from every figure: they exist whether or not the graph does."
)


def _kib(value: int) -> str:
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
            f"{component.name:<20}{component.total_bytes:>12,}"
            f"{component.bytes_per_node:>10.1f}{component.bytes_per_edge:>10.1f}"
            f"{component.exclusive_bytes:>12,}{component.shared_bytes:>12,}"
            f"{component.scaffolding_bytes:>11,}{component.object_count:>10,}"
        )
    naive = sum(component.total_bytes for component in report.components)
    lines += [
        "-" * len(_HEADER),
        f"{'graph total':<20}{report.total_bytes:>12,}"
        f"{report.bytes_per_node:>10.1f}{report.bytes_per_edge:>10.1f}",
        "",
        f"graph total          {report.total_bytes:>12,} B ({_kib(report.total_bytes)})",
        f"  of which shared    {report.shared_bytes:>12,} B ({_kib(report.shared_bytes)})",
        f"sum of component rows{naive:>12,} B "
        f"(over-counts shared objects by {naive - report.total_bytes:,} B)",
        f"process singletons   {report.singleton_bytes:>12,} B (excluded from the total)",
        "",
        "notes:",
    ]
    lines += [f"  {component.name:<20}{component.note}" for component in report.components]
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
