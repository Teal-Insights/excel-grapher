"""Plan per-shape export helpers that bind child calls through holes.

Given a warm `graph.formula_shapes` overlay, classify each address hole on
each interned skeleton:

- *Leaf* — the hole always names an input (or an ineligible formula). The
  helper body reads it with `xl_cell` / `xl_range`.
- *Passthrough* — every instance's hole names the same child shape, and that
  child's params are either a constant or a uniform cell-coordinate offset of
  the parent hole. The helper calls the child shape directly.
- *Lookup* — the child shape or its params vary by instance. The helper
  indexes a dict from parent-hole address to `(child_fn, child_params)`.

See GitHub discussion of shape-dispatch codegen: one function per formula
shape rather than one function per cell.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from typing import Literal, Protocol, TypeAlias, cast

from fastpyxl.utils.cell import column_index_from_string

from excel_grapher.core.address_keys import (
    CellKey,
    parse_node_key,
)
from excel_grapher.core.formula_shape import (
    AddressKind,
    AddressLeaf,
    FormulaShapeTable,
    iter_address_holes,
    resolve_address_leaf,
)
from excel_grapher.grapher.node import NodeView

__all__ = [
    "ConstantArg",
    "GeometricCellArg",
    "LeafHolePlan",
    "LookupEntry",
    "LookupHolePlan",
    "PassthroughHolePlan",
    "ShapeDispatchLayout",
    "ShapeDispatchPlan",
    "analyze_shape_dispatch",
]


@dataclass(frozen=True, slots=True)
class GeometricCellArg:
    """Child cell address as a uniform `(dcol, drow)` offset of the parent hole."""

    dcol: int
    drow: int
    sheet: str | None = None


@dataclass(frozen=True, slots=True)
class ConstantArg:
    """Child param that is the same resolved address for every parent instance."""

    address: str


ChildArg: TypeAlias = GeometricCellArg | ConstantArg


@dataclass(frozen=True, slots=True)
class LeafHolePlan:
    """Hole that always reads a leaf (or ineligible formula) through the runtime."""

    kind: AddressKind
    mode: Literal["leaf"] = "leaf"


@dataclass(frozen=True, slots=True)
class LookupEntry:
    """One parent-hole address and the child call it should make.

    `child_shape_key` is `None` when the hole names a leaf (or an ineligible
    formula cell that must go through `xl_cell`).
    """

    host: str
    child_shape_key: str | None
    child_params: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class LookupHolePlan:
    """Hole whose child shape or params vary; resolved via a generated dict."""

    kind: AddressKind
    table_name: str
    entries: tuple[LookupEntry, ...]
    mode: Literal["lookup"] = "lookup"


@dataclass(frozen=True, slots=True)
class PassthroughHolePlan:
    """Hole that always names one child shape whose params are closed-form."""

    kind: AddressKind
    child_shape_key: str
    args: tuple[ChildArg, ...]
    mode: Literal["passthrough"] = "passthrough"


HolePlan: TypeAlias = LeafHolePlan | PassthroughHolePlan | LookupHolePlan


@dataclass(frozen=True, slots=True)
class ShapeDispatchPlan:
    """Export plan for one interned formula shape."""

    shape_key: str
    helper_name: str
    holes: tuple[HolePlan, ...]
    hosts: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class ShapeDispatchLayout:
    """Per-shape plans plus the cell→`(shape, params)` dispatch table."""

    plans: tuple[ShapeDispatchPlan, ...]
    cell_bindings: tuple[tuple[str, str, tuple[str, ...]], ...]

    @property
    def needs_offset_cell(self) -> bool:
        """True when any passthrough arg is a geometric cell offset."""
        return any(
            isinstance(hole, PassthroughHolePlan)
            and any(isinstance(arg, GeometricCellArg) for arg in hole.args)
            for plan in self.plans
            for hole in plan.holes
        )

    @property
    def needs_eval_shape(self) -> bool:
        """True when any hole calls a child shape (passthrough or lookup)."""
        return any(
            isinstance(hole, (PassthroughHolePlan, LookupHolePlan))
            for plan in self.plans
            for hole in plan.holes
        )

    @property
    def needs_eval_lookup(self) -> bool:
        """True when any hole uses a parent-param → child-param dict."""
        return any(isinstance(hole, LookupHolePlan) for plan in self.plans for hole in plan.holes)


class _GraphView(Protocol):
    """Minimal graph surface used while classifying holes."""

    def get_node(self, address: str) -> NodeView | None: ...

    @property
    def formula_shapes(self) -> FormulaShapeTable | None: ...


def _node_is_formula(node: object | None) -> bool:
    if node is None:
        return False
    has = getattr(node, "has_formula", None)
    if isinstance(has, bool):
        return has
    return (
        getattr(node, "formula_ast", None) is not None
        or getattr(node, "normalized_formula", None) is not None
    )


def _cell_col_row(key: CellKey) -> tuple[int, int]:
    return int(column_index_from_string(key.column)), int(key.row)


def _try_geometric_cell(parent: str, child: str) -> GeometricCellArg | None:
    try:
        parent_key = parse_node_key(parent)
        child_key = parse_node_key(child)
    except ValueError:
        return None
    if not isinstance(parent_key, CellKey) or not isinstance(child_key, CellKey):
        return None
    pcol, prow = _cell_col_row(parent_key)
    ccol, crow = _cell_col_row(child_key)
    sheet = None if parent_key.sheet == child_key.sheet else child_key.sheet
    return GeometricCellArg(dcol=ccol - pcol, drow=crow - prow, sheet=sheet)


def _uniform_child_args(
    pairs: Sequence[tuple[str, tuple[str, ...]]],
) -> tuple[ChildArg, ...] | None:
    """Return closed-form child args, or `None` when instances disagree."""
    if not pairs:
        return None
    arities = {len(params) for _, params in pairs}
    if len(arities) != 1:
        return None
    arity = next(iter(arities))
    args: list[ChildArg] = []
    for index in range(arity):
        addresses = tuple(params[index] for _, params in pairs)
        if len(set(addresses)) == 1:
            args.append(ConstantArg(addresses[0]))
            continue
        parents = [parent for parent, _ in pairs]
        geos = [
            _try_geometric_cell(parent, child)
            for parent, child in zip(parents, addresses, strict=True)
        ]
        if any(geo is None for geo in geos):
            return None
        first = geos[0]
        assert first is not None
        if any(geo != first for geo in geos):
            return None
        args.append(first)
    return tuple(args)


def _classify_hole(
    *,
    kind: AddressKind,
    table_name: str,
    instances: Sequence[tuple[str, str | None, tuple[str, ...] | None]],
    eligible: Mapping[str, str],
) -> HolePlan:
    """Classify one hole from per-instance `(parent_param, child_key, child_params)`."""
    if kind != "CELL":
        return LeafHolePlan(kind=kind)

    formula_instances = [
        (parent, child_key, child_params)
        for parent, child_key, child_params in instances
        if child_key is not None and child_key in eligible and child_params is not None
    ]
    if not formula_instances:
        return LeafHolePlan(kind=kind)

    if len(formula_instances) != len(instances):
        entries = tuple(
            LookupEntry(
                host=parent,
                child_shape_key=child_key,
                child_params=child_params or (),
            )
            for parent, child_key, child_params in formula_instances
        )
        return LookupHolePlan(kind=kind, table_name=table_name, entries=entries)

    child_keys = {child_key for _, child_key, _ in formula_instances}
    if len(child_keys) != 1:
        entries = tuple(
            LookupEntry(host=parent, child_shape_key=child_key, child_params=child_params or ())
            for parent, child_key, child_params in formula_instances
        )
        return LookupHolePlan(kind=kind, table_name=table_name, entries=entries)

    child_shape_key = next(iter(child_keys))
    assert child_shape_key is not None
    pairs = [(parent, child_params or ()) for parent, _, child_params in formula_instances]
    uniform = _uniform_child_args(pairs)
    if uniform is not None:
        return PassthroughHolePlan(
            kind=kind,
            child_shape_key=child_shape_key,
            args=uniform,
        )
    entries = tuple(
        LookupEntry(host=parent, child_shape_key=child_shape_key, child_params=child_params or ())
        for parent, _, child_params in formula_instances
    )
    return LookupHolePlan(kind=kind, table_name=table_name, entries=entries)


def _resolved_params(
    params: tuple[AddressLeaf, ...],
    host: str,
) -> tuple[str, ...]:
    return tuple(resolve_address_leaf(leaf, host) for leaf in params)


def _child_binding(
    graph: _GraphView,
    table: FormulaShapeTable,
    parent_param: str,
    eligible: Mapping[str, str],
) -> tuple[str | None, tuple[str, ...] | None]:
    node = graph.get_node(parent_param)
    if not _node_is_formula(node):
        return None, None
    found = table.lookup(parent_param)
    if found is None:
        return None, None
    child_key, _skeleton, child_params = found
    if child_key not in eligible:
        return None, None
    return child_key, _resolved_params(child_params, parent_param)


def analyze_shape_dispatch(
    graph: object,
    formula_addresses: Sequence[str],
    helper_names: Mapping[str, str],
) -> ShapeDispatchLayout:
    """Build per-shape hole plans for `formula_addresses`.

    Args:
        graph: Dependency graph with a warm `formula_shapes` overlay.
        formula_addresses: Formula cells in the export closure.
        helper_names: `shape_key` → generated helper identifier (`_shape_0`).

    Returns:
        A layout of per-shape hole plans and the cell dispatch table.

    Raises:
        ValueError: If `graph.formula_shapes` is missing.
    """
    table = getattr(graph, "formula_shapes", None)
    if table is None:
        raise ValueError("shape_dispatch requires graph.formula_shapes to be warm")
    table = cast(FormulaShapeTable, table)
    graph_view = cast(_GraphView, graph)

    hosts_by_shape: dict[str, list[str]] = {key: [] for key in helper_names}
    bindings: list[tuple[str, str, tuple[str, ...]]] = []
    for address in formula_addresses:
        found = table.lookup(address)
        if found is None:
            continue
        shape_key, _skeleton, params = found
        if shape_key not in helper_names:
            continue
        resolved = _resolved_params(params, address)
        hosts_by_shape[shape_key].append(address)
        bindings.append((address, shape_key, resolved))

    plans: list[ShapeDispatchPlan] = []
    for shape_key, helper_name in helper_names.items():
        skeleton = table.shapes[shape_key]
        holes_meta = list(iter_address_holes(skeleton))
        hosts = tuple(hosts_by_shape.get(shape_key, ()))
        hole_plans: list[HolePlan] = []
        for hole in holes_meta:
            instances: list[tuple[str, str | None, tuple[str, ...] | None]] = []
            for host in hosts:
                found = table.lookup(host)
                assert found is not None
                _key, _skel, params = found
                parent_param = resolve_address_leaf(params[hole.index], host)
                child_key, child_params = _child_binding(
                    graph_view, table, parent_param, helper_names
                )
                instances.append((parent_param, child_key, child_params))
            table_name = f"{helper_name}_p{hole.index}"
            hole_plans.append(
                _classify_hole(
                    kind=hole.kind,
                    table_name=table_name,
                    instances=instances,
                    eligible=helper_names,
                )
            )
        plans.append(
            ShapeDispatchPlan(
                shape_key=shape_key,
                helper_name=helper_name,
                holes=tuple(hole_plans),
                hosts=hosts,
            )
        )

    return ShapeDispatchLayout(plans=tuple(plans), cell_bindings=tuple(bindings))
