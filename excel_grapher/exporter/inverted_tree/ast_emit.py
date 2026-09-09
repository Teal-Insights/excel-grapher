"""Translate a bound series' Excel AST into a first-level-dep Python helper."""

from __future__ import annotations

import re
from collections.abc import Mapping, Sequence
from contextvars import ContextVar
from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING

from excel_grapher.core.address_keys import CanonicalAddress, as_canonical, parse_cell_coords
from excel_grapher.core.excel_function_names import normalize_excel_function_name
from excel_grapher.core.formula_ast import (
    AbsoluteAxis,
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
    resolve_cell_ref,
)
from excel_grapher.core.formula_shape import fingerprint_formula_shape
from excel_grapher.exporter.inverted_tree import runtime as inverted_runtime
from excel_grapher.exporter.inverted_tree.access import (
    AccessFunction,
    AxisAccess,
    catalog_index_affine,
    cell_ref_catalog_pairs,
    classify_cell_ref_access,
    classify_producer_access,
    indirect_argument_addresses,
    indirect_target_addresses,
)
from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    KeyPoint,
    SeriesCatalog,
    Statement,
    covering_series,
    covering_series_of_column,
    covering_series_of_range,
    fit_affine_map,
    preferred_fields,
    schedule_axis_coord,
    schedule_partition,
)
from excel_grapher.exporter.inverted_tree.deps import (
    PositionalRangeCell,
    SeriesDeps,
    _field_binding,
    _host_follow_key_maps,
    _host_follow_pair_maps,
    _host_join_key,
    _host_producer_slots,
    _lookup_axis_fields,
    _ref_pinned_fields,
    addresses_outside_blank_ranges,
    covering_series_for_index_window,
    current_blank_rects,
    index_call_is_ref,
    index_window_corners,
    iter_range_addresses,
    iter_ref_addresses,
    node_formula_ast,
    offset_index_destination,
    predecessor_address,
    range_ref_label,
    resolve_offset_destination_series,
    resolve_positional_range,
    successor_address,
    try_formula_ast,
)
from excel_grapher.exporter.inverted_tree.domains import (
    DomainEmitPlan,
    series_domain_points,
    uses_datetime_values,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.schedule import (
    FusedPlan,
    IndexSourceIntern,
    indices_to_source,
    slot_table_needs_intern,
    slot_table_to_source,
    wrap_slot_table_source,
)
from excel_grapher.grapher.blank_ranges import BlankRangeRect, address_in_blank_ranges
from excel_grapher.series_bindings.types import Scalar

if TYPE_CHECKING:
    from excel_grapher.exporter.inverted_tree.named_axes import NamedAxes
    from excel_grapher.grapher.graph import DependencyGraph

_ARITHMETIC_HELPERS = {
    "+": "xl_add",
    "-": "xl_sub",
    "*": "xl_mul",
    "/": "xl_div",
    "^": "xl_pow",
}
_COMPARE_HELPERS = {
    "=": "xl_eq",
    "<>": "xl_ne",
    "<": "xl_lt",
    ">": "xl_gt",
    "<=": "xl_le",
    ">=": "xl_ge",
}
_UNARY_HELPERS = {
    "-": "xl_neg",
    "+": "xl_pos",
}
_RUNTIME_FUNCTIONS = frozenset(
    name
    for name, value in vars(inverted_runtime).items()
    if name.startswith("xl_") and callable(value)
)
_AGGREGATE_FUNCTIONS = frozenset(
    {"SUM", "SUMPRODUCT", "AVERAGE", "MAX", "MIN", "NPV", "STDEV", "RANK", "LARGE", "COUNTIF"}
)
_RANGE_REDUCE_FUNCTIONS = _AGGREGATE_FUNCTIONS | frozenset({"AND", "OR"})
_LOOKUP_TABLE_FUNCTIONS = frozenset({"VLOOKUP", "HLOOKUP", "LOOKUP", "XLOOKUP"})
_PURE_LOOKUP_FUNCTIONS = _LOOKUP_TABLE_FUNCTIONS | frozenset({"INDEX", "MATCH"})
_ARRAY_IF_VALUE_OPS = frozenset(_ARITHMETIC_HELPERS) | frozenset(_COMPARE_HELPERS)
_ARRAY_IF_UNSOUND_FNS = frozenset(
    {
        "AND",
        "OR",
        "IFS",
        "CHOOSE",
        "SWITCH",
        "SUM",
        "SUMPRODUCT",
        "AVERAGE",
        "AGGREGATE",
    }
)


@dataclass
class KeyedReadIntern:
    """Intern domain tuples, slot maps, catalog-index tables, and member tables.

    `emit_internals_module` binds one instance for a package so repeated keyed
    reads share metadata instead of embedding it at every site (#786). Nested
    aggregate member tables that need construction are interned the same way
    so compact source does not reallocate per host index (#795).
    """

    plan: DomainEmitPlan
    _sequences: dict[tuple[object, ...], str] = field(default_factory=dict)
    _slot_maps: dict[str, str] = field(default_factory=dict)
    _index_intern: IndexSourceIntern = field(
        default_factory=lambda: IndexSourceIntern(prefix="_SLOTS_")
    )
    _member_tables: dict[tuple[tuple[int, ...], ...], str] = field(default_factory=dict)
    uses_data: bool = False
    uses_datetime: bool = False

    def sequence(self, values: tuple[object, ...], *, field: str | None = None) -> str:
        """Return an expression for `values`, interned or reused from `data`."""
        expr = self.plan.sequence_expr(values, field=field)
        if expr is not None:
            if "data." in expr:
                self.uses_data = True
            return expr
        existing = self._sequences.get(values)
        if existing is not None:
            return existing
        if uses_datetime_values(values):
            self.uses_datetime = True
        name = f"_KEYS_{len(self._sequences)}"
        self._sequences[values] = name
        return name

    def slot_map(self, series_id: str) -> str:
        """Return the interned `{key: catalog_index}` name for `series_id`."""
        existing = self._slot_maps.get(series_id)
        if existing is not None:
            return existing
        name = f"_INDEX_{series_id}"
        if not name.isidentifier():
            name = f"_INDEX_{len(self._slot_maps)}"
        self._slot_maps[series_id] = name
        if "data." in self.plan.series_expr.get(series_id, ""):
            self.uses_data = True
        return name

    def slots(self, indices: tuple[int, ...]) -> str:
        """Return a compact or interned expression for catalog-index `indices`."""
        return self._index_intern.expr(indices)

    def member_slots(self, table: tuple[tuple[int, ...], ...]) -> str:
        """Return a compact or interned expression for a nested member-slot table.

        Constructor forms (`*`, concatenation, comprehensions) are bound once at
        module level so compact source does not reallocate on every host index.
        """
        source = slot_table_to_source(table)
        if not slot_table_needs_intern(source):
            return source
        existing = self._member_tables.get(table)
        if existing is not None:
            return existing
        name = f"_MEMBERS_{len(self._member_tables)}"
        self._member_tables[table] = name
        return name

    def emit_lines(self) -> list[str]:
        """Return module-level assignments for interned keyed-read metadata."""
        lines: list[str] = []
        for values, name in self._sequences.items():
            lines.append(f"{name} = {values!r}")
        lines.extend(self._index_intern.emit_lines())
        for table, name in self._member_tables.items():
            lines.append(f"{name} = {slot_table_to_source(table)}")
        for series_id, name in self._slot_maps.items():
            domain = self.plan.series_expr[series_id]
            lines.append(f"{name} = {{key: i for i, key in enumerate({domain})}}")
        return lines


_KEYED_INTERN: ContextVar[KeyedReadIntern | None] = ContextVar(
    "excel_grapher_inverted_tree_keyed_intern",
    default=None,
)


def current_keyed_intern() -> KeyedReadIntern | None:
    """Keyed-read intern for the current emit walk, if any."""
    return _KEYED_INTERN.get()


@dataclass
class LookupReuse:
    """Intern positional tables and identical eager lookups for one helper.

    Parameter-only tables are assigned once per helper call. Identical
    `xl_vlookup` (and similar) calls in one eager formula share a local.
    Lazy IF/IFERROR operands are not interned, so unused branches stay cold
    and error-catching lambdas keep their boundaries (#796).
    """

    _tables: dict[str, str] = field(default_factory=dict)
    _table_order: list[tuple[str, str]] = field(default_factory=list)
    _lookup_expr_to_name: dict[str, str] = field(default_factory=dict)
    _lookup_count: dict[str, int] = field(default_factory=dict)
    _lookup_expr: dict[str, str] = field(default_factory=dict)
    _lookup_order: list[str] = field(default_factory=list)

    def intern_table(self, expr: str) -> str:
        """Return a helper-local name for a stable positional table `expr`."""
        existing = self._tables.get(expr)
        if existing is not None:
            return existing
        name = f"_TABLE_{len(self._tables)}"
        self._tables[expr] = name
        self._table_order.append((name, expr))
        return name

    def begin_formula(self) -> None:
        """Reset per-formula lookup CSE state."""
        self._lookup_expr_to_name.clear()
        self._lookup_count.clear()
        self._lookup_expr.clear()
        self._lookup_order.clear()

    def intern_lookup(self, expr: str) -> str:
        """Return a formula-local name for an eager lookup `expr`."""
        existing = self._lookup_expr_to_name.get(expr)
        if existing is not None:
            self._lookup_count[existing] += 1
            return existing
        name = f"_LOOKUP_{len(self._lookup_order)}"
        self._lookup_expr_to_name[expr] = name
        self._lookup_count[name] = 1
        self._lookup_expr[name] = expr
        self._lookup_order.append(name)
        return name

    def finish_formula(self, expr: str) -> tuple[str, tuple[tuple[str, str], ...]]:
        """Inline singleton lookups and return repeated-lookup assignments.

        Nested interned names are rewritten innermost-first so a unique
        inner lookup does not leak into a kept outer assignment.
        """
        resolved = dict(self._lookup_expr)
        for name in self._lookup_order:
            text = resolved[name]
            for inner in self._lookup_order:
                if inner == name or self._lookup_count[inner] > 1:
                    continue
                text = _replace_ident(text, inner, resolved[inner])
            resolved[name] = text
        result = expr
        kept: list[tuple[str, str]] = []
        for name in self._lookup_order:
            text = resolved[name]
            if self._lookup_count[name] <= 1:
                result = _replace_ident(result, name, text)
            else:
                kept.append((name, text))
        return result, tuple(kept)

    def emit_table_lines(self, indent: str = "    ") -> list[str]:
        """Return helper-level assignments for interned positional tables."""
        return [f"{indent}{name} = {expr}" for name, expr in self._table_order]


_LOOKUP_REUSE: ContextVar[LookupReuse | None] = ContextVar(
    "excel_grapher_inverted_tree_lookup_reuse",
    default=None,
)


def current_lookup_reuse() -> LookupReuse | None:
    """Lookup intern for the current helper emit, if any."""
    return _LOOKUP_REUSE.get()


def _replace_ident(text: str, name: str, replacement: str) -> str:
    """Replace identifier `name` in `text` with `replacement`."""
    return re.sub(rf"\b{re.escape(name)}\b", lambda _match: replacement, text)


def _reuse_lookup(expr: str, ctx: EmitContext) -> str:
    """Intern `expr` when it is an eager lookup in the current helper."""
    reuse = current_lookup_reuse()
    if reuse is None or not ctx.eager or ctx.fused_use_area:
        return expr
    return reuse.intern_lookup(expr)


@dataclass
class EmitContext:
    """How to read bound series while lowering one host formula."""

    host: BoundSeries
    catalog: SeriesCatalog
    deps: SeriesDeps
    host_index: int
    host_cell: CanonicalAddress
    index_var: str | None
    prior_var: str | None
    used_runtime: set[str] = field(default_factory=set)
    scc_ids: frozenset[str] = field(default_factory=frozenset)
    instance_mode: bool = False
    compute_names: dict[str, str] = field(default_factory=dict)
    fused_mode: bool = False
    fused_plan: FusedPlan | None = None
    fused_ready: frozenset[str] = field(default_factory=frozenset)
    fused_buffer_suffix: str = ""
    fused_partition: tuple[Scalar, ...] | None = None
    fused_use_area: bool = False
    graph: DependencyGraph | None = None
    lookup_anchor_slot: int = 0
    blank_rects: tuple[BlankRangeRect, ...] = field(default_factory=current_blank_rects)
    array_context: bool = False
    eager: bool = True
    coordinate_vars: dict[str, str] | None = None
    named_axes: NamedAxes | None = None

    def param(self, series_id: str) -> str:
        return series_id

    def use(self, symbol: str) -> str:
        self.used_runtime.add(symbol)
        return symbol


def python_measure_type(series: BoundSeries) -> str:
    """Return the Python type of one observation (`float | str` for numbers)."""
    base = f"{series.python_dtype} | str" if series.python_dtype != "str" else series.python_dtype
    if series.has_none_holes or series.layout != "scalar":
        return f"{base} | None"
    return base


def _number_literal(value: float) -> str:
    if value == int(value) and abs(value) < 1e15:
        as_int = int(value)
        if float(as_int) == value:
            return repr(float(as_int)) if value != as_int else repr(as_int)
    return repr(value)


def emit_expr(node: AstNode, ctx: EmitContext) -> str:
    """Lower `node` to a Python expression against `ctx` parameters."""
    match node:
        case NumberNode(value):
            return _number_literal(value)
        case StringNode(value):
            return repr(value)
        case BoolNode(value):
            return "True" if value else "False"
        case ErrorNode(error):
            return f"{ctx.use('xl_raise')}({error.value!r})"
        case EmptyArgNode():
            return "0"
        case CellRefNode():
            return _emit_cell_ref(node, ctx)
        case RangeNode():
            raise InvertedTreeExportError(
                f"series {ctx.host.series_id!r}: bare range in value position"
            )
        case BinaryOpNode():
            return _emit_binary(node, ctx)
        case UnaryOpNode():
            return _emit_unary(node, ctx)
        case FunctionCallNode():
            return _emit_function(node, ctx)
        case _:
            raise InvertedTreeExportError(
                f"series {ctx.host.series_id!r}: unsupported AST node {type(node).__name__}"
            )


def _is_scan_prior_ref(address: CanonicalAddress, ctx: EmitContext) -> bool:
    """Return whether an address is the preceding self observation.

    External seeds retain their producer-coordinate reads. Substituting
    `prior` for an external read would turn a formula family over several
    producer coordinates into a recurrence.
    """
    pred = predecessor_address(ctx.host, ctx.host_index, ctx.catalog, ctx.graph)
    if pred is not None and address == pred:
        return ctx.host_index > 0
    succ = successor_address(ctx.host, ctx.host_index, ctx.catalog, ctx.graph)
    if succ is not None and address == succ:
        return ctx.host_index < len(ctx.host.cells) - 1
    return False


def _emit_cell_ref(node: CellRefNode, ctx: EmitContext) -> str:
    return _emit_address(as_canonical(resolve_cell_ref(node, ctx.host_cell)), ctx, ref=node)


def _current_statement(ctx: EmitContext) -> Statement | None:
    """Return the host statement covering `ctx.host_index`."""
    index = ctx.host_index
    for stmt in ctx.host.statements:
        if stmt.start <= index < stmt.stop:
            return stmt
    return next(
        (stmt for stmt in ctx.host.statements if ctx.host_cell in stmt.cells),
        None,
    )


def _statement_cells(ctx: EmitContext) -> tuple[CanonicalAddress, ...] | None:
    stmt = _current_statement(ctx)
    return None if stmt is None else stmt.cells


def _static_catalog_literal(
    owner: BoundSeries,
    address: CanonicalAddress,
    ctx: EmitContext,
    ref: CellRefNode | None,
) -> str | None:
    """Return a literal catalog index when `ref` is static with `coeff = 0`."""
    if (
        ref is None
        or ctx.graph is None
        or ctx.fused_mode
        or owner.is_scalar
        or owner.series_id == ctx.host.series_id
        or owner.series_id in ctx.deps.aligned_ids
    ):
        return None
    idx = owner.index_of(address)
    if idx is None:
        return None
    try:
        access = classify_cell_ref_access(
            ctx.host,
            owner,
            ctx.catalog,
            ctx.graph,
            host_cell=ctx.host_cell,
            ref=ref,
            cells=_statement_cells(ctx),
        )
        coeff, offset = catalog_index_affine(access)
    except InvertedTreeExportError:
        return None
    if coeff != 0:
        return None
    return str(offset)


def _host_follow_for(owner: BoundSeries, ctx: EmitContext) -> dict[str, dict[object, object]]:
    """Return hostâ†’producer key maps for `owner` from resolved host edges."""
    return _host_follow_key_maps(
        ctx.host, owner, _host_producer_slots(ctx.host, owner, ctx.deps.edges)
    )


def _host_pair_for(
    owner: BoundSeries, ctx: EmitContext
) -> dict[str, dict[object, tuple[object, object]]]:
    """Return host->`(then, else)` pair maps for `owner` from host edges."""
    return _host_follow_pair_maps(
        ctx.host,
        owner,
        _host_producer_slots(ctx.host, owner, ctx.deps.edges),
        ctx.graph,
    )


def _lookup_axes_for(owner: BoundSeries, ctx: EmitContext) -> frozenset[str]:
    """Return lookup-axis fields for `owner` from resolved host edges."""
    return _lookup_axis_fields(
        ctx.host, owner, _host_producer_slots(ctx.host, owner, ctx.deps.edges)
    )


def _follow_is_identity(
    host_follow: Mapping[str, Mapping[object, object]], fields: Sequence[str]
) -> bool:
    """True when every `fields` entry is missing from `host_follow` or is identity."""
    for name in fields:
        mapped = host_follow.get(name)
        if mapped is not None and any(key != item for key, item in mapped.items()):
            return False
    return True


def _remap_host_key(
    value: object, field: str, host_follow: Mapping[str, Mapping[object, object]]
) -> object:
    """Rewrite `value` through `host_follow[field]`, or return `value` unchanged."""
    mapped = host_follow.get(field)
    if mapped is None:
        return value
    return mapped.get(value, value)


def _host_key_value_expr(
    field: str,
    ctx: EmitContext,
    *,
    host_follow: Mapping[str, Mapping[object, object]] | None = None,
) -> str:
    """Return a Python expr for `host[field]` as `index_var` walks the host.

    When `host_follow` remaps this field onto a distinct producer vocabulary
    (`B1` -> `Bounds Test 1: â€¦`), emit the producer values in host order.
    """
    follow = host_follow or {}

    def _value(raw: object) -> object:
        return _remap_host_key(raw, field, follow)

    if ctx.index_var is None or field not in ctx.host.key_fields:
        return repr(_value(ctx.host.domain[ctx.host_index][field]))
    fields = ctx.host.key_fields
    points = series_domain_points(ctx.host)
    if len(fields) == 1:
        column = tuple(_value(point) for point in points)
        hint = field if column == points else None
        return f"{_interned_sequence(column, field=hint)}[{ctx.index_var}]"
    pos = fields.index(field)
    raw_column = tuple(point[pos] if isinstance(point, tuple) else point for point in points)
    column = tuple(_value(item) for item in raw_column)
    hint = field if column == raw_column else None
    return f"{_interned_sequence(column, field=hint)}[{ctx.index_var}]"


def _pair_key_value_expr(
    field: str,
    branch: int,
    ctx: EmitContext,
    pair_maps: Mapping[str, Mapping[object, tuple[object, object]]],
) -> str:
    """Return a Python expr for `pair_maps[field][host[field]][branch]`.

    The pair table covers the current host statement only. Catalog index
    `index_var` is the host origin, so a statement that starts at `start`
    is subscripted as `table[i - start]`.
    """
    mapped = pair_maps.get(field)
    if mapped is None:
        raise InvertedTreeExportError(f"series {ctx.host.series_id!r}: no IF pair map for {field}")
    if ctx.index_var is None or field not in ctx.host.key_fields:
        try:
            host_value = ctx.host.domain[ctx.host_index][field]
        except KeyError as exc:
            raise InvertedTreeExportError(
                f"series {ctx.host.series_id!r}: host member {ctx.host_index} "
                f"has no {field} for an IF pair read"
            ) from exc
        pair = mapped.get(host_value)
        if pair is None:
            raise InvertedTreeExportError(
                f"series {ctx.host.series_id!r}: host {field}={host_value!r} "
                f"has no IF pair for {field}"
            )
        return repr(pair[branch])
    stmt = _current_statement(ctx)
    domain = ctx.host.domain if stmt is None else stmt.domain
    origin = 0 if stmt is None else stmt.start
    column: list[tuple[object, object]] = []
    for point in domain:
        try:
            host_value = point[field]
        except KeyError as exc:
            raise InvertedTreeExportError(
                f"series {ctx.host.series_id!r}: host walk is missing {field} for an IF pair read"
            ) from exc
        pair = mapped.get(host_value)
        if pair is None:
            raise InvertedTreeExportError(
                f"series {ctx.host.series_id!r}: host {field}={host_value!r} "
                f"has no IF pair for {field}"
            )
        column.append(pair)
    index_expr = _index_expr(-origin, ctx.index_var)
    return f"{_interned_sequence(tuple(column))}[{index_expr}][{branch}]"


def _producer_field_binding(
    owner: BoundSeries,
    address: CanonicalAddress,
    ctx: EmitContext,
    ref: CellRefNode | None = None,
    *,
    host_follow: Mapping[str, Mapping[object, object]] | None = None,
    pair_maps: Mapping[str, Mapping[object, tuple[object, object]]] | None = None,
) -> dict[str, str | tuple[str, object]] | None:
    """Map each producer key field to `host` or a literal from `address`."""
    idx = owner.index_of(address)
    if idx is None:
        return None
    pinned = _ref_pinned_fields(ref, ctx.host_cell, owner) if ref is not None else frozenset()
    follow = host_follow if host_follow is not None else _host_follow_for(owner, ctx)
    pairs = pair_maps if pair_maps is not None else _host_pair_for(owner, ctx)
    raw = _field_binding(
        ctx.host,
        ctx.host_index,
        owner,
        idx,
        pinned_fields=pinned,
        literal_fields=_lookup_axes_for(owner, ctx),
        host_follow=follow,
        pair_maps=pairs,
    )
    if raw is None:
        return None
    binding: dict[str, str | tuple[str, object]] = {}
    for key, spec in raw:
        if spec == "host":
            binding[key] = "host"
        elif isinstance(spec, tuple) and len(spec) == 2:
            kind, payload = spec
            if kind == "lit":
                binding[key] = ("lit", payload)
            elif kind == "pair":
                binding[key] = ("pair", payload)
            else:
                raise InvertedTreeExportError(
                    f"series {ctx.host.series_id!r}: invalid keyed binding {spec!r} for {key}"
                )
        else:
            raise InvertedTreeExportError(
                f"series {ctx.host.series_id!r}: invalid keyed binding {spec!r} for {key}"
            )
    return binding


def _expected_producer_point(
    owner: BoundSeries,
    binding: dict[str, str | tuple[str, object]],
    host_index: int,
    host: BoundSeries,
    *,
    host_follow: Mapping[str, Mapping[object, object]] | None = None,
    pair_maps: Mapping[str, Mapping[object, tuple[object, object]]] | None = None,
) -> object:
    values: list[object] = []
    host_point = host.domain[host_index]
    follow = host_follow or {}
    pairs = pair_maps or {}
    for key_name in owner.key_fields:
        spec = binding[key_name]
        if spec == "host":
            values.append(_remap_host_key(host_point[key_name], key_name, follow))
        elif isinstance(spec, tuple) and spec[0] == "pair":
            try:
                host_value = host_point[key_name]
            except KeyError as exc:
                raise InvertedTreeExportError(
                    f"series {host.series_id!r}: host member {host_index} "
                    f"has no {key_name} for an IF pair read"
                ) from exc
            pair = pairs.get(key_name, {}).get(host_value)
            if pair is None:
                raise InvertedTreeExportError(
                    f"series {host.series_id!r}: host {key_name}={host_value!r} "
                    f"has no IF pair for {key_name}"
                )
            if spec[1] == 0:
                values.append(pair[0])
            elif spec[1] == 1:
                values.append(pair[1])
            else:
                raise InvertedTreeExportError(
                    f"series {host.series_id!r}: invalid IF pair branch {spec[1]!r}"
                )
        elif isinstance(spec, tuple) and spec[0] == "lit":
            values.append(spec[1])
        else:
            raise InvertedTreeExportError(
                f"series {host.series_id!r}: invalid keyed binding {spec!r} for {key_name}"
            )
    return values[0] if len(values) == 1 else tuple(values)


def _verify_keyed_binding(
    owner: BoundSeries,
    binding: dict[str, str | tuple[str, object]],
    ctx: EmitContext,
    *,
    host_follow: Mapping[str, Mapping[object, object]] | None = None,
    pair_maps: Mapping[str, Mapping[object, tuple[object, object]]] | None = None,
) -> None:
    """Fail closed when a host member has no producer cell for `binding`.

    Only members of the current statement that share this formula's
    host-followed keys are checked. A literal year listed by IDA need
    not exist on IMF (#760).
    """
    domain = series_domain_points(owner)
    known = set(domain)
    seen: set[int] = set()
    follow = host_follow if host_follow is not None else _host_follow_for(owner, ctx)
    pairs = pair_maps if pair_maps is not None else _host_pair_for(owner, ctx)
    follow_fields = tuple(
        key
        for key, spec in binding.items()
        if spec == "host" or (isinstance(spec, tuple) and spec[0] == "pair")
    )
    current_part = _host_join_key(ctx.host, ctx.host_index, follow_fields)
    stmt = _current_statement(ctx)
    for edge in ctx.deps.edges:
        if edge.producer_id != owner.series_id or edge.consumer_id != ctx.host.series_id:
            continue
        host_i = ctx.host.index_of(edge.consumer_cell)
        if host_i is None or host_i in seen or host_i >= len(ctx.host.domain):
            continue
        if stmt is not None and not (stmt.start <= host_i < stmt.stop):
            continue
        if follow_fields and _host_join_key(ctx.host, host_i, follow_fields) != current_part:
            continue
        seen.add(host_i)
        expected = _expected_producer_point(
            owner, binding, host_i, ctx.host, host_follow=follow, pair_maps=pairs
        )
        if expected not in known:
            raise InvertedTreeExportError(
                f"series {ctx.host.series_id!r} cell {ctx.host.cells[host_i]} "
                f"has no {owner.series_id!r} member {expected!r}"
            )


def _interned_sequence(values: tuple[object, ...], *, field: str | None = None) -> str:
    """Return `values` as an interned name, a `data` domain, or a tuple literal."""
    intern = current_keyed_intern()
    if intern is None:
        return repr(values)
    return intern.sequence(values, field=field)


def _catalog_indices_expr(indices: Sequence[int]) -> str:
    """Return a compact or interned expression for catalog-index `indices`."""
    intern = current_keyed_intern()
    items = tuple(indices)
    if intern is None:
        return indices_to_source(items)
    return intern.slots(items)


def _keyed_index_from_catalog_pairs(
    owner: BoundSeries,
    address: CanonicalAddress,
    ctx: EmitContext,
    ref: CellRefNode | None,
) -> str | None:
    """Return an affine or interned catalog-index expression when the site is static.

    Statement-local `(host, producer)` pairs are preferred over embedding the
    producer domain and calling `.index`. A one-cell host (no loop) uses the
    precomputed catalog slot.
    """
    if ctx.index_var is None:
        idx = owner.index_of(address)
        return None if idx is None else str(idx)
    if ref is None or ctx.graph is None:
        return None
    try:
        pairs = cell_ref_catalog_pairs(
            ctx.host,
            owner,
            ctx.graph,
            host_cell=ctx.host_cell,
            ref=ref,
            cells=_statement_cells(ctx),
        )
    except InvertedTreeExportError:
        return None
    if len(pairs) >= 2:
        fitted = fit_affine_map(pairs)
        if fitted is not None:
            return _linear_index_expr(fitted[0], fitted[1], ctx.index_var)
        stmt = _current_statement(ctx)
        members = ctx.host.cells if stmt is None else stmt.cells
        origin = 0 if stmt is None else stmt.start
        by_host = dict(pairs)
        table: list[int] = []
        for cell in members:
            host_i = ctx.host.index_of(cell)
            if host_i is None or host_i not in by_host:
                return None
            table.append(by_host[host_i])
        slots = _catalog_indices_expr(table)
        return f"{slots}[{_index_expr(-origin, ctx.index_var)}]"
    if len(pairs) == 1:
        return str(pairs[0][1])
    return None


def _keyed_catalog_index_expr(
    owner: BoundSeries,
    address: CanonicalAddress,
    ctx: EmitContext,
    ref: CellRefNode | None = None,
) -> str:
    """Return a catalog index for a keyed multi-read of `address`.

    Prefers an affine map or interned slot table. Otherwise looks up a shared
    key-to-slot map instead of embedding `domain.index(...)` at each site.
    """
    follow = _host_follow_for(owner, ctx)
    pairs = _host_pair_for(owner, ctx)
    binding = _producer_field_binding(
        owner, address, ctx, ref=ref, host_follow=follow, pair_maps=pairs
    )
    if binding is None:
        raise InvertedTreeExportError(
            f"series {ctx.host.series_id!r}: cannot emit keyed read of "
            f"{owner.series_id!r} at {address}"
        )
    _verify_keyed_binding(owner, binding, ctx, host_follow=follow, pair_maps=pairs)
    static = _keyed_index_from_catalog_pairs(owner, address, ctx, ref)
    if static is not None:
        return static
    intern = current_keyed_intern()
    if intern is not None:
        map_name = intern.slot_map(owner.series_id)
    else:
        domain = series_domain_points(owner)
        map_name = f"{{key: i for i, key in enumerate({domain!r})}}"
    if (
        ctx.index_var is not None
        and owner.key_fields == ctx.host.key_fields
        and all(binding[key_name] == "host" for key_name in owner.key_fields)
        and _follow_is_identity(follow, owner.key_fields)
    ):
        host_domain = series_domain_points(ctx.host)
        return f"{map_name}[{_interned_sequence(host_domain)}[{ctx.index_var}]]"
    key_parts: list[str] = []
    for key_name in owner.key_fields:
        spec = binding[key_name]
        if spec == "host":
            key_parts.append(_host_key_value_expr(key_name, ctx, host_follow=follow))
        elif isinstance(spec, tuple) and spec[0] == "pair":
            if spec[1] == 0:
                key_parts.append(_pair_key_value_expr(key_name, 0, ctx, pairs))
            elif spec[1] == 1:
                key_parts.append(_pair_key_value_expr(key_name, 1, ctx, pairs))
            else:
                raise InvertedTreeExportError(
                    f"series {ctx.host.series_id!r}: invalid IF pair branch {spec[1]!r}"
                )
        elif isinstance(spec, tuple) and spec[0] == "lit":
            key_parts.append(repr(spec[1]))
        else:
            raise InvertedTreeExportError(
                f"series {ctx.host.series_id!r}: invalid keyed binding {spec!r} for {key_name}"
            )
    key_expr = key_parts[0] if len(key_parts) == 1 else f"({', '.join(key_parts)})"
    return f"{map_name}[{key_expr}]"


def _emit_address(
    address: CanonicalAddress, ctx: EmitContext, *, ref: CellRefNode | None = None
) -> str:
    if address_in_blank_ranges(address, ctx.blank_rects):
        return "None"
    if ctx.coordinate_vars is not None:
        return _emit_named_address(address, ctx, ref=ref)
    if ctx.fused_mode:
        return _emit_fused_ref(address, ctx, ref=ref)
    if ctx.instance_mode:
        return _emit_instance_ref(address, ctx, ref=ref)
    if ctx.prior_var and _is_scan_prior_ref(address, ctx):
        return ctx.prior_var
    owner = ctx.catalog.require_series_for(address)
    if owner.series_id == ctx.host.series_id:
        if ctx.prior_var is not None:
            return ctx.prior_var
        raise InvertedTreeExportError(
            f"series {ctx.host.series_id!r}: self-ref {address} without a scan prior"
        )
    name = ctx.param(owner.series_id)
    if owner.is_scalar:
        return name
    literal = _static_catalog_literal(owner, address, ctx, ref)
    if literal is not None:
        return f"{name}[{literal}]"
    if owner.series_id in ctx.deps.keyed_ids:
        return f"{name}[{_keyed_catalog_index_expr(owner, address, ctx, ref=ref)}]"
    idx = owner.index_of(address)
    if owner.series_id in ctx.deps.lagged_ids and ctx.index_var is not None and idx is not None:
        offset = idx - ctx.host_index
        if offset == 0:
            return f"{name}[{ctx.index_var}]"
        if offset > 0:
            return f"{name}[{ctx.index_var} + {offset}]"
        return f"{name}[{ctx.index_var} - {-offset}]"
    if owner.series_id in ctx.deps.aligned_ids:
        if ctx.index_var is not None:
            return f"{name}[{ctx.index_var}]"
        if idx is not None:
            return f"{name}[{_aligned_taken_index(owner.series_id, idx, ctx)}]"
        return name
    if idx is not None and ctx.index_var is not None and owner.is_sequence:
        return f"{name}[{_instance_index_expr(owner, idx, ctx, ref)}]"
    if idx is not None and ctx.index_var is None:
        return f"{name}[{idx}]" if not owner.is_scalar else name
    if owner.series_id in ctx.deps.lookup_ids:
        return name
    return name


def _emit_named_address(
    address: CanonicalAddress, ctx: EmitContext, *, ref: CellRefNode | None = None
) -> str:
    """Read a producer by semantic coordinate relative to the host coordinate."""
    owner = ctx.catalog.require_series_for(address)
    name = ctx.param(owner.series_id)
    if owner.layout == "scalar":
        # A scalar member of the recurrence group is a demand-driven reader.
        return f"{name}[()]" if owner.series_id in ctx.scc_ids else name
    index = owner.index_of(address)
    if index is None:
        raise InvertedTreeExportError(f"series {owner.series_id!r}: no coordinate for {address}")
    return f"{name}[{', '.join(_named_keys(owner, owner.domain[index], ctx, ref=ref))}]"


def _pinned_axes(ref: CellRefNode | None, host_cell: CanonicalAddress) -> set[str]:
    """Worksheet axes on which `ref` is absolute and differs from the host."""
    if ref is None:
        return set()
    address = resolve_cell_ref(ref, host_cell)
    _host_sheet, host_row, host_col = parse_cell_coords(host_cell)
    _sheet, row, col = parse_cell_coords(address)
    pinned: set[str] = set()
    if col != host_col and isinstance(ref.ref.col, AbsoluteAxis):
        pinned.add("col")
    if row != host_row and isinstance(ref.ref.row, AbsoluteAxis):
        pinned.add("row")
    return pinned


def _named_keys(
    owner: BoundSeries,
    point: KeyPoint,
    ctx: EmitContext,
    *,
    ref: CellRefNode | None = None,
) -> list[str]:
    """Express each key of `point` relative to the host loop variables.

    A key equal to the host's own key is the loop variable. An integer
    `TIME_PERIOD` reached through a relative reference is the loop variable
    plus the authored difference; formula-family grouping then verifies the
    same difference at every coordinate sharing the expression. Keys reached
    through an absolute (`$`) reference, or on other axes, stay literal.
    """
    from excel_grapher.exporter.inverted_tree.deps import _key_field_axis

    assert ctx.coordinate_vars is not None
    host_point = ctx.host.domain[ctx.host_index].as_mapping()
    pinned = _pinned_axes(ref, ctx.host_cell)
    keys = []
    for key_field in owner.key_fields:
        variable = ctx.coordinate_vars.get(key_field)
        current = host_point.get(key_field)
        target = point[key_field]
        if variable is not None and current == target:
            keys.append(variable)
        elif (
            variable is not None
            and key_field == "TIME_PERIOD"
            and type(current) is int
            and type(target) is int
            and _key_field_axis(owner, key_field) not in pinned
        ):
            difference = target - current
            keys.append(f"{variable} {'+' if difference > 0 else '-'} {abs(difference)}")
        else:
            keys.append(repr(target))
    return keys


def _aligned_taken_index(producer_id: str, catalog_idx: int, ctx: EmitContext) -> int:
    """Return `catalog_idx` in the window `_aligned_call_arg` takes to.

    Aligned arguments are remapped into the host's index space. A non-looping
    helper (`index_var` is `None`) and a rung-3 instance subscript must honour
    that window, not the producer catalog slot.
    """
    index_map = ctx.deps.index_maps.get(producer_id)
    if index_map is None:
        raise InvertedTreeExportError(
            f"series {ctx.host.series_id!r}: aligned {producer_id!r} has no index map"
        )
    try:
        return index_map.index(catalog_idx)
    except ValueError:
        raise InvertedTreeExportError(
            f"series {ctx.host.series_id!r}: {producer_id}[{catalog_idx}] "
            f"is outside the aligned window {index_map}"
        ) from None


def _index_expr(offset: int, index_var: str) -> str:
    if offset == 0:
        return index_var
    if offset > 0:
        return f"{index_var} + {offset}"
    return f"{index_var} - {-offset}"


def _affine_index_expr(offset: int, index_var: str, *, step: int) -> str:
    """Return `offset + step * index_var` for a fused live-measure subscript."""
    if step == 1:
        return _index_expr(offset, index_var)
    if step == -1:
        if offset == 0:
            return f"-{index_var}"
        return f"{offset} - {index_var}"
    raise InvertedTreeExportError(f"unsupported fused index step {step}")


def _union_t(plan: FusedPlan, address: CanonicalAddress, ctx: EmitContext) -> int:
    """Return the union index of `address`, or fail closed naming the host."""
    coord = schedule_axis_coord(address, ctx.catalog)
    mapped = plan.coord_to_t.get(coord)
    if mapped is None:
        raise InvertedTreeExportError(
            f"series {ctx.host.series_id!r}: fused ref {address} is not on the union schedule"
        )
    return mapped


def _host_outer_fields(host: BoundSeries, catalog: SeriesCatalog) -> tuple[str, ...]:
    """Return the host's instance-partition key fields (`TIME_PERIOD` stripped)."""
    fields = preferred_fields(host, catalog)
    if fields is None:
        return ()
    return tuple(name for name in fields if name != "TIME_PERIOD")


def _producer_partition_of(
    owner: BoundSeries,
    address: CanonicalAddress,
    catalog: SeriesCatalog,
    host_outer: Sequence[str],
) -> tuple[Scalar, ...] | None:
    """Return the plan-partition key of `address`, including REF_AREA-only seeds."""
    part = schedule_partition(address, catalog)
    if part:
        return part
    idx = owner.index_of(address)
    if idx is None or idx >= len(owner.domain) or not host_outer:
        return None
    point = owner.domain[idx]
    try:
        return tuple(point[field] for field in host_outer)
    except KeyError:
        return None


def _aligned_area_index(
    owner: BoundSeries,
    address: CanonicalAddress,
    ctx: EmitContext,
    plan: FusedPlan,
) -> tuple[int, int] | None:
    """Return `(area_index, stride)` when `address` lines up with `plan.partitions`.

    A producer keyed by the outer partition fields â€” with or without
    `TIME_PERIOD` â€” has one block per area. Uniform block length is the
    catalog stride so `_area * stride + t` is identical across partitions.
    """
    host_outer = _host_outer_fields(ctx.host, ctx.catalog)
    if not host_outer or not plan.partitions:
        return None
    prod_part = _producer_partition_of(owner, address, ctx.catalog, host_outer)
    if prod_part not in plan.partitions:
        return None
    counts = [
        sum(
            1
            for cell in owner.cells
            if _producer_partition_of(owner, cell, ctx.catalog, host_outer) == part
        )
        for part in plan.partitions
    ]
    if not counts or counts[0] == 0 or len(set(counts)) != 1:
        return None
    return plan.partitions.index(prod_part), counts[0]


def _combine_area_index(area_expr: str, inner: str) -> str:
    """Join `_area * stride` with a schedule-axis term."""
    if inner in {"0", "0.0"}:
        return area_expr
    if area_expr in {"0", "0.0"}:
        return inner
    if inner.startswith("-"):
        return f"{area_expr} - {inner[1:]}"
    return f"{area_expr} + {inner}"


def _area_stride_expr(stride: int, delta_area: int) -> str:
    """Return `(_area + delta_area) * stride`."""
    area = "_area" if delta_area == 0 else _index_expr(delta_area, "_area")
    if stride == 1:
        return area if area == "_area" else f"({area})"
    if area == "_area":
        return f"_area * {stride}"
    return f"({area}) * {stride}"


def _emit_fused_ref(
    address: CanonicalAddress, ctx: EmitContext, *, ref: CellRefNode | None = None
) -> str:
    owner = ctx.catalog.require_series_for(address)
    literal = _static_catalog_literal(owner, address, ctx, ref)
    if literal is not None:
        ctx.use("live_measure")
        return f"live_measure({ctx.param(owner.series_id)}[{literal}])"
    if owner.series_id in ctx.deps.keyed_ids and owner.series_id not in ctx.scc_ids:
        ctx.use("live_measure")
        keyed_ctx = ctx
        if ctx.fused_plan is not None:
            # Keyed maps consume a host catalog slot. The fused loop variable
            # is a local schedule slot, including in an unrolled partition.
            step = -1 if ctx.fused_plan.direction == "reversed" else 1
            host_union = _union_t(ctx.fused_plan, ctx.host_cell, ctx)
            catalog_index = _affine_index_expr(
                ctx.host_index - step * host_union, ctx.index_var or "t", step=step
            )
            keyed_ctx = replace(ctx, index_var=f"({catalog_index})")
        return f"live_measure({ctx.param(owner.series_id)}[{_keyed_catalog_index_expr(owner, address, keyed_ctx, ref=ref)}])"
    idx = owner.index_of(address)
    if idx is None or ctx.fused_plan is None:
        raise InvertedTreeExportError(
            f"series {ctx.host.series_id!r}: fused ref {address} is unbound"
        )
    plan = ctx.fused_plan
    index_var = ctx.index_var or "t"
    ctx.use("live_measure")
    suffix = ctx.fused_buffer_suffix
    if owner.series_id in ctx.scc_ids:
        host_part = schedule_partition(ctx.host_cell, ctx.catalog)
        prod_part = schedule_partition(address, ctx.catalog)
        host_union = _union_t(plan, ctx.host_cell, ctx)
        prod_union = _union_t(plan, address, ctx)
        delta = prod_union - host_union
        if prod_part != host_part and plan.partitions:
            prod_start = plan.domain[owner.series_id][0]
            domain_len = plan.domain[owner.series_id][1] - prod_start
            prod_i = plan.partitions.index(prod_part)
            index_expr = _index_expr(prod_i * domain_len + delta - prod_start, index_var)
            return f"live_measure({owner.series_id}[{index_expr}])"
        local = f"{owner.series_id}{suffix}"
        if delta == 0:
            if owner.series_id not in ctx.fused_ready:
                raise InvertedTreeExportError(
                    f"series {ctx.host.series_id!r}: same-index read of "
                    f"{owner.series_id!r} before it is written"
                )
            return f"live_measure({local}_t)"
        prod_start = plan.domain[owner.series_id][0]
        index_expr = _index_expr(delta - prod_start, index_var)
        return f"live_measure({local}[{index_expr}])"
    name = ctx.param(owner.series_id)
    if owner.is_scalar:
        return f"live_measure({name})"
    if "TIME_PERIOD" not in ctx.host.key_fields and ref is not None:
        # Categorical schedule order need not advance through producer slots
        # in either catalog direction. Resolve each authored reference site.
        pairs = _instance_ref_index_pairs(owner, ctx, ref)
        schedule_pairs = [
            (plan.coord_to_t[schedule_axis_coord(ctx.host.cells[host_i], ctx.catalog)], producer_i)
            for host_i, producer_i in pairs
        ]
        if schedule_pairs:
            fitted = fit_affine_map(schedule_pairs)
            if fitted is not None:
                index_expr = _linear_index_expr(fitted[0], fitted[1], index_var)
            else:
                slots = [idx] * len(plan.coord_to_t)
                for position, producer_i in schedule_pairs:
                    slots[position] = producer_i
                index_expr = f"{_catalog_indices_expr(slots)}[{index_var}]"
            return f"live_measure({name}[{index_expr}])"
    host_union = _union_t(plan, ctx.host_cell, ctx)
    step = -1 if plan.direction == "reversed" else 1
    pinned = _ref_pinned_fields(ref, ctx.host_cell, owner) if ref is not None else frozenset()
    if ref is not None and ctx.graph is not None and "TIME_PERIOD" not in pinned:
        host_partition = schedule_partition(ctx.host_cell, ctx.catalog)
        members = tuple(
            cell
            for cell in ctx.host.cells
            if schedule_partition(cell, ctx.catalog) == host_partition
            and try_formula_ast(ctx.graph, cell) is not None
        )
        pairs = cell_ref_catalog_pairs(
            ctx.host, owner, ctx.graph, host_cell=ctx.host_cell, ref=ref, cells=members
        )
        if len(pairs) > 1 and {producer for _, producer in pairs} == {idx}:
            # Authored formulas can repeat a fixed reference without `$`.
            # Only the resolved dependencies establish that it stays fixed.
            pinned = pinned | {"TIME_PERIOD"}
    if ctx.fused_use_area:
        aligned = _aligned_area_index(owner, address, ctx, plan)
        if aligned is not None:
            area_i, stride = aligned
            host_part = schedule_partition(ctx.host_cell, ctx.catalog)
            host_i = plan.partitions.index(host_part) if host_part in plan.partitions else 0
            fields = preferred_fields(owner, ctx.catalog) or ()
            if "TIME_PERIOD" in fields and "TIME_PERIOD" not in pinned:
                local = idx - step * host_union - area_i * stride
                inner = _affine_index_expr(local, index_var, step=step)
            else:
                local = idx - area_i * stride
                inner = str(local)
            index_expr = _combine_area_index(_area_stride_expr(stride, area_i - host_i), inner)
            return f"live_measure({name}[{index_expr}])"
    index_expr = (
        str(idx)
        if "TIME_PERIOD" in pinned
        else _affine_index_expr(idx - step * host_union, index_var, step=step)
    )
    return f"live_measure({name}[{index_expr}])"


def _remap_instance_index(owner: BoundSeries, idx: int, ctx: EmitContext) -> int:
    """Map a producer catalog slot into the aligned argument window when needed."""
    if (
        owner.series_id not in ctx.scc_ids
        and owner.series_id in ctx.deps.aligned_ids
        and not owner.is_scalar
    ):
        return _aligned_taken_index(owner.series_id, idx, ctx)
    return idx


def _instance_ref_index_pairs(
    owner: BoundSeries,
    ctx: EmitContext,
    ref: CellRefNode | None,
) -> list[tuple[int, int]]:
    """Return `(host_index, producer_index)` for `ref` over the host statement."""
    if ref is None or ctx.graph is None:
        return []
    try:
        raw = cell_ref_catalog_pairs(
            ctx.host,
            owner,
            ctx.graph,
            host_cell=ctx.host_cell,
            ref=ref,
            cells=_statement_cells(ctx),
        )
    except InvertedTreeExportError:
        return []
    return [(host_i, _remap_instance_index(owner, prod_i, ctx)) for host_i, prod_i in raw]


def _instance_index_expr(
    owner: BoundSeries,
    idx: int,
    ctx: EmitContext,
    ref: CellRefNode | None,
) -> str:
    """Return the producer catalog index as a function of the host index.

    A formula-shape statement may map `i -> a*i + b` (opposite
    cross-partition reads) rather than `i + offset`. When the site is not
    affine, emit an explicit slot table indexed by `i`.
    """
    index_var = ctx.index_var or "i"
    idx = _remap_instance_index(owner, idx, ctx)
    pairs = _instance_ref_index_pairs(owner, ctx, ref)
    if len(pairs) >= 2:
        fitted = fit_affine_map(pairs)
        if fitted is not None:
            return _linear_index_expr(fitted[0], fitted[1], index_var)
        table = [idx] * len(ctx.host.cells)
        for host_i, prod_i in pairs:
            if 0 <= host_i < len(table):
                table[host_i] = prod_i
        return f"{_catalog_indices_expr(table)}[{index_var}]"
    return _index_expr(idx - ctx.host_index, index_var)


def _emit_instance_ref(
    address: CanonicalAddress, ctx: EmitContext, *, ref: CellRefNode | None = None
) -> str:
    owner = ctx.catalog.require_series_for(address)
    literal = _static_catalog_literal(owner, address, ctx, ref)
    if literal is not None and owner.series_id not in ctx.scc_ids:
        return f"{ctx.param(owner.series_id)}[{literal}]"
    if owner.series_id in ctx.deps.keyed_ids and owner.series_id not in ctx.scc_ids:
        return f"{ctx.param(owner.series_id)}[{_keyed_catalog_index_expr(owner, address, ctx, ref=ref)}]"
    idx = owner.index_of(address)
    if idx is None:
        raise InvertedTreeExportError(
            f"series {ctx.host.series_id!r}: instance ref {address} is unbound"
        )
    index_expr = _instance_index_expr(owner, idx, ctx, ref)
    if owner.series_id in ctx.scc_ids:
        fn = ctx.compute_names[owner.series_id]
        ctx.use("demand_instance")
        return f"demand_instance({owner.series_id!r}, {index_expr}, {fn}, memo, stack)"
    name = ctx.param(owner.series_id)
    if owner.is_scalar:
        return name
    return f"{name}[{index_expr}]"


def _emit_value_or_range(node: AstNode, ctx: EmitContext) -> str:
    """Emit a scalar expression or a positional range table."""
    if isinstance(node, (RangeNode, WholeColumnNode, WholeRowNode)):
        return _emit_range_table(node, ctx)
    return emit_expr(node, ctx)


def _emit_binary(node: BinaryOpNode, ctx: EmitContext) -> str:
    op = node.op
    if op == "&":
        left = emit_expr(node.left, ctx)
        right = emit_expr(node.right, ctx)
        return f"{ctx.use('xl_concat')}({left}, {right})"
    helper = _ARITHMETIC_HELPERS.get(op) or _COMPARE_HELPERS.get(op)
    if helper is not None:
        left = _emit_value_or_range(node.left, ctx)
        right = _emit_value_or_range(node.right, ctx)
        return f"{ctx.use(helper)}({left}, {right})"
    raise InvertedTreeExportError(f"series {ctx.host.series_id!r}: unsupported operator {op!r}")


def _emit_unary(node: UnaryOpNode, ctx: EmitContext) -> str:
    operand = emit_expr(node.operand, ctx)
    helper = _UNARY_HELPERS.get(node.op)
    if helper is not None:
        return f"{ctx.use(helper)}({operand})"
    if node.op == "%":
        return f"{ctx.use('xl_div')}({operand}, 100)"
    raise InvertedTreeExportError(
        f"series {ctx.host.series_id!r}: unsupported unary operator {node.op!r}"
    )


def _emit_function(node: FunctionCallNode, ctx: EmitContext) -> str:
    name = normalize_excel_function_name(node.name)
    if name in {"IFERROR", "IFNA", "ISERROR", "ISNA", "ISBLANK", "ISNUMBER"}:
        required = 2 if name in {"IFERROR", "IFNA"} else 1
        if len(node.args) != required:
            return f"{ctx.use('xl_raise')}('#VALUE!')"
        helper = "xl_isnumber_lazy" if name == "ISNUMBER" else f"xl_{name.lower()}"
        lazy = replace(ctx, eager=False)
        args = ", ".join(f"lambda: {emit_expr(arg, lazy)}" for arg in node.args)
        return f"{ctx.use(helper)}({args})"
    if name == "NA":
        return f"{ctx.use('xl_raise')}('#N/A')"
    if name == "IF":
        return _emit_if(node, ctx)
    if name == "CHOOSE":
        return _emit_choose(node, ctx)
    if name == "OFFSET":
        return _emit_offset(node, ctx)
    if name == "INDEX":
        return _emit_index(node, ctx)
    if name == "ROW":
        return _emit_row(node, ctx)
    if name == "COLUMN":
        return _emit_column(node, ctx)
    if name == "COLUMNS":
        return _emit_reference_geometry(node, ctx)
    if name == "INDIRECT":
        return _emit_indirect(node, ctx)
    if name == "MATCH":
        return _emit_match(node, ctx)
    if name == "TRUE":
        return "True"
    if name == "FALSE":
        return "False"
    if name in _RANGE_REDUCE_FUNCTIONS:
        return _emit_aggregate(node, ctx)
    if name in _LOOKUP_TABLE_FUNCTIONS:
        args = ", ".join(_emit_lookup_arg(arg, ctx) for arg in node.args)
    else:
        args = ", ".join(emit_expr(arg, ctx) for arg in node.args)
    func = f"xl_{name.lower()}"
    if func not in _RUNTIME_FUNCTIONS:
        raise InvertedTreeExportError(
            f"series {ctx.host.series_id!r}: Excel function {name} has no "
            "inverted-tree runtime helper"
        )
    ctx.use(func)
    call = f"{func}({args})"
    if name in _PURE_LOOKUP_FUNCTIONS:
        return _reuse_lookup(call, ctx)
    return call


def _emit_reference_geometry(node: FunctionCallNode, ctx: EmitContext) -> str:
    """Lower reference metadata without evaluating the referenced cell values."""
    name = normalize_excel_function_name(node.name)
    if len(node.args) > 1 or (name == "COLUMNS" and not node.args):
        return f"{ctx.use('xl_raise')}('#VALUE!')"
    ref = node.args[0] if node.args else None
    if ref is not None and not isinstance(
        ref, (CellRefNode, RangeNode, WholeColumnNode, WholeRowNode)
    ):
        raise _host_export_error(ctx, f"{name} requires a worksheet reference")

    def coordinate(host_cell: CanonicalAddress) -> int:
        if ref is None:
            _sheet, row, col = parse_cell_coords(host_cell)
            return row if name == "ROW" else col
        if isinstance(ref, CellRefNode):
            addresses = (as_canonical(resolve_cell_ref(ref, host_cell)),)
        elif isinstance(ref, RangeNode):
            addresses = (
                as_canonical(resolve_cell_ref(CellRefNode(ref.start_ref), host_cell)),
                as_canonical(resolve_cell_ref(CellRefNode(ref.end_ref), host_cell)),
            )
        else:
            addresses = tuple(iter_ref_addresses(ref, host_cell, ctx.graph))
        _sheet, row, col = parse_cell_coords(addresses[0])
        if name == "COLUMNS":
            return abs(parse_cell_coords(addresses[-1])[2] - col) + 1
        return row if name == "ROW" else col

    if ctx.index_var is None:
        return str(coordinate(ctx.host_cell))
    values = tuple(coordinate(cell) for cell in ctx.host.cells)
    if len(set(values)) == 1:
        return str(values[0])
    return f"{values!r}[{ctx.index_var}]"


def _emit_if(node: FunctionCallNode, ctx: EmitContext) -> str:
    if len(node.args) < 2:
        return f"{ctx.use('xl_raise')}('#VALUE!')"
    if _contains_array_if_operand(node):
        if not ctx.array_context:
            raise InvertedTreeExportError(
                f"series {ctx.host.series_id!r}: bare range in value position"
            )
        _assert_array_if_sound(node, ctx)
        cond = _emit_value_or_range(node.args[0], ctx)
        then = _emit_value_or_range(node.args[1], ctx)
        otherwise = _emit_value_or_range(node.args[2], ctx) if len(node.args) > 2 else "False"
        return f"{ctx.use('xl_if')}({cond}, {then}, {otherwise})"
    cond = emit_expr(node.args[0], ctx)
    then = emit_expr(node.args[1], replace(ctx, eager=False))
    # Array omitted else is Excel FALSE; scalar emit still uses 0.
    otherwise = emit_expr(node.args[2], replace(ctx, eager=False)) if len(node.args) > 2 else "0"
    return f"({then} if {cond} else {otherwise})"


def _contains_array_if_operand(node: AstNode) -> bool:
    """True when `node` has a range in element-wise IF/operator position.

    Ranges consumed by lookups or aggregates (`VLOOKUP`, `INDEX`, `SUM`, â€¦)
    are not array-`IF` operands; those functions keep their own lowering.
    """
    match node:
        case RangeNode() | WholeColumnNode() | WholeRowNode():
            return True
        case BinaryOpNode(left=left, right=right):
            return _contains_array_if_operand(left) or _contains_array_if_operand(right)
        case UnaryOpNode(operand=operand):
            return _contains_array_if_operand(operand)
        case FunctionCallNode(name=name, args=args):
            if normalize_excel_function_name(name) != "IF":
                return False
            return any(_contains_array_if_operand(arg) for arg in args)
        case _:
            return False


def _range_shape(node: RangeNode, ctx: EmitContext) -> tuple[int, int]:
    start = resolve_cell_ref(node.start_ref, ctx.host_cell)
    end = resolve_cell_ref(node.end_ref, ctx.host_cell)
    sheet1, row1, col1 = parse_cell_coords(start)
    sheet2, row2, col2 = parse_cell_coords(end)
    if sheet1 != sheet2:
        raise _host_export_error(ctx, "array IF does not support cross-sheet ranges")
    return (abs(row2 - row1) + 1, abs(col2 - col1) + 1)


def _assert_array_if_sound(node: FunctionCallNode, ctx: EmitContext) -> None:
    """Fail closed when array `IF` alignment is not element-wise.

    Supported interiors are ranges, cell refs, scalars, nested `IF`, and
    binary arithmetic/compare. Unary operators, concatenation, and other
    functions are rejected with an array-`IF`-specific message.
    """
    shapes: set[tuple[int, int]] = set()

    def walk(item: AstNode) -> None:
        if isinstance(item, (WholeColumnNode, WholeRowNode)):
            kind = "whole-column" if isinstance(item, WholeColumnNode) else "whole-row"
            raise _host_export_error(ctx, f"array IF does not support {kind} refs")
        if isinstance(item, RangeNode):
            shapes.add(_range_shape(item, ctx))
            return
        if isinstance(item, BinaryOpNode):
            if item.op not in _ARRAY_IF_VALUE_OPS:
                raise _host_export_error(ctx, f"array IF operator {item.op!r} is unsupported")
            walk(item.left)
            walk(item.right)
            return
        if isinstance(item, UnaryOpNode):
            raise _host_export_error(ctx, f"array IF unary {item.op!r} is unsupported")
        if isinstance(item, FunctionCallNode):
            name = normalize_excel_function_name(item.name)
            if name in _ARRAY_IF_UNSOUND_FNS:
                if name in {"AND", "OR"}:
                    detail = "AND/OR collapse"
                elif name in {"SUM", "SUMPRODUCT", "AVERAGE", "AGGREGATE"}:
                    detail = "nested aggregate"
                else:
                    detail = name
                raise _host_export_error(ctx, f"array IF {detail} is unsupported")
            if name != "IF":
                raise _host_export_error(ctx, f"array IF interior {name} is unsupported")
            for arg in item.args:
                walk(arg)

    for arg in node.args:
        walk(arg)
    if len(shapes) > 1:
        raise _host_export_error(ctx, f"array IF shape mismatch: {sorted(shapes)}")


def _emit_choose(node: FunctionCallNode, ctx: EmitContext) -> str:
    if len(node.args) < 2:
        return f"{ctx.use('xl_raise')}('#VALUE!')"
    index = emit_expr(node.args[0], ctx)
    lazy = replace(ctx, eager=False)
    choices = ", ".join(f"lambda: {emit_expr(arg, lazy)}" for arg in node.args[1:])
    return f"{ctx.use('xl_choose_lazy')}({index}, {choices})"


def _emit_aggregate(node: FunctionCallNode, ctx: EmitContext) -> str:
    """Emit a range-reducing call (`SUM`, `AND`, â€¦) with bound range arguments."""
    name = normalize_excel_function_name(node.name)
    func = f"xl_{name.lower()}"
    if func not in _RUNTIME_FUNCTIONS:
        raise InvertedTreeExportError(
            f"series {ctx.host.series_id!r}: Excel function {name} has no "
            "inverted-tree runtime helper"
        )
    array_ctx = replace(ctx, array_context=True)
    args = ", ".join(_emit_aggregate_arg(arg, array_ctx) for arg in node.args)
    ctx.use(func)
    return f"{func}({args})"


def _emit_aggregate_arg(node: AstNode, ctx: EmitContext) -> str:
    if isinstance(node, (RangeNode, WholeColumnNode, WholeRowNode)):
        return _emit_range_values(node, ctx)
    return emit_expr(node, ctx)


def _emit_range_values(node: AstNode, ctx: EmitContext) -> str:
    addresses = addresses_outside_blank_ranges(
        iter_ref_addresses(node, ctx.host_cell, ctx.graph),
        ctx.blank_rects,
    )
    if not addresses:
        return "()"
    if ctx.coordinate_vars is not None:
        named_view = _named_range_view(node, ctx)
        if named_view is not None:
            return named_view
        return _python_tuple([_emit_address(address, ctx) for address in addresses])
    covered = covering_series(ctx.catalog, addresses)
    if covered is not None:
        return _emit_aggregate_covering(covered, node, addresses, ctx)
    missing = [addr for addr in addresses if ctx.catalog.series_id_for(addr) is None]
    if missing:
        raise _host_export_error(ctx, f"range is not a bound series (unbound cells: {missing[:8]})")
    if ctx.graph is not None and ctx.index_var is not None:
        return _emit_mixed_aggregate_members(node, ctx)
    parts = [_emit_address(addr, ctx) for addr in addresses]
    if len(parts) == 1:
        return parts[0]
    return f"({', '.join(parts)})"


def _emit_mixed_aggregate_members(node: AstNode, ctx: EmitContext) -> str:
    """Resolve each mixed-owner window independently, preserving Excel order.

    A relative range can cross different binding owners in adjacent periods.
    Deferred rows avoid reading unused rows or unevaluated SCC members.
    """
    assert ctx.graph is not None and ctx.index_var is not None
    range_slot = _range_nodes(node_formula_ast(ctx.graph, ctx.host_cell)).index(node)
    statement = _current_statement(ctx)
    members = statement.cells if statement is not None else ctx.host.cells
    origin = statement.start if statement is not None else 0
    rows = []
    for member in members:
        ranges = _range_nodes(node_formula_ast(ctx.graph, member))
        addresses = addresses_outside_blank_ranges(
            iter_ref_addresses(ranges[range_slot], member, ctx.graph), ctx.blank_rects
        )
        values = []
        for address in addresses:
            owner = ctx.catalog.require_series_for(address)
            index = owner.index_of(address)
            assert index is not None
            if owner.series_id in ctx.scc_ids:
                values.append(f"{_emit_demanded_slots(owner, repr((index,)), ctx)}[0]")
            elif owner.is_scalar:
                values.append(ctx.param(owner.series_id))
            else:
                values.append(f"{ctx.param(owner.series_id)}[{index}]")
        rows.append(_python_tuple(values))
    if len(set(rows)) == 1:
        return rows[0]
    readers = _python_tuple([f"lambda: {row}" for row in rows])
    return f"{readers}[{_index_expr(-origin, ctx.index_var)}]()"


def _catalog_slots(
    covered: BoundSeries,
    addresses: Sequence[CanonicalAddress],
    ctx: EmitContext,
) -> tuple[int, ...]:
    """Return catalog indices of `addresses` inside `covered`."""
    indices: list[int] = []
    for addr in addresses:
        idx = covered.index_of(addr)
        if idx is None:
            raise _host_export_error(
                ctx, f"range cell {addr} is not inside bound series {covered.series_id!r}"
            )
        indices.append(idx)
    return tuple(indices)


def _member_slots_source(slots: Sequence[Sequence[int]]) -> str:
    """Return a Python tuple of per-member `take` index sequences."""
    intern = current_keyed_intern()
    table = tuple(tuple(row) for row in slots)
    if intern is not None:
        return intern.member_slots(table)
    return slot_table_to_source(table)


def _emit_demanded_slots(covered: BoundSeries, indices_expr: str, ctx: EmitContext) -> str:
    """Demand each selected catalog instance of an in-SCC producer."""
    fn = ctx.compute_names.get(covered.series_id)
    if fn is None:
        raise _host_export_error(
            ctx, f"in-SCC producer {covered.series_id!r} has no instance compute"
        )
    ctx.use("demand_instance")
    return (
        f"tuple({ctx.use('demand_instance')}("
        f"{covered.series_id!r}, j, {fn}, memo, stack) for j in {indices_expr})"
    )


def _emit_static_slots(covered: BoundSeries, slots: Sequence[int], ctx: EmitContext) -> str:
    """Emit a covering window whose catalog slots are the same for every member."""
    name = ctx.param(covered.series_id)
    if covered.series_id in ctx.scc_ids:
        return _emit_demanded_slots(covered, _catalog_indices_expr(slots), ctx)
    if tuple(slots) == tuple(range(len(covered.cells))):
        return name
    return f"{ctx.use('take')}({name}, {_catalog_indices_expr(slots)})"


def _emit_indexed_slots(
    covered: BoundSeries,
    table: Sequence[Sequence[int]],
    ctx: EmitContext,
    index_var: str,
    *,
    origin: int = 0,
) -> str:
    """Emit a covering window whose catalog slots are a function of the host index.

    Statement-local tables are subscripted as `table[i - start]` when `origin`
    is the statement's catalog start.
    """
    table_expr = wrap_slot_table_source(_member_slots_source(table))
    selector = f"{table_expr}[{_index_expr(-origin, index_var)}]"
    if covered.series_id in ctx.scc_ids:
        return _emit_demanded_slots(covered, selector, ctx)
    name = ctx.param(covered.series_id)
    return f"{ctx.use('take')}({name}, {selector})"


def _range_nodes(node: AstNode) -> list[AstNode]:
    """Return range sites in expression order without walking their endpoints."""
    if isinstance(node, (RangeNode, WholeColumnNode, WholeRowNode)):
        return [node]
    if isinstance(node, BinaryOpNode):
        return _range_nodes(node.left) + _range_nodes(node.right)
    if isinstance(node, UnaryOpNode):
        return _range_nodes(node.operand)
    if isinstance(node, FunctionCallNode):
        return [item for arg in node.args for item in _range_nodes(arg)]
    return []


def _aggregate_member_slot_table(
    covered: BoundSeries,
    node: AstNode,
    ctx: EmitContext,
) -> tuple[tuple[tuple[int, ...], ...], int] | None:
    """Return per-statement catalog slots of `node` and the catalog origin.

    Relative range endpoints are resolved against each statement member. The
    table is statement-local: generated code indexes it as `table[i - start]`.
    """
    if ctx.graph is None:
        return None
    template_ranges = _range_nodes(node_formula_ast(ctx.graph, ctx.host_cell))
    range_slot = template_ranges.index(node)
    stmt = _current_statement(ctx)
    if stmt is not None:
        members = stmt.cells
        origin = stmt.start
        length = stmt.stop - stmt.start
    else:
        members = tuple(ctx.host.cells)
        origin = 0
        length = len(ctx.host.cells)
    table: list[tuple[int, ...]] = [() for _ in range(length)]
    seen = False
    for cell in members:
        host_i = ctx.host.index_of(cell)
        if host_i is None:
            continue
        member_ranges = _range_nodes(node_formula_ast(ctx.graph, cell))
        if range_slot >= len(member_ranges):
            raise _host_export_error(ctx, f"aggregate range site is missing at {cell}")
        addresses = addresses_outside_blank_ranges(
            iter_ref_addresses(member_ranges[range_slot], cell, ctx.graph),
            ctx.blank_rects,
        )
        table[host_i - origin] = _catalog_slots(covered, addresses, ctx)
        seen = True
    if not seen:
        return None
    return tuple(table), origin


def _emit_aggregate_covering(
    covered: BoundSeries,
    node: AstNode,
    addresses: Sequence[CanonicalAddress],
    ctx: EmitContext,
) -> str:
    """Emit a literal aggregate range as catalog slots, not a block-axis access.

    Literal `SUM`/`AND` windows are member-local cell sets. Fitting them as an
    affine row/col origin fails on sparse catalogs and freezes dense take indices
    to the template member.
    """
    if covered.is_scalar:
        return ctx.param(covered.series_id)
    current = _catalog_slots(covered, addresses, ctx)
    packed = _aggregate_member_slot_table(covered, node, ctx)
    if packed is None:
        return _emit_static_slots(covered, current, ctx)
    table, origin = packed
    members = _statement_cells(ctx) or tuple(ctx.host.cells)
    unique: set[tuple[int, ...]] = set()
    for cell in members:
        host_i = ctx.host.index_of(cell)
        if host_i is not None:
            unique.add(table[host_i - origin])
    if len(unique) <= 1:
        return _emit_static_slots(covered, next(iter(unique), current), ctx)
    if ctx.index_var is None:
        raise _host_export_error(
            ctx, "aggregate range slots vary across statement members without a host index"
        )
    return _emit_indexed_slots(covered, table, ctx, ctx.index_var, origin=origin)


def _emit_lookup_arg(node: AstNode, ctx: EmitContext) -> str:
    if isinstance(node, (RangeNode, WholeColumnNode, WholeRowNode)):
        return _emit_range_table(node, ctx)
    return emit_expr(node, ctx)


def _python_tuple(items: Sequence[str]) -> str:
    """Emit a Python tuple literal; one-element rows keep a trailing comma."""
    if len(items) == 1:
        return f"({items[0]},)"
    return f"({', '.join(items)})"


def _series_slice_source(name: str, start: int, stop: int, step: int, length: int) -> str:
    """Return a slice expression covering `range(start, stop, step)` of `name`."""
    if step == 1 and start == 0 and stop == length:
        return name
    if step == 1:
        if start == 0:
            return f"{name}[:{stop}]"
        if stop == length:
            return f"{name}[{start}:]"
        return f"{name}[{start}:{stop}]"
    if start == 0 and stop >= length:
        return f"{name}[::{step}]"
    if start == 0:
        return f"{name}[:{stop}:{step}]"
    if stop >= length:
        return f"{name}[{start}::{step}]"
    return f"{name}[{start}:{stop}:{step}]"


def _positional_table_hoistable(cells: Sequence[PositionalRangeCell], ctx: EmitContext) -> bool:
    """True when every occupied cell is a bound parameter, not a live recurrence."""
    for cell in cells:
        if cell.blank:
            continue
        if cell.series_id is None:
            return False
        if cell.series_id == ctx.host.series_id or cell.series_id in ctx.scc_ids:
            return False
    return True


def _regular_column_operand(column: Sequence[PositionalRangeCell], ctx: EmitContext) -> str | None:
    """Return a zip operand for a regular catalog column, or `None`."""
    count = len(column)
    if count == 0:
        return None
    if all(cell.blank for cell in column):
        return f"(None,) * {count}"
    first = column[0]
    if first.blank or first.series_id is None:
        return None
    owner = ctx.catalog.get(first.series_id)
    name = ctx.param(owner.series_id)
    if owner.is_scalar:
        if all(not cell.blank and cell.series_id == first.series_id for cell in column):
            return f"({name},) * {count}"
        return None
    if any(cell.blank or cell.series_id != first.series_id for cell in column):
        return None
    indices: list[int] = []
    for cell in column:
        if cell.catalog_index is None:
            return None
        indices.append(cell.catalog_index)
    start = indices[0]
    step = 1 if count == 1 else indices[1] - indices[0]
    if step == 0:
        return None
    stop = start + step * count
    if tuple(indices) != tuple(range(start, stop, step)):
        return None
    return _series_slice_source(name, start, stop, step, len(owner.cells))


def _group_positional_rows(
    cells: Sequence[PositionalRangeCell],
) -> list[list[PositionalRangeCell]]:
    """Group positional cells into worksheet rows, preserving column order."""
    rows: list[list[PositionalRangeCell]] = []
    current_row: int | None = None
    current: list[PositionalRangeCell] = []
    for cell in cells:
        _sheet, row, _col = parse_cell_coords(cell.address)
        if current_row is None or row != current_row:
            if current:
                rows.append(current)
            current = [cell]
            current_row = row
        else:
            current.append(cell)
    if current:
        rows.append(current)
    return rows


def _compact_zip_table(
    rows: Sequence[Sequence[PositionalRangeCell]], ctx: EmitContext
) -> str | None:
    """Return `tuple(zip(...))` when every column is a regular slice or blank run."""
    if len(rows) < 2:
        return None
    width = len(rows[0])
    if width == 0 or any(len(row) != width for row in rows):
        return None
    flat = [cell for row in rows for cell in row]
    if not _positional_table_hoistable(flat, ctx):
        return None
    operands: list[str] = []
    for col in range(width):
        operand = _regular_column_operand([row[col] for row in rows], ctx)
        if operand is None:
            return None
        operands.append(operand)
    return f"tuple(zip({', '.join(operands)}))"


def _positional_table_source(cells: Sequence[PositionalRangeCell], ctx: EmitContext) -> str:
    """Emit a nested-tuple grid, using `zip` when rows are regular slices."""
    rows = _group_positional_rows(cells)
    if ctx.coordinate_vars is not None or (
        ctx.instance_mode and any(cell.series_id in ctx.scc_ids for cell in cells)
    ):
        table_ctx = replace(ctx, coordinate_vars={}) if ctx.coordinate_vars is not None else ctx
        callbacks = _python_tuple(
            [
                _python_tuple([f"lambda: {_emit_positional_cell(cell, table_ctx)}" for cell in row])
                for row in rows
            ]
        )
        return f"{ctx.use('lazy_table')}({callbacks})"
    compact = None if ctx.coordinate_vars is not None else _compact_zip_table(rows, ctx)
    if compact is not None:
        return compact
    return _python_tuple(
        [_python_tuple([_emit_positional_cell(cell, ctx) for cell in row]) for row in rows]
    )


def _emit_positional_cell(cell: PositionalRangeCell, ctx: EmitContext) -> str:
    """Emit one MATCH/INDEX window cell by catalog slot, not host index."""
    if cell.blank:
        return "None"
    if ctx.coordinate_vars is not None:
        return _emit_address(cell.address, ctx)
    if cell.series_id is None:
        raise _host_export_error(ctx, f"range cell {cell.address} is not a bound series")
    owner = ctx.catalog.get(cell.series_id)
    name = ctx.param(owner.series_id)
    if ctx.instance_mode and owner.series_id in ctx.scc_ids:
        index = cell.catalog_index
        if index is None:
            raise _host_export_error(ctx, f"range cell {cell.address} has no catalog index")
        return f"{ctx.use('demand_instance')}({owner.series_id!r}, {index}, {ctx.compute_names[owner.series_id]}, memo, stack)"
    if owner.is_scalar:
        return name
    if cell.catalog_index is None:
        raise _host_export_error(
            ctx, f"range cell {cell.address} is not inside bound series {owner.series_id!r}"
        )
    return f"{name}[{cell.catalog_index}]"


def _named_range_view(node: AstNode, ctx: EmitContext) -> str | None:
    """Emit a lazy `view` when a range lies inside one series' dense block.

    Each worksheet axis of the range maps onto the key field that enumerates
    that axis. A whole axis reads its keys; a partial run is a `span` between
    the corner keys, expressed relative to the host coordinate so sliding and
    growing windows share one formula family.
    """
    if not isinstance(node, RangeNode) or ctx.coordinate_vars is None or ctx.named_axes is None:
        return None
    start = as_canonical(resolve_cell_ref(node.start_ref, ctx.host_cell))
    end = as_canonical(resolve_cell_ref(node.end_ref, ctx.host_cell))
    sheet_start, row_start, col_start = parse_cell_coords(start)
    sheet_end, row_end, col_end = parse_cell_coords(end)
    if sheet_start != sheet_end:
        return None
    first_row, last_row = sorted((row_start, row_end))
    first_col, last_col = sorted((col_start, col_end))
    addresses = iter_range_addresses(start, end)
    if any(address_in_blank_ranges(address, ctx.blank_rects) for address in addresses):
        return None
    owner = covering_series(ctx.catalog, addresses)
    if owner is None or owner.is_scalar or owner.layout == "scalar" or owner.rect is None:
        return None
    try:
        field_axes = _axis_positions_are_worksheet_positions(owner)
    except InvertedTreeExportError:
        return None
    if set(field_axes) != set(owner.key_fields):
        return None
    start_index = owner.index_of(start)
    end_index = owner.index_of(end)
    if start_index is None or end_index is None:
        return None
    start_keys = _named_keys(owner, owner.domain[start_index], ctx, ref=CellRefNode(node.start_ref))
    end_keys = _named_keys(owner, owner.domain[end_index], ctx, ref=CellRefNode(node.end_ref))
    _sheet, owner_row, owner_col, _row2, _col2 = owner.rect
    domain = owner.tensor_domain
    row_expr = col_expr = None
    for position, key_field in enumerate(owner.key_fields):
        axis = field_axes[key_field]
        keys = domain.axes[position].keys
        constant = ctx.named_axes.constant(domain.axes[position])
        if axis == "row":
            first, last = first_row - owner_row, last_row - owner_row
        else:
            first, last = first_col - owner_col, last_col - owner_col
        start_key, end_key = start_keys[position], end_keys[position]
        literal_corners = start_key == repr(keys[first]) and end_key == repr(keys[last])
        if literal_corners and first == 0 and last == len(keys) - 1:
            expr = f"data.{constant}.keys"
        elif start_key == end_key:
            expr = f"({start_key},)"
        else:
            expr = f"{ctx.use('span')}(data.{constant}, {start_key}, {end_key})"
        if axis == "row":
            row_expr = expr
        else:
            col_expr = expr
    args = [ctx.param(owner.series_id)]
    if row_expr is not None:
        args.append(f"rows={row_expr}")
    if col_expr is not None:
        args.append(f"cols={col_expr}")
    if len(owner.key_fields) == 2 and field_axes[owner.key_fields[0]] == "col":
        args.append("cols_first=True")
    return f"{ctx.use('view')}({', '.join(args)})"


def _emit_range_table(node: AstNode, ctx: EmitContext) -> str:
    """Emit a nested-tuple grid, filling declared blanks with `None`."""
    named_view = _named_range_view(node, ctx)
    if named_view is not None:
        return named_view
    addresses = iter_ref_addresses(node, ctx.host_cell, ctx.graph)
    if not addresses:
        raise _host_export_error(ctx, "range is empty")
    cells, missing = resolve_positional_range(addresses, ctx.catalog, ctx.blank_rects, ctx.graph)
    if missing:
        label = range_ref_label(node, ctx.host_cell)
        raise _host_export_error(
            ctx,
            f"range {label} is not a bound series (unbound cells: {list(missing[:8])})",
        )
    expr = _positional_table_source(cells, ctx)
    reuse = current_lookup_reuse()
    if reuse is not None and _positional_table_hoistable(cells, ctx):
        return reuse.intern_table(expr)
    return expr


def _emit_covering_values(
    covered: BoundSeries,
    addresses: Sequence[CanonicalAddress],
    ctx: EmitContext,
) -> str:
    name = ctx.param(covered.series_id)
    if covered.is_scalar:
        return name
    indices: list[int] = []
    for addr in addresses:
        idx = covered.index_of(addr)
        if idx is None:
            raise _host_export_error(
                ctx, f"range cell {addr} is not inside bound series {covered.series_id!r}"
            )
        indices.append(idx)
    if indices == list(range(len(covered.cells))):
        return name
    return f"{ctx.use('take')}({name}, {_catalog_indices_expr(indices)})"


def _host_export_error(ctx: EmitContext, message: str) -> InvertedTreeExportError:
    return InvertedTreeExportError(f"series {ctx.host.series_id!r} cell {ctx.host_cell}: {message}")


def _ref_anchor_address(node: AstNode, host_cell: CanonicalAddress) -> CanonicalAddress | None:
    if isinstance(node, CellRefNode):
        return as_canonical(resolve_cell_ref(node, host_cell))
    if isinstance(node, RangeNode):
        return as_canonical(resolve_cell_ref(node.start_ref, host_cell))
    if isinstance(node, FunctionCallNode):
        name = normalize_excel_function_name(node.name)
        if name == "INDEX":
            window = index_window_corners(node, host_cell)
            if window is not None:
                return window[0]
        if name == "OFFSET":
            dest = offset_index_destination(node, host_cell)
            if dest is not None:
                return dest[0]
    return None


def _lookup_anchors(node: AstNode, host_cell: CanonicalAddress) -> list[CanonicalAddress]:
    """Return INDEX range starts and OFFSET anchors in preorder."""
    found: list[CanonicalAddress] = []

    def walk(item: AstNode, *, skip_index_array: bool = False) -> None:
        if isinstance(item, FunctionCallNode):
            name = normalize_excel_function_name(item.name)
            if name == "OFFSET" and item.args:
                base = item.args[0]
                if isinstance(base, FunctionCallNode):
                    dest = offset_index_destination(item, host_cell)
                    if dest is not None and not (
                        normalize_excel_function_name(base.name) == "INDEX"
                        and index_call_is_ref(base, host_cell)
                    ):
                        found.append(dest[0])
                    walk(base, skip_index_array=True)
                else:
                    start = _ref_anchor_address(base, host_cell)
                    if start is not None:
                        found.append(start)
                    walk(base)
                for arg in item.args[1:]:
                    walk(arg)
                return
            if (
                name == "INDEX"
                and item.args
                and isinstance(item.args[0], RangeNode)
                and not skip_index_array
            ):
                start = _ref_anchor_address(item.args[0], host_cell)
                if start is not None:
                    found.append(start)
            for arg in item.args:
                walk(arg)
            return
        if isinstance(item, BinaryOpNode):
            walk(item.left)
            walk(item.right)
            return
        if isinstance(item, UnaryOpNode):
            walk(item.operand)

    walk(node)
    return found


def _join_index_terms(terms: tuple[str, ...]) -> str:
    """Join additive index terms, dropping literal zeros."""
    kept = [term for term in terms if term not in {"0", "0.0", "(0)", "(0.0)"}]
    return " + ".join(kept) if kept else "0"


def _linear_index_expr(coeff: int, offset: int, index_var: str | None) -> str:
    """Return `coeff * index_var + offset` for a flat-block subscript."""
    if index_var is None or coeff == 0:
        return str(offset)
    if coeff == 1:
        var = index_var
    elif coeff == -1:
        if offset == 0:
            return f"-{index_var}"
        return f"{offset} - {index_var}"
    else:
        var = f"{coeff} * {index_var}"
    if offset == 0:
        return var
    if offset > 0:
        return f"{var} + {offset}"
    return f"{var} - {-offset}"


def _block_anchor_map(
    block: BoundSeries,
    ctx: EmitContext,
    slot: int,
    current_anchor: CanonicalAddress,
) -> tuple[int, int]:
    """Return `(coeff, offset)` mapping host index to the lookup's block slot.

    Fit the affine map over the current statement only. Same-shape formulas
    in other statements may read a different bound block; their anchors
    must not be required to belong to this one.
    """
    current_idx = block.index_of(current_anchor)
    if current_idx is None:
        raise _host_export_error(
            ctx,
            f"range {current_anchor} is not inside bound block {block.series_id!r}",
        )
    if ctx.graph is None or ctx.index_var is None:
        return 0, current_idx
    shape = fingerprint_formula_shape(node_formula_ast(ctx.graph, ctx.host_cell)).shape_key
    pairs: list[tuple[int, int]] = []
    members = _statement_cells(ctx) or ctx.host.cells
    for cell in members:
        index = ctx.host.index_of(cell)
        if index is None:
            continue
        ast = try_formula_ast(ctx.graph, cell)
        if ast is None:
            continue
        if fingerprint_formula_shape(ast).shape_key != shape:
            continue
        anchors = _lookup_anchors(ast, cell)
        if slot >= len(anchors):
            continue
        idx = block.index_of(anchors[slot])
        if idx is None:
            raise InvertedTreeExportError(
                f"series {ctx.host.series_id!r} cell {cell}: "
                f"range {anchors[slot]} is not inside bound block {block.series_id!r}"
            )
        pairs.append((index, idx))
    if len(pairs) < 2:
        return 0, current_idx
    fit = fit_affine_map(pairs)
    if fit is None:
        raise _host_export_error(
            ctx, "INDEX/OFFSET window is not an affine function of the host index"
        )
    return fit


def _access_or_fail(producer: BoundSeries, ctx: EmitContext) -> AccessFunction:
    if ctx.graph is None:
        raise _host_export_error(ctx, f"producer {producer.series_id!r} has no graph to classify")
    # Classify over the statement that owns the host cell, not the whole series:
    # an INDEX seed followed by a recurrence has block reads in one statement only.
    return classify_producer_access(
        ctx.host, producer, ctx.catalog, ctx.graph, cells=_statement_cells(ctx)
    )


def _emit_offset(node: FunctionCallNode, ctx: EmitContext) -> str:
    if len(node.args) < 3:
        raise _host_export_error(ctx, "OFFSET expects anchor, rows, cols")
    if ctx.coordinate_vars is not None:
        return _emit_named_offset(node, ctx)
    if isinstance(node.args[0], FunctionCallNode):
        return _emit_offset_from_expr(node, ctx)
    table = _series_for_ref(node.args[0], ctx)
    try:
        _access_or_fail(table, ctx)
    except InvertedTreeExportError as exc:
        if table.layout != "matrix":
            raise _host_export_error(
                ctx,
                f"OFFSET row offset into non-matrix series {table.series_id!r} is not supported",
            ) from exc
        raise
    rows = emit_expr(node.args[1], ctx)
    cols = emit_expr(node.args[2], ctx)
    name = ctx.param(table.series_id)
    if table.is_scalar:
        name = f"({name},)"
    anchor = _ref_anchor_address(node.args[0], ctx.host_cell)
    if anchor is None:
        raise _host_export_error(ctx, "OFFSET anchor must be a cell or range")
    slot = ctx.lookup_anchor_slot
    ctx.lookup_anchor_slot += 1
    coeff, offset = _block_anchor_map(table, ctx, slot, anchor)
    anchor_expr = _linear_index_expr(coeff, offset, ctx.index_var)
    if table.layout == "matrix":
        width = table.block_width
        index = _join_index_terms((f"({rows}) * {width}", f"({cols})", anchor_expr))
        return f"{ctx.use('xl_at')}({name}, {index})"
    if rows not in {"0", "0.0"}:
        raise _host_export_error(
            ctx,
            f"OFFSET row offset into non-matrix series {table.series_id!r} is not supported",
        )
    index = f"({cols})" if anchor_expr == "0" else f"({cols}) + ({anchor_expr})"
    return f"{ctx.use('xl_at')}({name}, {index})"


def _axis_positions_are_worksheet_positions(table: BoundSeries) -> dict[str, str]:
    """Map each key field of `table` to `"row"` or `"col"` when positions coincide.

    `OFFSET` moves along worksheet rows and columns. The move is expressed on
    a semantic axis only when the authored cells form a dense row-major block
    whose row positions enumerate one key field and whose column positions
    enumerate another, in axis order.
    """
    from excel_grapher.exporter.inverted_tree.deps import _key_field_axis

    width = table.block_width
    cells = table.cells
    height = max(1, (len(cells) + width - 1) // width)
    if height * width != len(cells):
        raise InvertedTreeExportError(
            f"series {table.series_id!r}: OFFSET requires a dense rectangular block"
        )
    axes: dict[str, str] = {}
    for key_field in table.key_fields:
        keys = tuple(dict.fromkeys(point[key_field] for point in table.domain))
        axis = _key_field_axis(table, key_field)
        if axis == "row" and len(keys) == height:
            expected = [keys[index // width] for index in range(len(cells))]
        elif axis == "col" and len(keys) == width:
            expected = [keys[index % width] for index in range(len(cells))]
        elif len(keys) == 1:
            continue
        else:
            raise InvertedTreeExportError(
                f"series {table.series_id!r}: key {key_field!r} does not enumerate a worksheet axis"
            )
        if [point[key_field] for point in table.domain] != expected:
            raise InvertedTreeExportError(
                f"series {table.series_id!r}: key {key_field!r} order differs from worksheet order"
            )
        axes[key_field] = axis
    return axes


def _emit_named_offset(node: FunctionCallNode, ctx: EmitContext) -> str:
    """Emit `OFFSET(anchor, rows, cols)` as a step along the producer's axes."""
    if isinstance(node.args[0], FunctionCallNode):
        return _emit_named_offset_index(node, ctx)
    table = _series_for_ref(node.args[0], ctx)
    rows = emit_expr(node.args[1], ctx)
    cols = emit_expr(node.args[2], ctx)
    name = ctx.param(table.series_id)
    anchor = _ref_anchor_address(node.args[0], ctx.host_cell)
    if anchor is None:
        raise _host_export_error(ctx, "OFFSET anchor must be a cell or range")
    steps = {"row": rows, "col": cols}
    static = {axis for axis, step in steps.items() if step in {"0", "0.0"}}
    if table.layout == "scalar" or table.is_scalar:
        if static == {"row", "col"}:
            return name
        # The constrained reference set is this one cell; any other
        # displacement is outside the bound model, as `xl_at` reports.
        return f"{ctx.use('at_anchor')}({name}, {rows}, {cols})"
    index = table.index_of(anchor)
    if index is None:
        raise _host_export_error(ctx, f"OFFSET anchor {anchor} is not in {table.series_id!r}")
    field_axes = _axis_positions_are_worksheet_positions(table)
    moved = {axis for axis in ("row", "col") if axis not in static}
    if moved - set(field_axes.values()):
        raise _host_export_error(
            ctx,
            f"OFFSET row offset into non-matrix series {table.series_id!r} is not supported"
            if "row" in moved
            else f"OFFSET column offset into series {table.series_id!r} is not supported",
        )
    if ctx.named_axes is None:
        raise _host_export_error(ctx, "OFFSET needs named axis constants")
    keys = _named_keys(table, table.domain[index], ctx)
    domain = table.tensor_domain
    for position, key_field in enumerate(table.key_fields):
        axis = field_axes.get(key_field)
        if axis is None or axis in static:
            continue
        constant = ctx.named_axes.constant(domain.axes[position])
        keys[position] = f"{ctx.use('axis_step')}(data.{constant}, {keys[position]}, {steps[axis]})"
    return f"{name}[{', '.join(keys)}]"


def _emit_named_offset_index(node: FunctionCallNode, ctx: EmitContext) -> str:
    """Emit `OFFSET(INDEX(range, row, col), rows, cols)` as INDEX over the moved range.

    Literal row and column displacements move the whole INDEX array; the
    INDEX selectors then pick from the destination cells. A selector that is
    `#REF!` on every host cell emits `xl_raise('#REF!')`.
    """
    base = node.args[0]
    if not (
        isinstance(base, FunctionCallNode)
        and normalize_excel_function_name(base.name) == "INDEX"
        and len(base.args) >= 2
    ):
        raise _host_export_error(ctx, "OFFSET from a computed reference must wrap INDEX")
    if _offset_index_provably_ref(base, ctx):
        return f"{ctx.use('xl_raise')}('#REF!')"
    destination = offset_index_destination(node, ctx.host_cell)
    if destination is None:
        raise _host_export_error(ctx, "OFFSET of INDEX needs literal row and column moves")
    start, end = destination
    moved = FunctionCallNode(base.name, (RangeNode(start, end), *base.args[1:]))
    return _emit_index(moved, ctx)


def _emit_offset_from_expr(node: FunctionCallNode, ctx: EmitContext) -> str:
    base = node.args[0]
    if isinstance(base, FunctionCallNode) and (normalize_excel_function_name(base.name) == "INDEX"):
        return _emit_offset_index(node, ctx)
    resolved = resolve_offset_destination_series(
        node,
        ctx.host_cell,
        ctx.catalog,
        ctx.graph,
        blank_rects=ctx.blank_rects,
    )
    if resolved is None:
        raise _host_export_error(ctx, "reference is not a bound series")
    table, _anchor = resolved
    return _emit_bound_series_lookup(table, ctx)


def _offset_index_provably_ref(index_node: FunctionCallNode, ctx: EmitContext) -> bool:
    """True when every host cell in the current statement yields INDEX `#REF!`."""
    cells = _statement_cells(ctx) or (ctx.host_cell,)
    return all(index_call_is_ref(index_node, cell) for cell in cells)


def _emit_offset_index(node: FunctionCallNode, ctx: EmitContext) -> str:
    """Emit `OFFSET(INDEX(...), rows, cols)` from the INDEX selector, not graph edges.

    Zero-displacement wrappers keep the INDEX pick even when constraint
    extraction attaches a whole-array edge set. A selector that is `#REF!` on
    every host cell emits `xl_raise('#REF!')`.
    """
    base = node.args[0]
    if not isinstance(base, FunctionCallNode):
        raise _host_export_error(ctx, "OFFSET base must be INDEX")
    if len(base.args) < 2:
        raise _host_export_error(ctx, "INDEX expects a range and row")
    if _offset_index_provably_ref(base, ctx):
        return f"{ctx.use('xl_raise')}('#REF!')"
    row_expr = emit_expr(base.args[1], ctx)
    col_arg = base.args[2] if len(base.args) > 2 else None
    col_expr, _col_literal = _emit_index_column_arg(col_arg, ctx)
    resolved = resolve_offset_destination_series(
        node,
        ctx.host_cell,
        ctx.catalog,
        ctx.graph,
        blank_rects=ctx.blank_rects,
    )
    if resolved is None:
        raise _host_export_error(ctx, "reference is not a bound series")
    table, anchor = resolved
    slot = ctx.lookup_anchor_slot
    ctx.lookup_anchor_slot += 1
    return _emit_index_into_block(
        table, anchor, row_expr, col_expr, ctx, slot, require_access=False
    )


def _row_column_args_omitted(node: FunctionCallNode) -> bool:
    return not node.args or (len(node.args) == 1 and isinstance(node.args[0], EmptyArgNode))


def _host_coord_expr(ctx: EmitContext, *, axis: str) -> str:
    """Return workbook geometry using the active evaluation order."""
    _sheet, row, col = parse_cell_coords(ctx.host_cell)
    current = row if axis == "row" else col
    if ctx.index_var is None:
        return str(current)
    cells = _statement_cells(ctx) or ctx.host.cells
    pairs: list[tuple[int, int]] = []
    for cell in cells:
        if (
            ctx.fused_partition is not None
            and schedule_partition(cell, ctx.catalog) != ctx.fused_partition
        ):
            continue
        idx = ctx.host.index_of(cell)
        if idx is None:
            continue
        if ctx.fused_plan is not None:
            idx = _union_t(ctx.fused_plan, cell, ctx)
        _cell_sheet, cell_row, cell_col = parse_cell_coords(cell)
        pairs.append((idx, cell_row if axis == "row" else cell_col))
    if len(pairs) < 2:
        return str(current)
    fit = fit_affine_map(pairs)
    if fit is None:
        return f"{dict(pairs)!r}[{ctx.index_var}]"
    return _linear_index_expr(fit[0], fit[1], ctx.index_var)


def _emit_row_or_column_ref(arg: AstNode, ctx: EmitContext, *, axis: str) -> str:
    if isinstance(arg, CellRefNode):
        address = as_canonical(resolve_cell_ref(arg, ctx.host_cell))
        _sheet, row, col = parse_cell_coords(address)
        return str(row if axis == "row" else col)
    if isinstance(arg, RangeNode):
        start = as_canonical(resolve_cell_ref(arg.start_ref, ctx.host_cell))
        _sheet, row, col = parse_cell_coords(start)
        return str(row if axis == "row" else col)
    raise _host_export_error(ctx, f"{axis.upper()} argument cannot be lowered")


def _emit_row(node: FunctionCallNode, ctx: EmitContext) -> str:
    if _row_column_args_omitted(node):
        return _host_coord_expr(ctx, axis="row")
    return _emit_row_or_column_ref(node.args[0], ctx, axis="row")


def _emit_column(node: FunctionCallNode, ctx: EmitContext) -> str:
    if _row_column_args_omitted(node):
        return _host_coord_expr(ctx, axis="col")
    return _emit_row_or_column_ref(node.args[0], ctx, axis="col")


def _emit_bound_series_lookup(covered: BoundSeries, ctx: EmitContext) -> str:
    """Emit a catalog lookup of `covered` from graph-classified access."""
    access = _access_or_fail(covered, ctx)
    name = ctx.param(covered.series_id)
    if covered.is_scalar:
        return name
    width = covered.block_width
    n_rows = max(1, (len(covered.cells) + width - 1) // width)
    row_expr = _indirect_axis_index(access.row, n_rows, ctx)
    col_expr = _indirect_axis_index(access.col, width, ctx)
    if row_expr in {"0", "0.0"}:
        row_term = "0"
    elif width <= 1:
        row_term = row_expr
    else:
        row_term = f"{row_expr} * {width}"
    index = _join_index_terms((row_term, col_expr))
    return f"{ctx.use('xl_at')}({name}, {index})"


def _indirect_axis_index(axis: AxisAccess, size: int, ctx: EmitContext) -> str:
    """Return a catalog-axis subscript, or fail closed when it is not static."""
    if axis.kind == "static":
        return _linear_index_expr(axis.coeff, axis.offset, ctx.index_var)
    if axis.kind == "whole" and size == 1:
        return "0"
    raise _host_export_error(ctx, "INDIRECT edge sets do not fit a static catalog index")


def _emit_indirect(node: FunctionCallNode, ctx: EmitContext) -> str:
    if ctx.graph is None:
        raise _host_export_error(ctx, "INDIRECT has no graph to classify")
    exclude = indirect_argument_addresses(node, ctx.host_cell)
    targets = indirect_target_addresses(ctx.graph, ctx.host_cell, exclude=tuple(exclude))
    if not targets:
        raise _host_export_error(ctx, "INDIRECT has no resolved edges")
    if ctx.coordinate_vars is not None:
        if len(targets) != 1:
            raise _host_export_error(ctx, "INDIRECT resolves to more than one cell")
        return _emit_address(targets[0], ctx)
    covered = covering_series(ctx.catalog, targets)
    if covered is None:
        raise _host_export_error(ctx, "INDIRECT targets are not one bound series")
    return _emit_bound_series_lookup(covered, ctx)


def _emit_index_into_block(
    block: BoundSeries,
    start: CanonicalAddress,
    row_expr: str,
    col_expr: str,
    ctx: EmitContext,
    slot: int,
    *,
    require_access: bool = True,
) -> str:
    if require_access:
        _access_or_fail(block, ctx)
    if block.is_scalar:
        return ctx.param(block.series_id)
    width = block.block_width
    coeff, offset = _block_anchor_map(block, ctx, slot, start)
    anchor_expr = _linear_index_expr(coeff, offset, ctx.index_var)
    col_term = "0" if col_expr in {"1", "1.0"} else f"({col_expr} - 1)"
    index = _join_index_terms((f"({row_expr} - 1) * {width}", col_term, anchor_expr))
    name = ctx.param(block.series_id)
    return f"{ctx.use('xl_at')}({name}, {index})"


def _emit_index_column_arg(col_arg: AstNode | None, ctx: EmitContext) -> tuple[str, int | None]:
    if col_arg is None or isinstance(col_arg, EmptyArgNode):
        return "None", None
    col_literal = int(col_arg.value) if isinstance(col_arg, NumberNode) else None
    try:
        return emit_expr(col_arg, ctx), col_literal
    except InvertedTreeExportError as exc:
        raise _host_export_error(ctx, f"INDEX column cannot be lowered ({exc})") from exc


def _emit_index(node: FunctionCallNode, ctx: EmitContext) -> str:
    if len(node.args) < 2:
        raise _host_export_error(ctx, "INDEX expects a range and row")
    row_arg = node.args[1]
    col_arg = node.args[2] if len(node.args) > 2 else None
    row_expr = "None" if isinstance(row_arg, EmptyArgNode) else emit_expr(row_arg, ctx)
    col_expr, col_literal = _emit_index_column_arg(col_arg, ctx)
    if row_expr in {"None", "0", "0.0"} or col_expr in {"None", "0", "0.0"}:
        # Omitted and zero selectors request vectors. Flat block indexing
        # cannot preserve their shape or Excel's single-row special case.
        table = (
            _emit_range_table(node.args[0], ctx)
            if isinstance(node.args[0], RangeNode)
            else _emit_value_or_range(node.args[0], ctx)
        )
        return _reuse_lookup(f"{ctx.use('xl_index')}({table}, {row_expr}, {col_expr})", ctx)
    if ctx.coordinate_vars is not None:
        if isinstance(node.args[0], RangeNode) and col_literal is not None and col_literal > 0:
            start = as_canonical(resolve_cell_ref(node.args[0].start_ref, ctx.host_cell))
            end = as_canonical(resolve_cell_ref(node.args[0].end_ref, ctx.host_cell))
            first_col = min(parse_cell_coords(start)[2], parse_cell_coords(end)[2])
            selected = [
                address
                for address in iter_ref_addresses(node.args[0], ctx.host_cell, ctx.graph)
                if parse_cell_coords(address)[2] == first_col + col_literal - 1
            ]
            if selected:
                # A proven column selection must not introduce dependencies on
                # other columns excluded by the extracted graph.
                column_view = _named_range_view(RangeNode(selected[0], selected[-1]), ctx)
                if column_view is not None:
                    return f"{ctx.use('xl_index')}({column_view}, {row_expr}, 1)"
                cells, missing = resolve_positional_range(
                    selected, ctx.catalog, ctx.blank_rects, ctx.graph
                )
                if missing:
                    raise _host_export_error(
                        ctx, f"INDEX selected column has unbound cells: {list(missing[:8])}"
                    )
                table = (
                    "("
                    + ", ".join(f"({_emit_positional_cell(cell, ctx)},)" for cell in cells)
                    + ",)"
                )
                return f"{ctx.use('xl_index')}({table}, {row_expr}, 1)"
        table = _emit_value_or_range(node.args[0], ctx)
        return f"{ctx.use('xl_index')}({table}, {row_expr}, {col_expr})"
    if isinstance(node.args[0], RangeNode):
        start = as_canonical(resolve_cell_ref(node.args[0].start_ref, ctx.host_cell))
        end = as_canonical(resolve_cell_ref(node.args[0].end_ref, ctx.host_cell))
        covered_full = covering_series_of_range(ctx.catalog, start, end)
        if covered_full is not None:
            slot = ctx.lookup_anchor_slot
            ctx.lookup_anchor_slot += 1
            return _emit_index_into_block(covered_full, start, row_expr, col_expr, ctx, slot)
        covered = None
        if col_literal is not None:
            try:
                covered = covering_series_of_column(ctx.catalog, start, end, col_literal)
            except InvertedTreeExportError:
                covered = None
        if covered is None:
            table = _emit_range_table(node.args[0], ctx)
            return _reuse_lookup(f"{ctx.use('xl_index')}({table}, {row_expr}, {col_expr})", ctx)
        if covered.layout == "matrix" or covered.block_width > 1:
            # The range overhangs the bound block (Q-CRAFT: a 28-column window
            # over a 22-column block) but the accessed column is inside it.
            # Index the block by row stride and column, never the flat row.
            slot = ctx.lookup_anchor_slot
            ctx.lookup_anchor_slot += 1
            return _emit_index_into_block(covered, start, row_expr, col_expr, ctx, slot)
        name = ctx.param(covered.series_id)
        return f"{ctx.use('xl_at')}({name}, ({row_expr}) - 1)"
    table = _emit_value_or_range(node.args[0], ctx)
    if col_arg is None or isinstance(col_arg, EmptyArgNode):
        return _reuse_lookup(f"{ctx.use('xl_index')}({table}, {row_expr})", ctx)
    return _reuse_lookup(f"{ctx.use('xl_index')}({table}, {row_expr}, {col_expr})", ctx)


def _emit_match_array(node: AstNode, ctx: EmitContext) -> str:
    """Emit a MATCH lookup vector that preserves Excel positions."""
    if ctx.coordinate_vars is not None:
        return _emit_value_or_range(node, ctx)
    if isinstance(node, CellRefNode):
        series = _series_for_ref(node, ctx)
        name = ctx.param(series.series_id)
        return name if not series.is_scalar else f"({name},)"
    if isinstance(node, (RangeNode, WholeColumnNode, WholeRowNode)):
        addresses = iter_ref_addresses(node, ctx.host_cell, ctx.graph)
        cells, missing = resolve_positional_range(
            addresses, ctx.catalog, ctx.blank_rects, ctx.graph
        )
        if missing:
            label = range_ref_label(node, ctx.host_cell)
            raise _host_export_error(
                ctx,
                f"range {label} is not a bound series (unbound cells: {list(missing[:8])})",
            )
        owned = [cell.address for cell in cells if not cell.blank]
        if owned and not any(cell.blank for cell in cells):
            covered = covering_series(ctx.catalog, owned)
            if covered is not None:
                values = _emit_covering_values(covered, owned, ctx)
                return values if not covered.is_scalar else f"({values},)"
        return _emit_range_table(node, ctx)
    return emit_expr(node, ctx)


def _emit_match(node: FunctionCallNode, ctx: EmitContext) -> str:
    if len(node.args) < 2:
        raise InvertedTreeExportError(
            f"series {ctx.host.series_id!r}: MATCH expects lookup and array"
        )
    lookup = emit_expr(node.args[0], ctx)
    array = _emit_match_array(node.args[1], ctx)
    match_type = emit_expr(node.args[2], ctx) if len(node.args) > 2 else "0"
    return _reuse_lookup(f"{ctx.use('xl_match')}({lookup}, {array}, {match_type})", ctx)


def _series_for_ref(node: AstNode, ctx: EmitContext) -> BoundSeries:
    if isinstance(node, CellRefNode):
        address = as_canonical(resolve_cell_ref(node, ctx.host_cell))
        if address_in_blank_ranges(address, ctx.blank_rects):
            raise _host_export_error(ctx, "reference is not a bound series")
        return ctx.catalog.require_series_for(address)
    if isinstance(node, (RangeNode, WholeColumnNode, WholeRowNode)):
        addresses = addresses_outside_blank_ranges(
            iter_ref_addresses(node, ctx.host_cell, ctx.graph),
            ctx.blank_rects,
        )
        covered = covering_series(ctx.catalog, addresses) if addresses else None
        if covered is None:
            raise _host_export_error(ctx, "reference is not a bound series")
        return covered
    if isinstance(node, FunctionCallNode):
        name = normalize_excel_function_name(node.name)
        if name == "INDEX":
            covered = covering_series_for_index_window(
                node, ctx.host_cell, ctx.catalog, blank_rects=ctx.blank_rects
            )
            if covered is not None:
                return covered
            raise _host_export_error(ctx, "reference is not a bound series")
        if name == "OFFSET":
            resolved = resolve_offset_destination_series(
                node,
                ctx.host_cell,
                ctx.catalog,
                ctx.graph,
                blank_rects=ctx.blank_rects,
            )
            if resolved is not None:
                return resolved[0]
            raise _host_export_error(ctx, "reference is not a bound series")
    raise _host_export_error(ctx, "OFFSET/MATCH base must be a cell or range")


def _as_measure_call(expr: str, series: BoundSeries) -> str:
    if series.python_dtype == "float":
        return f"as_measure({expr})"
    return f"as_measure({expr}, {series.python_dtype!r})"
