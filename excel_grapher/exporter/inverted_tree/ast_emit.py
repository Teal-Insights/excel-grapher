"""Translate a bound series' Excel AST into a first-level-dep Python helper."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from excel_grapher.core.address_keys import CanonicalAddress, as_canonical, parse_cell_coords
from excel_grapher.core.excel_function_names import normalize_excel_function_name
from excel_grapher.core.formula_ast import (
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
    classify_cell_ref_access,
    classify_producer_access,
    indirect_argument_addresses,
    indirect_target_addresses,
)
from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    SeriesCatalog,
    covering_series,
    covering_series_of_column,
    covering_series_of_range,
    fit_affine_map,
    preferred_fields,
    schedule_axis_coord,
    schedule_partition,
)
from excel_grapher.exporter.inverted_tree.deps import (
    DependenceEdge,
    SeriesDeps,
    addresses_outside_blank_ranges,
    current_blank_rects,
    iter_ref_addresses,
    node_formula_ast,
    predecessor_address,
    successor_address,
    try_formula_ast,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.schedule import (
    FusedPlan,
    FusedRegion,
    collect_dependence_edges,
    indices_to_source,
    plan_fused_scc,
)
from excel_grapher.grapher.blank_ranges import BlankRangeRect, address_in_blank_ranges
from excel_grapher.series_bindings.types import Scalar

if TYPE_CHECKING:
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
_AGGREGATE_FUNCTIONS = frozenset({"SUM", "SUMPRODUCT"})
_LOOKUP_TABLE_FUNCTIONS = frozenset({"VLOOKUP", "HLOOKUP", "LOOKUP", "XLOOKUP"})


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

    def param(self, series_id: str) -> str:
        return series_id

    def use(self, symbol: str) -> str:
        self.used_runtime.add(symbol)
        return symbol


def python_measure_type(series: BoundSeries) -> str:
    """Return the Python type of one observation (`float | str` for numbers)."""
    if series.python_dtype in {"float", "int"}:
        base = f"{series.python_dtype} | str"
    else:
        base = series.python_dtype
    if series.has_none_holes:
        return f"{base} | None"
    return base


def _python_param_inner(series: BoundSeries) -> str:
    """Inner type of a helper parameter for `series`.

    Numeric leaves use the measure type (`float | str` / `int | str`) so
    cached text and error-code strings in `data.py` stay assignable.
    """
    return python_measure_type(series)


def python_annotation(series: BoundSeries) -> str:
    """Return a typing annotation for a helper parameter."""
    inner = _python_param_inner(series)
    if series.is_scalar:
        return inner
    return f"Sequence[{inner}]"


def python_data_annotation(series: BoundSeries) -> str:
    """Return a typing annotation for a `data.py` constant or default.

    Uses the same inner type as `python_annotation` so workbook defaults are
    assignable to `compute_*` parameters. Sequence leaves are stored as tuples.
    """
    inner = _python_param_inner(series)
    if series.is_scalar:
        return inner
    return f"tuple[{inner}, ...]"


def python_return_annotation(series: BoundSeries) -> str:
    """Return a typing annotation for a helper result."""
    inner = python_measure_type(series)
    if series.is_scalar:
        return inner
    return f"tuple[{inner}, ...]"


def _cast_scalar(expr: str, dtype: str) -> str:
    if dtype == "float":
        return f"float({expr})"
    if dtype == "int":
        return f"int({expr})"
    if dtype == "str":
        return f"str({expr})"
    if dtype == "bool":
        return f"bool({expr})"
    return expr


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
    """True when `address` is the scan accumulator, not a shared selector.

    Index 0 (or the last member of a reversed scan) may sit next to a bound
    scalar. That neighbor is `prior` only when `SeriesDeps` classified it as
    the seed. An absolute selector read by every member stays a parameter.
    """
    pred = predecessor_address(ctx.host, ctx.host_index, ctx.catalog, ctx.graph)
    if pred is not None and address == pred:
        if ctx.host_index > 0:
            return True
        owner = ctx.catalog.series_for(address)
        return owner is not None and owner.series_id == ctx.deps.seed_id
    succ = successor_address(ctx.host, ctx.host_index, ctx.catalog, ctx.graph)
    if succ is not None and address == succ:
        if ctx.host_index < len(ctx.host.cells) - 1:
            return True
        owner = ctx.catalog.series_for(address)
        return owner is not None and owner.series_id == ctx.deps.seed_id
    return False


def _emit_cell_ref(node: CellRefNode, ctx: EmitContext) -> str:
    return _emit_address(as_canonical(resolve_cell_ref(node, ctx.host_cell)), ctx, ref=node)


def _statement_cells(ctx: EmitContext) -> tuple[CanonicalAddress, ...] | None:
    return next(
        (stmt.cells for stmt in ctx.host.statements if ctx.host_cell in stmt.cells),
        None,
    )


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


def _emit_address(
    address: CanonicalAddress, ctx: EmitContext, *, ref: CellRefNode | None = None
) -> str:
    if address_in_blank_ranges(address, ctx.blank_rects):
        return "None"
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
        return f"{name}[{_index_expr(idx - ctx.host_index, ctx.index_var)}]"
    if idx is not None and ctx.index_var is None:
        return f"{name}[{idx}]" if not owner.is_scalar else name
    if owner.series_id in ctx.deps.lookup_ids:
        return name
    return name


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

    A producer keyed by the outer partition fields — with or without
    `TIME_PERIOD` — has one block per area. Uniform block length is the
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
    host_union = _union_t(plan, ctx.host_cell, ctx)
    step = -1 if plan.direction == "reversed" else 1
    if ctx.fused_use_area:
        aligned = _aligned_area_index(owner, address, ctx, plan)
        if aligned is not None:
            area_i, stride = aligned
            host_part = schedule_partition(ctx.host_cell, ctx.catalog)
            host_i = plan.partitions.index(host_part) if host_part in plan.partitions else 0
            fields = preferred_fields(owner, ctx.catalog) or ()
            if "TIME_PERIOD" in fields:
                local = idx - step * host_union - area_i * stride
                inner = _affine_index_expr(local, index_var, step=step)
            else:
                local = idx - area_i * stride
                inner = str(local)
            index_expr = _combine_area_index(_area_stride_expr(stride, area_i - host_i), inner)
            return f"live_measure({name}[{index_expr}])"
    index_expr = _affine_index_expr(idx - step * host_union, index_var, step=step)
    return f"live_measure({name}[{index_expr}])"


def _emit_instance_ref(
    address: CanonicalAddress, ctx: EmitContext, *, ref: CellRefNode | None = None
) -> str:
    owner = ctx.catalog.require_series_for(address)
    literal = _static_catalog_literal(owner, address, ctx, ref)
    if literal is not None and owner.series_id not in ctx.scc_ids:
        return f"{ctx.param(owner.series_id)}[{literal}]"
    idx = owner.index_of(address)
    if idx is None:
        raise InvertedTreeExportError(
            f"series {ctx.host.series_id!r}: instance ref {address} is unbound"
        )
    if (
        owner.series_id not in ctx.scc_ids
        and owner.series_id in ctx.deps.aligned_ids
        and not owner.is_scalar
    ):
        idx = _aligned_taken_index(owner.series_id, idx, ctx)
    offset = idx - ctx.host_index
    index_var = ctx.index_var or "i"
    index_expr = _index_expr(offset, index_var)
    if owner.series_id in ctx.scc_ids:
        fn = ctx.compute_names[owner.series_id]
        ctx.use("demand_instance")
        return f"demand_instance({owner.series_id!r}, {index_expr}, {fn}, memo, stack)"
    name = ctx.param(owner.series_id)
    if owner.is_scalar:
        return name
    return f"{name}[{index_expr}]"


def formula_shape_runs(
    series: BoundSeries,
    graph: DependencyGraph,
) -> list[tuple[str, int, int]]:
    """Return consecutive `(shape_key, start, stop)` runs over `series.cells`."""
    statements = series.statements
    if len(statements) > 1 or (statements and statements[0].shape_key is not None):
        return [(stmt.shape_key or "", stmt.start, stmt.stop) for stmt in statements]
    runs: list[tuple[str, int, int]] = []
    for index, address in enumerate(series.cells):
        ast = try_formula_ast(graph, address)
        key = fingerprint_formula_shape(ast).shape_key if ast is not None else ""
        if runs and runs[-1][0] == key:
            runs[-1] = (key, runs[-1][1], index + 1)
        else:
            runs.append((key, index, index + 1))
    return runs


def member_peel_stop(series: BoundSeries, graph: DependencyGraph) -> int:
    """Return the exclusive end of a peeled prefix, or 0 if the series is uniform."""
    runs = formula_shape_runs(series, graph)
    if len(runs) <= 1:
        return 0
    if len(runs) == 2 and runs[0][1] == 0:
        return runs[0][2]
    raise InvertedTreeExportError(
        f"series {series.series_id!r} has {len({run[0] for run in runs})} formula "
        "shapes; members must share one shape or a peeled prefix"
    )


def _emit_binary(node: BinaryOpNode, ctx: EmitContext) -> str:
    left = emit_expr(node.left, ctx)
    right = emit_expr(node.right, ctx)
    op = node.op
    if op == "&":
        return f"(str({left}) + str({right}))"
    helper = _ARITHMETIC_HELPERS.get(op) or _COMPARE_HELPERS.get(op)
    if helper is not None:
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
    if name == "IF":
        return _emit_if(node, ctx)
    if name == "CHOOSE":
        return _emit_choose(node, ctx)
    if name == "OFFSET":
        return _emit_offset(node, ctx)
    if name == "INDEX":
        return _emit_index(node, ctx)
    if name == "INDIRECT":
        return _emit_indirect(node, ctx)
    if name == "MATCH":
        return _emit_match(node, ctx)
    if name == "TRUE":
        return "True"
    if name == "FALSE":
        return "False"
    if name in _AGGREGATE_FUNCTIONS:
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
    return f"{func}({args})"


def _emit_if(node: FunctionCallNode, ctx: EmitContext) -> str:
    if len(node.args) < 2:
        return f"{ctx.use('xl_raise')}('#VALUE!')"
    cond = emit_expr(node.args[0], ctx)
    then = emit_expr(node.args[1], ctx)
    otherwise = emit_expr(node.args[2], ctx) if len(node.args) > 2 else "0"
    return f"({then} if {cond} else {otherwise})"


def _emit_choose(node: FunctionCallNode, ctx: EmitContext) -> str:
    if len(node.args) < 2:
        return f"{ctx.use('xl_raise')}('#VALUE!')"
    index = emit_expr(node.args[0], ctx)
    choices = ", ".join(emit_expr(arg, ctx) for arg in node.args[1:])
    return f"{ctx.use('xl_choose')}({index}, {choices})"


def _emit_aggregate(node: FunctionCallNode, ctx: EmitContext) -> str:
    name = normalize_excel_function_name(node.name)
    func = f"xl_{name.lower()}"
    if func not in _RUNTIME_FUNCTIONS:
        raise InvertedTreeExportError(
            f"series {ctx.host.series_id!r}: Excel function {name} has no "
            "inverted-tree runtime helper"
        )
    args = ", ".join(_emit_aggregate_arg(arg, ctx) for arg in node.args)
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
    covered = covering_series(ctx.catalog, addresses)
    if covered is not None:
        _access_or_fail(covered, ctx)
        return _emit_covering_values(covered, addresses, ctx)
    missing = [addr for addr in addresses if ctx.catalog.series_id_for(addr) is None]
    if missing:
        raise _host_export_error(ctx, f"range is not a bound series (unbound cells: {missing[:8]})")
    parts = [_emit_address(addr, ctx) for addr in addresses]
    if len(parts) == 1:
        return parts[0]
    return f"({', '.join(parts)})"


def _emit_lookup_arg(node: AstNode, ctx: EmitContext) -> str:
    if isinstance(node, (RangeNode, WholeColumnNode, WholeRowNode)):
        return _emit_range_table(node, ctx)
    return emit_expr(node, ctx)


def _emit_range_table(node: AstNode, ctx: EmitContext) -> str:
    """Emit a nested-tuple grid, filling declared blanks with `None`."""
    addresses = iter_ref_addresses(node, ctx.host_cell, ctx.graph)
    if not addresses:
        raise _host_export_error(ctx, "range is empty")
    missing = [
        addr
        for addr in addresses
        if ctx.catalog.series_id_for(addr) is None
        and not address_in_blank_ranges(addr, ctx.blank_rects)
    ]
    if missing:
        raise _host_export_error(ctx, f"range is not a bound series (unbound cells: {missing[:8]})")
    rows: list[list[str]] = []
    current_row: int | None = None
    current: list[str] = []
    for addr in addresses:
        _sheet, row, _col = parse_cell_coords(addr)
        cell = (
            "None" if address_in_blank_ranges(addr, ctx.blank_rects) else _emit_address(addr, ctx)
        )
        if current_row is None or row != current_row:
            if current:
                rows.append(current)
            current = [cell]
            current_row = row
        else:
            current.append(cell)
    if current:
        rows.append(current)
    row_srcs = [f"({', '.join(row)})" for row in rows]
    if len(row_srcs) == 1:
        return f"({row_srcs[0]},)"
    return f"({', '.join(row_srcs)})"


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
    return f"{ctx.use('take')}({name}, {indices_to_source(indices)})"


def _host_export_error(ctx: EmitContext, message: str) -> InvertedTreeExportError:
    return InvertedTreeExportError(f"series {ctx.host.series_id!r} cell {ctx.host_cell}: {message}")


def _ref_anchor_address(node: AstNode, host_cell: CanonicalAddress) -> CanonicalAddress | None:
    if isinstance(node, CellRefNode):
        return as_canonical(resolve_cell_ref(node, host_cell))
    if isinstance(node, RangeNode):
        return as_canonical(resolve_cell_ref(node.start_ref, host_cell))
    return None


def _lookup_anchors(node: AstNode, host_cell: CanonicalAddress) -> list[CanonicalAddress]:
    """Return INDEX range starts and OFFSET anchors in preorder."""
    found: list[CanonicalAddress] = []

    def walk(item: AstNode) -> None:
        if isinstance(item, FunctionCallNode):
            name = normalize_excel_function_name(item.name)
            if item.args and (
                name == "OFFSET" or (name == "INDEX" and isinstance(item.args[0], RangeNode))
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
    """Return `(coeff, offset)` mapping host index to the lookup's block slot."""
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
    for index, cell in enumerate(ctx.host.cells):
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
    cells = next(
        (stmt.cells for stmt in ctx.host.statements if ctx.host_cell in stmt.cells),
        None,
    )
    return classify_producer_access(ctx.host, producer, ctx.catalog, ctx.graph, cells=cells)


def _emit_offset(node: FunctionCallNode, ctx: EmitContext) -> str:
    if len(node.args) < 3:
        raise _host_export_error(ctx, "OFFSET expects anchor, rows, cols")
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
    covered = covering_series(ctx.catalog, targets)
    if covered is None:
        raise _host_export_error(ctx, "INDIRECT targets are not one bound series")
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


def _emit_index_into_block(
    block: BoundSeries,
    start: CanonicalAddress,
    row_expr: str,
    col_expr: str,
    ctx: EmitContext,
    slot: int,
) -> str:
    _access_or_fail(block, ctx)
    width = block.block_width
    coeff, offset = _block_anchor_map(block, ctx, slot, start)
    anchor_expr = _linear_index_expr(coeff, offset, ctx.index_var)
    col_term = "0" if col_expr in {"1", "1.0"} else f"({col_expr} - 1)"
    index = _join_index_terms((f"({row_expr} - 1) * {width}", col_term, anchor_expr))
    name = ctx.param(block.series_id)
    return f"{ctx.use('xl_at')}({name}, {index})"


def _emit_index_column_arg(col_arg: AstNode | None, ctx: EmitContext) -> tuple[str, int | None]:
    if col_arg is None or isinstance(col_arg, EmptyArgNode):
        return "1", 1
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
    row_expr = emit_expr(row_arg, ctx)
    col_expr, col_literal = _emit_index_column_arg(col_arg, ctx)
    if isinstance(node.args[0], RangeNode):
        start = as_canonical(resolve_cell_ref(node.args[0].start_ref, ctx.host_cell))
        end = as_canonical(resolve_cell_ref(node.args[0].end_ref, ctx.host_cell))
        covered_full = covering_series_of_range(ctx.catalog, start, end)
        if covered_full is not None:
            slot = ctx.lookup_anchor_slot
            ctx.lookup_anchor_slot += 1
            return _emit_index_into_block(covered_full, start, row_expr, col_expr, ctx, slot)
        if col_literal is None:
            raise _host_export_error(
                ctx, "INDEX column is not a literal and the range is not one bound block"
            )
        covered = covering_series_of_column(ctx.catalog, start, end, col_literal)
        if covered is None:
            raise _host_export_error(ctx, "INDEX column is not a bound series")
        if covered.layout == "matrix" or covered.block_width > 1:
            # The range overhangs the bound block (Q-CRAFT: a 28-column window
            # over a 22-column block) but the accessed column is inside it.
            # Index the block by row stride and column, never the flat row.
            slot = ctx.lookup_anchor_slot
            ctx.lookup_anchor_slot += 1
            return _emit_index_into_block(covered, start, row_expr, col_expr, ctx, slot)
        name = ctx.param(covered.series_id)
        return f"{ctx.use('xl_at')}({name}, ({row_expr}) - 1)"
    table = emit_expr(node.args[0], ctx)
    return f"{ctx.use('xl_at')}({table}, ({row_expr}) - 1)"


def _emit_match(node: FunctionCallNode, ctx: EmitContext) -> str:
    if len(node.args) < 2:
        raise InvertedTreeExportError(
            f"series {ctx.host.series_id!r}: MATCH expects lookup and array"
        )
    lookup = emit_expr(node.args[0], ctx)
    array_series = _series_for_ref(node.args[1], ctx)
    array = ctx.param(array_series.series_id)
    match_type = emit_expr(node.args[2], ctx) if len(node.args) > 2 else "0"
    return f"{ctx.use('xl_match')}({lookup}, {array}, {match_type})"


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
    raise _host_export_error(ctx, "OFFSET/MATCH base must be a cell or range")


def _emit_hole_expr(series: BoundSeries, host_index: int, graph: DependencyGraph) -> str:
    """Return a Python literal for a retained matrix hole cell."""
    from excel_grapher.exporter.inverted_tree.emit import (
        _cell_value,
        _coerce_cached_value,
        _py_literal,
    )

    hole = series.hole_at(host_index)
    address = series.cells[host_index]
    if hole is None or hole.kind in {"blank", "off_closure"}:
        return "None"
    if hole.kind == "graph_leaf":
        node = graph.get_node(address)
        if node is None or node.value is None:
            raise InvertedTreeExportError(
                f"series {series.series_id!r} cell {address}: graph leaf has no cached value"
            )
        return _py_literal(_cell_value(graph, address, series.dtype))
    if hole.literal is None:
        raise InvertedTreeExportError(
            f"series {series.series_id!r} cell {address}: cached value is unavailable"
        )
    return _py_literal(_coerce_cached_value(hole.literal, series.dtype, address))


def _region_measure(
    series: BoundSeries,
    *,
    catalog: SeriesCatalog,
    deps: SeriesDeps,
    graph: DependencyGraph,
    host_index: int,
    index_var: str | None,
    prior_var: str | None,
) -> tuple[str, set[str]]:
    if try_formula_ast(graph, series.cells[host_index]) is None:
        return _emit_hole_expr(series, host_index, graph), set()
    ctx = EmitContext(
        host=series,
        catalog=catalog,
        deps=deps,
        host_index=host_index,
        host_cell=series.cells[host_index],
        index_var=index_var,
        prior_var=prior_var,
        graph=graph,
    )
    expr = emit_expr(node_formula_ast(graph, series.cells[host_index]), ctx)
    return _as_measure_call(expr, series), set(ctx.used_runtime)


def _emit_region_chain(
    series: BoundSeries,
    *,
    catalog: SeriesCatalog,
    deps: SeriesDeps,
    graph: DependencyGraph,
    runs: list[tuple[str, int, int]],
    prior_var: str | None,
    prefix: str,
    suffix: str,
    indent: str,
) -> tuple[list[str], set[str]]:
    """Emit an if/elif chain that assigns each shape-run formula."""
    used: set[str] = set()
    lines: list[str] = []
    for run_index, (_key, start, stop) in enumerate(runs):
        coerce, expr_used = _region_measure(
            series,
            catalog=catalog,
            deps=deps,
            graph=graph,
            host_index=start,
            index_var="i",
            prior_var=prior_var,
        )
        used |= expr_used
        statement = f"{prefix}{coerce}{suffix}"
        if run_index < len(runs) - 1:
            keyword = "if" if run_index == 0 else "elif"
            lines.append(f"{indent}{keyword} i < {stop}:")
            lines.append(f"{indent}    {statement}")
        elif run_index == 0:
            lines.append(f"{indent}{statement}")
        else:
            lines.append(f"{indent}else:")
            lines.append(f"{indent}    {statement}")
    return lines, used


def _is_identity_aligned(deps: SeriesDeps, param_id: str, host_n: int) -> bool:
    """True when `param_id` is identity-aligned to a `host_n`-long walk."""
    if param_id not in deps.aligned_ids:
        return False
    if param_id in deps.affine_maps:
        return False
    index_map = deps.index_maps.get(param_id)
    return index_map is None or index_map == tuple(range(host_n))


def emit_sequence_length_guards(
    series: BoundSeries,
    deps: SeriesDeps,
    catalog: SeriesCatalog,
    *,
    emit_n: bool = True,
) -> tuple[list[str], set[str]]:
    """Emit length guards so helpers reject producer arrays of the wrong size.

    Identity-aligned params use `require_aligned` so scan helpers still accept
    a shorter working buffer. Affine and index-mapped params are taken to the
    host walk at the call site, so the helper requires host length. Other
    sequence params (lags, lookups, mixed reads) must be dense over the
    producer's `__domain__`.
    """
    host_n = len(series.cells)
    used: set[str] = set()
    lines: list[str] = []
    identity: list[str] = []
    for param_id in deps.param_ids:
        producer = catalog.get(param_id)
        if not producer.is_sequence:
            continue
        if param_id in deps.aligned_ids and _is_identity_aligned(deps, param_id, host_n):
            identity.append(param_id)
            continue
        if param_id in deps.aligned_ids:
            used.add("require_length")
            lines.append(f"    require_length({param_id}, {host_n})")
            continue
        used.add("require_length")
        lines.append(f"    require_length({param_id}, {len(producer.cells)})")
    if identity:
        used.add("require_aligned")
        if emit_n:
            lines.append(f"    n = require_aligned({', '.join(identity)})")
        else:
            used.add("require_length")
            for param_id in identity:
                lines.append(f"    require_length({param_id}, {host_n})")
    elif emit_n:
        lines.append(f"    n = {host_n}")
    return lines, used


def emit_helper_body(
    series: BoundSeries,
    *,
    catalog: SeriesCatalog,
    deps: SeriesDeps,
    graph: DependencyGraph,
) -> tuple[list[str], set[str]]:
    """Return indented body lines and the runtime symbols they use."""
    used: set[str] = set()
    if series.is_scalar:
        guard_lines, guard_used = emit_sequence_length_guards(series, deps, catalog, emit_n=False)
        used |= guard_used
        ctx = EmitContext(
            host=series,
            catalog=catalog,
            deps=deps,
            host_index=0,
            host_cell=series.cells[0],
            index_var=None,
            prior_var=None,
            graph=graph,
        )
        ast = node_formula_ast(graph, series.cells[0])
        expr = emit_expr(ast, ctx)
        used |= ctx.used_runtime
        used.add("as_measure")
        used.add("XlError")
        if series.python_dtype in {"float", "int"}:
            coerce = (
                f"as_measure({expr})"
                if series.python_dtype == "float"
                else f"as_measure({expr}, {series.python_dtype!r})"
            )
            return [
                "    try:",
                f"        return {coerce}",
                "    except XlError as err:",
                "        return err.code",
            ], used
        return [f"    return {_cast_scalar(expr, series.python_dtype)}"], used

    if deps.is_scan:
        return _emit_scan_body(series, catalog=catalog, deps=deps, graph=graph)

    runs = formula_shape_runs(series, graph)
    used.add("as_measure")
    used.add("XlError")
    guard_lines, guard_used = emit_sequence_length_guards(series, deps, catalog)
    used |= guard_used
    lines: list[str] = []
    lines.extend(guard_lines)
    measure = python_measure_type(series)
    lines.append(f"    out: list[{measure}] = []")
    lines.append("    for i in range(n):")
    lines.append("        try:")
    if len(runs) <= 1:
        coerce, expr_used = _region_measure(
            series,
            catalog=catalog,
            deps=deps,
            graph=graph,
            host_index=0,
            index_var="i",
            prior_var=None,
        )
        used |= expr_used
        lines.append(f"            out.append({coerce})")
    else:
        region_lines, region_used = _emit_region_chain(
            series,
            catalog=catalog,
            deps=deps,
            graph=graph,
            runs=runs,
            prior_var=None,
            prefix="out.append(",
            suffix=")",
            indent="            ",
        )
        used |= region_used
        lines.extend(region_lines)
    lines.append("        except XlError as err:")
    lines.append("            out.append(err.code)")
    lines.append("    return tuple(out)")
    return lines, used


def _emit_scan_body(
    series: BoundSeries,
    *,
    catalog: SeriesCatalog,
    deps: SeriesDeps,
    graph: DependencyGraph,
) -> tuple[list[str], set[str]]:
    runs = formula_shape_runs(series, graph)
    seed = deps.seed_id
    used: set[str] = {"as_measure", "XlError", "is_error"}
    guard_lines, guard_used = emit_sequence_length_guards(series, deps, catalog)
    used |= guard_used
    lines: list[str] = list(guard_lines)
    seed_expr = seed if seed is not None else "0"
    measure = python_measure_type(series)
    lines.append(f"    path: list[{measure}] = []")
    lines.append(f"    prior: {measure} = {seed_expr}")
    if deps.scan_direction == "reversed":
        lines.append("    for i in reversed(range(n)):")
    else:
        lines.append("    for i in range(n):")
    lines.append("        if is_error(prior):")
    lines.append("            path.append(prior)")
    lines.append("            continue")
    lines.append("        try:")
    region_lines, region_used = _emit_region_chain(
        series,
        catalog=catalog,
        deps=deps,
        graph=graph,
        runs=runs or [("", 0, len(series.cells))],
        prior_var="prior",
        prefix="prior = ",
        suffix="",
        indent="            ",
    )
    used |= region_used
    lines.extend(region_lines)
    lines.append("        except XlError as err:")
    lines.append("            prior = err.code")
    lines.append("        path.append(prior)")
    if deps.scan_direction == "reversed":
        lines.append("    return tuple(reversed(path))")
    else:
        lines.append("    return tuple(path)")
    return lines, used


def _as_measure_call(expr: str, series: BoundSeries) -> str:
    if series.python_dtype == "float":
        return f"as_measure({expr})"
    return f"as_measure({expr}, {series.python_dtype!r})"


def _compute_fn_name(series_id: str) -> str:
    return f"{series_id}_compute"


def _emit_region_return(
    series: BoundSeries,
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
    graph: DependencyGraph,
    host_index: int,
    scc_ids: frozenset[str],
    compute_names: dict[str, str],
) -> tuple[str, set[str]]:
    ctx = EmitContext(
        host=series,
        catalog=catalog,
        deps=deps[series.series_id],
        host_index=host_index,
        host_cell=series.cells[host_index],
        index_var="i",
        prior_var=None,
        scc_ids=scc_ids,
        instance_mode=True,
        compute_names=compute_names,
        graph=graph,
    )
    if try_formula_ast(graph, series.cells[host_index]) is None:
        return _emit_hole_expr(series, host_index, graph), set()
    expr = emit_expr(node_formula_ast(graph, series.cells[host_index]), ctx)
    return _as_measure_call(expr, series), set(ctx.used_runtime)


def emit_rung3_scc(
    scc: tuple[str, ...],
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
    graph: DependencyGraph,
    edges: Sequence[DependenceEdge] | None = None,
) -> tuple[list[str], set[str]]:
    """Emit demand-driven instance evaluation for an SCC (the rung-3 floor)."""
    used: set[str] = {"as_measure", "XlError", "eval_instance"}
    scc_ids = frozenset(scc)
    compute_names = {sid: _compute_fn_name(sid) for sid in scc}
    lines: list[str] = [
        "    memo: dict[tuple[str, int], object] = {}",
        "    stack: set[tuple[str, int]] = set()",
        "",
    ]
    for sid in scc:
        series = catalog.get(sid)
        runs = formula_shape_runs(series, graph)
        if not runs:
            raise InvertedTreeExportError(f"series {sid!r} has no members to evaluate")
        fn = compute_names[sid]
        lines.append(f"    def {fn}(i: int) -> {python_measure_type(series)}:")
        for run_index, (_key, start, stop) in enumerate(runs):
            expr, expr_used = _emit_region_return(
                series,
                catalog=catalog,
                deps=deps,
                graph=graph,
                host_index=start,
                scc_ids=scc_ids,
                compute_names=compute_names,
            )
            used |= expr_used
            guarded = run_index < len(runs) - 1
            if guarded:
                lines.append(f"        if i < {stop}:")
                indent = "            "
            else:
                indent = "        "
            lines.append(f"{indent}try:")
            lines.append(f"{indent}    return {expr}")
            lines.append(f"{indent}except XlError as err:")
            lines.append(f"{indent}    return err.code")
        lines.append("")
    edges = collect_dependence_edges(catalog, graph, scc, edges=edges)
    intra = [edge for edge in edges if edge.consumer_id in scc_ids and edge.producer_id in scc_ids]
    reverse_drive = (
        bool(intra)
        and all(edge.distance <= 0 for edge in intra)
        and any(edge.distance < 0 for edge in intra)
    )
    returned: list[str] = []
    for sid in scc:
        series = catalog.get(sid)
        n = len(series.cells)
        fn = compute_names[sid]
        if series.is_scalar:
            lines.append(f"    {sid} = eval_instance({sid!r}, 0, {fn}, memo, stack)")
        elif reverse_drive:
            lines.append(f"    for i in reversed(range({n})):")
            lines.append(f"        eval_instance({sid!r}, i, {fn}, memo, stack)")
            lines.append(
                f"    {sid} = tuple(eval_instance({sid!r}, i, {fn}, memo, stack) for i in range({n}))"
            )
        else:
            lines.append(
                f"    {sid} = tuple(eval_instance({sid!r}, i, {fn}, memo, stack) for i in range({n}))"
            )
        returned.append(sid)
    lines.append(f"    return {', '.join(returned)}")
    return lines, used


def _indented(lines: list[str], spaces: int) -> list[str]:
    pad = " " * spaces
    return [f"{pad}{line}" if line else line for line in lines]


def _fused_template_index(
    series: BoundSeries,
    plan: FusedPlan,
    region: FusedRegion,
    catalog: SeriesCatalog,
    partition: tuple[Scalar, ...] | None = None,
) -> int:
    start, stop = plan.domain[series.series_id]
    union_t = max(region.start, start)
    if union_t >= min(region.stop, stop):
        union_t = max(0, stop - start - 1)
    target_coord = plan.schedule[union_t]
    for i, cell in enumerate(series.cells):
        if schedule_axis_coord(cell, catalog) != target_coord:
            continue
        if partition is not None and schedule_partition(cell, catalog) != partition:
            continue
        return i
    return max(0, stop - start - 1)


def _region_guard(region: FusedRegion) -> str:
    if region.stop == region.start + 1:
        return f"t == {region.start}"
    return f"{region.start} <= t < {region.stop}"


def _emit_fused_expr(
    series: BoundSeries,
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
    graph: DependencyGraph,
    host_index: int,
    plan: FusedPlan,
    ready: set[str],
    suffix: str = "",
    partition: tuple[Scalar, ...] | None = None,
    use_area_var: bool = False,
) -> tuple[str, set[str]]:
    ctx = EmitContext(
        host=series,
        catalog=catalog,
        deps=deps[series.series_id],
        host_index=host_index,
        host_cell=series.cells[host_index],
        index_var="t",
        prior_var=None,
        scc_ids=frozenset(plan.scc),
        fused_mode=True,
        fused_plan=plan,
        fused_ready=frozenset(ready),
        fused_buffer_suffix=suffix,
        fused_partition=partition,
        fused_use_area=use_area_var,
        graph=graph,
    )
    if try_formula_ast(graph, series.cells[host_index]) is None:
        return _emit_hole_expr(series, host_index, graph), set()
    expr = emit_expr(node_formula_ast(graph, series.cells[host_index]), ctx)
    return _as_measure_call(expr, series), set(ctx.used_runtime)


def _as_measure_literal(expr: str) -> str | None:
    """Return the inner literal of `as_measure(<literal>)`, if that is all it is."""
    prefix = "as_measure("
    if not expr.startswith(prefix) or not expr.endswith(")"):
        return None
    inner = expr[len(prefix) : -1]
    if not inner or inner[0] in {"(", "["} or "live_measure" in inner or "(" in inner:
        return None
    return inner


def _unify_area_exprs(exprs: Sequence[str]) -> str:
    """Return one expression, or an `_area`-indexed literal tuple.

    Structurally different expressions are joined with a short-circuiting
    `if/else` so side-effecting reads are not evaluated for every area.
    """
    if len(exprs) == 1 or len(set(exprs)) == 1:
        return exprs[0]
    literals = [_as_measure_literal(expr) for expr in exprs]
    if all(item is not None for item in literals):
        return f"as_measure(({', '.join(item for item in literals if item is not None)})[_area])"
    chain = exprs[0]
    for index, expr in enumerate(exprs[1:], start=1):
        chain = f"{chain} if _area == {index - 1} else {expr}"
    return chain


def _emit_fused_assign(series: BoundSeries, expr: str, suffix: str = "") -> list[str]:
    sid = f"{series.series_id}{suffix}"
    return [
        "try:",
        f"    {sid}_t = {expr}",
        "except XlError as err:",
        f"    {sid}_t = err.code",
        f"{sid}.append({sid}_t)",
    ]


def _emit_fused_region(
    plan: FusedPlan,
    region: FusedRegion,
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
    graph: DependencyGraph,
    suffix: str = "",
    partition: tuple[Scalar, ...] | None = None,
    area_partitions: Sequence[tuple[Scalar, ...]] = (),
) -> tuple[list[str], set[str]]:
    used: set[str] = set()
    lines: list[str] = []
    ready: set[str] = set()
    for sid in region.body_order:
        start, stop = plan.domain[sid]
        if stop <= region.start or start >= region.stop:
            continue
        series = catalog.get(sid)
        if area_partitions:
            exprs: list[str] = []
            for part in area_partitions:
                expr, expr_used = _emit_fused_expr(
                    series,
                    catalog=catalog,
                    deps=deps,
                    graph=graph,
                    host_index=_fused_template_index(series, plan, region, catalog, part),
                    plan=plan,
                    ready=ready,
                    suffix=suffix,
                    partition=part,
                    use_area_var=True,
                )
                used |= expr_used
                exprs.append(expr)
            expr = _unify_area_exprs(exprs)
        else:
            expr, expr_used = _emit_fused_expr(
                series,
                catalog=catalog,
                deps=deps,
                graph=graph,
                host_index=_fused_template_index(series, plan, region, catalog, partition),
                plan=plan,
                ready=ready,
                suffix=suffix,
                partition=partition,
            )
            used |= expr_used
        assign = _emit_fused_assign(series, expr, suffix)
        if start <= region.start and stop >= region.stop:
            lines.extend(assign)
        else:
            lines.append(f"if {start} <= t < {stop}:")
            lines.extend(_indented(assign, 4))
        ready.add(sid)
    return lines, used


def _emit_fused_loop(
    plan: FusedPlan,
    regions: Sequence[FusedRegion],
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
    graph: DependencyGraph,
    n: int,
    t_header: str,
    suffix: str = "",
    partition: tuple[Scalar, ...] | None = None,
    area_partitions: Sequence[tuple[Scalar, ...]] = (),
    indent: int = 4,
) -> tuple[list[str], set[str]]:
    """Emit `for t in ...` plus region bodies at `indent` spaces."""
    used: set[str] = set()
    lines = [t_header]
    multi = len(regions) > 1
    inner_indent = indent + 4
    for index, region in enumerate(regions):
        body, body_used = _emit_fused_region(
            plan,
            region,
            catalog=catalog,
            deps=deps,
            graph=graph,
            suffix=suffix,
            partition=partition,
            area_partitions=area_partitions,
        )
        used |= body_used
        if multi:
            if index == 0:
                lines.append(f"{' ' * inner_indent}if {_region_guard(region)}:")
            elif index + 1 == len(regions):
                lines.append(f"{' ' * inner_indent}else:")
            else:
                lines.append(f"{' ' * inner_indent}elif {_region_guard(region)}:")
            lines.extend(_indented(body, inner_indent + 4))
        else:
            lines.extend(_indented(body, inner_indent))
    return lines, used


def emit_rung2_scc(
    scc: tuple[str, ...],
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
    graph: DependencyGraph,
    plan: FusedPlan | None = None,
    edges: Sequence[DependenceEdge] | None = None,
) -> tuple[list[str], set[str]]:
    """Emit a fused union-domain loop for a fusible zipper SCC.

    Each `FusedRegion` is one residual-order / access-class span. The union
    schedule is the loop; look-ahead or non-contiguous domains must use
    `emit_rung3_scc`. A singleton SCC reuses this body as the rung-1 scan:
    self-lags index the growing buffer at compile-time offsets. A matrix
    nest wraps that body in `for _area in range(n)` (or unrolls when
    partitions are not isomorphic).

    Raises:
        InvertedTreeExportError: The SCC is not fusible, or the residual is a
            real same-index cycle.
    """
    if plan is None:
        plan = plan_fused_scc(scc, catalog=catalog, graph=graph, edges=edges)
    if plan is None:
        raise InvertedTreeExportError(
            f"zipper series {list(scc)!r} is not fusible; use demand-driven evaluation"
        )
    used: set[str] = {"as_measure", "XlError"}
    lines: list[str] = []
    for sid in scc:
        series = catalog.get(sid)
        lines.append(f"    {sid}: list[{python_measure_type(series)}] = []")
    n = len(plan.schedule)
    seq_params: list[str] = []
    if len(scc) == 1 and not plan.is_nested:
        guard_lines, guard_used = emit_sequence_length_guards(
            catalog.get(scc[0]), deps[scc[0]], catalog
        )
        used |= guard_used
        lines.extend(guard_lines)
        seq_params = ["n"]
        t_header = "    for t in range(n):"
    else:
        t_header = f"    for t in range({n}):"

    def _part_locals(indent: str) -> list[str]:
        return [
            f"{indent}{sid}_p: list[{python_measure_type(catalog.get(sid))}] = []" for sid in scc
        ]

    def _part_extend(indent: str) -> list[str]:
        return [f"{indent}{sid}.extend({sid}_p)" for sid in scc]

    if plan.unroll and plan.is_nested:
        for part, regions in zip(plan.partitions, plan.partition_regions, strict=True):
            lines.extend(_part_locals("    "))
            loop, loop_used = _emit_fused_loop(
                plan,
                regions,
                catalog=catalog,
                deps=deps,
                graph=graph,
                n=n,
                t_header=t_header,
                suffix="_p",
                partition=part,
            )
            used |= loop_used
            lines.extend(loop)
            lines.extend(_part_extend("    "))
    elif plan.is_nested:
        lines.append(f"    for _area in range({len(plan.partitions)}):")
        lines.extend(_part_locals("        "))
        loop, loop_used = _emit_fused_loop(
            plan,
            plan.regions,
            catalog=catalog,
            deps=deps,
            graph=graph,
            n=n,
            t_header="        for t in range(n):"
            if seq_params
            else f"        for t in range({n}):",
            suffix="_p",
            partition=plan.partitions[0],
            area_partitions=plan.partitions,
            indent=8,
        )
        used |= loop_used
        lines.extend(loop)
        lines.extend(_part_extend("        "))
    else:
        loop, loop_used = _emit_fused_loop(
            plan,
            plan.regions,
            catalog=catalog,
            deps=deps,
            graph=graph,
            n=n,
            t_header=t_header,
        )
        used |= loop_used
        lines.extend(loop)
    returned_items: list[str] = []
    for sid in scc:
        series = catalog.get(sid)
        if series.is_scalar:
            returned_items.append(f"{sid}[0]")
            continue
        if len(series.cells) > 1 and not plan.is_nested:
            u_first = plan.coord_to_t[schedule_axis_coord(series.cells[0], catalog)]
            u_last = plan.coord_to_t[schedule_axis_coord(series.cells[-1], catalog)]
            if u_first > u_last:
                returned_items.append(f"tuple(reversed({sid}))")
                continue
        returned_items.append(f"tuple({sid})")
    returned = ", ".join(returned_items)
    lines.append(f"    return {returned}")
    return lines, used
