"""Translate a bound series' Excel AST into a first-level-dep Python helper."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from excel_grapher.core.address_keys import CanonicalAddress, as_canonical
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
    resolve_cell_ref,
)
from excel_grapher.core.formula_shape import fingerprint_formula_shape
from excel_grapher.exporter.inverted_tree import runtime as inverted_runtime
from excel_grapher.exporter.inverted_tree.access import (
    AccessFunction,
    classify_producer_access,
)
from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    SeriesCatalog,
    covering_series,
    fit_affine_map,
    schedule_axis_coord,
    schedule_partition,
)
from excel_grapher.exporter.inverted_tree.deps import (
    DependenceEdge,
    SeriesDeps,
    iter_range_addresses,
    node_formula_ast,
    predecessor_address,
    range_column_addresses,
    successor_address,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.schedule import (
    FusedPlan,
    FusedRegion,
    collect_dependence_edges,
    plan_fused_scc,
)
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
    graph: DependencyGraph | None = None
    lookup_anchor_slot: int = 0

    def param(self, series_id: str) -> str:
        return series_id

    def use(self, symbol: str) -> str:
        self.used_runtime.add(symbol)
        return symbol


def python_measure_type(series: BoundSeries) -> str:
    """Return the Python type of one observation (`float | str` for numbers)."""
    if series.python_dtype in {"float", "int"}:
        return f"{series.python_dtype} | str"
    return series.python_dtype


def python_annotation(series: BoundSeries) -> str:
    """Return a typing annotation for a helper parameter."""
    inner = python_measure_type(series) if series.is_formula_series else series.python_dtype
    if series.is_scalar:
        return inner
    return f"Sequence[{inner}]"


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
    address = as_canonical(resolve_cell_ref(node, ctx.host_cell))
    if ctx.fused_mode:
        return _emit_fused_ref(address, ctx)
    if ctx.instance_mode:
        return _emit_instance_ref(address, ctx)
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


def _emit_fused_ref(address: CanonicalAddress, ctx: EmitContext) -> str:
    owner = ctx.catalog.require_series_for(address)
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
    index_expr = _affine_index_expr(idx - step * host_union, index_var, step=step)
    return f"live_measure({name}[{index_expr}])"


def _emit_instance_ref(address: CanonicalAddress, ctx: EmitContext) -> str:
    owner = ctx.catalog.require_series_for(address)
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
        key = fingerprint_formula_shape(node_formula_ast(graph, address)).shape_key
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
    if name == "MATCH":
        return _emit_match(node, ctx)
    if name == "TRUE":
        return "True"
    if name == "FALSE":
        return "False"
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
        ast = node_formula_ast(ctx.graph, cell)
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
    return classify_producer_access(ctx.host, producer, ctx.catalog, ctx.graph)


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
        covered_full = covering_series(ctx.catalog, iter_range_addresses(start, end))
        if covered_full is not None:
            slot = ctx.lookup_anchor_slot
            ctx.lookup_anchor_slot += 1
            return _emit_index_into_block(covered_full, start, row_expr, col_expr, ctx, slot)
        if col_literal is None:
            raise _host_export_error(
                ctx, "INDEX column is not a literal and the range is not one bound block"
            )
        column_cells = range_column_addresses(start, end, col_literal)
        covered = covering_series(ctx.catalog, column_cells)
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
        return ctx.catalog.require_series_for(as_canonical(resolve_cell_ref(node, ctx.host_cell)))
    if isinstance(node, RangeNode):
        start = resolve_cell_ref(node.start_ref, ctx.host_cell)
        end = resolve_cell_ref(node.end_ref, ctx.host_cell)
        covered = covering_series(ctx.catalog, iter_range_addresses(start, end))
        if covered is None:
            raise _host_export_error(ctx, f"reference {start}:{end} is not bound")
        return covered
    raise _host_export_error(ctx, "OFFSET/MATCH base must be a cell or range")


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


def emit_helper_body(
    series: BoundSeries,
    *,
    catalog: SeriesCatalog,
    deps: SeriesDeps,
    graph: DependencyGraph,
) -> tuple[list[str], set[str]]:
    """Return indented body lines and the runtime symbols they use."""
    used: set[str] = set()
    seq_params = [
        sid for sid in deps.param_ids if catalog.get(sid).is_sequence and sid in deps.aligned_ids
    ]
    if series.is_scalar:
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
        return _emit_scan_body(
            series, catalog=catalog, deps=deps, graph=graph, seq_params=seq_params
        )

    runs = formula_shape_runs(series, graph)
    used.add("as_measure")
    used.add("XlError")
    lines: list[str] = []
    if seq_params:
        joined = ", ".join(seq_params)
        used.add("require_aligned")
        lines.append(f"    n = require_aligned({joined})")
    else:
        lines.append(f"    n = {len(series.cells)}")
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
    seq_params: list[str],
) -> tuple[list[str], set[str]]:
    runs = formula_shape_runs(series, graph)
    seed = deps.seed_id
    used: set[str] = {"as_measure", "XlError", "is_error"}
    lines: list[str] = []
    if seq_params:
        used.add("require_aligned")
        lines.append(f"    n = require_aligned({', '.join(seq_params)})")
    else:
        lines.append(f"    n = {len(series.cells)}")
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
        graph=graph,
    )
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
        info = deps[scc[0]]
        seq_params = [
            sid
            for sid in info.param_ids
            if catalog.get(sid).is_sequence and sid in info.aligned_ids
        ]
    if seq_params:
        used.add("require_aligned")
        t_header = "    for t in range(n):"
        lines.append(f"    n = require_aligned({', '.join(seq_params)})")
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
