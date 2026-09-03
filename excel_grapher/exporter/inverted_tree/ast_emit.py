"""Translate a bound series' Excel AST into a first-level-dep Python helper."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from excel_grapher.core.address_keys import normalize_key as normalize_address
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
from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, SeriesCatalog, covering_series
from excel_grapher.exporter.inverted_tree.deps import (
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
    schedule_coord,
)

if TYPE_CHECKING:
    from excel_grapher.grapher.graph import DependencyGraph

_COMPARE_OPS = {
    "=": "==",
    "<>": "!=",
    "<": "<",
    ">": ">",
    "<=": "<=",
    ">=": ">=",
}
_ARITHMETIC_OPS = {"+", "-", "*", "/"}
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
    host_cell: str
    index_var: str | None
    prior_var: str | None
    used_runtime: set[str] = field(default_factory=set)
    scc_ids: frozenset[str] = field(default_factory=frozenset)
    instance_mode: bool = False
    compute_names: dict[str, str] = field(default_factory=dict)
    fused_mode: bool = False
    fused_plan: FusedPlan | None = None
    fused_ready: frozenset[str] = field(default_factory=frozenset)

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


def _emit_cell_ref(node: CellRefNode, ctx: EmitContext) -> str:
    address = resolve_cell_ref(node, ctx.host_cell)
    if ctx.fused_mode:
        return _emit_fused_ref(address, ctx)
    if ctx.instance_mode:
        return _emit_instance_ref(address, ctx)
    pred = predecessor_address(ctx.host, ctx.host_index, ctx.catalog)
    if pred is not None and normalize_address(address) == normalize_address(pred) and ctx.prior_var:
        return ctx.prior_var
    succ = successor_address(ctx.host, ctx.host_index, ctx.catalog)
    if succ is not None and normalize_address(address) == normalize_address(succ) and ctx.prior_var:
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
    if owner.series_id in ctx.deps.aligned_ids and ctx.index_var is not None:
        return f"{name}[{ctx.index_var}]"
    if idx is not None and ctx.index_var is not None and owner.is_sequence:
        return f"{name}[{_index_expr(idx - ctx.host_index, ctx.index_var)}]"
    if idx is not None and ctx.index_var is None:
        return f"{name}[{idx}]" if not owner.is_scalar else name
    if owner.series_id in ctx.deps.lookup_ids:
        return name
    return name


def _index_expr(offset: int, index_var: str) -> str:
    if offset == 0:
        return index_var
    if offset > 0:
        return f"{index_var} + {offset}"
    return f"{index_var} - {-offset}"


def _emit_fused_ref(address: str, ctx: EmitContext) -> str:
    owner = ctx.catalog.require_series_for(address)
    idx = owner.index_of(address)
    if idx is None or ctx.fused_plan is None:
        raise InvertedTreeExportError(
            f"series {ctx.host.series_id!r}: fused ref {address} is unbound"
        )
    plan = ctx.fused_plan
    coord_to_t = plan.coord_to_t
    host_coord = schedule_coord(ctx.host_cell, ctx.catalog)
    prod_coord = schedule_coord(address, ctx.catalog)
    host_union = coord_to_t[host_coord]
    prod_union = coord_to_t[prod_coord]
    index_var = ctx.index_var or "t"
    ctx.use("live_measure")
    if owner.series_id in ctx.scc_ids:
        delta = prod_union - host_union
        if delta == 0:
            if owner.series_id not in ctx.fused_ready:
                raise InvertedTreeExportError(
                    f"series {ctx.host.series_id!r}: same-index read of "
                    f"{owner.series_id!r} before it is written"
                )
            return f"live_measure({owner.series_id}_t)"
        prod_start = plan.domain[owner.series_id][0]
        index_expr = _index_expr(delta - prod_start, index_var)
        return f"live_measure({owner.series_id}[{index_expr}])"
    name = ctx.param(owner.series_id)
    if owner.is_scalar:
        return f"live_measure({name})"
    return f"live_measure({name}[{_index_expr(idx - host_union, index_var)}])"


def _emit_instance_ref(address: str, ctx: EmitContext) -> str:
    owner = ctx.catalog.require_series_for(address)
    idx = owner.index_of(address)
    if idx is None:
        raise InvertedTreeExportError(
            f"series {ctx.host.series_id!r}: instance ref {address} is unbound"
        )
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
    if op == "/":
        return f"{ctx.use('xl_div')}({left}, {right})"
    if op == "^":
        return f"({left} ** {right})"
    if op == "&":
        return f"(str({left}) + str({right}))"
    if op in _ARITHMETIC_OPS:
        return f"({left} {op} {right})"
    if op in _COMPARE_OPS:
        return f"({left} {_COMPARE_OPS[op]} {right})"
    raise InvertedTreeExportError(f"series {ctx.host.series_id!r}: unsupported operator {op!r}")


def _emit_unary(node: UnaryOpNode, ctx: EmitContext) -> str:
    operand = emit_expr(node.operand, ctx)
    if node.op == "-":
        return f"(-{operand})"
    if node.op == "+":
        return operand
    if node.op == "%":
        return f"({operand} / 100.0)"
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


def _emit_offset(node: FunctionCallNode, ctx: EmitContext) -> str:
    if len(node.args) < 3:
        raise InvertedTreeExportError(
            f"series {ctx.host.series_id!r}: OFFSET expects anchor, rows, cols"
        )
    table = _series_for_ref(node.args[0], ctx)
    rows = emit_expr(node.args[1], ctx)
    cols = emit_expr(node.args[2], ctx)
    name = ctx.param(table.series_id)
    # 1-row table: OFFSET(anchor, 0, n) → table[n] with 0-based n.
    index = f"({cols})" if rows == "0" or rows == "0.0" else f"(({rows}) * len({name}) + ({cols}))"
    return f"{ctx.use('xl_at')}({name}, {index})"


def _emit_index(node: FunctionCallNode, ctx: EmitContext) -> str:
    if len(node.args) < 2:
        raise InvertedTreeExportError(
            f"series {ctx.host.series_id!r}: INDEX expects a range and row"
        )
    row_expr = emit_expr(node.args[1], ctx)
    col_arg = node.args[2] if len(node.args) > 2 else None
    col_index = 1
    if isinstance(col_arg, NumberNode):
        col_index = int(col_arg.value)
    if isinstance(node.args[0], RangeNode):
        start = resolve_cell_ref(node.args[0].start_ref, ctx.host_cell)
        end = resolve_cell_ref(node.args[0].end_ref, ctx.host_cell)
        column_cells = range_column_addresses(start, end, col_index)
        covered = covering_series(ctx.catalog, column_cells)
        if covered is None:
            raise InvertedTreeExportError(
                f"series {ctx.host.series_id!r}: INDEX column is not a bound series"
            )
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
        return ctx.catalog.require_series_for(resolve_cell_ref(node, ctx.host_cell))
    if isinstance(node, RangeNode):
        start = resolve_cell_ref(node.start_ref, ctx.host_cell)
        end = resolve_cell_ref(node.end_ref, ctx.host_cell)
        covered = covering_series(ctx.catalog, iter_range_addresses(start, end))
        if covered is None:
            raise InvertedTreeExportError(
                f"series {ctx.host.series_id!r}: reference {start}:{end} is not bound"
            )
        return covered
    raise InvertedTreeExportError(
        f"series {ctx.host.series_id!r}: OFFSET/MATCH base must be a cell or range"
    )


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
    )
    expr = emit_expr(node_formula_ast(graph, series.cells[host_index]), ctx)
    return _as_measure_call(expr, series), set(ctx.used_runtime)


def emit_rung3_scc(
    scc: tuple[str, ...],
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
    graph: DependencyGraph,
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
    edges = collect_dependence_edges(catalog, graph, scc)
    intra = [edge for edge in edges if edge.consumer_id in scc_ids and edge.producer_id in scc_ids]
    reverse_drive = (
        bool(intra)
        and all(edge.distance <= 0 for edge in intra)
        and any(edge.distance < 0 for edge in intra)
    )
    returned: list[str] = []
    for sid in scc:
        n = len(catalog.get(sid).cells)
        fn = compute_names[sid]
        if reverse_drive:
            lines.append(
                f"    {sid} = tuple(eval_instance({sid!r}, i, {fn}, memo, stack) for i in reversed(range({n})))[::-1]"
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
) -> int:
    start, stop = plan.domain[series.series_id]
    union_t = max(region.start, start)
    if union_t >= min(region.stop, stop):
        union_t = max(0, stop - start - 1)
    target_coord = plan.schedule[union_t]
    for i, cell in enumerate(series.cells):
        if schedule_coord(cell, catalog) == target_coord:
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
    )
    expr = emit_expr(node_formula_ast(graph, series.cells[host_index]), ctx)
    return _as_measure_call(expr, series), set(ctx.used_runtime)


def _emit_fused_assign(series: BoundSeries, expr: str) -> list[str]:
    sid = series.series_id
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
) -> tuple[list[str], set[str]]:
    used: set[str] = set()
    lines: list[str] = []
    ready: set[str] = set()
    for sid in region.body_order:
        start, stop = plan.domain[sid]
        if stop <= region.start or start >= region.stop:
            continue
        series = catalog.get(sid)
        expr, expr_used = _emit_fused_expr(
            series,
            catalog=catalog,
            deps=deps,
            graph=graph,
            host_index=_fused_template_index(series, plan, region, catalog),
            plan=plan,
            ready=ready,
        )
        used |= expr_used
        assign = _emit_fused_assign(series, expr)
        if start <= region.start and stop >= region.stop:
            lines.extend(assign)
        else:
            lines.append(f"if {start} <= t < {stop}:")
            lines.extend(_indented(assign, 4))
        ready.add(sid)
    return lines, used


def emit_rung2_scc(
    scc: tuple[str, ...],
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
    graph: DependencyGraph,
) -> tuple[list[str], set[str]]:
    """Emit a fused union-domain loop for a fusible zipper SCC.

    Each `FusedRegion` is one residual-order / access-class span. The union
    schedule is the loop; look-ahead or non-contiguous domains must use
    `emit_rung3_scc`.

    Raises:
        InvertedTreeExportError: The SCC is not fusible, or the residual is a
            real same-index cycle.
    """
    plan = plan_fused_scc(scc, catalog=catalog, graph=graph)
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
    lines.append(f"    for t in range({n}):")
    multi = len(plan.regions) > 1
    for index, region in enumerate(plan.regions):
        body, body_used = _emit_fused_region(plan, region, catalog=catalog, deps=deps, graph=graph)
        used |= body_used
        if multi:
            if index == 0:
                lines.append(f"        if {_region_guard(region)}:")
            elif index + 1 == len(plan.regions):
                lines.append("        else:")
            else:
                lines.append(f"        elif {_region_guard(region)}:")
            lines.extend(_indented(body, 12))
        else:
            lines.extend(_indented(body, 8))
    returned_items: list[str] = []
    for sid in scc:
        series = catalog.get(sid)
        if len(series.cells) > 1:
            u_first = plan.coord_to_t[schedule_coord(series.cells[0], catalog)]
            u_last = plan.coord_to_t[schedule_coord(series.cells[-1], catalog)]
            if u_first > u_last:
                returned_items.append(f"tuple(reversed({sid}))")
                continue
        returned_items.append(f"tuple({sid})")
    returned = ", ".join(returned_items)
    lines.append(f"    return {returned}")
    return lines, used
