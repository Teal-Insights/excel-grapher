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
from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, SeriesCatalog, covering_series
from excel_grapher.exporter.inverted_tree.deps import (
    SeriesDeps,
    iter_range_addresses,
    node_formula_ast,
    predecessor_address,
    range_column_addresses,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError

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
    pred = predecessor_address(ctx.host, ctx.host_index, ctx.catalog)
    if pred is not None and normalize_address(address) == normalize_address(pred) and ctx.prior_var:
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
    if idx is not None and ctx.index_var is None:
        return f"{name}[{idx}]" if not owner.is_scalar else name
    if owner.series_id in ctx.deps.lookup_ids:
        return name
    if ctx.index_var is not None and owner.is_sequence:
        return f"{name}[{ctx.index_var}]"
    return name


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

    ctx = EmitContext(
        host=series,
        catalog=catalog,
        deps=deps,
        host_index=0,
        host_cell=series.cells[0],
        index_var="i",
        prior_var=None,
    )
    ast = node_formula_ast(graph, series.cells[0])
    expr = emit_expr(ast, ctx)
    used |= ctx.used_runtime
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
    coerce = (
        f"as_measure({expr})"
        if series.python_dtype == "float"
        else f"as_measure({expr}, {series.python_dtype!r})"
    )
    lines.append(f"    out: list[{measure}] = []")
    lines.append("    for i in range(n):")
    lines.append("        try:")
    lines.append(f"            out.append({coerce})")
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
    seed = deps.seed_id
    ctx = EmitContext(
        host=series,
        catalog=catalog,
        deps=deps,
        host_index=0,
        host_cell=series.cells[0],
        index_var="i",
        prior_var="prior",
    )
    ast = node_formula_ast(graph, series.cells[0])
    expr = emit_expr(ast, ctx)
    used = set(ctx.used_runtime)
    used.add("as_measure")
    used.add("XlError")
    used.add("is_error")
    lines: list[str] = []
    if seq_params:
        used.add("require_aligned")
        lines.append(f"    n = require_aligned({', '.join(seq_params)})")
    else:
        lines.append(f"    n = {len(series.cells)}")
    seed_expr = seed if seed is not None else "0"
    measure = python_measure_type(series)
    coerce = (
        f"as_measure({expr})"
        if series.python_dtype == "float"
        else f"as_measure({expr}, {series.python_dtype!r})"
    )
    lines.append(f"    path: list[{measure}] = []")
    lines.append(f"    prior: {measure} = {seed_expr}")
    lines.append("    for i in range(n):")
    lines.append("        if is_error(prior):")
    lines.append("            path.append(prior)")
    lines.append("            continue")
    lines.append("        try:")
    lines.append(f"            prior = {coerce}")
    lines.append("        except XlError as err:")
    lines.append("            prior = err.code")
    lines.append("        path.append(prior)")
    lines.append("    return tuple(path)")
    return lines, used
