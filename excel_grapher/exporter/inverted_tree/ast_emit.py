"""Translate a bound series' Excel AST into a first-level-dep Python helper."""

from __future__ import annotations

from collections.abc import Iterator, Sequence
from dataclasses import dataclass, field, replace
from itertools import product
from typing import TYPE_CHECKING, Any, cast

from excel_grapher.core.address_keys import CanonicalAddress, as_canonical, parse_cell_coords
from excel_grapher.core.excel_function_names import normalize_excel_function_name
from excel_grapher.core.formula_ast import (
    AbsoluteAxis,
    AstNode,
    BinaryOpNode,
    BoolNode,
    CellRef,
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
from excel_grapher.exporter.inverted_tree import runtime as inverted_runtime
from excel_grapher.exporter.inverted_tree.access import (
    indirect_argument_addresses,
    indirect_target_addresses,
)
from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    KeyPoint,
    SeriesCatalog,
    Statement,
    covering_series,
)
from excel_grapher.exporter.inverted_tree.deps import (
    PositionalRangeCell,
    SeriesDeps,
    addresses_outside_blank_ranges,
    covering_series_for_index_window,
    current_blank_rects,
    index_call_is_ref,
    index_window_corners,
    iter_range_addresses,
    iter_ref_addresses,
    offset_index_destination,
    range_ref_label,
    resolve_offset_destination_series,
    resolve_positional_range,
    try_formula_ast,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher.blank_ranges import BlankRangeRect, address_in_blank_ranges

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
class EmitContext:
    """How to read bound series while lowering one host formula."""

    host: BoundSeries
    catalog: SeriesCatalog
    deps: SeriesDeps
    host_index: int
    host_cell: CanonicalAddress
    coordinate_vars: dict[str, str]
    used_runtime: set[str] = field(default_factory=set)
    scc_ids: frozenset[str] = field(default_factory=frozenset)
    graph: DependencyGraph | None = None
    blank_rects: tuple[BlankRangeRect, ...] = field(default_factory=current_blank_rects)
    array_context: bool = False
    named_axes: NamedAxes | None = None

    def param(self, series_id: str) -> str:
        return series_id

    def use(self, symbol: str) -> str:
        self.used_runtime.add(symbol)
        return symbol


def python_measure_type(series: BoundSeries) -> str:
    """Return the Python type of one observation (`float | str` for numbers)."""
    base = f"{series.python_dtype} | str" if series.python_dtype != "str" else series.python_dtype
    if series.has_none_holes or not series.single_valued:
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


def _emit_address(
    address: CanonicalAddress, ctx: EmitContext, *, ref: CellRefNode | None = None
) -> str:
    if address_in_blank_ranges(address, ctx.blank_rects):
        return "None"
    return _emit_named_address(address, ctx, ref=ref)


def _emit_named_address(
    address: CanonicalAddress, ctx: EmitContext, *, ref: CellRefNode | None = None
) -> str:
    """Read a producer by semantic coordinate relative to the host coordinate."""
    owner = ctx.catalog.require_series_for(address)
    name = ctx.param(owner.series_id)
    if owner.single_valued:
        # A scalar member of the recurrence group is a demand-driven reader.
        return f"{name}[()]" if owner.series_id in ctx.scc_ids else name
    index = owner.index_of(address)
    if index is None:
        raise InvertedTreeExportError(f"series {owner.series_id!r}: no coordinate for {address}")
    return f"{name}[{', '.join(_named_keys(owner, owner.domain[index], ctx, ref=ref))}]"


def _iter_cell_refs(node: AstNode | None) -> Iterator[CellRef]:
    """Yield every cell and range endpoint reference of `node` in formula order."""
    if node is None:
        return
    match node:
        case CellRefNode(ref):
            yield ref
        case RangeNode(start_ref, end_ref):
            yield start_ref
            yield end_ref
        case FunctionCallNode(_name, args):
            for arg in args:
                yield from _iter_cell_refs(arg)
        case BinaryOpNode(_op, left, right):
            yield from _iter_cell_refs(left)
            yield from _iter_cell_refs(right)
        case UnaryOpNode(_op, operand):
            yield from _iter_cell_refs(operand)
        case _:
            return


def _neighbor_host_cell(ctx: EmitContext, axis: str) -> CanonicalAddress | None:
    """The host series cell one step along `axis`, or the previous one at the edge."""
    sheet, row, col = parse_cell_coords(ctx.host_cell)
    candidates = (
        [(row, col + 1), (row, col - 1)] if axis == "col" else [(row + 1, col), (row - 1, col)]
    )
    cache = ctx.host._emit_cache
    cells = cache.get("positions")
    if cells is None:
        cells = cache["positions"] = {parse_cell_coords(cell): cell for cell in ctx.host.cells}
    for candidate_row, candidate_col in candidates:
        cell = cells.get((sheet, candidate_row, candidate_col))
        if cell is not None:
            return cell
    return None


def _cell_refs_of(ctx: EmitContext, cell: CanonicalAddress) -> list[CellRef]:
    """References of the formula at `cell`, cached on the host series."""
    assert ctx.graph is not None
    cache = ctx.host._emit_cache
    refs = cache.get(("refs", cell))
    if refs is None:
        refs = cache[("refs", cell)] = list(_iter_cell_refs(try_formula_ast(ctx.graph, cell)))
    return refs


def _reference_stays_put(ctx: EmitContext, ref: CellRef, axis: str) -> bool:
    """True when the neighbouring host cell reaches the same target through `ref`.

    A relative reference authored per cell can still point at one fixed
    cell; the family then keys that cell literally instead of unrolling a
    different displacement at every coordinate.
    """
    if ctx.graph is None:
        return False
    neighbor = _neighbor_host_cell(ctx, axis)
    if neighbor is None:
        return False
    host_refs = _cell_refs_of(ctx, ctx.host_cell)
    neighbor_refs = _cell_refs_of(ctx, neighbor)
    if len(host_refs) != len(neighbor_refs):
        return False
    index = next((i for i, candidate in enumerate(host_refs) if candidate is ref), None)
    if index is None:
        return False
    target = resolve_cell_ref(ref, ctx.host_cell)
    return resolve_cell_ref(neighbor_refs[index], neighbor) == target


def _pinned_axes(ref: CellRefNode | None, ctx: EmitContext) -> set[str]:
    """Worksheet axes on which `ref` reaches one fixed cell that is not the host."""
    if ref is None:
        return set()
    address = resolve_cell_ref(ref, ctx.host_cell)
    _host_sheet, host_row, host_col = parse_cell_coords(ctx.host_cell)
    _sheet, row, col = parse_cell_coords(address)
    pinned: set[str] = set()
    for axis, host_position, position, axis_ref in (
        ("col", host_col, col, ref.ref.col),
        ("row", host_row, row, ref.ref.row),
    ):
        if _reference_stays_put(ctx, ref.ref, axis) or (
            position != host_position and isinstance(axis_ref, AbsoluteAxis)
        ):
            pinned.add(axis)
    return pinned


def _integer_driver(
    ctx: EmitContext, key_field: str, field_axis: str | None
) -> tuple[str, int] | None:
    """The host loop variable an integer key of `key_field` moves with."""
    from excel_grapher.exporter.inverted_tree.deps import _key_field_axis

    host_point = ctx.host.domain[ctx.host_index].as_mapping()
    candidates = [
        (field, variable)
        for field, variable in ctx.coordinate_vars.items()
        if type(host_point.get(field)) is int
    ]
    if not candidates:
        return None
    ranked = sorted(
        candidates,
        key=lambda item: (
            item[0] != key_field,
            _key_field_axis(ctx.host, item[0]) != field_axis,
            item[0] != "TIME_PERIOD",
        ),
    )
    field, variable = ranked[0]
    return variable, cast(int, host_point[field])


def _key_template(target: str, ctx: EmitContext) -> str | None:
    """An f-string over a host key embedded in the label `target`."""
    host_point = ctx.host.domain[ctx.host_index].as_mapping()
    best: tuple[str, str] | None = None
    for key_field, variable in ctx.coordinate_vars.items():
        value = host_point.get(key_field)
        if not isinstance(value, str) or len(value) < 2 or value == target or value not in target:
            continue
        if any(char in value for char in "'\\\"{}"):
            continue
        if best is None or len(value) > len(best[0]):
            best = (value, variable)
    if best is None:
        return None
    value, variable = best
    quoted = repr(target)
    inner = quoted[1:-1].replace("{", "{{").replace("}", "}}")
    return f"f{quoted[0]}{inner.replace(value, '{' + variable + '}')}{quoted[0]}"


def _named_keys(
    owner: BoundSeries,
    point: KeyPoint,
    ctx: EmitContext,
    *,
    ref: CellRefNode | None = None,
) -> list[str]:
    """Express each key of `point` relative to the host loop variables.

    A key equal to the host's own key is the loop variable. An integer key
    reached through a moving reference is the host's period variable plus
    the authored difference; formula-family grouping then verifies the same
    difference at every coordinate sharing the expression. A label that
    embeds the host's own key is a template over it. Keys reached through a
    fixed reference (`$`, or the same cell from every host cell) stay literal.
    """
    from excel_grapher.exporter.inverted_tree.deps import _key_field_axis

    assert ctx.coordinate_vars is not None
    host_point = ctx.host.domain[ctx.host_index].as_mapping()
    pinned = _pinned_axes(ref, ctx)
    keys = []
    for key_field in owner.key_fields:
        variable = ctx.coordinate_vars.get(key_field)
        current = host_point.get(key_field)
        target = point[key_field]
        field_axis = _key_field_axis(owner, key_field)
        if field_axis in pinned:
            keys.append(repr(target))
            continue
        if variable is not None and current == target:
            keys.append(variable)
            continue
        if type(target) is int:
            driver = _integer_driver(ctx, key_field, field_axis)
            if driver is not None:
                variable, current = driver
                difference = target - current
                sign = "+" if difference > 0 else "-"
                keys.append(variable if difference == 0 else f"{variable} {sign} {abs(difference)}")
                continue
        if isinstance(target, str):
            template = _key_template(target, ctx)
            if template is not None:
                keys.append(template)
                continue
        keys.append(repr(target))
    return keys


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
        lazy = ctx
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
        return call
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

    return str(coordinate(ctx.host_cell))


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
    then = emit_expr(node.args[1], ctx)
    # Array omitted else is Excel FALSE; scalar emit still uses 0.
    otherwise = emit_expr(node.args[2], ctx) if len(node.args) > 2 else "0"
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
    run = _cell_run(node.args[1:], ctx)
    if run is not None:
        view = _named_range_view(run, ctx)
        if view is not None:
            return f"{ctx.use('xl_choose_range')}({index}, {view})"
    lazy = ctx
    choices = ", ".join(f"lambda: {emit_expr(arg, lazy)}" for arg in node.args[1:])
    return f"{ctx.use('xl_choose_lazy')}({index}, {choices})"


def _cell_run(args: Sequence[AstNode], ctx: EmitContext) -> RangeNode | None:
    """The range spanned by cell references listed along one worksheet row or column."""
    if len(args) < 2 or not all(isinstance(arg, CellRefNode) for arg in args):
        return None
    refs = [cast(CellRefNode, arg) for arg in args]
    located = [parse_cell_coords(resolve_cell_ref(ref, ctx.host_cell)) for ref in refs]
    sheets = {sheet for sheet, _row, _col in located}
    if len(sheets) != 1:
        return None
    rows = [row for _sheet, row, _col in located]
    cols = [col for _sheet, _row, col in located]
    across = len(set(rows)) == 1 and cols == list(range(cols[0], cols[0] + len(cols)))
    down = len(set(cols)) == 1 and rows == list(range(rows[0], rows[0] + len(rows)))
    if not (across or down):
        return None
    return RangeNode(start_ref=refs[0].ref, end_ref=refs[-1].ref)


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
    named_view = _named_range_view(node, ctx, addresses)
    if named_view is not None:
        return named_view
    return _python_tuple([_emit_address(address, ctx) for address in addresses])


def _emit_lookup_arg(node: AstNode, ctx: EmitContext) -> str:
    if isinstance(node, (RangeNode, WholeColumnNode, WholeRowNode)):
        return _emit_range_table(node, ctx)
    return emit_expr(node, ctx)


def _python_tuple(items: Sequence[str]) -> str:
    """Emit a Python tuple literal; one-element rows keep a trailing comma."""
    if len(items) == 1:
        return f"({items[0]},)"
    return f"({', '.join(items)})"


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


def _positional_table_source(cells: Sequence[PositionalRangeCell], ctx: EmitContext) -> str:
    """Emit an irregular range as a table of per-cell callbacks.

    Cells are read by literal coordinate so the table never depends on the
    host loop variable and can be built once per call.
    """
    rows = _group_positional_rows(cells)
    table_ctx = replace(ctx, coordinate_vars={})
    callbacks = _python_tuple([_python_tuple(_table_row_parts(row, table_ctx)) for row in rows])
    return f"{ctx.use('lazy_table')}({callbacks})"


def _table_row_parts(row: Sequence[PositionalRangeCell], ctx: EmitContext) -> list[str]:
    """One view per run of a series' cells in a table row, callbacks elsewhere."""
    parts: list[str] = []
    index = 0
    while index < len(row):
        cell = row[index]
        end = index
        while (
            cell.series_id is not None
            and end + 1 < len(row)
            and row[end + 1].series_id == cell.series_id
            and not row[end + 1].blank
        ):
            end += 1
        view = None
        if end > index:
            view = _named_range_view(RangeNode(cell.address, row[end].address), ctx)
        if view is not None:
            parts.append(view)
            index = end + 1
            continue
        parts.append(f"lambda: {_emit_positional_cell(cell, ctx)}")
        index += 1
    return parts


def _emit_positional_cell(cell: PositionalRangeCell, ctx: EmitContext) -> str:
    """Emit one MATCH/INDEX window cell by literal coordinate."""
    if cell.blank:
        return "None"
    return _emit_address(cell.address, ctx)


def _named_range_view(
    node: AstNode,
    ctx: EmitContext,
    addresses: Sequence[CanonicalAddress] | None = None,
) -> str | None:
    """Emit a lazy `view` when a range enumerates a product of one series' keys.

    Every key field of the owning series selects a run of its axis: the
    whole axis, a `span` between the corner keys expressed relative to the
    host coordinate, or one key. Fields that vary down the worksheet form
    the row product and fields that vary across it the column product; the
    range lowers only when that product reproduces the cells in worksheet
    order. Nested layouts select keys per field name.
    """
    from excel_grapher.exporter.inverted_tree.deps import _key_field_axis

    if not isinstance(node, RangeNode) or ctx.coordinate_vars is None or ctx.named_axes is None:
        return None
    start = as_canonical(resolve_cell_ref(node.start_ref, ctx.host_cell))
    end = as_canonical(resolve_cell_ref(node.end_ref, ctx.host_cell))
    if parse_cell_coords(start)[0] != parse_cell_coords(end)[0]:
        return None
    if addresses is None:
        addresses = iter_range_addresses(start, end)
        if any(address_in_blank_ranges(address, ctx.blank_rects) for address in addresses):
            return None
    if not addresses:
        return None
    owner = covering_series(ctx.catalog, addresses)
    if owner is None or owner.is_scalar or owner.single_valued:
        return None
    indices = [owner.index_of(address) for address in addresses]
    if any(index is None for index in indices):
        return None
    points = [owner.domain[cast(int, index)] for index in indices]
    first_keys = _named_keys(owner, points[0], ctx, ref=CellRefNode(node.start_ref))
    last_keys = _named_keys(owner, points[-1], ctx, ref=CellRefNode(node.end_ref))
    positions = [parse_cell_coords(address)[1:] for address in addresses]
    domain = owner.tensor_domain
    selections: dict[str, str] = {}
    seen_keys: dict[str, tuple[Any, ...]] = {}
    row_fields: list[str] = []
    col_fields: list[str] = []
    for position, key_field in enumerate(owner.key_fields):
        axis = domain.axes[position]
        keys = axis.keys
        seen = tuple(dict.fromkeys(point[key_field] for point in points))
        first, last = keys.index(seen[0]), keys.index(seen[-1])
        if seen != keys[first : last + 1]:
            return None
        constant = ctx.named_axes.constant(axis)
        start_key, end_key = first_keys[position], last_keys[position]
        literal_corners = start_key == repr(keys[first]) and end_key == repr(keys[last])
        if literal_corners and first == 0 and last == len(keys) - 1:
            expr = f"data.{constant}.keys"
        elif start_key == end_key:
            expr = f"({start_key},)"
        else:
            expr = f"{ctx.use('span')}(data.{constant}, {start_key}, {end_key})"
        selections[key_field] = expr
        seen_keys[key_field] = seen
        by_row: dict[int, set[Any]] = {}
        by_col: dict[int, set[Any]] = {}
        for point, (row, col) in zip(points, positions, strict=True):
            by_row.setdefault(row, set()).add(point[key_field])
            by_col.setdefault(col, set()).add(point[key_field])
        varies_in_row = any(len(values) > 1 for values in by_row.values())
        varies_in_col = any(len(values) > 1 for values in by_col.values())
        if varies_in_row and not varies_in_col:
            col_fields.append(key_field)
        elif varies_in_col and not varies_in_row:
            row_fields.append(key_field)
        elif not varies_in_row and not varies_in_col:
            (row_fields if _key_field_axis(owner, key_field) == "row" else col_fields).append(
                key_field
            )
        else:
            return None

    def runs(key_field: str) -> int:
        return sum(
            1
            for before, after in zip(points, points[1:], strict=False)
            if before[key_field] != after[key_field]
        )

    row_fields.sort(key=runs)
    col_fields.sort(key=runs)
    expected = [
        {
            **dict(zip(row_fields, row_keys, strict=True)),
            **dict(zip(col_fields, col_keys, strict=True)),
        }
        for row_keys in product(*(seen_keys[field] for field in row_fields))
        for col_keys in product(*(seen_keys[field] for field in col_fields))
    ]
    actual = [{field: point[field] for field in owner.key_fields} for point in points]
    if expected != actual:
        return None
    args = [ctx.param(owner.series_id)]
    if len(row_fields) <= 1 and len(col_fields) <= 1:
        if row_fields:
            args.append(f"rows={selections[row_fields[0]]}")
        if col_fields:
            args.append(f"cols={selections[col_fields[0]]}")
        if row_fields and col_fields and owner.key_fields[0] == col_fields[0]:
            args.append("cols_first=True")
    else:
        if row_fields:
            args.append(
                "rows={"
                + ", ".join(f"{field!r}: {selections[field]}" for field in row_fields)
                + "}"
            )
        if col_fields:
            args.append(
                "cols={"
                + ", ".join(f"{field!r}: {selections[field]}" for field in col_fields)
                + "}"
            )
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
    return _positional_table_source(cells, ctx)


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


def _emit_offset(node: FunctionCallNode, ctx: EmitContext) -> str:
    if len(node.args) < 3:
        raise _host_export_error(ctx, "OFFSET expects anchor, rows, cols")
    return _emit_named_offset(node, ctx)


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
    if table.single_valued or table.is_scalar:
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


def _offset_index_provably_ref(index_node: FunctionCallNode, ctx: EmitContext) -> bool:
    """True when every host cell in the current statement yields INDEX `#REF!`."""
    cells = _statement_cells(ctx) or (ctx.host_cell,)
    return all(index_call_is_ref(index_node, cell) for cell in cells)


def _row_column_args_omitted(node: FunctionCallNode) -> bool:
    return not node.args or (len(node.args) == 1 and isinstance(node.args[0], EmptyArgNode))


def _host_coord_expr(ctx: EmitContext, *, axis: str) -> str:
    """Return the host cell's worksheet row or column."""
    _sheet, row, col = parse_cell_coords(ctx.host_cell)
    return str(row if axis == "row" else col)


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


def _emit_indirect(node: FunctionCallNode, ctx: EmitContext) -> str:
    if ctx.graph is None:
        raise _host_export_error(ctx, "INDIRECT has no graph to classify")
    exclude = indirect_argument_addresses(node, ctx.host_cell)
    targets = indirect_target_addresses(ctx.graph, ctx.host_cell, exclude=tuple(exclude))
    if not targets:
        raise _host_export_error(ctx, "INDIRECT has no resolved edges")
    if len(targets) != 1:
        raise _host_export_error(ctx, "INDIRECT resolves to more than one cell")
    return _emit_address(targets[0], ctx)


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
        # Omitted and zero selectors request vectors; the lazy table keeps
        # their shape and Excel's single-row special case.
        table = (
            _emit_range_table(node.args[0], ctx)
            if isinstance(node.args[0], RangeNode)
            else _emit_value_or_range(node.args[0], ctx)
        )
        return f"{ctx.use('xl_index')}({table}, {row_expr}, {col_expr})"
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
                "(" + ", ".join(f"({_emit_positional_cell(cell, ctx)},)" for cell in cells) + ",)"
            )
            return f"{ctx.use('xl_index')}({table}, {row_expr}, 1)"
    table = _emit_value_or_range(node.args[0], ctx)
    return f"{ctx.use('xl_index')}({table}, {row_expr}, {col_expr})"


def _emit_match_array(node: AstNode, ctx: EmitContext) -> str:
    """Emit a MATCH lookup vector that preserves Excel positions."""
    return _emit_value_or_range(node, ctx)


def _emit_match(node: FunctionCallNode, ctx: EmitContext) -> str:
    if len(node.args) < 2:
        raise InvertedTreeExportError(
            f"series {ctx.host.series_id!r}: MATCH expects lookup and array"
        )
    lookup = emit_expr(node.args[0], ctx)
    array = _emit_match_array(node.args[1], ctx)
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
