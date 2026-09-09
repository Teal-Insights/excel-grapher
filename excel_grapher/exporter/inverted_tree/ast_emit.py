"""Translate a bound series' Excel AST into a first-level-dep Python helper."""

from __future__ import annotations

from collections.abc import Sequence
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
    lazy = ctx
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
    named_view = _named_range_view(node, ctx)
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
    callbacks = _python_tuple(
        [
            _python_tuple([f"lambda: {_emit_positional_cell(cell, table_ctx)}" for cell in row])
            for row in rows
        ]
    )
    return f"{ctx.use('lazy_table')}({callbacks})"


def _emit_positional_cell(cell: PositionalRangeCell, ctx: EmitContext) -> str:
    """Emit one MATCH/INDEX window cell by literal coordinate."""
    if cell.blank:
        return "None"
    return _emit_address(cell.address, ctx)


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
