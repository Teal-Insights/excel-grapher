"""First-level series dependencies, leaf closures, and fail-closed checks."""

from __future__ import annotations

from collections.abc import Iterable, Iterator, Mapping, Sequence
from contextlib import contextmanager
from contextvars import ContextVar, Token
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import TYPE_CHECKING, Any, Literal, cast

from fastpyxl.utils.cell import get_column_letter

from excel_grapher.core.address_keys import (
    CanonicalAddress,
    as_canonical,
    canonical_address,
    format_cell_key,
    format_key,
    parse_cell_coords,
)
from excel_grapher.core.addressing import index_excel_range
from excel_grapher.core.excel_function_meta import is_ref_only_arg
from excel_grapher.core.excel_function_names import (
    normalize_excel_function_name as _normalize_excel_function_name,
)
from excel_grapher.core.formula_ast import (
    AbsoluteAxis,
    AstNode,
    BinaryOpNode,
    CellRefNode,
    EmptyArgNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    UnaryOpNode,
    WholeColumnNode,
    WholeRowNode,
    resolve_cell_ref,
    resolve_whole_column_ref,
    resolve_whole_row_ref,
)
from excel_grapher.core.range_shorthand import (
    expand_whole_column_span_deps,
    expand_whole_row_span_deps,
)
from excel_grapher.core.types import ExcelRange, XlError
from excel_grapher.exporter.inverted_tree import excel as inverted_excel
from excel_grapher.exporter.inverted_tree.access import (
    indirect_argument_addresses,
    indirect_target_addresses,
    is_seed_access,
    offset_target_addresses,
    overlapping_schedule_peer,
    unique_seed_or_none,
)
from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    KeyPoint,
    SeriesCatalog,
    covering_series,
    covering_series_of_column,
    covering_series_of_range,
    fit_affine_map,
    preferred_fields,
    range_column_origin,
    schedule_axis_coord,
    schedule_coord,
    schedule_partition,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher.blank_ranges import (
    BlankRangeRect,
    address_in_blank_ranges,
    blank_rects_for_addresses,
)
from excel_grapher.semantic_model.types import AccessClass
from excel_grapher.series_bindings.geometry import parse_value_map
from excel_grapher.series_bindings.normalize import is_override_input
from excel_grapher.series_bindings.resolve import (
    _bind_source_addresses,
    _execute_bind,
    _structure_source_addresses,
    _WorkbookValues,
)

if TYPE_CHECKING:
    from excel_grapher.grapher.graph import DependencyGraph

_EMITTED_BUILTINS = frozenset(
    name.removeprefix("xl_").removesuffix("_lazy").upper()
    for name, value in vars(inverted_excel).items()
    if name.startswith("xl_") and callable(value)
)


def normalize_excel_function_name(name: str) -> str:
    """Normalize a formula token, stripping `_XLUDF.` for emitted builtins."""
    return _normalize_excel_function_name(name, registered_builtins=_EMITTED_BUILTINS)


_BLANK_RECTS: ContextVar[tuple[BlankRangeRect, ...]] = ContextVar(
    "excel_grapher_inverted_tree_blank_rects",
    default=(),
)


def bind_blank_rects(rects: tuple[BlankRangeRect, ...]) -> Token[tuple[BlankRangeRect, ...]]:
    """Install `rects` for the current inverted-tree emit walk."""
    return _BLANK_RECTS.set(rects)


def reset_blank_rects(token: Token[tuple[BlankRangeRect, ...]]) -> None:
    """Restore the blank-range context installed by `bind_blank_rects`."""
    _BLANK_RECTS.reset(token)


def current_blank_rects() -> tuple[BlankRangeRect, ...]:
    """Blank-range rectangles for the current inverted-tree emit walk."""
    return _BLANK_RECTS.get()


def addresses_outside_blank_ranges(
    addresses: Sequence[CanonicalAddress],
    blank_rects: Sequence[BlankRangeRect] | None = None,
) -> list[CanonicalAddress]:
    """Drop addresses that lie in declared structural blank rectangles."""
    rects = current_blank_rects() if blank_rects is None else blank_rects
    if not rects:
        return list(addresses)
    relevant = blank_rects_for_addresses(addresses, rects)
    if not relevant:
        return list(addresses)
    return [addr for addr in addresses if not address_in_blank_ranges(addr, relevant)]


@dataclass(frozen=True, slots=True)
class PositionalRangeCell:
    """One worksheet cell in a MATCH/INDEX window, in sheet order.

    `label_value` holds a row-label or header used when the cell sits
    outside the occupying series' `data_range`.
    """

    address: CanonicalAddress
    series_id: str | None
    catalog_index: int | None
    blank: bool
    label_value: object | None = None


def _is_excel_blank_cell(
    address: CanonicalAddress,
    graph: DependencyGraph | None,
) -> bool:
    """True when `address` is an Excel blank, not an unbound catalog cell.

    Range expansion may create graph leaves for empty cells. Those holes have
    no formula and no cached value. A formula cell or a valued leaf in the
    same window is a real catalog gap and must fail closed.
    """
    if graph is None:
        return False
    node = graph.get_node(address)
    if node is None:
        return True
    if node.has_formula:
        return False
    return node.value is None


def resolve_positional_range(
    addresses: Sequence[CanonicalAddress],
    catalog: SeriesCatalog,
    blank_rects: Sequence[BlankRangeRect] | None = None,
    graph: DependencyGraph | None = None,
) -> tuple[tuple[PositionalRangeCell, ...], tuple[CanonicalAddress, ...]]:
    """Map each address to a bound catalog cell or a positional blank.

    Returns `(cells, missing)`. `missing` lists addresses that are neither
    bound, declared blank, nor an empty off-catalog hole in a window that
    already has a bound cell. Worksheet order and rectangle size are
    preserved; blanks stay in `cells` so MATCH/INDEX positions do not
    shift. Unique cell ownership (`covering_series`) is unchanged.

    An entirely unbound window still fails closed so empty VLOOKUP tables
    require `blank_ranges`. On-graph formula cells and valued leaves
    without a series are always missing.
    """
    rects = current_blank_rects() if blank_rects is None else blank_rects
    relevant = blank_rects_for_addresses(addresses, rects) if rects else ()
    cells: list[PositionalRangeCell] = []
    missing: list[CanonicalAddress] = []
    unbound_blanks: list[CanonicalAddress] = []
    owned = False
    for address in addresses:
        if relevant and address_in_blank_ranges(address, relevant):
            cells.append(PositionalRangeCell(address, None, None, True))
            continue
        owner = catalog.series_for(address)
        if owner is not None:
            owned = True
            cells.append(
                PositionalRangeCell(address, owner.series_id, owner.index_of(address), False)
            )
            continue
        if _is_excel_blank_cell(address, graph):
            unbound_blanks.append(address)
            cells.append(PositionalRangeCell(address, None, None, True))
            continue
        missing.append(address)
    if unbound_blanks and not owned:
        missing.extend(unbound_blanks)
    return tuple(cells), tuple(missing)


def _label_values_for_window(
    catalog: SeriesCatalog, parent_addresses: Sequence[CanonicalAddress]
) -> dict[CanonicalAddress, object]:
    """Map row_label/column_header cells inside `parent_addresses` to series keys."""
    parent = set(parent_addresses)
    values: dict[CanonicalAddress, object] = {}
    for series in catalog.series.values():
        data_cells = series.authored_cells or series.cells
        occupies = [cell for cell in data_cells if cell in parent]
        if not occupies:
            continue
        for key_field in series.key_fields:
            bind = series.dimension_bind(key_field)
            if not isinstance(bind, Mapping) or bind.get("kind") not in {
                "row_label",
                "column_header",
            }:
                continue
            for data_cell in occupies:
                point = series.key_point_for(data_cell)
                if point is None:
                    continue
                for source in _bind_source_addresses(dict(bind), data_cell):
                    values[as_canonical(source)] = point[key_field]
    return values


def _attach_index_label_cells(
    selected: Sequence[CanonicalAddress],
    parent_addresses: Sequence[CanonicalAddress],
    cells: Sequence[PositionalRangeCell],
    missing: Sequence[CanonicalAddress],
    catalog: SeriesCatalog,
    graph: DependencyGraph | None,
) -> tuple[tuple[PositionalRangeCell, ...], tuple[CanonicalAddress, ...]]:
    """Fill INDEX column holes that are row-labels or the header leaf."""
    if not missing:
        return tuple(cells), tuple(missing)
    labels = _label_values_for_window(catalog, parent_addresses)
    by_addr = {cell.address: cell for cell in cells}
    header = selected[0] if selected else None
    filled: list[PositionalRangeCell] = []
    still: list[CanonicalAddress] = []
    for address in selected:
        if address in by_addr:
            filled.append(by_addr[address])
            continue
        if address in labels:
            filled.append(PositionalRangeCell(address, None, None, False, labels[address]))
            continue
        node = None if graph is None else graph.get_node(address)
        if address == header and node is not None and not node.has_formula:
            filled.append(PositionalRangeCell(address, None, None, False, node.value))
            continue
        still.append(address)
    return tuple(filled), tuple(still)


def index_column_literal(node: AstNode | None, host_cell: CanonicalAddress) -> int | None:
    """Return a static INDEX column (`1`, `COLUMNS(range)`), else `None`."""
    if node is None or isinstance(node, EmptyArgNode):
        return None
    literal = ast_literal_int(node)
    if literal is not None:
        return literal
    if not isinstance(node, FunctionCallNode):
        return None
    if normalize_excel_function_name(node.name) != "COLUMNS" or len(node.args) != 1:
        return None
    ref = node.args[0]
    if not isinstance(ref, RangeNode):
        return None
    start = as_canonical(resolve_cell_ref(ref.start_ref, host_cell))
    end = as_canonical(resolve_cell_ref(ref.end_ref, host_cell))
    if parse_cell_coords(start)[0] != parse_cell_coords(end)[0]:
        return None
    return abs(parse_cell_coords(end)[2] - parse_cell_coords(start)[2]) + 1


def range_ref_label(node: AstNode, host_cell: CanonicalAddress) -> str:
    """Return a sheet-qualified A1 label for a cell or range ref."""
    if isinstance(node, CellRefNode):
        return resolve_cell_ref(node, host_cell)
    if isinstance(node, RangeNode):
        start = resolve_cell_ref(node.start_ref, host_cell)
        end = resolve_cell_ref(node.end_ref, host_cell)
        return f"{start}:{end}"
    return type(node).__name__


def _layout_distance(
    consumer: CanonicalAddress, producer: CanonicalAddress, catalog: SeriesCatalog
) -> int:
    """Return schedule-axis distance (`axis(consumer) - axis(producer)`).

    Distance is the `TIME_PERIOD` component when both cells sit on a key
    nest (#638). Flattened full-tuple coordinates are not used here — a
    Kenya 2021 read of Kenya 2020 is distance 1, not the gap across France.
    """
    consumer_series = catalog.series_for(consumer)
    producer_series = catalog.series_for(producer)
    if (
        consumer_series is not None
        and producer_series is not None
        and preferred_fields(consumer_series, catalog) != preferred_fields(producer_series, catalog)
    ):
        consumer_index = consumer_series.index_of(consumer)
        producer_index = producer_series.index_of(producer)
        return (0 if consumer_index is None else consumer_index) - (
            0 if producer_index is None else producer_index
        )
    return schedule_axis_coord(consumer, catalog) - schedule_axis_coord(producer, catalog)


def identity_join_indices(
    host: BoundSeries,
    producer: BoundSeries,
    catalog: SeriesCatalog,
) -> tuple[int, ...]:
    """Return producer catalog slots for each host member on the schedule axis.

    A hole is `-1` when no producer member shares the host's schedule
    coordinate. The join uses the same distance as `DependenceEdge` (`0`
    means identity). The producer's coordinate -> index map is cached on
    the catalog `ScheduleIndex`.

    Raises:
        InvertedTreeExportError: Two producer members share one host
            coordinate, or a host cell has no schedule coordinate.
    """
    if preferred_fields(host, catalog) != preferred_fields(producer, catalog):
        return tuple(slot if slot < len(producer.cells) else -1 for slot in range(len(host.cells)))
    producer_by_coord = catalog.schedule.index_by_coord[producer.series_id]
    slots: list[int] = []
    for host_cell in host.cells:
        host_coord = catalog.schedule.coord_of.get(host_cell)
        if host_coord is None:
            raise InvertedTreeExportError(f"cell {host_cell} has no schedule coordinate")
        matches = producer_by_coord.get(host_coord, ())
        if len(matches) > 1:
            raise InvertedTreeExportError(
                f"identity join of {host.series_id!r} onto {producer.series_id!r} "
                f"is ambiguous at {host_cell}: duplicate schedule keys "
                f"{tuple(producer.cells[slot] for slot in matches)}"
            )
        slots.append(-1 if not matches else matches[0])
    return tuple(slots)


def _member_access(
    consumer_cell: CanonicalAddress,
    producer: BoundSeries,
    producer_cell: CanonicalAddress,
    catalog: SeriesCatalog,
) -> AccessClass:
    """Classify a cell-ref by schedule distance (`0` is identity).

    Instance labels are identity or shift. `refine_access_classes` upgrades a
    consumer-producer bundle to `affine` when catalog indices lie on
    `f(i) = a*i + b` with `a != 1`.
    """
    if producer.index_of(producer_cell) is None:
        return "gather"
    if schedule_partition(consumer_cell, catalog) != schedule_partition(producer_cell, catalog):
        return "cross_partition"
    if _layout_distance(consumer_cell, producer_cell, catalog) == 0:
        return "identity"
    return "shift"


def iter_range_addresses(start: str, end: str) -> list[CanonicalAddress]:
    """Expand a same-sheet A1 range into canonical cell addresses (row-major)."""
    sheet1, row1, col1 = parse_cell_coords(start)
    sheet2, row2, col2 = parse_cell_coords(end)
    if sheet1 != sheet2:
        raise InvertedTreeExportError(f"cross-sheet range {start}:{end} is not supported")
    r1, r2 = min(row1, row2), max(row1, row2)
    c1, c2 = min(col1, col2), max(col1, col2)
    return [
        as_canonical(format_cell_key(sheet1, get_column_letter(col), row))
        for row in range(r1, r2 + 1)
        for col in range(c1, c2 + 1)
    ]


def iter_cross_sheet_addresses(
    start: str,
    end: str,
    graph: DependencyGraph,
) -> list[CanonicalAddress]:
    """Expand a 3-D range across workbook `sheet_order` (row-major per sheet)."""
    sheet1, row1, col1 = parse_cell_coords(start)
    sheet2, row2, col2 = parse_cell_coords(end)
    order = graph.sheet_order
    if not order:
        raise InvertedTreeExportError(f"cross-sheet range {start}:{end} is not supported")
    try:
        first = order.index(sheet1)
        last = order.index(sheet2)
    except ValueError as exc:
        raise InvertedTreeExportError(f"cross-sheet range {start}:{end} is not supported") from exc
    lo, hi = (first, last) if first <= last else (last, first)
    r1, r2 = min(row1, row2), max(row1, row2)
    c1, c2 = min(col1, col2), max(col1, col2)
    return [
        as_canonical(format_cell_key(sheet, get_column_letter(col), row))
        for sheet in order[lo : hi + 1]
        for row in range(r1, r2 + 1)
        for col in range(c1, c2 + 1)
    ]


def iter_ref_addresses(
    node: AstNode,
    host_cell: CanonicalAddress,
    graph: DependencyGraph | None,
) -> list[CanonicalAddress]:
    """Expand a range, whole-column, or whole-row leaf to canonical cells.

    Same-sheet rectangles use `iter_range_addresses`. Cross-sheet ranges walk
    `graph.sheet_order` between the endpoints. Whole-column / whole-row refs
    expand to the workbook used extent in `graph.sheet_bounds`.
    """
    if isinstance(node, RangeNode):
        start = resolve_cell_ref(node.start_ref, host_cell)
        end = resolve_cell_ref(node.end_ref, host_cell)
        sheet1, _row1, _col1 = parse_cell_coords(start)
        sheet2, _row2, _col2 = parse_cell_coords(end)
        if sheet1 == sheet2:
            return iter_range_addresses(start, end)
        if graph is None:
            raise InvertedTreeExportError(f"cross-sheet range {start}:{end} is not supported")
        return iter_cross_sheet_addresses(start, end, graph)
    if isinstance(node, WholeColumnNode):
        if graph is None or graph.sheet_bounds is None:
            raise InvertedTreeExportError("whole-column ref requires workbook used-range bounds")
        sheet, start_letter, end_letter = resolve_whole_column_ref(node, host_cell)
        return [
            as_canonical(format_key(dep_sheet, a1))
            for dep_sheet, a1 in expand_whole_column_span_deps(
                sheet, start_letter, end_letter, graph.sheet_bounds
            )
        ]
    if isinstance(node, WholeRowNode):
        if graph is None or graph.sheet_bounds is None:
            raise InvertedTreeExportError("whole-row ref requires workbook used-range bounds")
        sheet, start_row, end_row = resolve_whole_row_ref(node, host_cell)
        return [
            as_canonical(format_key(dep_sheet, a1))
            for dep_sheet, a1 in expand_whole_row_span_deps(
                sheet, start_row, end_row, graph.sheet_bounds
            )
        ]
    raise InvertedTreeExportError(
        f"expected a range, whole-column, or whole-row ref, got {type(node).__name__}"
    )


def _ref_only_argument_is_address_shaped(node: AstNode) -> bool:
    """True when a ref-only argument has no nested function call.

    Matches `ref_only_function_spans`: `ROW($B$1)` and `ROWS(A1:A2)` use only
    address or shape, so they are not value dependencies. A nested call such
    as `ROW(INDEX(...))` stays visible.
    """
    match node:
        case FunctionCallNode():
            return False
        case BinaryOpNode():
            return _ref_only_argument_is_address_shaped(
                node.left
            ) and _ref_only_argument_is_address_shaped(node.right)
        case UnaryOpNode():
            return _ref_only_argument_is_address_shaped(node.operand)
        case _:
            return True


def ast_literal_int(node: AstNode) -> int | None:
    """Return an integer encoded as `NumberNode` or unary `+` / `-`."""
    if isinstance(node, NumberNode):
        value = node.value
        if isinstance(value, bool) or value != int(value):
            return None
        return int(value)
    if isinstance(node, UnaryOpNode) and node.op in {"+", "-"}:
        inner = ast_literal_int(node.operand)
        if inner is None:
            return None
        return inner if node.op == "+" else -inner
    return None


def eval_host_selector(node: AstNode, host_cell: CanonicalAddress) -> int | None:
    """Evaluate a ROW/COLUMN/arithmetic INDEX selector at `host_cell`.

    Returns `None` when the selector is not a closed integer form of literals,
    `ROW` / `COLUMN` of the host or of a cell, and `+` `-` `*` `/`.
    """
    literal = ast_literal_int(node)
    if literal is not None:
        return literal
    if isinstance(node, UnaryOpNode) and node.op in {"+", "-"}:
        inner = eval_host_selector(node.operand, host_cell)
        if inner is None:
            return None
        return inner if node.op == "+" else -inner
    if isinstance(node, BinaryOpNode) and node.op in {"+", "-", "*", "/"}:
        left = eval_host_selector(node.left, host_cell)
        right = eval_host_selector(node.right, host_cell)
        if left is None or right is None:
            return None
        if node.op == "+":
            return left + right
        if node.op == "-":
            return left - right
        if node.op == "*":
            return left * right
        if right == 0:
            return None
        return int(left / right)
    if not isinstance(node, FunctionCallNode):
        return None
    name = normalize_excel_function_name(node.name)
    if name not in {"ROW", "COLUMN"}:
        return None
    if not node.args or (len(node.args) == 1 and isinstance(node.args[0], EmptyArgNode)):
        _sheet, row, col = parse_cell_coords(host_cell)
        return row if name == "ROW" else col
    if len(node.args) != 1:
        return None
    arg = node.args[0]
    corners = None
    if isinstance(arg, CellRefNode):
        address = as_canonical(resolve_cell_ref(arg, host_cell))
        corners = address, address
    elif isinstance(arg, RangeNode):
        corners = (
            as_canonical(resolve_cell_ref(arg.start_ref, host_cell)),
            as_canonical(resolve_cell_ref(arg.end_ref, host_cell)),
        )
    if corners is None:
        return None
    _sheet, row, col = parse_cell_coords(corners[0])
    return row if name == "ROW" else col


def ref_window_corners(
    node: AstNode, host_cell: CanonicalAddress
) -> tuple[CanonicalAddress, CanonicalAddress] | None:
    """Return `(start, end)` for a cell or range reference."""
    if isinstance(node, CellRefNode):
        address = as_canonical(resolve_cell_ref(node, host_cell))
        return address, address
    if isinstance(node, RangeNode):
        start = as_canonical(resolve_cell_ref(node.start_ref, host_cell))
        end = as_canonical(resolve_cell_ref(node.end_ref, host_cell))
        return start, end
    return None


def index_window_corners(
    node: FunctionCallNode, host_cell: CanonicalAddress
) -> tuple[CanonicalAddress, CanonicalAddress] | None:
    """Return the INDEX array window, narrowed to a literal column when present."""
    if normalize_excel_function_name(node.name) != "INDEX" or len(node.args) < 2:
        return None
    corners = ref_window_corners(node.args[0], host_cell)
    if corners is None:
        return None
    start, end = corners
    col_arg = node.args[2] if len(node.args) > 2 else None
    if col_arg is None:
        return start, end
    col_index = ast_literal_int(col_arg)
    if col_index is None:
        return start, end
    sheet1, row1, col1 = parse_cell_coords(start)
    sheet2, row2, col2 = parse_cell_coords(end)
    if sheet1 != sheet2:
        return None
    r1, r2 = min(row1, row2), max(row1, row2)
    c1, c2 = min(col1, col2), max(col1, col2)
    width = c2 - c1 + 1
    if col_index < 1 or col_index > width:
        return None
    col = c1 + col_index - 1
    return (
        as_canonical(format_cell_key(sheet1, get_column_letter(col), r1)),
        as_canonical(format_cell_key(sheet1, get_column_letter(col), r2)),
    )


def index_call_is_ref(node: FunctionCallNode, host_cell: CanonicalAddress) -> bool:
    """Return whether `INDEX` is out of bounds at `host_cell`.

    Unevaluable selectors return `False` so export stays fail-closed until
    runtime `INDEX` can raise `#REF!`.
    """
    if normalize_excel_function_name(node.name) != "INDEX" or len(node.args) < 2:
        return False
    corners = ref_window_corners(node.args[0], host_cell)
    if corners is None:
        return False
    row_sel = eval_host_selector(node.args[1], host_cell)
    if row_sel is None:
        return False
    col_arg = node.args[2] if len(node.args) > 2 else None
    if col_arg is None:
        col_sel: int | None = None
    elif isinstance(col_arg, EmptyArgNode):
        col_sel = 0
    else:
        col_sel = eval_host_selector(col_arg, host_cell)
        if col_sel is None:
            return False
    sheet, row1, col1 = parse_cell_coords(corners[0])
    _sheet2, row2, col2 = parse_cell_coords(corners[1])
    if sheet != _sheet2:
        return False
    base = ExcelRange(
        sheet,
        min(row1, row2),
        min(col1, col2),
        max(row1, row2),
        max(col1, col2),
    )
    result = index_excel_range(base, row_sel, col_sel)
    return isinstance(result, XlError) and result == XlError.REF


def shift_range_corners(
    start: CanonicalAddress,
    end: CanonicalAddress,
    rows: int,
    cols: int,
) -> tuple[CanonicalAddress, CanonicalAddress] | None:
    """Translate a same-sheet window by `rows` and `cols`, or `None` if off-grid."""
    sheet1, row1, col1 = parse_cell_coords(start)
    sheet2, row2, col2 = parse_cell_coords(end)
    if sheet1 != sheet2:
        return None
    new_r1 = row1 + rows
    new_c1 = col1 + cols
    new_r2 = row2 + rows
    new_c2 = col2 + cols
    if min(new_r1, new_r2) < 1 or min(new_c1, new_c2) < 1:
        return None
    if max(new_r1, new_r2) > 1_048_576 or max(new_c1, new_c2) > 16_384:
        return None
    return (
        as_canonical(format_cell_key(sheet1, get_column_letter(new_c1), new_r1)),
        as_canonical(format_cell_key(sheet1, get_column_letter(new_c2), new_r2)),
    )


def offset_cell_destination(
    node: FunctionCallNode, host_cell: CanonicalAddress
) -> CanonicalAddress | None:
    """Return the landing cell of a literal `OFFSET(cell, rows, cols)`.

    Extra height/width arguments make a range `OFFSET`; those stay on the
    `OFFSET(INDEX(...))` path. Off-grid moves return `None`.
    """
    if normalize_excel_function_name(node.name) != "OFFSET" or len(node.args) < 3:
        return None
    base = node.args[0]
    if not isinstance(base, CellRefNode) or len(node.args) >= 4:
        return None
    rows = ast_literal_int(node.args[1])
    cols = ast_literal_int(node.args[2])
    if rows is None or cols is None:
        return None
    anchor = as_canonical(resolve_cell_ref(base, host_cell))
    dest = shift_range_corners(anchor, anchor, rows, cols)
    return dest[0] if dest is not None else None


def offset_index_destination(
    node: FunctionCallNode, host_cell: CanonicalAddress
) -> tuple[CanonicalAddress, CanonicalAddress] | None:
    """Return dest corners of `OFFSET(INDEX(...), rows, cols)` when static.

    Classifies the INDEX window the same way as `_visit_index`, then applies
    literal row/column offsets. Non-literal offsets return `None`.
    """
    if normalize_excel_function_name(node.name) != "OFFSET" or len(node.args) < 3:
        return None
    base = node.args[0]
    if not isinstance(base, FunctionCallNode):
        return None
    if normalize_excel_function_name(base.name) != "INDEX":
        return None
    rows = ast_literal_int(node.args[1])
    cols = ast_literal_int(node.args[2])
    if rows is None or cols is None:
        return None
    if len(node.args) >= 4 and ast_literal_int(node.args[3]) is None:
        return None
    if len(node.args) >= 5 and ast_literal_int(node.args[4]) is None:
        return None
    window = index_window_corners(base, host_cell)
    if window is None:
        return None
    dest = shift_range_corners(window[0], window[1], rows, cols)
    if dest is None:
        return None
    start, end = dest
    if len(node.args) >= 4:
        height = ast_literal_int(node.args[3])
        width = ast_literal_int(node.args[4]) if len(node.args) >= 5 else None
        if height is None or height < 1:
            return None
        sheet, row, col = parse_cell_coords(start)
        end_row = row + height - 1
        end_col = (
            col + (width - 1) if width is not None and width >= 1 else parse_cell_coords(end)[2]
        )
        if end_row < 1 or end_col < 1 or end_row > 1_048_576 or end_col > 16_384:
            return None
        end = as_canonical(format_cell_key(sheet, get_column_letter(end_col), end_row))
    return start, end


def offset_expr_exclude_addresses(
    node: FunctionCallNode,
    host_cell: CanonicalAddress,
    graph: DependencyGraph | None,
) -> list[CanonicalAddress]:
    """Return INDEX-array and OFFSET-argument cells that are not the destination."""
    found: list[CanonicalAddress] = []
    if node.args:
        base = node.args[0]
        if isinstance(base, FunctionCallNode) and (
            normalize_excel_function_name(base.name) == "INDEX" and base.args
        ):
            array = base.args[0]
            corners = ref_window_corners(array, host_cell)
            if corners is not None:
                found.extend(iter_range_addresses(corners[0], corners[1]))
            elif isinstance(array, (RangeNode, WholeColumnNode, WholeRowNode)):
                found.extend(iter_ref_addresses(array, host_cell, graph))
        for arg in node.args[1:]:
            if isinstance(arg, CellRefNode):
                found.append(as_canonical(resolve_cell_ref(arg, host_cell)))
    return found


def resolve_offset_destination_series(
    node: FunctionCallNode,
    host_cell: CanonicalAddress,
    catalog: SeriesCatalog,
    graph: DependencyGraph | None,
    *,
    blank_rects: Sequence[BlankRangeRect] | None = None,
) -> tuple[BoundSeries, CanonicalAddress] | None:
    """Return `(series, anchor)` for OFFSET whose base yields a reference.

    Prefers a literal `OFFSET(cell, rows, cols)` landing cell, then a
    classified `OFFSET(INDEX(...), rows, cols)` window. Falls back to graph
    `dynamic_offset` edges when the offset is not a literal.
    """
    cell_dest = offset_cell_destination(node, host_cell)
    if cell_dest is not None:
        addresses = addresses_outside_blank_ranges([cell_dest], blank_rects)
        covered = covering_series(catalog, addresses) if addresses else None
        if covered is None:
            return None
        return covered, cell_dest
    dest = offset_index_destination(node, host_cell)
    if dest is not None:
        addresses = addresses_outside_blank_ranges(
            iter_range_addresses(dest[0], dest[1]),
            blank_rects,
        )
        covered = covering_series(catalog, addresses) if addresses else None
        if covered is None:
            return None
        return covered, dest[0]
    if graph is None:
        return None
    exclude = offset_expr_exclude_addresses(node, host_cell, graph)
    targets = addresses_outside_blank_ranges(
        offset_target_addresses(graph, host_cell, exclude=exclude),
        blank_rects,
    )
    covered = covering_series(catalog, targets) if targets else None
    if covered is None:
        return None
    return covered, targets[0]


_OFFSET_COLUMN_READER: ContextVar[_WorkbookValues | None] = ContextVar(
    "excel_grapher_offset_column_reader",
    default=None,
)


@contextmanager
def offset_column_workbook(workbook: Path | str | None) -> Iterator[None]:
    """Open `workbook` for anchor column-key reads during one emit.

    A column-only `OFFSET` whose anchor series has no column axis reads the
    anchor column with the landing series' column bind. The reader stays open
    for the whole emit so each formula does not reload the workbook.
    """
    if workbook is None:
        yield
        return
    reader = _WorkbookValues(workbook)
    token = _OFFSET_COLUMN_READER.set(reader)
    try:
        yield
    finally:
        _OFFSET_COLUMN_READER.reset(token)
        reader.close()


@dataclass(frozen=True, slots=True)
class OffsetColumnMember:
    """One worksheet column of a column-only `OFFSET` span."""

    address: CanonicalAddress
    series: BoundSeries
    column_key: object


@dataclass(frozen=True, slots=True)
class OffsetColumnSpan:
    """Contiguous columns from the anchor through every in-domain landing.

    `members` follows worksheet order. `anchor_index` is the anchor column.
    `column_field` is the key the landing series enumerates. The anchor column
    carries that same key even when its own series does not declare the field,
    so a later `axis_step` counts worksheet columns from the anchor. Columns
    between the anchor and the farthest landing stay on the span when the
    selector domain skips them: the displacement is a worksheet distance.
    """

    column_field: str
    members: tuple[OffsetColumnMember, ...]
    anchor_index: int

    @property
    def anchor_key(self) -> object:
        """Column key of the anchor column, the origin of `axis_step`."""
        return self.members[self.anchor_index].column_key


def _column_key_fields(series: BoundSeries) -> tuple[str, ...]:
    return tuple(field for field in series.key_fields if _key_field_axis(series, field) == "col")


def _offset_column_error(
    host_series_id: str | None,
    host_cell: CanonicalAddress,
    message: str,
) -> InvertedTreeExportError:
    if host_series_id is None:
        return InvertedTreeExportError(f"cell {host_cell}: {message}")
    return InvertedTreeExportError(f"series {host_series_id!r} cell {host_cell}: {message}")


def _offset_argument_cells(
    node: FunctionCallNode,
    host_cell: CanonicalAddress,
) -> set[CanonicalAddress]:
    found: set[CanonicalAddress] = set()
    for arg in node.args[1:]:
        for ref in _iter_cell_ref_nodes(arg):
            found.add(as_canonical(resolve_cell_ref(ref, host_cell)))
    return found


def _offset_provenance_captured(
    graph: DependencyGraph,
    host_cell: CanonicalAddress,
) -> bool:
    for dep in graph.get_dependencies(host_cell):
        if graph.get_edge_attrs(host_cell, dep).provenance is not None:
            return True
    return False


def _infer_offset_landings(
    graph: DependencyGraph,
    host_cell: CanonicalAddress,
    *,
    blank_rects: Sequence[BlankRangeRect] | None,
) -> list[CanonicalAddress] | None:
    """Return inferred `OFFSET` targets, or `None` when the domain is unknown."""
    node = graph.get_node(host_cell)
    formula = None if node is None else node.normalized_formula
    env = graph.cell_type_env
    if not formula or env is None:
        return None
    from excel_grapher.grapher.dynamic_refs import DynamicRefError, infer_dynamic_offset_targets

    sheet = parse_cell_coords(host_cell)[0]
    try:
        found = infer_dynamic_offset_targets(
            formula,
            current_sheet=sheet,
            cell_type_env=env,
            blank_rects=blank_rects,
        )
    except DynamicRefError:
        return None
    return [as_canonical(address) for address in found]


def _same_row_landings(
    anchor: CanonicalAddress,
    targets: Sequence[CanonicalAddress],
    excluded: set[CanonicalAddress],
) -> list[CanonicalAddress]:
    sheet, row, _col = parse_cell_coords(anchor)
    landings: list[CanonicalAddress] = []
    for target in targets:
        if target in excluded:
            continue
        target_sheet, target_row, _target_col = parse_cell_coords(target)
        if target_sheet != sheet or target_row != row:
            continue
        landings.append(as_canonical(target))
    return landings


def _in_domain_column_landings(
    node: FunctionCallNode,
    host_cell: CanonicalAddress,
    anchor: CanonicalAddress,
    graph: DependencyGraph | None,
    *,
    blank_rects: Sequence[BlankRangeRect] | None,
) -> list[CanonicalAddress] | None:
    """Return same-row in-domain landings, or `None` when they are unknown."""
    if graph is None:
        return None
    excluded = _offset_argument_cells(node, host_cell)
    if _offset_provenance_captured(graph, host_cell):
        targets: list[CanonicalAddress] | None = offset_target_addresses(
            graph, host_cell, exclude=tuple(excluded)
        )
    else:
        targets = _infer_offset_landings(graph, host_cell, blank_rects=blank_rects)
    if targets is None:
        return None
    return _same_row_landings(anchor, targets, excluded)


class _UnusedColumnReader:
    """Stand-in reader for binds that do not touch the workbook."""

    def read(self, address: str) -> Any:
        raise InvertedTreeExportError(f"OFFSET column key at {address} needs the bindings workbook")


def _bind_column_key(
    bind: Mapping[str, Any],
    address: CanonicalAddress,
    graph: DependencyGraph | None,
    *,
    read_as: str,
    evaluate_addresses: set[str],
    evaluators: dict[int, Any],
) -> object:
    """Return the column key `bind` assigns to `address`.

    `read_as` is the landing series' effective key type. A bind that omits
    `read` would otherwise coerce with `auto` and disagree with catalog keys.
    `evaluate_addresses` is the catalog's labeller set, so an uncached formula
    header is evaluated the same way as series resolution.
    """
    applied = dict(bind)
    if read_as and "read" not in applied:
        applied["read"] = read_as
    reader = _OFFSET_COLUMN_READER.get()
    sources = _bind_source_addresses(applied, str(address))
    if reader is None:
        if sources:
            raise ValueError(f"OFFSET column key at {address} needs the bindings workbook")
        reader = cast(_WorkbookValues, _UnusedColumnReader())
    return _execute_bind(
        applied,
        graph=graph,
        reader=reader,
        data_address=str(address),
        evaluate_addresses=evaluate_addresses,
        evaluators=evaluators,
    )


def _column_key_value(key: object) -> str | int:
    if isinstance(key, bool) or not isinstance(key, str | int):
        raise ValueError(f"OFFSET column key {key!r} must be a string or integer")
    return key


@dataclass(frozen=True, slots=True)
class _OffsetColumnLayout:
    """Bound columns of a cross-series column `OFFSET`, before key reads."""

    anchor_series_id: str
    anchor_col: int
    field_name: str
    bind: Mapping[str, Any]
    read_as: str
    columns: tuple[tuple[int, CanonicalAddress, BoundSeries], ...]


def _key_read_as(series: BoundSeries, field: str) -> str:
    """Return the catalog read mode for `field` on `series`."""
    try:
        index = series.key_fields.index(field)
    except ValueError:
        return "auto"
    if index < len(series.key_types) and series.key_types[index]:
        return series.key_types[index]
    bind = series.dimension_bind(field) or {}
    return str(bind.get("read", "auto"))


def _labeller_evaluate_addresses(catalog: SeriesCatalog) -> set[str]:
    """Return labeller cells and header sources the catalog may evaluate."""
    addresses: set[str] = set()
    for series in catalog.series.values():
        if not series.axis_labels:
            continue
        cells = [
            str(cell)
            for cell in (
                series.authored_cells if series.authored_cells is not None else series.cells
            )
        ]
        addresses.update(cells)
        addresses.update(_structure_source_addresses(dict(series.raw), cells))
    return addresses


@dataclass(frozen=True, slots=True)
class _DynamicColumnMove:
    """A dynamic `OFFSET(cell, 0, cols)` whose anchor has no column axis."""

    anchor: CanonicalAddress
    anchor_series: BoundSeries
    sheet: str
    row: int
    anchor_col: int
    landings: tuple[CanonicalAddress, ...]
    rects: tuple[BlankRangeRect, ...]


def _dynamic_column_move(
    node: FunctionCallNode,
    host_cell: CanonicalAddress,
    catalog: SeriesCatalog,
    graph: DependencyGraph | None,
    *,
    blank_rects: Sequence[BlankRangeRect] | None = None,
) -> _DynamicColumnMove | None:
    """Return a column-only dynamic `OFFSET`, or `None` when it is some other call."""
    if len(node.args) != 3 or not isinstance(node.args[0], CellRefNode):
        return None
    if ast_literal_int(node.args[1]) != 0 or ast_literal_int(node.args[2]) is not None:
        return None
    anchor = as_canonical(resolve_cell_ref(node.args[0], host_cell))
    anchor_series = catalog.series_for(anchor)
    if anchor_series is None or _column_key_fields(anchor_series):
        return None
    rects = current_blank_rects() if blank_rects is None else tuple(blank_rects)
    landings = _in_domain_column_landings(node, host_cell, anchor, graph, blank_rects=rects)
    if not landings:
        return None
    sheet, row, anchor_col = parse_cell_coords(anchor)
    return _DynamicColumnMove(
        anchor,
        anchor_series,
        sheet,
        row,
        anchor_col,
        tuple(landings),
        rects,
    )


def anchor_only_column_target(
    node: FunctionCallNode,
    host_cell: CanonicalAddress,
    catalog: SeriesCatalog,
    graph: DependencyGraph | None,
    *,
    blank_rects: Sequence[BlankRangeRect] | None = None,
) -> tuple[BoundSeries, CanonicalAddress] | None:
    """Return the anchor when every in-domain landing stays on its column.

    A language domain of `{0}` never leaves the anchor. The caller reads that
    cell. A domain that reaches another column is a span, not this case.
    """
    move = _dynamic_column_move(
        node,
        host_cell,
        catalog,
        graph,
        blank_rects=blank_rects,
    )
    if move is None:
        return None
    columns = {parse_cell_coords(landing)[2] for landing in move.landings}
    if columns != {move.anchor_col}:
        return None
    return move.anchor_series, move.anchor


def _offset_column_layout(
    node: FunctionCallNode,
    host_cell: CanonicalAddress,
    catalog: SeriesCatalog,
    graph: DependencyGraph | None,
    *,
    blank_rects: Sequence[BlankRangeRect] | None = None,
    host_series_id: str | None = None,
) -> _OffsetColumnLayout | None:
    """Return bound columns for a dynamic column-only `OFFSET`, or `None`.

    `None` means the formula is not a cross-series column step, or every
    in-domain landing stays on the anchor column. A column step that reaches
    another column fails closed when a worksheet column in the span is unbound,
    the landing series do not share one column key, or that key has no bind.
    """
    move = _dynamic_column_move(
        node,
        host_cell,
        catalog,
        graph,
        blank_rects=blank_rects,
    )
    if move is None:
        return None
    landing_cols = [parse_cell_coords(landing)[2] for landing in move.landings]
    first = min(move.anchor_col, min(landing_cols))
    last = max(move.anchor_col, max(landing_cols))
    if first == last:
        return None
    columns: list[tuple[int, CanonicalAddress, BoundSeries]] = []
    for col in range(first, last + 1):
        address = as_canonical(format_cell_key(move.sheet, get_column_letter(col), move.row))
        if address_in_blank_ranges(address, move.rects) or catalog.series_for(address) is None:
            raise _offset_column_error(
                host_series_id,
                host_cell,
                f"OFFSET column offset into series {move.anchor_series.series_id!r} "
                f"crosses unbound cell {address}",
            )
        columns.append((col, address, catalog.require_series_for(address)))
    field_name: str | None = None
    bind: Mapping[str, Any] | None = None
    read_as: str | None = None
    for col, _address, series in columns:
        if col == move.anchor_col:
            continue
        fields = _column_key_fields(series)
        if len(fields) != 1:
            if not fields:
                continue
            raise _offset_column_error(
                host_series_id,
                host_cell,
                f"OFFSET column offset into series {move.anchor_series.series_id!r} "
                f"uses more than one column key on {series.series_id!r}",
            )
        name = fields[0]
        series_bind = series.dimension_bind(name) or {}
        series_read = _key_read_as(series, name)
        if field_name is None:
            field_name = name
            bind = series_bind
            read_as = series_read
            continue
        if name != field_name or dict(series_bind) != dict(bind or {}) or series_read != read_as:
            raise _offset_column_error(
                host_series_id,
                host_cell,
                f"OFFSET column offset into series {move.anchor_series.series_id!r} "
                "column key bind differs across the landing series",
            )
    if field_name is None or bind is None or read_as is None:
        raise _offset_column_error(
            host_series_id,
            host_cell,
            f"OFFSET column offset into series {move.anchor_series.series_id!r} "
            "has no column key on the landing cells",
        )
    return _OffsetColumnLayout(
        move.anchor_series.series_id,
        move.anchor_col,
        field_name,
        dict(bind),
        read_as,
        tuple(columns),
    )


def resolve_offset_column_span(
    node: FunctionCallNode,
    host_cell: CanonicalAddress,
    catalog: SeriesCatalog,
    graph: DependencyGraph | None,
    *,
    blank_rects: Sequence[BlankRangeRect] | None = None,
    host_series_id: str | None = None,
) -> OffsetColumnSpan | None:
    """Return a cross-series column span for `OFFSET(cell, 0, cols)`.

    The anchor series must not already enumerate a column axis. `cols` is not
    a literal. In-domain landings share one column key. `None` means this
    pattern does not apply; the caller keeps its existing failure.

    The anchor column uses the landing series' column bind, so the key
    sequence starts at the anchor even when that series has no column field.

    Raises:
        InvertedTreeExportError: The move reaches another column, and a
            worksheet column in the span is unbound, the landing series do not
            share one column key, a column key is missing, or two columns
            share one key.
    """
    layout = _offset_column_layout(
        node,
        host_cell,
        catalog,
        graph,
        blank_rects=blank_rects,
        host_series_id=host_series_id,
    )
    if layout is None:
        return None
    members: list[OffsetColumnMember] = []
    seen: dict[object, CanonicalAddress] = {}
    anchor_index: int | None = None
    labellers = _labeller_evaluate_addresses(catalog)
    evaluators: dict[int, Any] = {}
    for col, address, series in layout.columns:
        point = series.key_point_for(address)
        try:
            if layout.field_name in series.key_fields and (
                _key_field_axis(series, layout.field_name) == "col"
            ):
                if point is None or layout.field_name not in point.as_mapping():
                    raise ValueError(
                        f"OFFSET column offset into series {layout.anchor_series_id!r} "
                        f"has no {layout.field_name!r} key at {address}"
                    )
                raw_key = point[layout.field_name]
            else:
                raw_key = _bind_column_key(
                    layout.bind,
                    address,
                    graph,
                    read_as=layout.read_as,
                    evaluate_addresses=labellers,
                    evaluators=evaluators,
                )
            key = _column_key_value(raw_key)
        except ValueError as exc:
            raise _offset_column_error(host_series_id, host_cell, str(exc)) from exc
        previous = seen.get(key)
        if previous is not None:
            raise _offset_column_error(
                host_series_id,
                host_cell,
                f"OFFSET column key {key!r} is shared by {previous} and {address}",
            )
        seen[key] = address
        if col == layout.anchor_col:
            anchor_index = len(members)
        members.append(OffsetColumnMember(address, series, key))
    if anchor_index is None:
        return None
    return OffsetColumnSpan(layout.field_name, tuple(members), anchor_index)


def covering_series_for_index_window(
    node: FunctionCallNode,
    host_cell: CanonicalAddress,
    catalog: SeriesCatalog,
    *,
    blank_rects: Sequence[BlankRangeRect] | None = None,
) -> BoundSeries | None:
    """Return the unique series that owns an INDEX array window."""
    window = index_window_corners(node, host_cell)
    if window is None:
        return None
    addresses = addresses_outside_blank_ranges(
        iter_range_addresses(window[0], window[1]),
        blank_rects,
    )
    return covering_series(catalog, addresses) if addresses else None


def try_formula_ast(graph: DependencyGraph, address: CanonicalAddress) -> AstNode | None:
    """Return the formula AST for `address`, or `None` when the cell is a hole."""
    node = graph.get_node(address)
    if node is None:
        return None
    return getattr(node, "formula_ast", None)


def node_formula_ast(graph: DependencyGraph, address: CanonicalAddress) -> AstNode:
    """Return the formula AST for `address`, or fail closed."""
    ast = try_formula_ast(graph, address)
    if ast is None:
        node = graph.get_node(address)
        if node is None:
            raise InvertedTreeExportError(f"graph is missing bound cell {address}")
        raise InvertedTreeExportError(
            f"bound cell {address} has no formula AST (cannot verify first-level refs)"
        )
    return ast


def _iter_cell_ref_nodes(node: AstNode) -> Iterator[CellRefNode]:
    """Yield every `CellRefNode` in `node` (not range endpoints)."""
    match node:
        case CellRefNode():
            yield node
        case BinaryOpNode(left=left, right=right):
            yield from _iter_cell_ref_nodes(left)
            yield from _iter_cell_ref_nodes(right)
        case UnaryOpNode(operand=operand):
            yield from _iter_cell_ref_nodes(operand)
        case FunctionCallNode(args=args):
            for arg in args:
                yield from _iter_cell_ref_nodes(arg)
        case _:
            return


def _ref_shifts_with_host(ref: CellRefNode, host_cell: str) -> bool:
    """True when `ref` moves with `host_cell` on every axis that differs.

    An absolute axis (`$A$2`, `A$2`, `$A2` on the axis that changes) is a
    scalar read every period and cannot be a lag.
    """
    address = resolve_cell_ref(ref, host_cell)
    _host_sheet, host_row, host_col = parse_cell_coords(host_cell)
    _addr_sheet, addr_row, addr_col = parse_cell_coords(address)
    cell = ref.ref
    col_locked = addr_col != host_col and isinstance(cell.col, AbsoluteAxis)
    row_locked = addr_row != host_row and isinstance(cell.row, AbsoluteAxis)
    return not col_locked and not row_locked


def _has_schedule_axis(series: BoundSeries, catalog: SeriesCatalog) -> bool:
    """True when `series` declares `TIME_PERIOD` as a key field."""
    fields = preferred_fields(series, catalog)
    return fields is not None and "TIME_PERIOD" in fields


def _boundary_index(series: BoundSeries, catalog: SeriesCatalog, *, delta: int) -> int:
    """Return the catalog index of the min (`delta < 0`) or max schedule coord."""
    key = min if delta < 0 else max
    return key(range(len(series.cells)), key=lambda i: schedule_coord(series.cells[i], catalog))


def _non_peer_seed_ref(
    ast: AstNode,
    host_cell: CanonicalAddress,
    series: BoundSeries,
    catalog: SeriesCatalog,
) -> CanonicalAddress | None:
    """Return the unique relative year-0 read of a series outside the host key.

    A unique relative read of a differently keyed series is a seed only when
    it is not the same `TIME_PERIOD` as `host_cell`. Copying one matrix row
    (`B4=B2`, `C4=C2`) is an identity join on the inner axis, not a scan of
    the whole matrix sequence (#649). A `T+1` take of a richer-keyed
    overlapping peer is a lagged or cross-partition read, not a seed (#747).
    A scalar or unkeyed neighbor has no schedule axis and remains a year-0
    seed.
    """
    host_fields = preferred_fields(series, catalog)
    found: list[CanonicalAddress] = []
    for ref in _iter_cell_ref_nodes(ast):
        if not _ref_shifts_with_host(ref, host_cell):
            continue
        address = as_canonical(resolve_cell_ref(ref, host_cell))
        owner = catalog.series_for(address)
        if owner is None or owner.series_id == series.series_id:
            continue
        if overlapping_schedule_peer(series, owner, catalog):
            continue
        if preferred_fields(owner, catalog) == host_fields:
            continue
        if (
            _has_schedule_axis(series, catalog)
            and _has_schedule_axis(owner, catalog)
            and schedule_axis_coord(address, catalog) == schedule_axis_coord(host_cell, catalog)
        ):
            continue
        found.append(address)
    if len(found) == 1:
        return found[0]
    return None


def _adjacent_schedule_ref(
    series: BoundSeries,
    index: int,
    catalog: SeriesCatalog,
    graph: DependencyGraph | None,
    *,
    delta: int,
) -> CanonicalAddress | None:
    """Return a relative other-series ref at `schedule_coord` + `delta`.

    When the host and producer do not share join fields, a unique relative
    read of that producer at the series' schedule boundary is the seed or
    terminal. An overlapping keyed peer is a lagged read, not a seed,
    including a producer whose fields are a superset of the host's (#747).
    A same-`TIME_PERIOD` read of a richer key nest (a matrix row) is an
    identity join, not a year-0 seed (#649). A `T+1` `IF` look-ahead into
    a richer-keyed producer is not a seed; callers must probe only at the
    host boundary (#745).
    """
    if graph is None or index < 0 or index >= len(series.cells):
        return None
    host_cell = series.cells[index]
    try:
        ast = node_formula_ast(graph, host_cell)
    except InvertedTreeExportError:
        return None
    matched: list[CanonicalAddress] = []
    for ref in _iter_cell_ref_nodes(ast):
        if not _ref_shifts_with_host(ref, host_cell):
            continue
        address = as_canonical(resolve_cell_ref(ref, host_cell))
        owner = catalog.series_for(address)
        if owner is None or owner.series_id == series.series_id:
            continue
        if not is_seed_access(series, owner, address, host_cell, catalog, delta=delta):
            continue
        matched.append(address)
    unique = unique_seed_or_none(series, host_cell, catalog, matched, delta=delta)
    if unique is not None or matched:
        return unique
    if index != _boundary_index(series, catalog, delta=delta):
        return None
    return _non_peer_seed_ref(ast, host_cell, series, catalog)


def predecessor_address(
    series: BoundSeries,
    index: int,
    catalog: SeriesCatalog,
    graph: DependencyGraph | None = None,
) -> CanonicalAddress | None:
    """Return the lagged predecessor of `series.cells[index]`, if any.

    For index > 0 this is the previous cell in the series. For index 0 a seed
    is a relative read of a producer with no schedule axis, or of one at
    schedule-axis coordinate host - 1. Everything else is an aligned read.
    """
    if index < 0 or index >= len(series.cells):
        return None
    if index > 0:
        return series.cells[index - 1]
    return _adjacent_schedule_ref(series, index, catalog, graph, delta=-1)


def successor_address(
    series: BoundSeries,
    index: int,
    catalog: SeriesCatalog,
    graph: DependencyGraph | None = None,
) -> CanonicalAddress | None:
    """Return the look-ahead successor of `series.cells[index]`, if any.

    For index < len(series.cells) - 1 this is the next cell in the series. For
    the last index a terminal is a relative read of a producer with no
    schedule axis, or of one at schedule-axis coordinate host + 1.
    """
    if index < 0 or index >= len(series.cells):
        return None
    if index < len(series.cells) - 1:
        return series.cells[index + 1]
    return _adjacent_schedule_ref(series, index, catalog, graph, delta=+1)


@dataclass(frozen=True, slots=True)
class DependenceEdge:
    """One instance-level read, annotated with access class and schedule distance.

    `distance` is `axis(consumer) - axis(producer)` on the inner schedule
    axis (`TIME_PERIOD` when the key is a nest). Positive means the
    producer is an earlier period (a `pre` / lag). `access` is `identity`
    when that distance is `0` and the outer keys match, `shift` when it is
    a nonzero same-partition step, `cross_partition` when the outer keys
    differ, and `affine` when catalog-index `f(i) = coeff * i + offset`
    has `coeff != 1`. `guarded` is True when the read is guarded by a
    conditional in the dependency graph.
    """

    consumer_id: str
    producer_id: str
    consumer_cell: CanonicalAddress
    producer_cell: CanonicalAddress
    distance: int
    access: AccessClass = "identity"
    coeff: int | None = None
    offset: int | None = None
    guarded: bool = False


@dataclass(frozen=True, slots=True)
class CatalogEdges:
    """Instance-level edges for every formula series, walked once.

    `by_consumer` groups `edges` by `DependenceEdge.consumer_id` so later
    planning steps do not re-walk formula ASTs. `by_producer` groups the same
    edges by `DependenceEdge.producer_id` for incoming-adjacency lookups.
    """

    edges: tuple[DependenceEdge, ...]
    by_consumer: dict[str, tuple[DependenceEdge, ...]]
    by_producer: dict[str, tuple[DependenceEdge, ...]]


@dataclass
class SeriesDeps:
    """Emit-facing projection of one host's `DependenceEdge`s.

    `DependenceEdge` is the source of truth. This view groups those edges by
    access class so helpers can zip, lag-index, and scan without walking the
    edge list again. It does not retain the `DependencyGraph` it was
    projected from; `edges` is the host-local source of truth:

    - `aligned_ids` / `index_maps` / `affine_maps` — identity, affine, or
      irregular-gather joins already taken to the host walk. When
      `fit_affine_map` is `None`, `index_maps` keeps the observed catalog
      slots (#695).
    - `lookup_ids` — `whole` / `dynamic` table reads
    - `is_scan` / `seed_id` / `scan_direction` — self-lags discharged by
      loop order. A relative other-series read at `schedule_coord` ± 1 is
      a seed; an absolute selector read by every member is a scalar
      parameter, not a scan seed. A same-`TIME_PERIOD` slice of a matrix
      is an aligned take, not a seed (#649). A `T+1` take of a richer
      key nest is a lagged or cross-partition read, not competing scan
      terminals (#745, #747). Seed probes run only at the host
      schedule-coord boundary (`_boundary_index`), not at every catalog
      index.
    """

    host_id: str
    param_ids: tuple[str, ...]
    is_scan: bool
    seed_id: str | None
    aligned_ids: frozenset[str]
    lookup_ids: frozenset[str]
    index_maps: dict[str, tuple[int, ...]]
    affine_maps: dict[str, tuple[int, int]]
    scan_direction: Literal["forward", "reversed"] = "forward"
    edges: tuple[DependenceEdge, ...] = ()


@dataclass
class _DepCollector:
    host: BoundSeries
    catalog: SeriesCatalog
    graph: DependencyGraph | None = None
    edges: list[DependenceEdge] = field(default_factory=list)
    blank_rects: tuple[BlankRangeRect, ...] = field(default_factory=current_blank_rects)

    def _emit(
        self,
        *,
        producer_id: str,
        host_cell: CanonicalAddress,
        producer_cell: CanonicalAddress,
        access: AccessClass,
    ) -> None:
        consumer_cell = host_cell
        guarded = (
            self.graph.is_guarded(consumer_cell, producer_cell) if self.graph is not None else False
        )
        self.edges.append(
            DependenceEdge(
                consumer_id=self.host.series_id,
                producer_id=producer_id,
                consumer_cell=consumer_cell,
                producer_cell=producer_cell,
                distance=_layout_distance(consumer_cell, producer_cell, self.catalog),
                access=access,
                guarded=guarded,
            )
        )

    def emit_cell(
        self, address: CanonicalAddress, host_cell: CanonicalAddress, host_index: int
    ) -> None:
        if (
            address_in_blank_ranges(address, self.blank_rects)
            and self.catalog.series_for(address) is None
        ):
            return
        owner = self.catalog.require_series_for(address)
        if owner.series_id == self.host.series_id:
            if address == self.host.cells[host_index]:
                return
            self._emit(
                producer_id=owner.series_id,
                host_cell=host_cell,
                producer_cell=address,
                access=_member_access(host_cell, owner, address, self.catalog),
            )
            return
        self._emit(
            producer_id=owner.series_id,
            host_cell=host_cell,
            producer_cell=address,
            access=_member_access(host_cell, owner, address, self.catalog),
        )

    def emit_lookup(
        self,
        producer: BoundSeries,
        host_cell: CanonicalAddress,
        producer_cell: CanonicalAddress,
        access: AccessClass,
    ) -> None:
        if producer.series_id == self.host.series_id:
            return
        self._emit(
            producer_id=producer.series_id,
            host_cell=host_cell,
            producer_cell=producer_cell,
            access=access,
        )

    def _visit_range_addresses(
        self,
        addresses: list[CanonicalAddress],
        host_cell: CanonicalAddress,
        *,
        ref: AstNode | None = None,
        access: AccessClass = "whole",
    ) -> None:
        cells, missing = resolve_positional_range(
            addresses, self.catalog, self.blank_rects, self.graph
        )
        if missing:
            label = f"range {range_ref_label(ref, host_cell)}" if ref is not None else "range"
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r} cell {host_cell}: "
                f"{label} is not a bound series (unbound cells: {list(missing[:8])})"
            )
        seen: set[str] = set()
        for cell in cells:
            if cell.blank or cell.series_id is None or cell.series_id in seen:
                continue
            seen.add(cell.series_id)
            self.emit_lookup(self.catalog.get(cell.series_id), host_cell, cell.address, access)

    def _visit_lookup_array(
        self,
        node: AstNode,
        host_cell: CanonicalAddress,
        host_index: int,
    ) -> None:
        if isinstance(node, (RangeNode, WholeColumnNode, WholeRowNode)):
            self._visit_range_addresses(
                iter_ref_addresses(node, host_cell, self.graph),
                host_cell,
                ref=node,
                access="dynamic",
            )
            return
        if isinstance(node, CellRefNode):
            address = as_canonical(resolve_cell_ref(node, host_cell))
            if address_in_blank_ranges(address, self.blank_rects):
                raise InvertedTreeExportError(
                    f"series {self.host.series_id!r} cell {host_cell}: "
                    f"{address} is not a bound series"
                )
            owner = self.catalog.require_series_for(address)
            self.emit_lookup(owner, host_cell, address, "dynamic")
            return
        self.visit(node, host_cell=host_cell, host_index=host_index)

    def visit(self, node: AstNode, *, host_cell: CanonicalAddress, host_index: int) -> None:
        match node:
            case CellRefNode():
                address = as_canonical(resolve_cell_ref(node, host_cell))
                self.emit_cell(address, host_cell, host_index)
            case RangeNode() | WholeColumnNode() | WholeRowNode():
                self._visit_range_addresses(
                    iter_ref_addresses(node, host_cell, self.graph),
                    host_cell,
                    ref=node,
                )
            case FunctionCallNode():
                self._visit_function(node, host_cell=host_cell, host_index=host_index)
            case BinaryOpNode():
                self.visit(node.left, host_cell=host_cell, host_index=host_index)
                self.visit(node.right, host_cell=host_cell, host_index=host_index)
            case UnaryOpNode():
                self.visit(node.operand, host_cell=host_cell, host_index=host_index)
            case _:
                return

    def _visit_function(
        self,
        node: FunctionCallNode,
        *,
        host_cell: CanonicalAddress,
        host_index: int,
    ) -> None:
        name = normalize_excel_function_name(node.name)
        if name == "OFFSET":
            self._visit_offset(node, host_cell=host_cell, host_index=host_index)
            return
        if name == "INDEX":
            self._visit_index(node, host_cell=host_cell, host_index=host_index)
            return
        if name == "INDIRECT":
            self._visit_indirect(node, host_cell=host_cell, host_index=host_index)
            return
        if name == "MATCH":
            self._visit_match(node, host_cell=host_cell, host_index=host_index)
            return
        for index, arg in enumerate(node.args):
            if is_ref_only_arg(name, index) and _ref_only_argument_is_address_shaped(arg):
                continue
            self.visit(arg, host_cell=host_cell, host_index=host_index)

    def _visit_offset(
        self,
        node: FunctionCallNode,
        *,
        host_cell: CanonicalAddress,
        host_index: int,
    ) -> None:
        if not node.args:
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r}: OFFSET with no arguments"
            )
        base = node.args[0]
        if isinstance(base, FunctionCallNode):
            self._visit_offset_from_expr(node, host_cell=host_cell, host_index=host_index)
            return
        layout = _offset_column_layout(
            node,
            host_cell,
            self.catalog,
            self.graph,
            blank_rects=self.blank_rects,
            host_series_id=self.host.series_id,
        )
        if layout is not None:
            seen: set[str] = set()
            for _col, address, series in layout.columns:
                if series.series_id in seen:
                    continue
                seen.add(series.series_id)
                self.emit_lookup(series, host_cell, address, "dynamic")
            for arg in node.args[1:]:
                self.visit(arg, host_cell=host_cell, host_index=host_index)
            return
        resolved = resolve_offset_destination_series(
            node,
            host_cell,
            self.catalog,
            self.graph,
            blank_rects=self.blank_rects,
        )
        if resolved is not None:
            table, anchor = resolved
        else:
            table = self._series_for_ref(base, host_cell)
            anchor = self._ref_anchor(base, host_cell)
        self.emit_lookup(table, host_cell, anchor, "dynamic")
        for arg in node.args[1:]:
            self.visit(arg, host_cell=host_cell, host_index=host_index)

    def _visit_offset_from_expr(
        self,
        node: FunctionCallNode,
        *,
        host_cell: CanonicalAddress,
        host_index: int,
    ) -> None:
        base = node.args[0]
        if isinstance(base, FunctionCallNode) and (
            normalize_excel_function_name(base.name) == "INDEX"
        ):
            if len(base.args) < 2:
                raise InvertedTreeExportError(
                    f"series {self.host.series_id!r}: INDEX expects a range and row"
                )
            self._visit_index_selectors(base, host_cell=host_cell, host_index=host_index)
            if index_call_is_ref(base, host_cell):
                for arg in node.args[1:]:
                    self.visit(arg, host_cell=host_cell, host_index=host_index)
                return
        else:
            self.visit(base, host_cell=host_cell, host_index=host_index)
        resolved = resolve_offset_destination_series(
            node,
            host_cell,
            self.catalog,
            self.graph,
            blank_rects=self.blank_rects,
        )
        if resolved is None:
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r}: reference is not a bound series"
            )
        table, anchor = resolved
        self.emit_lookup(table, host_cell, anchor, "dynamic")
        for arg in node.args[1:]:
            self.visit(arg, host_cell=host_cell, host_index=host_index)

    def _visit_index_selectors(
        self,
        node: FunctionCallNode,
        *,
        host_cell: CanonicalAddress,
        host_index: int,
    ) -> None:
        row_arg = node.args[1]
        col_arg = node.args[2] if len(node.args) > 2 else None
        if isinstance(row_arg, FunctionCallNode) and (
            normalize_excel_function_name(row_arg.name) == "MATCH"
        ):
            self._visit_match(row_arg, host_cell=host_cell, host_index=host_index)
        else:
            self.visit(row_arg, host_cell=host_cell, host_index=host_index)
        if col_arg is not None and index_column_literal(col_arg, host_cell) is None:
            try:
                self.visit(col_arg, host_cell=host_cell, host_index=host_index)
            except InvertedTreeExportError as exc:
                raise InvertedTreeExportError(
                    f"series {self.host.series_id!r} cell {host_cell}: "
                    f"INDEX column cannot be lowered ({exc})"
                ) from exc

    def _visit_indirect(
        self,
        node: FunctionCallNode,
        *,
        host_cell: CanonicalAddress,
        host_index: int,
    ) -> None:
        del host_index
        if self.graph is None:
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r} cell {host_cell}: "
                "INDIRECT has no graph to classify"
            )
        exclude = indirect_argument_addresses(node, host_cell)
        targets = indirect_target_addresses(self.graph, host_cell, exclude=tuple(exclude))
        if not targets:
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r} cell {host_cell}: INDIRECT has no resolved edges"
            )
        covered = covering_series(self.catalog, targets)
        if covered is None:
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r} cell {host_cell}: "
                "INDIRECT targets are not one bound series"
            )
        self.emit_lookup(covered, host_cell, targets[0], "dynamic")

    def _visit_index(
        self,
        node: FunctionCallNode,
        *,
        host_cell: CanonicalAddress,
        host_index: int,
    ) -> None:
        if len(node.args) < 2:
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r}: INDEX expects a range and row"
            )
        col_arg = node.args[2] if len(node.args) > 2 else None
        self._visit_index_selectors(node, host_cell=host_cell, host_index=host_index)
        if col_arg is None:
            col_index, col_literal = 1, True
        else:
            col_index = index_column_literal(col_arg, host_cell)
            col_literal = col_index is not None
            if col_index is None:
                col_index = 1
        if isinstance(node.args[0], RangeNode):
            start = resolve_cell_ref(node.args[0].start_ref, host_cell)
            end = resolve_cell_ref(node.args[0].end_ref, host_cell)
            covered_full = covering_series_of_range(self.catalog, start, end)
            covered_col = None
            if col_literal:
                try:
                    covered_col = covering_series_of_column(self.catalog, start, end, col_index)
                except InvertedTreeExportError:
                    covered_col = None
            if covered_full is not None:
                self.emit_lookup(covered_full, host_cell, as_canonical(start), "dynamic")
            elif covered_col is not None:
                self.emit_lookup(
                    covered_col, host_cell, range_column_origin(start, end, col_index), "dynamic"
                )
            elif col_literal and self._visit_index_worksheet_column(
                node.args[0], col_index, host_cell=host_cell
            ):
                return
            else:
                self._visit_range_addresses(
                    iter_range_addresses(start, end),
                    host_cell,
                    ref=node.args[0],
                    access="dynamic",
                )
        else:
            self.visit(node.args[0], host_cell=host_cell, host_index=host_index)

    def _visit_index_worksheet_column(
        self,
        node: RangeNode,
        col_index: int,
        *,
        host_cell: CanonicalAddress,
    ) -> bool:
        """Record deps for one INDEX worksheet column, including row-label holes.

        Returns False when the selector is outside the range so the caller can
        keep the full-range visit that lets Excel `INDEX` return `#REF!`.
        """
        start = as_canonical(resolve_cell_ref(node.start_ref, host_cell))
        end = as_canonical(resolve_cell_ref(node.end_ref, host_cell))
        parent = iter_range_addresses(start, end)
        first_col = min(parse_cell_coords(start)[2], parse_cell_coords(end)[2])
        selected = [
            address
            for address in parent
            if parse_cell_coords(address)[2] == first_col + col_index - 1
        ]
        if not selected:
            return False
        cells, missing = resolve_positional_range(
            selected, self.catalog, self.blank_rects, self.graph
        )
        cells, missing = _attach_index_label_cells(
            selected, parent, cells, missing, self.catalog, self.graph
        )
        if missing:
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r} cell {host_cell}: "
                f"range {range_ref_label(node, host_cell)} is not a bound series "
                f"(unbound cells: {list(missing[:8])})"
            )
        seen: set[str] = set()
        for cell in cells:
            if cell.blank or cell.series_id is None or cell.series_id in seen:
                continue
            seen.add(cell.series_id)
            self.emit_lookup(self.catalog.get(cell.series_id), host_cell, cell.address, "dynamic")
        return True

    def _visit_match(
        self,
        node: FunctionCallNode,
        *,
        host_cell: CanonicalAddress,
        host_index: int,
    ) -> None:
        if len(node.args) < 2:
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r}: MATCH expects lookup and array"
            )
        self.visit(node.args[0], host_cell=host_cell, host_index=host_index)
        self._visit_lookup_array(node.args[1], host_cell, host_index)
        for arg in node.args[2:]:
            self.visit(arg, host_cell=host_cell, host_index=host_index)

    def _series_for_ref(self, node: AstNode, host_cell: CanonicalAddress) -> BoundSeries:
        if isinstance(node, CellRefNode):
            address = as_canonical(resolve_cell_ref(node, host_cell))
            if address_in_blank_ranges(address, self.blank_rects):
                raise InvertedTreeExportError(
                    f"series {self.host.series_id!r}: reference is not a bound series"
                )
            return self.catalog.require_series_for(address)
        if isinstance(node, (RangeNode, WholeColumnNode, WholeRowNode)):
            addresses = addresses_outside_blank_ranges(
                iter_ref_addresses(node, host_cell, self.graph),
                self.blank_rects,
            )
            covered = covering_series(self.catalog, addresses) if addresses else None
            if covered is None:
                raise InvertedTreeExportError(
                    f"series {self.host.series_id!r}: reference is not a bound series"
                )
            return covered
        if isinstance(node, FunctionCallNode):
            name = normalize_excel_function_name(node.name)
            if name == "INDEX":
                covered = covering_series_for_index_window(
                    node, host_cell, self.catalog, blank_rects=self.blank_rects
                )
                if covered is not None:
                    return covered
                raise InvertedTreeExportError(
                    f"series {self.host.series_id!r}: reference is not a bound series"
                )
            if name == "OFFSET":
                resolved = resolve_offset_destination_series(
                    node,
                    host_cell,
                    self.catalog,
                    self.graph,
                    blank_rects=self.blank_rects,
                )
                if resolved is not None:
                    return resolved[0]
                raise InvertedTreeExportError(
                    f"series {self.host.series_id!r}: reference is not a bound series"
                )
        raise InvertedTreeExportError(
            f"series {self.host.series_id!r}: expected a cell or range reference, "
            f"got {type(node).__name__}"
        )

    def _ref_anchor(self, node: AstNode, host_cell: CanonicalAddress) -> CanonicalAddress:
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
        raise InvertedTreeExportError(
            f"series {self.host.series_id!r}: expected a cell or range reference, "
            f"got {type(node).__name__}"
        )


def refine_access_classes(
    edges: Sequence[DependenceEdge],
    catalog: SeriesCatalog,
) -> list[DependenceEdge]:
    """Rewrite identity/shift bundles that lie on an integer affine map."""
    groups: dict[tuple[str, str], list[int]] = {}
    result = list(edges)
    for index, edge in enumerate(result):
        if edge.access in {"whole", "dynamic", "gather", "cross_partition"}:
            continue
        groups.setdefault((edge.consumer_id, edge.producer_id), []).append(index)
    for (consumer_id, producer_id), indexes in groups.items():
        consumer = catalog.get(consumer_id)
        producer = catalog.get(producer_id)
        pairs: list[tuple[int, int]] = []
        valid = True
        for index in indexes:
            edge = result[index]
            host = consumer.index_of(edge.consumer_cell)
            prod = producer.index_of(edge.producer_cell)
            if host is None or prod is None:
                valid = False
                break
            pairs.append((host, prod))
        if not valid:
            continue
        fitted = fit_affine_map(pairs)
        if fitted is None or fitted[0] == 1:
            continue
        coeff, offset = fitted
        for index in indexes:
            result[index] = replace(result[index], access="affine", coeff=coeff, offset=offset)
    return result


def collect_series_edges(
    series: BoundSeries,
    *,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
    blank_rects: tuple[BlankRangeRect, ...] | None = None,
) -> list[DependenceEdge]:
    """Walk `series` formulas and return classified instance-level edges."""
    rects = current_blank_rects() if blank_rects is None else blank_rects
    collector = _DepCollector(host=series, catalog=catalog, graph=graph, blank_rects=rects)
    for index, address in enumerate(series.cells):
        ast = try_formula_ast(graph, address)
        if ast is None:
            hole = series.hole_at(index)
            if hole is not None and hole.kind == "bound_leaf" and hole.claimant_id is not None:
                claimant = catalog.get(hole.claimant_id)
                collector._emit(
                    producer_id=claimant.series_id,
                    host_cell=address,
                    producer_cell=address,
                    access=_member_access(address, claimant, address, catalog),
                )
            continue
        collector.visit(ast, host_cell=address, host_index=index)
    return refine_access_classes(collector.edges, catalog)


def collect_catalog_edges(
    catalog: SeriesCatalog,
    graph: DependencyGraph,
    *,
    blank_rects: tuple[BlankRangeRect, ...] | None = None,
) -> CatalogEdges:
    """Walk each formula series once and return catalog-wide classified edges."""
    by_consumer: dict[str, tuple[DependenceEdge, ...]] = {}
    collected: list[DependenceEdge] = []
    for series in catalog.formula_series():
        series_edges = tuple(
            collect_series_edges(series, catalog=catalog, graph=graph, blank_rects=blank_rects)
        )
        by_consumer[series.series_id] = series_edges
        collected.extend(series_edges)
    by_producer: dict[str, list[DependenceEdge]] = {}
    for edge in collected:
        by_producer.setdefault(edge.producer_id, []).append(edge)
    return CatalogEdges(
        edges=tuple(collected),
        by_consumer=by_consumer,
        by_producer={producer_id: tuple(edges) for producer_id, edges in by_producer.items()},
    )


def _int_key(value: object) -> int | None:
    """Return `value` when it is a non-bool integer year or offset."""
    if isinstance(value, bool) or not isinstance(value, int):
        return None
    return value


def _time_period_at(series: BoundSeries, index: int) -> int | None:
    """Return the `TIME_PERIOD` of catalog member `index`, if present."""
    if "TIME_PERIOD" not in series.key_fields or index >= len(series.domain):
        return None
    try:
        return _int_key(series.domain[index]["TIME_PERIOD"])
    except KeyError:
        return None


def _year_deltas(
    host: BoundSeries,
    host_index: int,
    producer: BoundSeries,
    indices: set[int],
) -> set[int] | None:
    """Return `producer_year - host_year` for each slot, or `None`."""
    host_year = _time_period_at(host, host_index)
    if host_year is None:
        return None
    deltas: set[int] = set()
    for index in indices:
        year = _time_period_at(producer, index)
        if year is None:
            return None
        deltas.add(year - host_year)
    return deltas


def _host_ref_kinds(
    host: BoundSeries,
    host_index: int,
    producer: BoundSeries,
    indices: set[int],
    graph: DependencyGraph | None,
) -> tuple[set[int], set[int]] | None:
    """Return `(shifting, pinned)` producer slots from formula `$` axes.

    `None` means the host formula is unavailable, so callers must not treat
    catalog adjacency as a lag.
    """
    if graph is None or host_index < 0 or host_index >= len(host.cells):
        return None
    host_cell = host.cells[host_index]
    ast = try_formula_ast(graph, host_cell)
    if ast is None:
        return None
    shifting: set[int] = set()
    pinned: set[int] = set()
    for ref in _iter_cell_ref_nodes(ast):
        address = as_canonical(resolve_cell_ref(ref, host_cell))
        index = producer.index_of(address)
        if index is None or index not in indices:
            continue
        if _ref_shifts_with_host(ref, host_cell):
            shifting.add(index)
        else:
            pinned.add(index)
    return shifting, pinned


def _is_consistent_lag(
    host: BoundSeries,
    producer: BoundSeries,
    per_host: dict[int, set[int]],
    graph: DependencyGraph | None,
) -> bool:
    """True when every multi-slot read is the same shifting year-delta set.

    A lag requires both positions to move with the host. An absolute pin
    (`$F$2`) is not a lag even when the catalog slots are adjacent
    (`hi - lo == 1`). Same-year dual reads have one delta and are keyed
    accesses, not a lag. Without a formula AST, a single multi-read host
    is underdetermined and is not classified as a lag.
    """
    multi = {host_i: indices for host_i, indices in per_host.items() if len(indices) > 1}
    if not multi:
        return False
    delta_sets: list[set[int]] = []
    for host_i, indices in multi.items():
        kinds = _host_ref_kinds(host, host_i, producer, indices, graph)
        if kinds is not None:
            shifting, pinned = kinds
            if pinned & indices or not indices <= shifting:
                return False
        elif graph is None and len(multi) < 2:
            return False
        deltas = _year_deltas(host, host_i, producer, indices)
        if deltas is None or len(deltas) < 2:
            return False
        delta_sets.append(deltas)
    first = delta_sets[0]
    return all(item == first for item in delta_sets)


def _dimension_bind(series: BoundSeries, field: str) -> Mapping[str, Any] | None:
    """Return the bind mapping for `field`, if the series declares one."""
    return series.dimension_bind(field)


def _infer_key_field_axis(series: BoundSeries, field: str) -> Literal["sheet", "row", "col"] | None:
    """Infer the Excel axis that determines `field` from catalog geometry."""
    by_sheet: dict[str, set[object]] = {}
    by_row: dict[int, set[object]] = {}
    by_col: dict[int, set[object]] = {}
    for cell, point in zip(series.cells, series.domain, strict=False):
        try:
            value = point[field]
        except KeyError:
            return None
        sheet, row, col = parse_cell_coords(cell)
        by_sheet.setdefault(sheet, set()).add(value)
        by_row.setdefault(row, set()).add(value)
        by_col.setdefault(col, set()).add(value)

    def _is_axis(groups: dict[Any, set[object]]) -> bool:
        if not groups or not all(len(values) == 1 for values in groups.values()):
            return False
        return len({next(iter(values)) for values in groups.values()}) > 1

    candidates: list[Literal["sheet", "row", "col"]] = []
    if _is_axis(by_sheet):
        candidates.append("sheet")
    if _is_axis(by_row):
        candidates.append("row")
    if _is_axis(by_col):
        candidates.append("col")
    return candidates[0] if len(candidates) == 1 else None


def _key_field_axis(series: BoundSeries, field: str) -> Literal["sheet", "row", "col"] | None:
    """Return the Excel axis that binds `field`, or `None` when unknown.

    Bind kind wins when declared (`sheet_name`, `column_header`, `row_label`,
    `row_index`, `column_letter`, `value_map`). Otherwise the catalog geometry
    is used: a field that is
    constant on one axis and varies on another belongs to the varying axis.

    Results, including unknown axes, are cached on the `BoundSeries` instance
    so formula-reference walks do not rescan catalog geometry. Caches are not
    shared by `series_id` across catalog instances.
    """
    cache = series._key_axis_cache
    if field in cache:
        return cache[field]
    bind = _dimension_bind(series, field)
    axis: Literal["sheet", "row", "col"] | None = None
    if bind is not None:
        kind = bind.get("kind")
        if kind == "sheet_name":
            axis = "sheet"
        elif kind in {"column_header", "column_letter"}:
            axis = "col"
        elif kind in {"row_label", "row_index"}:
            axis = "row"
        elif kind == "value_map":
            try:
                parsed_axis, _parsed = parse_value_map(bind.get("values") or {})
            except ValueError:
                parsed_axis = None
            if parsed_axis == "columns":
                axis = "col"
            elif parsed_axis == "rows":
                axis = "row"
            else:
                axis = _infer_key_field_axis(series, field)
        else:
            axis = _infer_key_field_axis(series, field)
    else:
        axis = _infer_key_field_axis(series, field)
    cache[field] = axis
    return axis


def _ref_pinned_fields(
    ref: CellRefNode,
    host_cell: CanonicalAddress,
    producer: BoundSeries,
) -> frozenset[str]:
    """Return producer keys frozen by `ref`'s `$` axes or a cross-sheet address.

    `$` locks row and column. It does not lock the sheet: a same-sheet
    `Stress!$D$2` still walks `sheet_name` with the host. A ref that
    resolves onto another sheet freezes sheet-bound keys. An unknown axis
    is frozen by any `$` so classification stays fail-closed.
    """
    address = as_canonical(resolve_cell_ref(ref, host_cell))
    host_sheet, _host_row, _host_col = parse_cell_coords(host_cell)
    addr_sheet, _addr_row, _addr_col = parse_cell_coords(address)
    cell = ref.ref
    col_locked = isinstance(cell.col, AbsoluteAxis)
    row_locked = isinstance(cell.row, AbsoluteAxis)
    same_sheet = addr_sheet == host_sheet
    pinned: set[str] = set()
    for key_name in producer.key_fields:
        axis = _key_field_axis(producer, key_name)
        axis_locked = (
            (axis == "col" and col_locked)
            or (axis == "row" and row_locked)
            or (axis == "sheet" and not same_sheet)
            or (axis is None and (col_locked or row_locked))
        )
        if axis_locked:
            pinned.add(key_name)
    return frozenset(pinned)


def _host_slot_pinned_fields(
    host: BoundSeries,
    host_index: int,
    producer: BoundSeries,
    indices: set[int],
    graph: DependencyGraph | None,
) -> dict[int, frozenset[str]] | None:
    """Map producer slots to keys frozen by the host formula, if available."""
    if graph is None or host_index < 0 or host_index >= len(host.cells):
        return None
    host_cell = host.cells[host_index]
    ast = try_formula_ast(graph, host_cell)
    if ast is None:
        return None
    pinned: dict[int, set[str]] = {}
    for ref in _iter_cell_ref_nodes(ast):
        address = as_canonical(resolve_cell_ref(ref, host_cell))
        index = producer.index_of(address)
        if index is None or index not in indices:
            continue
        pinned.setdefault(index, set()).update(_ref_pinned_fields(ref, host_cell, producer))
    return {index: frozenset(fields) for index, fields in pinned.items()}


_HOST_FOLLOW_UNSET = object()


def _host_follow_key_maps(
    host: BoundSeries,
    producer: BoundSeries,
    per_host: Mapping[int, set[int]],
) -> dict[str, dict[object, object]]:
    """Return host→producer value maps for keys that follow the host walk.

    Shared producer values that every host member reads stay literals
    (`Threshold` on every breach row). After dropping that intersection,
    each member may have at most one remaining producer value, and those
    remainders must be a function of the host value. The strings need not
    match: host `Baseline breach` may join producer `Baseline` (#741);
    host `B1` may join producer `Bounds Test 1: Real GDP Growth Shock`
    (#739). Two remainders per member are not a function of the host, so
    the field is not followed.
    """
    shared = [name for name in producer.key_fields if name in host.key_fields]
    maps: dict[str, dict[object, object]] = {name: {} for name in shared}
    valid = set(shared)
    member_values: dict[str, list[tuple[object, set[object]]]] = {name: [] for name in shared}
    for host_index, indices in per_host.items():
        if host_index >= len(host.domain):
            return {}
        host_point = host.domain[host_index]
        for name in tuple(valid):
            try:
                host_value = host_point[name]
            except KeyError:
                valid.discard(name)
                continue
            producer_values: set[object] = set()
            for index in indices:
                if index >= len(producer.domain):
                    valid.discard(name)
                    producer_values = set()
                    break
                try:
                    producer_values.add(producer.domain[index][name])
                except KeyError:
                    valid.discard(name)
                    producer_values = set()
                    break
            if name not in valid:
                continue
            member_values[name].append((host_value, producer_values))
    for name in tuple(valid):
        reads = member_values[name]
        if not reads:
            continue
        intersection = set.intersection(*(values for _, values in reads))
        for host_value, producer_values in reads:
            remaining = producer_values - intersection
            if len(remaining) > 1:
                valid.discard(name)
                break
            if not remaining:
                continue
            producer_value = next(iter(remaining))
            existing = maps[name].get(host_value, _HOST_FOLLOW_UNSET)
            if existing is not _HOST_FOLLOW_UNSET and existing != producer_value:
                valid.discard(name)
                break
            maps[name][host_value] = producer_value
    return {name: maps[name] for name in valid}


def _producer_slots_in_node(
    node: AstNode,
    host_cell: CanonicalAddress,
    producer: BoundSeries,
    indices: set[int],
) -> set[int]:
    """Return producer catalog slots among `indices` referenced by `node`."""
    slots: set[int] = set()
    for ref in _iter_cell_ref_nodes(node):
        address = as_canonical(resolve_cell_ref(ref, host_cell))
        index = producer.index_of(address)
        if index is not None and index in indices:
            slots.add(index)
    return slots


def _if_pair_slot_candidates(
    node: AstNode,
    host_cell: CanonicalAddress,
    producer: BoundSeries,
    indices: set[int],
) -> list[tuple[int, int]]:
    """Collect `IF` then/else producer-slot pairs that cover `indices`."""
    found: list[tuple[int, int]] = []
    match node:
        case FunctionCallNode(name=name, args=args):
            if normalize_excel_function_name(name) == "IF" and len(args) >= 3:
                then_slots = _producer_slots_in_node(args[1], host_cell, producer, indices)
                else_slots = _producer_slots_in_node(args[2], host_cell, producer, indices)
                if (
                    len(then_slots) == 1
                    and len(else_slots) == 1
                    and then_slots != else_slots
                    and then_slots | else_slots == indices
                ):
                    found.append((next(iter(then_slots)), next(iter(else_slots))))
            for arg in args:
                found.extend(_if_pair_slot_candidates(arg, host_cell, producer, indices))
        case BinaryOpNode(left=left, right=right):
            found.extend(_if_pair_slot_candidates(left, host_cell, producer, indices))
            found.extend(_if_pair_slot_candidates(right, host_cell, producer, indices))
        case UnaryOpNode(operand=operand):
            found.extend(_if_pair_slot_candidates(operand, host_cell, producer, indices))
        case _:
            pass
    return found


def _if_pair_slots(
    ast: AstNode,
    host_cell: CanonicalAddress,
    producer: BoundSeries,
    indices: set[int],
) -> tuple[int, int] | None:
    """Return `(then_slot, else_slot)` when one `IF` covers `indices`.

    Correspondence is then/else operand position, not catalog order. Two
    matching `IF`s, or none, stay unclassifiable.
    """
    found = _if_pair_slot_candidates(ast, host_cell, producer, indices)
    if len(found) != 1:
        return None
    return found[0]


def _host_follow_pair_maps(
    host: BoundSeries,
    producer: BoundSeries,
    per_host: Mapping[int, set[int]],
    graph: DependencyGraph | None,
) -> dict[str, dict[object, tuple[object, object]]]:
    """Return host->`(then, else)` maps for one `IF`-aligned pair field.

    A same-year dual of two producer `SCENARIO`s is keyed when each host
    member reads its own pair (`B2` -> `B2.1`/`B2.2`, `B6` -> `B6.1`/`B6.2`).
    `_host_follow_key_maps` cannot remap that 1:2 leftover. Branch order
    comes from the host `IF`, not catalog adjacency or string prefixes
    (#752). A single distinct pair stays on the literal path; two or more
    distinct pairs are required here. Missing AST, a non-`IF` dual, or a
    second differing key fail closed.
    """
    if graph is None or not producer.key_fields:
        return {}
    shared = [name for name in producer.key_fields if name in host.key_fields]
    if not shared:
        return {}
    multi = {host_i: indices for host_i, indices in per_host.items() if len(indices) > 1}
    if not multi or any(len(indices) != 2 for indices in multi.values()):
        return {}
    members: list[tuple[KeyPoint, KeyPoint, KeyPoint]] = []
    for host_index, indices in multi.items():
        if host_index >= len(host.domain) or host_index >= len(host.cells):
            return {}
        ast = try_formula_ast(graph, host.cells[host_index])
        if ast is None:
            return {}
        ordered = _if_pair_slots(ast, host.cells[host_index], producer, indices)
        if ordered is None:
            return {}
        then_index, else_index = ordered
        if then_index >= len(producer.domain) or else_index >= len(producer.domain):
            return {}
        members.append(
            (host.domain[host_index], producer.domain[then_index], producer.domain[else_index])
        )
    if len(members) < 2:
        return {}
    pair_fields: set[str] | None = None
    for _host_point, then_point, else_point in members:
        differing: set[str] = set()
        for name in producer.key_fields:
            try:
                then_value = then_point[name]
                else_value = else_point[name]
            except KeyError:
                return {}
            if then_value != else_value:
                differing.add(name)
        differing &= set(shared)
        if pair_fields is None:
            pair_fields = differing
        elif pair_fields != differing:
            return {}
    if pair_fields is None or len(pair_fields) != 1:
        return {}
    field = next(iter(pair_fields))
    maps: dict[object, tuple[object, object]] = {}
    for host_point, then_point, else_point in members:
        try:
            host_value = host_point[field]
            pair = (then_point[field], else_point[field])
        except KeyError:
            return {}
        existing = maps.get(host_value)
        if existing is not None and existing != pair:
            return {}
        maps[host_value] = pair
    if len(set(maps.values())) < 2:
        return {}
    return {field: maps}


def _host_join_key(
    host: BoundSeries,
    host_index: int,
    fields: Sequence[str],
) -> tuple[object, ...] | None:
    """Return host values for `fields`, or `None` when a field is missing."""
    if host_index >= len(host.domain):
        return None
    if not fields:
        return ()
    point = host.domain[host_index]
    parts: list[object] = []
    for name in fields:
        try:
            parts.append(point[name])
        except KeyError:
            return None
    return tuple(parts)


def _lookup_axis_fields(
    host: BoundSeries,
    producer: BoundSeries,
    per_host: Mapping[int, set[int]],
) -> frozenset[str]:
    """Return producer keys whose multi-read value set is partition-invariant.

    A field whose values are the same set for every multi-slot host
    member of a joined partition, and that set has more than one value,
    is a catalog lookup axis: the host indexes those values (`CHOOSE` /
    a listed year-row) rather than following its own key. Bind the field
    as a literal even when one value equals the host year (#754).
    Partitions may disagree on the set (#760): IMF `{2024…2051}` and
    IDA's longer horizon are still lookup axes. A leftover that changes
    with the host (`{t, t-1}` or `{t, 2026}`) is not a lookup axis. One
    multi-read member cannot distinguish a lookup axis from a mixed
    host-year plus pin, so every partition that multi-reads needs at
    least two members, and at least two multi-read members are required
    overall.
    """
    multi = {host_i: indices for host_i, indices in per_host.items() if len(indices) > 1}
    if len(multi) < 2:
        return frozenset()
    axes: set[str] = set()
    for name in producer.key_fields:
        partition_fields = tuple(
            field for field in producer.key_fields if field in host.key_fields and field != name
        )
        groups: dict[tuple[object, ...] | None, list[set[object]]] = {}
        valid = True
        for host_i, indices in multi.items():
            values: set[object] = set()
            for index in indices:
                if index >= len(producer.domain):
                    valid = False
                    break
                try:
                    values.add(producer.domain[index][name])
                except KeyError:
                    valid = False
                    break
            if not valid:
                break
            part = _host_join_key(host, host_i, partition_fields)
            if part is None and partition_fields:
                valid = False
                break
            groups.setdefault(part, []).append(values)
        if not valid or not groups:
            continue
        if any(len(value_sets) < 2 for value_sets in groups.values()):
            continue
        if any(
            len(value_sets[0]) <= 1 or not all(item == value_sets[0] for item in value_sets)
            for value_sets in groups.values()
        ):
            continue
        axes.add(name)
    return frozenset(axes)


def _field_binding(
    host: BoundSeries,
    host_index: int,
    producer: BoundSeries,
    producer_index: int,
    *,
    pinned_fields: frozenset[str] = frozenset(),
    literal_fields: frozenset[str] = frozenset(),
    host_follow: Mapping[str, Mapping[object, object]] | None = None,
    pair_maps: Mapping[str, Mapping[object, tuple[object, object]]] | None = None,
) -> tuple[tuple[str, object], ...] | None:
    """Return `(field, 'host' | ('lit', value) | ('pair', branch))` keys.

    A field is `host` when it equals the consumer's value, or when
    `host_follow` maps the host value onto this producer value, and the
    field is not frozen by a `$` pin on that field's bind axis or by a
    lookup axis (`literal_fields`). `$F$2` keeps `TIME_PERIOD` a literal
    even when that year equals the host year. A same-sheet `$D$2` does
    not freeze `sheet_name` (`SCENARIO` stays `host`). An invariant
    multi-value year-row is a lookup axis even without `$` (#754),
    including when joined partitions list different years (#760).
    Shared constants stay literals beside a remapped path (#741).
    `pair_maps` mark an `IF` then/else leftover as `('pair', 0)` /
    `('pair', 1)` so two host rows that dual-read different producer
    pairs still share a pattern (#752). Emit can replay a binding
    across the host walk only when every member agrees on this spec.
    """
    if host_index >= len(host.domain) or producer_index >= len(producer.domain):
        return None
    host_point = host.domain[host_index]
    prod = producer.domain[producer_index]
    parts: list[tuple[str, object]] = []
    follow = host_follow or {}
    pairs = pair_maps or {}
    frozen_fields = pinned_fields | literal_fields
    for key_name in producer.key_fields:
        try:
            value = prod[key_name]
        except KeyError:
            return None
        if key_name not in frozen_fields and key_name in host.key_fields:
            try:
                host_value = host_point[key_name]
            except KeyError:
                host_value = _HOST_FOLLOW_UNSET
            if host_value is not _HOST_FOLLOW_UNSET:
                if host_value == value:
                    parts.append((key_name, "host"))
                    continue
                mapped = follow.get(key_name)
                if mapped is not None and mapped.get(host_value) == value:
                    parts.append((key_name, "host"))
                    continue
                pair_mapped = pairs.get(key_name)
                if pair_mapped is not None:
                    pair = pair_mapped.get(host_value)
                    if pair is not None:
                        if value == pair[0]:
                            parts.append((key_name, ("pair", 0)))
                            continue
                        if value == pair[1]:
                            parts.append((key_name, ("pair", 1)))
                            continue
        parts.append((key_name, ("lit", value)))
    return tuple(parts)


def _multi_read_pattern_sets(
    host: BoundSeries,
    producer: BoundSeries,
    per_host: dict[int, set[int]],
    graph: DependencyGraph | None,
    host_follow: Mapping[str, Mapping[object, object]],
    pair_maps: Mapping[str, Mapping[object, tuple[object, object]]] | None = None,
) -> dict[int, frozenset[tuple[tuple[str, object], ...]]] | None:
    """Return per-member keyed pattern sets, or `None` when a slot is unbound."""
    lookup_axes = _lookup_axis_fields(host, producer, per_host)
    pattern_sets: dict[int, frozenset[tuple[tuple[str, object], ...]]] = {}
    for host_i, indices in per_host.items():
        if len(indices) < 2:
            continue
        pin_map = _host_slot_pinned_fields(host, host_i, producer, indices, graph)
        patterns: list[tuple[tuple[str, object], ...]] = []
        for index in indices:
            pinned_fields = pin_map.get(index, frozenset()) if pin_map is not None else frozenset()
            binding = _field_binding(
                host,
                host_i,
                producer,
                index,
                pinned_fields=pinned_fields,
                literal_fields=lookup_axes,
                host_follow=host_follow,
                pair_maps=pair_maps,
            )
            if binding is None:
                return None
            patterns.append(binding)
        unique = frozenset(patterns)
        if len(unique) != len(indices):
            return None
        pattern_sets[host_i] = unique
    return pattern_sets


def _keyed_patterns_agree(
    host: BoundSeries,
    producer: BoundSeries,
    per_host: Mapping[int, set[int]],
    pattern_sets: Mapping[int, frozenset[tuple[tuple[str, object], ...]]],
) -> bool:
    """True when multi-read pattern sets agree globally or per lookup partition.

    Without a lookup axis, every host member must share one pattern set.
    With a lookup axis, members of the same joined partition must share
    a pattern; partitions may list different literal years (#760).
    """
    if not pattern_sets:
        return False
    lookup_axes = _lookup_axis_fields(host, producer, per_host)
    if not lookup_axes:
        first = next(iter(pattern_sets.values()))
        return all(item == first for item in pattern_sets.values())
    partition_fields = tuple(
        name for name in producer.key_fields if name in host.key_fields and name not in lookup_axes
    )
    groups: dict[tuple[object, ...] | None, list[frozenset[tuple[tuple[str, object], ...]]]] = {}
    for host_i, unique in pattern_sets.items():
        part = _host_join_key(host, host_i, partition_fields)
        if part is None and partition_fields:
            return False
        groups.setdefault(part, []).append(unique)
    return all(all(item == group[0] for item in group) for group in groups.values())


def _is_keyed_multi_read(
    host: BoundSeries,
    producer: BoundSeries,
    per_host: dict[int, set[int]],
    graph: DependencyGraph | None,
) -> bool:
    """True when every multi-slot host shares the same host-or-literal keys.

    Each slot is then `domain.index` of those fields: two scenarios at one
    year, `baseline[t]` plus `baseline[2026]`, two same-sheet variants
    (`stats[s, mean]` / `stats[s, stdev]`), two instruments at one remapped
    host scenario (#739), a remapped path plus a constant cap (#741), an
    `IF` then/else pair of producer scenarios that is a function of the
    host row (#752), or an N-way catalog along a lookup axis of a joined
    producer (`CHOOSE(k, p[t0], …)` / a sum of that year-row, #754). A
    `t-1` read is a literal that changes per member and cannot be keyed.
    A `$` pin whose year happens to equal the host year is still a
    literal; a `$` pin on the host sheet is not a `SCENARIO` literal.
    An invariant year set is a lookup axis even without `$`. Joined
    partitions may list different years on that axis (#760).
    """
    if not producer.key_fields:
        return False
    host_follow = _host_follow_key_maps(host, producer, per_host)
    pattern_sets = _multi_read_pattern_sets(host, producer, per_host, graph, host_follow)
    if pattern_sets is not None and _keyed_patterns_agree(host, producer, per_host, pattern_sets):
        return True
    pair_maps = _host_follow_pair_maps(host, producer, per_host, graph)
    if not pair_maps:
        return False
    pattern_sets = _multi_read_pattern_sets(host, producer, per_host, graph, host_follow, pair_maps)
    return pattern_sets is not None and _keyed_patterns_agree(
        host, producer, per_host, pattern_sets
    )


def series_deps_from_edges(
    host: BoundSeries,
    edges: Sequence[DependenceEdge],
    catalog: SeriesCatalog,
    graph: DependencyGraph | None = None,
) -> SeriesDeps:
    """Derive the emit-facing `SeriesDeps` view from classified edges of `host`."""
    params: dict[str, None] = {}
    lookup_ids: set[str] = set()
    aligned_hits: dict[str, list[tuple[int, int]]] = {}
    saw_self_lag = False
    saw_forward_lag = False
    saw_backward_lag = False
    seed_id: str | None = None
    terminal_seed_id: str | None = None
    pred_boundary = succ_boundary = -1
    pred_ref: CanonicalAddress | None = None
    succ_ref: CanonicalAddress | None = None
    if host.cells:
        pred_boundary = _boundary_index(host, catalog, delta=-1)
        succ_boundary = _boundary_index(host, catalog, delta=+1)
        pred_ref = _adjacent_schedule_ref(host, pred_boundary, catalog, graph, delta=-1)
        succ_ref = _adjacent_schedule_ref(host, succ_boundary, catalog, graph, delta=+1)
    for edge in edges:
        if edge.consumer_id != host.series_id:
            continue
        if edge.producer_id == host.series_id:
            host_index = host.index_of(edge.consumer_cell)
            if host_index is None:
                continue
            if edge.distance > 0:
                saw_self_lag = True
            pred = predecessor_address(host, host_index, catalog, graph)
            if pred is not None and edge.producer_cell == pred:
                saw_forward_lag = True
            succ = successor_address(host, host_index, catalog, graph)
            if succ is not None and edge.producer_cell == succ:
                saw_backward_lag = True
            continue
        params.setdefault(edge.producer_id, None)
        if edge.access in {"whole", "dynamic"}:
            lookup_ids.add(edge.producer_id)
            continue
        host_index = host.index_of(edge.consumer_cell)
        if host_index is None:
            continue
        owner = catalog.get(edge.producer_id)
        idx = owner.index_of(edge.producer_cell)
        if idx is not None:
            aligned_hits.setdefault(edge.producer_id, []).append((host_index, idx))
        if host_index == pred_boundary and pred_ref is not None and edge.producer_cell == pred_ref:
            seed_id = edge.producer_id
        if host_index == succ_boundary and succ_ref is not None and edge.producer_cell == succ_ref:
            terminal_seed_id = edge.producer_id
    scan_direction: Literal["forward", "reversed"] = "forward"
    if saw_backward_lag and not saw_forward_lag:
        scan_direction = "reversed"
        is_scan = True
        if terminal_seed_id is not None:
            seed_id = terminal_seed_id
    elif saw_self_lag or seed_id is not None:
        scan_direction = "forward"
        is_scan = True
    else:
        is_scan = False
    remaining = [sid for sid in catalog.order if sid in params]
    if seed_id is not None and seed_id in remaining:
        remaining.remove(seed_id)
        param_ids = (seed_id, *remaining)
    else:
        param_ids = tuple(remaining)
    index_maps: dict[str, tuple[int, ...]] = {}
    affine_maps: dict[str, tuple[int, int]] = {}
    aligned: set[str] = set()
    host_n = len(host.cells)
    identity_by_producer: dict[str, list[DependenceEdge]] = {}
    for edge in edges:
        if (
            edge.consumer_id == host.series_id
            and edge.access == "identity"
            and edge.producer_id != host.series_id
        ):
            identity_by_producer.setdefault(edge.producer_id, []).append(edge)
    for series_id, pairs in aligned_hits.items():
        if series_id in lookup_ids or series_id == seed_id:
            continue
        dep = catalog.get(series_id)
        if dep.is_scalar:
            continue
        per_host: dict[int, set[int]] = {}
        for host_i, dep_i in pairs:
            per_host.setdefault(host_i, set()).add(dep_i)
        slots = [-1] * host_n
        origin = fit_affine_map(
            [(host_i, min(indices)) for host_i, indices in per_host.items() if indices]
        )
        multi = {host_i: indices for host_i, indices in per_host.items() if len(indices) > 1}
        if multi and _is_consistent_lag(host, dep, per_host, graph):
            continue
        if multi and _is_keyed_multi_read(host, dep, per_host, graph):
            continue
        static_catalog = origin is not None and origin[0] == 0
        if static_catalog and any(len(indices) > 1 for indices in per_host.values()):
            # Same catalog slots from every member (`labels[0]` / `labels[1]`),
            # or a mixed absolute + relative read of one producer (#681).
            continue
        if multi:
            # Several literal producer coordinates read by one member. Named
            # bodies spell each coordinate out; no positional class is needed.
            continue
        for host_i, indices in per_host.items():
            slots[host_i] = next(iter(indices))
        joined = identity_join_indices(host, dep, catalog)
        for edge in identity_by_producer.get(series_id, ()):
            host_index = host.index_of(edge.consumer_cell)
            if host_index is None:
                continue
            slot = joined[host_index]
            if slot < 0 or dep.cells[slot] != edge.producer_cell:
                raise InvertedTreeExportError(
                    f"series {host.series_id!r} cell {edge.consumer_cell} "
                    f"identity-reads {edge.producer_cell}, not the join slot "
                    f"of {series_id!r}"
                )
        if not all(slot >= 0 for slot in slots):
            continue
        fitted = fit_affine_map([(index, slots[index]) for index in range(host_n)])
        observed = tuple(slots)
        fields_differ = preferred_fields(host, catalog) != preferred_fields(dep, catalog)
        if observed == joined:
            index_maps[series_id] = joined
            aligned.add(series_id)
            if fitted is not None and fitted[0] != 1:
                affine_maps[series_id] = fitted
        elif fitted is None and fields_differ:
            # Positional join is dummy `(0, 1, …)`; keep the observed slots (#695).
            index_maps[series_id] = observed
            aligned.add(series_id)
        elif fitted is not None and (fitted[0] != 1 or fields_differ):
            index_maps[series_id] = observed
            aligned.add(series_id)
            if fitted[0] != 1:
                affine_maps[series_id] = fitted
    return SeriesDeps(
        host_id=host.series_id,
        param_ids=param_ids,
        is_scan=is_scan,
        seed_id=seed_id,
        aligned_ids=frozenset(aligned),
        lookup_ids=frozenset(lookup_ids),
        index_maps=index_maps,
        affine_maps=affine_maps,
        scan_direction=scan_direction,
        edges=tuple(edge for edge in edges if edge.consumer_id == host.series_id),
    )


def collect_all_deps(
    catalog: SeriesCatalog,
    graph: DependencyGraph,
    *,
    catalog_edges: CatalogEdges | None = None,
    blank_rects: tuple[BlankRangeRect, ...] | None = None,
) -> dict[str, SeriesDeps]:
    """Collect first-level deps for every formula series.

    Pass `catalog_edges` when the catalog has already been walked so this
    step does not re-visit formula ASTs.
    """
    if catalog_edges is None:
        catalog_edges = collect_catalog_edges(catalog, graph, blank_rects=blank_rects)
    deps = {
        series.series_id: series_deps_from_edges(
            series, catalog_edges.by_consumer.get(series.series_id, ()), catalog, graph
        )
        for series in catalog.formula_series()
    }
    return _with_labeller_edges(catalog, deps)


def _with_labeller_edges(
    catalog: SeriesCatalog, deps: dict[str, SeriesDeps]
) -> dict[str, SeriesDeps]:
    """Depend every series with a runtime axis on that axis's labeller."""
    for series in catalog.formula_series():
        info = deps.get(series.series_id)
        if info is None:
            continue
        extra = [
            labeller.series_id
            for axis in series.tensor_domain.axes
            if (labeller := catalog.runtime_labeller(axis.name, axis.keys)) is not None
            and labeller.series_id != series.series_id
            and labeller.series_id not in info.param_ids
        ]
        if not extra:
            continue
        extras = tuple(sid for sid in catalog.order if sid in set(extra))
        deps[series.series_id] = replace(info, param_ids=info.param_ids + extras)
    return deps


def leaf_closure(
    root_id: str,
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
) -> tuple[str, ...]:
    """Return input and constant series in the subgraph of `root_id`.

    Order follows the bindings catalog (`catalog.order`).
    """
    seen: set[str] = set()
    stack = [root_id]
    while stack:
        current = stack.pop()
        if current in seen:
            continue
        seen.add(current)
        info = deps.get(current)
        if info is None:
            continue
        for param_id in info.param_ids:
            stack.append(param_id)
    inputs = [sid for sid in catalog.order if sid in seen and catalog.get(sid).direction == "input"]
    constants = [
        sid for sid in catalog.order if sid in seen and catalog.get(sid).direction == "constant"
    ]
    return tuple(inputs + constants)


def assert_subgraph_bound(
    *,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
    roots: Iterable[str],
) -> None:
    """Fail closed when a subgraph formula is unbound or bound as a plain input."""
    seen: set[str] = set()
    stack = [canonical_address(addr) for addr in roots]
    while stack:
        address = stack.pop()
        if address in seen:
            continue
        seen.add(address)
        node = graph.get_node(address)
        if node is None:
            continue
        has_formula = bool(getattr(node, "has_formula", False))
        if has_formula:
            owner = catalog.series_for(address)
            if owner is None:
                raise InvertedTreeExportError(
                    f"unbound formula cell {address} is in the target subgraph"
                )
            if not owner.is_formula_series and not (
                owner.direction == "input" and is_override_input(cast(dict[str, Any], owner.raw))
            ):
                raise InvertedTreeExportError(
                    f"formula cell {address} is bound to {owner.direction} series "
                    f"{owner.series_id!r}, not an internal or output series"
                )
        for dep in graph.get_dependencies(address):
            stack.append(as_canonical(dep))


def all_formula_root_cells(catalog: SeriesCatalog) -> Iterator[str]:
    """Yield every cell of every formula series."""
    for series in catalog.formula_series():
        yield from series.cells
