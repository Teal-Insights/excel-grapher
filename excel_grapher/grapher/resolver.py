from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass

import fastpyxl
from fastpyxl.utils.cell import (
    column_index_from_string,
    coordinate_from_string,
    get_column_letter,
)

from excel_grapher.core.addressing import offset_range, split_sheet_qualified_address
from excel_grapher.core.coercions import to_number
from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    CellRefNode,
    FormulaParseError,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    UnaryOpNode,
)
from excel_grapher.core.formula_ast import (
    parse as parse_formula_ast,
)
from excel_grapher.core.formula_normalization import (
    expand_defined_names,
    expand_whole_column_row_for_parse,
)
from excel_grapher.core.operators_reference import apply_arithmetic
from excel_grapher.core.range_shorthand import resolve_whole_column_span, resolve_whole_row_span
from excel_grapher.core.types import CellValue, ExcelRange, XlError

_RECT_DEFINED_NAME_RE = re.compile(
    r"'?(?P<sheet>[^'!]+)'?!\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)"
    r":\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)$"
)
_WHOLE_ROW_DEFINED_NAME_RE = re.compile(r"'?(?P<sheet>[^'!]+)'?!\$?(?P<r1>\d+):\$?(?P<r2>\d+)$")
_WHOLE_COL_DEFINED_NAME_RE = re.compile(
    r"'?(?P<sheet>[^'!]+)'?!\$?(?P<c1>[A-Z]{1,3}):\$?(?P<c2>[A-Z]{1,3})$"
)
_CELL_DEFINED_NAME_RE = re.compile(r"'?([^'!]+)'?!\$?([A-Z]{1,3})\$?(\d+)$")


@dataclass(frozen=True)
class NamedRangeMaps:
    cell_map: dict[str, tuple[str, str]]
    range_map: dict[str, tuple[str, str, str]]


def _sheet_bounds(wb: fastpyxl.Workbook) -> dict[str, tuple[int, int]]:
    """Return per-sheet (max_row, max_col) from workbook dimensions."""
    bounds: dict[str, tuple[int, int]] = {}
    for name in wb.sheetnames:
        ws = wb[name]
        max_row = getattr(ws, "max_row", None) or 1
        max_col = getattr(ws, "max_column", None) or 1
        if max_row < 1:
            max_row = 1
        if max_col < 1:
            max_col = 1
        bounds[name] = (max_row, max_col)
    return bounds


def _range_node_to_excel_range_bounded(
    node: RangeNode,
    bounds: dict[str, tuple[int, int]],
) -> ExcelRange | None:
    """Convert a RangeNode to ExcelRange, capping to sheet bounds."""
    try:
        sheet_start, coord_start = node.start.split("!", 1)
        sheet_end, coord_end = node.end.split("!", 1)
    except ValueError:
        return None
    if sheet_start != sheet_end:
        return None
    sheet = sheet_start.strip("'")
    try:
        col_letter1, row1 = coordinate_from_string(coord_start)
        col_letter2, row2 = coordinate_from_string(coord_end)
    except Exception:
        return None
    start_row = min(row1, row2)
    end_row = max(row1, row2)
    start_col = min(column_index_from_string(col_letter1), column_index_from_string(col_letter2))
    end_col = max(column_index_from_string(col_letter1), column_index_from_string(col_letter2))
    max_r, max_c = bounds.get(sheet, (1048576, 16384))
    end_row = min(end_row, max_r)
    end_col = min(end_col, max_c)
    start_row = max(1, start_row)
    start_col = max(1, start_col)
    return ExcelRange(
        sheet=sheet,
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
    )


def _base_node_to_excel_range(
    node: CellRefNode | RangeNode,
    bounds: dict[str, tuple[int, int]],
) -> ExcelRange | None:
    """Interpret base argument of OFFSET as an ExcelRange."""
    if isinstance(node, CellRefNode):
        try:
            sheet, coord = node.address.split("!", 1)
            sheet = sheet.strip("'")
            col_letter, row = coordinate_from_string(coord)
            col = column_index_from_string(col_letter)
            return ExcelRange(
                sheet=sheet,
                start_row=row,
                start_col=col,
                end_row=row,
                end_col=col,
            )
        except Exception:
            return None
    if isinstance(node, RangeNode):
        return _range_node_to_excel_range_bounded(node, bounds)
    return None


def _as_int_or_float(n: float) -> int | float:
    """Return an int when *n* is integral, otherwise leave as float."""
    return int(n) if n == int(n) else n


def _eval_number_for_defined_name(
    node: AstNode,
    get_cell_value: Callable[[str], CellValue],
    bounds: dict[str, tuple[int, int]],
) -> int | float | None:
    """Evaluate an AST node to a number for OFFSET args (rows, cols, height, width).

    Supports literals, cell refs, bare `COUNTA(range)`, and arithmetic /
    unary expressions over those (e.g. `COUNTA(A:A)+5`). Unsupported nodes
    return `None` so callers can fail closed.
    """
    if isinstance(node, NumberNode):
        return _as_int_or_float(node.value)
    if isinstance(node, CellRefNode):
        val = get_cell_value(node.address)
        n = to_number(val)
        if isinstance(n, XlError):
            return None
        return _as_int_or_float(n)
    if isinstance(node, UnaryOpNode):
        inner = _eval_number_for_defined_name(node.operand, get_cell_value, bounds)
        if inner is None:
            return None
        if node.op == "-":
            return _as_int_or_float(-float(inner))
        if node.op == "+":
            return _as_int_or_float(float(inner))
        if node.op == "%":
            return float(inner) / 100.0
        return None
    if isinstance(node, BinaryOpNode):
        if node.op not in {"+", "-", "*", "/", "^"}:
            return None
        left = _eval_number_for_defined_name(node.left, get_cell_value, bounds)
        right = _eval_number_for_defined_name(node.right, get_cell_value, bounds)
        if left is None or right is None:
            return None
        result = apply_arithmetic(node.op, float(left), float(right))
        if isinstance(result, XlError):
            return None
        return _as_int_or_float(result)
    if isinstance(node, FunctionCallNode) and node.name.upper() == "COUNTA" and len(node.args) == 1:
        rng: ExcelRange | None = None
        if isinstance(node.args[0], RangeNode):
            rng = _range_node_to_excel_range_bounded(node.args[0], bounds)
        elif isinstance(node.args[0], CellRefNode):
            rng = _base_node_to_excel_range(node.args[0], bounds)
        if rng is None:
            return None
        count = 0
        for addr in rng.cell_addresses():
            v = get_cell_value(addr)
            if v is not None and v != "":
                count += 1
        return count
    return None


def _eval_offset_formula_to_range(
    node: FunctionCallNode,
    get_cell_value: Callable[[str], CellValue],
    bounds: dict[str, tuple[int, int]],
) -> tuple[str, str, str] | None:
    """Evaluate OFFSET(...) to (sheet, start_a1, end_a1) or None."""
    if node.name.upper() != "OFFSET" or len(node.args) < 3:
        return None
    base = (
        _base_node_to_excel_range(node.args[0], bounds)
        if isinstance(node.args[0], (CellRefNode, RangeNode))
        else None
    )
    if base is None:
        return None
    rows = _eval_number_for_defined_name(node.args[1], get_cell_value, bounds)
    cols = _eval_number_for_defined_name(node.args[2], get_cell_value, bounds)
    if rows is None or cols is None:
        return None
    # Explicit height/width that fail to evaluate must not fall back to the base
    # shape (which collapses cell-anchored OFFSETs to a poison 1x1 range).
    height: int | float | None
    if len(node.args) >= 4:
        height = _eval_number_for_defined_name(node.args[3], get_cell_value, bounds)
        if height is None:
            return None
    else:
        height = None
    width: int | float | None
    if len(node.args) >= 5:
        width = _eval_number_for_defined_name(node.args[4], get_cell_value, bounds)
        if width is None:
            return None
    else:
        width = None
    if height is not None and height <= 0:
        return None
    if width is not None and width <= 0:
        return None
    # Use Excel grid limits so OFFSET result is accepted even when sheet used range is smaller.
    max_r, max_c = 1048576, 16384

    class _Bounds:
        sheet = base.sheet
        min_row = 1
        min_col = 1
        max_row = max_r
        max_col = max_c

    result = offset_range(
        base,
        rows,
        cols,
        height,
        width,
        bounds=_Bounds(),
    )
    if isinstance(result, XlError):
        return None
    start_a1 = f"{get_column_letter(result.start_col)}{result.start_row}"
    end_a1 = f"{get_column_letter(result.end_col)}{result.end_row}"
    return (result.sheet, start_a1, end_a1)


def _eval_indirect_formula_to_range(
    node: FunctionCallNode,
    get_cell_value: Callable[[str], CellValue],
    bounds: dict[str, tuple[int, int]],
) -> tuple[str, str] | tuple[str, str, str] | None:
    """Evaluate INDIRECT(...) to (sheet, a1) or (sheet, start_a1, end_a1)."""
    from excel_grapher.core.formula_ast import StringNode

    if node.name.upper() != "INDIRECT" or len(node.args) < 1:
        return None
    if not isinstance(node.args[0], StringNode):
        return None
    text = node.args[0].value.strip()
    default_sheet = next(iter(bounds.keys()), "Sheet1")

    if ":" in text:
        start_text, end_text = text.split(":", 1)
        parsed_start = split_sheet_qualified_address(start_text)
        if parsed_start is None:
            sheet = default_sheet
            start_ref = start_text
        else:
            sheet, start_ref = parsed_start

        parsed_end = split_sheet_qualified_address(end_text)
        if parsed_end is None:
            end_ref = end_text
        else:
            end_sheet, end_ref = parsed_end
            if end_sheet != sheet:
                return None

        try:
            c1s, r1 = coordinate_from_string(start_ref)
            c2s, r2 = coordinate_from_string(end_ref)
            c1 = column_index_from_string(c1s)
            c2 = column_index_from_string(c2s)
        except Exception:
            return None

        if r1 > r2:
            r1, r2 = r2, r1
        if c1 > c2:
            c1, c2 = c2, c1

        start_a1 = f"{get_column_letter(c1)}{r1}"
        end_a1 = f"{get_column_letter(c2)}{r2}"
        return (sheet, start_a1, end_a1)

    parsed = split_sheet_qualified_address(text)
    if parsed is None:
        sheet = default_sheet
        addr_part = text
    else:
        sheet, addr_part = parsed
    try:
        c, r = coordinate_from_string(addr_part)
        a1 = f"{c}{r}"
        return (sheet, a1)
    except Exception:
        return None


def _normalize_formula_for_parse(formula: str, bounds: dict[str, tuple[int, int]]) -> str:
    """Strip ``$`` and expand whole-column/whole-row refs so formula_ast can parse."""
    return expand_whole_column_row_for_parse(formula, bounds)


def _excel_range_to_named_entry(rng: ExcelRange) -> tuple[str, str, str]:
    """Return `(sheet, start_a1, end_a1)` for a resolved named-range rectangle."""
    start = f"{get_column_letter(rng.start_col)}{rng.start_row}"
    end = f"{get_column_letter(rng.end_col)}{rng.end_row}"
    return (rng.sheet, start, end)


def _used_extent_for_named_sheet(
    bounds: dict[str, tuple[int, int]], sheet: str
) -> tuple[str, tuple[int, int]] | None:
    """Return `(canonical_sheet, (max_row, max_col))` or None if unknown.

    Whole-row/column names need a used-range extent. Missing sheets stay
    unresolved (fail closed) rather than expanding to `XFD`/`1048576`.
    """
    if sheet in bounds:
        return sheet, bounds[sheet]
    key = sheet.casefold()
    for name, extent in bounds.items():
        if name.casefold() == key:
            return name, extent
    return None


def _try_resolve_whole_axis_defined_name(
    attr_text: str,
    bounds: dict[str, tuple[int, int]],
) -> tuple[str, str, str] | None:
    """Expand `Sheet!$5:$10` / `Sheet!$A:$C` against the sheet used range."""
    m = _WHOLE_ROW_DEFINED_NAME_RE.match(attr_text)
    if m is not None:
        found = _used_extent_for_named_sheet(bounds, m.group("sheet"))
        if found is None:
            return None
        sheet, extent = found
        rng = resolve_whole_row_span(sheet, int(m.group("r1")), int(m.group("r2")), {sheet: extent})
        return _excel_range_to_named_entry(rng)
    m = _WHOLE_COL_DEFINED_NAME_RE.match(attr_text)
    if m is not None:
        found = _used_extent_for_named_sheet(bounds, m.group("sheet"))
        if found is None:
            return None
        sheet, extent = found
        rng = resolve_whole_column_span(sheet, m.group("c1"), m.group("c2"), {sheet: extent})
        return _excel_range_to_named_entry(rng)
    return None


def _formula_defined_name_body(attr_text: str) -> str:
    """Strip a leading `=` from a defined-name formula, if present."""
    body = attr_text.strip()
    if body.startswith("="):
        return body[1:].lstrip()
    return body


def _is_offset_or_indirect_defined_name(attr_text: str) -> bool:
    """Return True when `attr_text` is an OFFSET or INDIRECT formula."""
    upper = _formula_defined_name_body(attr_text).upper()
    return upper.startswith(("OFFSET(", "INDIRECT("))


def _store_resolved_defined_name(
    name: str,
    resolved: tuple[str, str] | tuple[str, str, str],
    cell_map: dict[str, tuple[str, str]],
    range_map: dict[str, tuple[str, str, str]],
) -> None:
    """Record a resolved defined name as a cell or rectangle."""
    if len(resolved) == 2:
        cell_map[name] = (resolved[0], resolved[1])
    elif len(resolved) == 3:
        range_map[name] = resolved


def _try_resolve_formula_defined_name(
    attr_text: str,
    wb: fastpyxl.Workbook,
    named_ranges: dict[str, tuple[str, str]] | None = None,
    named_range_ranges: dict[str, tuple[str, str, str]] | None = None,
) -> tuple[str, str, str] | tuple[str, str] | None:
    """If attr_text is an OFFSET/INDIRECT formula, evaluate to range or cell.

    Already-resolved `named_ranges` / `named_range_ranges` are substituted
    before parse so `OFFSET(START, ...)` can resolve when `START` is another
    defined name (cell or rectangle).

    Returns:
        `(sheet, a1)` or `(sheet, start_a1, end_a1)` when evaluation succeeds,
        otherwise `None`.
    """
    formula = _formula_defined_name_body(attr_text)
    if not formula.upper().startswith(("OFFSET(", "INDIRECT(")):
        return None
    formula = "=" + formula
    if named_ranges or named_range_ranges:
        formula = expand_defined_names(
            formula,
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
        )
    bounds = _sheet_bounds(wb)
    formula = _normalize_formula_for_parse(formula, bounds)
    try:
        ast = parse_formula_ast(formula)
    except FormulaParseError:
        return None
    if not isinstance(ast, FunctionCallNode):
        return None

    def get_cell_value(addr: str) -> CellValue:
        try:
            sheet_part, a1 = addr.split("!", 1)
            sheet = sheet_part.strip("'")
            if sheet in wb.sheetnames:
                return wb[sheet][a1].value
        except Exception:
            pass
        return None

    if ast.name.upper() == "OFFSET":
        return _eval_offset_formula_to_range(ast, get_cell_value, bounds)
    if ast.name.upper() == "INDIRECT":
        return _eval_indirect_formula_to_range(ast, get_cell_value, bounds)
    return None


def build_named_range_map(wb: fastpyxl.Workbook) -> NamedRangeMaps:
    """Map defined names to single-cell and range references.

    Accepted referents (optionally quoted sheet name):

    - Single cells: `Sheet1!$A$1`
    - Rectangles: `Sheet1!$A$1:$B$10`
    - Whole-row spans: `Sheet1!$5:$134` / `Sheet1!$5:$5`
    - Whole-column spans: `Sheet1!$A:$C` / `Sheet1!$A:$A`
    - Evaluated `OFFSET` / `INDIRECT` formulas using workbook values, including
      `OFFSET` whose base is another already-resolved defined name

    Whole-row and whole-column names expand the **implicit** axis against that
    sheet's used range (`Worksheet.max_row` / `max_column`), the same policy as
    formula shorthands in `excel_grapher.core.range_shorthand`. Named rows or
    columns are kept; they are not clipped to the used extent. A missing sheet
    is omitted rather than expanded to Excel's full grid (`XFD` / `1048576`).

    Multi-area unions, `#REF!`, array constants, and other non-range formulas
    are skipped.
    """
    cell_map: dict[str, tuple[str, str]] = {}
    range_map: dict[str, tuple[str, str, str]] = {}
    bounds: dict[str, tuple[int, int]] | None = None
    pending_formulas: list[tuple[str, str]] = []
    for name, defn in wb.defined_names.items():
        attr_text = getattr(defn, "attr_text", None)
        if not isinstance(attr_text, str) or not attr_text:
            continue
        if attr_text.startswith("{") or attr_text.startswith("#") or attr_text.startswith('"'):
            continue
        key = str(name)
        if _is_offset_or_indirect_defined_name(attr_text):
            pending_formulas.append((key, attr_text))
            continue
        if "," in attr_text:
            continue
        if ":" in attr_text:
            m = _RECT_DEFINED_NAME_RE.match(attr_text)
            if m is not None:
                sheet_name = m.group("sheet")
                start = f"{m.group('c1')}{m.group('r1')}"
                end = f"{m.group('c2')}{m.group('r2')}"
                range_map[key] = (sheet_name, start, end)
                continue
            if _WHOLE_ROW_DEFINED_NAME_RE.match(attr_text) or _WHOLE_COL_DEFINED_NAME_RE.match(
                attr_text
            ):
                if bounds is None:
                    bounds = _sheet_bounds(wb)
                resolved_axis = _try_resolve_whole_axis_defined_name(attr_text, bounds)
                if resolved_axis is not None:
                    range_map[key] = resolved_axis
                    continue
            resolved = _try_resolve_formula_defined_name(
                attr_text, wb, named_ranges=cell_map, named_range_ranges=range_map
            )
            if resolved is not None:
                _store_resolved_defined_name(key, resolved, cell_map, range_map)
            continue

        m = _CELL_DEFINED_NAME_RE.match(attr_text)
        if m is None:
            continue
        sheet_name = m.group(1)
        col = m.group(2)
        row = m.group(3)
        cell_map[key] = (sheet_name, f"{col}{row}")

    remaining = pending_formulas
    while remaining:
        unresolved: list[tuple[str, str]] = []
        progress = False
        for key, attr_text in remaining:
            resolved = _try_resolve_formula_defined_name(
                attr_text, wb, named_ranges=cell_map, named_range_ranges=range_map
            )
            if resolved is None:
                unresolved.append((key, attr_text))
                continue
            _store_resolved_defined_name(key, resolved, cell_map, range_map)
            progress = True
        if not progress:
            break
        remaining = unresolved

    return NamedRangeMaps(cell_map=cell_map, range_map=range_map)
