from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass
from typing import Protocol

import fastpyxl
from fastpyxl.utils.cell import (
    column_index_from_string,
    coordinate_from_string,
    get_column_letter,
)

from excel_grapher.core.address_keys import needs_quoting
from excel_grapher.core.addressing import offset_range, split_sheet_qualified_address
from excel_grapher.core.coercions import to_number
from excel_grapher.core.formula_ast import (
    AstNode,
    CellRefNode,
    FormulaParseError,
    FunctionCallNode,
    NumberNode,
    RangeNode,
)
from excel_grapher.core.formula_ast import (
    parse as parse_formula_ast,
)
from excel_grapher.core.formula_normalization import expand_whole_column_row_for_parse
from excel_grapher.core.range_shorthand import sheet_used_extent
from excel_grapher.core.types import CellValue, ExcelRange, XlError

from .parser import CellRef

_NAME_TOKEN_RE = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\b(?!\s*!)")


class NameResolver(Protocol):
    def resolve(self, name: str) -> tuple[str, str] | None:
        """Return (sheet, A1) for a defined name, or None if unknown/unsupported."""


class DictNameResolver:
    def __init__(self, mapping: dict[str, tuple[str, str]]):
        self._mapping = mapping

    def resolve(self, name: str) -> tuple[str, str] | None:
        return self._mapping.get(name)


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


def _eval_number_for_defined_name(
    node: AstNode,
    get_cell_value: Callable[[str], CellValue],
    bounds: dict[str, tuple[int, int]],
) -> int | float | None:
    """Evaluate an AST node to a number for OFFSET args (rows, cols, height, width)."""
    if isinstance(node, NumberNode):
        return int(node.value) if node.value == int(node.value) else node.value
    if isinstance(node, CellRefNode):
        val = get_cell_value(node.address)
        n = to_number(val)
        if isinstance(n, XlError):
            return None
        return int(n) if n == int(n) else float(n)
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
    height = (
        _eval_number_for_defined_name(node.args[3], get_cell_value, bounds)
        if len(node.args) >= 4
        else None
    )
    width = (
        _eval_number_for_defined_name(node.args[4], get_cell_value, bounds)
        if len(node.args) >= 5
        else None
    )
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


def _quote_sheet_for_formula(sheet: str) -> str:
    if needs_quoting(sheet):
        return "'" + sheet.replace("'", "''") + "'"
    return sheet


def _format_cell_ref_for_formula(sheet: str, a1: str) -> str:
    col, row = coordinate_from_string(a1)
    return f"{_quote_sheet_for_formula(sheet)}!${col}${row}"


def _substitute_defined_names_in_attr_text(
    attr_text: str,
    cell_map: dict[str, tuple[str, str]],
    range_map: dict[str, tuple[str, str, str]],
) -> str:
    """Inline already-resolved defined names so OFFSET/INDIRECT bodies can parse."""

    def repl(match: re.Match[str]) -> str:
        token = match.group(1)
        if token in cell_map:
            sheet, a1 = cell_map[token]
            return _format_cell_ref_for_formula(sheet, a1)
        if token in range_map:
            sheet, start_a1, end_a1 = range_map[token]
            c1, r1 = coordinate_from_string(start_a1)
            c2, r2 = coordinate_from_string(end_a1)
            return f"{_quote_sheet_for_formula(sheet)}!${c1}${r1}:${c2}${r2}"
        return token

    return _NAME_TOKEN_RE.sub(repl, attr_text)


def _store_resolved_name(
    resolved: tuple[str, str] | tuple[str, str, str],
    *,
    cell_map: dict[str, tuple[str, str]],
    range_map: dict[str, tuple[str, str, str]],
    name: str,
) -> None:
    if len(resolved) == 2:
        cell_map[name] = (resolved[0], resolved[1])
    elif len(resolved) == 3:
        range_map[name] = resolved


def _resolve_static_defined_name(
    attr_text: str,
    bounds: dict[str, tuple[int, int]],
) -> tuple[str, str] | tuple[str, str, str] | None:
    if "," in attr_text:
        return None

    m = re.match(r"'?([^'!]+)'?!\$?([A-Z]{1,3})\$?(\d+)$", attr_text, re.IGNORECASE)
    if m:
        return (m.group(1), f"{m.group(2).upper()}{m.group(3)}")

    if ":" in attr_text:
        m = re.match(
            r"'?(?P<sheet>[^'!]+)'?!\$?(?P<c>[A-Z]{1,3}):\$?(?P=c)$",
            attr_text,
            re.IGNORECASE,
        )
        if m:
            sheet = m.group("sheet")
            col = m.group("c").upper()
            max_row, _ = sheet_used_extent(bounds, sheet)
            return (sheet, f"{col}1", f"{col}{max_row}")

        m = re.match(
            r"'?(?P<sheet>[^'!]+)'?!\$?(?P<r>\d+):\$?(?P=r)$",
            attr_text,
        )
        if m:
            sheet = m.group("sheet")
            row = int(m.group("r"))
            _, max_col = sheet_used_extent(bounds, sheet)
            return (
                sheet,
                f"A{row}",
                f"{get_column_letter(max_col)}{row}",
            )

        m = re.match(
            r"'?(?P<sheet>[^'!]+)'?!\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+):\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)$",
            attr_text,
            re.IGNORECASE,
        )
        if m:
            sheet_name = m.group("sheet")
            start = f"{m.group('c1').upper()}{m.group('r1')}"
            end = f"{m.group('c2').upper()}{m.group('r2')}"
            return (sheet_name, start, end)

    return None


def _normalize_formula_for_parse(formula: str, bounds: dict[str, tuple[int, int]]) -> str:
    """Strip ``$`` and expand whole-column/whole-row refs so formula_ast can parse."""
    return expand_whole_column_row_for_parse(formula, bounds)


def _try_resolve_formula_defined_name(
    attr_text: str,
    wb: fastpyxl.Workbook,
) -> tuple[str, str, str] | tuple[str, str] | None:
    """If attr_text is an OFFSET/INDIRECT formula, evaluate to range or cell; else None."""
    formula = attr_text.strip()
    if not formula.upper().startswith("OFFSET(") and not formula.upper().startswith("INDIRECT("):
        return None
    if not formula.startswith("="):
        formula = "=" + formula
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

    Only includes simple definitions like Sheet1!$A$1 or Sheet1!$A$1:$B$10
    (optionally quoted sheet name). Skips multi-area and complex formulas.
    Formula-based names (OFFSET, INDIRECT) are evaluated using workbook values.
    """
    bounds = _sheet_bounds(wb)
    cell_map: dict[str, tuple[str, str]] = {}
    range_map: dict[str, tuple[str, str, str]] = {}
    pending: dict[str, str] = {}
    for name, defn in wb.defined_names.items():
        attr_text = getattr(defn, "attr_text", None)
        if not isinstance(attr_text, str) or not attr_text:
            continue
        stripped = attr_text.strip()
        if stripped.startswith("{") or stripped.startswith("#") or stripped.startswith('"'):
            continue
        pending[str(name)] = stripped

    changed = True
    while changed and pending:
        changed = False
        for name in list(pending):
            attr_text = _substitute_defined_names_in_attr_text(pending[name], cell_map, range_map)
            resolved = _resolve_static_defined_name(attr_text, bounds)
            if resolved is None:
                resolved = _try_resolve_formula_defined_name(attr_text, wb)
            if resolved is None:
                continue
            _store_resolved_name(
                resolved,
                cell_map=cell_map,
                range_map=range_map,
                name=name,
            )
            del pending[name]
            changed = True

    return NamedRangeMaps(cell_map=cell_map, range_map=range_map)


def qualify_cell_ref(ref: CellRef, current_sheet: str) -> tuple[str, str]:
    sheet = ref.sheet if ref.sheet is not None else current_sheet
    return sheet, f"{ref.column}{ref.row}"
