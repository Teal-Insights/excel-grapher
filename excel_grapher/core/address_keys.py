"""Sheet-qualified address parsing and normalization helpers.

These helpers are the single source of truth for translating between the
external address strings that Excel users write (e.g. `'Sheet1!A1'`,
`"'My Sheet'!B2"`) and the canonical :data:`NormalizedAddress` form stored in the
`DependencyGraph` and emitted by generated code.

They live in `excel_grapher.core` so that both the grapher and the
evaluator/exporter can import them without violating layering rules.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Iterable, Sequence
from enum import StrEnum
from typing import TypeAlias

from fastpyxl.utils.cell import (
    column_index_from_string,
    coordinate_from_string,
)
from fastpyxl.utils.exceptions import CellCoordinatesException

# Canonical sheet-qualified cell (`Sheet1!B1`) or same-sheet range (`Sheet1!C4:D4`).
NormalizedAddress: TypeAlias = str

_A1_CELL_COORD_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
_WHOLE_COL_COORD_RE = re.compile(r"^\$?([A-Za-z]{1,3})$")
_WHOLE_ROW_COORD_RE = re.compile(r"^\$?(\d+)$")


def needs_quoting(sheet: str) -> bool:
    """Return True if a sheet name must be wrapped in single quotes in a formula."""
    return " " in sheet or "-" in sheet or "'" in sheet


def _escape_sheet_for_formula(sheet: str) -> str:
    """Escape apostrophes for use inside quoted sheet names."""
    return sheet.replace("'", "''")


def quote_sheet_if_needed(sheet: str) -> str:
    """Return a sheet name quoted for formulas when quoting is required."""
    if not needs_quoting(sheet):
        return sheet
    return "'" + _escape_sheet_for_formula(sheet) + "'"


# Regex fragment for the inner part of a quoted Excel sheet name in formulas.
_QUOTED_SHEET_NAME_INNER = r"(?:[^']|'')+"


def quoted_sheet_name_regex(*, capture_group: str = "sheet") -> str:
    """Return a regex fragment matching a quoted Excel sheet name (no trailing ``!``)."""
    return rf"'(?P<{capture_group}>{_QUOTED_SHEET_NAME_INNER})'"


def quoted_sheet_prefix_regex(*, capture_group: str = "sheet") -> str:
    """Return a regex fragment matching ``'Sheet Name'!`` with Excel ``''`` escape."""
    return quoted_sheet_name_regex(capture_group=capture_group) + "!"


def unescape_formula_sheet_name(escaped: str) -> str:
    """Unescape a sheet name captured from a quoted formula reference."""
    return escaped.replace("''", "'")


def parse_address(address: str) -> tuple[str, str]:
    """Parse a sheet-qualified address into `(sheet, cell_coord)`.

    The returned sheet name has any surrounding single quotes stripped and any
    escaped apostrophes (`''`) unescaped to a single apostrophe.

    Examples:
        >>> parse_address("Sheet1!A1")
        ('Sheet1', 'A1')
        >>> parse_address("'My Sheet'!B2")
        ('My Sheet', 'B2')
        >>> parse_address("'It''s Data'!C3")
        ("It's Data", 'C3')
    """
    if address.startswith("'"):
        i = 1
        while i < len(address):
            if address[i] == "'":
                if i + 1 < len(address) and address[i + 1] == "'":
                    i += 2
                    continue
                break
            i += 1
        sheet = address[1:i].replace("''", "'")
        rest = address[i + 1 :]
        if rest.startswith("!"):
            return sheet, rest[1:]
        raise ValueError(f"Invalid address format: {address}")

    if "!" in address:
        sheet, cell = address.rsplit("!", 1)
        return sheet, cell

    raise ValueError(f"Address must be sheet-qualified: {address}")


def format_key(sheet: str, cell: str) -> NormalizedAddress:
    """Format a sheet and A1 cell coordinate into a canonical address string."""
    return f"{quote_sheet_if_needed(sheet)}!{cell}"


def format_cell_key(sheet: str, column: str, row: int) -> NormalizedAddress:
    """Format a (sheet, column_letters, row) triple into a canonical address."""
    return f"{quote_sheet_if_needed(sheet)}!{column}{row}"


def format_range_key(sheet: str, start_cell: str, end_cell: str) -> NormalizedAddress:
    """Format a same-sheet range as a single-prefix canonical address.

    Examples:
        >>> format_range_key("Sheet1", "A1", "A3")
        'Sheet1!A1:A3'
        >>> format_range_key("My Sheet", "A1", "B2")
        "'My Sheet'!A1:B2"
    """
    return f"{quote_sheet_if_needed(sheet)}!{start_cell}:{end_cell}"


def canonical_cell_coord(cell: str) -> str:
    """Canonicalize an A1 / whole-column / whole-row coordinate fragment.

    Strips `$` markers, uppercases column letters, and normalizes row numbers
    (`01` -> `1`). Non-matching fragments are returned unchanged.
    """
    m = _A1_CELL_COORD_RE.fullmatch(cell)
    if m is not None:
        return f"{m.group(1).upper()}{int(m.group(2))}"
    m_col = _WHOLE_COL_COORD_RE.fullmatch(cell)
    if m_col is not None:
        return m_col.group(1).upper()
    m_row = _WHOLE_ROW_COORD_RE.fullmatch(cell)
    if m_row is not None:
        return str(int(m_row.group(1)))
    return cell


def _parse_a1_cell(cell: str) -> tuple[str, int]:
    """Parse an A1 cell coordinate into uppercase column letters and row."""
    try:
        column, row = coordinate_from_string(canonical_cell_coord(cell))
    except CellCoordinatesException as exc:
        raise ValueError(f"Expected A1 cell coordinate, got: {cell!r}") from exc
    return str(column).upper(), int(row)


def parse_cell_coords(address: str) -> tuple[str, int, int]:
    """Parse a sheet-qualified A1 cell into `(sheet, row, col)` (1-based).

    Raises:
        ValueError: If `address` is not a sheet-qualified single cell.
    """
    sheet, cell = parse_address(address)
    col_letters, row = _parse_a1_cell(cell)
    return sheet, row, int(column_index_from_string(col_letters))


def _ordered_columns(left: str, right: str) -> tuple[str, str]:
    """Return two column letters ordered from left to right."""
    left_u = left.upper()
    right_u = right.upper()
    if column_index_from_string(left_u) <= column_index_from_string(right_u):
        return left_u, right_u
    return right_u, left_u


def split_address_on_colon(address: str) -> tuple[str, str] | None:
    """Split an address on the first colon outside quoted sheet names.

    Handles colons embedded in quoted sheet names (`'A:B'!C1:D2`).
    Returns `None` when there is no top-level colon.
    """
    in_quote = False
    i = 0
    while i < len(address):
        ch = address[i]
        if ch == "'":
            if in_quote and i + 1 < len(address) and address[i + 1] == "'":
                i += 2
                continue
            in_quote = not in_quote
        elif ch == ":" and not in_quote:
            return address[:i], address[i + 1 :]
        i += 1
    return None


def _split_top_level_comma(address: str) -> list[str]:
    """Split an address on commas outside quoted sheet names."""
    parts: list[str] = []
    start = 0
    in_quote = False
    i = 0
    while i < len(address):
        ch = address[i]
        if ch == "'":
            if in_quote and i + 1 < len(address) and address[i + 1] == "'":
                i += 2
                continue
            in_quote = not in_quote
        elif ch == "," and not in_quote:
            parts.append(address[start:i])
            start = i + 1
        i += 1
    parts.append(address[start:])
    return parts


def normalize_key(key: str) -> NormalizedAddress:
    """Normalize an address to canonical :data:`NormalizedAddress` form.

    Unnecessary quoting is stripped; sheet names with spaces, hyphens, or
    apostrophes are quoted. Absolute markers (`$`) are stripped and column
    letters are uppercased. For single cells, the result matches `Node.key`.
    Ranges and unions use `parse_node_key` (1x1 ranges collapse to cells).
    Non-parseable colon forms fall back to `format_range_key` for same-sheet
    ranges or keep both sheet-qualified endpoints when sheets differ.

    Examples:
        >>> normalize_key("'Sheet1'!A1")
        'Sheet1!A1'
        >>> normalize_key("'My Sheet'!B2")
        "'My Sheet'!B2"
        >>> normalize_key("Sheet1!Y63:D63")
        'Sheet1!D63:Y63'
        >>> normalize_key("Sheet1!A1:Sheet1!B2")
        'Sheet1!A1:B2'
        >>> normalize_key("Sheet1!A1:$A$3")
        'Sheet1!A1:A3'
        >>> normalize_key("Sheet1!E5,A1:D1")
        'Sheet1!A1:D1,E5'
        >>> normalize_key("Sheet1!D63:D63")
        'Sheet1!D63'
    """
    if "," in key and len(_split_top_level_comma(key)) > 1:
        return str(parse_node_key(key))

    parts = split_address_on_colon(key)
    if parts is not None:
        try:
            return str(parse_node_key(key))
        except ValueError:
            pass
        start_raw, end_raw = parts
        start_sheet, start_cell = parse_address(start_raw)
        start_cell = canonical_cell_coord(start_cell)
        if "!" in end_raw or end_raw.startswith("'"):
            end_sheet, end_cell = parse_address(end_raw)
            end_cell = canonical_cell_coord(end_cell)
            if end_sheet == start_sheet:
                return format_range_key(start_sheet, start_cell, end_cell)
            start_fmt = format_key(start_sheet, start_cell)
            end_fmt = format_key(end_sheet, end_cell)
            return f"{start_fmt}:{end_fmt}"
        return format_range_key(start_sheet, start_cell, canonical_cell_coord(end_raw))

    try:
        return str(parse_node_key(key))
    except ValueError:
        sheet, cell = parse_address(key)
        return format_key(sheet, canonical_cell_coord(cell))


def make_node_key_sort_key(
    sheet_order: Sequence[str],
) -> Callable[[NormalizedAddress], tuple[int, str, int, int]]:
    """Build a key function for workbook-aligned `NodeKey` sorting.

    Keys are ordered by:
    1) workbook sheet order (from `sheet_order`),
    2) top-left row (cell row, or range/union min_row),
    3) top-left column (cell column, or range/union min_col).

    Sheets not present in `sheet_order` are placed after known sheets and
    sorted by sheet name. Cross-sheet unions sort by their first canonical
    member's sheet.
    """
    sheet_rank = {name: idx for idx, name in enumerate(sheet_order)}
    fallback_rank = len(sheet_rank)

    def _anchor(node_key: NormalizedAddress) -> tuple[str, int, int]:
        try:
            parsed = parse_node_key(node_key)
        except ValueError:
            # Non-canonical junk (e.g. whole-column refs): sort late, stably.
            return ("\uffff", 10**9, 10**9)

        if isinstance(parsed, CellKey):
            return (
                parsed.sheet,
                parsed.row,
                int(column_index_from_string(parsed.column)),
            )
        if isinstance(parsed, RangeKey):
            return (
                parsed.sheet,
                parsed.min_row,
                int(column_index_from_string(parsed.min_col)),
            )
        # UnionKey — first member after canonical sort
        first = parsed.members[0]
        if isinstance(first, CellKey):
            return (
                first.sheet,
                first.row,
                int(column_index_from_string(first.column)),
            )
        return (
            first.sheet,
            first.min_row,
            int(column_index_from_string(first.min_col)),
        )

    def _sort_key(node_key: NormalizedAddress) -> tuple[int, str, int, int]:
        sheet, row, col = _anchor(node_key)
        return (sheet_rank.get(sheet, fallback_rank), sheet, row, col)

    return _sort_key


def sort_node_keys(
    node_keys: Iterable[NormalizedAddress], *, sheet_order: Sequence[str]
) -> list[NormalizedAddress]:
    """Return `node_keys` sorted by workbook sheet order, then row, then column."""
    return sorted(node_keys, key=make_node_key_sort_key(sheet_order))


# ---------------------------------------------------------------------------
# Address key model (CellKey / RangeKey / UnionKey)
# ---------------------------------------------------------------------------


class NodeShape(StrEnum):
    """Geometry kind inferred from a canonical node key."""

    cell = "cell"
    row = "row"
    column = "column"
    range = "range"
    union = "union"


class CellKey(str):
    """Canonical sheet-qualified single-cell key (`Sheet1!E63`)."""

    __slots__ = ()

    @property
    def shape(self) -> NodeShape:
        return NodeShape.cell

    @property
    def sheet(self) -> str:
        return parse_address(self)[0]

    @property
    def column(self) -> str:
        _sheet, cell = parse_address(self)
        col, _row = _parse_a1_cell(cell)
        return col

    @property
    def row(self) -> int:
        _sheet, cell = parse_address(self)
        _col, row = _parse_a1_cell(cell)
        return row


class RangeKey(str):
    """Canonical sheet-qualified rectangle (`Sheet1!D63:Y63`, `Sheet1!E4:I18`)."""

    __slots__ = ()

    @property
    def shape(self) -> NodeShape:
        if self.min_row == self.max_row:
            return NodeShape.row
        if self.min_col == self.max_col:
            return NodeShape.column
        return NodeShape.range

    @property
    def sheet(self) -> str:
        return self._corners()[0]

    @property
    def min_col(self) -> str:
        return self._corners()[1]

    @property
    def max_col(self) -> str:
        return self._corners()[3]

    @property
    def min_row(self) -> int:
        return self._corners()[2]

    @property
    def max_row(self) -> int:
        return self._corners()[4]

    @property
    def column(self) -> str | None:
        if self.min_col == self.max_col:
            return self.min_col
        return None

    @property
    def row(self) -> int | None:
        if self.min_row == self.max_row:
            return self.min_row
        return None

    def _corners(self) -> tuple[str, str, int, str, int]:
        parts = split_address_on_colon(self)
        if parts is None:
            raise ValueError(f"RangeKey requires a colon range: {self!r}")
        start_raw, end_raw = parts
        sheet, start_coord = parse_address(start_raw)
        if "!" in end_raw or end_raw.startswith("'"):
            end_sheet, end_coord = parse_address(end_raw)
            if end_sheet != sheet:
                raise ValueError(f"Range endpoints must share the same sheet: {self!r}")
        else:
            end_coord = end_raw
        start_col, start_row = _parse_a1_cell(start_coord)
        end_col, end_row = _parse_a1_cell(end_coord)
        min_col, max_col = _ordered_columns(start_col, end_col)
        min_row, max_row = sorted((start_row, end_row))
        return sheet, min_col, min_row, max_col, max_row


class UnionKey(str):
    """Canonical multi-area key (`Sheet1!A1:D1,E5` or `Sheet1!A1,Sheet2!B2`)."""

    __slots__ = ()

    @property
    def shape(self) -> NodeShape:
        return NodeShape.union

    @property
    def members(self) -> tuple[CellKey | RangeKey, ...]:
        return tuple(_parse_union_member_areas(self))


NodeKey: TypeAlias = CellKey | RangeKey | UnionKey


def _format_bare_area(area: CellKey | RangeKey) -> str:
    """Format an area without the leading `Sheet!` prefix."""
    if isinstance(area, CellKey):
        return f"{area.column}{area.row}"
    if area.min_col == area.max_col and area.min_row == area.max_row:
        return f"{area.min_col}{area.min_row}"
    return f"{area.min_col}{area.min_row}:{area.max_col}{area.max_row}"


def _area_sort_key(area: CellKey | RangeKey) -> tuple[str, int, int, int, int]:
    if isinstance(area, CellKey):
        col = int(column_index_from_string(area.column))
        return (area.sheet, area.row, col, area.row, col)
    return (
        area.sheet,
        area.min_row,
        int(column_index_from_string(area.min_col)),
        area.max_row,
        int(column_index_from_string(area.max_col)),
    )


def _make_cell_key(sheet: str, column: str, row: int) -> CellKey:
    return CellKey(format_cell_key(sheet, column.upper(), int(row)))


def _make_range_or_cell_key(
    sheet: str, min_col: str, min_row: int, max_col: str, max_row: int
) -> CellKey | RangeKey:
    min_col_u, max_col_u = _ordered_columns(min_col.upper(), max_col.upper())
    lo_row, hi_row = sorted((int(min_row), int(max_row)))
    if min_col_u == max_col_u and lo_row == hi_row:
        return _make_cell_key(sheet, min_col_u, lo_row)
    text = f"{quote_sheet_if_needed(sheet)}!{min_col_u}{lo_row}:{max_col_u}{hi_row}"
    return RangeKey(text)


def _make_union_key(areas: Sequence[CellKey | RangeKey]) -> NodeKey:
    if not areas:
        raise ValueError("Union key cannot be empty")
    unique: dict[str, CellKey | RangeKey] = {}
    for area in areas:
        unique[str(area)] = area
    ordered = sorted(unique.values(), key=_area_sort_key)
    if len(ordered) == 1:
        return ordered[0]
    sheets = {a.sheet for a in ordered}
    if len(sheets) == 1:
        sheet = next(iter(sheets))
        bare = ",".join(_format_bare_area(a) for a in ordered)
        return UnionKey(f"{quote_sheet_if_needed(sheet)}!{bare}")
    return UnionKey(",".join(str(a) for a in ordered))


def _parse_single_area(raw: str, *, default_sheet: str | None = None) -> CellKey | RangeKey:
    """Parse one cell or range area into a canonical `CellKey` or `RangeKey`."""
    text = raw.strip()
    if not text:
        raise ValueError("Empty address area")

    parts = split_address_on_colon(text)
    if parts is None:
        if "!" in text or text.startswith("'"):
            sheet, cell = parse_address(text)
        else:
            if default_sheet is None:
                raise ValueError(f"Address must be sheet-qualified: {text!r}")
            sheet, cell = default_sheet, text
        col, row = _parse_a1_cell(cell)
        return _make_cell_key(sheet, col, row)

    start_raw, end_raw = parts
    if "!" in start_raw or start_raw.startswith("'"):
        start_sheet, start_coord = parse_address(start_raw)
    else:
        if default_sheet is None:
            raise ValueError(f"Address must be sheet-qualified: {text!r}")
        start_sheet, start_coord = default_sheet, start_raw

    if "!" in end_raw or end_raw.startswith("'"):
        end_sheet, end_coord = parse_address(end_raw)
        if end_sheet != start_sheet:
            raise ValueError(
                f"Range endpoints must share the same sheet: {text!r} "
                f"({start_sheet!r} vs {end_sheet!r})"
            )
    else:
        end_coord = end_raw

    start_col, start_row = _parse_a1_cell(start_coord)
    end_col, end_row = _parse_a1_cell(end_coord)
    return _make_range_or_cell_key(start_sheet, start_col, start_row, end_col, end_row)


def _parse_union_member_areas(key: str) -> list[CellKey | RangeKey]:
    chunks = [p.strip() for p in _split_top_level_comma(key)]
    if not chunks or all(c == "" for c in chunks):
        raise ValueError(f"Empty union key: {key!r}")
    if any(c == "" for c in chunks):
        raise ValueError(f"Empty union area in key: {key!r}")

    areas: list[CellKey | RangeKey] = []
    inherited_sheet: str | None = None
    for chunk in chunks:
        area = _parse_single_area(chunk, default_sheet=inherited_sheet)
        inherited_sheet = area.sheet
        areas.append(area)
    return areas


def parse_node_key(value: str | NodeKey) -> NodeKey:
    """Canonicalize `value` into a `CellKey`, `RangeKey`, or `UnionKey`.

    Strips `$`, orders range corners, sorts/dedupes union members, collapses
    one-member unions, and formats same-sheet unions with a single sheet prefix.
    """
    text = str(value) if isinstance(value, (CellKey, RangeKey, UnionKey)) else str(value).strip()
    if not text:
        raise ValueError("Empty node key")

    chunks = [p.strip() for p in _split_top_level_comma(text)]
    if len(chunks) > 1:
        return _make_union_key(_parse_union_member_areas(text))

    area_raw = chunks[0]
    if not area_raw or area_raw.endswith("!"):
        raise ValueError(f"Empty union key: {text!r}")
    return _parse_single_area(area_raw)
