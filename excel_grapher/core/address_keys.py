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
from dataclasses import dataclass
from enum import StrEnum
from typing import TypeAlias

from fastpyxl.utils.cell import (
    column_index_from_string,
    coordinate_from_string,
    get_column_letter,
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


@dataclass(frozen=True, slots=True)
class ParsedRowKey:
    """Parsed one-row node key (`Sheet1!D63:Y63`).

    Columns are ordered so `min_col` <= `max_col` by index. `row` is the single
    worksheet row for both endpoints.
    """

    sheet: str
    row: int
    min_col: str
    max_col: str


def _parse_a1_cell(coord: str) -> tuple[str, int]:
    """Parse an A1 coordinate into `(column_letters, row)`, stripping `$`."""
    try:
        col, row = coordinate_from_string(coord.replace("$", ""))
    except CellCoordinatesException as exc:
        raise ValueError(f"Invalid A1 cell coordinate: {coord!r}") from exc
    return str(col).upper(), int(row)


def _ordered_columns(col_a: str, col_b: str) -> tuple[str, str]:
    """Return `(min_col, max_col)` ordered by column index."""
    a = column_index_from_string(col_a)
    b = column_index_from_string(col_b)
    if a <= b:
        return col_a, col_b
    return col_b, col_a


def format_row_key(sheet: str, min_col: str, row: int, max_col: str) -> NormalizedAddress:
    """Format a one-row span as a canonical row-node key.

    Columns are ordered so the left endpoint is the lesser column index
    (e.g. `Y, D` becomes `Sheet1!D63:Y63`).
    """
    left, right = _ordered_columns(min_col.upper(), max_col.upper())
    return f"{quote_sheet_if_needed(sheet)}!{left}{row}:{right}{row}"


def parse_row_key(key: str) -> ParsedRowKey:
    """Parse a sheet-qualified one-row key into sheet, row, and column span.

    Accepts `Sheet1!D63:Y63`, both-end forms (`Sheet1!D63:Sheet1!Y63`), quoted
    sheets, and absolute markers (`$D$63`). Raises `ValueError` for cell-only
    keys, multi-row extents, or cross-sheet ranges.
    """
    parts = split_address_on_colon(key)
    if parts is None:
        raise ValueError(f"Row key must be a one-row range (got cell-only key): {key!r}")

    start_raw, end_raw = parts
    start_sheet, start_coord = parse_address(start_raw)

    if "!" in end_raw or end_raw.startswith("'"):
        end_sheet, end_coord = parse_address(end_raw)
        if end_sheet != start_sheet:
            raise ValueError(
                f"Row key endpoints must share the same sheet: {key!r} "
                f"({start_sheet!r} vs {end_sheet!r})"
            )
    else:
        end_coord = end_raw

    start_col, start_row = _parse_a1_cell(start_coord)
    end_col, end_row = _parse_a1_cell(end_coord)
    if start_row != end_row:
        raise ValueError(f"Row key must be a one-row extent (same row on both ends): {key!r}")

    min_col, max_col = _ordered_columns(start_col, end_col)
    return ParsedRowKey(
        sheet=start_sheet,
        row=start_row,
        min_col=min_col,
        max_col=max_col,
    )


def normalize_row_key(key: str) -> NormalizedAddress:
    """Normalize a one-row key to canonical :data:`NormalizedAddress` form.

    Unnecessary sheet quoting is stripped, columns are ordered, absolute markers
    are removed, and both-end sheet qualification collapses to a single prefix.
    """
    parsed = parse_row_key(key)
    return format_row_key(parsed.sheet, parsed.min_col, parsed.row, parsed.max_col)


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
    One-row range keys are canonicalized via `normalize_row_key` (ordered
    columns, including 1x1 forms such as `Sheet1!D63:D63`). Multi-area keys
    use `parse_node_key`. Other ranges collapse via single-prefix
    `format_range_key` (same-sheet) or keep both endpoints when sheets differ.

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
    """
    if "," in key and len(_split_top_level_comma(key)) > 1:
        return str(parse_node_key(key))

    parts = split_address_on_colon(key)
    if parts is not None:
        try:
            return normalize_row_key(key)
        except ValueError:
            pass
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
# Node key model (CellKey / RangeKey / UnionKey) — formula-group Issue 1 sprint 1
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


def members_to_node_key(members: Sequence[str | NodeKey]) -> NodeKey:
    """Build a canonical node key that exactly covers the given cell members.

    Uses sheet partition, sort by `(row, col)`, greedy horizontal runs, then
    greedy vertical merge of equal-width runs. Same member set always yields
    the same key regardless of input order.
    """
    if not members:
        raise ValueError("Cannot build node key from empty member set")

    cells: dict[tuple[str, int, int], tuple[str, str, int]] = {}
    for raw in members:
        parsed = parse_node_key(raw)
        if not isinstance(parsed, CellKey):
            raise ValueError(
                f"members_to_node_key requires cell members; got {type(parsed).__name__}: {raw!r}"
            )
        col_i = int(column_index_from_string(parsed.column))
        cells[(parsed.sheet, parsed.row, col_i)] = (parsed.sheet, parsed.column, parsed.row)

    if not cells:
        raise ValueError("Cannot build node key from empty member set")

    by_sheet: dict[str, list[tuple[int, int, str]]] = {}
    for sheet, row, col_i in cells:
        _sheet, col_letters, _row = cells[(sheet, row, col_i)]
        by_sheet.setdefault(sheet, []).append((row, col_i, col_letters))

    areas: list[CellKey | RangeKey] = []
    for sheet, items in by_sheet.items():
        items.sort(key=lambda t: (t[0], t[1]))
        # Horizontal runs: (row, min_col_i, max_col_i, min_col_letters, max_col_letters)
        runs: list[tuple[int, int, int]] = []
        for row, col_i, _col_letters in items:
            if runs and runs[-1][0] == row and col_i == runs[-1][2] + 1:
                prev = runs[-1]
                runs[-1] = (prev[0], prev[1], col_i)
            else:
                runs.append((row, col_i, col_i))

        # Vertical merge: (min_col_i, max_col_i, min_row, max_row)
        rects: list[tuple[int, int, int, int]] = [
            (min_c, max_c, row, row) for row, min_c, max_c in runs
        ]
        rects.sort(key=lambda r: (r[0], r[1], r[2]))
        merged: list[tuple[int, int, int, int]] = []
        for min_c, max_c, min_r, max_r in rects:
            if (
                merged
                and merged[-1][0] == min_c
                and merged[-1][1] == max_c
                and min_r == merged[-1][3] + 1
            ):
                prev = merged[-1]
                merged[-1] = (prev[0], prev[1], prev[2], max_r)
            else:
                merged.append((min_c, max_c, min_r, max_r))

        for min_c, max_c, min_r, max_r in merged:
            areas.append(
                _make_range_or_cell_key(
                    sheet,
                    get_column_letter(min_c),
                    min_r,
                    get_column_letter(max_c),
                    max_r,
                )
            )

    return _make_union_key(areas)


def expand_node_cells(key: str | NodeKey) -> tuple[CellKey, ...]:
    """Expand a cell, range, or union key into canonical member `CellKey`s.

    Order is deterministic: sheet order as in the key, then row, then column.
    """
    parsed = parse_node_key(key)
    if isinstance(parsed, CellKey):
        return (parsed,)

    cells: list[CellKey] = []
    areas: Sequence[CellKey | RangeKey] = (
        (parsed,) if isinstance(parsed, RangeKey) else parsed.members
    )

    for area in areas:
        if isinstance(area, CellKey):
            cells.append(area)
            continue
        min_c = column_index_from_string(area.min_col)
        max_c = column_index_from_string(area.max_col)
        for row in range(area.min_row, area.max_row + 1):
            for col_i in range(min_c, max_c + 1):
                cells.append(_make_cell_key(area.sheet, get_column_letter(col_i), row))
    return tuple(cells)
