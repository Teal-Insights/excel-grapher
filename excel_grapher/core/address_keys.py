"""Sheet-qualified address parsing and normalization helpers.

These helpers are the single source of truth for translating between the
external address strings that Excel users write (e.g. `'Sheet1!A1'`,
`"'My Sheet'!B2"`) and the canonical :data:`NormalizedAddress` form stored in the
`DependencyGraph` and emitted by generated code.

They live in `excel_grapher.core` so that both the grapher and the
evaluator/exporter can import them without violating layering rules.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Sequence
from dataclasses import dataclass
from typing import TypeAlias

from fastpyxl.utils.cell import column_index_from_string, coordinate_from_string

# Canonical sheet-qualified cell (`Sheet1!B1`) or range (`Sheet1!C4:Sheet1!D4`).
NormalizedAddress: TypeAlias = str


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


def _split_top_level_colon(address: str) -> tuple[str, str] | None:
    """Split an address on the first colon outside quoted sheet names."""
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


def _parse_a1_cell(coord: str) -> tuple[str, int]:
    """Parse an A1 coordinate into `(column_letters, row)`, stripping `$`."""
    col, row = coordinate_from_string(coord.replace("$", ""))
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
    parts = _split_top_level_colon(key)
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


def normalize_key(key: str) -> NormalizedAddress:
    """Normalize an address to canonical :data:`NormalizedAddress` form.

    Unnecessary quoting is stripped; sheet names with spaces, hyphens, or
    apostrophes are quoted. For single cells, the result matches `Node.key`.

    Examples:
        >>> normalize_key("'Sheet1'!A1")
        'Sheet1!A1'
        >>> normalize_key("'My Sheet'!B2")
        "'My Sheet'!B2"
    """
    sheet, cell = parse_address(key)
    return format_key(sheet, cell)


def make_node_key_sort_key(
    sheet_order: Sequence[str],
) -> Callable[[NormalizedAddress], tuple[int, str, int, int]]:
    """Build a key function for workbook-aligned `NodeKey` sorting.

    Keys are ordered by:
    1) workbook sheet order (from `sheet_order`),
    2) row number,
    3) column number.

    Sheets not present in `sheet_order` are placed after known sheets and
    sorted by sheet name.
    """
    sheet_rank = {name: idx for idx, name in enumerate(sheet_order)}
    fallback_rank = len(sheet_rank)

    def _sort_key(node_key: NormalizedAddress) -> tuple[int, str, int, int]:
        sheet, cell = parse_address(node_key)
        col_letters, row = coordinate_from_string(cell.replace("$", ""))
        col = int(column_index_from_string(col_letters))
        return (sheet_rank.get(sheet, fallback_rank), sheet, int(row), col)

    return _sort_key


def sort_node_keys(
    node_keys: Iterable[NormalizedAddress], *, sheet_order: Sequence[str]
) -> list[NormalizedAddress]:
    """Return `node_keys` sorted by workbook sheet order, then row, then column."""
    return sorted(node_keys, key=make_node_key_sort_key(sheet_order))
