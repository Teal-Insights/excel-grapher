"""Sheet-qualified address parsing and normalization helpers.

These helpers are the single source of truth for translating between the
external address strings that Excel users write (e.g. ``'Sheet1!A1'``,
``"'My Sheet'!B2"``) and the canonical ``NodeKey`` form stored in the
``DependencyGraph`` and emitted by generated code.

They live in :mod:`excel_grapher.core` so that both the grapher and the
evaluator/exporter can import them without violating layering rules.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Sequence

from fastpyxl.utils.cell import column_index_from_string, coordinate_from_string


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


def parse_address(address: str) -> tuple[str, str]:
    """Parse a sheet-qualified address into ``(sheet, cell_coord)``.

    The returned sheet name has any surrounding single quotes stripped and any
    escaped apostrophes (``''``) unescaped to a single apostrophe.

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


def format_key(sheet: str, cell: str) -> str:
    """Format a sheet and A1 cell coordinate into a canonical address string."""
    return f"{quote_sheet_if_needed(sheet)}!{cell}"


def format_cell_key(sheet: str, column: str, row: int) -> str:
    """Format a (sheet, column_letters, row) triple into a canonical address."""
    return f"{quote_sheet_if_needed(sheet)}!{column}{row}"


def normalize_key(key: str) -> str:
    """Normalize an address to the canonical ``NodeKey`` form.

    Unnecessary quoting is stripped; sheet names with spaces, hyphens, or
    apostrophes are quoted. The result matches ``Node.key`` exactly.

    Examples:
        >>> normalize_key("'Sheet1'!A1")
        'Sheet1!A1'
        >>> normalize_key("'My Sheet'!B2")
        "'My Sheet'!B2"
    """
    sheet, cell = parse_address(key)
    return format_key(sheet, cell)


def _normalize_cell_coord(cell: str) -> str:
    """Return an A1 coordinate without absolute markers."""
    col, row = coordinate_from_string(cell.replace("$", ""))
    return f"{col}{int(row)}"


def normalize_range_key(range_key: str, *, current_sheet: str | None = None) -> str:
    """Normalize a single-sheet A1 range to canonical ``Sheet!A1:Sheet!B2`` form.

    Accepted forms include local ranges (when ``current_sheet`` is provided),
    sheet-qualified ranges with one or both endpoints qualified, and optional ``$``
    absolute markers.
    """

    def _resolve_endpoint(endpoint: str, *, default_sheet: str | None) -> tuple[str, str]:
        token = endpoint.strip()
        if not token:
            raise ValueError(f"Invalid range endpoint: {endpoint!r}")

        if "!" in token or token.startswith("'"):
            sheet, cell = parse_address(token)
        else:
            if default_sheet is None:
                raise ValueError(f"Range endpoint must be sheet-qualified: {endpoint!r}")
            sheet, cell = default_sheet, token
        return sheet, _normalize_cell_coord(cell)

    token = range_key.strip()
    if not token:
        raise ValueError("Range key must not be empty")

    if ":" not in token:
        sheet, cell = _resolve_endpoint(token, default_sheet=current_sheet)
        return format_key(sheet, cell)

    start_raw, end_raw = token.split(":", 1)
    start_sheet, start_cell = _resolve_endpoint(start_raw, default_sheet=current_sheet)
    end_sheet, end_cell = _resolve_endpoint(end_raw, default_sheet=start_sheet)
    if end_sheet != start_sheet:
        raise ValueError(
            f"Range endpoints must be on the same sheet: {start_sheet!r} vs {end_sheet!r}"
        )
    return f"{format_key(start_sheet, start_cell)}:{format_key(end_sheet, end_cell)}"


def make_node_key_sort_key(
    sheet_order: Sequence[str],
) -> Callable[[str], tuple[int, str, int, int]]:
    """Build a key function for workbook-aligned ``NodeKey`` sorting.

    Keys are ordered by:
    1) workbook sheet order (from ``sheet_order``),
    2) row number,
    3) column number.

    Sheets not present in ``sheet_order`` are placed after known sheets and
    sorted by sheet name.
    """
    sheet_rank = {name: idx for idx, name in enumerate(sheet_order)}
    fallback_rank = len(sheet_rank)

    def _sort_key(node_key: str) -> tuple[int, str, int, int]:
        sheet, cell = parse_address(node_key)
        col_letters, row = coordinate_from_string(cell.replace("$", ""))
        col = int(column_index_from_string(col_letters))
        return (sheet_rank.get(sheet, fallback_rank), sheet, int(row), col)

    return _sort_key


def sort_node_keys(node_keys: Iterable[str], *, sheet_order: Sequence[str]) -> list[str]:
    """Return ``node_keys`` sorted by workbook sheet order, then row, then column."""
    return sorted(node_keys, key=make_node_key_sort_key(sheet_order))
