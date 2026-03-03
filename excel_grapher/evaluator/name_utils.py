"""Utilities for converting Excel names to Python identifiers."""

import re


def parse_address(address: str) -> tuple[str, str]:
    """Parse a sheet-qualified address into (sheet_name, cell) tuple.

    The sheet name is returned WITHOUT surrounding quotes, with any
    escaped quotes ('') converted to single quotes (').

    Args:
        address: Sheet-qualified cell address (e.g., "'My Sheet'!A1" or "Sheet1!A1")

    Returns:
        Tuple of (sheet_name, cell_coordinate).

    Raises:
        ValueError: If address is not sheet-qualified or has invalid format.

    Examples:
        >>> parse_address("Sheet1!A1")
        ('Sheet1', 'A1')
        >>> parse_address("'My Sheet'!B2")
        ('My Sheet', 'B2')
        >>> parse_address("'It''s Data'!C3")
        ("It's Data", 'C3')
    """
    if address.startswith("'"):
        # Find the closing quote (handling escaped quotes)
        i = 1
        while i < len(address):
            if address[i] == "'":
                if i + 1 < len(address) and address[i + 1] == "'":
                    i += 2  # Skip escaped quote
                    continue
                break
            i += 1
        sheet = address[1:i]  # Strip surrounding quotes
        # Unescape internal quotes
        sheet = sheet.replace("''", "'")
        rest = address[i + 1 :]
        if rest.startswith("!"):
            return sheet, rest[1:]
        raise ValueError(f"Invalid address format: {address}")

    if "!" in address:
        sheet, cell = address.rsplit("!", 1)
        return sheet, cell

    raise ValueError(f"Address must be sheet-qualified: {address}")


def quote_sheet_if_needed(sheet: str) -> str:
    """Add quotes around a sheet name if it contains special characters.

    Quotes are added if the sheet name contains spaces, hyphens, or apostrophes.
    This matches the behavior of excel_grapher's Node.key property.

    Args:
        sheet: Sheet name (without quotes).

    Returns:
        Sheet name, quoted if necessary.

    Examples:
        >>> quote_sheet_if_needed("Sheet1")
        'Sheet1'
        >>> quote_sheet_if_needed("My Sheet")
        "'My Sheet'"
        >>> quote_sheet_if_needed("It's Data")
        "'It's Data'"
    """
    if " " in sheet or "-" in sheet or "'" in sheet:
        return f"'{sheet}'"
    return sheet


def format_address(sheet: str, cell: str) -> str:
    """Format a sheet name and cell coordinate into a normalized address.

    The sheet name will be quoted if it contains special characters.

    Args:
        sheet: Sheet name (without quotes).
        cell: Cell coordinate (e.g., "A1").

    Returns:
        Normalized sheet-qualified address.

    Examples:
        >>> format_address("Sheet1", "A1")
        'Sheet1!A1'
        >>> format_address("My Sheet", "B2")
        "'My Sheet'!B2"
    """
    return f"{quote_sheet_if_needed(sheet)}!{cell}"


def normalize_address(address: str) -> str:
    """Normalize an address to a canonical form.

    This parses the address and re-formats it, ensuring consistent quoting
    based on whether the sheet name contains special characters.

    Args:
        address: Sheet-qualified cell address.

    Returns:
        Normalized address matching Node.key format.

    Examples:
        >>> normalize_address("'2024'!A1")  # Quotes not needed
        '2024!A1'
        >>> normalize_address("'My Sheet'!A1")  # Quotes needed (space)
        "'My Sheet'!A1"
    """
    sheet, cell = parse_address(address)
    return format_address(sheet, cell)


def address_to_python_name(address: str) -> str:
    """Convert an Excel cell address to a valid Python function name.

    Examples:
        'Sheet1!A1' -> 'cell_sheet1_a1'
        "'My Sheet'!B2" -> 'cell_my_sheet_b2'
        'B1_GDP_ext!A35' -> 'cell_b1_gdp_ext_a35'

    Args:
        address: Sheet-qualified Excel cell address (e.g., 'Sheet1!A1')

    Returns:
        Valid Python identifier suitable for use as a function name.
    """
    sheet, cell = parse_address(address)

    # Combine sheet and cell
    combined = f"{sheet}_{cell}"

    # Lowercase
    combined = combined.lower()

    # Remove apostrophes (they're word-internal and shouldn't create separators)
    combined = combined.replace("'", "")

    # Replace any non-alphanumeric characters with underscore
    combined = re.sub(r"[^a-z0-9]+", "_", combined)

    # Collapse multiple underscores
    combined = re.sub(r"_+", "_", combined)

    # Remove leading/trailing underscores
    combined = combined.strip("_")

    return f"cell_{combined}"


def excel_func_to_python(name: str) -> str:
    """Convert an Excel function name to a Python function name.

    Examples:
        'SUM' -> 'xl_sum'
        'VLOOKUP' -> 'xl_vlookup'

    Args:
        name: Excel function name (e.g., 'SUM', 'VLOOKUP')

    Returns:
        Python function name with 'xl_' prefix.
    """
    # Lowercase
    result = name.lower()

    # Replace dots with underscores (e.g., NORM.DIST -> norm_dist)
    result = result.replace(".", "_")

    return f"xl_{result}"
