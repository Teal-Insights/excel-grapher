"""Utilities for converting Excel names to Python identifiers.

Address parsing/formatting/normalization live in
`excel_grapher.core.address_keys` as the canonical implementation.
"""

from __future__ import annotations

import re

from excel_grapher.core.address_keys import (
    parse_address,
)
from excel_grapher.core.excel_function_names import (
    EXCEL_FUNCTION_PREFIXES,
    excel_func_to_python_runtime_name,
    normalize_excel_function_name,
)


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
    result = normalize_excel_function_name(name)
    return excel_func_to_python_runtime_name(result)


__all__ = [
    "EXCEL_FUNCTION_PREFIXES",
    "address_to_python_name",
    "excel_func_to_python",
    "normalize_excel_function_name",
]
