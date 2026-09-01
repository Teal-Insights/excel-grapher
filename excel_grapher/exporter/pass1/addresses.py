"""Workbook address parsing helpers for Pass-1 synthesis."""

from __future__ import annotations


def parse_workbook_address(address: str) -> tuple[str, str, int]:
    """Split a sheet-qualified A1 address into ``(sheet, column, row)``."""
    sheet, colrow = address.split("!", 1)
    column = "".join(character for character in colrow if character.isalpha())
    row = int("".join(character for character in colrow if character.isdigit()))
    return sheet, column, row
