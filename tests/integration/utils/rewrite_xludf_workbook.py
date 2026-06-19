"""Rewrite allowlisted built-in calls to ``_xludf.`` spelling for regression fixtures."""

from __future__ import annotations

from tests.integration.utils.rewrite_prefixed_workbook import (
    rewrite_formula_to_xludf,
    write_xludf_workbook_copy,
)

__all__ = ["rewrite_formula_to_xludf", "write_xludf_workbook_copy"]
