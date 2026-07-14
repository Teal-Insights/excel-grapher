"""Lazy range values for the exported Python runtime.

`Range` lives in `excel_grapher.core.grid` and is re-exported here so generated
code and export consumers keep a stable import path.
"""

from __future__ import annotations

from excel_grapher.core.grid import Range

__all__ = ["Range"]
