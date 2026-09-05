"""Generate standalone packages with explicit input-leaf compute functions.

Constant leaves are read from `data`. Key domains are published on
`data.{FIELD}_DOMAIN` and as `__key__` / `__domain__` on each `compute_*`.
"""

from __future__ import annotations

from excel_grapher.exporter.inverted_tree.emit import generate_inverted_tree_modules
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError

__all__ = [
    "InvertedTreeExportError",
    "generate_inverted_tree_modules",
]
