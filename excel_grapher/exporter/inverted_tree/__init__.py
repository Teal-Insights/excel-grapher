"""Inverted-tree codegen: leaf-closure functions, no `EvalContext`.

This is the competing export paradigm from issue #597. Default ctx export is
unchanged; callers opt in with `generate_modules(..., paradigm="inverted_tree")`.
"""

from __future__ import annotations

from excel_grapher.exporter.inverted_tree.emit import generate_inverted_tree_modules
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError

__all__ = [
    "InvertedTreeExportError",
    "generate_inverted_tree_modules",
]
