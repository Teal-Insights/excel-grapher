"""Inverted-tree codegen: leaf-closure functions, no `EvalContext`.

Opt in with `generate_modules(..., paradigm="inverted_tree")`. That is the
recommended series-binding export; the library default remains `ctx` until
the issue 662 default-flip gate.
"""

from __future__ import annotations

from excel_grapher.exporter.inverted_tree.emit import generate_inverted_tree_modules
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError

__all__ = [
    "InvertedTreeExportError",
    "generate_inverted_tree_modules",
]
