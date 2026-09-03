"""Export errors for inverted-tree codegen."""

from __future__ import annotations


class InvertedTreeExportError(ValueError):
    """Bound series could not be inverted (unbound ref or verification mismatch)."""
