"""Shared Excel runtime.

Houses the implementations used by `excel_grapher.evaluator` at eval time.
The inverted-tree exporter embeds shared `core` (and selected export-runtime)
symbols via `excel_grapher.exporter.embed.emit_runtime`. This package must
not import from `evaluator`, `exporter`, or `grapher`.
"""

from __future__ import annotations

__all__: list[str] = []
