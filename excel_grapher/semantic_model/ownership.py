"""Ownership and import-direction rules for the semantic catalog contract.

`excel_grapher.semantic_model` is the module owner for catalog topology
shared by visualization and inverted-tree export (#722). This package
declares the contract only; catalog construction and edge classification
still run in `excel_grapher.exporter.inverted_tree` until extraction PRs.

Dependency direction (no cycles):

- `core` is imported by everyone.
- `series_bindings` declares series; it does not import this package.
- This package may import `core` and `series_bindings` types only.
- `grapher` may import this package (occupancy for series-level views).
- `exporter.inverted_tree` implements the contract and may import this package.
- Visualization (`exporter.semantic_*`, `grapher.series_graph`) consumes the
  contract and must not import emitter-owned inverted-tree modules.

The canonical cell graph remains `DependencyGraph`. A catalog is derived
from graph + bindings + workbook; it is not a `ProjectionResult` and is not
TACO.
"""

from __future__ import annotations

SHARED_CATALOG_TYPE_NAMES = (
    "KeyPoint",
    "BoundSeries",
    "Statement",
    "SeriesCatalog",
    "ScheduleIndex",
    "SeriesHole",
    "DependenceEdge",
    "CatalogEdges",
)
"""Inventory types that visualization and inverted-tree both consume."""

EMITTER_OWNED_MODULES = frozenset(
    {
        "excel_grapher.exporter.inverted_tree.emit",
        "excel_grapher.exporter.inverted_tree.ast_emit",
        "excel_grapher.exporter.inverted_tree.named_emit",
        "excel_grapher.exporter.inverted_tree.named_axes",
        "excel_grapher.exporter.inverted_tree.schedule",
        "excel_grapher.exporter.inverted_tree.standalone",
        "excel_grapher.exporter.inverted_tree.runtime",
        "excel_grapher.exporter.inverted_tree.excel",
    }
)
"""Inverted-tree modules that stay exporter-owned (rung, fusion, codegen)."""

EMITTER_OWNED_CONCEPTS = (
    "rung selection",
    "demand / call planning",
    "Python helper naming",
    "fallback emission",
    "fused-region planning that exists only to emit loops",
    "SeriesDeps emit projection",
    "seed / scan access helpers in access.py",
    "INDEX/OFFSET covering-range helpers",
)
"""Concepts that remain inverted-tree policy unless a second consumer needs them."""

SEMANTIC_CONSUMER_MODULES = (
    "excel_grapher.exporter.semantic_catalog",
    "excel_grapher.exporter.semantic_graph",
    "excel_grapher.exporter.semantic_viz",
    "excel_grapher.exporter.semantic_drilldown",
    "excel_grapher.exporter.series_graph",
    "excel_grapher.grapher.series_graph",
)
"""Modules that may consume the catalog contract but not emitter policy."""
