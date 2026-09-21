"""Shared series-catalog topology contract.

This package owns the **vocabulary and protocols** for bound series, statements,
instance partitions, schedule axes, and classified instance edges. Visualization
and inverted-tree codegen both consume this surface.

It does **not** own Python emission. Rung selection, demand planning, helper
naming, fallback emission, and fused-region loop planning stay in
`excel_grapher.exporter.inverted_tree`. Catalog construction still runs there
until extraction PRs; those implementations must satisfy these protocols.

See `excel_grapher.semantic_model.ownership` for the import-direction decision.
"""

from __future__ import annotations

from excel_grapher.semantic_model.catalog import (
    BoundSeries,
    CatalogOccupancy,
    KeyPoint,
    ScheduleIndex,
    SeriesCatalog,
    SeriesHole,
    Statement,
)
from excel_grapher.semantic_model.edges import (
    CatalogEdges,
    DependenceEdge,
    SemanticCatalogLoader,
    SemanticCatalogView,
)
from excel_grapher.semantic_model.ownership import (
    EMITTER_OWNED_CONCEPTS,
    EMITTER_OWNED_MODULES,
    SEMANTIC_CONSUMER_MODULES,
    SHARED_CATALOG_TYPE_NAMES,
)
from excel_grapher.semantic_model.types import (
    SCHEDULE_AXIS,
    AccessClass,
    Direction,
    HoleKind,
    Layout,
)

__all__ = [
    "AccessClass",
    "BoundSeries",
    "CatalogEdges",
    "CatalogOccupancy",
    "DependenceEdge",
    "Direction",
    "EMITTER_OWNED_CONCEPTS",
    "EMITTER_OWNED_MODULES",
    "HoleKind",
    "KeyPoint",
    "Layout",
    "SCHEDULE_AXIS",
    "SEMANTIC_CONSUMER_MODULES",
    "SHARED_CATALOG_TYPE_NAMES",
    "ScheduleIndex",
    "SemanticCatalogLoader",
    "SemanticCatalogView",
    "SeriesCatalog",
    "SeriesHole",
    "Statement",
]
