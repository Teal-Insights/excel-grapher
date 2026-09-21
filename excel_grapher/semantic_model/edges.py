"""Classified instance-edge and catalog-view protocols."""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from pathlib import Path
from typing import Protocol

from excel_grapher.core.address_keys import CanonicalAddress
from excel_grapher.semantic_model.catalog import SeriesCatalog
from excel_grapher.semantic_model.types import AccessClass
from excel_grapher.series_bindings.types import WorkbookSeriesBindings


class DependenceEdge(Protocol):
    """One instance-level read, annotated with access class and schedule distance.

    `distance` is `axis(consumer) - axis(producer)` on the inner schedule
    axis (`TIME_PERIOD` when the key is a nest). `access` is the shared
    access class (`identity` / `shift` / `affine` / `gather` / `whole` /
    `dynamic` / `cross_partition`), distinct from emit-only access functions.
    """

    consumer_id: str
    producer_id: str
    consumer_cell: CanonicalAddress
    producer_cell: CanonicalAddress
    distance: int
    access: AccessClass
    coeff: int | None
    offset: int | None
    guarded: bool


class CatalogEdges(Protocol):
    """Instance-level edges for every formula series, walked once."""

    edges: tuple[DependenceEdge, ...]
    by_consumer: Mapping[str, tuple[DependenceEdge, ...]]
    by_producer: Mapping[str, tuple[DependenceEdge, ...]]


class SemanticCatalogView(Protocol):
    """Bound catalog plus classified instance edges for one workbook."""

    catalog: SeriesCatalog
    edges: CatalogEdges
    concepts: Mapping[str, tuple[str | None, str | None]]


class SemanticCatalogLoader(Protocol):
    """Build a catalog view from graph + bindings + workbook.

    The current implementation is
    `excel_grapher.exporter.semantic_catalog.load_semantic_catalog`.
    It must not import inverted-tree emission.
    """

    def __call__(
        self,
        graph: object,
        bindings: WorkbookSeriesBindings,
        *,
        workbook: Path | str,
        blank_ranges: Iterable[str] | None = None,
    ) -> SemanticCatalogView:
        """Return the statement-partitioned catalog and classified edges."""
        ...
