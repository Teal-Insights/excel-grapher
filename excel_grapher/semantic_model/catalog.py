"""Catalog topology protocols: series, statements, occupancy, and schedule."""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any, Protocol

from excel_grapher.core.address_keys import CanonicalAddress
from excel_grapher.semantic_model.types import Direction, HoleKind, Layout
from excel_grapher.series_bindings.types import Scalar


class KeyPoint(Protocol):
    """Resolved key coordinates for one series member."""

    items: tuple[tuple[str, Scalar], ...]

    def __getitem__(self, field: str) -> Scalar:
        """Return the value for `field`."""
        ...

    def as_mapping(self) -> dict[str, Scalar]:
        """Return a new dict of field names to values."""
        ...


class Statement(Protocol):
    """One formula-shape run over an ordered index domain.

    A bound series with mixed member formulas partitions into one statement
    per consecutive shape run. Uniform series are one statement.
    """

    statement_id: str
    series_id: str
    shape_key: str | None
    start: int
    stop: int
    cells: tuple[CanonicalAddress, ...]
    domain: tuple[KeyPoint, ...]


class SeriesHole(Protocol):
    """One catalog cell that is not an on-graph formula."""

    index: int
    address: CanonicalAddress
    kind: HoleKind
    literal: object | None
    claimant_id: str | None


class BoundSeries(Protocol):
    """One bindings-catalog series with expanded cells and an index domain."""

    series_id: str
    layout: Layout
    direction: Direction
    cells: tuple[CanonicalAddress, ...]
    key_fields: tuple[str, ...]
    dtype: str
    compute_name: str | None
    raw: Mapping[str, Any]
    domain: tuple[KeyPoint, ...]
    statements: tuple[Statement, ...]
    axis_labels: str | None
    holes: tuple[SeriesHole, ...]
    authored_cells: tuple[CanonicalAddress, ...] | None
    authored_domain: tuple[KeyPoint, ...] | None

    def index_of(self, address: CanonicalAddress) -> int | None:
        """Return the 0-based index of canonical `address` in `cells`, if present."""
        ...

    def key_point_for(self, address: CanonicalAddress) -> KeyPoint | None:
        """Return resolved key coordinates for `address`, if this series owns it."""
        ...

    @property
    def is_formula_series(self) -> bool:
        """True when internals/outputs emit a helper for this series."""
        ...

    @property
    def is_scalar(self) -> bool:
        """True when the series is a single value."""
        ...


class ScheduleIndex(Protocol):
    """Precomputed schedule coordinates and join fields for one catalog.

    `partition_of` is the outer instance-partition tuple. `axis_of` is the
    inner schedule-axis coordinate (`TIME_PERIOD` when the key is a nest).
    """

    preferred: Mapping[str, tuple[str, ...] | None]
    coord_of: Mapping[CanonicalAddress, int]
    index_by_coord: Mapping[str, Mapping[int, tuple[int, ...]]]
    statement_id_by_coord: Mapping[str, Mapping[int, str]]
    partition_of: Mapping[CanonicalAddress, tuple[Scalar, ...]]
    axis_of: Mapping[CanonicalAddress, int]
    coords_of: Mapping[str, frozenset[int]]


class CatalogOccupancy(Protocol):
    """Cell-to-series owner lookup without catalog construction details.

    `SeriesCatalog` satisfies this. Series-level visualization uses occupancy
    without importing inverted-tree emission.
    """

    def series_id_for(self, address: CanonicalAddress) -> str | None:
        """Return the bound series owning canonical `address`, if any."""
        ...

    def get(self, series_id: str) -> Any:
        """Return series metadata for `series_id`."""
        ...


class SeriesCatalog(CatalogOccupancy, Protocol):
    """Bindings series keyed by id, with reverse address lookup.

    Graph node identity stays `CellKey`. Use `key_point_for` to reach
    per-cell key coordinates without copying them onto `Node.metadata`.
    """

    series: Mapping[str, BoundSeries]
    order: tuple[str, ...]
    address_to_id: Mapping[CanonicalAddress, str]
    schedule: ScheduleIndex

    def series_for(self, address: CanonicalAddress) -> BoundSeries | None:
        """Return the bound series owning canonical `address`, if any."""
        ...

    def require_series_for(self, address: CanonicalAddress) -> BoundSeries:
        """Return the series owning `address`, or fail closed."""
        ...

    def key_point_for(self, address: CanonicalAddress) -> KeyPoint | None:
        """Return resolved key coordinates for `address`, or `None` if unbound."""
        ...

    def binds_for(self, address: CanonicalAddress) -> Mapping[str, Mapping[str, Any]] | None:
        """Return series-level dimension binds for the owner of `address`."""
        ...

    def formula_series(self) -> list[BoundSeries]:
        """Return internals and outputs in bindings order."""
        ...

    def output_series(self) -> list[BoundSeries]:
        """Return output series in bindings order."""
        ...

    def input_series(self) -> list[BoundSeries]:
        """Return mutable input series in bindings order."""
        ...

    def constant_series(self) -> list[BoundSeries]:
        """Return constant series in bindings order."""
        ...
