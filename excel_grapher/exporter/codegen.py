"""Generate inverted-tree Python packages from Excel dependency graphs."""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path
from typing import TYPE_CHECKING, ClassVar, cast

from excel_grapher.grapher.graph import DependencyGraph

__all__ = ["CodeGenerator"]

REPRESENTATION_VERSION = "named-axis-v1"

if TYPE_CHECKING:
    from excel_grapher.series_bindings.types import InputSeries, WorkbookSeriesBindings


class CodeGenerator:
    """Generate inverted-tree packages from a `DependencyGraph`.

    Package export is `generate_modules()` and requires series bindings.
    """

    representation_version: ClassVar[str] = REPRESENTATION_VERSION

    def __init__(self, graph: DependencyGraph) -> None:
        """Initialize the code generator.

        Args:
            graph: Dependency graph covering the cells to export. A projection
                facade with `original_graph` is also accepted; export uses that
                original graph.
        """
        self.graph = graph

    def __enter__(self) -> CodeGenerator:
        return self

    def __exit__(self, *args: object) -> None:
        return None

    def _public_graph(self) -> DependencyGraph:
        original = getattr(self.graph, "original_graph", None)
        if original is not None:
            return cast("DependencyGraph", original)
        return self.graph

    def derive_input_series(
        self,
        bindings: WorkbookSeriesBindings,
        *,
        workbook: Path | str,
    ) -> list[InputSeries]:
        """Derive input-series metadata from explicit series bindings."""
        from excel_grapher.series_bindings import derive_input_series

        return derive_input_series(self._public_graph(), bindings, workbook=workbook)

    def generate_modules(
        self,
        *,
        series_bindings: WorkbookSeriesBindings,
        bindings_workbook: Path | str,
        blank_ranges: Sequence[str] | None = None,
    ) -> dict[str, str]:
        """Generate a standalone package with explicit input-leaf computes.

        Outputs come from the bindings catalog, not from a cell-address list.

        Args:
            series_bindings: Catalog of input, constant, internal, and output series.
            bindings_workbook: Workbook used to expand binding ranges.
            blank_ranges: Sheet-qualified rectangles omitted from the graph that
                resolve as empty rather than unbound catalog cells.

        Returns:
            Module filenames mapped to standalone Python source.
        """
        from excel_grapher.exporter.inverted_tree import generate_inverted_tree_modules

        return generate_inverted_tree_modules(
            self._public_graph(),
            series_bindings=series_bindings,
            bindings_workbook=bindings_workbook,
            blank_ranges=blank_ranges,
        )
