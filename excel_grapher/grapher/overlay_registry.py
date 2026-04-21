"""In-process overlay builder registry for lightweight viz (analysis-agnostic)."""

from __future__ import annotations

from collections.abc import Callable, Sequence
from typing import Any, TypeAlias

from .graph import DependencyGraph
from .lightweight_viz import LightweightVizCore, LightweightVizOverlay

OverlayBuilderFn: TypeAlias = Callable[..., LightweightVizOverlay]

_builders: dict[str, OverlayBuilderFn] = {}


def register_overlay_builder(overlay_id: str, builder: OverlayBuilderFn) -> None:
    if overlay_id in _builders:
        raise ValueError(f"duplicate overlay_id registration: {overlay_id!r}")
    _builders[overlay_id] = builder


def list_overlay_builders() -> tuple[str, ...]:
    return tuple(sorted(_builders))


def clear_overlay_builders_for_tests() -> None:
    """Reset registry (tests only)."""
    _builders.clear()


def build_overlays(
    graph: DependencyGraph,
    core: LightweightVizCore,
    requested: Sequence[str],
    *,
    context: Any | None = None,
) -> list[LightweightVizOverlay]:
    overlays: list[LightweightVizOverlay] = []
    for oid in requested:
        if oid not in _builders:
            raise ValueError(f"unknown overlay_id: {oid!r}")
        try:
            overlays.append(_builders[oid](graph, core, context=context))
        except Exception as e:
            raise RuntimeError(f"overlay builder failed for overlay_id={oid!r}") from e
    return overlays
