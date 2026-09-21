"""Shared type aliases for the series catalog topology contract.

These aliases are the vocabulary visualization and inverted-tree export both
use. Implementations still live under `excel_grapher.exporter.inverted_tree`
until catalog extraction (#722); they must re-export the same objects.
"""

from __future__ import annotations

from typing import Literal

Direction = Literal["input", "constant", "internal", "output"]
Layout = Literal["scalar", "series", "matrix"]
HoleKind = Literal["blank", "off_closure", "literal", "graph_leaf", "bound_leaf"]
AccessClass = Literal[
    "identity", "shift", "affine", "gather", "whole", "dynamic", "cross_partition"
]

SCHEDULE_AXIS = "TIME_PERIOD"
"""Inner ordered dimension of a keyed nest (`TIME_PERIOD` when present)."""
