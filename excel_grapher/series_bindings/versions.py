"""Schema version and feature support matrix for series bindings."""

from __future__ import annotations

SUPPORTED_SCHEMA_VERSIONS: frozenset[str] = frozenset({"1.0.0", "1.1.0", "1.2.0", "1.3.0", "1.4.0"})

IMPLEMENTED_BIND_KINDS: frozenset[str] = frozenset(
    {
        "data_cell",
        "cell",
        "column_header",
        "row_label",
        "constant",
    }
)

PLANNED_BIND_KINDS: frozenset[str] = frozenset({"row_hierarchy"})

IMPLEMENTED_LAYOUTS: frozenset[str] = frozenset({"series", "scalar"})

PLANNED_LAYOUTS: frozenset[str] = frozenset({"matrix"})


def is_bind_implemented(kind: str | None) -> bool:
    return kind in IMPLEMENTED_BIND_KINDS


def is_layout_implemented(layout: str | None) -> bool:
    return layout in IMPLEMENTED_LAYOUTS
