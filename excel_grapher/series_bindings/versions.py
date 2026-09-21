"""Schema version and feature support matrix for series bindings."""

from __future__ import annotations

SUPPORTED_SCHEMA_VERSIONS: frozenset[str] = frozenset(
    {
        "1.0.0",
        "1.1.0",
        "1.2.0",
        "1.3.0",
        "1.4.0",
        "1.5.0",
        "1.6.0",
        "1.7.0",
        "1.8.0",
        "1.9.0",
        "1.10.0",
        "1.11.0",
        "1.12.0",
        "1.13.0",
        "1.14.0",
        "1.15.0",
        "1.16.0",
        "1.17.0",
        "1.18.0",
        "1.19.0",
        "1.20.0",
    }
)

CURRENT_SCHEMA_VERSION: str = max(
    SUPPORTED_SCHEMA_VERSIONS,
    key=lambda version: tuple(int(part) for part in version.split(".")),
)

IMPLEMENTED_BIND_KINDS: frozenset[str] = frozenset(
    {
        "data_cell",
        "cell",
        "column_header",
        "row_label",
        "value_map",
        "constant",
        "sheet_name",
    }
)

IMPLEMENTED_LAYOUTS: frozenset[str] = frozenset({"series", "scalar", "matrix"})


def is_bind_implemented(kind: str | None) -> bool:
    return kind in IMPLEMENTED_BIND_KINDS


def is_layout_implemented(layout: str | None) -> bool:
    return layout in IMPLEMENTED_LAYOUTS
