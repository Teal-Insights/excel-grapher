"""Normalize series binding documents for schema validation and codegen."""

from __future__ import annotations

from typing import Any, Literal

InputMode = Literal["leaf", "override"]

_STRUCTURAL_FIELDS = frozenset(
    {
        "id",
        "sheet",
        "data_range",
        "exclude_rows",
        "exclude_columns",
        "layout",
        "structure",
        "key",
        "groups",
        "series_context",
        "validation",
        "editable",
        "sdmx_notes",
        "notes",
    }
)


def effective_dimension_id(component: dict[str, Any]) -> str:
    """Return the record-field identifier for a dimension or attribute.

    Structure components may declare an `id` distinct from their `concept`
    (schema 1.8.0), so two dimensions in one series can share a concept.
    The effective id defaults to the concept when no `id` is declared.
    """
    return str(component.get("id") or component.get("concept") or "")


def component_for_field(series: dict[str, Any], field_name: str) -> dict[str, Any] | None:
    """Return the dimension or attribute whose effective id matches `field_name`."""
    structure = series.get("structure") or {}
    components = [
        *(structure.get("dimensions") or []),
        *(structure.get("attributes") or []),
    ]
    for component in components:
        if not isinstance(component, dict):
            continue
        if effective_dimension_id(component) == field_name:
            return component
    return None


def concept_for_field(series: dict[str, Any], field_name: str) -> str:
    """Map a record field name to the concept id used for scheme lookups.

    A dimension or attribute whose effective id matches `field_name` may
    reference a different concept (schema 1.8.0); dtype inheritance from the
    concept scheme keys on that concept.
    """
    component = component_for_field(series, field_name)
    if component is not None:
        concept = component.get("concept")
        if concept:
            return str(concept)
    return field_name


def _setter_block(series: dict[str, Any]) -> dict[str, Any] | None:
    input_block = series.get("input")
    if isinstance(input_block, dict):
        setter = input_block.get("setter")
        if isinstance(setter, dict):
            return setter
    legacy = series.get("setter")
    if isinstance(legacy, dict):
        return legacy
    return None


def has_input_direction(series: dict[str, Any]) -> bool:
    return _setter_block(series) is not None


def input_mode(series: dict[str, Any]) -> InputMode:
    """Return the declared input binding mode (default ``leaf``)."""
    input_block = series.get("input")
    if not isinstance(input_block, dict):
        return "leaf"
    mode = input_block.get("mode", "leaf")
    if mode in ("leaf", "override"):
        return mode
    return "leaf"


def is_override_input(series: dict[str, Any]) -> bool:
    """Return True when the series declares override input semantics."""
    return input_mode(series) == "override"


def has_output_direction(series: dict[str, Any]) -> bool:
    output = series.get("output")
    return isinstance(output, dict) and isinstance(output.get("compute"), dict)


def has_internal_direction(series: dict[str, Any]) -> bool:
    """Return True when the series declares internal (non-I/O) binding semantics."""
    return "internal" in series and isinstance(series.get("internal"), dict)


def has_constant_direction(series: dict[str, Any]) -> bool:
    """Return True when the series declares constant (reader-only leaf) semantics."""
    return "constant" in series and isinstance(series.get("constant"), dict)


def has_reader_direction(series: dict[str, Any]) -> bool:
    """Return True when the series emits a public `read_*` (input or constant)."""
    return has_input_direction(series) or has_constant_direction(series)


def effective_validation(series: dict[str, Any]) -> dict[str, Any]:
    """Return series validation flags with direction-specific defaults applied."""
    validation = dict(series.get("validation") or {})
    if is_override_input(series):
        validation["intersect_graph_leaves"] = False
    if has_internal_direction(series) and "intersect_graph_formulas" not in validation:
        validation["intersect_graph_formulas"] = True
    if has_constant_direction(series) and "intersect_graph_leaves" not in validation:
        validation["intersect_graph_leaves"] = True
    return validation


def normalize_series_entry(series: dict[str, Any]) -> dict[str, Any]:
    """Return a copy with legacy aliases normalized for schema validation and codegen."""
    out = dict(series)
    if out.get("layout") == "row_series":
        out["layout"] = "series"
    legacy_setter = out.pop("setter", None)
    input_block = out.get("input")
    input_block = {} if not isinstance(input_block, dict) else dict(input_block)

    if legacy_setter is not None:
        if "setter" not in input_block:
            input_block["setter"] = legacy_setter
        elif input_block["setter"] != legacy_setter:
            raise ValueError(
                f"series {series.get('id')!r}: conflicting top-level setter and input.setter"
            )

    if input_block:
        out["input"] = input_block
    elif "input" in out:
        del out["input"]

    return out


def normalize_bindings_document(document: dict[str, Any]) -> dict[str, Any]:
    """Normalize all series entries in a binding manifest."""
    out = dict(document)
    series_list = out.get("series")
    if not isinstance(series_list, list):
        return out
    out["series"] = [normalize_series_entry(s) if isinstance(s, dict) else s for s in series_list]
    return out


def structural_fields_match(left: dict[str, Any], right: dict[str, Any]) -> bool:
    """Return True when shared structural fields agree (missing fields ignored)."""
    for field in _STRUCTURAL_FIELDS:
        if field not in left and field not in right:
            continue
        if left.get(field) != right.get(field):
            return False
    return True


def merge_series_entries(
    existing: dict[str, Any],
    incoming: dict[str, Any],
    *,
    shard_index: int,
) -> dict[str, Any]:
    """Merge two series entries with the same id from different shards."""
    left = normalize_series_entry(existing)
    right = normalize_series_entry(incoming)
    series_id = str(left.get("id", ""))
    if not structural_fields_match(left, right):
        raise ValueError(
            f"Cannot merge series {series_id!r}: structural fields differ across shards "
            f"(shard {shard_index})"
        )

    merged = dict(left)
    for direction in ("input", "output", "internal", "constant"):
        if direction not in right:
            continue
        if direction in merged and merged[direction] != right[direction]:
            raise ValueError(
                f"Cannot merge series {series_id!r}: duplicate conflicting {direction} block "
                f"(shard {shard_index})"
            )
        merged[direction] = right[direction]

    for field in ("sdmx_notes", "notes"):
        if field in right and field not in merged:
            merged[field] = right[field]

    return merged
