"""Normalize series binding documents for schema validation and codegen."""

from __future__ import annotations

from typing import Any

_STRUCTURAL_FIELDS = frozenset(
    {
        "id",
        "sheet",
        "data_range",
        "layout",
        "structure",
        "key",
        "series_context",
        "validation",
        "editable",
        "sdmx_notes",
        "notes",
    }
)


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


def has_output_direction(series: dict[str, Any]) -> bool:
    output = series.get("output")
    return isinstance(output, dict) and isinstance(output.get("compute"), dict)


def normalize_series_entry(series: dict[str, Any]) -> dict[str, Any]:
    """Return a copy with legacy ``setter`` folded into ``input.setter`` when needed."""
    out = dict(series)
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
    for direction in ("input", "output"):
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
