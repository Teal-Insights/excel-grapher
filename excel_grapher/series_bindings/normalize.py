"""Normalize series binding documents for schema validation and codegen."""

from __future__ import annotations

import copy
from typing import Any, Literal

from excel_grapher.series_bindings.ranges import series_data_ranges, series_sheets

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
_COMPLEMENTARY_EQUAL_FIELDS = frozenset(
    {
        "exclude_rows",
        "exclude_columns",
        "layout",
        "key",
        "groups",
        "series_context",
        "validation",
        "editable",
    }
)
_DIRECTION_FIELDS = ("input", "output", "internal", "constant")
_CONSTANT_LIKE_KINDS = frozenset({"constant", "sheet_name"})
_BIND_VALUE_KEYS = frozenset({"kind", "value", "values"})


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


def _field_equal(left: dict[str, Any], right: dict[str, Any], field: str) -> bool:
    if field not in left and field not in right:
        return True
    return left.get(field) == right.get(field)


def _direction_names(series: dict[str, Any]) -> frozenset[str]:
    return frozenset(name for name in _DIRECTION_FIELDS if name in series)


def _component_bind(component: dict[str, Any]) -> dict[str, Any] | None:
    bind = component.get("bind")
    if isinstance(bind, dict):
        return bind
    if "value" in component:
        return {"kind": "constant", "value": component["value"]}
    return None


def _bind_shape(bind: dict[str, Any] | None) -> Any:
    if bind is None:
        return None
    kind = bind.get("kind")
    if kind in _CONSTANT_LIKE_KINDS:
        rest = {key: value for key, value in bind.items() if key not in _BIND_VALUE_KEYS}
        return ("sheet_or_constant", rest)
    return bind


def _component_shape(component: dict[str, Any]) -> tuple[Any, ...]:
    return (
        component.get("id"),
        component.get("concept"),
        component.get("role"),
        component.get("scope"),
        component.get("dtype"),
        component.get("include_in_record"),
        _bind_shape(_component_bind(component)),
    )


def _structure_shape(structure: Any) -> Any:
    if not isinstance(structure, dict):
        return structure
    dimensions = structure.get("dimensions") or []
    attributes = structure.get("attributes") or []
    return (
        structure.get("measure"),
        [_component_shape(item) for item in dimensions if isinstance(item, dict)],
        [_component_shape(item) for item in attributes if isinstance(item, dict)],
    )


def complementary_fields_match(left: dict[str, Any], right: dict[str, Any]) -> bool:
    """Return True when shards share direction, key, and dimension shape.

    `sheet` / `data_range` may differ. Constant and `sheet_name` bind *values*
    may differ; other bind geometry must match.
    """
    if left.get("sheet") == right.get("sheet") and left.get("data_range") == right.get(
        "data_range"
    ):
        return False
    for field in _COMPLEMENTARY_EQUAL_FIELDS:
        if not _field_equal(left, right, field):
            return False
    if _direction_names(left) != _direction_names(right):
        return False
    return _structure_shape(left.get("structure")) == _structure_shape(right.get("structure"))


def _unique_extend(existing: list[str], incoming: list[str]) -> list[str]:
    out = list(existing)
    for item in incoming:
        if item not in out:
            out.append(item)
    return out


def _merge_sheet_field(left: dict[str, Any], right: dict[str, Any]) -> str | list[str]:
    sheets = _unique_extend(series_sheets(left), series_sheets(right))
    if len(sheets) == 1:
        return sheets[0]
    return sheets


def _merge_data_range_field(left: dict[str, Any], right: dict[str, Any]) -> str | list[str]:
    ranges = _unique_extend(series_data_ranges(left), series_data_ranges(right))
    if len(ranges) == 1:
        return ranges[0]
    return ranges


def _bind_read_fields(bind: dict[str, Any]) -> dict[str, Any]:
    return {key: value for key, value in bind.items() if key not in _BIND_VALUE_KEYS}


def _sheet_value_mapping(series: dict[str, Any], bind: dict[str, Any]) -> dict[str, Any]:
    kind = bind.get("kind")
    sheets = series_sheets(series)
    if kind == "constant":
        return {sheet: bind.get("value") for sheet in sheets}
    if kind == "sheet_name":
        values = bind.get("values")
        if isinstance(values, dict) and values:
            return dict(values)
        return {sheet: sheet for sheet in sheets}
    raise ValueError(f"Unsupported complementary bind kind {kind!r}")


def _union_sheet_mappings(
    left_map: dict[str, Any],
    right_map: dict[str, Any],
    *,
    series_id: str,
    shard_index: int,
) -> dict[str, Any]:
    merged = dict(left_map)
    for sheet, value in right_map.items():
        if sheet in merged and merged[sheet] != value:
            raise ValueError(
                f"Cannot merge series {series_id!r}: sheet {sheet!r} maps to "
                f"conflicting values {merged[sheet]!r} and {value!r} (shard {shard_index})"
            )
        merged[sheet] = value
    return merged


def _compress_sheet_mapping(
    mapping: dict[str, Any], *, read_fields: dict[str, Any]
) -> dict[str, Any]:
    values = list(mapping.values())
    if values and all(value == values[0] for value in values):
        return {"kind": "constant", "value": values[0], **read_fields}
    if mapping and all(value == sheet for sheet, value in mapping.items()):
        return {"kind": "sheet_name", **read_fields}
    return {"kind": "sheet_name", "values": dict(mapping), **read_fields}


def _merge_component(
    left_series: dict[str, Any],
    right_series: dict[str, Any],
    left_component: dict[str, Any],
    right_component: dict[str, Any],
    *,
    series_id: str,
    shard_index: int,
) -> dict[str, Any]:
    left_bind = _component_bind(left_component) or {}
    right_bind = _component_bind(right_component) or {}
    merged = dict(left_component)
    if left_bind.get("kind") not in _CONSTANT_LIKE_KINDS:
        return merged
    if left_bind == right_bind:
        return merged
    mapping = _union_sheet_mappings(
        _sheet_value_mapping(left_series, left_bind),
        _sheet_value_mapping(right_series, right_bind),
        series_id=series_id,
        shard_index=shard_index,
    )
    merged["bind"] = _compress_sheet_mapping(mapping, read_fields=_bind_read_fields(left_bind))
    if "value" in merged and merged["bind"].get("kind") != "constant":
        del merged["value"]
    return merged


def _merge_structure(
    left: dict[str, Any],
    right: dict[str, Any],
    *,
    series_id: str,
    shard_index: int,
) -> dict[str, Any]:
    left_structure = left.get("structure") or {}
    right_structure = right.get("structure") or {}
    if left_structure == right_structure:
        return copy.deepcopy(left_structure)
    merged: dict[str, Any] = {"measure": copy.deepcopy(left_structure.get("measure"))}
    left_dims = left_structure.get("dimensions") or []
    right_dims = right_structure.get("dimensions") or []
    merged["dimensions"] = [
        _merge_component(
            left, right, left_dim, right_dim, series_id=series_id, shard_index=shard_index
        )
        for left_dim, right_dim in zip(left_dims, right_dims, strict=True)
    ]
    left_attrs = left_structure.get("attributes") or []
    right_attrs = right_structure.get("attributes") or []
    if left_attrs or right_attrs:
        merged["attributes"] = [
            _merge_component(
                left, right, left_attr, right_attr, series_id=series_id, shard_index=shard_index
            )
            for left_attr, right_attr in zip(left_attrs, right_attrs, strict=True)
        ]
    return merged


def _compose_direction_blocks(
    merged: dict[str, Any],
    right: dict[str, Any],
    *,
    series_id: str,
    shard_index: int,
) -> dict[str, Any]:
    for direction in _DIRECTION_FIELDS:
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


def merge_series_entries(
    existing: dict[str, Any],
    incoming: dict[str, Any],
    *,
    shard_index: int,
) -> dict[str, Any]:
    """Merge two series entries with the same id from different shards.

    Identical `sheet` / `data_range` / structure compose direction blocks
    (`input` + `output` on one rectangle). Complementary shards that share
    direction, `key`, and dimension shape concatenate `data_range` even when
    `sheet` differs; differing constant / `sheet_name` binds become one
    `sheet_name` dimension.
    """
    left = normalize_series_entry(existing)
    right = normalize_series_entry(incoming)
    series_id = str(left.get("id", ""))
    if structural_fields_match(left, right):
        return _compose_direction_blocks(
            dict(left), right, series_id=series_id, shard_index=shard_index
        )
    if not complementary_fields_match(left, right):
        raise ValueError(
            f"Cannot merge series {series_id!r}: structural fields differ across shards "
            f"(shard {shard_index})"
        )

    merged = copy.deepcopy(left)
    merged["sheet"] = _merge_sheet_field(left, right)
    merged["data_range"] = _merge_data_range_field(left, right)
    merged["structure"] = _merge_structure(
        left, right, series_id=series_id, shard_index=shard_index
    )
    return _compose_direction_blocks(merged, right, series_id=series_id, shard_index=shard_index)
