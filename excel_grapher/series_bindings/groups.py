"""View-level API grouping for series binding export sequencing and manifests."""

from __future__ import annotations

import json
from typing import Any, Literal

from excel_grapher.series_bindings.normalize import has_input_direction, has_output_direction
from excel_grapher.series_bindings.types import WorkbookSeriesBindings

Direction = Literal["input", "output"]
API_MANIFEST_SCHEMA_VERSION = "1.0.0"


def any_series_has_groups(bindings: WorkbookSeriesBindings) -> bool:
    """Return whether any series entry declares API groups."""
    for series in bindings.get("series", []):
        if isinstance(series, dict) and series.get("groups"):
            return True
    return False


def _primary_group(series: dict[str, Any]) -> dict[str, Any] | None:
    groups = series.get("groups")
    if not isinstance(groups, list) or not groups:
        return None
    first = groups[0]
    return first if isinstance(first, dict) else None


def _group_sort_key(
    series: dict[str, Any],
    manifest_index: int,
    fn_name: str,
) -> tuple[Any, ...]:
    primary = _primary_group(series)
    if primary is None:
        return (1, (), manifest_index, manifest_index, fn_name)
    path = tuple(str(part) for part in primary.get("path") or [])
    order = primary.get("order")
    order_key = manifest_index if order is None else (0, order)
    return (0, path, order_key, manifest_index, fn_name)


def _series_with_index(
    bindings: WorkbookSeriesBindings,
    *,
    direction: Direction | None = None,
) -> list[tuple[int, dict[str, Any]]]:
    indexed: list[tuple[int, dict[str, Any]]] = []
    for index, series in enumerate(bindings.get("series", [])):
        if not isinstance(series, dict):
            continue
        if direction == "input" and not has_input_direction(series):
            continue
        if direction == "output" and not has_output_direction(series):
            continue
        indexed.append((index, series))
    return indexed


def _function_name(series: dict[str, Any], *, direction: Direction) -> str | None:
    if direction == "input":
        input_block = series.get("input") or {}
        setter = input_block.get("setter") or series.get("setter")
        if isinstance(setter, dict) and setter.get("name"):
            return str(setter["name"])
        return None
    output_block = series.get("output") or {}
    compute = output_block.get("compute")
    if isinstance(compute, dict) and compute.get("name"):
        return str(compute["name"])
    return None


def ordered_series_for_direction(
    bindings: WorkbookSeriesBindings,
    direction: Direction,
) -> list[dict[str, Any]]:
    """Return series entries ordered for export when groups are declared."""
    indexed = _series_with_index(bindings, direction=direction)
    if not any_series_has_groups(bindings):
        return [series for _, series in indexed]

    def sort_key(item: tuple[int, dict[str, Any]]) -> tuple[Any, ...]:
        index, series = item
        fn_name = _function_name(series, direction=direction) or series.get("id", "")
        return _group_sort_key(series, index, str(fn_name))

    return [series for _, series in sorted(indexed, key=sort_key)]


def _ordered_function_names(
    bindings: WorkbookSeriesBindings,
    *,
    direction: Direction,
) -> list[str]:
    names: list[str] = []
    for series in ordered_series_for_direction(bindings, direction):
        fn_name = _function_name(series, direction=direction)
        if fn_name is not None:
            names.append(fn_name)
    if not any_series_has_groups(bindings):
        return sorted(set(names))
    seen: set[str] = set()
    ordered: list[str] = []
    for name in names:
        if name not in seen:
            seen.add(name)
            ordered.append(name)
    return ordered


def ordered_setter_names(bindings: WorkbookSeriesBindings) -> list[str]:
    """Return setter function names ordered by declared API groups."""
    return _ordered_function_names(bindings, direction="input")


def ordered_compute_names(bindings: WorkbookSeriesBindings) -> list[str]:
    """Return compute function names ordered by declared API groups."""
    return _ordered_function_names(bindings, direction="output")


def _insert_group_path(
    tree: dict[str, Any],
    path: list[str],
    *,
    direction: Direction,
    fn_name: str,
) -> None:
    node = tree
    for label in path:
        if label not in node:
            node[label] = {}
        node = node[label]
    node.setdefault("setters", [])
    node.setdefault("computes", [])
    bucket = "setters" if direction == "input" else "computes"
    if fn_name not in node[bucket]:
        node[bucket].append(fn_name)


def _serialize_group_tree(node: dict[str, Any]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for label in sorted(node):
        entry = node[label]
        payload: dict[str, Any] = {}
        if entry.get("setters"):
            payload["setters"] = list(entry["setters"])
        if entry.get("computes"):
            payload["computes"] = list(entry["computes"])
        child_labels = [key for key in entry if key not in {"setters", "computes"}]
        for child_label in sorted(child_labels):
            payload.update(_serialize_group_tree({child_label: entry[child_label]}))
        result[label] = payload
    return result


def build_api_group_manifest(bindings: WorkbookSeriesBindings) -> dict[str, Any]:
    """Build a machine-readable manifest of grouped export API symbols."""
    tree: dict[str, Any] = {}
    members: list[dict[str, Any]] = []

    for direction in ("input", "output"):
        for series in ordered_series_for_direction(bindings, direction):
            fn_name = _function_name(series, direction=direction)
            if fn_name is None:
                continue
            groups = series.get("groups") or []
            member: dict[str, Any] = {
                "series_id": series["id"],
                "name": fn_name,
                "direction": direction,
            }
            if groups:
                member["groups"] = groups
            members.append(member)
            primary = _primary_group(series)
            if primary is not None:
                path = [str(part) for part in primary.get("path") or []]
                if path:
                    _insert_group_path(
                        tree,
                        path,
                        direction=direction,
                        fn_name=fn_name,
                    )

    return {
        "schema_version": API_MANIFEST_SCHEMA_VERSION,
        "bindings_schema_version": bindings.get("schema_version"),
        "members": members,
        "group_tree": _serialize_group_tree(tree),
        "flat": {
            "setters": ordered_setter_names(bindings),
            "computes": ordered_compute_names(bindings),
        },
    }


def emit_api_manifest_json(bindings: WorkbookSeriesBindings) -> str:
    """Serialize the API group manifest as formatted JSON."""
    return json.dumps(build_api_group_manifest(bindings), indent=2, sort_keys=False) + "\n"
