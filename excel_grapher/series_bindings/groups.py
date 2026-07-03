"""View-level series binding groups: export sequencing and group manifest.

Groups are a presentation concern of the generated `set_*`/`compute_*` API.
They sequence code export and feed documentation tooling; they never affect
graph extraction, binding resolution, or record semantics.
"""

from __future__ import annotations

import re
from typing import Any, TypedDict

from excel_grapher.series_bindings.normalize import normalize_series_entry
from excel_grapher.series_bindings.types import WorkbookSeriesBindings


class GroupMember(TypedDict):
    """One binding's membership entry inside a group manifest node."""

    id: str
    setter: str | None
    compute: str | None
    order: int | None


class GroupNode(TypedDict):
    """One group in the nested manifest tree."""

    label: str
    path: list[str]
    slug: str
    members: list[GroupMember]
    children: list[GroupNode]


class GroupsManifest(TypedDict):
    """Machine-readable group structure for documentation tooling."""

    groups: list[GroupNode]
    ungrouped: list[GroupMember]


_SLUG_RE = re.compile(r"[^a-z0-9]+")


def group_slug(label: str) -> str:
    """Derive a Python-identifier-safe slug from a human-readable group label."""
    slug = _SLUG_RE.sub("_", label.lower()).strip("_")
    if not slug:
        slug = "group"
    if slug[0].isdigit():
        slug = f"g{slug}"
    return slug


def _group_refs(series: dict[str, Any]) -> list[dict[str, Any]]:
    groups = series.get("groups")
    if not isinstance(groups, list):
        return []
    return [ref for ref in groups if isinstance(ref, dict) and ref.get("path")]


def bindings_have_groups(bindings: WorkbookSeriesBindings | dict[str, Any]) -> bool:
    """Return True when any series declares at least one group ref."""
    return any(
        _group_refs(series) for series in bindings.get("series", []) if isinstance(series, dict)
    )


def _setter_name(series: dict[str, Any]) -> str | None:
    normalized = normalize_series_entry(series)
    setter = (normalized.get("input") or {}).get("setter")
    if isinstance(setter, dict) and setter.get("name"):
        return str(setter["name"])
    return None


def _compute_name(series: dict[str, Any]) -> str | None:
    compute = (series.get("output") or {}).get("compute")
    if isinstance(compute, dict) and compute.get("name"):
        return str(compute["name"])
    return None


def _member(series: dict[str, Any], order: int | None) -> GroupMember:
    return {
        "id": str(series.get("id", "")),
        "setter": _setter_name(series),
        "compute": _compute_name(series),
        "order": order,
    }


class _TreeNode:
    """Mutable group tree node used while scanning bindings in declaration order."""

    def __init__(self, label: str, path: tuple[str, ...]) -> None:
        self.label = label
        self.path = path
        self.children: dict[str, _TreeNode] = {}
        # (order-is-missing, order, declaration_index) sort key per member.
        self.members: list[tuple[bool, int, int, dict[str, Any], int | None]] = []

    def child(self, label: str) -> _TreeNode:
        if label not in self.children:
            self.children[label] = _TreeNode(label, self.path + (label,))
        return self.children[label]

    def sorted_members(self) -> list[tuple[dict[str, Any], int | None]]:
        return [
            (series, order)
            for _, _, _, series, order in sorted(
                self.members, key=lambda item: (item[0], item[1], item[2])
            )
        ]


def _build_group_tree(
    bindings: WorkbookSeriesBindings | dict[str, Any],
    *,
    placement_only: bool,
) -> tuple[_TreeNode, list[tuple[dict[str, Any], int | None]]]:
    """Scan series in declaration order into a group tree.

    When `placement_only` is true each binding is attached to its first group
    ref only (definition sequencing); otherwise every group ref contributes a
    membership entry (manifest view). Returns the synthetic root node and the
    ungrouped bindings in declaration order.
    """
    root = _TreeNode("", ())
    ungrouped: list[tuple[dict[str, Any], int | None]] = []
    for index, series in enumerate(bindings.get("series", [])):
        if not isinstance(series, dict):
            continue
        refs = _group_refs(series)
        if not refs:
            ungrouped.append((series, None))
            continue
        if placement_only:
            refs = refs[:1]
        for ref in refs:
            node = root
            for label in ref["path"]:
                node = node.child(str(label))
            order = ref.get("order")
            order = int(order) if isinstance(order, int) else None
            node.members.append((order is None, order or 0, index, series, order))
    return root, ungrouped


def _flatten_tree(node: _TreeNode, out: list[dict[str, Any]]) -> None:
    for series, _ in node.sorted_members():
        out.append(series)
    for child in node.children.values():
        _flatten_tree(child, out)


def bindings_export_order(
    bindings: WorkbookSeriesBindings | dict[str, Any],
) -> list[dict[str, Any]]:
    """Return series entries in grouped export order.

    Sibling groups follow first-appearance order over the series list; members
    within a leaf group sort by explicit `order` then declaration order; each
    group's own members precede its nested children; ungrouped bindings trail
    in declaration order. A multi-membership binding is placed by its first
    group ref. Without any groups the declaration order is returned unchanged.
    """
    if not bindings_have_groups(bindings):
        return [s for s in bindings.get("series", []) if isinstance(s, dict)]
    root, ungrouped = _build_group_tree(bindings, placement_only=True)
    ordered: list[dict[str, Any]] = []
    _flatten_tree(root, ordered)
    ordered.extend(series for series, _ in ungrouped)
    return ordered


def grouped_public_names(
    bindings: WorkbookSeriesBindings | dict[str, Any],
) -> tuple[list[str], list[str]]:
    """Return unique setter and compute names in grouped export order."""
    setters: list[str] = []
    computes: list[str] = []
    for series in bindings_export_order(bindings):
        setter = _setter_name(series)
        if setter is not None and setter not in setters:
            setters.append(setter)
        compute = _compute_name(series)
        if compute is not None and compute not in computes:
            computes.append(compute)
    return setters, computes


def _manifest_node(node: _TreeNode) -> GroupNode:
    return {
        "label": node.label,
        "path": list(node.path),
        "slug": group_slug(node.label),
        "members": [_member(series, order) for series, order in node.sorted_members()],
        "children": [_manifest_node(child) for child in node.children.values()],
    }


def group_manifest(
    bindings: WorkbookSeriesBindings | dict[str, Any],
) -> GroupsManifest:
    """Build the nested group manifest consumed by documentation tooling.

    Unlike `bindings_export_order`, a multi-membership binding appears under
    every group it references.
    """
    root, ungrouped = _build_group_tree(bindings, placement_only=False)
    return {
        "groups": [_manifest_node(child) for child in root.children.values()],
        "ungrouped": [_member(series, order) for series, order in ungrouped],
    }
