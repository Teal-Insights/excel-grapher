"""Series-level relation declarations that compile to `CellType.relations`."""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Iterable, Mapping, Sequence
from pathlib import Path
from typing import Any

from excel_grapher.core.cell_types import (
    CellRelation,
    GreaterThanCell,
    NotEqualCell,
    normalize_cell_type_env_key,
)
from excel_grapher.series_bindings.ranges import expand_bound_series_addresses
from excel_grapher.series_bindings.resolve import (
    PartialKeyDomainError,
    UnknownBindKindError,
    resolve_key_domain,
)
from excel_grapher.series_bindings.types import ValidationIssue, make_issue

RELATION_KINDS: tuple[str, ...] = ("greater_than", "not_equal")
RELATION_TYPES: dict[str, type[CellRelation]] = {
    "greater_than": GreaterThanCell,
    "not_equal": NotEqualCell,
}

_NUMERIC_DTYPES = frozenset({"int", "integer", "float", "number"})
_DTYPE_ALIASES = {"integer": "int", "str": "string"}
FrozenKey = tuple[tuple[str, Any], ...]


class SeriesRelationError(ValueError):
    """Raised when series relations cannot be compiled fail-closed."""


def iter_series_relations(series: Mapping[str, Any]) -> list[tuple[str, str]]:
    """Return `(kind, partner_id)` pairs declared on `series`."""
    relations: list[tuple[str, str]] = []
    for item in series.get("relations") or []:
        if not isinstance(item, Mapping):
            continue
        for kind in RELATION_KINDS:
            partner = item.get(kind)
            if isinstance(partner, str) and partner:
                relations.append((kind, partner))
    return relations


def series_by_id(bindings: Mapping[str, Any]) -> dict[str, dict[str, Any]]:
    """Index series entries by id, keeping the first occurrence."""
    indexed: dict[str, dict[str, Any]] = {}
    for series in bindings.get("series") or []:
        if not isinstance(series, dict):
            continue
        series_id = series.get("id")
        if isinstance(series_id, str) and series_id and series_id not in indexed:
            indexed[series_id] = series
    return indexed


def authored_measure_dtype(series: Mapping[str, Any]) -> str | None:
    """Return the measure dtype export will use, or `None` when omitted."""
    structure = series.get("structure") or {}
    measure = structure.get("measure")
    if not isinstance(measure, Mapping):
        return None
    if measure.get("dtype") is not None:
        return str(measure["dtype"])
    bind = measure.get("bind")
    if isinstance(bind, Mapping):
        read = bind.get("read")
        if read not in (None, "auto"):
            return str(read)
    return None


def canonical_measure_dtype(series: Mapping[str, Any]) -> str:
    """Canonical measure dtype; omitted dtype follows export's float default."""
    dtype = authored_measure_dtype(series)
    if dtype is None:
        return "float"
    return _DTYPE_ALIASES.get(dtype, dtype)


def dtypes_comparable(left: Mapping[str, Any], right: Mapping[str, Any]) -> bool:
    """Return True when two series' measure dtypes can be compared."""
    left_dtype = canonical_measure_dtype(left)
    right_dtype = canonical_measure_dtype(right)
    if left_dtype in _NUMERIC_DTYPES and right_dtype in _NUMERIC_DTYPES:
        return True
    return left_dtype == right_dtype


def key_field_set(series: Mapping[str, Any]) -> frozenset[str]:
    """Return declared key field names, ignoring list order."""
    return frozenset(str(field) for field in (series.get("key") or []))


def freeze_key(key: Mapping[str, Any]) -> FrozenKey:
    """Return a hashable key tuple in field-name order."""
    return tuple(sorted(key.items()))


def _issue(
    code: str,
    message: str,
    *,
    series_id: str | None = None,
    address: str | None = None,
) -> ValidationIssue:
    return make_issue("error", code, message, series_id=series_id, address=address)


def relation_declaration_issues(bindings: Mapping[str, Any]) -> list[ValidationIssue]:
    """Validate partner identity, dtypes, irreflexive relations, and DAG shape.

    Cycle detection walks `greater_than` edges only: `not_equal` is symmetric
    and does not impose an order. Any relation to the declaring series itself
    is rejected.
    """
    indexed = series_by_id(bindings)
    issues: list[ValidationIssue] = []
    greater_than_edges: list[tuple[str, str]] = []
    for series_id, series in indexed.items():
        declaring_key = key_field_set(series)
        for kind, partner_id in iter_series_relations(series):
            if partner_id == series_id:
                issues.append(
                    _issue(
                        "reflexive_relation",
                        f"series {series_id!r} {kind} itself is not allowed",
                        series_id=series_id,
                    )
                )
                continue
            partner = indexed.get(partner_id)
            if partner is None:
                issues.append(
                    _issue(
                        "unknown_relation_partner",
                        f"series {series_id!r} {kind} partner {partner_id!r} is not a series id",
                        series_id=series_id,
                    )
                )
                continue
            if not dtypes_comparable(series, partner):
                issues.append(
                    _issue(
                        "incomparable_relation_dtype",
                        f"series {series_id!r} {kind} partner {partner_id!r} has "
                        f"incomparable dtype {canonical_measure_dtype(partner)!r} "
                        f"(declaring dtype {canonical_measure_dtype(series)!r})",
                        series_id=series_id,
                    )
                )
            partner_key = key_field_set(partner)
            if declaring_key != partner_key:
                issues.append(
                    _issue(
                        "incompatible_relation_key",
                        f"series {series_id!r} {kind} partner {partner_id!r} key "
                        f"{sorted(partner_key)!r} must match {sorted(declaring_key)!r}",
                        series_id=series_id,
                    )
                )
            if kind == "greater_than":
                greater_than_edges.append((series_id, partner_id))
    cycle = _greater_than_cycle(greater_than_edges)
    if cycle is not None:
        issues.append(
            _issue(
                "cyclic_relation",
                "greater_than relations form a cycle: " + " -> ".join(cycle),
                series_id=cycle[0],
            )
        )
    return issues


def keyed_addresses(
    series: Mapping[str, Any],
    *,
    workbook: Path | str,
) -> tuple[dict[FrozenKey, str], list[ValidationIssue]]:
    """Map each unique resolved key to the series cell at that key.

    Duplicate keys on the partner (or declaring) series are an error: the
    compiler cannot choose which cell `GreaterThanCell` / `NotEqualCell`
    should name.
    """
    series_id = str(series.get("id") or "") or None
    try:
        addresses = expand_bound_series_addresses(series, workbook=workbook)
        points = resolve_key_domain(workbook, dict(series), addresses)
    except PartialKeyDomainError as exc:
        return {}, [
            _issue(
                "unresolved_relation_key",
                str(exc),
                series_id=series_id or exc.series_id,
            )
        ]
    except (UnknownBindKindError, ValueError, TypeError) as exc:
        return {}, [
            _issue(
                "invalid_data_range",
                str(exc),
                series_id=series_id,
            )
        ]

    index: dict[FrozenKey, str] = {}
    issues: list[ValidationIssue] = []
    for address, key in zip(addresses, points, strict=True):
        frozen = freeze_key(key)
        norm = normalize_cell_type_env_key(address)
        existing = index.get(frozen)
        if existing is not None and existing != norm:
            issues.append(
                _issue(
                    "ambiguous_relation_partner_key",
                    f"series {series_id!r} maps key {dict(frozen)!r} to both {existing} and {norm}",
                    series_id=series_id,
                    address=norm,
                )
            )
            continue
        index[frozen] = norm
    return index, issues


def relation_cell_indexes(
    bindings: Mapping[str, Any],
    *,
    workbook: Path | str,
    series_ids: Iterable[str] | None = None,
) -> tuple[dict[str, dict[FrozenKey, str]], list[ValidationIssue]]:
    """Resolve key-to-address indexes for `series_ids` (or every series)."""
    indexed = series_by_id(bindings)
    wanted = set(series_ids) if series_ids is not None else set(indexed)
    indexes: dict[str, dict[FrozenKey, str]] = {}
    issues: list[ValidationIssue] = []
    for series_id in wanted:
        series = indexed.get(series_id)
        if series is None:
            continue
        index, index_issues = keyed_addresses(series, workbook=workbook)
        issues.extend(index_issues)
        indexes[series_id] = index
    return indexes, issues


def missing_partner_cell_issues(
    bindings: Mapping[str, Any],
    indexes: Mapping[str, dict[FrozenKey, str]],
) -> list[ValidationIssue]:
    """Fail closed when a declaring cell has no unique partner cell at its key."""
    indexed = series_by_id(bindings)
    issues: list[ValidationIssue] = []
    for series_id, series in indexed.items():
        relations = iter_series_relations(series)
        if not relations:
            continue
        declaring_index = indexes.get(series_id) or {}
        declaring_key = key_field_set(series)
        for kind, partner_id in relations:
            if partner_id == series_id:
                continue
            partner = indexed.get(partner_id)
            if partner is None or declaring_key != key_field_set(partner):
                continue
            partner_index = indexes.get(partner_id) or {}
            for frozen, address in declaring_index.items():
                if frozen not in partner_index:
                    issues.append(
                        _issue(
                            "missing_relation_partner_key",
                            f"series {series_id!r} cell {address} {kind} partner "
                            f"{partner_id!r} has no cell at key {dict(frozen)!r}",
                            series_id=series_id,
                            address=address,
                        )
                    )
    return issues


def relation_alignment_issues(
    bindings: Mapping[str, Any],
    *,
    workbook: Path | str,
) -> list[ValidationIssue]:
    """Resolve relation series and report missing or ambiguous partner cells."""
    indexed = series_by_id(bindings)
    needed: set[str] = set()
    for series_id, series in indexed.items():
        relations = iter_series_relations(series)
        if not relations:
            continue
        needed.add(series_id)
        needed.update(partner_id for _kind, partner_id in relations if partner_id in indexed)
    indexes, issues = relation_cell_indexes(bindings, workbook=workbook, series_ids=needed)
    issues.extend(missing_partner_cell_issues(bindings, indexes))
    return issues


def raise_if_relation_errors(issues: Sequence[ValidationIssue]) -> None:
    """Raise `SeriesRelationError` listing every error-level issue."""
    messages = [issue["message"] for issue in issues if issue["level"] == "error"]
    if messages:
        raise SeriesRelationError("; ".join(messages))


def _greater_than_cycle(edges: Sequence[tuple[str, str]]) -> list[str] | None:
    graph: dict[str, list[str]] = defaultdict(list)
    for src, dst in edges:
        graph[src].append(dst)
    visiting: set[str] = set()
    visited: set[str] = set()
    stack: list[str] = []

    def dfs(node: str) -> list[str] | None:
        if node in visiting:
            return stack[stack.index(node) :] + [node]
        if node in visited:
            return None
        visiting.add(node)
        stack.append(node)
        for nxt in graph[node]:
            found = dfs(nxt)
            if found is not None:
                return found
        stack.pop()
        visiting.remove(node)
        visited.add(node)
        return None

    for node in list(graph):
        found = dfs(node)
        if found is not None:
            return found
    return None
