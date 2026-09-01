"""Fingerprint summaries and ref-relation inference for cluster refactor prompts.

Under ``series_ast`` clustering, members of one refactor unit share a structural
skeleton and differ only in which concrete cells fill each ref slot. A complete
description of the cluster is therefore: one exemplar translation + a per-ref
relation matrix relating each slot's binding keys to the member's own keys.

When a shared skeleton still mixes binding series behind any ``ref_N`` (regime
boundaries inside one owning series), members are partitioned into separate
fingerprint groups so each group keeps a uniform ``series_id`` per slot.

Incomplete ``address_to_series_id`` maps leave unbound operands under regime
``None``. Those mates stay together only when they share operand sheet/row
geometry; mixed unbound geometry falls back rather than emitting one semantic
group.
"""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from itertools import combinations
from pathlib import Path
from typing import TYPE_CHECKING, Literal

import fastpyxl.utils.cell as xl_cell_utils

from excel_grapher.exporter.pass1.addresses import parse_workbook_address
from excel_grapher.exporter.pass1.bindings import BindingKeyValue
from excel_grapher.exporter.pass1.clustering import (
    BoundAddressKeys,
    _ClusteringKeyCache,
    _ref_position_key_values,
    format_structural_skeleton,
    structural_fingerprint,
)

if TYPE_CHECKING:
    from excel_grapher.exporter.pass1.models import MemberContext


@dataclass(frozen=True)
class SemanticDependencyRef:
    """Minimal semantic-helper handle used for ref resolution (avoids cycles)."""

    helper_name: str
    call_form: str
    address_template: str
    addresses: tuple[str, ...]


KeyCombo = tuple[tuple[str, BindingKeyValue], ...]
LookupKey = BindingKeyValue | tuple[BindingKeyValue, ...]
"""A lookup-table key: one member key value, or a tuple of them (issue #163)."""
RelationTier = Literal["constant", "identity", "offset", "lookup", "explicit"]
ResolutionKind = Literal["semantic_helper", "xl_cell", "self_recurrence", "unresolved"]


@dataclass(frozen=True)
class RefResolution:
    """How the referenced cells are resolved at runtime."""

    kind: ResolutionKind
    helper_name: str | None = None
    call_form: str | None = None
    sheet: str | None = None
    address_template: str | None = None
    row_by_dim: tuple[tuple[str, tuple[tuple[BindingKeyValue, int], ...]], ...] = ()
    col_by_dim: tuple[tuple[str, tuple[tuple[BindingKeyValue, str], ...]], ...] = ()


@dataclass(frozen=True)
class RefRelation:
    """How one ref slot's binding keys relate to the member's own keys."""

    ref_index: int
    tier: RelationTier
    series_id: str | None
    fixed_keys: dict[str, BindingKeyValue]
    identity_dims: tuple[str, ...]
    offsets: dict[str, int]
    lookups: dict[str, dict[LookupKey, BindingKeyValue]]
    explicit: tuple[tuple[KeyCombo, KeyCombo], ...] | None
    resolution: RefResolution
    lookup_bases: dict[str, str] = field(default_factory=dict)
    """Map ref-dim -> member dim used as the base for a lag lookup.

    When ``lookups[dim]`` holds lag deltas, ``lookup_bases[dim]`` is the
    member dimension that is lagged (usually the same as ``dim``). Absent for
    direct value lookups (``ref.d == table[member.k]``).
    """
    lookup_keys: dict[str, str | tuple[str, ...]] = field(default_factory=dict)
    """Map ref-dim -> member dim(s) whose value indexes ``lookups[dim]``.

    For a lag lookup this is the dimension whose value selects the lag delta;
    for a direct value lookup it is the dimension whose value selects the ref
    key. A tuple names jointly determining dimensions whose value tuple keys
    ``lookups[dim]`` (issue #163). Present for every entry in ``lookups``.
    """


@dataclass(frozen=True)
class FingerprintGroup:
    skeleton_text: str
    members: tuple[str, ...]
    exemplar: MemberContext
    ref_relations: tuple[RefRelation, ...]
    ref_addresses_by_member: tuple[tuple[str, tuple[str, ...]], ...] = ()
    ref_keys_by_member: tuple[tuple[str, tuple[dict[str, BindingKeyValue], ...]], ...] = ()
    """Per member address, the binding keys of each ref slot in slot order."""


@dataclass(frozen=True)
class ClusterFingerprintSummary:
    groups: tuple[FingerprintGroup, ...]
    key_space: dict[str, tuple[BindingKeyValue, ...]]
    key_to_column: dict[BindingKeyValue, str] | None
    fallback_reason: str | None
    fingerprint_group_count: int = 0
    relation_tiers: tuple[str, ...] = ()
    legacy_token_estimate: int | None = None
    fingerprint_token_estimate: int | None = None

    def __post_init__(self) -> None:
        if not self.fingerprint_group_count:
            object.__setattr__(self, "fingerprint_group_count", len(self.groups))
        if not self.relation_tiers and self.groups:
            tiers = tuple(
                relation.tier for group in self.groups for relation in group.ref_relations
            )
            object.__setattr__(self, "relation_tiers", tiers)


def _key_combo(keys: Mapping[str, BindingKeyValue]) -> KeyCombo:
    return tuple(sorted(keys.items()))


def _is_numeric(value: BindingKeyValue) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _as_number(value: BindingKeyValue) -> int | float:
    if not _is_numeric(value):
        raise TypeError(f"expected numeric binding key, got {value!r}")
    assert isinstance(value, (int, float))
    return value


def _numeric_delta(left: BindingKeyValue, right: BindingKeyValue) -> int | float:
    """Return ``left - right`` for numeric binding keys."""
    return _as_number(left) - _as_number(right)


def _constant_offset(
    member_value: BindingKeyValue,
    ref_value: BindingKeyValue,
) -> int | None:
    if member_value == ref_value:
        return 0
    if _is_numeric(member_value) and _is_numeric(ref_value):
        assert isinstance(member_value, (int, float))
        assert isinstance(ref_value, (int, float))
        delta = ref_value - member_value
        if isinstance(delta, float) and not delta.is_integer():
            return None
        return int(delta)
    return None


def _single_valued_table(
    pairs: Sequence[tuple[BindingKeyValue, BindingKeyValue]],
) -> dict[BindingKeyValue, BindingKeyValue] | None:
    table: dict[BindingKeyValue, BindingKeyValue] = {}
    for key, value in pairs:
        existing = table.get(key)
        if existing is None:
            table[key] = value
        elif existing != value:
            return None
    return table


def _as_lookup_table(
    table: Mapping[BindingKeyValue, BindingKeyValue],
) -> dict[LookupKey, BindingKeyValue]:
    """Widen a scalar-keyed table to the ``LookupKey``-keyed schema."""
    widened: dict[LookupKey, BindingKeyValue] = {}
    for key, value in table.items():
        widened[key] = value
    return widened


def _subset_routing_lookup(
    dimension_id: str,
    addresses: Sequence[str],
    member_keys: Mapping[str, Mapping[str, BindingKeyValue]],
    ref_keys_by_member: Mapping[str, Mapping[str, BindingKeyValue]],
    member_dims: Sequence[str],
) -> tuple[str | tuple[str, ...], dict[LookupKey, BindingKeyValue]] | None:
    """Find the smallest member-dim subset whose value table routes ``dimension_id``.

    Fallback for ref dims the strict single-dim tiers cannot express (issue
    #163): search subsets of ascending arity for a single-valued table over the
    recorded per-member ref keys — the ground truth for how the original graph
    routed each call site. Unlike the strict search, restating tables are
    accepted; per-member verification in mechanical synthesis keeps them exact.
    Returns ``(key_dims, table)`` with a scalar key dim (and scalar-keyed
    table) at arity 1, tuples otherwise; ``None`` when even the full member
    tuple is not single-valued (duplicate member key combos with conflicting
    refs).
    """
    usable_dims = [
        dim for dim in member_dims if all(dim in member_keys[address] for address in addresses)
    ]
    for arity in range(1, len(usable_dims) + 1):
        for combo in combinations(usable_dims, arity):
            table: dict[LookupKey, BindingKeyValue] = {}
            single_valued = True
            for address in addresses:
                key: LookupKey = (
                    member_keys[address][combo[0]]
                    if arity == 1
                    else tuple(member_keys[address][dim] for dim in combo)
                )
                value = ref_keys_by_member[address][dimension_id]
                existing = table.get(key)
                if existing is None:
                    table[key] = value
                elif existing != value:
                    single_valued = False
                    break
            if single_valued:
                return (combo[0] if arity == 1 else combo), table
    return None


def classify_ref_relation(
    ref_index: int,
    member_keys: Mapping[str, Mapping[str, BindingKeyValue]],
    ref_keys_by_member: Mapping[str, Mapping[str, BindingKeyValue]],
    *,
    series_id: str | None = None,
    resolution: RefResolution | None = None,
) -> RefRelation:
    """Infer how one ref position's keys relate to each member's own keys."""
    addresses = tuple(member_keys)
    if not addresses:
        raise ValueError("member_keys must be non-empty")
    if set(addresses) != set(ref_keys_by_member):
        raise ValueError("ref_keys_by_member must cover the same members")

    resolved = resolution or RefResolution(kind="unresolved")
    ref_dims = sorted(
        {dimension_id for keys in ref_keys_by_member.values() for dimension_id in keys}
    )
    member_dims = sorted({dimension_id for keys in member_keys.values() for dimension_id in keys})

    fixed_keys: dict[str, BindingKeyValue] = {}
    identity_dims: list[str] = []
    offsets: dict[str, int] = {}
    lookups: dict[str, dict[LookupKey, BindingKeyValue]] = {}
    lookup_bases: dict[str, str] = {}
    lookup_keys: dict[str, str | tuple[str, ...]] = {}
    needs_explicit = False

    for dimension_id in ref_dims:
        ref_values = [ref_keys_by_member[address][dimension_id] for address in addresses]
        if len(set(ref_values)) == 1:
            fixed_keys[dimension_id] = ref_values[0]
            continue

        if dimension_id in member_dims:
            deltas = [
                _constant_offset(
                    member_keys[address][dimension_id],
                    ref_keys_by_member[address][dimension_id],
                )
                for address in addresses
            ]
            if all(delta is not None for delta in deltas) and len(set(deltas)) == 1:
                delta = deltas[0]
                assert delta is not None
                if delta == 0:
                    identity_dims.append(dimension_id)
                else:
                    offsets[dimension_id] = delta
                continue

            # Ragged lag: member.d - ref.d is a function of exactly one other
            # member dimension (prompt Example 2).
            if all(_is_numeric(member_keys[a][dimension_id]) for a in addresses) and all(
                _is_numeric(ref_keys_by_member[a][dimension_id]) for a in addresses
            ):
                lag_candidates: list[tuple[str, dict[BindingKeyValue, BindingKeyValue]]] = []
                for key_dim in member_dims:
                    if key_dim == dimension_id:
                        continue
                    pairs = [
                        (
                            member_keys[address][key_dim],
                            int(
                                _numeric_delta(
                                    member_keys[address][dimension_id],
                                    ref_keys_by_member[address][dimension_id],
                                )
                            ),
                        )
                        for address in addresses
                    ]
                    table = _single_valued_table(pairs)
                    if table is not None and len(table) >= 1:
                        lag_candidates.append((key_dim, table))
                if len(lag_candidates) == 1:
                    key_dim, table = lag_candidates[0]
                    lookups[dimension_id] = _as_lookup_table(table)
                    lookup_bases[dimension_id] = dimension_id
                    lookup_keys[dimension_id] = key_dim
                    continue

        # Direct value lookup: ref.d == table[member.k] for exactly one k.
        # Reject tables that merely restate a unique-per-member key (any irregular
        # assignment is a function of a unique id); require compression
        # (``len(table) < len(addresses)``) or a ref dim absent from member keys.
        value_candidates: list[tuple[str, dict[BindingKeyValue, BindingKeyValue]]] = []
        for key_dim in member_dims:
            pairs = [
                (
                    member_keys[address][key_dim],
                    ref_keys_by_member[address][dimension_id],
                )
                for address in addresses
            ]
            table = _single_valued_table(pairs)
            if table is None:
                continue
            compresses = len(table) < len(addresses)
            external_dim = dimension_id not in member_dims
            if compresses or external_dim:
                value_candidates.append((key_dim, table))
        if len(value_candidates) == 1:
            key_dim, table = value_candidates[0]
            lookups[dimension_id] = _as_lookup_table(table)
            lookup_keys[dimension_id] = key_dim
            continue
        if len(value_candidates) > 1:
            # Prefer the key dim that is not the ref dim itself when multiple fit.
            non_self = [c for c in value_candidates if c[0] != dimension_id]
            if len(non_self) == 1:
                key_dim, table = non_self[0]
                lookups[dimension_id] = _as_lookup_table(table)
                lookup_keys[dimension_id] = key_dim
                continue

        # No (unambiguous) single-dim relation: search member-dim subsets of
        # ascending arity for a single-valued routing table (issue #163).
        subset = _subset_routing_lookup(
            dimension_id, addresses, member_keys, ref_keys_by_member, member_dims
        )
        if subset is not None:
            key_dims, routing_table = subset
            lookups[dimension_id] = routing_table
            lookup_keys[dimension_id] = key_dims
            continue

        needs_explicit = True
        break

    if needs_explicit:
        explicit = tuple(
            (
                _key_combo(member_keys[address]),
                _key_combo(ref_keys_by_member[address]),
            )
            for address in addresses
        )
        return RefRelation(
            ref_index=ref_index,
            tier="explicit",
            series_id=series_id,
            fixed_keys={},
            identity_dims=(),
            offsets={},
            lookups={},
            explicit=explicit,
            resolution=resolved,
            lookup_bases={},
        )

    if lookups:
        tier: RelationTier = "lookup"
    elif offsets:
        tier = "offset"
    elif identity_dims:
        tier = "identity"
    else:
        tier = "constant"

    return RefRelation(
        ref_index=ref_index,
        tier=tier,
        series_id=series_id,
        fixed_keys=fixed_keys,
        identity_dims=tuple(sorted(identity_dims)),
        offsets=dict(sorted(offsets.items())),
        lookups={key: dict(sorted(table.items())) for key, table in sorted(lookups.items())},
        explicit=None,
        resolution=resolved,
        lookup_bases=dict(sorted(lookup_bases.items())),
        lookup_keys=dict(sorted(lookup_keys.items())),
    )


def _series_id_for_refs(
    ref_addresses: Sequence[str],
    address_to_series_id: Mapping[str, str] | None,
) -> str | None:
    if not address_to_series_id:
        return None
    series_ids = {
        address_to_series_id[address]
        for address in ref_addresses
        if address in address_to_series_id
    }
    if len(series_ids) == 1:
        return next(iter(series_ids))
    return None


def _address_to_helper_map(
    semantic_dependencies: Sequence[SemanticDependencyRef],
) -> dict[str, str]:
    """Map each covered operand address to the semantic helper that serves it.

    A single binding *series* can be published as several helpers when the
    scheduler peels its source cluster (e.g. ``baseline_interest_rate`` /
    ``baseline_interest_rate_2`` split by year for cycle breaking). Downstream
    operands that sweep across the split boundary must partition by the concrete
    helper, not just the series id, so each fingerprint group's slot resolves to
    one helper (``_resolve_ref`` requires every member ref ⊆ one dependency).
    """
    mapping: dict[str, str] = {}
    for dependency in semantic_dependencies:
        for address in dependency.addresses:
            mapping.setdefault(address, dependency.helper_name)
    return mapping


_SELF_REGIME = "\0self"
"""Regime marker for an operand this refactor unit computes itself.

Never a legal helper name or series id, so an in-cluster operand can never
collide with a sibling helper that carries the operand's series id (issue #138).
"""


def _ref_series_regime_key(
    ref_addresses: Sequence[str],
    address_to_series_id: Mapping[str, str] | None,
    address_to_helper: Mapping[str, str] | None = None,
    cluster_member_addresses: frozenset[str] = frozenset(),
) -> tuple[str | None, ...]:
    """Per-slot resolution identity for one member.

    In-cluster operands take :data:`_SELF_REGIME` (they resolve as
    ``self_recurrence``, not through any helper). Otherwise prefers the concrete
    semantic helper serving the operand (so a series split across helpers
    partitions correctly); falls back to the binding series id, then ``None``
    when unbound/unknown.
    """
    key: list[str | None] = []
    for address in ref_addresses:
        if address in cluster_member_addresses:
            key.append(_SELF_REGIME)
            continue
        helper = address_to_helper.get(address) if address_to_helper else None
        if helper is not None:
            key.append(helper)
        elif address_to_series_id:
            key.append(address_to_series_id.get(address))
        else:
            key.append(None)
    return tuple(key)


def _partition_members_by_ref_series_regime(
    group_members: Sequence[MemberContext],
    refs_by_address: Mapping[str, tuple[str, ...]],
    address_to_series_id: Mapping[str, str] | None,
    address_to_helper: Mapping[str, str] | None = None,
    cluster_member_addresses: frozenset[str] = frozenset(),
) -> tuple[tuple[MemberContext, ...], ...]:
    """Split skeleton-mates whose ref slots land in different helpers/series.

    Mechanical synthesis and LLM prompts assume each fingerprint group has one
    helper behind each ``ref_N``. When AST clustering keeps regime boundaries
    together — or a series is published as several helpers — partition here so
    each group gets uniform reads; existing multi-group synthesis then routes by
    the varying key.

    A peeled recurrence mixes both regimes in one slot: the first member of the
    later unit lags into the *sibling* helper while its mates lag into the unit
    itself. Both carry the same series id, so the in-cluster marker is what
    keeps them apart.

    Unbound / incomplete maps: slots whose identity is ``None`` are not
    further partitioned here. Callers must reject groups whose unbound members
    disagree on operand ``(sheet, row)`` geometry — see
    ``_unbound_ref_slot_geometry_conflict``. Same-geometry unbound mates (for
    example a column sweep on one row) may remain one group.
    """
    if (not address_to_series_id and not address_to_helper) or len(group_members) < 2:
        return (tuple(group_members),)

    partitions: dict[tuple[str | None, ...], list[MemberContext]] = {}
    order: list[tuple[str | None, ...]] = []
    for member in group_members:
        regime = _ref_series_regime_key(
            refs_by_address[member.address],
            address_to_series_id,
            address_to_helper,
            cluster_member_addresses,
        )
        if regime not in partitions:
            partitions[regime] = []
            order.append(regime)
        partitions[regime].append(member)
    return tuple(tuple(partitions[regime]) for regime in order)


def _unbound_ref_slot_geometry_conflict(
    group_members: Sequence[MemberContext],
    refs_by_address: Mapping[str, tuple[str, ...]],
    address_to_series_id: Mapping[str, str] | None,
) -> str | None:
    """Return a fallback reason when unbound ref slots mix sheet/row geometry.

    Regime split keys only on ``series_id`` (``None`` when the operand is
    missing from ``address_to_series_id``). After that split, a multi-member
    group whose unbound slot operands disagree on ``(sheet, row)`` is unsafe
    to treat as one semantic helper — fall back. Column variation on a shared
    sheet/row is allowed. Applies only when a series map is present; a
    missing/empty map leaves prior behavior unchanged.
    """
    if not address_to_series_id or len(group_members) < 2:
        return None

    regime = _ref_series_regime_key(refs_by_address[group_members[0].address], address_to_series_id)
    for slot_index, series_id in enumerate(regime):
        if series_id is not None:
            continue
        geometries: set[tuple[str, int]] = set()
        for member in group_members:
            sheet, _column, row = parse_workbook_address(
                refs_by_address[member.address][slot_index]
            )
            geometries.add((sheet, row))
        if len(geometries) > 1:
            return (
                "unbound_ref_slot_geometry_conflict:"
                f"ref_{slot_index} mixes sheet/row among unbound operands"
            )
    return None


def _resolve_ref(
    ref_addresses_by_member: Mapping[str, str],
    *,
    cluster_member_addresses: frozenset[str],
    semantic_dependencies: Sequence[SemanticDependencyRef] = (),
    member_keys: Mapping[str, Mapping[str, BindingKeyValue]] | None = None,
) -> RefResolution:
    addresses = tuple(ref_addresses_by_member.values())
    if addresses and all(address in cluster_member_addresses for address in addresses):
        return RefResolution(kind="self_recurrence")

    for dependency in semantic_dependencies:
        covered = set(dependency.addresses)
        if addresses and set(addresses) <= covered:
            return RefResolution(
                kind="semantic_helper",
                helper_name=dependency.helper_name,
                call_form=dependency.call_form,
                address_template=dependency.address_template,
            )

    sheets = {parse_workbook_address(address)[0] for address in addresses}
    if len(sheets) != 1:
        return RefResolution(kind="unresolved")
    sheet = next(iter(sheets))
    rows = {parse_workbook_address(address)[2] for address in addresses}
    cols = {parse_workbook_address(address)[1] for address in addresses}

    row_by_dim: list[tuple[str, tuple[tuple[BindingKeyValue, int], ...]]] = []
    col_by_dim: list[tuple[str, tuple[tuple[BindingKeyValue, str], ...]]] = []
    if member_keys is not None:
        for dimension_id in sorted({dim for keys in member_keys.values() for dim in keys}):
            row_pairs = [
                (
                    member_keys[member][dimension_id],
                    parse_workbook_address(ref_addresses_by_member[member])[2],
                )
                for member in ref_addresses_by_member
                if dimension_id in member_keys[member]
            ]
            row_table = _single_valued_table(row_pairs)
            if row_table is not None and len(set(row_table.values())) > 1 and len(rows) > 1:
                row_by_dim.append(
                    (
                        dimension_id,
                        tuple(sorted((k, int(v)) for k, v in row_table.items())),
                    )
                )
            col_pairs = [
                (
                    member_keys[member][dimension_id],
                    parse_workbook_address(ref_addresses_by_member[member])[1],
                )
                for member in ref_addresses_by_member
                if dimension_id in member_keys[member]
            ]
            col_table = _single_valued_table(
                [(k, v) for k, v in col_pairs]  # type: ignore[misc]
            )
            if col_table is not None and len(set(col_table.values())) > 1 and len(cols) > 1:
                col_by_dim.append(
                    (
                        dimension_id,
                        tuple(sorted((k, str(v)) for k, v in col_table.items())),
                    )
                )

    template = f"{sheet}!{{col}}{{row}}"
    if len(rows) == 1 and len(cols) == 1:
        only = addresses[0]
        _sheet, column, row = parse_workbook_address(only)
        template = f"{sheet}!{column}{row}"
    elif len(rows) == 1:
        only_row = next(iter(rows))
        template = f"{sheet}!{{col}}{only_row}"
    elif len(cols) == 1:
        only_col = next(iter(cols))
        template = f"{sheet}!{only_col}{{row}}"

    return RefResolution(
        kind="xl_cell",
        sheet=sheet,
        address_template=template,
        row_by_dim=tuple(row_by_dim),
        col_by_dim=tuple(col_by_dim),
    )


def _key_space_from_expected(
    expected_member_keys: Mapping[str, Mapping[str, BindingKeyValue]],
) -> dict[str, tuple[BindingKeyValue, ...]]:
    values_by_dim: dict[str, set[BindingKeyValue]] = defaultdict(set)
    for keys in expected_member_keys.values():
        for dimension_id, value in keys.items():
            values_by_dim[dimension_id].add(value)

    def _sort_key(value: BindingKeyValue) -> tuple[int, object]:
        if isinstance(value, bool):
            return (3, value)
        if isinstance(value, (int, float)):
            return (0, value)
        if isinstance(value, str):
            return (1, value)
        return (2, str(value))

    return {
        dimension_id: tuple(sorted(values, key=_sort_key))
        for dimension_id, values in sorted(values_by_dim.items())
    }


def _key_to_column_from_members(
    key_space: Mapping[str, tuple[BindingKeyValue, ...]],
    members: Sequence[MemberContext],
    expected_member_keys: Mapping[str, Mapping[str, BindingKeyValue]],
) -> dict[BindingKeyValue, str] | None:
    """Map each member's projection-period key to that member's column letter."""
    if "TIME_PERIOD" in key_space:
        period_dimension_id = "TIME_PERIOD"
    elif "PROJECTION_PERIOD" in key_space:
        period_dimension_id = "PROJECTION_PERIOD"
    else:
        return None
    mapping: dict[BindingKeyValue, str] = {}
    for member in members:
        keys = expected_member_keys.get(member.address)
        if keys is None:
            continue
        value = keys.get(period_dimension_id)
        if value is None:
            continue
        mapping.setdefault(value, member.engine_column)
    return mapping or None


def _fallback_summary(
    reason: str,
    *,
    members: Sequence[MemberContext],
    expected_member_keys: Mapping[str, Mapping[str, BindingKeyValue]],
) -> ClusterFingerprintSummary:
    key_space = _key_space_from_expected(expected_member_keys)
    return ClusterFingerprintSummary(
        groups=(),
        key_space=key_space,
        key_to_column=_key_to_column_from_members(key_space, members, expected_member_keys),
        fallback_reason=reason,
        legacy_token_estimate=estimate_legacy_dump_tokens(members),
        fingerprint_token_estimate=None,
    )


def build_cluster_fingerprint_summary(
    members: Sequence[MemberContext],
    *,
    expected_member_keys: Mapping[str, Mapping[str, BindingKeyValue]],
    bound_address_keys: BoundAddressKeys,
    workbook_path: Path | None = None,
    address_to_series_id: Mapping[str, str] | None = None,
    semantic_dependencies: Sequence[SemanticDependencyRef] = (),
    key_cache: _ClusteringKeyCache | None = None,
) -> ClusterFingerprintSummary:
    """Build a lossless fingerprint summary for one cluster refactor unit."""
    if len(members) < 1:
        return _fallback_summary(
            "empty_cluster",
            members=members,
            expected_member_keys=expected_member_keys,
        )

    cluster_addresses = frozenset(member.address for member in members)
    formula_by_address = {member.address: member.normalized_formula for member in members}

    # Group members by structural skeleton (future-proof for series clustering).
    groups_by_skeleton: dict[tuple, list[MemberContext]] = defaultdict(list)
    refs_by_address: dict[str, tuple[str, ...]] = {}
    for member in members:
        if key_cache is not None:
            fingerprint = key_cache.fingerprint_for_formula(
                member.address, member.normalized_formula
            )
        else:
            fingerprint = structural_fingerprint(
                member.normalized_formula,
                bound_address_keys=bound_address_keys,
                workbook_path=workbook_path,
            )
        if fingerprint is None:
            return _fallback_summary(
                "unparseable_formula",
                members=members,
                expected_member_keys=expected_member_keys,
            )
        skeleton, refs = fingerprint
        refs_by_address[member.address] = refs
        groups_by_skeleton[skeleton].append(member)

    address_to_helper = _address_to_helper_map(semantic_dependencies)
    group_records: list[FingerprintGroup] = []
    for skeleton, skeleton_members in groups_by_skeleton.items():
        member_partitions = _partition_members_by_ref_series_regime(
            skeleton_members,
            refs_by_address,
            address_to_series_id,
            address_to_helper,
            cluster_addresses,
        )
        for group_members in member_partitions:
            geometry_conflict = _unbound_ref_slot_geometry_conflict(
                group_members,
                refs_by_address,
                address_to_series_id,
            )
            if geometry_conflict is not None:
                return _fallback_summary(
                    geometry_conflict,
                    members=members,
                    expected_member_keys=expected_member_keys,
                )

            member_addresses = tuple(member.address for member in group_members)
            member_keys = {
                address: dict(expected_member_keys.get(address, {})) for address in member_addresses
            }
            relation_member_keys = member_keys

            ref_count = len(refs_by_address[member_addresses[0]])
            if any(len(refs_by_address[address]) != ref_count for address in member_addresses):
                return _fallback_summary(
                    "ref_count_mismatch",
                    members=members,
                    expected_member_keys=expected_member_keys,
                )

            ref_values_by_member: dict[str, list[dict[str, BindingKeyValue]]] = {}
            for address in member_addresses:
                ref_values = _ref_position_key_values(
                    address,
                    formula_by_address[address],
                    bound_address_keys,
                    workbook_path=workbook_path,
                    key_cache=key_cache,
                )
                if ref_values is None:
                    return _fallback_summary(
                        "missing_ref_key_values",
                        members=members,
                        expected_member_keys=expected_member_keys,
                    )
                ref_values_by_member[address] = ref_values

            relations: list[RefRelation] = []
            for ref_index in range(ref_count):
                ref_keys_by_member = {
                    address: ref_values_by_member[address][ref_index]
                    for address in member_addresses
                }
                ref_addresses = {
                    address: refs_by_address[address][ref_index] for address in member_addresses
                }
                series_id = _series_id_for_refs(tuple(ref_addresses.values()), address_to_series_id)
                resolution = _resolve_ref(
                    ref_addresses,
                    cluster_member_addresses=cluster_addresses,
                    semantic_dependencies=semantic_dependencies,
                    member_keys=relation_member_keys,
                )
                relations.append(
                    classify_ref_relation(
                        ref_index,
                        relation_member_keys,
                        ref_keys_by_member,
                        series_id=series_id,
                        resolution=resolution,
                    )
                )

            exemplar = group_members[0]
            group_records.append(
                FingerprintGroup(
                    skeleton_text=format_structural_skeleton(skeleton),
                    members=member_addresses,
                    exemplar=exemplar,
                    ref_relations=tuple(relations),
                    ref_addresses_by_member=tuple(
                        (address, refs_by_address[address]) for address in member_addresses
                    ),
                    ref_keys_by_member=tuple(
                        (address, tuple(ref_values_by_member[address]))
                        for address in member_addresses
                    ),
                )
            )

    key_space = _key_space_from_expected(expected_member_keys)
    key_to_column = _key_to_column_from_members(key_space, members, expected_member_keys)
    summary = ClusterFingerprintSummary(
        groups=tuple(group_records),
        key_space=key_space,
        key_to_column=key_to_column,
        fallback_reason=None,
        legacy_token_estimate=estimate_legacy_dump_tokens(members),
    )
    object.__setattr__(
        summary,
        "fingerprint_token_estimate",
        estimate_fingerprint_dump_tokens(summary),
    )
    return summary


def _format_key_value(value: LookupKey) -> str:
    if isinstance(value, tuple):
        return "(" + ", ".join(_format_key_value(item) for item in value) + ")"
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    return str(value)


def _format_mapping(table: Mapping[LookupKey, object]) -> str:
    items = ", ".join(f"{_format_key_value(key)}: {value}" for key, value in table.items())
    return "{" + items + "}"


def _format_col_letter_with_index(letter: str) -> str:
    """Format an Excel column letter with its 1-based index for geometry tuples."""
    index = xl_cell_utils.column_index_from_string(letter)
    return f"{letter}={index}"


def _format_col_by_mapping(table: Mapping[LookupKey, str]) -> str:
    """Format ``col by`` maps as ``{key: Letter=N, ...}`` for ``xl_index_ref``."""
    items = ", ".join(
        f"{_format_key_value(key)}: {_format_col_letter_with_index(str(value))}"
        for key, value in table.items()
    )
    return "{" + items + "}"


def _format_key_space_line(
    dimension_id: str,
    values: Sequence[BindingKeyValue],
    *,
    key_to_column: Mapping[BindingKeyValue, str] | None,
) -> str:
    if (
        values
        and all(_is_numeric(value) for value in values)
        and list(values) == list(range(int(values[0]), int(values[0]) + len(values)))  # type: ignore[arg-type]
    ):
        span = f"{values[0]}..{values[-1]}"
    else:
        span = "[" + ", ".join(_format_key_value(value) for value in values) + "]"
    suffix = ""
    if key_to_column and any(value in key_to_column for value in values):
        engine: dict[LookupKey, object] = {
            value: key_to_column[value] for value in values if value in key_to_column
        }
        suffix = f"   (engine columns: {_format_mapping(engine)})"
    return f"  {dimension_id}: {span}{suffix}"


def _format_ref_relation_lines(relation: RefRelation) -> list[str]:
    parts: list[str] = []
    if relation.series_id is not None:
        parts.append(f"series {relation.series_id}")
    for dimension_id, value in sorted(relation.fixed_keys.items()):
        parts.append(f"{dimension_id} = {_format_key_value(value)}")
    for dimension_id in relation.identity_dims:
        parts.append(f"{dimension_id} = member.{dimension_id}")
    for dimension_id, delta in sorted(relation.offsets.items()):
        if delta < 0:
            parts.append(f"{dimension_id} = member.{dimension_id} - {abs(delta)}")
        elif delta > 0:
            parts.append(f"{dimension_id} = member.{dimension_id} + {delta}")
        else:
            parts.append(f"{dimension_id} = member.{dimension_id}")
    for dimension_id, table in sorted(relation.lookups.items()):
        if dimension_id in relation.lookup_bases:
            lag_key_dims = [dim for dim in relation.identity_dims if dim != dimension_id]
            lag_key = relation.lookup_keys.get(
                dimension_id, lag_key_dims[0] if lag_key_dims else "KEY"
            )
            parts.append(
                f"{dimension_id} = member.{dimension_id} - lag, "
                f"lag by {lag_key} {_format_mapping(table)}"
            )
        else:
            key_dim_candidates = [dim for dim in relation.identity_dims if dim != dimension_id]
            key_dims = relation.lookup_keys.get(
                dimension_id, key_dim_candidates[0] if key_dim_candidates else "keys"
            )
            key_label = "(" + ", ".join(key_dims) + ")" if isinstance(key_dims, tuple) else key_dims
            parts.append(f"{dimension_id} = table[{key_label}] {_format_mapping(table)}")

    header = (
        f"  ref_{relation.ref_index}: " + "; ".join(parts)
        if parts
        else (f"  ref_{relation.ref_index}: (no key relation)")
    )
    lines = [header]
    if relation.tier == "explicit" and relation.explicit is not None:
        lines.append("         explicit member keys -> ref keys:")
        for member_combo, ref_combo in relation.explicit:
            member_text = ", ".join(f"{k}={v}" for k, v in member_combo)
            ref_text = ", ".join(f"{k}={v}" for k, v in ref_combo)
            lines.append(f"           {{{member_text}}} -> {{{ref_text}}}")

    resolution = relation.resolution
    if resolution.kind == "semantic_helper" and resolution.helper_name is not None:
        lines.append(
            f"         reads: semantic helper {resolution.helper_name}"
            + (f" via {resolution.call_form}" if resolution.call_form else "")
        )
    elif resolution.kind == "self_recurrence":
        lines.append("         reads: in-cluster self-recurrence (use this helper)")
    elif resolution.kind == "xl_cell" and resolution.address_template is not None:
        detail = f"         reads: xl_cell '{resolution.address_template}'"
        extras: list[str] = []
        for dim, pairs in resolution.row_by_dim:
            extras.append(f"row by {dim} {_format_mapping(dict(pairs))}")
        for dim, pairs in resolution.col_by_dim:
            extras.append(f"col by {dim} {_format_col_by_mapping(dict(pairs))}")
        if extras:
            detail += ", " + ", ".join(extras)
        lines.append(detail)
    return lines


def format_fingerprint_group_section(
    group: FingerprintGroup,
    *,
    group_index: int,
    key_space: Mapping[str, tuple[BindingKeyValue, ...]],
    key_to_column: Mapping[BindingKeyValue, str] | None,
    exemplar_keys: Mapping[str, BindingKeyValue] | None = None,
) -> str:
    total = len(group.members)
    lines = [
        f"## Fingerprint F{group_index} ({total} of {total} members)",
        "",
        f"`{group.skeleton_text}`",
        "",
        "Member key space (every member, no sampling):",
    ]
    for dimension_id, values in key_space.items():
        lines.append(_format_key_space_line(dimension_id, values, key_to_column=key_to_column))
    lines.append("")
    lines.append("Reference relations:")
    for relation in group.ref_relations:
        lines.extend(_format_ref_relation_lines(relation))
    exemplar = group.exemplar
    key_bits = ""
    if exemplar_keys:
        key_bits = "; " + ", ".join(
            f"{dimension_id}={_format_key_value(value)}"
            for dimension_id, value in sorted(exemplar_keys.items())
        )
    lines.append("")
    lines.append(f"Exemplar translation ({exemplar.address}{key_bits}):")
    lines.append("```python")
    lines.append(exemplar.python_source.strip())
    lines.append("```")
    return "\n".join(lines)


def format_cluster_fingerprint_dump(
    summary: ClusterFingerprintSummary,
    *,
    exemplar_keys_by_address: Mapping[str, Mapping[str, BindingKeyValue]] | None = None,
    key_vocabulary_yaml: str | None = None,
    member_metadata_yaml: str | None = None,
    dependency_stubs: str | None = None,
    helper_name: str | None = None,
) -> str:
    """Render fingerprint groups for inclusion in a cluster refactor prompt dump."""
    if summary.fallback_reason is not None:
        raise ValueError(
            f"cannot format fingerprint dump with fallback_reason={summary.fallback_reason!r}"
        )
    blocks: list[str] = []
    for index, group in enumerate(summary.groups, start=1):
        exemplar_keys = None
        if exemplar_keys_by_address is not None:
            exemplar_keys = exemplar_keys_by_address.get(group.exemplar.address)
        blocks.append(
            format_fingerprint_group_section(
                group,
                group_index=index,
                key_space=summary.key_space,
                key_to_column=summary.key_to_column,
                exemplar_keys=exemplar_keys,
            )
        )

    body = "\n\n".join(blocks)
    header = "Cluster to refactor:\n"
    if helper_name is not None:
        header = f"Cluster to refactor (helper_name={helper_name}):\n"
    sections = [header, body]
    if key_vocabulary_yaml is not None:
        sections.append(f"Key vocabulary:\n\n```yaml\n{key_vocabulary_yaml.strip()}\n```")
    if member_metadata_yaml is not None:
        sections.append(
            f"Member metadata (exemplars only):\n\n```yaml\n{member_metadata_yaml.strip()}\n```"
        )
    if dependency_stubs is not None:
        sections.append(f"Dependencies:\n\n```python\n{dependency_stubs.strip()}\n```")
    return "\n".join(section for section in sections if section).strip()


def estimate_legacy_dump_tokens(
    members: Sequence[MemberContext],
    *,
    member_limit: int = 30,
) -> int:
    """Rough token estimate for the legacy sampled member-source dump."""
    if not members:
        return 0
    if len(members) <= member_limit:
        selected = members
    else:
        # Match sample_indices_for_prompt spacing.
        step = (len(members) - 1) / (member_limit - 1)
        indices = sorted({round(i * step) for i in range(member_limit)})
        selected = tuple(members[index] for index in indices)
    text = "\n\n\n".join(member.python_source for member in selected)
    return max(1, len(text) // 4)


def estimate_fingerprint_dump_tokens(summary: ClusterFingerprintSummary) -> int:
    """Rough token estimate for the fingerprint dump (characters / 4)."""
    if summary.fallback_reason is not None:
        return 0
    text = format_cluster_fingerprint_dump(summary)
    return max(1, len(text) // 4)
