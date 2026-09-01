"""Semantic binding record hints and refactor identifier validation."""

from __future__ import annotations

import ast
import builtins
import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from typing import Any

from excel_grapher.exporter.pass1.bindings import BindingKeyValue

_SNAKE_CASE_PATTERN = re.compile(r"^[a-z][a-z0-9_]*$")

RESERVED_HELPER_NAMES: frozenset[str] = frozenset(
    {
        "_address_to_func_name",
        "_resolve_formula",
        "_RESOLVED_FORMULAS",
        "_ADDRESS_DISPATCH",
        "_SYMBOL_DISPATCH",
    }
)


@dataclass(frozen=True)
class BindingRecordHints:
    binding_keys: dict[str, BindingKeyValue] | None = None
    binding_record: dict[str, BindingKeyValue] | None = None

    def to_payload(self) -> dict[str, object]:
        payload: dict[str, object] = {}
        if self.binding_keys:
            payload["binding_keys"] = dict(self.binding_keys)
        if self.binding_record:
            payload["binding_record"] = dict(self.binding_record)
        return payload

    def has_hints(self) -> bool:
        return bool(self.to_payload())


def binding_record_hints_from_cell(
    cell: Mapping[str, Any] | None,
) -> BindingRecordHints:
    if cell is None:
        return BindingRecordHints()
    key = cell.get("key")
    record = cell.get("record")
    binding_keys = dict(key) if isinstance(key, Mapping) and key else None
    binding_record = dict(record) if isinstance(record, Mapping) and record else None
    return BindingRecordHints(
        binding_keys=binding_keys,
        binding_record=binding_record,
    )


def cluster_binding_naming_hints(
    member_hints: tuple[BindingRecordHints, ...],
) -> dict[str, object]:
    cluster_payload: dict[str, object] = {}
    member_payloads = [hint.to_payload() for hint in member_hints if hint.has_hints()]
    if member_payloads:
        cluster_payload["members"] = member_payloads
    return cluster_payload


def sole_series_id_for_addresses(
    addresses: Sequence[str],
    address_to_series_id: Mapping[str, str],
) -> str:
    """Return the single series_id covering every address in a refactor unit."""
    missing = sorted(address for address in addresses if address not in address_to_series_id)
    if missing:
        raise ValueError(f"addresses missing series_id mapping: {missing}")
    series_ids = {address_to_series_id[address] for address in addresses}
    if len(series_ids) != 1:
        raise ValueError(
            f"expected exactly one series_id for refactor unit, got {sorted(series_ids)}"
        )
    return next(iter(series_ids))


def _dedupe_helper_name(base: str, taken: set[str]) -> str:
    """Return ``base`` or ``base_2``, ``base_3``, … not present in ``taken``."""
    if base not in taken:
        return base
    index = 2
    while True:
        candidate = f"{base}_{index}"
        if candidate not in taken:
            return candidate
        index += 1


def allocate_schedule_helper_names(
    unit_members: Sequence[Sequence[str]],
    address_to_series_id: Mapping[str, str],
    *,
    existing_names: frozenset[str] = frozenset(),
) -> tuple[str, ...]:
    """Lock a unique helper name for each schedule unit before any LLM call.

    Sole units for a series keep the bare ``series_id`` when free. When a
    series is sliced into multiple units, later peels receive deterministic
    ``series_id_2``, ``series_id_3``, … suffixes. Every unit skips names
    already reserved by earlier units or present in ``existing_names``;
    existing helpers are never overwritten.
    """
    series_ids = tuple(
        sole_series_id_for_addresses(members, address_to_series_id) for members in unit_members
    )
    reserved: set[str] = set(existing_names)
    allocated: list[str] = []

    for series_id in series_ids:
        name = _dedupe_helper_name(series_id, reserved)
        validate_semantic_identifier(name, existing_names=frozenset(reserved))
        allocated.append(name)
        reserved.add(name)

    return tuple(allocated)


def validate_semantic_identifier(
    name: str,
    *,
    existing_names: frozenset[str],
    allow_name: str | None = None,
    reserved_names: frozenset[str] = RESERVED_HELPER_NAMES,
) -> None:
    if not name.isidentifier():
        raise ValueError(f"name is not a valid identifier: {name!r}")
    if not _SNAKE_CASE_PATTERN.fullmatch(name):
        raise ValueError(f"name must be snake_case: {name!r}")
    if name.startswith("cell_"):
        raise ValueError(f"name must not use cell_* prefix: {name!r}")
    if name in builtins.__dict__:
        raise ValueError(f"name collides with Python builtin: {name!r}")
    if name in reserved_names:
        raise ValueError(f"name collides with reserved runtime symbol: {name!r}")
    if name in existing_names and name != allow_name:
        raise ValueError(f"name {name!r} collides with existing function")


def _is_semantic_helper_def(function_def: ast.FunctionDef) -> bool:
    if function_def.name.startswith("cell_") or function_def.name.startswith("_"):
        return False
    if not function_def.args.args:
        return False
    return function_def.args.args[0].arg == "ctx"


def collect_semantic_helper_names(source: str) -> frozenset[str]:
    module = ast.parse(source)
    return frozenset(
        node.name
        for node in module.body
        if isinstance(node, ast.FunctionDef) and _is_semantic_helper_def(node)
    )


def semantic_helpers_available_for_calls(
    source: str,
    existing_names: frozenset[str],
) -> frozenset[str]:
    """Semantic helpers defined in source plus already-allocated helper names."""
    defined = collect_semantic_helper_names(source)
    allocated = frozenset(
        name
        for name in existing_names
        if name not in RESERVED_HELPER_NAMES
        and not name.startswith("cell_")
        and not name.startswith("_")
    )
    return defined | allocated
