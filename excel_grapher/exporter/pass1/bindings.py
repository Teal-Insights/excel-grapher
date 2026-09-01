"""Binding key vocabulary and address-key resolution for internals refactor."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from excel_grapher.series_bindings import load_series_bindings
from excel_grapher.series_bindings.normalize import effective_dimension_id
from excel_grapher.series_bindings.setter_codegen import dimension_id_to_param_name
from excel_grapher.series_bindings.types import Scalar, WorkbookSeriesBindings

BindingKeyValue = str | int | float | bool


@dataclass(frozen=True)
class KeyConceptSpec:
    """Vocabulary entry for one cell-scoped key dimension.

    ``dimension_id`` is the effective binding key used for parameter naming,
    member keys, clustering, and dispatch. ``concept`` is the SDMX-style
    semantic category the dimension references.
    """

    dimension_id: str
    concept: str
    dtype: str
    suggested_param_name: str


def resolve_dimension_key(
    name: str,
    vocabulary: Sequence[KeyConceptSpec],
) -> str:
    """Resolve a binding key name to an effective dimension id.

    Prefers an exact ``dimension_id`` match. Falls back to ``concept`` only when
    exactly one vocabulary entry uses that concept.
    """
    by_id = {item.dimension_id: item for item in vocabulary}
    if name in by_id:
        return name
    concept_matches = [item for item in vocabulary if item.concept == name]
    if len(concept_matches) == 1:
        return concept_matches[0].dimension_id
    if len(concept_matches) > 1:
        ids = sorted(item.dimension_id for item in concept_matches)
        raise ValueError(
            f"ambiguous binding key {name!r}: matches dimension ids {ids}; "
            "use an explicit dimension id"
        )
    raise ValueError(f"unknown binding key: {name!r}")


def key_concept_vocabulary_from_bindings(
    bindings: WorkbookSeriesBindings,
) -> tuple[KeyConceptSpec, ...]:
    """Aggregate cell-scoped key dimensions from already-loaded series bindings."""
    concept_scheme = bindings.get("concept_scheme") or {}
    concept_dtypes: dict[str, str] = {}
    for concept in concept_scheme.get("concepts") or []:
        if not isinstance(concept, dict):
            continue
        concept_id = concept.get("id")
        dtype = concept.get("dtype")
        if isinstance(concept_id, str) and dtype is not None:
            concept_dtypes[concept_id] = str(dtype)

    seen: dict[str, KeyConceptSpec] = {}
    for series in bindings.get("series") or []:
        if not isinstance(series, dict):
            continue
        structure = series.get("structure") or {}
        for dimension in structure.get("dimensions") or []:
            if not isinstance(dimension, dict):
                continue
            if dimension.get("role") != "key":
                continue
            if dimension.get("scope") == "series":
                continue
            concept = dimension.get("concept")
            if not isinstance(concept, str):
                continue
            dimension_id = effective_dimension_id(dimension)
            if not dimension_id or dimension_id in seen:
                continue
            dtype = concept_dtypes.get(concept) or dimension.get("dtype") or "string"
            seen[dimension_id] = KeyConceptSpec(
                dimension_id=dimension_id,
                concept=concept,
                dtype=str(dtype),
                suggested_param_name=dimension_id_to_param_name(dimension_id),
            )
    return tuple(sorted(seen.values(), key=lambda item: item.dimension_id))


def load_key_concept_vocabulary(bindings_path: Path) -> tuple[KeyConceptSpec, ...]:
    """Aggregate cell-scoped key dimensions from ``bindings/*.bindings.yaml``."""
    return key_concept_vocabulary_from_bindings(load_series_bindings(bindings_path))


def build_bound_address_keys(
    input_series: Sequence[Mapping[str, Any]],
    output_series: Sequence[Mapping[str, Any]],
    internal_series: Sequence[Mapping[str, Any]] = (),
    *,
    constant_series: Sequence[Mapping[str, Any]] = (),
) -> dict[str, dict[str, BindingKeyValue]]:
    """Index every bound cell address to its coerced cell-scope binding keys.

    Constant series are folded in first (lowest priority) so a keyed reader-only
    leaf still contributes its per-cell keys. Input/output/internal series
    overwrite the constant baseline for any shared address so formula-bearing
    ownership keeps precedence.
    """
    index: dict[str, dict[str, BindingKeyValue]] = {}
    for series_list in (constant_series, input_series, output_series, internal_series):
        for series in series_list:
            for cell in series["cells"]:
                index[str(cell["address"])] = _coerce_binding_keys(cell["key"])
    return index


def series_cell_owners(
    series_list: Sequence[Mapping[str, Any]],
) -> dict[str, tuple[str, ...]]:
    """Map each series cell address to the series ids that claim it."""
    owners: dict[str, list[str]] = {}
    for series in series_list:
        series_id = series.get("id")
        if not isinstance(series_id, str) or not series_id:
            continue
        for cell in series.get("cells", []):
            address = cell.get("address")
            if isinstance(address, str):
                owners.setdefault(address, []).append(series_id)
    return {address: tuple(series_ids) for address, series_ids in owners.items()}


def _unique_series_id_by_address(
    series_list: Sequence[Mapping[str, Any]],
    *,
    ownership_kind: str,
) -> dict[str, str]:
    owners_by_address = series_cell_owners(series_list)
    duplicates = {
        address: series_ids
        for address, series_ids in owners_by_address.items()
        if len(series_ids) > 1
    }
    if duplicates:
        sample_address, sample_series_ids = min(duplicates.items())
        raise ValueError(
            f"{ownership_kind} series cell address must map to exactly one series_id; "
            f"got {sample_address!r} in {list(sample_series_ids)}"
            + (
                f" and {len(duplicates) - 1} more duplicate address(es)"
                if len(duplicates) > 1
                else ""
            )
        )
    return {
        address: series_ids[0]
        for address, series_ids in owners_by_address.items()
        if len(series_ids) == 1
    }


def build_address_to_series_id(
    internal_series: Sequence[Mapping[str, Any]],
    *,
    output_series: Sequence[Mapping[str, Any]] = (),
    input_series: Sequence[Mapping[str, Any]] = (),
    constant_series: Sequence[Mapping[str, Any]] = (),
) -> dict[str, str]:
    """Map cell addresses to refactor partition series ids.

    Internal ownership wins. Addresses without an internal owner fall back to
    constant (reader-only leaf), then public output, then input binding series
    ids, so ``series_ast`` / ``series`` clustering can keep time-sweep public
    series together and mechanical synthesis can claim ``read_<id>(ctx)`` sites
    for constant operands.
    """
    address_to_series_id = _unique_series_id_by_address(
        internal_series,
        ownership_kind="internal",
    )
    for ownership_kind, series_list in (
        ("constant", constant_series),
        ("output", output_series),
        ("input", input_series),
    ):
        public_ids = _unique_series_id_by_address(
            series_list,
            ownership_kind=ownership_kind,
        )
        for address, series_id in public_ids.items():
            address_to_series_id.setdefault(address, series_id)
    return address_to_series_id


def expected_keys_for_address(
    address: str,
    *,
    bound_address_keys: Mapping[str, Mapping[str, BindingKeyValue]],
) -> dict[str, BindingKeyValue]:
    """Return the coerced cell-scope binding keys declared for one address."""
    bound = bound_address_keys.get(address)
    if bound is None:
        return {}
    return _coerce_binding_keys(bound)


def _coerce_binding_keys(keys: Mapping[str, Scalar]) -> dict[str, BindingKeyValue]:
    coerced: dict[str, BindingKeyValue] = {}
    for key, value in keys.items():
        if isinstance(value, (bool, str, int, float)):
            coerced[key] = value
    return coerced


def varying_key_concepts(
    addresses: Sequence[str],
    *,
    bound_address_keys: Mapping[str, Mapping[str, BindingKeyValue]],
) -> frozenset[str]:
    keys_by_address = [
        expected_keys_for_address(address, bound_address_keys=bound_address_keys)
        for address in addresses
    ]
    dimension_ids = {key for keys in keys_by_address for key in keys}
    return frozenset(
        dimension_id
        for dimension_id in dimension_ids
        if len({keys.get(dimension_id) for keys in keys_by_address}) > 1
    )


def expected_member_keys_for_cluster(
    addresses: Sequence[str],
    *,
    bound_address_keys: Mapping[str, Mapping[str, BindingKeyValue]],
) -> dict[str, dict[str, BindingKeyValue]]:
    varying = varying_key_concepts(
        addresses,
        bound_address_keys=bound_address_keys,
    )
    return {
        address: {
            dimension_id: expected_keys_for_address(
                address,
                bound_address_keys=bound_address_keys,
            )[dimension_id]
            for dimension_id in varying
        }
        for address in addresses
    }


def format_binding_key_literal(value: BindingKeyValue) -> str:
    if isinstance(value, str):
        return repr(value)
    if isinstance(value, bool):
        return "True" if value else "False"
    return str(value)


def render_literal_helper_call(
    helper_name: str,
    parameters: Sequence[tuple[str, str]],
    keys: Mapping[str, BindingKeyValue],
) -> str:
    """Render ``helper(ctx, param=literal, ...)`` for one collapsed cell address."""
    kwargs = ", ".join(
        f"{name}={format_binding_key_literal(keys[dimension_id])}"
        for name, dimension_id in parameters
    )
    return f"{helper_name}(ctx, {kwargs})"


def helper_parameters_for_varying_keys(
    varying_dimension_ids: frozenset[str],
    vocabulary: Sequence[KeyConceptSpec],
) -> tuple[KeyConceptSpec, ...]:
    by_id = {item.dimension_id: item for item in vocabulary}
    missing = sorted(
        dimension_id for dimension_id in varying_dimension_ids if dimension_id not in by_id
    )
    if missing:
        raise ValueError(f"unknown varying key dimensions: {missing}")
    return tuple(by_id[dimension_id] for dimension_id in sorted(varying_dimension_ids))
