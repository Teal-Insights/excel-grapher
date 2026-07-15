"""Reverse address/range → semantic reader call mapping (Phase 1b).

Builds on Phase 1 leaf resolution (`_LEAF_INDEX_*` / `resolved["leaves"]`) so
consumers (notably QCraft fingerprints) can describe data-layer reads as
`read_<id>(ctx, …)` / `read_<id>_range(ctx)` instead of bare Excel geometry.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from pathlib import Path
from typing import Literal, NotRequired, TypedDict

from excel_grapher.core.address_keys import normalize_key, split_address_on_colon
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.normalize import has_input_direction
from excel_grapher.series_bindings.resolve import resolve_series_bindings
from excel_grapher.series_bindings.setter_codegen import (
    _reader_function_name,
    _should_emit_reader,
    _should_emit_reader_range,
    dimension_id_to_param_name,
)
from excel_grapher.series_bindings.types import Scalar, SeriesResolution, WorkbookSeriesBindings

ReaderLeafKind = Literal["keyed", "scalar", "address_keyed"]
ReaderCallMode = Literal["reader", "reader_range", "xl_cell", "xl_range"]
ReaderFallbackReason = Literal[
    "unbound",
    "ambiguous_owner",
    "not_binding_aligned_range",
]


class ReaderLeafEntry(TypedDict):
    """One uniquely-owned input leaf mapped to its semantic reader call."""

    series_id: str
    reader: str
    keys: dict[str, Scalar]
    kwargs: dict[str, Scalar]
    kind: ReaderLeafKind
    call_form: str


class ReaderRangeEntry(TypedDict):
    """One binding-aligned `data_range` mapped to its `read_*_range` helper."""

    series_id: str
    reader: str
    data_range: str
    call_form: str


class ReaderIndex(TypedDict):
    """Machine-readable reverse map from Excel geometry to reader call forms."""

    leaves: dict[str, ReaderLeafEntry]
    ranges: dict[str, ReaderRangeEntry]
    ambiguous: tuple[str, ...]


class ReaderCallResolution(TypedDict):
    """Result of resolving a cell or range reference to a reader (or fallback)."""

    mode: ReaderCallMode
    call_form: str
    reason: str | None
    series_id: str | None
    reader: str | None
    keys: NotRequired[dict[str, Scalar]]
    kwargs: NotRequired[dict[str, Scalar]]
    kind: NotRequired[ReaderLeafKind]


def format_reader_call_form(
    reader: str,
    *,
    kwargs: Mapping[str, object] | None = None,
    address: str | None = None,
    range_reader: bool = False,
) -> str:
    """Build a bindings-aware call-form string for fingerprints / prompts.

    Args:
        reader: Generated reader function name (`read_<id>` or `read_<id>_range`).
        kwargs: Snake_case dimension kwargs for keyed readers.
        address: Address kwarg for duplicate-key (`requires_address`) readers.
        range_reader: When True, emit the no-arg range form `reader(ctx)`.

    Returns:
        Call form such as `read_gdp(ctx, time_period=2020)` or
        `read_gdp_range(ctx)`.
    """
    if range_reader or (not kwargs and address is None):
        return f"{reader}(ctx)"
    if address is not None:
        return f"{reader}(ctx, address={address!r})"
    assert kwargs is not None
    parts = ", ".join(f"{name}={value!r}" for name, value in kwargs.items())
    return f"{reader}(ctx, {parts})"


def _leaf_kind(resolved: SeriesResolution, key_fields: list[str]) -> ReaderLeafKind:
    if resolved["requires_address"]:
        return "address_keyed"
    if not key_fields:
        return "scalar"
    return "keyed"


def _leaf_entry(
    *,
    series_id: str,
    reader: str,
    address: str,
    key: Mapping[str, Scalar],
    key_fields: list[str],
    kind: ReaderLeafKind,
) -> ReaderLeafEntry:
    keys = {field: key[field] for field in key_fields if field in key}
    if kind == "address_keyed":
        kwargs: dict[str, Scalar] = {}
        call_form = format_reader_call_form(reader, address=address)
    elif kind == "scalar":
        kwargs = {}
        call_form = format_reader_call_form(reader)
    else:
        kwargs = {dimension_id_to_param_name(field): keys[field] for field in key_fields}
        call_form = format_reader_call_form(reader, kwargs=kwargs)
    return {
        "series_id": series_id,
        "reader": reader,
        "keys": keys,
        "kwargs": kwargs,
        "kind": kind,
        "call_form": call_form,
    }


def build_reader_index(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str,
    export_addresses: Iterable[str] | None = None,
) -> ReaderIndex:
    """Build reverse address/range indexes from Phase 1 leaf resolution.

    Leaves owned by more than one input series are recorded in `ambiguous` and
    omitted from `leaves` so consumers fall back to `xl_cell` with a clear
    reason. Range entries are emitted only when codegen would emit
    `read_<id>_range`.
    """
    report = resolve_series_bindings(
        graph,
        bindings,
        workbook=workbook,
        direction="input",
        export_addresses=export_addresses,
    )
    by_id = {
        s["id"]: s
        for s in bindings.get("series", [])
        if isinstance(s, dict) and has_input_direction(s)
    }

    owners: dict[str, list[ReaderLeafEntry]] = {}
    ranges: dict[str, ReaderRangeEntry] = {}

    for resolved in report["series"]:
        series = by_id.get(resolved["series_id"])
        if series is None or not _should_emit_reader(resolved):
            continue
        reader = _reader_function_name(series, resolved)
        key_fields = [str(field) for field in (series.get("key") or [])]
        kind = _leaf_kind(resolved, key_fields)
        for leaf in resolved["leaves"]:
            address = normalize_key(leaf["address"])
            entry = _leaf_entry(
                series_id=resolved["series_id"],
                reader=reader,
                address=address,
                key=leaf["key"],
                key_fields=key_fields,
                kind=kind,
            )
            owners.setdefault(address, []).append(entry)

        if _should_emit_reader_range(series, resolved):
            data_range = series.get("data_range")
            if isinstance(data_range, str) and data_range:
                normalized_range = normalize_key(data_range)
                range_reader = f"{reader}_range"
                ranges[normalized_range] = {
                    "series_id": resolved["series_id"],
                    "reader": range_reader,
                    "data_range": normalized_range,
                    "call_form": format_reader_call_form(range_reader, range_reader=True),
                }

    leaves: dict[str, ReaderLeafEntry] = {}
    ambiguous: list[str] = []
    for address, entries in sorted(owners.items()):
        # Distinct series ids — duplicate identical entries from one series are fine.
        series_ids = {entry["series_id"] for entry in entries}
        if len(series_ids) != 1:
            ambiguous.append(address)
            continue
        leaves[address] = entries[0]

    return {
        "leaves": leaves,
        "ranges": ranges,
        "ambiguous": tuple(ambiguous),
    }


def _is_range_ref(ref: str) -> bool:
    return split_address_on_colon(ref) is not None


def resolve_reader_ref(
    ref: str,
    *,
    graph: DependencyGraph | None = None,
    bindings: WorkbookSeriesBindings | None = None,
    workbook: Path | str | None = None,
    index: ReaderIndex | None = None,
    export_addresses: Iterable[str] | None = None,
) -> ReaderCallResolution:
    """Resolve a cell or range reference to a reader call form or Excel fallback.

    Prefer passing a prebuilt `index` when resolving many refs. Otherwise supply
    `graph`, `bindings`, and `workbook` so the index can be built from Phase 1
    leaf resolution.
    """
    if index is None:
        if graph is None or bindings is None or workbook is None:
            raise ValueError("resolve_reader_ref requires index= or graph/bindings/workbook")
        index = build_reader_index(
            graph,
            bindings,
            workbook=workbook,
            export_addresses=export_addresses,
        )

    normalized = normalize_key(ref)
    if _is_range_ref(normalized):
        range_entry = index["ranges"].get(normalized)
        if range_entry is not None:
            return {
                "mode": "reader_range",
                "call_form": range_entry["call_form"],
                "reason": None,
                "series_id": range_entry["series_id"],
                "reader": range_entry["reader"],
            }
        return {
            "mode": "xl_range",
            "call_form": f"xl_range(ctx, {normalized!r})",
            "reason": "not_binding_aligned_range",
            "series_id": None,
            "reader": None,
        }

    if normalized in index["ambiguous"]:
        return {
            "mode": "xl_cell",
            "call_form": f"xl_cell(ctx, {normalized!r})",
            "reason": "ambiguous_owner",
            "series_id": None,
            "reader": None,
        }

    leaf = index["leaves"].get(normalized)
    if leaf is None:
        return {
            "mode": "xl_cell",
            "call_form": f"xl_cell(ctx, {normalized!r})",
            "reason": "unbound",
            "series_id": None,
            "reader": None,
        }

    return {
        "mode": "reader",
        "call_form": leaf["call_form"],
        "reason": None,
        "series_id": leaf["series_id"],
        "reader": leaf["reader"],
        "keys": leaf["keys"],
        "kwargs": leaf["kwargs"],
        "kind": leaf["kind"],
    }


def reader_index_as_discovery_dicts(
    index: ReaderIndex,
) -> tuple[dict[str, dict[str, object]], dict[str, dict[str, object]]]:
    """Convert a `ReaderIndex` into JSON-literal-friendly discovery payloads."""
    leaves = {address: dict(entry) for address, entry in index["leaves"].items()}
    ranges = {address: dict(entry) for address, entry in index["ranges"].items()}
    return leaves, ranges
