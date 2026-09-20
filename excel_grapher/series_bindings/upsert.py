"""Surgical insert or replace of one series binding sidecar entry."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

import yaml

from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from excel_grapher.series_bindings.audit import audit_binding_resolutions, format_audit_findings
from excel_grapher.series_bindings.load import (
    SeriesBindingsLoadError,
    parse_bindings_file,
)
from excel_grapher.series_bindings.occupancy import Direction, binding_direction
from excel_grapher.series_bindings.schema import SeriesBindingsSchemaError
from excel_grapher.series_bindings.versions import CURRENT_SCHEMA_VERSION
from excel_grapher.series_bindings.workflow import validate_bindings_workbook

SHARD_FILENAMES: dict[Direction, str] = {
    "input": "inputs.bindings.yaml",
    "output": "outputs.bindings.yaml",
    "internal": "internals.bindings.yaml",
    "constant": "constants.bindings.yaml",
}


class BindingUpsertError(ValueError):
    """Raised when a single-series upsert is not a valid insertion."""


@dataclass(frozen=True)
class BindingUpsertResult:
    """Filesystem result of a successful single-series upsert."""

    series_id: str
    shard_path: Path
    replaced: bool
    direction: Direction

    @property
    def action(self) -> Literal["inserted", "replaced"]:
        """Whether this call inserted a new id or replaced an existing one."""
        return "replaced" if self.replaced else "inserted"


def bootstrap_binding_shards(
    directory: Path,
    *,
    schema_version: str | None = None,
    workbook: str | None = None,
) -> list[Path]:
    """Create empty four-file sidecar shards when missing.

    Existing shard files are left unchanged. Empty `series: []` placeholders
    are valid excel-grapher documents.
    """
    directory.mkdir(parents=True, exist_ok=True)
    version = schema_version or CURRENT_SCHEMA_VERSION
    written: list[Path] = []
    for filename in SHARD_FILENAMES.values():
        path = directory / filename
        if not path.is_file():
            document: dict[str, Any] = {"schema_version": version, "series": []}
            if workbook is not None:
                document["workbook"] = workbook
            path.write_text(
                yaml.safe_dump(document, sort_keys=False, allow_unicode=True),
                encoding="utf-8",
            )
        written.append(path)
    return written


def _dump_binding_document(path: Path, document: Mapping[str, Any]) -> None:
    path.write_text(
        yaml.safe_dump(dict(document), sort_keys=False, allow_unicode=True),
        encoding="utf-8",
    )


def _empty_shard_document(
    *,
    schema_version: str,
    workbook: str | None,
) -> dict[str, Any]:
    document: dict[str, Any] = {"schema_version": schema_version, "series": []}
    if workbook is not None:
        document["workbook"] = workbook
    return document


def _group_placement(
    series: Mapping[str, Any],
) -> tuple[tuple[str, ...] | None, int | None]:
    groups = series.get("groups")
    if not isinstance(groups, list) or not groups:
        return None, None
    first = groups[0]
    if not isinstance(first, dict):
        return None, None
    path_raw = first.get("path")
    if not isinstance(path_raw, list) or not path_raw:
        return None, None
    path = tuple(str(part) for part in path_raw)
    order = first.get("order")
    return path, int(order) if isinstance(order, int) else None


def _insert_series(
    existing: Sequence[Mapping[str, Any]],
    series: Mapping[str, Any],
    *,
    replace: bool,
) -> list[dict[str, Any]]:
    series_id = str(series["id"])
    copied = dict(series)
    if replace:
        return [copied if item.get("id") == series_id else dict(item) for item in existing]
    path, order = _group_placement(series)
    if path is None or order is None:
        return [dict(item) for item in existing] + [copied]
    result: list[dict[str, Any]] = []
    inserted = False
    for item in existing:
        item_path, item_order = _group_placement(item)
        if not inserted and item_path == path and item_order is not None and item_order > order:
            result.append(copied)
            inserted = True
        result.append(dict(item))
    if not inserted:
        result.append(copied)
    return result


def _require_direction(series: Mapping[str, Any]) -> Direction:
    direction = binding_direction(dict(series))
    if direction is None:
        raise BindingUpsertError(
            "Series must declare exactly one of input, output, internal, or constant"
        )
    return direction


def _load_shard_document(
    path: Path,
    *,
    schema_version: str,
    workbook: str | None,
) -> dict[str, Any]:
    if not path.is_file():
        return _empty_shard_document(schema_version=schema_version, workbook=workbook)
    document = parse_bindings_file(path)
    series = document.get("series")
    if not isinstance(series, list):
        raise BindingUpsertError(f"Shard {path} must contain a series list")
    return document


def upsert_series_binding(
    workbook: Path,
    bindings_path: Path,
    series: Mapping[str, Any],
    *,
    replace: bool = False,
    dynamic_refs: DynamicRefConfig | None = None,
    use_cached_dynamic_refs: bool = True,
    blank_ranges: Sequence[str] | None = None,
) -> BindingUpsertResult:
    """Insert or replace one series after fail-closed checks.

    The series is written only when schema load, cell occupancy, and
    resolution audit all succeed. On failure the previous shard bytes are
    restored.
    """
    if "id" not in series or not isinstance(series.get("id"), str):
        raise BindingUpsertError("Series requires a string id")
    series_id = str(series["id"])
    direction = _require_direction(series)
    workbook_name = workbook.name

    if bindings_path.is_file():
        shard_path = bindings_path
        document = _load_shard_document(
            shard_path,
            schema_version=CURRENT_SCHEMA_VERSION,
            workbook=workbook_name,
        )
    else:
        bootstrap_binding_shards(
            bindings_path,
            schema_version=CURRENT_SCHEMA_VERSION,
            workbook=workbook_name,
        )
        shard_path = bindings_path / SHARD_FILENAMES[direction]
        document = _load_shard_document(
            shard_path,
            schema_version=CURRENT_SCHEMA_VERSION,
            workbook=workbook_name,
        )

    existing_series = list(document.get("series") or [])
    existing_ids = [str(item.get("id", "")) for item in existing_series if isinstance(item, dict)]
    replaced = series_id in existing_ids
    if replaced and not replace:
        raise BindingUpsertError(
            f"Series id {series_id!r} already exists; pass replace=True to update it"
        )
    if replace and not replaced:
        raise BindingUpsertError(f"Series id {series_id!r} does not exist to replace")

    document["series"] = _insert_series(existing_series, series, replace=replaced)
    previous = shard_path.read_text(encoding="utf-8") if shard_path.is_file() else None
    _dump_binding_document(shard_path, document)
    try:
        result = validate_bindings_workbook(
            workbook,
            bindings_path,
            dynamic_refs=dynamic_refs,
            use_cached_dynamic_refs=use_cached_dynamic_refs,
            blank_ranges=blank_ranges,
        )
        report = result["report"]
        if not report["ok"]:
            errors = [issue for issue in report["issues"] if issue["level"] == "error"]
            preview = "\n".join(str(issue) for issue in errors[:5]) or "validation failed"
            raise BindingUpsertError(f"Upsert failed validation:\n{preview}")
        audit = audit_binding_resolutions(
            result["graph"],
            result["bindings"],
            workbook=workbook,
        )
        if not audit.ok:
            preview = "\n".join(format_audit_findings(audit.findings)[:12])
            raise BindingUpsertError(f"Upsert failed resolution audit:\n{preview}")
    except (SeriesBindingsLoadError, SeriesBindingsSchemaError) as exc:
        if previous is None:
            shard_path.unlink(missing_ok=True)
        else:
            shard_path.write_text(previous, encoding="utf-8")
        raise BindingUpsertError(str(exc)) from exc
    except BindingUpsertError:
        if previous is None:
            shard_path.unlink(missing_ok=True)
        else:
            shard_path.write_text(previous, encoding="utf-8")
        raise
    except Exception:
        if previous is None:
            shard_path.unlink(missing_ok=True)
        else:
            shard_path.write_text(previous, encoding="utf-8")
        raise
    return BindingUpsertResult(
        series_id=series_id,
        shard_path=shard_path,
        replaced=replaced,
        direction=direction,
    )
