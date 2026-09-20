"""Surgical insert or replace of one series binding sidecar entry."""

from __future__ import annotations

import shutil
import tempfile
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

from ruamel.yaml import YAML

from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from excel_grapher.series_bindings.audit import audit_binding_resolutions, format_audit_findings
from excel_grapher.series_bindings.load import SeriesBindingsLoadError
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

_FILE_SUFFIXES = {".yaml", ".yml", ".json"}


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


def _round_trip_yaml() -> YAML:
    yaml_rt = YAML()
    yaml_rt.preserve_quotes = True
    yaml_rt.default_flow_style = False
    yaml_rt.width = 4096
    return yaml_rt


def _dump_binding_document(path: Path, document: Any) -> None:
    yaml_rt = _round_trip_yaml()
    with path.open("w", encoding="utf-8") as handle:
        yaml_rt.dump(document, handle)


def _empty_shard_document(
    *,
    schema_version: str,
    workbook: str | None,
) -> dict[str, Any]:
    document: dict[str, Any] = {"schema_version": schema_version, "series": []}
    if workbook is not None:
        document["workbook"] = workbook
    return document


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
            _dump_binding_document(
                path,
                _empty_shard_document(schema_version=version, workbook=workbook),
            )
        written.append(path)
    return written


def _group_placement(
    series: Mapping[str, Any],
) -> tuple[tuple[str, ...] | None, int | None]:
    groups = series.get("groups")
    if not isinstance(groups, list) or not groups:
        return None, None
    best_path: tuple[str, ...] | None = None
    best_order: int | None = None
    for item in groups:
        if not isinstance(item, dict):
            continue
        path_raw = item.get("path")
        if not isinstance(path_raw, list) or not path_raw:
            continue
        path = tuple(str(part) for part in path_raw)
        order = item.get("order")
        if not isinstance(order, int):
            continue
        if best_order is None or order < best_order:
            best_path = path
            best_order = order
    return best_path, best_order


def _insert_series(
    existing: Sequence[Mapping[str, Any]],
    series: Mapping[str, Any],
    *,
    replace: bool,
) -> list[Any]:
    series_id = str(series["id"])
    copied = dict(series)
    if replace:
        return [copied if item.get("id") == series_id else item for item in existing]
    path, order = _group_placement(series)
    if path is None or order is None:
        return [*existing, copied]
    result: list[Any] = []
    inserted = False
    for item in existing:
        item_path, item_order = _group_placement(item)
        if not inserted and item_path == path and item_order is not None and item_order > order:
            result.append(copied)
            inserted = True
        result.append(item)
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
) -> Any:
    if not path.is_file():
        return _empty_shard_document(schema_version=schema_version, workbook=workbook)
    yaml_rt = _round_trip_yaml()
    with path.open(encoding="utf-8") as handle:
        document = yaml_rt.load(handle)
    if not isinstance(document, dict):
        raise BindingUpsertError(f"Shard {path} must contain a mapping document")
    series = document.get("series")
    if not isinstance(series, list):
        raise BindingUpsertError(f"Shard {path} must contain a series list")
    return document


def _is_file_sidecar(path: Path) -> bool:
    if path.is_file():
        return True
    if path.is_dir():
        return False
    return path.suffix.lower() in _FILE_SUFFIXES


def _copy_binding_tree(source: Path, dest: Path) -> None:
    dest.mkdir(parents=True, exist_ok=True)
    for src in source.iterdir():
        if src.is_file():
            shutil.copy2(src, dest / src.name)


def _prepare_work_tree(
    bindings_path: Path,
    *,
    direction: Direction,
    workbook_name: str,
) -> tuple[Path, Path]:
    """Copy or bootstrap bindings into a temp tree. Return (work_root, work_shard)."""
    work_root = Path(tempfile.mkdtemp(prefix="excel-grapher-upsert-"))
    try:
        if _is_file_sidecar(bindings_path):
            work_shard = work_root / bindings_path.name
            if bindings_path.is_file():
                shutil.copy2(bindings_path, work_shard)
            return work_root, work_shard

        work_dir = work_root / "bindings"
        if bindings_path.is_dir():
            shutil.copytree(bindings_path, work_dir)
        else:
            work_dir.mkdir()
        bootstrap_binding_shards(
            work_dir,
            schema_version=CURRENT_SCHEMA_VERSION,
            workbook=workbook_name,
        )
        return work_root, work_dir / SHARD_FILENAMES[direction]
    except Exception:
        shutil.rmtree(work_root, ignore_errors=True)
        raise


def _commit_work_tree(
    *,
    bindings_path: Path,
    work_shard: Path,
    direction: Direction,
) -> Path:
    if _is_file_sidecar(bindings_path):
        bindings_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(work_shard, bindings_path)
        return bindings_path
    work_dir = work_shard.parent
    if bindings_path.exists() and bindings_path.is_dir():
        _copy_binding_tree(work_dir, bindings_path)
    else:
        shutil.copytree(work_dir, bindings_path)
    return bindings_path / SHARD_FILENAMES[direction]


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

    Schema load, occupancy, and resolution audit run against a temporary copy.
    The destination tree is written only after those checks succeed.

    Args:
        workbook: Workbook the sidecar describes.
        bindings_path: Existing sidecar file/directory, or the path to create.
        series: One series mapping to insert or replace.
        replace: Replace an existing series with the same id.
        dynamic_refs: Constraint-based OFFSET/INDEX/INDIRECT config. Ignored
            when `use_cached_dynamic_refs` is True.
        use_cached_dynamic_refs: Resolve dynamic refs from cached workbook
            values. Library default is True (same as `validate_bindings_workbook`).
            The CLI flag `--use-cached-dynamic-refs` defaults to False, matching
            other `excel-grapher bindings` commands.
        blank_ranges: Sheet-qualified rectangles omitted from the graph.

    Raises:
        BindingUpsertError: The series is not a valid insertion. The destination
            bindings tree is left unchanged, including when the path did not exist.
    """
    if "id" not in series or not isinstance(series.get("id"), str):
        raise BindingUpsertError("Series requires a string id")
    series_id = str(series["id"])
    direction = _require_direction(series)
    workbook_name = workbook.name

    work_root: Path | None = None
    try:
        work_root, work_shard = _prepare_work_tree(
            bindings_path,
            direction=direction,
            workbook_name=workbook_name,
        )
        work_bindings = work_shard if _is_file_sidecar(bindings_path) else work_shard.parent
        document = _load_shard_document(
            work_shard,
            schema_version=CURRENT_SCHEMA_VERSION,
            workbook=workbook_name,
        )
        existing_series = list(document.get("series") or [])
        existing_ids = [
            str(item.get("id", "")) for item in existing_series if isinstance(item, Mapping)
        ]
        replaced = series_id in existing_ids
        if replaced and not replace:
            raise BindingUpsertError(
                f"Series id {series_id!r} already exists; pass replace=True to update it"
            )
        if replace and not replaced:
            raise BindingUpsertError(f"Series id {series_id!r} does not exist to replace")

        document["series"] = _insert_series(existing_series, series, replace=replaced)
        _dump_binding_document(work_shard, document)

        try:
            result = validate_bindings_workbook(
                workbook,
                work_bindings,
                dynamic_refs=dynamic_refs,
                use_cached_dynamic_refs=use_cached_dynamic_refs,
                blank_ranges=blank_ranges,
            )
        except (SeriesBindingsLoadError, SeriesBindingsSchemaError) as exc:
            raise BindingUpsertError(str(exc)) from exc
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

        shard_path = _commit_work_tree(
            bindings_path=bindings_path,
            work_shard=work_shard,
            direction=direction,
        )
    finally:
        if work_root is not None:
            shutil.rmtree(work_root, ignore_errors=True)

    return BindingUpsertResult(
        series_id=series_id,
        shard_path=shard_path,
        replaced=replaced,
        direction=direction,
    )
