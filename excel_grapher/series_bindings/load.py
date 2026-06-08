from __future__ import annotations

import json
from pathlib import Path
from typing import Any, cast

import yaml

from excel_grapher.series_bindings.normalize import merge_series_entries, normalize_series_entry
from excel_grapher.series_bindings.schema import validate_bindings_document
from excel_grapher.series_bindings.types import WorkbookSeriesBindings

BINDINGS_GLOB_NAMES = ("*.bindings.yaml", "*.bindings.yml", "*.bindings.json")


class SeriesBindingsLoadError(ValueError):
    """Raised when binding files cannot be parsed or merged."""


def _parse_raw_text(text: str, *, path: Path) -> Any:
    suffix = path.suffix.lower()
    name = path.name.lower()
    if suffix == ".json" or name.endswith(".bindings.json"):
        return json.loads(text)
    if (
        suffix in {".yaml", ".yml"}
        or name.endswith(".bindings.yaml")
        or name.endswith(".bindings.yml")
    ):
        loaded = yaml.safe_load(text)
        if loaded is None:
            raise SeriesBindingsLoadError(f"Empty YAML binding file: {path}")
        if not isinstance(loaded, dict):
            raise SeriesBindingsLoadError(f"Binding file root must be a mapping: {path}")
        return loaded
    raise SeriesBindingsLoadError(
        f"Unsupported binding file extension {path.suffix!r} (expected .bindings.yaml or .bindings.json): {path}"
    )


def parse_bindings_file(path: Path | str) -> dict[str, Any]:
    """Parse one binding sidecar file (YAML or JSON) without schema validation."""
    p = Path(path)
    text = p.read_text(encoding="utf-8")
    try:
        return _parse_raw_text(text, path=p)
    except json.JSONDecodeError as exc:
        raise SeriesBindingsLoadError(f"Invalid JSON in {p}: {exc}") from exc
    except yaml.YAMLError as exc:
        raise SeriesBindingsLoadError(f"Invalid YAML in {p}: {exc}") from exc


def _binding_files_in_directory(directory: Path) -> list[Path]:
    files_found: list[Path] = []
    for pattern in BINDINGS_GLOB_NAMES:
        files_found.extend(directory.glob(pattern))
    return sorted(set(files_found), key=lambda p: p.name.lower())


def merge_series_binding_documents(documents: list[dict[str, Any]]) -> dict[str, Any]:
    """Merge partial manifests (e.g. one per sheet) into one workbook document."""
    if not documents:
        raise SeriesBindingsLoadError("No binding documents to merge")

    schema_version: str | None = None
    workbook: str | None = None
    concept_scheme: dict[str, Any] | None = None
    series_by_id: dict[str, dict[str, Any]] = {}

    for index, doc in enumerate(documents):
        if not isinstance(doc, dict):
            raise SeriesBindingsLoadError(f"Document {index} must be a mapping")

        doc_schema = doc.get("schema_version")
        if not isinstance(doc_schema, str):
            raise SeriesBindingsLoadError(f"Document {index} missing string schema_version")
        if schema_version is None:
            schema_version = doc_schema
        elif doc_schema != schema_version:
            raise SeriesBindingsLoadError(
                f"schema_version mismatch: expected {schema_version!r}, got {doc_schema!r} in shard {index}"
            )

        doc_workbook = doc.get("workbook")
        if doc_workbook is not None:
            if not isinstance(doc_workbook, str):
                raise SeriesBindingsLoadError(f"Document {index} workbook must be a string")
            if workbook is None:
                workbook = doc_workbook
            elif doc_workbook != workbook:
                raise SeriesBindingsLoadError(
                    f"workbook mismatch: expected {workbook!r}, got {doc_workbook!r} in shard {index}"
                )

        doc_concepts = doc.get("concept_scheme")
        if doc_concepts is not None:
            if concept_scheme is None:
                concept_scheme = doc_concepts
            elif doc_concepts != concept_scheme:
                raise SeriesBindingsLoadError(
                    f"concept_scheme mismatch across shards (first difference at shard {index})"
                )

        series = doc.get("series")
        if not isinstance(series, list) or not series:
            raise SeriesBindingsLoadError(f"Document {index} must contain a non-empty series list")
        seen_ids_in_document: set[str] = set()

        for entry in series:
            if not isinstance(entry, dict):
                raise SeriesBindingsLoadError(f"series[] entries must be mappings in shard {index}")
            series_id = entry.get("id")
            if not isinstance(series_id, str):
                raise SeriesBindingsLoadError(
                    f"Each series entry requires string id (shard {index})"
                )
            if series_id in seen_ids_in_document:
                raise SeriesBindingsLoadError(
                    f"Duplicate series id {series_id!r} within shard {index}"
                )
            seen_ids_in_document.add(series_id)
            normalized = normalize_series_entry(entry)
            if series_id in series_by_id:
                try:
                    series_by_id[series_id] = merge_series_entries(
                        series_by_id[series_id],
                        normalized,
                        shard_index=index,
                    )
                except ValueError as exc:
                    raise SeriesBindingsLoadError(str(exc)) from exc
            else:
                series_by_id[series_id] = normalized

    merged_series = list(series_by_id.values())

    if schema_version is None:
        raise SeriesBindingsLoadError("Merged document missing schema_version")

    merged: dict[str, Any] = {
        "schema_version": schema_version,
        "series": merged_series,
    }
    if workbook is not None:
        merged["workbook"] = workbook
    if concept_scheme is not None:
        merged["concept_scheme"] = concept_scheme
    return merged


def load_series_bindings(path: Path | str, *, validate: bool = True) -> WorkbookSeriesBindings:
    """Load a binding sidecar file or directory of shards.

    When `path` is a directory, all `*.bindings.yaml` / `*.bindings.json` files
    are merged in sorted filename order before schema validation.
    """
    p = Path(path)
    if not p.exists():
        raise SeriesBindingsLoadError(f"Binding path does not exist: {p}")

    if p.is_dir():
        binding_files = _binding_files_in_directory(p)
        if not binding_files:
            raise SeriesBindingsLoadError(
                f"No binding files matching {BINDINGS_GLOB_NAMES!r} in directory: {p}"
            )
        documents = [parse_bindings_file(f) for f in binding_files]
        document = merge_series_binding_documents(documents)
    elif p.is_file():
        document = parse_bindings_file(p)
    else:
        raise SeriesBindingsLoadError(f"Binding path is not a file or directory: {p}")

    if validate:
        return validate_bindings_document(document)
    return cast(WorkbookSeriesBindings, document)
