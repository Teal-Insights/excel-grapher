"""Workbook-level series binding validation and optional codegen checks."""

from __future__ import annotations

import hashlib
import json
from collections.abc import Iterable, Sequence
from pathlib import Path
from typing import TYPE_CHECKING, Literal, NotRequired, TypedDict

from excel_grapher.core.address_keys import normalize_key as normalize_address
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings.canonical import bindings_canonical_sha256
from excel_grapher.series_bindings.input_series import derive_input_series
from excel_grapher.series_bindings.load import SeriesBindingsLoadError, load_series_bindings
from excel_grapher.series_bindings.normalize import has_input_direction
from excel_grapher.series_bindings.ranges import (
    expand_bound_series_addresses,
    expand_bound_series_addresses_for_graph,
)
from excel_grapher.series_bindings.types import (
    InputSeries,
    ValidationReport,
    WorkbookSeriesBindings,
)
from excel_grapher.series_bindings.validate import validate_series_bindings

if TYPE_CHECKING:
    from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
    from excel_grapher.grapher.graph import DependencyGraph


class BindingsCheckResult(TypedDict):
    """Result of validating a workbook against a binding sidecar."""

    bindings: WorkbookSeriesBindings
    graph: DependencyGraph
    targets: list[str]
    report: ValidationReport
    canonical_sha256: str
    inputs: list[str]
    computes: list[str]
    input_series: list[InputSeries]
    generated_files: NotRequired[dict[str, str]]


def input_ids(bindings: WorkbookSeriesBindings) -> list[str]:
    """Return sorted unique input series ids."""
    names: list[str] = []
    for series in bindings["series"]:
        series_id = series.get("id")
        if series_id and has_input_direction(series):
            names.append(str(series_id))
    return sorted(set(names))


def compute_names(bindings: WorkbookSeriesBindings) -> list[str]:
    """Return sorted unique declared output compute function names."""
    names: list[str] = []
    for series in bindings["series"]:
        output_block = series.get("output") or {}
        compute = output_block.get("compute")
        if isinstance(compute, dict) and compute.get("name"):
            names.append(str(compute["name"]))
    return sorted(set(names))


def all_series_targets(
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path,
) -> list[str]:
    """Expand every series `data_range` into graph target addresses.

    Applies `exclude_rows` / `exclude_columns` so punched holes are not treated
    as bound targets.
    """
    targets: list[str] = []
    for series in bindings["series"]:
        targets.extend(expand_bound_series_addresses(series, workbook=workbook))
    return sorted(set(targets))


TargetsFrom = Literal["all", "outputs"]
"""Which series seed graph extraction: every series, or only `output` series."""


def output_series_targets(
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path,
    extra_targets: Iterable[str] = (),
) -> list[str]:
    """Expand only `output` series `data_range`s into graph target addresses.

    Use this instead of a hand-maintained target list: output `data_range`s
    are authored from workbook structure, so the intended order is author
    outputs, derive targets, extract the graph, then author and validate
    inputs / internals / constants against that graph.

    Applies `exclude_rows` / `exclude_columns`, so output holes are the single
    place to skip cells inside a block. Output series over blank or literal
    cells stay in the set; they become graph leaves.

    Args:
        bindings: Loaded bindings document.
        workbook: Path to the `.xlsx` workbook (resolves named ranges).
        extra_targets: Additional roots to union with the derived set.

    Returns:
        Sorted, de-duplicated sheet-qualified addresses. Hash the result with
        `target_set_sha256` for a graph cache key that ignores non-output edits.
    """
    targets: list[str] = list(extra_targets)
    for series in bindings["series"]:
        if isinstance(series.get("output"), dict):
            targets.extend(expand_bound_series_addresses(series, workbook=workbook))
    return sorted(set(targets))


def series_targets(
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path,
    targets_from: TargetsFrom = "all",
) -> list[str]:
    """Return graph targets from every series or only from `output` series."""
    if targets_from == "outputs":
        return output_series_targets(bindings, workbook=workbook)
    return all_series_targets(bindings, workbook=workbook)


def target_set_sha256(targets: Iterable[str]) -> str:
    """Return a stable SHA-256 fingerprint of a graph target set.

    Order and duplicates do not affect the digest, and addresses are
    normalized first. Fold this (not the whole bindings document) into graph
    cache keys so edits to inputs, internals, labels or dimensions do not
    invalidate a cached graph.
    """
    canonical = sorted({normalize_address(target) for target in targets})
    payload = json.dumps(canonical, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def series_binding_public_addresses(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str,
) -> frozenset[str]:
    """Return normalized addresses published by series binding `data_range`s.

    Applies `exclude_rows` / `exclude_columns` so punched holes are not treated
    as bound. Pass the result as `preserve` to `OptimalCompression` /
    `compress_optimal` or `IdentityTransitCompression` /
    `compress_identity_transits` (or via `series_bindings=...` on either
    projection) so series-bound leaves that are not export targets stay in
    the projected graph. Both compressors always union `preserve` with
    `target_keys()`.
    """
    addresses: set[str] = set()
    for series in bindings.get("series", []):
        if not isinstance(series, dict):
            continue
        addresses.update(
            normalize_address(addr)
            for addr in expand_bound_series_addresses_for_graph(
                graph,
                series,
                workbook=workbook,
            )
        )
    return frozenset(addresses)


def _explicit_bindings_candidates(workbook: Path, bindings: Path) -> list[Path]:
    """Return candidate paths for an explicit ``--bindings`` argument."""
    candidates = [bindings]
    if not bindings.is_absolute():
        candidates.append(workbook.parent / bindings)
    return candidates


def resolve_bindings_path(
    workbook: Path,
    bindings: Path | None = None,
    *,
    create_if_missing: bool = False,
) -> Path:
    """Resolve a binding sidecar path from an explicit path or workbook conventions.

    Args:
        workbook: Path to the `.xlsx` workbook.
        bindings: Optional explicit sidecar file or shard directory.
        create_if_missing: When True, return the path that should be created
            instead of raising if no sidecar exists yet. Relative `--bindings`
            values resolve next to the workbook. This does not create files.

    Returns:
        Binding sidecar file or directory path. Existing paths are preferred;
        with `create_if_missing` the returned path may not exist yet.

    Raises:
        SeriesBindingsLoadError: When no sidecar can be resolved and
            `create_if_missing` is False.
    """
    if bindings is not None:
        candidates = _explicit_bindings_candidates(workbook, bindings)
        for candidate in candidates:
            if candidate.is_file() or candidate.is_dir():
                return candidate
        if create_if_missing:
            if not bindings.is_absolute():
                return workbook.parent / bindings
            return bindings
        tried = ", ".join(str(path) for path in candidates)
        raise SeriesBindingsLoadError(f"Binding path does not exist: {bindings} (tried: {tried})")

    candidates = [
        workbook.with_suffix(".bindings.yaml"),
        workbook.parent / f"{workbook.stem}.bindings",
    ]
    for candidate in candidates:
        if candidate.is_file() or candidate.is_dir():
            return candidate
    if create_if_missing:
        return workbook.parent / f"{workbook.stem}.bindings"

    tried = ", ".join(str(path) for path in candidates)
    raise SeriesBindingsLoadError(
        f"No binding sidecar found for {workbook} (tried: {tried}). "
        "Pass --bindings or colocate {workbook.stem}.bindings.yaml."
    )


def validate_bindings_workbook(
    workbook: Path,
    bindings_path: Path,
    *,
    dynamic_refs: DynamicRefConfig | None = None,
    use_cached_dynamic_refs: bool = True,
    blank_ranges: Sequence[str] | None = None,
    targets_from: TargetsFrom = "all",
) -> BindingsCheckResult:
    """Load bindings, build the graph, and validate against the workbook.

    Args:
        workbook: Path to the `.xlsx` workbook.
        bindings_path: Binding sidecar file or shard directory.
        dynamic_refs: Constraint-based OFFSET/INDEX/INDIRECT config. Ignored
            when `use_cached_dynamic_refs` is True.
        use_cached_dynamic_refs: Resolve dynamic refs from cached workbook
            values. Default True preserves the previous library behavior.
        blank_ranges: Sheet-qualified rectangles omitted from the graph.
        targets_from: `"all"` roots the graph at every series `data_range`;
            `"outputs"` roots it only at `output` series, matching a pipeline
            that extracts from `output_series_targets`.
    """
    bindings = load_series_bindings(bindings_path)
    targets = series_targets(bindings, workbook=workbook, targets_from=targets_from)
    graph = create_dependency_graph(
        workbook,
        targets,
        load_values=True,
        use_cached_dynamic_refs=use_cached_dynamic_refs,
        dynamic_refs=dynamic_refs,
        blank_ranges=blank_ranges,
    )
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    return {
        "bindings": bindings,
        "graph": graph,
        "targets": targets,
        "report": report,
        "canonical_sha256": bindings_canonical_sha256(bindings),
        "inputs": input_ids(bindings),
        "computes": compute_names(bindings),
        "input_series": derive_input_series(graph, bindings, workbook=workbook),
    }


def generate_bindings_modules(
    graph: DependencyGraph,
    *,
    bindings: WorkbookSeriesBindings,
    workbook: Path,
) -> dict[str, str]:
    """Generate a modular export package for the binding closure."""
    from excel_grapher.exporter import CodeGenerator

    with CodeGenerator(graph) as gen:
        return gen.generate_modules(
            series_bindings=bindings,
            bindings_workbook=workbook,
        )


def run_binding_checks(
    workbook: Path,
    bindings_path: Path,
    *,
    module_dir: Path,
    package_name: str = "bindings_module",
    smoke_test: bool = True,
    dynamic_refs: DynamicRefConfig | None = None,
    use_cached_dynamic_refs: bool = True,
) -> BindingsCheckResult:
    """Validate bindings, optionally smoke-test generated public functions."""
    result = validate_bindings_workbook(
        workbook,
        bindings_path,
        dynamic_refs=dynamic_refs,
        use_cached_dynamic_refs=use_cached_dynamic_refs,
    )
    report = result["report"]
    if not report["ok"]:
        errors = [issue for issue in report["issues"] if issue["level"] == "error"]
        raise ValueError(f"Binding validation failed: {errors}")

    files = generate_bindings_modules(
        result["graph"],
        bindings=result["bindings"],
        workbook=workbook,
    )
    if smoke_test:
        from excel_grapher.series_bindings.smoke import smoke_test_bindings_module

        smoke_test_bindings_module(
            files,
            bindings=result["bindings"],
            graph=result["graph"],
            workbook=workbook,
            module_dir=module_dir,
            package_name=package_name,
        )
    else:
        module_dir.mkdir(parents=True, exist_ok=True)
        for filename, content in files.items():
            (module_dir / filename).write_text(content, encoding="utf-8")

    result["generated_files"] = files
    return result
