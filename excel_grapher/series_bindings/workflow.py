"""Workbook-level series binding validation and optional codegen checks."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, NotRequired, TypedDict

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings.canonical import bindings_canonical_sha256
from excel_grapher.series_bindings.input_series import derive_input_series
from excel_grapher.series_bindings.load import SeriesBindingsLoadError, load_series_bindings
from excel_grapher.series_bindings.ranges import expand_data_range
from excel_grapher.series_bindings.types import (
    InputSeries,
    ValidationReport,
    WorkbookSeriesBindings,
)
from excel_grapher.series_bindings.validate import validate_series_bindings

if TYPE_CHECKING:
    from excel_grapher.grapher.graph import DependencyGraph


class BindingsCheckResult(TypedDict):
    """Result of validating a workbook against a binding sidecar."""

    bindings: WorkbookSeriesBindings
    graph: DependencyGraph
    targets: list[str]
    report: ValidationReport
    canonical_sha256: str
    setters: list[str]
    computes: list[str]
    input_series: list[InputSeries]
    generated_files: NotRequired[dict[str, str]]


def setter_names(bindings: WorkbookSeriesBindings) -> list[str]:
    """Return sorted unique declared input setter function names."""
    names: list[str] = []
    for series in bindings["series"]:
        input_block = series.get("input") or {}
        setter = input_block.get("setter") or series.get("setter")
        if isinstance(setter, dict) and setter.get("name"):
            names.append(str(setter["name"]))
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
    """Expand every series ``data_range`` into graph target addresses."""
    targets: list[str] = []
    for series in bindings["series"]:
        data_range = series.get("data_range")
        if isinstance(data_range, str):
            targets.extend(expand_data_range(data_range, workbook=workbook))
    return sorted(set(targets))


def _explicit_bindings_candidates(workbook: Path, bindings: Path) -> list[Path]:
    """Return candidate paths for an explicit ``--bindings`` argument."""
    candidates = [bindings]
    if not bindings.is_absolute():
        candidates.append(workbook.parent / bindings)
    return candidates


def resolve_bindings_path(workbook: Path, bindings: Path | None = None) -> Path:
    """Resolve a binding sidecar path from an explicit path or workbook conventions.

    Args:
        workbook: Path to the ``.xlsx`` workbook.
        bindings: Optional explicit sidecar file or shard directory.

    Returns:
        Existing binding sidecar file or directory path.

    Raises:
        SeriesBindingsLoadError: When no sidecar can be resolved.
    """
    if bindings is not None:
        for candidate in _explicit_bindings_candidates(workbook, bindings):
            if candidate.is_file() or candidate.is_dir():
                return candidate
        tried = ", ".join(str(path) for path in _explicit_bindings_candidates(workbook, bindings))
        raise SeriesBindingsLoadError(f"Binding path does not exist: {bindings} (tried: {tried})")

    candidates: list[Path] = [
        workbook.with_suffix(".bindings.yaml"),
        workbook.parent / f"{workbook.stem}.bindings",
    ]
    for candidate in candidates:
        if candidate.is_file() or candidate.is_dir():
            return candidate

    tried = ", ".join(str(path) for path in candidates)
    raise SeriesBindingsLoadError(
        f"No binding sidecar found for {workbook} (tried: {tried}). "
        "Pass --bindings or colocate {workbook.stem}.bindings.yaml."
    )


def validate_bindings_workbook(
    workbook: Path,
    bindings_path: Path,
) -> BindingsCheckResult:
    """Load bindings, build the graph, and validate against the workbook."""
    bindings = load_series_bindings(bindings_path)
    targets = all_series_targets(bindings, workbook=workbook)
    graph = create_dependency_graph(
        workbook,
        targets,
        load_values=True,
        use_cached_dynamic_refs=True,
    )
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    return {
        "bindings": bindings,
        "graph": graph,
        "targets": targets,
        "report": report,
        "canonical_sha256": bindings_canonical_sha256(bindings),
        "setters": setter_names(bindings),
        "computes": compute_names(bindings),
        "input_series": derive_input_series(graph, bindings, workbook=workbook),
    }


def generate_bindings_modules(
    graph: DependencyGraph,
    *,
    targets: list[str],
    bindings: WorkbookSeriesBindings,
    workbook: Path,
) -> dict[str, str]:
    """Generate a modular export package for the binding closure."""
    from excel_grapher.exporter import CodeGenerator

    with CodeGenerator(graph) as gen:
        return gen.generate_modules(
            targets,
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
) -> BindingsCheckResult:
    """Validate bindings, optionally smoke-test generated setters and computes."""
    result = validate_bindings_workbook(workbook, bindings_path)
    report = result["report"]
    if not report["ok"]:
        errors = [issue for issue in report["issues"] if issue["level"] == "error"]
        raise ValueError(f"Binding validation failed: {errors}")

    files = generate_bindings_modules(
        result["graph"],
        targets=result["targets"],
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
