"""Workbook-level series binding validation and optional codegen checks."""

from __future__ import annotations

from collections.abc import Iterable, Sequence
from pathlib import Path
from typing import TYPE_CHECKING, Literal, NotRequired, TypedDict

from excel_grapher.core.address_keys import normalize_key as normalize_address
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings.canonical import bindings_canonical_sha256
from excel_grapher.series_bindings.input_series import derive_input_series
from excel_grapher.series_bindings.load import SeriesBindingsLoadError, load_series_bindings
from excel_grapher.series_bindings.ranges import (
    expand_data_range,
    expand_data_range_for_graph,
    series_data_ranges,
)
from excel_grapher.series_bindings.resolve import resolve_series_bindings
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
    readers: list[str]
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


def reader_names(bindings: WorkbookSeriesBindings) -> list[str]:
    """Return sorted unique reader function names for input and constant series.

    Every input series that declares a setter gets a `read_<series_id>` dual
    (or an explicit `input.reader.name` override). Constant series emit the same
    reader surface without a setter (`constant.reader.name` or `read_<series_id>`).
    Range duals (`read_<id>_range`) are omitted from this list; they are
    auxiliary helpers.
    """
    names: list[str] = []
    for series in bindings["series"]:
        series_id = series.get("id")
        if not series_id:
            continue
        input_block = series.get("input") or {}
        setter = input_block.get("setter") or series.get("setter")
        constant_block = series.get("constant")
        if isinstance(setter, dict) and setter.get("name"):
            reader = input_block.get("reader")
            if isinstance(reader, dict) and reader.get("name"):
                names.append(str(reader["name"]))
            else:
                names.append(f"read_{series_id}")
        elif isinstance(constant_block, dict):
            reader = constant_block.get("reader")
            if isinstance(reader, dict) and reader.get("name"):
                names.append(str(reader["name"]))
            else:
                names.append(f"read_{series_id}")
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
        for data_range in series_data_ranges(series):
            targets.extend(expand_data_range(data_range, workbook=workbook))
    return sorted(set(targets))


def series_binding_public_addresses(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str,
) -> frozenset[str]:
    """Return normalized addresses published by series binding `data_range`s.

    Pass the result as `preserve` to `OptimalCompression` / `compress_optimal`
    or `IdentityTransitCompression` / `compress_identity_transits` (or via
    `series_bindings=...` on either projection) so series-bound leaves that
    are not export targets stay in the projected graph. Both compressors
    always union `preserve` with `target_keys()`.
    """
    addresses: set[str] = set()
    for series in bindings.get("series", []):
        if not isinstance(series, dict):
            continue
        for data_range in series_data_ranges(series):
            addresses.update(
                normalize_address(addr)
                for addr in expand_data_range_for_graph(
                    graph,
                    data_range,
                    workbook=workbook,
                )
            )
    return frozenset(addresses)


def output_binding_covered_addresses(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str,
    export_addresses: Iterable[str] | None = None,
) -> frozenset[str]:
    """Return addresses covered by successfully resolved output computes.

    Uses the same resolution path as compute codegen (including `exclude_rows` /
    `exclude_columns` and optional export-closure intersection), so coverage
    matches the addresses semantic `compute_*` functions actually return.
    """
    report = resolve_series_bindings(
        graph,
        bindings,
        workbook=workbook,
        direction="output",
        export_addresses=export_addresses,
    )
    addresses: set[str] = set()
    for resolved in report["series"]:
        if not resolved["ok"]:
            continue
        for leaf in resolved["leaves"]:
            addresses.add(normalize_address(leaf["address"]))
    return frozenset(addresses)


def should_emit_compute_all(
    targets: Sequence[str],
    *,
    covered_by_output: frozenset[str],
    include_compute_all: bool | None = None,
) -> bool:
    """Decide whether generated exports should include public `compute_all`.

    Args:
        targets: Normalized export target cell addresses.
        covered_by_output: Addresses covered by resolved output series computes.
        include_compute_all: Explicit override. `True` always emits, `False`
            never emits, and `None` (default) omits only when every target is
            covered by an output binding.

    Returns:
        True when `compute_all` should be part of the generated public API.
    """
    if include_compute_all is True:
        return True
    if include_compute_all is False:
        return False
    if not targets:
        return True
    return not all(normalize_address(target) in covered_by_output for target in targets)


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
        "readers": reader_names(bindings),
        "computes": compute_names(bindings),
        "input_series": derive_input_series(graph, bindings, workbook=workbook),
    }


def generate_bindings_modules(
    graph: DependencyGraph,
    *,
    targets: list[str],
    bindings: WorkbookSeriesBindings,
    workbook: Path,
    paradigm: Literal["ctx", "inverted_tree"] = "ctx",
) -> dict[str, str]:
    """Generate a modular export package for the binding closure."""
    from excel_grapher.exporter import CodeGenerator

    with CodeGenerator(graph) as gen:
        return gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
            paradigm=paradigm,
        )


def run_binding_checks(
    workbook: Path,
    bindings_path: Path,
    *,
    module_dir: Path,
    package_name: str = "bindings_module",
    smoke_test: bool = True,
    paradigm: Literal["ctx", "inverted_tree"] = "ctx",
) -> BindingsCheckResult:
    """Validate bindings, optionally smoke-test generated public functions."""
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
        paradigm=paradigm,
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
            paradigm=paradigm,
        )
    else:
        module_dir.mkdir(parents=True, exist_ok=True)
        for filename, content in files.items():
            (module_dir / filename).write_text(content, encoding="utf-8")

    result["generated_files"] = files
    return result
