"""``excel-grapher bindings`` command group."""

from __future__ import annotations

import argparse
import json
import sys
import tempfile
from collections.abc import Mapping
from pathlib import Path
from typing import Any, cast

import yaml

from excel_grapher.exporter.inverted_tree import InvertedTreeExportError
from excel_grapher.exporter.semantic_catalog import SemanticCatalogError
from excel_grapher.exporter.semantic_viz import to_semantic_viz_payload, write_semantic_viz_html
from excel_grapher.grapher.blank_ranges import BlankRangesLoadError, load_blank_ranges_module
from excel_grapher.grapher.builder import list_dynamic_ref_constraint_candidates
from excel_grapher.grapher.dynamic_refs import (
    EXACT_MATCH_UNDOMAINED_WARN_CAP,
    DynamicRefConfig,
    DynamicRefError,
)
from excel_grapher.series_bindings.audit import (
    DIRECTIONS,
    audit_binding_resolutions,
    format_audit_findings,
)
from excel_grapher.series_bindings.burndown import (
    format_burndown_report,
    internal_binding_burndown,
    load_exempt_addresses,
)
from excel_grapher.series_bindings.domains import SeriesRelationError, undomained_leaves
from excel_grapher.series_bindings.load import SeriesBindingsLoadError, load_series_bindings
from excel_grapher.series_bindings.resolve import BindingDirection
from excel_grapher.series_bindings.schema import SeriesBindingsSchemaError
from excel_grapher.series_bindings.smoke import BindingsSmokeError
from excel_grapher.series_bindings.types import (
    ValidationIssue,
    ValidationReport,
    WorkbookSeriesBindings,
)
from excel_grapher.series_bindings.upsert import (
    BindingUpsertError,
    upsert_series_binding,
)
from excel_grapher.series_bindings.versions import CURRENT_SCHEMA_VERSION
from excel_grapher.series_bindings.workflow import (
    all_series_targets,
    generate_bindings_modules,
    resolve_bindings_path,
    run_binding_checks,
    validate_bindings_workbook,
)

_PY_DYNAMIC_REF_HINT = (
    "Pass dynamic_refs=DynamicRefConfig.from_bindings(...) / "
    "DynamicRefConfig.from_constraints(...) or set use_cached_dynamic_refs=True."
)
_CLI_DYNAMIC_REF_HINT = (
    "Declare series domain in the bindings sidecar, or set --use-cached-dynamic-refs."
)


def register(subparsers: argparse._SubParsersAction[argparse.ArgumentParser]) -> None:
    """Register the ``bindings`` command group."""
    bindings_parser = subparsers.add_parser("bindings", help="Series binding sidecar tools")
    bindings_sub = bindings_parser.add_subparsers(dest="bindings_command", required=True)
    validate_parser = bindings_sub.add_parser(
        "validate",
        help="Validate a workbook binding sidecar",
    )
    validate_parser.add_argument("workbook", type=Path, help="Path to the .xlsx workbook")
    validate_parser.add_argument(
        "--bindings",
        type=Path,
        default=None,
        help="Binding sidecar file or shard directory (default: colocated sidecar)",
    )
    validate_parser.add_argument(
        "--json",
        action="store_true",
        help="Print the validation report as JSON",
    )
    validate_parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Print validation warnings (errors are always printed on failure)",
    )
    validate_parser.add_argument(
        "--smoke-test",
        action="store_true",
        help="Generate modules and smoke-test after validation (compute_* with data.py defaults)",
    )
    validate_parser.add_argument(
        "--emit-dir",
        type=Path,
        default=None,
        help="Write generated module files to this directory",
    )
    validate_parser.add_argument(
        "--package-name",
        default="bindings_module",
        help="Package directory name for smoke tests (default: bindings_module)",
    )
    validate_parser.add_argument(
        "--use-cached-dynamic-refs",
        action="store_true",
        help="Resolve OFFSET/INDEX/INDIRECT from the workbook's cached values instead of "
        "series domain.",
    )
    validate_parser.add_argument(
        "--blank-ranges",
        type=Path,
        default=None,
        help="Python module exposing BLANK_RANGES: Sequence[str] "
        "(sheet-qualified rectangles omitted from the graph)",
    )
    viz_parser = bindings_sub.add_parser(
        "viz",
        help="Write a statement-graph HTML visualization from series bindings",
    )
    viz_parser.add_argument("workbook", type=Path, help="Path to the .xlsx workbook")
    viz_parser.add_argument(
        "--bindings",
        type=Path,
        default=None,
        help="Binding sidecar file or shard directory (default: colocated sidecar)",
    )
    viz_parser.add_argument(
        "--output",
        type=Path,
        required=True,
        help="Path to write the HTML viewer",
    )
    viz_parser.add_argument(
        "--use-cached-dynamic-refs",
        action="store_true",
        help="Resolve OFFSET/INDEX/INDIRECT from the workbook's cached values instead of "
        "series domain.",
    )
    viz_parser.add_argument(
        "--json",
        action="store_true",
        help="Also write the statement-graph payload next to --output as .viz.json",
    )
    viz_parser.add_argument(
        "--blank-ranges",
        type=Path,
        default=None,
        help="Python module exposing BLANK_RANGES: Sequence[str] "
        "(sheet-qualified rectangles omitted from the graph)",
    )
    undomained_parser = bindings_sub.add_parser(
        "undomained",
        help="List graph leaves that have no bindings domain",
    )
    undomained_parser.add_argument("workbook", type=Path, help="Path to the .xlsx workbook")
    undomained_parser.add_argument(
        "--bindings",
        type=Path,
        default=None,
        help="Binding sidecar file or shard directory (default: colocated sidecar)",
    )
    undomained_parser.add_argument(
        "--use-cached-dynamic-refs",
        action="store_true",
        help="Resolve OFFSET/INDEX/INDIRECT from the workbook's cached values instead of "
        "bindings domains.",
    )
    undomained_parser.add_argument(
        "--json",
        action="store_true",
        help="Print the undomained leaf list as JSON",
    )
    candidates_parser = bindings_sub.add_parser(
        "candidates",
        help="List dynamic-ref leaves that still need a bindings domain",
    )
    candidates_parser.add_argument("workbook", type=Path, help="Path to the .xlsx workbook")
    candidates_parser.add_argument(
        "--bindings",
        type=Path,
        default=None,
        help="Binding sidecar file or shard directory (default: colocated sidecar)",
    )
    candidates_parser.add_argument(
        "--target",
        action="append",
        default=None,
        help="Extraction root (sheet-qualified address, range, or defined name). "
        "Repeatable. Unioned with series data_range cells.",
    )
    candidates_parser.add_argument(
        "--json",
        action="store_true",
        help="Print the candidate address list as JSON",
    )
    candidates_parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit 1 when any dynamic-ref leaf is still missing a domain",
    )
    candidates_parser.add_argument(
        "--undomained-match",
        action="store_true",
        help="Also report lookup cells exact MATCH kept only because they have no domain. "
        "Printed separately from the must-bind list.",
    )

    audit_parser = bindings_sub.add_parser(
        "audit",
        help="Fail if series resolution would fail codegen",
    )
    _add_workbook_wiring_args(audit_parser)
    audit_parser.add_argument(
        "--json",
        action="store_true",
        help="Print the audit report as JSON",
    )
    audit_parser.add_argument(
        "--direction",
        action="append",
        choices=list(DIRECTIONS),
        help="Limit to one direction (repeatable). Default: all.",
    )
    audit_parser.add_argument(
        "--strict",
        action="store_true",
        help="Treat warnings as fatal (exit 1 when any warning is present)",
    )

    burndown_parser = bindings_sub.add_parser(
        "burndown",
        help="Print unbound formula cells as a coverage worklist",
    )
    _add_workbook_wiring_args(burndown_parser)
    burndown_parser.add_argument(
        "--json",
        action="store_true",
        help="Print the coverage worklist as JSON",
    )
    burndown_parser.add_argument(
        "--per-sheet",
        default=None,
        help="Only print row detail for this sheet (summary always prints)",
    )
    burndown_parser.add_argument(
        "--max-rows",
        type=int,
        default=None,
        help="Limit row detail lines per sheet",
    )
    burndown_parser.add_argument(
        "--exempt",
        type=Path,
        default=None,
        help="Text file of reviewed sheet-qualified addresses to omit from the worklist",
    )
    burndown_parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit 1 when any unbound internal formula cells remain",
    )

    upsert_parser = bindings_sub.add_parser(
        "upsert",
        help="Insert or replace one series after fail-closed checks",
    )
    _add_workbook_wiring_args(upsert_parser)
    upsert_parser.add_argument(
        "--series",
        type=Path,
        required=True,
        help="YAML file with one series mapping (or a document with series: [one entry])",
    )
    upsert_parser.add_argument(
        "--replace",
        action="store_true",
        help="Replace an existing series with the same id",
    )


def _add_workbook_wiring_args(parser: argparse.ArgumentParser) -> None:
    """Add workbook / sidecar arguments shared by bindings commands."""
    parser.add_argument("workbook", type=Path, help="Path to the .xlsx workbook")
    parser.add_argument(
        "--bindings",
        type=Path,
        default=None,
        help="Binding sidecar file or shard directory (default: colocated sidecar)",
    )
    parser.add_argument(
        "--use-cached-dynamic-refs",
        action="store_true",
        help="Resolve OFFSET/INDEX/INDIRECT from the workbook's cached values instead of "
        "series domain.",
    )
    parser.add_argument(
        "--blank-ranges",
        type=Path,
        default=None,
        help="Python module exposing BLANK_RANGES: Sequence[str] "
        "(sheet-qualified rectangles omitted from the graph)",
    )


def dispatch(args: argparse.Namespace) -> int:
    """Dispatch a ``bindings`` subcommand."""
    if args.bindings_command == "validate":
        return cmd_validate(args)
    if args.bindings_command == "viz":
        return cmd_viz(args)
    if args.bindings_command == "undomained":
        return cmd_undomained(args)
    if args.bindings_command == "candidates":
        return cmd_candidates(args)
    if args.bindings_command == "audit":
        return cmd_audit(args)
    if args.bindings_command == "burndown":
        return cmd_burndown(args)
    if args.bindings_command == "upsert":
        return cmd_upsert(args)
    print(f"Unknown bindings command: {args.bindings_command}", file=sys.stderr)
    return 2


def cmd_viz(args: argparse.Namespace) -> int:
    """Run ``excel-grapher bindings viz``."""
    workbook = args.workbook
    if not workbook.is_file():
        print(f"Workbook not found: {workbook}", file=sys.stderr)
        return 1
    try:
        bindings_path = resolve_bindings_path(workbook, args.bindings)
    except SeriesBindingsLoadError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    try:
        bindings_doc = load_series_bindings(bindings_path)
        dynamic_refs = _derive_dynamic_refs(workbook, bindings_doc, bindings_path)
        blank_ranges = (
            load_blank_ranges_module(args.blank_ranges) if args.blank_ranges is not None else None
        )
        result = validate_bindings_workbook(
            workbook,
            bindings_path,
            dynamic_refs=dynamic_refs,
            use_cached_dynamic_refs=args.use_cached_dynamic_refs,
            blank_ranges=blank_ranges,
        )
        payload = to_semantic_viz_payload(
            result["graph"],
            result["bindings"],
            workbook=workbook,
            blank_ranges=blank_ranges,
        )
        write_semantic_viz_html(payload, args.output, title=workbook.name)
        if args.json:
            json_path = args.output.with_suffix(".viz.json")
            json_path.write_text(
                json.dumps(payload.to_dict(), indent=2, default=str),
                encoding="utf-8",
            )
            print(f"wrote {json_path}")
        print(
            f"wrote {args.output} "
            f"statements={payload.graph.stats.statement_count} "
            f"bundles={payload.graph.stats.bundle_count} "
            f"instance_edges={payload.graph.stats.instance_edge_count} "
            f"cells={payload.graph.stats.cell_count}"
        )
        return 0
    except SeriesBindingsLoadError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except SeriesBindingsSchemaError as exc:
        print(f"Binding sidecar schema error:\n  {exc}", file=sys.stderr)
        return 1
    except BlankRangesLoadError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except (SemanticCatalogError, InvertedTreeExportError) as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except (DynamicRefError, ValueError) as exc:
        print(_format_cli_dynamic_ref_error(exc), file=sys.stderr)
        return 1


def cmd_validate(args: argparse.Namespace) -> int:
    """Run ``excel-grapher bindings validate``."""
    workbook = args.workbook
    if not workbook.is_file():
        print(f"Workbook not found: {workbook}", file=sys.stderr)
        return 1

    try:
        bindings_path = resolve_bindings_path(workbook, args.bindings)
    except SeriesBindingsLoadError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    try:
        bindings_doc = load_series_bindings(bindings_path)
        dynamic_refs = _derive_dynamic_refs(workbook, bindings_doc, bindings_path)
        result = validate_bindings_workbook(
            workbook,
            bindings_path,
            dynamic_refs=dynamic_refs,
            use_cached_dynamic_refs=args.use_cached_dynamic_refs,
            blank_ranges=(
                load_blank_ranges_module(args.blank_ranges)
                if getattr(args, "blank_ranges", None) is not None
                else None
            ),
        )
    except SeriesBindingsLoadError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except SeriesBindingsSchemaError as exc:
        print(f"Binding sidecar schema error:\n  {exc}", file=sys.stderr)
        return 1
    except BlankRangesLoadError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except (DynamicRefError, ValueError) as exc:
        print(_format_cli_dynamic_ref_error(exc), file=sys.stderr)
        return 1

    report = result["report"]
    if args.json:
        print(json.dumps(report, indent=2, default=str))
    else:
        _print_summary(result)
        _print_issues(report, include_warnings=args.verbose)

    if not report["ok"]:
        return 1

    if not args.smoke_test and args.emit_dir is None:
        return 0

    try:
        if args.smoke_test:
            if args.emit_dir is None:
                with tempfile.TemporaryDirectory() as temp_dir:
                    module_dir = Path(temp_dir) / args.package_name
                    run_binding_checks(
                        workbook,
                        bindings_path,
                        module_dir=module_dir,
                        package_name=args.package_name,
                        smoke_test=True,
                        dynamic_refs=dynamic_refs,
                        use_cached_dynamic_refs=args.use_cached_dynamic_refs,
                    )
            else:
                module_dir = _module_dir(args.emit_dir, args.package_name)
                check_result = run_binding_checks(
                    workbook,
                    bindings_path,
                    module_dir=module_dir,
                    package_name=args.package_name,
                    smoke_test=True,
                    dynamic_refs=dynamic_refs,
                    use_cached_dynamic_refs=args.use_cached_dynamic_refs,
                )
                if not args.json:
                    _write_generated_files(check_result, module_dir)
        elif args.emit_dir is not None:
            module_dir = _module_dir(args.emit_dir, args.package_name)
            files = generate_bindings_modules(
                result["graph"],
                bindings=result["bindings"],
                workbook=workbook,
            )
            _write_generated_files({"generated_files": files}, module_dir)
    except (BindingsSmokeError, InvertedTreeExportError) as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except (DynamicRefError, ValueError) as exc:
        print(_format_cli_dynamic_ref_error(exc), file=sys.stderr)
        return 1

    if not args.json and args.smoke_test:
        print("All compute functions passed smoke checks.")
    return 0


def cmd_undomained(args: argparse.Namespace) -> int:
    """Run ``excel-grapher bindings undomained``."""
    workbook = args.workbook
    if not workbook.is_file():
        print(f"Workbook not found: {workbook}", file=sys.stderr)
        return 1
    try:
        bindings_path = resolve_bindings_path(workbook, args.bindings)
    except SeriesBindingsLoadError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    try:
        bindings_doc = load_series_bindings(bindings_path)
        dynamic_refs = _derive_dynamic_refs(workbook, bindings_doc, bindings_path)
        result = validate_bindings_workbook(
            workbook,
            bindings_path,
            dynamic_refs=dynamic_refs,
            use_cached_dynamic_refs=args.use_cached_dynamic_refs,
        )
        leaves = undomained_leaves(result["graph"], result["bindings"], workbook=workbook)
    except SeriesBindingsLoadError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except SeriesBindingsSchemaError as exc:
        print(f"Binding sidecar schema error:\n  {exc}", file=sys.stderr)
        return 1
    except (DynamicRefError, ValueError) as exc:
        print(_format_cli_dynamic_ref_error(exc), file=sys.stderr)
        return 1
    if args.json:
        print(json.dumps(leaves, indent=2))
    elif leaves:
        print("\n".join(leaves))
    else:
        print("ok: every graph leaf has a bindings domain")
    return 0


def cmd_candidates(args: argparse.Namespace) -> int:
    """Run ``excel-grapher bindings candidates``.

    Lists leaf addresses that feed OFFSET / INDEX / INDIRECT and have no
    bindings domain yet. Does not build the dependency graph and does not
    resolve dynamic refs from cached values.

    Cells reachable only through an unresolved dynamic ref are omitted until
    the controlling domain is declared. Re-run after each domain edit on a
    dynamic-ref argument.
    """
    workbook = args.workbook
    if not workbook.is_file():
        print(f"Workbook not found: {workbook}", file=sys.stderr)
        return 1

    bindings_doc: WorkbookSeriesBindings = {
        "schema_version": CURRENT_SCHEMA_VERSION,
        "series": [],
    }
    bindings_path: Path | None = None
    try:
        bindings_path = resolve_bindings_path(workbook, args.bindings)
    except SeriesBindingsLoadError as exc:
        if not args.target:
            print(
                f"{exc}\n"
                "Author an output (or other) series, or pass --target with an extraction root.",
                file=sys.stderr,
            )
            return 1
    else:
        try:
            bindings_doc = load_series_bindings(bindings_path)
        except SeriesBindingsLoadError as exc:
            print(str(exc), file=sys.stderr)
            return 1
        except SeriesBindingsSchemaError as exc:
            print(f"Binding sidecar schema error:\n  {exc}", file=sys.stderr)
            return 1

    targets = list(args.target or [])
    if bindings_path is not None:
        targets.extend(all_series_targets(bindings_doc, workbook=workbook))
    targets = sorted(set(targets))
    if not targets:
        print(
            "No extraction targets. Author a series data_range or pass --target.",
            file=sys.stderr,
        )
        return 1

    undomained_exact_match: list[str] = []
    try:
        dynamic_refs = DynamicRefConfig.from_bindings(
            bindings_doc, workbook, bindings_path=bindings_path
        )
        leaves = list_dynamic_ref_constraint_candidates(
            workbook,
            targets,
            dynamic_refs=dynamic_refs,
            undomained_exact_match=undomained_exact_match if args.undomained_match else None,
        )
    except SeriesRelationError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except (DynamicRefError, ValueError) as exc:
        print(_format_cli_dynamic_ref_error(exc), file=sys.stderr)
        return 1

    if args.json and args.undomained_match:
        print(
            json.dumps(
                {"missing": leaves, "undomained_exact_match": undomained_exact_match},
                indent=2,
            )
        )
    elif args.json:
        print(json.dumps(leaves, indent=2))
    elif leaves:
        print("\n".join(leaves))
    else:
        print("ok: no dynamic-ref leaves are missing a bindings domain")
    if args.undomained_match and not args.json:
        _print_undomained_exact_match(undomained_exact_match)
    if args.strict and leaves:
        return 1
    return 0


def _print_undomained_exact_match(addresses: list[str]) -> None:
    """Print the opt-in exact-MATCH list without merging it into must-bind output."""
    if not addresses:
        print("undomained exact MATCH: none")
        return
    total = len(addresses)
    if total > EXACT_MATCH_UNDOMAINED_WARN_CAP:
        shown = addresses[:EXACT_MATCH_UNDOMAINED_WARN_CAP]
        print(
            f"undomained exact MATCH ({total} cells; showing {len(shown)}; "
            "pass --json for the full list):"
        )
        print("\n".join(shown))
        return
    print(f"undomained exact MATCH ({total}):")
    print("\n".join(addresses))


def _parse_audit_directions(raw: list[str] | None) -> tuple[BindingDirection, ...]:
    if not raw:
        return DIRECTIONS
    parsed: list[BindingDirection] = []
    allowed = set(DIRECTIONS)
    for item in raw:
        if item not in allowed:
            raise SystemExit(f"Unknown direction {item!r}; expected one of {sorted(allowed)}")
        direction = cast(BindingDirection, item)
        if direction not in parsed:
            parsed.append(direction)
    return tuple(parsed)


def _load_series_document(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise BindingUpsertError(f"Series file not found: {path}")
    loaded = yaml.safe_load(path.read_text(encoding="utf-8"))
    if not isinstance(loaded, dict):
        raise BindingUpsertError(f"Series file root must be a mapping: {path}")
    if "series" in loaded:
        series = loaded["series"]
        if not isinstance(series, list) or len(series) != 1 or not isinstance(series[0], dict):
            raise BindingUpsertError(
                f"Series file {path} must contain exactly one series entry when using "
                "a bindings document shape"
            )
        return series[0]
    return loaded


def _bindings_context(args: argparse.Namespace):
    workbook = args.workbook
    if not workbook.is_file():
        print(f"Workbook not found: {workbook}", file=sys.stderr)
        return None
    try:
        bindings_path = resolve_bindings_path(workbook, args.bindings)
        bindings_doc = load_series_bindings(bindings_path)
        dynamic_refs = _derive_dynamic_refs(workbook, bindings_doc, bindings_path)
        blank_ranges = (
            load_blank_ranges_module(args.blank_ranges) if args.blank_ranges is not None else None
        )
        result = validate_bindings_workbook(
            workbook,
            bindings_path,
            dynamic_refs=dynamic_refs,
            use_cached_dynamic_refs=args.use_cached_dynamic_refs,
            blank_ranges=blank_ranges,
        )
    except SeriesBindingsLoadError as exc:
        print(str(exc), file=sys.stderr)
        return None
    except SeriesBindingsSchemaError as exc:
        print(f"Binding sidecar schema error:\n  {exc}", file=sys.stderr)
        return None
    except BlankRangesLoadError as exc:
        print(str(exc), file=sys.stderr)
        return None
    except (DynamicRefError, ValueError) as exc:
        print(_format_cli_dynamic_ref_error(exc), file=sys.stderr)
        return None
    return workbook, bindings_path, result


def cmd_audit(args: argparse.Namespace) -> int:
    """Run ``excel-grapher bindings audit``."""
    loaded = _bindings_context(args)
    if loaded is None:
        return 1
    workbook, _bindings_path, result = loaded
    directions = _parse_audit_directions(args.direction)
    report = audit_binding_resolutions(
        result["graph"],
        result["bindings"],
        workbook=workbook,
        directions=directions,
    )
    if args.json:
        print(json.dumps(report.to_dict(), indent=2, default=str))
    else:
        print(
            f"Binding resolution audit: {report.error_count} error(s), "
            f"{report.warning_count} warning(s) "
            f"(directions={','.join(directions)})"
        )
        lines = format_audit_findings(report.findings)
        if lines:
            print()
            print("\n".join(lines))
        else:
            print("No resolution issues found.")
    if not report.ok:
        return 1
    if args.strict and report.warning_count > 0:
        return 1
    return 0


def cmd_burndown(args: argparse.Namespace) -> int:
    """Run ``excel-grapher bindings burndown``."""
    loaded = _bindings_context(args)
    if loaded is None:
        return 1
    workbook, _bindings_path, result = loaded
    exempt = load_exempt_addresses(args.exempt) if args.exempt is not None else frozenset()
    report = internal_binding_burndown(
        result["graph"],
        result["bindings"],
        exempt_cells=exempt,
        workbook=workbook,
        per_sheet=args.per_sheet,
        max_rows=args.max_rows,
    )
    if args.json:
        print(json.dumps(report.to_dict(), indent=2, default=str))
    else:
        print("\n".join(format_burndown_report(report)))
    if args.strict and report.unbound_count > 0:
        return 1
    return 0


def cmd_upsert(args: argparse.Namespace) -> int:
    """Run ``excel-grapher bindings upsert``."""
    workbook = args.workbook
    if not workbook.is_file():
        print(f"Workbook not found: {workbook}", file=sys.stderr)
        return 1
    try:
        series = _load_series_document(args.series)
        bindings_path = resolve_bindings_path(workbook, args.bindings, create_if_missing=True)
        if bindings_path.is_file() or bindings_path.is_dir():
            bindings_doc = load_series_bindings(bindings_path)
        else:
            bindings_doc = {}
        dynamic_refs = _derive_dynamic_refs(workbook, bindings_doc, bindings_path)
        blank_ranges = (
            load_blank_ranges_module(args.blank_ranges) if args.blank_ranges is not None else None
        )
        result = upsert_series_binding(
            workbook,
            bindings_path,
            series,
            replace=args.replace,
            dynamic_refs=dynamic_refs,
            use_cached_dynamic_refs=args.use_cached_dynamic_refs,
            blank_ranges=blank_ranges,
        )
    except BindingUpsertError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except SeriesBindingsLoadError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except BlankRangesLoadError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except (DynamicRefError, ValueError) as exc:
        print(_format_cli_dynamic_ref_error(exc), file=sys.stderr)
        return 1
    print(f"{result.action} {result.series_id} -> {result.shard_path}")
    return 0


def _derive_dynamic_refs(
    workbook: Path,
    bindings: Mapping[str, Any],
    bindings_path: Path,
) -> DynamicRefConfig | None:
    """Derive a dynamic-ref config from bindings series `domain`."""
    bindings_config = DynamicRefConfig.from_bindings(
        bindings, workbook, bindings_path=bindings_path
    )
    return bindings_config if len(bindings_config.cell_type_env) else None


def _format_cli_dynamic_ref_error(exc: BaseException) -> str:
    text = str(exc)
    if _PY_DYNAMIC_REF_HINT in text:
        return text.replace(_PY_DYNAMIC_REF_HINT, _CLI_DYNAMIC_REF_HINT)
    return f"{text}\n{_CLI_DYNAMIC_REF_HINT}"


def _module_dir(emit_dir: Path, package_name: str) -> Path:
    if emit_dir.name == package_name:
        return emit_dir
    return emit_dir / package_name


def _print_summary(result: Mapping[str, Any]) -> None:
    report = result["report"]
    errors = sum(1 for issue in report["issues"] if issue["level"] == "error")
    warnings = sum(1 for issue in report["issues"] if issue["level"] == "warning")
    print(f"ok={report['ok']} errors={errors} warnings={warnings}")
    print(f"canonical_sha256={result['canonical_sha256']}")
    print(f"inputs={result['inputs']}")
    print(f"computes={result['computes']}")


def _format_issue(issue: ValidationIssue) -> str:
    location_parts: list[str] = []
    if issue.get("series_id"):
        location_parts.append(str(issue["series_id"]))
    if issue.get("address"):
        location_parts.append(str(issue["address"]))
    location = ":".join(location_parts) if location_parts else "-"
    return f"{issue['level']} [{issue['code']}] {location}: {issue['message']}"


def _print_issues(report: ValidationReport, *, include_warnings: bool) -> None:
    for issue in report["issues"]:
        if issue["level"] == "error" or include_warnings:
            print(_format_issue(issue))


def _write_generated_files(result: Mapping[str, Any], output_dir: Path) -> None:
    files = result.get("generated_files")
    if not isinstance(files, dict):
        return
    output_dir.mkdir(parents=True, exist_ok=True)
    for filename, content in files.items():
        path = output_dir / filename
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding="utf-8")
        print(f"wrote {path}")
