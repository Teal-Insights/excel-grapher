"""``excel-grapher bindings`` command group."""

from __future__ import annotations

import argparse
import json
import sys
import tempfile
from collections.abc import Mapping
from pathlib import Path
from typing import Any

from excel_grapher.series_bindings.load import SeriesBindingsLoadError
from excel_grapher.series_bindings.schema import SeriesBindingsSchemaError
from excel_grapher.series_bindings.smoke import BindingsSmokeError
from excel_grapher.series_bindings.types import ValidationIssue, ValidationReport
from excel_grapher.series_bindings.workflow import (
    generate_bindings_modules,
    resolve_bindings_path,
    run_binding_checks,
    validate_bindings_workbook,
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
        help="Generate modules and smoke-test setters/computes after validation",
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


def dispatch(args: argparse.Namespace) -> int:
    """Dispatch a ``bindings`` subcommand."""
    if args.bindings_command == "validate":
        return cmd_validate(args)
    print(f"Unknown bindings command: {args.bindings_command}", file=sys.stderr)
    return 2


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
        result = validate_bindings_workbook(workbook, bindings_path)
    except SeriesBindingsLoadError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except SeriesBindingsSchemaError as exc:
        print(f"Binding sidecar schema error:\n  {exc}", file=sys.stderr)
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
                    )
            else:
                module_dir = _module_dir(args.emit_dir, args.package_name)
                check_result = run_binding_checks(
                    workbook,
                    bindings_path,
                    module_dir=module_dir,
                    package_name=args.package_name,
                    smoke_test=True,
                )
                if not args.json:
                    _write_generated_files(check_result, module_dir)
        elif args.emit_dir is not None:
            module_dir = _module_dir(args.emit_dir, args.package_name)
            files = generate_bindings_modules(
                result["graph"],
                targets=result["targets"],
                bindings=result["bindings"],
                workbook=workbook,
            )
            _write_generated_files({"generated_files": files}, module_dir)
    except (BindingsSmokeError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 1

    if not args.json and args.smoke_test:
        print("All setter and compute functions passed smoke checks.")
    return 0


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
    print(f"setters={result['setters']}")
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
