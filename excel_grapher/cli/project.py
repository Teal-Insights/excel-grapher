"""``excel-grapher project`` command group."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from excel_grapher.exporter.compression_workflow import (
    CompressionMethod,
    compress_workbook,
    format_report_text,
)


def register(subparsers: argparse._SubParsersAction[argparse.ArgumentParser]) -> None:
    """Register the ``project`` command group."""
    project_parser = subparsers.add_parser(
        "project",
        help="Graph projection and compression tools",
    )
    project_sub = project_parser.add_subparsers(dest="project_command", required=True)
    compress_parser = project_sub.add_parser(
        "compress",
        help="Build a dependency graph and compress it",
    )
    compress_parser.add_argument("workbook", type=Path, help="Path to the .xlsx workbook")
    compress_parser.add_argument(
        "--targets",
        nargs="+",
        required=True,
        help="Target cell addresses (e.g. Engine!C20 Engine!D20)",
    )
    compress_parser.add_argument(
        "--method",
        choices=("similarity", "optimal", "identity"),
        default="similarity",
        help="Compression strategy (default: similarity)",
    )
    compress_parser.add_argument(
        "--json",
        action="store_true",
        help="Print the compression report as JSON",
    )
    compress_parser.add_argument(
        "--manifest-out",
        type=Path,
        default=None,
        help="Write the full projection manifest JSON to this path",
    )


def dispatch(args: argparse.Namespace) -> int:
    """Dispatch a ``project`` subcommand."""
    if args.project_command == "compress":
        return cmd_compress(args)
    print(f"Unknown project command: {args.project_command}", file=sys.stderr)
    return 2


def cmd_compress(args: argparse.Namespace) -> int:
    """Run ``excel-grapher project compress``."""
    workbook: Path = args.workbook
    if not workbook.is_file():
        print(f"Workbook not found: {workbook}", file=sys.stderr)
        return 1

    method: CompressionMethod = args.method
    try:
        projection, report = compress_workbook(
            workbook,
            list(args.targets),
            method=method,
        )
    except Exception as exc:
        print(f"Compression failed: {exc}", file=sys.stderr)
        return 1

    if args.json:
        print(json.dumps(report.to_dict(), indent=2))
    else:
        print(format_report_text(report))

    if args.manifest_out is not None:
        manifest = projection.manifest
        args.manifest_out.parent.mkdir(parents=True, exist_ok=True)
        args.manifest_out.write_text(
            json.dumps(manifest.to_dict(), indent=2),
            encoding="utf-8",
        )
        if not args.json:
            print()
            print(f"Wrote manifest: {args.manifest_out}")

    return 0
