"""``excel-grapher project`` command group."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from excel_grapher.exporter.compression_workflow import (
    CompressionMethod,
    build_embedding_provider,
    build_similarity_config,
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
        "--preserve",
        nargs="+",
        default=None,
        metavar="NODE",
        help="Node keys that must not be inlined away",
    )
    compress_parser.add_argument(
        "--embedding-provider",
        choices=("mock", "openai"),
        default="mock",
        help="Embedding backend for similarity scoring (default: mock)",
    )
    compress_parser.add_argument(
        "--embedding-model",
        default="text-embedding-3-small",
        help="OpenAI embedding model when --embedding-provider=openai",
    )
    compress_parser.add_argument(
        "--max-candidates",
        type=int,
        default=None,
        help="Cap compressible candidates during similarity search",
    )
    compress_parser.add_argument(
        "--top-n-packings",
        type=int,
        default=None,
        help="Cap packings scored before similarity selection",
    )
    compress_parser.add_argument(
        "--score-flatness-epsilon",
        type=float,
        default=None,
        help="When top packing scores differ by less than this, prefer max reduction",
    )
    compress_parser.add_argument(
        "--no-fallback-to-optimal",
        action="store_true",
        help="Disable reduction-only fallback when similarity scores are flat",
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
    preserve = set(args.preserve) if args.preserve else None
    provider = None
    config = None
    if method == "similarity":
        try:
            provider = build_embedding_provider(
                args.embedding_provider,
                model=args.embedding_model,
            )
        except (ImportError, ValueError) as exc:
            print(f"Embedding provider error: {exc}", file=sys.stderr)
            return 1
        config = build_similarity_config(
            max_candidates=args.max_candidates,
            top_n_packings=args.top_n_packings,
            score_flatness_epsilon=args.score_flatness_epsilon,
            fallback_to_optimal=False if args.no_fallback_to_optimal else None,
            embedding_model=args.embedding_model if args.embedding_provider == "openai" else None,
        )

    try:
        projection, report = compress_workbook(
            workbook,
            list(args.targets),
            method=method,
            provider=provider,
            config=config,
            preserve=preserve,
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
