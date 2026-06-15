"""Top-level ``excel-grapher`` command-line interface."""

from __future__ import annotations

import argparse
import sys
from collections.abc import Sequence

from excel_grapher.cli import bindings


def build_parser() -> argparse.ArgumentParser:
    """Build the root ``excel-grapher`` argument parser."""
    parser = argparse.ArgumentParser(prog="excel-grapher")
    subparsers = parser.add_subparsers(dest="command", required=True)
    bindings.register(subparsers)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    """Run the ``excel-grapher`` CLI."""
    parser = build_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)
    if args.command == "bindings":
        return bindings.dispatch(args)
    parser.error(f"unknown command: {args.command!r}")
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
