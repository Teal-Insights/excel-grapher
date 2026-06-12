#!/usr/bin/env python3
"""Validate ffv2 series bindings and smoke-test every generated setter/compute.

Usage (workbook colocated with bindings sidecar):

    uv run python scripts/test_ffv2_bindings.py
    uv run python scripts/test_ffv2_bindings.py --workbook /path/to/ffv2.xlsx

Expects ``ffv2.bindings.yaml`` next to the workbook, or pass ``--bindings``.
Uses the repo fixture when ``--generate-fixture`` is set (no local xlsx required).
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tests.integration.user_flows.ffv2_harness import (  # noqa: E402
    compute_names,
    run_ffv2_binding_checks,
    setter_names,
    validate_ffv2_bindings,
)
from tests.integration.user_flows.utils import write_ffv2_workbook  # noqa: E402

DEFAULT_FIXTURE_BINDINGS = ROOT / "tests" / "fixtures" / "series_bindings" / "ffv2.yaml"


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--workbook",
        type=Path,
        default=Path("ffv2.xlsx"),
        help="Path to ffv2.xlsx (default: ./ffv2.xlsx)",
    )
    parser.add_argument(
        "--bindings",
        type=Path,
        default=None,
        help="Path to bindings YAML (default: <workbook>.bindings.yaml)",
    )
    parser.add_argument(
        "--module-dir",
        type=Path,
        default=Path("module"),
        help="Directory for generated package files (default: ./module)",
    )
    parser.add_argument(
        "--generate-fixture",
        action="store_true",
        help="Build ffv2.xlsx from the repo test helper instead of reading a local file",
    )
    parser.add_argument(
        "--use-fixture-bindings",
        action="store_true",
        help=f"Use repo fixture bindings at {DEFAULT_FIXTURE_BINDINGS}",
    )
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    workbook = args.workbook
    if args.generate_fixture:
        write_ffv2_workbook(workbook)

    if not workbook.is_file():
        print(f"Workbook not found: {workbook}", file=sys.stderr)
        return 1

    bindings_path = args.bindings
    if bindings_path is None:
        bindings_path = (
            DEFAULT_FIXTURE_BINDINGS
            if args.use_fixture_bindings
            else workbook.with_suffix(".bindings.yaml")
        )
    if not bindings_path.is_file():
        print(f"Bindings not found: {bindings_path}", file=sys.stderr)
        return 1

    from excel_grapher.series_bindings import load_series_bindings

    bindings = load_series_bindings(bindings_path)
    validation = validate_ffv2_bindings(workbook, bindings)
    print(json.dumps(validation["report"], indent=2, default=str))
    if not validation["report"]["ok"]:
        return 1

    print(f"canonical_sha256={validation['canonical_sha256']}")
    print(f"setters={setter_names(bindings)}")
    print(f"computes={compute_names(bindings)}")

    result = run_ffv2_binding_checks(
        workbook,
        bindings_path,
        module_dir=args.module_dir,
    )
    for filename, content in result["generated_files"].items():
        path = args.module_dir / filename
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding="utf-8")
        print(f"wrote {path}")

    print("All setter and compute functions passed smoke checks.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
