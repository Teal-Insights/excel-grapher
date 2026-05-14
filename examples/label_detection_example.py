#!/usr/bin/env python3
"""
Example: attach row/column labels to graph nodes via ``create_dependency_graph``.

This uses the in-library label pipeline (``label_detection`` / ``label_behaviors``)
instead of post-processing the graph. For graph cache JSON, include the same
settings under ``extraction_params`` using ``label_detection_config_to_jsonable``.

Run from the repository root::

    uv run python examples/label_detection_example.py

Or with an explicit region rule (see ``--mode region``)::

    uv run python examples/label_detection_example.py --mode region
"""

from __future__ import annotations

import argparse
import json
import tempfile
from pathlib import Path

import xlsxwriter

from excel_grapher import (
    BehaviorRule,
    LabelDetectionConfig,
    RegionLabelParams,
    RegionSelector,
    create_dependency_graph,
    label_detection_config_to_jsonable,
    region_specs_from_ranges,
)


def write_demo_workbook(path: Path) -> None:
    """Minimal sheet: left labels (col A), header row 1, formula in C2."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_string(0, 0, "Item")
    ws.write_number(0, 1, 2024)
    ws.write_string(0, 2, "Value")
    ws.write_string(1, 0, "Revenue")
    ws.write_number(1, 1, 2024)
    ws.write_formula(1, 2, "=B2*10", None, 100.0)
    wb.close()


def _config_heuristic() -> LabelDetectionConfig:
    return LabelDetectionConfig(enabled=True)


def _config_region() -> LabelDetectionConfig:
    return LabelDetectionConfig(
        enabled=True,
        rules=(
            BehaviorRule(
                name="dataBlock",
                selector=RegionSelector(
                    include=region_specs_from_ranges(["Sheet1!A1:C10"]),
                ),
                behaviors=("region_left_label_columns", "region_header_rows"),
                stop_after_match=True,
                region_params=RegionLabelParams(
                    label_columns=("A",),
                    header_rows=(1,),
                ),
            ),
        ),
        fallback_behaviors=(),
    )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--mode",
        choices=("heuristic", "region"),
        default="heuristic",
        help="heuristic: scan left/up; region: explicit label column + header row",
    )
    args = parser.parse_args()

    cfg = _config_heuristic() if args.mode == "heuristic" else _config_region()

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        wb_path = Path(tmp.name)
    try:
        write_demo_workbook(wb_path)
        graph = create_dependency_graph(
            wb_path,
            ["Sheet1!C2"],
            load_values=True,
            label_detection=cfg,
        )
        node = graph.get_node("Sheet1!C2")
        if node is None:
            raise SystemExit("expected node Sheet1!C2")
        print("Sheet1!C2 metadata:")
        print(f"  row_labels:    {node.metadata.get('row_labels')}")
        print(f"  column_labels: {node.metadata.get('column_labels')}")
        print()
        print("extraction_params fragment for graph cache meta:")
        fragment = {"label_detection": label_detection_config_to_jsonable(cfg)}
        print(json.dumps(fragment, indent=2))
    finally:
        wb_path.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
