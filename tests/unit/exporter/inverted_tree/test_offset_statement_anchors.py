"""OFFSET anchors infer only over the current statement (#772).

Same-shape OFFSET formulas in one host series may read different bound
lookup blocks. Statement partitioning already separates those producer
accesses; block-anchor inference must not require the second block's
anchor to belong to the first.
"""

from __future__ import annotations

from pathlib import Path
from typing import Literal

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)


def _offset_blocks_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "offset_statement_anchors.xlsx",
        {
            "Tables": {
                "A1": 2020,
                "B1": 2021,
                "A2": 1,
                "B2": 2,
                "A3": 10,
                "B3": 20,
            },
            "Engine": {
                "A1": 0,
                "C2": 2020,
                "C3": 2021,
                "B2": "=OFFSET(Tables!$A$2,0,$A$1)",
                "B3": "=OFFSET(Tables!$A$3,0,$A$1)",
            },
        },
    )


def _offset_blocks_bindings() -> dict:
    return bindings_document(
        series_entry("selector", "Engine!A1", layout="scalar", direction="input", dtype="int"),
        series_entry(
            "first_labels",
            "Tables!A2:B2",
            layout="series",
            direction="constant",
            header_row=1,
        ),
        series_entry(
            "second_labels",
            "Tables!A3:B3",
            layout="series",
            direction="constant",
            header_row=1,
        ),
        series_entry(
            "labels",
            "Engine!B2:B3",
            layout="series",
            direction="output",
            label_column="C",
            compute_name="compute_labels",
        ),
    )


def test_offset_host_partitions_on_distinct_lookup_blocks(tmp_path: Path) -> None:
    catalog, _deps, _graph = inverted_graph_parts(
        _offset_blocks_workbook(tmp_path),
        _offset_blocks_bindings(),
    )
    labels = catalog.get("labels")
    assert [(stmt.start, stmt.stop) for stmt in labels.statements] == [(0, 1), (1, 2)]


@pytest.mark.parametrize("force_rung", [None, 3])
def test_offset_anchor_inference_stays_in_statement(
    tmp_path: Path,
    force_rung: Literal[3] | None,
) -> None:
    workbook = _offset_blocks_workbook(tmp_path)
    modules = generate_inverted(
        workbook,
        _offset_blocks_bindings(),
        force_rung=force_rung,
    )
    pkg = load_package(modules, tmp_path, name=f"offset_stmt_{force_rung}")
    assert pkg.compute_labels(selector=0) == pytest.approx((1.0, 10.0))
    assert pkg.compute_labels(selector=1) == pytest.approx((2.0, 20.0))
