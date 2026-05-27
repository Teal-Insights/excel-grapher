"""Unit tests for derive_output_series."""

from __future__ import annotations

from pathlib import Path

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    derive_output_series,
    expand_data_range,
    load_series_bindings,
)

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "series_bindings"


def test_derive_output_series_from_merged_manifest(tmp_path: Path) -> None:
    from tests.unit.series_bindings.test_resolve import _write_borvelia_workbook

    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    shard_dir = tmp_path / "shards"
    shard_dir.mkdir()
    (shard_dir / "in.bindings.yaml").write_text(
        (FIXTURES / "shard_borvelia_input.yaml").read_text(encoding="utf-8"),
        encoding="utf-8",
    )
    (shard_dir / "out.bindings.yaml").write_text(
        (FIXTURES / "shard_borvelia_output.yaml").read_text(encoding="utf-8"),
        encoding="utf-8",
    )
    bindings = load_series_bindings(shard_dir)
    graph = create_dependency_graph(
        wb_path,
        expand_data_range("Inputs!F5:J5"),
        load_values=True,
    )
    output_series = derive_output_series(graph, bindings, workbook=wb_path)
    assert len(output_series) == 1
    assert output_series[0]["compute_name"] == "compute_borvelia_primary_balance"
    assert len(output_series[0]["cells"]) == 5
