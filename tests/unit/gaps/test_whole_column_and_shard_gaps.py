"""Binding shard merge and whole-column formula gaps (minimal reproduction)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
import xlsxwriter
import yaml

from excel_grapher import DependencyGraph, Node, create_dependency_graph
from excel_grapher.evaluator.errors import ParseError
from excel_grapher.evaluator.parser import parse
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.series_bindings import SeriesBindingsLoadError, load_series_bindings
from excel_grapher.series_bindings.load import merge_series_binding_documents

NESTED_INDEX_MATCH_WHOLE_COLUMN_FORMULA = (
    "=INDEX('QB - Stafford'!C:C,"
    "MATCH(INDEX(C3:V3,MATCH(MAX(C4:V4),C4:V4,0)),'QB - Stafford'!A:A,0))"
)


def _minimal_binding_shard(*, concept_id: str, series_id: str) -> dict[str, Any]:
    return {
        "schema_version": "1.3.0",
        "workbook": "w.xlsx",
        "concept_scheme": {
            "id": concept_id,
            "concepts": [{"id": "TIME_PERIOD", "dtype": "string"}],
        },
        "series": [
            {
                "id": series_id,
                "sheet": "S1",
                "data_range": "S1!A1",
                "layout": "scalar",
                "structure": {
                    "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
                    "dimensions": [],
                },
                "key": [],
            }
        ],
    }


def _write_binding_shard(path: Path, *, concept_id: str, series_id: str) -> None:
    path.write_text(
        yaml.safe_dump(_minimal_binding_shard(concept_id=concept_id, series_id=series_id)),
        encoding="utf-8",
    )


def _write_whole_column_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    qb = wb.add_worksheet("QB - Stafford")
    sheet1 = wb.add_worksheet("Sheet1")
    qb.write_string(4, 0, "target")
    qb.write_string(4, 2, "A SEA")
    sheet1.write_formula(
        0,
        0,
        "=INDEX('QB - Stafford'!C:C,MATCH(\"target\",'QB - Stafford'!A:A,0))",
        None,
        "A SEA",
    )
    wb.close()


def test_binding_shards_reject_mismatched_concept_scheme() -> None:
    aggregate = _minimal_binding_shard(concept_id="team_summary", series_id="agg_stat")
    player = _minimal_binding_shard(concept_id="player_game_log", series_id="player_stat")
    with pytest.raises(SeriesBindingsLoadError, match="concept_scheme mismatch"):
        merge_series_binding_documents([aggregate, player])


def test_bindings_directory_load_fails_concept_scheme_merge(tmp_path: Path) -> None:
    shard_dir = tmp_path / "shards"
    shard_dir.mkdir()
    _write_binding_shard(
        shard_dir / "summary.bindings.yaml", concept_id="team_summary", series_id="agg_stat"
    )
    _write_binding_shard(
        shard_dir / "player.bindings.yaml", concept_id="player_game_log", series_id="player_stat"
    )
    with pytest.raises(SeriesBindingsLoadError, match="concept_scheme mismatch"):
        load_series_bindings(shard_dir, validate=False)


def test_nested_index_match_formula_raw_parse_rejects_unqualified_local_refs() -> None:
    with pytest.raises(ParseError, match="sheet-qualified"):
        parse(NESTED_INDEX_MATCH_WHOLE_COLUMN_FORMULA)


def test_codegen_export_rejects_whole_column_ast_node(tmp_path: Path) -> None:
    workbook = tmp_path / "whole_column.xlsx"
    _write_whole_column_workbook(workbook)
    graph = create_dependency_graph(
        workbook,
        ["Sheet1!A1"],
        load_values=True,
        max_range_cells=2,
        use_cached_dynamic_refs=True,
    )
    with pytest.raises(ValueError, match="WholeColumnNode"):
        CodeGenerator(graph).generate(["Sheet1!A1"])


def test_codegen_export_rejects_whole_column_in_multi_target_batch(tmp_path: Path) -> None:
    workbook = tmp_path / "whole_column_batch.xlsx"
    _write_whole_column_workbook(workbook)
    graph = create_dependency_graph(
        workbook,
        ["Sheet1!A1", "QB - Stafford!A5"],
        load_values=True,
        max_range_cells=2,
        use_cached_dynamic_refs=True,
    )
    with pytest.raises(ValueError, match="WholeColumnNode"):
        CodeGenerator(graph).generate(["Sheet1!A1", "QB - Stafford!A5"])


def test_codegen_modules_reject_whole_column_ast_node(tmp_path: Path) -> None:
    workbook = tmp_path / "whole_column_modular.xlsx"
    _write_whole_column_workbook(workbook)
    graph = create_dependency_graph(
        workbook,
        ["Sheet1!A1"],
        load_values=True,
        max_range_cells=2,
        use_cached_dynamic_refs=True,
    )
    with pytest.raises(ValueError, match="WholeColumnNode"):
        CodeGenerator(graph).generate_modules(["Sheet1!A1"])


def test_manual_graph_codegen_rejects_whole_column_without_workbook_io() -> None:
    graph = DependencyGraph()
    graph.add_node(
        Node(
            sheet="QB - Stafford",
            column="A",
            row=5,
            formula=None,
            normalized_formula=None,
            value="target",
            is_leaf=True,
        )
    )
    graph.add_node(
        Node(
            sheet="QB - Stafford",
            column="C",
            row=5,
            formula=None,
            normalized_formula=None,
            value="A SEA",
            is_leaf=True,
        )
    )
    graph.add_node(
        Node(
            sheet="Sheet1",
            column="A",
            row=1,
            formula=NESTED_INDEX_MATCH_WHOLE_COLUMN_FORMULA,
            normalized_formula=(
                "=INDEX('QB - Stafford'!C:C,MATCH(\"target\",'QB - Stafford'!A:A,0))"
            ),
            value=None,
            is_leaf=False,
        )
    )
    with pytest.raises(ValueError, match="WholeColumnNode"):
        CodeGenerator(graph).generate(["Sheet1!A1"])
