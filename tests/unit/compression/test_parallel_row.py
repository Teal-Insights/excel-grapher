"""Unit tests for parallel row run detection."""

from __future__ import annotations

from excel_grapher.compression.parallel_row import (
    ParallelRun,
    RowCell,
    find_parallel_runs,
    find_parallel_runs_in_map,
    group_row_cells,
    merge_adjacent_runs,
    split_contiguous_row_segments,
)
from excel_grapher.compression.template_signature import template_signature, with_column_variable
from excel_grapher.core.address_keys import format_cell_key
from excel_grapher.core.formula_ast import FunctionCallNode

from .conftest import parse_formula


def _row_cell(sheet: str, col: str, row: int, formula: str) -> RowCell:
    key = format_cell_key(sheet, col, row)
    return RowCell(sheet=sheet, col=col, row=row, key=key, ast=parse_formula(formula))


def _if_row_cells() -> list[RowCell]:
    return [
        _row_cell("Chart Data", col, 177, f'=IF(Ext!D3="No",NA(),Ext!{col}87)')
        for col in ("D", "E", "F")
    ]


def test_group_row_cells_sorts_columns() -> None:
    ast_map = {
        format_cell_key("Sheet1", "F", 1): parse_formula("=Sheet1!F10*2"),
        format_cell_key("Sheet1", "D", 1): parse_formula("=Sheet1!D10*2"),
        format_cell_key("Sheet1", "E", 1): parse_formula("=Sheet1!E10*2"),
        format_cell_key("Sheet1", "A", 2): parse_formula("=1+1"),
    }
    grouped = group_row_cells(ast_map)
    assert list(grouped) == [("Sheet1", 1), ("Sheet1", 2)]
    row_one = grouped[("Sheet1", 1)]
    assert [cell.col for cell in row_one] == ["D", "E", "F"]


def test_split_contiguous_row_segments_gap() -> None:
    cells = [
        _row_cell("Sheet1", "D", 1, "=Sheet1!D10*2"),
        _row_cell("Sheet1", "E", 1, "=Sheet1!E10*2"),
        _row_cell("Sheet1", "G", 1, "=Sheet1!G10*2"),
        _row_cell("Sheet1", "H", 1, "=Sheet1!H10*2"),
        _row_cell("Sheet1", "I", 1, "=Sheet1!I10*2"),
    ]
    segments = split_contiguous_row_segments(cells)
    assert len(segments) == 2
    assert [cell.col for cell in segments[0]] == ["D", "E"]
    assert [cell.col for cell in segments[1]] == ["G", "H", "I"]


def test_find_parallel_runs_if_row() -> None:
    runs = find_parallel_runs(_if_row_cells())
    assert len(runs) == 1
    run = runs[0]
    assert run.sheet == "Chart Data"
    assert run.row == 177
    assert run.start_col == "D"
    assert run.end_col == "F"
    assert len(run.cells) == 3


def test_find_parallel_runs_multiply_row() -> None:
    cells = [_row_cell("Sheet1", col, 1, f"=Sheet1!{col}10*2") for col in ("D", "E", "F", "G")]
    runs = find_parallel_runs(cells)
    assert len(runs) == 1
    assert runs[0].start_col == "D"
    assert runs[0].end_col == "G"


def test_find_parallel_runs_rejects_length_two() -> None:
    cells = [
        _row_cell("Sheet1", "D", 1, "=Sheet1!D10*2"),
        _row_cell("Sheet1", "E", 1, "=Sheet1!E10*2"),
    ]
    assert find_parallel_runs(cells) == []


def test_find_parallel_runs_rejects_mismatched_formulas() -> None:
    cells = [
        _row_cell("Sheet1", "D", 1, "=Sheet1!D10*2"),
        _row_cell("Sheet1", "E", 1, "=Sheet1!E10*2"),
        _row_cell("Sheet1", "F", 1, "=Sheet1!F10*3"),
    ]
    assert find_parallel_runs(cells) == []


def test_find_parallel_runs_gap_limits_segment() -> None:
    ast_map = {
        format_cell_key("Sheet1", col, 10): parse_formula(f"=Sheet1!{col}10*2")
        for col in ("D", "E", "G", "H", "I")
    }
    runs = find_parallel_runs_in_map(ast_map)
    assert len(runs) == 1
    assert runs[0].start_col == "G"
    assert runs[0].end_col == "I"


def test_merge_adjacent_runs() -> None:
    cells_a = _if_row_cells()[:2]
    cells_b = [_if_row_cells()[2]]
    peer_a = tuple(cell.ast for cell in cells_a)
    peer_b = tuple(cell.ast for cell in _if_row_cells())
    sig_a = template_signature(
        with_column_variable(
            cells_a[0].ast,
            output_sheet="Chart Data",
            output_col="D",
            output_row=177,
            peer_asts=peer_a,
        )
    )
    sig_b = template_signature(
        with_column_variable(
            cells_b[0].ast,
            output_sheet="Chart Data",
            output_col="F",
            output_row=177,
            peer_asts=peer_b,
        )
    )
    left = ParallelRun(
        sheet="Chart Data",
        row=177,
        start_col="D",
        end_col="E",
        cells=tuple(cells_a),
        signature=sig_a,
    )
    right = ParallelRun(
        sheet="Chart Data",
        row=177,
        start_col="F",
        end_col="F",
        cells=tuple(cells_b),
        signature=sig_b,
    )
    merged = merge_adjacent_runs((left, right))
    assert len(merged) == 1
    assert merged[0].start_col == "D"
    assert merged[0].end_col == "F"
    assert len(merged[0].cells) == 3


def test_find_parallel_runs_in_map_multiple_rows() -> None:
    ast_map = {
        format_cell_key("Chart Data", col, 177): parse_formula(f'=IF(Ext!D3="No",NA(),Ext!{col}87)')
        for col in ("D", "E", "F")
    }
    ast_map.update(
        {
            format_cell_key("Sheet1", col, 1): parse_formula(f"=Sheet1!{col}10*2")
            for col in ("D", "E", "F")
        }
    )
    runs = find_parallel_runs_in_map(ast_map)
    assert len(runs) == 2
    by_sheet = {(run.sheet, run.row): run for run in runs}
    assert by_sheet[("Chart Data", 177)].end_col == "F"
    assert by_sheet[("Sheet1", 1)].end_col == "F"


def test_parallel_run_normalized_template() -> None:
    runs = find_parallel_runs(_if_row_cells())
    template = runs[0].normalized_template()
    assert isinstance(template, FunctionCallNode)
    assert template.name.upper() == "IF"
