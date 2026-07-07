"""Detect contiguous parallel formula runs on the same row."""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Mapping, Sequence
from dataclasses import dataclass

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import normalize_key, parse_address, quote_sheet_if_needed
from excel_grapher.core.formula_ast import AstNode, FunctionCallNode

from .nodes import ParallelFormulaNode
from .stats import CompressionStats
from .template_signature import TemplateSignature, template_signature, with_column_variable
from .types import CompressedNode

__all__ = [
    "ParallelRun",
    "RowCell",
    "apply_parallel_row",
    "build_parallel_node",
    "find_parallel_runs",
    "find_parallel_runs_in_map",
    "group_row_cells",
    "materialize_parallel_node",
    "merge_adjacent_runs",
    "parallel_artifact_key",
    "split_contiguous_row_segments",
]


@dataclass(frozen=True, slots=True)
class RowCell:
    """One formula cell positioned on a workbook row."""

    sheet: str
    col: str
    row: int
    key: str
    ast: AstNode


@dataclass(frozen=True, slots=True)
class ParallelRun:
    """A contiguous same-row group sharing one normalized template signature."""

    sheet: str
    row: int
    start_col: str
    end_col: str
    cells: tuple[RowCell, ...]
    signature: TemplateSignature

    def normalized_template(self) -> AstNode:
        """Return the column-normalized template AST for this run."""
        peer_asts = tuple(cell.ast for cell in self.cells)
        anchor = self.cells[0]
        normalized = with_column_variable(
            anchor.ast,
            output_sheet=self.sheet,
            output_col=anchor.col,
            output_row=self.row,
            peer_asts=peer_asts,
        )
        return normalized


def group_row_cells(ast_map: Mapping[str, AstNode]) -> dict[tuple[str, int], list[RowCell]]:
    """Group per-cell ASTs by `(sheet, row)` with columns sorted left to right."""
    grouped: dict[tuple[str, int], list[RowCell]] = defaultdict(list)
    for cell_key, ast in ast_map.items():
        normalized_key = normalize_key(cell_key)
        sheet, coord = parse_address(normalized_key)
        col = "".join(character for character in coord if character.isalpha())
        row = int("".join(character for character in coord if character.isdigit()))
        grouped[(sheet, row)].append(
            RowCell(sheet=sheet, col=col, row=row, key=normalized_key, ast=ast)
        )

    for cells in grouped.values():
        cells.sort(key=lambda cell: fastpyxl.utils.cell.column_index_from_string(cell.col))
    return dict(grouped)


def split_contiguous_row_segments(cells: Sequence[RowCell]) -> list[list[RowCell]]:
    """Split a sorted row into maximal contiguous column segments."""
    if not cells:
        return []

    segments: list[list[RowCell]] = []
    current = [cells[0]]
    for previous, current_cell in zip(cells, cells[1:], strict=False):
        prev_index = fastpyxl.utils.cell.column_index_from_string(previous.col)
        curr_index = fastpyxl.utils.cell.column_index_from_string(current_cell.col)
        if curr_index - prev_index == 1:
            current.append(current_cell)
            continue
        segments.append(current)
        current = [current_cell]
    segments.append(current)
    return segments


def find_parallel_runs(
    cells: Sequence[RowCell],
    *,
    min_len: int = 3,
) -> list[ParallelRun]:
    """Find maximal contiguous runs with matching normalized template signatures."""
    if len(cells) < min_len:
        return []

    runs: list[ParallelRun] = []
    index = 0
    while index <= len(cells) - min_len:
        best_end: int | None = None
        best_signature: TemplateSignature | None = None
        for end in range(index + min_len, len(cells) + 1):
            chunk = cells[index:end]
            signature = _matching_run_signature(chunk)
            if signature is not None:
                best_end = end
                best_signature = signature
        if best_end is None or best_signature is None:
            index += 1
            continue
        chunk = cells[index:best_end]
        runs.append(
            ParallelRun(
                sheet=chunk[0].sheet,
                row=chunk[0].row,
                start_col=chunk[0].col,
                end_col=chunk[-1].col,
                cells=tuple(chunk),
                signature=best_signature,
            )
        )
        index = best_end
    return runs


def merge_adjacent_runs(runs: Sequence[ParallelRun]) -> list[ParallelRun]:
    """Merge neighboring runs on the same row that share a template signature."""
    if not runs:
        return []

    ordered = sorted(
        runs,
        key=lambda run: fastpyxl.utils.cell.column_index_from_string(run.start_col),
    )
    merged: list[ParallelRun] = [ordered[0]]
    for run in ordered[1:]:
        previous = merged[-1]
        prev_end = fastpyxl.utils.cell.column_index_from_string(previous.end_col)
        run_start = fastpyxl.utils.cell.column_index_from_string(run.start_col)
        if (
            previous.sheet == run.sheet
            and previous.row == run.row
            and previous.signature == run.signature
            and run_start - prev_end == 1
        ):
            merged[-1] = ParallelRun(
                sheet=previous.sheet,
                row=previous.row,
                start_col=previous.start_col,
                end_col=run.end_col,
                cells=previous.cells + run.cells,
                signature=previous.signature,
            )
            continue
        merged.append(run)
    return merged


def find_parallel_runs_in_map(
    ast_map: Mapping[str, AstNode],
    *,
    min_len: int = 3,
) -> list[ParallelRun]:
    """Detect parallel runs across every row in `ast_map`."""
    runs: list[ParallelRun] = []
    for cells in group_row_cells(ast_map).values():
        for segment in split_contiguous_row_segments(cells):
            runs.extend(merge_adjacent_runs(find_parallel_runs(segment, min_len=min_len)))
    return runs


def _matching_run_signature(cells: Sequence[RowCell]) -> TemplateSignature | None:
    if not cells:
        return None
    peer_asts = tuple(cell.ast for cell in cells)
    signatures: list[TemplateSignature] = []
    for cell in cells:
        normalized = with_column_variable(
            cell.ast,
            output_sheet=cell.sheet,
            output_col=cell.col,
            output_row=cell.row,
            peer_asts=peer_asts,
        )
        signatures.append(template_signature(normalized))
    first = signatures[0]
    if any(signature != first for signature in signatures[1:]):
        return None
    return first


def parallel_artifact_key(run: ParallelRun) -> str:
    """Return a synthetic map key for a compressed parallel row artifact."""
    return f"parallel:{quote_sheet_if_needed(run.sheet)}!{run.row}:{run.start_col}:{run.end_col}"


def build_parallel_node(run: ParallelRun) -> ParallelFormulaNode:
    """Build a `ParallelFormulaNode` artifact from a detected run."""
    template = run.normalized_template()
    condition, if_true, if_false = _if_projection(template)
    return ParallelFormulaNode(
        sheet=run.sheet,
        template=template,
        start_col=run.start_col,
        end_col=run.end_col,
        output_row=run.row,
        condition=condition,
        if_true=if_true,
        if_false=if_false,
    )


def apply_parallel_row(
    ast_map: Mapping[str, AstNode],
    stats: CompressionStats | None = None,
    *,
    min_run_length: int = 3,
) -> dict[str, CompressedNode]:
    """Merge parallel row runs into `ParallelFormulaNode` artifacts.

    Args:
        ast_map: Per-cell formula AST map.
        stats: Optional stats object for rule contribution counters.
        min_run_length: Minimum contiguous columns required to form a group.

    Returns:
        Mixed compressed map with per-cell keys removed when absorbed into
        parallel artifacts.
    """
    normalized = {normalize_key(cell_key): ast for cell_key, ast in ast_map.items()}
    runs = find_parallel_runs_in_map(normalized, min_len=min_run_length)
    absorbed = {cell.key for run in runs for cell in run.cells}

    result: dict[str, CompressedNode] = {
        key: ast for key, ast in normalized.items() if key not in absorbed
    }
    for run in runs:
        result[parallel_artifact_key(run)] = build_parallel_node(run)

    if stats is not None:
        stats.contribution_for("parallel_if_row").record(
            cells_affected=sum(len(run.cells) for run in runs),
            emission_units_saved=sum(len(run.cells) - 1 for run in runs),
        )
    return result


def materialize_parallel_node(node: ParallelFormulaNode) -> dict[str, AstNode]:
    """Expand a `ParallelFormulaNode` to per-cell ASTs."""
    from .expand import materialize_parallel_node as _materialize

    return _materialize(node)


def _if_projection(
    template: AstNode,
) -> tuple[AstNode | None, AstNode | None, AstNode | None]:
    if (
        isinstance(template, FunctionCallNode)
        and template.name.upper() == "IF"
        and len(template.args) == 3
    ):
        return template.args[0], template.args[1], template.args[2]
    return None, None, None
