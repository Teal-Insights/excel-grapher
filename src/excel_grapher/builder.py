from __future__ import annotations

from collections.abc import Iterable
from collections import deque
from pathlib import Path
import re

import openpyxl
import openpyxl.utils.cell

from .graph import DependencyGraph, NodeHook
from .node import Node
from .parser import expand_range, mask_spans, normalize_formula, parse_cell_refs, parse_range_refs_with_spans
from .resolver import build_named_range_map


def create_dependency_graph(
    workbook: Path | str | openpyxl.Workbook,
    targets: Iterable[str],
    *,
    max_depth: int = 50,
    expand_ranges: bool = True,
    max_range_cells: int = 5000,
    hooks: list[NodeHook] | None = None,
    load_values: bool = True,
) -> DependencyGraph:
    """
    Build a dependency graph starting from target cells.

    Current implementation supports basic A1 references and sheet-qualified references.
    Range expansion and named range resolution will be added incrementally via TDD.
    """

    def load_wb(data_only: bool) -> openpyxl.Workbook:
        if isinstance(workbook, openpyxl.Workbook):
            if data_only:
                raise ValueError("load_values=True is not supported when passing a Workbook instance")
            return workbook
        path = Path(workbook)
        keep_vba = path.suffix.lower() == ".xlsm"
        return openpyxl.load_workbook(path, data_only=data_only, keep_vba=keep_vba)

    wb_formulas = load_wb(data_only=False)
    wb_values = load_wb(data_only=True) if load_values and not isinstance(workbook, openpyxl.Workbook) else None

    graph = DependencyGraph()
    for h in hooks or []:
        graph.register_hook(h)

    named_ranges = build_named_range_map(wb_formulas)
    _NAME_TOKEN_RE = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\b(?!\s*!)")

    def parse_target(t: str) -> tuple[str, str]:
        if "!" not in t:
            raise ValueError(f"Target must be sheet-qualified: {t}")
        sheet, a1 = t.split("!", 1)
        if sheet not in wb_formulas.sheetnames:
            raise ValueError(f"Sheet not found: {sheet}")
        return sheet, a1

    def extract_deps(formula: str, current_sheet: str) -> list[tuple[str, str]]:
        if not formula.startswith("="):
            return []
        deps: list[tuple[str, str]] = []

        # 0) Expand ranges first, then mask them so later cell-ref parsing doesn't
        # misinterpret the range endpoint as a same-sheet ref (e.g. 'Sheet 2'!A1:B2).
        masked = formula
        if expand_ranges:
            spans: list[tuple[int, int]] = []
            for start, end, span in parse_range_refs_with_spans(formula):
                spans.append(span)
                sheet = start.sheet if start.sheet is not None else current_sheet
                deps.extend(
                    expand_range(
                        sheet=sheet,
                        start_col=start.column,
                        start_row=start.row,
                        end_col=end.column,
                        end_row=end.row,
                        max_cells=max_range_cells,
                    )
                )
            masked = mask_spans(masked, spans)

        for ref in parse_cell_refs(masked):
            sh = ref.sheet if ref.sheet is not None else current_sheet
            deps.append((sh, f"{ref.column}{ref.row}"))

        # 3) Named ranges
        for m in _NAME_TOKEN_RE.finditer(masked):
            token = m.group(1)
            resolved = named_ranges.get(token)
            if resolved is None:
                continue
            deps.append(resolved)

        # Deduplicate while preserving order
        seen: set[tuple[str, str]] = set()
        out: list[tuple[str, str]] = []
        for d in deps:
            if d in seen:
                continue
            seen.add(d)
            out.append(d)
        return out

    visited: set[str] = set()
    q: deque[tuple[str, str, int]] = deque()
    for t in targets:
        sh, a1 = parse_target(str(t))
        q.append((sh, a1, 0))

    try:
        while q:
            sheet, a1, depth = q.popleft()
            key = f"{sheet}!{a1}"
            if key in visited:
                continue
            visited.add(key)
            if depth > max_depth:
                continue

            ws_f = wb_formulas[sheet]
            raw = ws_f[a1].value
            is_formula = isinstance(raw, str) and raw.startswith("=")

            if is_formula:
                formula = str(raw)
                normalized = normalize_formula(formula, sheet, named_ranges)
                value = None
                if wb_values is not None:
                    value = wb_values[sheet][a1].value
                is_leaf = False
            else:
                formula = None
                normalized = None
                value = raw
                is_leaf = True

            col, row = openpyxl.utils.cell.coordinate_from_string(a1)
            node = Node(
                sheet=sheet,
                column=col,
                row=int(row),
                formula=formula,
                normalized_formula=normalized,
                value=value,
                is_leaf=is_leaf,
            )
            graph.add_node(node)

            if not is_formula:
                continue

            for dep_sheet, dep_a1 in extract_deps(formula, sheet):
                dep_key = f"{dep_sheet}!{dep_a1}"
                graph.add_edge(key, dep_key)
                if dep_key not in visited:
                    if dep_sheet not in wb_formulas.sheetnames:
                        continue
                    q.append((dep_sheet, dep_a1, depth + 1))
    finally:
        if wb_values is not None:
            wb_values.close()
        if not isinstance(workbook, openpyxl.Workbook):
            wb_formulas.close()

    return graph

