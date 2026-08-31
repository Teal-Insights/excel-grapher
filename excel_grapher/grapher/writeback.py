"""Write a `GraphReadView` to a new Excel workbook.

This is workbook I/O next to `create_dependency_graph`, not Python codegen.
`write_workbook` is an export backend on the same read surface `CodeGenerator`
already accepts (`DependencyGraph` or `ProjectionResult`).
"""

from __future__ import annotations

import os
from datetime import date, datetime
from pathlib import Path

from fastpyxl import Workbook
from fastpyxl.utils.cell import coordinate_from_string
from fastpyxl.utils.exceptions import CellCoordinatesException
from fastpyxl.workbook.defined_name import DefinedName
from fastpyxl.worksheet.formula import ArrayFormula
from fastpyxl.worksheet.worksheet import Worksheet

from excel_grapher.core.address_keys import quote_sheet_if_needed
from excel_grapher.core.formula_ast import FormulaStyle, render_formula
from excel_grapher.core.types import XlError

from .graph import GraphReadView
from .node import NodeView
from .shared_formulas import (
    SharedFormulasMode,
    parse_shared_formulas_mode,
    shared_formula_cell_values,
)


def write_workbook(
    graph: GraphReadView,
    destination: Path | str,
    *,
    formula_style: FormulaStyle = FormulaStyle.A1_EXCEL,
    coerce_relative_refs: bool = False,
    overwrite: bool = False,
    include_defined_names: bool = True,
    shared_formulas: SharedFormulasMode = "auto",
) -> None:
    """Write `graph` to a new `.xlsx` at `destination`.

    Accepts any `GraphReadView`: a mutable `DependencyGraph` (the workbook
    after `set_node_ast` / `set_node_value` / `move_node` / in-place
    `compress_*`) or a non-mutating `ProjectionResult`. Formula rewriting
    is an intended write-back use case; compressed or projected views are
    written as they are. Output contains only sheets and cells from the
    view. Styles, charts, VBA, and cells outside the view are omitted
    (accepted v1 lossiness). Vacated `move_node` addresses are simply
    absent; there is no template to clear.

    Two write orders: move then write persists current keys on this
    `DependencyGraph` (relatives already rewritten so resolved targets
    match the pre-move meaning). Project then write exports the projected
    clone. A `ProjectionResult` is a snapshot from `project()`; a later
    `move_node` on the original graph does not update it. Re-run
    projection after geometry edits to include them in a projected
    workbook.

    Formula cells are spelled from current `formula_ast` via `render_formula`
    (never from the opt-in raw `Node.formula` audit string). Defined names are
    expanded to A1 before parse, so written formulas are expanded A1 (for
    example `=A1*Settings!$B$2`, not `=A1*TaxRate`). That is expected, not a
    bug. When `include_defined_names` is True (default), workbook-global names
    from `graph.named_ranges` and `graph.named_range_ranges` are written so
    aliases still exist in Excel. Names the writer cannot express as a cell
    or rectangle are refused. Scope is not stored on the graph maps, so
    names are emitted workbook-global. The writer does not invent a name from
    an expanded AST.

    Leaves write `node.value`. Formula cells do not receive
    evaluator-computed cached results. Cells extracted as `ArrayFormula` are
    written back as `ArrayFormula(text, ref=...)` using the observed spill /
    CSE `ref`. fastpyxl does not distinguish legacy CSE from dynamic-array
    spills, so write-back emits `t="array"` for both. A flagged array cell
    with no observed `ref` is refused rather than written as a scalar formula.

    When `graph.formula_shapes` is warm, contiguous autofill runs that share
    an interned relative shape become one Excel shared formula (`t="shared"`).
    The overlay is opt-in and is not rewarmed here; missing or stale shapes,
    gaps, mixed axes, array formulas, and `INDIRECT` stay per-cell A1. See
    `shared_formulas`.

    Args:
        graph: Read view whose cells are written.
        destination: Output path. Parent directories must already exist.
        formula_style: Reference spelling. Default `A1_EXCEL` keeps `$` on
            absolute axes and omits the host sheet prefix. `R1C1` is not
            persisted for normal formula cells; shared-formula groups use
            interned relative shapes (the R1C1 dialect) and store the master
            cell in A1, which is what Excel and fastpyxl persist on disk.
        coerce_relative_refs: If True, bind relative axes to absolute
            indexes before spelling (`$` on both axes in `A1_EXCEL`).
            Shared-formula grouping is skipped (absolute fills are not a
            shared relative shape).
        overwrite: If False (default), raise when `destination` exists.
            The source workbook is never opened or saved.
        include_defined_names: If True (default), write `named_ranges` and
            `named_range_ranges` as workbook-global defined names. Set False
            to omit the name table.
        shared_formulas: Group interned autofill runs into Excel shared
            formulas (`t="shared"`). `auto` (default) groups when
            `graph.formula_shapes` is warm and skips grouping when it is
            missing. `off` always writes per-cell A1. `require` fails if the
            overlay is missing. The writer does not auto-rewarm (GitHub
            #560). Stale shapes, non-contiguous or mixed-axis leftovers,
            array formulas, and `INDIRECT` emit per-cell rather than an
            invalid shared formula.

    Raises:
        FileExistsError: If `destination` exists and `overwrite` is False.
        ValueError: If the view is empty, spans multiple sheets without
            `sheet_order`, a formula cell has no `formula_ast`,
            `formula_style` is `R1C1`, a relative axis has no host
            address, an array formula is missing its observed spill / CSE
            `ref`, a defined name cannot be expressed as a cell or
            rectangle, `shared_formulas` is not a known mode, or
            `shared_formulas='require'` and `formula_shapes` is missing.
    """
    dest = Path(destination)
    style = FormulaStyle(formula_style)
    shared_mode = parse_shared_formulas_mode(shared_formulas)
    if style is FormulaStyle.R1C1:
        raise ValueError(
            "R1C1 formula style is not persisted for normal formula cells; "
            "shared-formula groups store the master in A1 (Excel's xlsx dialect)"
        )
    if dest.exists() and not overwrite:
        raise FileExistsError(f"Refusing to overwrite existing file: {dest}")
    if dest.exists() and dest.is_dir():
        raise IsADirectoryError(f"Destination is a directory: {dest}")
    if len(graph) == 0:
        raise ValueError("Cannot write an empty graph view")

    sheet_names = _ordered_sheet_names(graph)
    planned = _plan_cells(
        graph,
        style=style,
        coerce_relative_refs=coerce_relative_refs,
        shared_formulas=shared_mode,
    )
    planned_names = _plan_defined_names(graph) if include_defined_names else []

    wb = Workbook()
    tmp: Path | None = None
    try:
        sheets = _create_sheets(wb, sheet_names)
        for sheet_name, coord, value in planned:
            sheets[sheet_name][coord] = value
        for name, attr_text in planned_names:
            wb.defined_names.add(DefinedName(name=name, attr_text=attr_text))
        tmp = dest.with_name(f".{dest.name}.{os.getpid()}.tmp")
        wb.save(tmp)
        os.replace(tmp, dest)
        tmp = None
    finally:
        wb.close()
        if tmp is not None:
            tmp.unlink(missing_ok=True)


def _cell_label(node: NodeView, fallback: str) -> str:
    if node.address is not None:
        return str(node.address)
    if node.sheet and node.column and node.row is not None:
        return f"{node.sheet}!{node.column}{node.row}"
    return fallback


def _ordered_sheet_names(graph: GraphReadView) -> list[str]:
    present: set[str] = set()
    for key in graph:
        node = graph.get_node(key)
        if node is None:
            continue
        if not node.sheet:
            raise ValueError(f"Cannot write cell {key} without a sheet name")
        present.add(node.sheet)
    if not present:
        raise ValueError("Cannot write an empty graph view")

    order = graph.sheet_order
    if order:
        seen: set[str] = set()
        names: list[str] = []
        for name in order:
            if name in present and name not in seen:
                names.append(name)
                seen.add(name)
        names.extend(sorted(present - seen))
        return names
    if len(present) > 1:
        raise ValueError("sheet_order is required when writing a view that spans multiple sheets")
    return list(present)


def _create_sheets(wb: Workbook, sheet_names: list[str]) -> dict[str, Worksheet]:
    first, *rest = sheet_names
    active = wb.active
    if active is None:
        active = wb.create_sheet(first)
    else:
        active.title = first
    sheets: dict[str, Worksheet] = {first: active}
    for name in rest:
        sheets[name] = wb.create_sheet(name)
    return sheets


def _plan_cells(
    graph: GraphReadView,
    *,
    style: FormulaStyle,
    coerce_relative_refs: bool,
    shared_formulas: SharedFormulasMode,
) -> list[tuple[str, str, object]]:
    shared_values = shared_formula_cell_values(
        graph,
        style=style,
        coerce_relative_refs=coerce_relative_refs,
        mode=shared_formulas,
    )
    planned: list[tuple[str, str, object]] = []
    for key in graph:
        node = graph.get_node(key)
        if node is None:
            continue
        if not node.sheet or not node.column or node.row is None:
            raise ValueError(f"Cannot write cell {key} without sheet/column/row")
        coord = f"{node.column}{node.row}"
        value = shared_values.get(key)
        if value is None:
            value = _cell_value(
                node,
                key=key,
                style=style,
                coerce_relative_refs=coerce_relative_refs,
            )
        planned.append((node.sheet, coord, value))
    return planned


def _cell_value(
    node: NodeView,
    *,
    key: str,
    style: FormulaStyle,
    coerce_relative_refs: bool,
) -> object:
    if node.has_formula:
        if node.formula_ast is None:
            raise ValueError(f"Cannot write unparseable formula at {_cell_label(node, key)}")
        try:
            text = render_formula(
                node.formula_ast,
                anchor=node.address,
                style=style,
                coerce_relative_refs=coerce_relative_refs,
            )
        except ValueError as exc:
            raise ValueError(f"Cannot render formula at {_cell_label(node, key)}: {exc}") from exc
        if node.is_array_formula:
            ref = node.array_formula_ref
            if not ref:
                raise ValueError(
                    f"Cannot write array formula at {_cell_label(node, key)} "
                    "without an observed spill/CSE ref"
                )
            return ArrayFormula(ref, text)
        return text
    if node.is_array_formula:
        raise ValueError(
            f"Cannot write array formula at {_cell_label(node, key)} without formula_ast"
        )
    return _excel_leaf_value(node.value, key=_cell_label(node, key))


def _excel_leaf_value(value: object, *, key: str) -> object:
    if value is None:
        return None
    if isinstance(value, XlError):
        return str(value)
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float, str, datetime, date)):
        return value
    raise ValueError(f"Unsupported leaf value type {type(value).__name__} at {key}")


def _plan_defined_names(graph: GraphReadView) -> list[tuple[str, str]]:
    cells = dict(graph.named_ranges or {})
    ranges = dict(graph.named_range_ranges or {})
    overlap = sorted(set(cells) & set(ranges))
    if overlap:
        shown = ", ".join(overlap)
        raise ValueError(
            f"Defined name(s) {shown} appear in both named_ranges and named_range_ranges"
        )
    planned: list[tuple[str, str]] = []
    for name, dest in cells.items():
        sheet, coord = dest
        planned.append((name, _defined_name_cell_attr(name, sheet, coord)))
    for name, dest in ranges.items():
        sheet, start, end = dest
        planned.append((name, _defined_name_range_attr(name, sheet, start, end)))
    return planned


def _require_defined_name(name: object) -> str:
    if not isinstance(name, str) or not name.strip():
        raise ValueError(f"Cannot write defined name {name!r}: name is empty")
    return name


def _require_defined_name_sheet(name: str, sheet: object) -> str:
    if not isinstance(sheet, str) or not sheet.strip():
        raise ValueError(
            f"Cannot write defined name {name!r}: expected a sheet-qualified cell or range"
        )
    return sheet


def _absolute_a1_coord(coord: object, *, name: str) -> str:
    if not isinstance(coord, str) or not coord.strip():
        raise ValueError(f"Cannot write defined name {name!r}: expected an A1 cell, got {coord!r}")
    try:
        col, row = coordinate_from_string(coord.replace("$", ""))
    except (CellCoordinatesException, IndexError, TypeError, ValueError) as exc:
        raise ValueError(
            f"Cannot write defined name {name!r}: expected an A1 cell, got {coord!r}"
        ) from exc
    if not col or not isinstance(row, int) or row < 1:
        raise ValueError(f"Cannot write defined name {name!r}: expected an A1 cell, got {coord!r}")
    return f"${col}${row}"


def _defined_name_cell_attr(name: object, sheet: object, coord: object) -> str:
    label = _require_defined_name(name)
    sheet_name = _require_defined_name_sheet(label, sheet)
    return f"{quote_sheet_if_needed(sheet_name)}!{_absolute_a1_coord(coord, name=label)}"


def _defined_name_range_attr(name: object, sheet: object, start: object, end: object) -> str:
    label = _require_defined_name(name)
    sheet_name = _require_defined_name_sheet(label, sheet)
    start_a1 = _absolute_a1_coord(start, name=label)
    end_a1 = _absolute_a1_coord(end, name=label)
    return f"{quote_sheet_if_needed(sheet_name)}!{start_a1}:{end_a1}"
