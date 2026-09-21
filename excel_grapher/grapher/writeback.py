"""Write a `GraphReadView` to a new Excel workbook.

This is workbook I/O next to `create_dependency_graph`, not Python codegen.
`write_workbook` is an export backend on the same read surface `CodeGenerator`
already accepts (`DependencyGraph` or `ProjectionResult`).
"""

from __future__ import annotations

import os
from datetime import date, datetime
from pathlib import Path
from typing import TYPE_CHECKING

from fastpyxl import Workbook
from fastpyxl.utils.cell import coordinate_from_string
from fastpyxl.utils.exceptions import CellCoordinatesException
from fastpyxl.workbook.defined_name import DefinedName
from fastpyxl.worksheet.formula import ArrayFormula
from fastpyxl.worksheet.worksheet import Worksheet

from excel_grapher.core.address_keys import format_key, parse_address, quote_sheet_if_needed
from excel_grapher.core.excel_function_names import normalize_excel_function_name
from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    CellRef,
    CellRefNode,
    FormulaStyle,
    FunctionCallNode,
    RangeNode,
    UnaryOpNode,
    WholeColumnNode,
    WholeRowNode,
    bind_axes,
    parse_preserving_axes_optional,
    render_formula,
    resolve_cell_ref,
)
from excel_grapher.core.types import XlError

from .graph import GraphReadView
from .node import NodeView
from .parser import DEFAULT_MAX_RANGE_CELLS, expand_range
from .shared_formulas import (
    SharedFormulasMode,
    parse_shared_formulas_mode,
    shared_formula_cell_values,
)

_DYNAMIC_OVERLAY_FNS = frozenset({"OFFSET", "INDIRECT"})

if TYPE_CHECKING:
    from excel_grapher.series_bindings.types import WorkbookSeriesBindings


def write_workbook(
    graph: GraphReadView,
    destination: Path | str,
    *,
    formula_style: FormulaStyle = FormulaStyle.A1_EXCEL,
    coerce_relative_refs: bool = False,
    overwrite: bool = False,
    include_defined_names: bool = True,
    shared_formulas: SharedFormulasMode = "auto",
    series_bindings: WorkbookSeriesBindings | None = None,
    bindings_workbook: Path | str | None = None,
    include_bound_labels: bool = True,
    include_cell_validation: bool = True,
) -> None:
    """Write `graph` to a new `.xlsx` at `destination`.

    Accepts any `GraphReadView`: a mutable `DependencyGraph` (the workbook
    after `set_node_ast` / `set_node_value` / `move_node` / in-place
    `compress_*`) or a non-mutating `ProjectionResult`. Formula rewriting
    is an intended write-back use case; compressed or projected views are
    written as they are. Output contains only sheets and cells from the
    view. Styles, charts, VBA, and cells outside the view are omitted
    (accepted v1 lossiness). Vacated `move_node` addresses are simply
    absent; there is no template to clear. When `series_bindings` is set
    and `include_bound_labels` is True (default), `row_label`,
    `column_header`, `kind: cell`, and attribute source cells that
    bindings read are written too, even if they sit outside the target
    closure. Formula labels are copied only when every static dependency
    is already in the view or is itself a bound label; otherwise the
    writer fails closed rather than extracting a second dependency
    closure. The bindings workbook is opened read-only for those extra
    cells; it is never overwritten. When `include_cell_validation` is
    True (default), constrained inputs are also written as Excel data
    validations: input-series domains and relations when
    `series_bindings` is set, otherwise value cells in `cell_type_env`.
    Constants, outputs, and formula cells are omitted.

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
            The destination is always a new file; `graph`'s original
            workbook is never saved in place. `bindings_workbook` is opened
            read-only when bound labels are included (`include_bound_labels`
            is True).
        include_defined_names: If True (default), write `named_ranges` and
            `named_range_ranges` as workbook-global defined names. Set False
            to omit the name table.
        shared_formulas: Group interned autofill runs into Excel shared
            formulas (`t="shared"`). `auto` (default) groups when
            `graph.formula_shapes` is warm and skips grouping when it is
            missing. `off` always writes per-cell A1. `require` fails if the
            overlay is missing. The writer does not auto-rewarm (GitHub
            #560). `project()` drops the overlay on the clone; rewarm
            `projection.projected_graph.formula_shapes` before writing a
            `ProjectionResult` if you want grouping. Stale shapes,
            non-contiguous or mixed-axis leftovers, array formulas, and
            `INDIRECT` emit per-cell rather than an invalid shared formula.
        series_bindings: Optional sidecar used to locate bound labels.
            Requires `bindings_workbook` when `include_bound_labels` is True.
        bindings_workbook: Workbook the sidecar describes. Required when
            `series_bindings` is set and `include_bound_labels` is True.
        include_bound_labels: If True (default) and `series_bindings` is
            set, write `row_label` / `column_header` / `kind: cell` /
            attribute source cells that are missing from `graph`. Formula
            labels whose static deps are not already in the view or bound
            labels are refused. The original graph is not mutated.
        include_cell_validation: If True (default), write Excel data
            validations for constrained inputs. With `series_bindings`,
            rules come from input-series `enum`, `between`,
            `real_between`, `from_workbook`, `value_map` needles, and
            `relations`. Constants, outputs, and internals are skipped.
            Without bindings, value cells in `cell_type_env` (including
            a projection's projected graph) are used. Inline lists cannot
            contain commas. Formulas longer than 255 characters, relation
            partners missing from the written view, and `from_workbook`
            inputs with no cached value fail closed.

    Raises:
        FileExistsError: If `destination` exists and `overwrite` is False.
        ValueError: If the view is empty, spans multiple sheets without
            `sheet_order`, a formula cell has no `formula_ast`,
            `formula_style` is `R1C1`, a relative axis has no host
            address, an array formula is missing its observed spill / CSE
            `ref`, a defined name cannot be expressed as a cell or
            rectangle, `shared_formulas` is not a known mode,
            `shared_formulas='require'` and `formula_shapes` is missing, or
            `series_bindings` is set with `include_bound_labels` and without
            `bindings_workbook`, a bound formula label cannot be copied
            as a static overlay, or a constrained input cannot be written
            as Excel data validation.
    """
    dest = Path(destination)
    style = FormulaStyle(formula_style)
    shared_mode = parse_shared_formulas_mode(shared_formulas)
    if series_bindings is not None and include_bound_labels and bindings_workbook is None:
        raise ValueError(
            "bindings_workbook is required when series_bindings is set "
            "and include_bound_labels is True"
        )
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

    planned = _plan_cells(
        graph,
        style=style,
        coerce_relative_refs=coerce_relative_refs,
        shared_formulas=shared_mode,
    )
    if series_bindings is not None and include_bound_labels:
        assert bindings_workbook is not None
        extra_planned = _plan_bound_label_cells(
            graph,
            series_bindings=series_bindings,
            bindings_workbook=bindings_workbook,
            style=style,
            coerce_relative_refs=coerce_relative_refs,
        )
        planned = _merge_planned_cells(extra_planned, planned)
    sheet_names = _ordered_sheet_names(
        graph,
        present={sheet for sheet, _coord, _value in planned},
    )
    planned_names = _plan_defined_names(graph) if include_defined_names else []

    wb = Workbook()
    tmp: Path | None = None
    try:
        sheets = _create_sheets(wb, sheet_names)
        for sheet_name, coord, value in planned:
            sheets[sheet_name][coord] = value
        if include_cell_validation:
            from .cell_validation import apply_constrained_input_validations

            apply_constrained_input_validations(
                sheets,
                graph,
                planned=planned,
                series_bindings=series_bindings,
                bindings_workbook=bindings_workbook,
            )
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


def _ordered_sheet_names(
    graph: GraphReadView,
    present: set[str] | None = None,
) -> list[str]:
    if present is None:
        present = set()
        for key in graph:
            node = graph.get_node(key)
            if node is None:
                continue
            if not node.sheet:
                raise ValueError(f"Cannot write cell {key} without a sheet name")
            present.add(node.sheet)
    if not present:
        raise ValueError("Cannot write an empty graph view")

    order = list(graph.sheet_order) if graph.sheet_order else []
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


def _merge_planned_cells(
    base: list[tuple[str, str, object]],
    overlay: list[tuple[str, str, object]],
) -> list[tuple[str, str, object]]:
    merged: dict[tuple[str, str], object] = {(sheet, coord): value for sheet, coord, value in base}
    for sheet, coord, value in overlay:
        merged[(sheet, coord)] = value
    return [(sheet, coord, value) for (sheet, coord), value in merged.items()]


def _plan_bound_label_cells(
    graph: GraphReadView,
    *,
    series_bindings: WorkbookSeriesBindings,
    bindings_workbook: Path | str,
    style: FormulaStyle,
    coerce_relative_refs: bool,
) -> list[tuple[str, str, object]]:
    from excel_grapher.series_bindings.resolve import _WorkbookValues, bound_label_addresses

    labels = bound_label_addresses(
        series_bindings,
        workbook=bindings_workbook,
        graph=graph,
    )
    missing = [address for address in sorted(labels) if address not in graph]
    if not missing:
        return []
    planned: list[tuple[str, str, object]] = []
    with _WorkbookValues(bindings_workbook, data_only=False) as reader:
        reader.prefetch(missing)
        for address in missing:
            sheet, coord = parse_address(address)
            value = _bound_label_cell_value(
                reader.read(address),
                address=address,
                graph=graph,
                labels=labels,
                style=style,
                coerce_relative_refs=coerce_relative_refs,
            )
            planned.append((sheet, coord, value))
    return planned


def _bound_label_cell_value(
    raw: object,
    *,
    address: str,
    graph: GraphReadView,
    labels: set[str],
    style: FormulaStyle,
    coerce_relative_refs: bool,
) -> object:
    if isinstance(raw, ArrayFormula):
        raise ValueError(f"Cannot write bound array formula at {address}")
    if isinstance(raw, str) and raw.startswith("="):
        return _bound_label_formula_text(
            raw,
            address=address,
            graph=graph,
            labels=labels,
            style=style,
            coerce_relative_refs=coerce_relative_refs,
        )
    return _excel_leaf_value(raw, key=address)


def _bound_label_formula_text(
    formula: str,
    *,
    address: str,
    graph: GraphReadView,
    labels: set[str],
    style: FormulaStyle,
    coerce_relative_refs: bool,
) -> str:
    ast = parse_preserving_axes_optional(
        formula,
        anchor=address,
        named_ranges=graph.named_ranges,
        named_range_ranges=graph.named_range_ranges,
    )
    if ast is None:
        raise ValueError(f"Cannot write unparseable bound label formula at {address}")
    deps = _overlay_formula_deps(bind_axes(ast, address), address=address)
    missing = sorted(dep for dep in deps if dep not in graph and dep not in labels)
    if missing:
        shown = ", ".join(missing)
        noun = "dependency" if len(missing) == 1 else "dependencies"
        verb = "is" if len(missing) == 1 else "are"
        raise ValueError(
            f"Cannot write bound label formula at {address}: "
            f"{noun} {shown} {verb} not in the graph view or bound labels"
        )
    try:
        return render_formula(
            ast,
            anchor=address,
            style=style,
            coerce_relative_refs=coerce_relative_refs,
        )
    except ValueError as exc:
        raise ValueError(f"Cannot render bound label formula at {address}: {exc}") from exc


def _overlay_formula_deps(ast: AstNode, *, address: str) -> set[str]:
    deps: set[str] = set()

    def walk(node: AstNode) -> None:
        match node:
            case CellRefNode(ref):
                deps.add(resolve_cell_ref(ref, address))
            case RangeNode(start_ref, end_ref):
                deps.update(_overlay_range_deps(start_ref, end_ref, address=address))
            case WholeColumnNode() | WholeRowNode():
                raise ValueError(
                    f"Cannot write bound label formula at {address}: "
                    "whole-column/row references cannot be copied as a static overlay"
                )
            case FunctionCallNode(name, args):
                if normalize_excel_function_name(name) in _DYNAMIC_OVERLAY_FNS:
                    raise ValueError(
                        f"Cannot write bound label formula at {address}: "
                        "OFFSET/INDIRECT cannot be copied as a static overlay"
                    )
                for arg in args:
                    walk(arg)
            case BinaryOpNode(_, left, right):
                walk(left)
                walk(right)
            case UnaryOpNode(_, operand):
                walk(operand)
            case _:
                return

    walk(ast)
    return deps


def _overlay_range_deps(start_ref: CellRef, end_ref: CellRef, *, address: str) -> set[str]:
    start = resolve_cell_ref(start_ref, address)
    end = resolve_cell_ref(end_ref, address)
    start_sheet, start_coord = parse_address(start)
    end_sheet, end_coord = parse_address(end)
    if start_sheet != end_sheet:
        raise ValueError(
            f"Cannot write bound label formula at {address}: "
            "cross-sheet ranges cannot be copied as a static overlay"
        )
    start_col, start_row = coordinate_from_string(start_coord)
    end_col, end_row = coordinate_from_string(end_coord)
    try:
        pairs = expand_range(
            sheet=start_sheet,
            start_col=start_col,
            start_row=int(start_row),
            end_col=end_col,
            end_row=int(end_row),
            max_cells=DEFAULT_MAX_RANGE_CELLS,
        )
    except ValueError as exc:
        raise ValueError(f"Cannot write bound label formula at {address}: {exc}") from exc
    return {format_key(sheet, coord) for sheet, coord in pairs}


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
