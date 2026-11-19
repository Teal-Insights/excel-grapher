"""Write a `GraphReadView` to a new or patched Excel workbook.

This is workbook I/O next to `create_dependency_graph`, not Python codegen.
`write_workbook` is an export backend on the same read surface `CodeGenerator`
already accepts (`DependencyGraph` or `ProjectionResult`), supporting both
clean-room new-file output and in-place / template-patched writeback.
"""

from __future__ import annotations

import os
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import TYPE_CHECKING

from fastpyxl import Workbook, load_workbook
from fastpyxl.utils.cell import (
    column_index_from_string,
    coordinate_from_string,
    get_column_letter,
)
from fastpyxl.utils.exceptions import CellCoordinatesException
from fastpyxl.workbook.defined_name import DefinedName
from fastpyxl.worksheet.formula import ArrayFormula
from fastpyxl.worksheet.worksheet import Worksheet

from excel_grapher.core.address_keys import (
    format_key,
    normalize_key,
    parse_address,
    quote_sheet_if_needed,
)
from excel_grapher.core.excel_function_names import normalize_excel_function_name
from excel_grapher.core.formula_ast import (
    AbsoluteAxis,
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
from .parser import (
    DEFAULT_MAX_RANGE_CELLS,
    FormulaNormalizer,
    _find_function_calls_with_spans,
    expand_range,
    mask_ref_only_function_calls,
    mask_spans,
    parse_range_refs_with_spans,
    parse_standalone_cell_refs,
)
from .resolver import build_named_range_map
from .shared_formulas import (
    SharedFormula,
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
    in_place: bool = False,
    template: Path | str | None = None,
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
    """Write `graph` to a new `.xlsx` or patch an existing workbook at `destination`.

    Accepts any `GraphReadView`: a mutable `DependencyGraph` (the workbook
    after `set_node_ast` / `set_node_value` / `replace_node_formula` /
    `move_node` / in-place `compress_*`) or a non-mutating `ProjectionResult`.

    Two output modes share this surface:

    - **Clean-room new file** (default, `in_place=False` and `template=None`):
      writes a brand-new `.xlsx` containing only sheets and cells from the
      view (plus optional bound labels). Styles, charts, VBA, and cells
      outside the view are omitted.
    - **In-place / template patch** (`in_place=True` or `template=...`):
      loads the existing `.xlsx` / `.xlsm` (`destination` when `in_place=True`,
      or `template` otherwise), preserves styles, number formats, merged
      cells, charts, VBA (`keep_vba=True` for `.xlsm`), off-closure sheets
      and cells, and sheet-local or formula-valued defined names. Cells
      originally extracted into `graph` (`graph.extracted_keys`) that are no
      longer present after `move_node`, `remove_node`, `compress_*`, or
      `ProjectionResult` have their values cleared while keeping cell borders
      and fills intact. Fails closed if a new or moved cell would overwrite an
      occupied off-closure cell or array block, or if vacating a cell leaves a
      dangling reference in an off-closure formula or defined name. Sets
      `wb.calculation.fullCalcOnLoad = True` so desktop Excel recalculates
      downstream formulas on open.

    Formula cells are spelled from current `formula_ast` via `render_formula`
    (never from the opt-in raw `Node.formula` audit string). Defined names are
    expanded to A1 before parse, so written formulas are expanded A1 (for
    example `=A1*Settings!$B$2`, not `=A1*TaxRate`). When
    `include_defined_names` is True (default), workbook-global names from
    `graph.named_ranges` and `graph.named_range_ranges` are written (upserted
    in patch mode so existing unrelated defined names remain).

    Args:
        graph: Read view whose cells are written.
        destination: Output path. Parent directories must already exist.
        in_place: If True, load `destination` as the base workbook and
            atomically rewrite it in place (implies overwriting `destination`).
        template: Optional path to an existing `.xlsx` / `.xlsm` workbook to
            copy and patch into `destination`. Cannot point to a different file
            than `destination` when `in_place` is True.
        formula_style: Reference spelling. Default `A1_EXCEL` keeps `$` on
            absolute axes and omits the host sheet prefix. `R1C1` is not
            persisted for normal formula cells.
        coerce_relative_refs: If True, bind relative axes to absolute
            indexes before spelling (`$` on both axes in `A1_EXCEL`).
        overwrite: If False (default) and `in_place` is False, raise when
            `destination` exists.
        include_defined_names: If True (default), write `named_ranges` and
            `named_range_ranges` as workbook-global defined names.
        shared_formulas: Group interned autofill runs into Excel shared
            formulas (`t="shared"`): `auto`, `off`, or `require`.
        series_bindings: Optional sidecar used to locate bound labels and
            input domains for cell validation.
        bindings_workbook: Workbook the sidecar describes. Defaults to the
            base workbook when `in_place=True` or `template` is set; required
            in clean-room mode when bound labels or input relations are used.
        include_bound_labels: If True (default) and `series_bindings` is
            set, include `row_label` / `column_header` / `kind: cell` /
            attribute source cells missing from `graph`.
        include_cell_validation: If True (default), write Excel data
            validations for constrained inputs. In patch mode, existing
            validation ranges on rewritten or vacated cells are split and
            replaced without duplicating rules.

    Raises:
        FileNotFoundError: If `in_place=True` and `destination` does not exist,
            or `template` does not exist.
        FileExistsError: If `destination` exists, `in_place` is False, and
            `overwrite` is False.
        IsADirectoryError: If `destination` or `template` is a directory.
        ValueError: On invalid options, unrenderable formulas, macro extension
            mismatch (`.xlsm` vs `.xlsx`), off-closure overwrite collision, or
            dangling off-closure references to vacated cells.
    """
    dest = Path(destination)
    style = FormulaStyle(formula_style)
    shared_mode = parse_shared_formulas_mode(shared_formulas)
    base_path = _resolve_base_workbook(
        dest, in_place=in_place, template=template, overwrite=overwrite
    )
    effective_bindings_wb: Path | str | None = (
        bindings_workbook if bindings_workbook is not None else base_path
    )

    if series_bindings is not None and include_bound_labels and effective_bindings_wb is None:
        raise ValueError(
            "bindings_workbook is required when series_bindings is set "
            "and include_bound_labels is True"
        )
    if style is FormulaStyle.R1C1:
        raise ValueError(
            "R1C1 formula style is not persisted for normal formula cells; "
            "shared-formula groups store the master in A1 (Excel's xlsx dialect)"
        )
    if len(graph) == 0:
        raise ValueError("Cannot write an empty graph view")

    planned = _plan_cells(
        graph,
        style=style,
        coerce_relative_refs=coerce_relative_refs,
        shared_formulas=shared_mode,
    )
    if series_bindings is not None and include_bound_labels:
        assert effective_bindings_wb is not None
        same_as_base = (
            base_path is not None and Path(effective_bindings_wb).resolve() == base_path.resolve()
        )
        if not same_as_base:
            extra_planned = _plan_bound_label_cells(
                graph,
                series_bindings=series_bindings,
                bindings_workbook=effective_bindings_wb,
                style=style,
                coerce_relative_refs=coerce_relative_refs,
            )
            planned = _merge_planned_cells(extra_planned, planned)
    planned_names = _plan_defined_names(graph) if include_defined_names else []

    if base_path is None:
        _write_clean_workbook(
            graph,
            dest,
            planned=planned,
            planned_names=planned_names,
            include_cell_validation=include_cell_validation,
            series_bindings=series_bindings,
            bindings_workbook=effective_bindings_wb,
        )
        return

    _write_patched_workbook(
        graph,
        dest,
        base_path=base_path,
        planned=planned,
        planned_names=planned_names,
        include_defined_names=include_defined_names,
        include_cell_validation=include_cell_validation,
        series_bindings=series_bindings,
        bindings_workbook=effective_bindings_wb,
    )


def _resolve_base_workbook(
    dest: Path,
    *,
    in_place: bool,
    template: Path | str | None,
    overwrite: bool,
) -> Path | None:
    if dest.exists() and dest.is_dir():
        raise IsADirectoryError(f"Destination is a directory: {dest}")

    base_path: Path | None = None
    if in_place:
        if template is not None and Path(template).resolve() != dest.resolve():
            raise ValueError(
                "Cannot combine in_place=True with a different template path; "
                "pass template=... without in_place=True to write to another file"
            )
        if not dest.exists():
            raise FileNotFoundError(f"Cannot patch missing workbook in place: {dest}")
        base_path = dest
    elif template is not None:
        base_path = Path(template)
        if base_path.exists() and base_path.is_dir():
            raise IsADirectoryError(f"Template is a directory: {base_path}")
        if not base_path.exists():
            raise FileNotFoundError(f"Template workbook not found: {base_path}")
        if dest.exists() and not overwrite:
            raise FileExistsError(f"Refusing to overwrite existing file: {dest}")
    else:
        if dest.exists() and not overwrite:
            raise FileExistsError(f"Refusing to overwrite existing file: {dest}")

    if base_path is not None:
        base_xlsm = base_path.suffix.lower() == ".xlsm"
        dest_xlsm = dest.suffix.lower() == ".xlsm"
        if base_xlsm != dest_xlsm:
            raise ValueError(
                f"Template extension ({base_path.suffix!r}) and destination extension "
                f"({dest.suffix!r}) must agree on macro support (.xlsm)"
            )
    return base_path


def _write_clean_workbook(
    graph: GraphReadView,
    dest: Path,
    *,
    planned: list[tuple[str, str, object]],
    planned_names: list[tuple[str, str]],
    include_cell_validation: bool,
    series_bindings: WorkbookSeriesBindings | None,
    bindings_workbook: Path | str | None,
) -> None:
    sheet_names = _ordered_sheet_names(
        graph,
        present={sheet for sheet, _coord, _value in planned},
    )
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


@dataclass(frozen=True, slots=True)
class _ArrayBlock:
    sheet: str
    anchor_coord: str
    ref: str
    min_col: int
    min_row: int
    max_col: int
    max_row: int

    def contains(self, col_idx: int, row_idx: int) -> bool:
        return self.min_col <= col_idx <= self.max_col and self.min_row <= row_idx <= self.max_row


def _is_true_array_formula(value: object) -> bool:
    return isinstance(value, ArrayFormula) and not isinstance(value, SharedFormula)


def _parse_a1_rect_bounds(ref: str) -> tuple[int, int, int, int]:
    cleaned = ref.replace("$", "").strip()
    start_str, sep, end_str = cleaned.partition(":")
    c1_str, r1 = coordinate_from_string(start_str)
    if not sep:
        col_i = int(column_index_from_string(c1_str))
        return col_i, int(r1), col_i, int(r1)
    c2_str, r2 = coordinate_from_string(end_str)
    col_a = int(column_index_from_string(c1_str))
    col_b = int(column_index_from_string(c2_str))
    return min(col_a, col_b), min(int(r1), int(r2)), max(col_a, col_b), max(int(r1), int(r2))


def _expand_rect_keys(
    sheet: str, min_col: int, min_row: int, max_col: int, max_row: int
) -> set[str]:
    out: set[str] = set()
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            out.add(format_key(sheet, f"{get_column_letter(col)}{row}"))
    return out


def _write_patched_workbook(
    graph: GraphReadView,
    dest: Path,
    *,
    base_path: Path,
    planned: list[tuple[str, str, object]],
    planned_names: list[tuple[str, str]],
    include_defined_names: bool,
    include_cell_validation: bool,
    series_bindings: WorkbookSeriesBindings | None,
    bindings_workbook: Path | str | None,
) -> None:
    from .cell_validation import apply_constrained_input_validations, strip_validation_coords

    keep_vba = base_path.suffix.lower() == ".xlsm"
    wb = load_workbook(base_path, data_only=False, keep_vba=keep_vba)
    tmp: Path | None = None
    try:
        raw_extracted = getattr(graph, "extracted_keys", None)
        current_keys = {normalize_key(k) for k in graph}
        baseline_keys = (
            {normalize_key(k) for k in raw_extracted}
            if raw_extracted is not None
            else set(current_keys)
        )

        baseline_owned_keys = set(baseline_keys)
        off_closure_arrays: dict[str, list[_ArrayBlock]] = defaultdict(list)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for anchor_coord, ref_str in ws.array_formulae.items():
                c1, r1, c2, r2 = _parse_a1_rect_bounds(str(ref_str))
                anchor_key = format_key(sheet_name, anchor_coord)
                if anchor_key in baseline_keys:
                    baseline_owned_keys.update(_expand_rect_keys(sheet_name, c1, r1, c2, r2))
                else:
                    off_closure_arrays[sheet_name].append(
                        _ArrayBlock(
                            sheet=sheet_name,
                            anchor_coord=anchor_coord,
                            ref=str(ref_str),
                            min_col=c1,
                            min_row=r1,
                            max_col=c2,
                            max_row=r2,
                        )
                    )

        planned_keys = {format_key(sheet, coord) for sheet, coord, _ in planned}
        covered_by_planned_arrays: set[str] = set()

        for sheet_name, coord, value in planned:
            cell_key = format_key(sheet_name, coord)
            col_str, row_i = coordinate_from_string(coord)
            col_i = int(column_index_from_string(col_str))
            for block in off_closure_arrays.get(sheet_name, ()):
                if block.contains(col_i, int(row_i)):
                    raise ValueError(
                        f"Cannot write {cell_key} during in-place writeback: overlaps "
                        f"off-graph array formula at {format_key(sheet_name, block.anchor_coord)} "
                        f"({block.ref})"
                    )
            if cell_key not in baseline_owned_keys and sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                if (int(row_i), col_i) in ws._cells and ws[coord].value is not None:
                    raise ValueError(
                        f"Cannot write {cell_key} during in-place writeback: "
                        "off-graph cell is already occupied"
                    )
            if _is_true_array_formula(value):
                assert isinstance(value, ArrayFormula)
                if value.ref:
                    ac1, ar1, ac2, ar2 = _parse_a1_rect_bounds(str(value.ref))
                    covered_by_planned_arrays.update(
                        _expand_rect_keys(sheet_name, ac1, ar1, ac2, ar2)
                    )
                    if sheet_name in wb.sheetnames:
                        ws = wb[sheet_name]
                        for r_idx in range(ar1, ar2 + 1):
                            for c_idx in range(ac1, ac2 + 1):
                                for block in off_closure_arrays.get(sheet_name, ()):
                                    if block.contains(c_idx, r_idx):
                                        raise ValueError(
                                            f"Cannot write array formula at {cell_key} ({value.ref}) "
                                            "during in-place writeback: overlaps off-graph array "
                                            f"formula at {format_key(sheet_name, block.anchor_coord)} "
                                            f"({block.ref})"
                                        )
                                member_coord = f"{get_column_letter(c_idx)}{r_idx}"
                                member_key = format_key(sheet_name, member_coord)
                                if (
                                    member_key not in baseline_owned_keys
                                    and (r_idx, c_idx) in ws._cells
                                    and ws[member_coord].value is not None
                                ):
                                    raise ValueError(
                                        f"Cannot write array formula at {cell_key} ({value.ref}) "
                                        f"during in-place writeback: spill cell {member_key} "
                                        "is occupied by an off-graph value"
                                    )

        dangling_candidates = baseline_owned_keys - planned_keys - covered_by_planned_arrays
        if dangling_candidates:
            _validate_no_dangling_off_closure_refs(
                wb,
                graph=graph,
                vacated_keys=dangling_candidates,
                baseline_owned_keys=baseline_owned_keys,
                planned_keys=planned_keys,
                planned_names=planned_names if include_defined_names else [],
                include_defined_names=include_defined_names,
            )

        for sheet_name, _coord, _value in planned:
            if sheet_name not in wb.sheetnames:
                wb.create_sheet(sheet_name)
        sheets: dict[str, Worksheet] = {name: wb[name] for name in wb.sheetnames}

        cells_to_clear = baseline_owned_keys - planned_keys
        for clear_key in sorted(cells_to_clear):
            c_sheet, c_coord = parse_address(clear_key)
            if c_sheet not in sheets:
                continue
            ws = sheets[c_sheet]
            col_str, row_i = coordinate_from_string(c_coord)
            col_i = int(column_index_from_string(col_str))
            if (int(row_i), col_i) in ws._cells or c_coord in ws.array_formulae:
                ws[c_coord] = None

        for sheet_name, coord, value in planned:
            sheets[sheet_name][coord] = value

        validation_strip_keys = (
            baseline_owned_keys | planned_keys if include_cell_validation else cells_to_clear
        )
        strip_by_sheet: dict[str, set[str]] = defaultdict(set)
        for vkey in validation_strip_keys:
            v_sheet, v_coord = parse_address(vkey)
            strip_by_sheet[v_sheet].add(v_coord)
        strip_validation_coords(sheets, strip_by_sheet)

        if include_cell_validation:
            apply_constrained_input_validations(
                sheets,
                graph,
                planned=planned,
                series_bindings=series_bindings,
                bindings_workbook=bindings_workbook,
            )

        if include_defined_names:
            for name, attr_text in planned_names:
                wb.defined_names.add(DefinedName(name=name, attr_text=attr_text))

        wb.calculation.fullCalcOnLoad = True
        tmp = dest.with_name(f".{dest.name}.{os.getpid()}.tmp")
        wb.save(tmp)
        os.replace(tmp, dest)
        tmp = None
    finally:
        vba_archive = getattr(wb, "vba_archive", None)
        if (
            vba_archive is not None
            and getattr(vba_archive, "fp", None) is not None
            and not vba_archive.fp.closed
        ):
            vba_archive.close()
        wb.close()
        if tmp is not None:
            tmp.unlink(missing_ok=True)


def _validate_no_dangling_off_closure_refs(
    wb: Workbook,
    *,
    graph: GraphReadView,
    vacated_keys: set[str],
    baseline_owned_keys: set[str],
    planned_keys: set[str],
    planned_names: list[tuple[str, str]],
    include_defined_names: bool,
) -> None:
    vacated_by_sheet: dict[str, list[tuple[int, int, str]]] = defaultdict(list)
    for vkey in sorted(vacated_keys):
        v_sheet, v_coord = parse_address(vkey)
        c_str, r_int = coordinate_from_string(v_coord)
        vacated_by_sheet[v_sheet].append((int(column_index_from_string(c_str)), int(r_int), vkey))

    wb_maps = build_named_range_map(wb)
    cell_names = dict(wb_maps.cell_map)
    range_names = dict(wb_maps.range_map)
    if include_defined_names:
        for name, dest_cell in (graph.named_ranges or {}).items():
            cell_names[name] = dest_cell
            range_names.pop(name, None)
        for name, dest_range in (graph.named_range_ranges or {}).items():
            range_names[name] = dest_range
            cell_names.pop(name, None)

    effective_global_names: dict[str, str] = {}
    for name, defn in wb.defined_names.items():
        attr_text = getattr(defn, "attr_text", None)
        if isinstance(attr_text, str) and attr_text:
            effective_global_names[str(name)] = attr_text
    for name, attr_text in planned_names:
        effective_global_names[name] = attr_text

    default_sheet = wb.sheetnames[0] if wb.sheetnames else "Sheet1"
    for name, attr_text in sorted(effective_global_names.items()):
        if attr_text.startswith(("{", '"', "#")):
            continue
        _check_formula_against_vacated(
            f"={attr_text}",
            host_sheet=default_sheet,
            host_coord="A1",
            source_label=f"defined name {name!r}",
            vacated_by_sheet=vacated_by_sheet,
            sheetnames=list(wb.sheetnames),
            named_ranges=cell_names,
            named_range_ranges=range_names,
        )

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for name, defn in ws.defined_names.items():
            attr_text = getattr(defn, "attr_text", None)
            if (
                not isinstance(attr_text, str)
                or not attr_text
                or attr_text.startswith(("{", '"', "#"))
            ):
                continue
            _check_formula_against_vacated(
                f"={attr_text}",
                host_sheet=sheet_name,
                host_coord="A1",
                source_label=f"sheet defined name {sheet_name}!{name}",
                vacated_by_sheet=vacated_by_sheet,
                sheetnames=list(wb.sheetnames),
                named_ranges=cell_names,
                named_range_ranges=range_names,
            )

        for (row_i, col_i), cell in ws._cells.items():
            coord = f"{get_column_letter(col_i)}{row_i}"
            cell_key = format_key(sheet_name, coord)
            if cell_key in baseline_owned_keys or cell_key in planned_keys:
                continue
            raw_val = cell.value
            formula_text: str | None = None
            if isinstance(raw_val, ArrayFormula):
                raw_text = raw_val.text or ""
                if raw_text:
                    formula_text = raw_text if raw_text.startswith("=") else f"={raw_text}"
            elif isinstance(raw_val, str) and raw_val.startswith("="):
                formula_text = raw_val
            if formula_text is None:
                continue
            _check_formula_against_vacated(
                formula_text,
                host_sheet=sheet_name,
                host_coord=coord,
                source_label=f"off-graph cell {cell_key}",
                vacated_by_sheet=vacated_by_sheet,
                sheetnames=list(wb.sheetnames),
                named_ranges=cell_names,
                named_range_ranges=range_names,
            )


def _check_formula_against_vacated(
    formula_text: str,
    *,
    host_sheet: str,
    host_coord: str,
    source_label: str,
    vacated_by_sheet: dict[str, list[tuple[int, int, str]]],
    sheetnames: list[str],
    named_ranges: dict[str, tuple[str, str]],
    named_range_ranges: dict[str, tuple[str, str, str]],
) -> None:
    anchor = format_key(host_sheet, host_coord)
    ast = parse_preserving_axes_optional(
        formula_text,
        anchor=anchor,
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
    )
    if ast is not None:
        _walk_bound_ast_for_vacated(
            bind_axes(ast, anchor),
            source_label=source_label,
            vacated_by_sheet=vacated_by_sheet,
            sheetnames=sheetnames,
        )
        return

    _check_unparsed_formula_for_vacated(
        formula_text,
        host_sheet=host_sheet,
        source_label=source_label,
        vacated_by_sheet=vacated_by_sheet,
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
    )


def _sheets_in_span(start_sheet: str, end_sheet: str, sheetnames: list[str]) -> list[str]:
    if start_sheet == end_sheet:
        return [start_sheet]
    if start_sheet in sheetnames and end_sheet in sheetnames:
        i1 = sheetnames.index(start_sheet)
        i2 = sheetnames.index(end_sheet)
        lo, hi = min(i1, i2), max(i1, i2)
        return sheetnames[lo : hi + 1]
    return [start_sheet, end_sheet]


def _raise_if_rect_hits_vacated(
    sheets: list[str],
    *,
    c_lo: int,
    r_lo: int,
    c_hi: int,
    r_hi: int,
    source_label: str,
    vacated_by_sheet: dict[str, list[tuple[int, int, str]]],
) -> None:
    for sheet in sheets:
        for v_col, v_row, v_key in vacated_by_sheet.get(sheet, ()):
            if c_lo <= v_col <= c_hi and r_lo <= v_row <= r_hi:
                raise ValueError(
                    f"Cannot vacate {v_key} during in-place writeback: "
                    f"{source_label} still references it"
                )


def _walk_bound_ast_for_vacated(
    node: AstNode,
    *,
    source_label: str,
    vacated_by_sheet: dict[str, list[tuple[int, int, str]]],
    sheetnames: list[str],
) -> None:
    match node:
        case CellRefNode(ref):
            assert isinstance(ref.col, AbsoluteAxis) and isinstance(ref.row, AbsoluteAxis)
            _raise_if_rect_hits_vacated(
                [ref.sheet],
                c_lo=ref.col.index,
                r_lo=ref.row.index,
                c_hi=ref.col.index,
                r_hi=ref.row.index,
                source_label=source_label,
                vacated_by_sheet=vacated_by_sheet,
            )
        case RangeNode(start_ref, end_ref):
            assert isinstance(start_ref.col, AbsoluteAxis) and isinstance(
                start_ref.row, AbsoluteAxis
            )
            assert isinstance(end_ref.col, AbsoluteAxis) and isinstance(end_ref.row, AbsoluteAxis)
            _raise_if_rect_hits_vacated(
                _sheets_in_span(start_ref.sheet, end_ref.sheet, sheetnames),
                c_lo=min(start_ref.col.index, end_ref.col.index),
                r_lo=min(start_ref.row.index, end_ref.row.index),
                c_hi=max(start_ref.col.index, end_ref.col.index),
                r_hi=max(start_ref.row.index, end_ref.row.index),
                source_label=source_label,
                vacated_by_sheet=vacated_by_sheet,
            )
        case WholeColumnNode(sheet=sheet, start_col=start_col, end_col=end_col):
            assert isinstance(start_col, AbsoluteAxis) and isinstance(end_col, AbsoluteAxis)
            _raise_if_rect_hits_vacated(
                [sheet],
                c_lo=min(start_col.index, end_col.index),
                r_lo=1,
                c_hi=max(start_col.index, end_col.index),
                r_hi=1_048_576,
                source_label=source_label,
                vacated_by_sheet=vacated_by_sheet,
            )
        case WholeRowNode(sheet=sheet, start_row=start_row, end_row=end_row):
            assert isinstance(start_row, AbsoluteAxis) and isinstance(end_row, AbsoluteAxis)
            _raise_if_rect_hits_vacated(
                [sheet],
                c_lo=1,
                r_lo=min(start_row.index, end_row.index),
                c_hi=16_384,
                r_hi=max(start_row.index, end_row.index),
                source_label=source_label,
                vacated_by_sheet=vacated_by_sheet,
            )
        case FunctionCallNode(name, args):
            fn_upper = normalize_excel_function_name(name)
            if fn_upper in _DYNAMIC_OVERLAY_FNS:
                raise ValueError(
                    f"Cannot vacate cells during in-place writeback: {source_label} uses "
                    f"dynamic {fn_upper}() whose targets cannot be statically verified"
                )
            for arg in args:
                _walk_bound_ast_for_vacated(
                    arg,
                    source_label=source_label,
                    vacated_by_sheet=vacated_by_sheet,
                    sheetnames=sheetnames,
                )
        case BinaryOpNode(_, left, right):
            _walk_bound_ast_for_vacated(
                left,
                source_label=source_label,
                vacated_by_sheet=vacated_by_sheet,
                sheetnames=sheetnames,
            )
            _walk_bound_ast_for_vacated(
                right,
                source_label=source_label,
                vacated_by_sheet=vacated_by_sheet,
                sheetnames=sheetnames,
            )
        case UnaryOpNode(_, operand):
            _walk_bound_ast_for_vacated(
                operand,
                source_label=source_label,
                vacated_by_sheet=vacated_by_sheet,
                sheetnames=sheetnames,
            )
        case _:
            return


def _check_unparsed_formula_for_vacated(
    formula_text: str,
    *,
    host_sheet: str,
    source_label: str,
    vacated_by_sheet: dict[str, list[tuple[int, int, str]]],
    named_ranges: dict[str, tuple[str, str]],
    named_range_ranges: dict[str, tuple[str, str, str]],
) -> None:
    dyn_hits = _find_function_calls_with_spans(formula_text, _DYNAMIC_OVERLAY_FNS)
    if dyn_hits:
        fn_name = dyn_hits[0][0].upper()
        raise ValueError(
            f"Cannot vacate cells during in-place writeback: {source_label} uses "
            f"dynamic {fn_name}() whose targets cannot be statically verified"
        )

    normalizer = FormulaNormalizer(named_ranges, named_range_ranges)
    masked = mask_ref_only_function_calls(formula_text)
    norm = normalizer.normalize(masked, host_sheet)

    range_spans: list[tuple[int, int]] = []
    for start, end, span in parse_range_refs_with_spans(norm):
        range_spans.append(span)
        sh = start.sheet if start.sheet is not None else host_sheet
        end_sh = end.sheet if end.sheet is not None else sh
        sheets = [sh] if sh == end_sh else [sh, end_sh]
        if start.range_kind == "whole_column":
            c1 = int(column_index_from_string(start.column))
            c2 = int(column_index_from_string(end.column))
            _raise_if_rect_hits_vacated(
                sheets,
                c_lo=min(c1, c2),
                r_lo=1,
                c_hi=max(c1, c2),
                r_hi=1_048_576,
                source_label=source_label,
                vacated_by_sheet=vacated_by_sheet,
            )
        elif start.range_kind == "whole_row":
            _raise_if_rect_hits_vacated(
                sheets,
                c_lo=1,
                r_lo=min(start.row, end.row),
                c_hi=16_384,
                r_hi=max(start.row, end.row),
                source_label=source_label,
                vacated_by_sheet=vacated_by_sheet,
            )
        else:
            c1 = int(column_index_from_string(start.column))
            c2 = int(column_index_from_string(end.column))
            _raise_if_rect_hits_vacated(
                sheets,
                c_lo=min(c1, c2),
                r_lo=min(start.row, end.row),
                c_hi=max(c1, c2),
                r_hi=max(start.row, end.row),
                source_label=source_label,
                vacated_by_sheet=vacated_by_sheet,
            )

    without_ranges = mask_spans(norm, range_spans) if range_spans else norm
    for ref in parse_standalone_cell_refs(without_ranges):
        sh = ref.sheet if ref.sheet is not None else host_sheet
        col_i = int(column_index_from_string(ref.column))
        _raise_if_rect_hits_vacated(
            [sh],
            c_lo=col_i,
            r_lo=ref.row,
            c_hi=col_i,
            r_hi=ref.row,
            source_label=source_label,
            vacated_by_sheet=vacated_by_sheet,
        )


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
