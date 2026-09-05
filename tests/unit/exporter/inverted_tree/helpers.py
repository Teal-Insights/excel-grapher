"""Shared helpers for inverted-tree shape-unit tests."""

from __future__ import annotations

import copy
import importlib
import inspect
import re
import sys
import types
from collections.abc import Callable, Mapping, Sequence
from pathlib import Path
from typing import Any, Literal

from fastpyxl.utils.cell import column_index_from_string, get_column_letter

from excel_grapher.core.address_keys import (
    format_key,
    format_range_key,
    parse_address,
    split_address_on_colon,
)
from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    SeriesCatalog,
    build_catalog,
    build_schedule_index,
)
from excel_grapher.exporter.inverted_tree.deps import SeriesDeps, collect_all_deps
from excel_grapher.exporter.inverted_tree.emit import generate_inverted_tree_modules
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.blank_ranges import normalize_blank_range_specs
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings import validate_bindings_document
from excel_grapher.series_bindings.types import WorkbookSeriesBindings
from excel_grapher.series_bindings.workflow import all_series_targets

_CELL_COORD_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$")
_FORMULA_A1_RE = re.compile(
    r"(?<![A-Za-z0-9_])"
    r"(?P<sheet>'(?:[^']|'')+'!|[A-Za-z_][\w.]*!)?"
    r"(?P<coord>\$?[A-Za-z]{1,3}\$?\d+)"
    r"(?![A-Za-z0-9_(])"
)


def make_catalog(
    series: dict[str, BoundSeries],
    order: tuple[str, ...],
    address_to_id: dict[str, str],
) -> SeriesCatalog:
    """Assemble a catalog, computing the schedule index from `series`."""
    return SeriesCatalog(
        series=series,
        order=order,
        address_to_id=address_to_id,
        schedule=build_schedule_index(series),
    )


def write_workbook(
    path: Path,
    sheets: Mapping[str, Mapping[str, object]],
    *,
    defined_names: Mapping[str, str] | None = None,
) -> Path:
    """Write a multi-sheet workbook with `fastpyxl` (formulas as `=...` strings)."""
    from fastpyxl import Workbook
    from fastpyxl.workbook.defined_name import DefinedName

    wb = Workbook()
    default = wb.active
    first = True
    for title, cells in sheets.items():
        if first:
            ws = default
            ws.title = title
            first = False
        else:
            ws = wb.create_sheet(title)
        for addr, value in cells.items():
            ws[addr] = value
    if defined_names:
        for name, attr_text in defined_names.items():
            wb.defined_names.add(DefinedName(name, attr_text=attr_text))
    wb.save(path)
    return path


def series_entry(
    series_id: str,
    data_range: str,
    *,
    layout: str = "scalar",
    direction: str = "input",
    dtype: str = "float",
    header_row: int | None = None,
    label_column: str | None = None,
    compute_name: str | None = None,
    key: Sequence[str] | None = None,
    key_concept: str = "TIME_PERIOD",
    key_read: str = "int",
    domain: Mapping[str, Any] | None = None,
    value_map: Mapping[Any, Any] | None = None,
) -> dict[str, Any]:
    """Build a minimal series dict that passes schema validation."""
    sheet = data_range.split("!", 1)[0]
    dimensions: list[dict[str, Any]] = []
    key_fields = list(key) if key is not None else []
    if layout in {"series", "row_series"} and header_row is not None:
        dimensions.append(
            {
                "concept": key_concept,
                "role": "key",
                "scope": "cell",
                "bind": {"kind": "column_header", "header_row": header_row, "read": key_read},
            }
        )
        if key is None:
            key_fields = [key_concept]
    elif layout in {"series", "row_series"} and label_column is not None:
        dimensions.append(
            {
                "concept": key_concept,
                "role": "key",
                "scope": "cell",
                "bind": {"kind": "row_label", "label_column": label_column, "read": key_read},
            }
        )
        if key is None:
            key_fields = [key_concept]
    entry: dict[str, Any] = {
        "id": series_id,
        "sheet": sheet,
        "data_range": data_range,
        "layout": layout,
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": dtype,
                "bind": {"kind": "data_cell", "read": dtype if dtype != "number" else "float"},
            },
            "dimensions": dimensions,
        },
        "key": key_fields,
    }
    if direction == "input":
        entry["input"] = {"setter": {"name": f"set_{series_id}"}}
        if domain is not None:
            entry["input"]["domain"] = dict(domain)
        if value_map is not None:
            entry["input"]["value_map"] = dict(value_map)
    elif direction == "output":
        entry["output"] = {"compute": {"name": compute_name or f"compute_{series_id}"}}
    elif direction == "internal":
        entry["internal"] = {}
    elif direction == "constant":
        entry["constant"] = {}
    else:
        raise ValueError(f"unknown direction {direction!r}")
    return entry


def bindings_document(*series: dict[str, Any], schema_version: str = "1.13.0") -> dict[str, Any]:
    return {
        "schema_version": schema_version,
        "concept_scheme": {
            "id": "inverted_tree_shape",
            "concepts": [
                {"id": "TIME_PERIOD", "dtype": "int"},
                {"id": "OBS_VALUE", "dtype": "number"},
                {"id": "COUNTRY", "dtype": "string"},
                {"id": "SCENARIO", "dtype": "string"},
                {"id": "SHOCK_PARAMETER", "dtype": "string"},
            ],
        },
        "series": list(series),
    }


def inverted_graph_parts(
    workbook: Path,
    document: dict[str, Any],
    *,
    dynamic_refs: DynamicRefConfig | None = None,
    blank_ranges: Sequence[str] | None = None,
) -> tuple[SeriesCatalog, dict[str, SeriesDeps], DependencyGraph]:
    """Return catalog, first-level deps, and graph for inverted-tree emit tests."""
    bindings: WorkbookSeriesBindings = validate_bindings_document(document)
    targets = all_series_targets(bindings, workbook=workbook)
    graph = create_dependency_graph(
        workbook,
        targets,
        load_values=True,
        use_cached_dynamic_refs=dynamic_refs is None,
        dynamic_refs=dynamic_refs,
        capture_dependency_provenance=True,
        blank_ranges=blank_ranges,
    )
    catalog = build_catalog(bindings, workbook=workbook, graph=graph)
    return (
        catalog,
        collect_all_deps(catalog, graph, blank_rects=normalize_blank_range_specs(blank_ranges)),
        graph,
    )


def generate_inverted(
    workbook: Path,
    document: dict[str, Any],
    *,
    dynamic_refs: DynamicRefConfig | None = None,
    force_rung: Literal[2, 3] | None = None,
    blank_ranges: Sequence[str] | None = None,
) -> dict[str, str]:
    bindings: WorkbookSeriesBindings = validate_bindings_document(document)
    targets = all_series_targets(bindings, workbook=workbook)
    graph = create_dependency_graph(
        workbook,
        targets,
        load_values=True,
        use_cached_dynamic_refs=dynamic_refs is None,
        dynamic_refs=dynamic_refs,
        capture_dependency_provenance=True,
        blank_ranges=blank_ranges,
    )
    return generate_inverted_tree_modules(
        graph,
        series_bindings=bindings,
        bindings_workbook=workbook,
        force_rung=force_rung,
        blank_ranges=blank_ranges,
    )


def load_package(
    modules: Mapping[str, str],
    tmp_path: Path,
    name: str = "inv_pkg",
) -> types.ModuleType:
    pkg = tmp_path / name
    pkg.mkdir(parents=True, exist_ok=True)
    for filename, content in modules.items():
        (pkg / filename).write_text(content, encoding="utf-8")
    if str(tmp_path) not in sys.path:
        sys.path.insert(0, str(tmp_path))
    for key in list(sys.modules):
        if key == name or key.startswith(name + "."):
            del sys.modules[key]
    pkg = importlib.import_module(name)
    for sub in ("api", "internals", "runtime", "data"):
        importlib.import_module(f"{name}.{sub}")
    return pkg


def load_forced_rung_packages(
    workbook: Path,
    document: dict[str, Any],
    tmp_path: Path,
    stem: str,
) -> tuple[types.ModuleType, types.ModuleType]:
    """Load packages generated at `force_rung=2` and `force_rung=3`."""
    fused = load_package(
        generate_inverted(workbook, document, force_rung=2),
        tmp_path,
        name=f"{stem}_r2",
    )
    demand = load_package(
        generate_inverted(workbook, document, force_rung=3),
        tmp_path,
        name=f"{stem}_r3",
    )
    return fused, demand


def required_param_names(function: Callable[..., object]) -> tuple[str, ...]:
    names: list[str] = []
    for name, parameter in inspect.signature(function).parameters.items():
        if parameter.default is inspect.Parameter.empty:
            names.append(name)
    return tuple(names)


def all_param_names(function: Callable[..., object]) -> tuple[str, ...]:
    return tuple(inspect.signature(function).parameters)


def transpose_cell_coord(coord: str) -> str:
    """Swap the column and row indexes of an A1 coordinate (`B1` -> `A2`)."""
    match = _CELL_COORD_RE.fullmatch(coord.strip())
    if match is None:
        raise ValueError(f"not an A1 cell coordinate: {coord!r}")
    abs_col, col, abs_row, row = match.groups()
    new_col = get_column_letter(int(row))
    new_row = column_index_from_string(col.upper())
    return f"{abs_row}{new_col}{abs_col}{new_row}"


def _order_cell_pair(start: str, end: str) -> tuple[str, str]:
    start_match = _CELL_COORD_RE.fullmatch(start)
    end_match = _CELL_COORD_RE.fullmatch(end)
    if start_match is None or end_match is None:
        return start, end
    start_col = column_index_from_string(start_match.group(2).upper())
    start_row = int(start_match.group(4))
    end_col = column_index_from_string(end_match.group(2).upper())
    end_row = int(end_match.group(4))
    if (end_col, end_row) < (start_col, start_row):
        return end, start
    return start, end


def transpose_address(address: str) -> str:
    """Transpose a bare, sheet-qualified, or same-sheet range address."""
    split = split_address_on_colon(address)
    if split is None:
        if "!" not in address:
            return transpose_cell_coord(address)
        sheet, cell = parse_address(address)
        return format_key(sheet, transpose_cell_coord(cell))
    left, right = split
    if "!" in left:
        sheet, start = parse_address(left)
    else:
        sheet, start = "", left
    if "!" in right:
        _end_sheet, end = parse_address(right)
    else:
        end = right
    start_t, end_t = _order_cell_pair(transpose_cell_coord(start), transpose_cell_coord(end))
    if not sheet:
        return f"{start_t}:{end_t}"
    return format_range_key(sheet, start_t, end_t)


def transpose_formula(text: str) -> str:
    """Rewrite A1 references in an Excel formula by swapping axes."""
    if not isinstance(text, str) or not text.startswith("="):
        return text
    parts: list[str] = []
    index = 0
    in_string = False
    while index < len(text):
        char = text[index]
        if char == '"':
            in_string = not in_string
            parts.append(char)
            index += 1
            continue
        if not in_string:
            match = _FORMULA_A1_RE.match(text, index)
            if match is not None:
                sheet = match.group("sheet") or ""
                parts.append(sheet + transpose_cell_coord(match.group("coord")))
                index = match.end()
                continue
        parts.append(char)
        index += 1
    return "".join(parts)


def transpose_sheets(
    sheets: Mapping[str, Mapping[str, object]],
) -> dict[str, dict[str, object]]:
    """Transpose every cell and formula reference in a multi-sheet mapping."""
    result: dict[str, dict[str, object]] = {}
    for title, cells in sheets.items():
        moved: dict[str, object] = {}
        for addr, value in cells.items():
            new_value = transpose_formula(value) if isinstance(value, str) else value
            moved[transpose_cell_coord(addr)] = new_value
        result[title] = moved
    return result


def _transpose_bind(bind: dict[str, Any]) -> dict[str, Any]:
    kind = bind.get("kind")
    if kind == "column_header":
        header_row = int(bind["header_row"])
        return {
            key: value
            for key, value in bind.items()
            if key not in {"kind", "header_row", "label_column"}
        } | {
            "kind": "row_label",
            "label_column": get_column_letter(header_row),
        }
    if kind == "row_label":
        label_column = str(bind["label_column"])
        return {
            key: value
            for key, value in bind.items()
            if key not in {"kind", "header_row", "label_column"}
        } | {
            "kind": "column_header",
            "header_row": column_index_from_string(label_column),
        }
    return bind


def _walk_transpose_binds(node: Any) -> Any:
    if isinstance(node, dict):
        if "kind" in node and node["kind"] in {"column_header", "row_label"}:
            return _transpose_bind(
                {key: _walk_transpose_binds(value) for key, value in node.items()}
            )
        return {key: _walk_transpose_binds(value) for key, value in node.items()}
    if isinstance(node, list):
        return [_walk_transpose_binds(item) for item in node]
    return node


def transpose_bindings(document: Mapping[str, Any]) -> dict[str, Any]:
    """Swap `column_header`/`row_label` binds and transpose `data_range`s."""
    doc = copy.deepcopy(dict(document))
    for series in doc.get("series", []):
        if "data_range" in series:
            series["data_range"] = transpose_address(str(series["data_range"]))
        if "structure" in series:
            series["structure"] = _walk_transpose_binds(series["structure"])
    return doc


def load_workbook_sheets(path: Path) -> dict[str, dict[str, object]]:
    """Read stored cell values (formulas as `=...` strings) from `path`."""
    from fastpyxl import load_workbook

    workbook = load_workbook(path)
    sheets: dict[str, dict[str, object]] = {}
    for worksheet in workbook.worksheets:
        cells: dict[str, object] = {}
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                value = cell.value
                if isinstance(value, str) and not value.startswith("="):
                    cells[cell.coordinate] = value
                elif getattr(value, "text", None) is not None:
                    text = str(value.text)
                    cells[cell.coordinate] = text if text.startswith("=") else f"={text}"
                else:
                    cells[cell.coordinate] = value
        sheets[worksheet.title] = cells
    return sheets


def write_oriented_workbook(
    path: Path,
    sheets: Mapping[str, Mapping[str, object]],
    *,
    orientation: str,
) -> Path:
    """Write `sheets`, transposing when `orientation` is `vertical`."""
    if orientation == "vertical":
        sheets = transpose_sheets(sheets)
    elif orientation != "horizontal":
        raise ValueError(f"orientation must be horizontal or vertical; got {orientation!r}")
    return write_workbook(path, sheets)


def oriented_document(document: Mapping[str, Any], orientation: str) -> dict[str, Any]:
    """Return `document`, transposed when `orientation` is `vertical`."""
    if orientation == "vertical":
        return transpose_bindings(document)
    if orientation != "horizontal":
        raise ValueError(f"orientation must be horizontal or vertical; got {orientation!r}")
    return dict(document)


def oriented_addresses(addresses: Sequence[str], orientation: str) -> tuple[str, ...]:
    """Transpose sheet-qualified addresses when `orientation` is `vertical`."""
    if orientation == "horizontal":
        return tuple(addresses)
    return tuple(transpose_address(address) for address in addresses)


def input_kwargs(catalog: SeriesCatalog, graph: DependencyGraph) -> dict[str, object]:
    """Build `compute_*` keyword arguments from loaded input-series values."""
    kwargs: dict[str, object] = {}
    for series in catalog.input_series():
        values: list[object] = []
        for cell in series.cells:
            node = graph.get_node(cell)
            values.append(None if node is None else node.value)
        kwargs[series.series_id] = values[0] if series.is_scalar else tuple(values)
    return kwargs


def call_compute(pkg: types.ModuleType, series_id: str, kwargs: Mapping[str, object]) -> object:
    """Call `pkg.compute_<series_id>` with the intersection of `kwargs`."""
    name = f"compute_{series_id}"
    function = getattr(pkg, name)
    accepted = set(inspect.signature(function).parameters)
    return function(**{key: value for key, value in kwargs.items() if key in accepted})
