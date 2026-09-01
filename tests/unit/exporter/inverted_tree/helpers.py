"""Shared helpers for inverted-tree shape-unit tests."""

from __future__ import annotations

import importlib
import inspect
import sys
import types
from collections.abc import Callable, Mapping, Sequence
from pathlib import Path
from typing import Any

from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from excel_grapher.series_bindings import validate_bindings_document
from excel_grapher.series_bindings.types import WorkbookSeriesBindings
from excel_grapher.series_bindings.workflow import all_series_targets


def write_workbook(path: Path, sheets: Mapping[str, Mapping[str, object]]) -> Path:
    """Write a multi-sheet workbook with `fastpyxl` (formulas as `=...` strings)."""
    from fastpyxl import Workbook

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
    compute_name: str | None = None,
    key: Sequence[str] | None = None,
    key_concept: str = "TIME_PERIOD",
    key_read: str = "int",
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
        if not key_fields:
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
                {"id": "SHOCK_PARAMETER", "dtype": "string"},
            ],
        },
        "series": list(series),
    }


def generate_inverted(
    workbook: Path,
    document: dict[str, Any],
    *,
    dynamic_refs: DynamicRefConfig | None = None,
) -> dict[str, str]:
    bindings: WorkbookSeriesBindings = validate_bindings_document(document)
    targets = all_series_targets(bindings, workbook=workbook)
    graph = create_dependency_graph(
        workbook,
        targets,
        load_values=True,
        use_cached_dynamic_refs=dynamic_refs is None,
        dynamic_refs=dynamic_refs,
    )
    with CodeGenerator(graph) as gen:
        return gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
            paradigm="inverted_tree",
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


def required_param_names(function: Callable[..., object]) -> tuple[str, ...]:
    names: list[str] = []
    for name, parameter in inspect.signature(function).parameters.items():
        if parameter.default is inspect.Parameter.empty:
            names.append(name)
    return tuple(names)


def all_param_names(function: Callable[..., object]) -> tuple[str, ...]:
    return tuple(inspect.signature(function).parameters)
