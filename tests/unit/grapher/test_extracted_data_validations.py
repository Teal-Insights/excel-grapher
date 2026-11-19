"""Graph-owned worksheet data validations (#1090)."""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from typing import Any

import fastpyxl
from fastpyxl.worksheet.datavalidation import DataValidation, expand_cell_ranges

from excel_grapher.grapher import create_dependency_graph, write_workbook
from excel_grapher.grapher.cache import (
    build_graph_cache_meta,
    dependency_graph_from_json,
    dependency_graph_to_json,
    save_graph_cache,
    try_load_graph_cache,
)
from excel_grapher.grapher.graph_pickle import dump_graph, load_graph
from excel_grapher.series_bindings import validate_bindings_document


def _dv(sqref: str, **kwargs: Any) -> DataValidation:
    rule = DataValidation(**kwargs)
    for part in sqref.split():
        rule.add(part)
    return rule


def _book(path: Path) -> Path:
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = "x"
    ws["B1"] = "=A1"
    ws["C1"] = 3
    ws["D1"] = "outside"
    ws.add_data_validation(_dv("A1", type="list", formula1='"x,y"', allow_blank=True))
    ws.add_data_validation(
        _dv("B1", type="custom", formula1="B1<>0", errorStyle="stop", error="no")
    )
    ws.add_data_validation(_dv("C1", type="whole", operator="between", formula1="1", formula2="9"))
    ws.add_data_validation(
        _dv(
            "A1:C1",
            type="decimal",
            operator="greaterThan",
            formula1="0",
            showErrorMessage=True,
            error="positive",
        )
    )
    ws.add_data_validation(_dv("D1", type="list", formula1='"keep"'))
    other = wb.create_sheet("Notes")
    other["A1"] = "note"
    wb.save(path)
    wb.close()
    return path


def _rows(path: Path, sheet: str = "Sheet1") -> list[dict[str, Any]]:
    wb = fastpyxl.load_workbook(path)
    try:
        found: list[dict[str, Any]] = []
        for rule in wb[sheet].data_validations.dataValidation:
            sqref = str(rule.sqref)
            found.append(
                {
                    "type": rule.type,
                    "operator": rule.operator,
                    "formula1": rule.formula1,
                    "formula2": rule.formula2,
                    "allow_blank": bool(rule.allowBlank),
                    "error_style": rule.errorStyle,
                    "error": rule.error,
                    "cells": expand_cell_ranges(sqref) if sqref else set(),
                }
            )
        return found
    finally:
        wb.close()


def _covers(rows: list[dict[str, Any]], coord: str) -> list[dict[str, Any]]:
    return [row for row in rows if coord in row["cells"]]


def test_extract_stores_compiled_and_original_rules_for_materialized_cells(
    tmp_path: Path,
) -> None:
    source = _book(tmp_path / "source.xlsx")
    graph = create_dependency_graph(source, ["Sheet1!B1", "Sheet1!C1"], load_values=True)
    by_address: dict[str, list[Any]] = {}
    for suggestion in graph.data_validations:
        by_address.setdefault(suggestion.address, []).append(suggestion)

    assert set(by_address) == {"Sheet1!A1", "Sheet1!B1", "Sheet1!C1"}
    listed = next(item for item in by_address["Sheet1!A1"] if item.validation.type == "list")
    assert listed.status == "domain"
    assert listed.domain == {"enum": ["x", "y"]}
    assert listed.allow_blank is True
    assert listed.validation.formula1 == '"x,y"'

    custom = next(item for item in by_address["Sheet1!B1"] if item.validation.type == "custom")
    assert custom.status == "unsupported"
    assert custom.reason == "unsupported_type"
    assert custom.domain is None
    assert custom.validation.formula1 == "B1<>0"
    assert custom.validation.error == "no"

    bounded = next(item for item in by_address["Sheet1!C1"] if item.validation.type == "whole")
    assert bounded.status == "domain"
    assert bounded.domain == {"between": {"min": 1, "max": 9}}

    one_sided = next(item for item in by_address["Sheet1!C1"] if item.validation.type == "decimal")
    assert one_sided.status == "unsupported"
    assert one_sided.reason == "unsupported_operator"
    assert one_sided.validation.operator == "greaterThan"
    assert one_sided.validation.formula1 == "0"


def test_vacated_cells_keep_validations_through_json_and_pickle(tmp_path: Path) -> None:
    source = _book(tmp_path / "source.xlsx")
    graph = create_dependency_graph(source, ["Sheet1!C1"], load_values=True)
    graph.remove_node("Sheet1!C1")
    assert "Sheet1!C1" not in graph
    assert any(item.address == "Sheet1!C1" for item in graph.data_validations)

    restored = dependency_graph_from_json(dependency_graph_to_json(graph))
    assert restored.data_validations == graph.data_validations
    assert "Sheet1!C1" not in restored

    blob = tmp_path / "graph.egdg"
    dump_graph(graph, blob)
    pickled = load_graph(blob)
    assert pickled.data_validations == graph.data_validations

    cache = tmp_path / "graph.json"
    meta = build_graph_cache_meta(source, ["Sheet1!C1"], extraction_params={})
    save_graph_cache(cache, graph, meta)
    loaded = try_load_graph_cache(cache, expected_meta=meta)
    assert loaded is not None
    assert loaded.data_validations == graph.data_validations


def test_clean_room_write_emits_stored_rules_and_honors_removal(tmp_path: Path) -> None:
    source = _book(tmp_path / "source.xlsx")
    graph = create_dependency_graph(source, ["Sheet1!B1", "Sheet1!C1"], load_values=True)
    dest = tmp_path / "out.xlsx"
    write_workbook(graph, dest)

    rows = _rows(dest)
    listed = _covers(rows, "A1")
    assert any(row["type"] == "list" and row["formula1"] == '"x,y"' for row in listed)
    assert any(
        row["type"] == "custom" and row["formula1"] == "B1<>0" and row["error"] == "no"
        for row in _covers(rows, "B1")
    )
    assert any(
        row["type"] == "whole" and row["formula1"] == "1" and row["formula2"] == "9"
        for row in _covers(rows, "C1")
    )
    assert any(row["type"] == "decimal" and row["operator"] == "greaterThan" for row in rows)
    assert _covers(rows, "D1") == []

    graph.data_validations = tuple(
        item
        for item in graph.data_validations
        if not (item.address == "Sheet1!C1" and item.validation.type == "whole")
    )
    dropped = tmp_path / "dropped.xlsx"
    write_workbook(graph, dropped)
    dropped_rows = _rows(dropped)
    assert not any(row["type"] == "whole" for row in _covers(dropped_rows, "C1"))
    assert any(row["type"] == "decimal" for row in _covers(dropped_rows, "C1"))

    edited_source = next(
        item for item in graph.data_validations if item.validation.type == "custom"
    )
    edited = replace(
        edited_source,
        validation=replace(edited_source.validation, formula1="B1>1", error="edited"),
    )
    graph.data_validations = tuple(
        edited if item.validation.type == "custom" else item for item in graph.data_validations
    )
    edited_path = tmp_path / "edited.xlsx"
    write_workbook(graph, edited_path)
    custom = next(row for row in _covers(_rows(edited_path), "B1") if row["type"] == "custom")
    assert custom["formula1"] == "B1>1"
    assert custom["error"] == "edited"


def test_authored_domain_wins_over_stored_worksheet_rule(tmp_path: Path) -> None:
    source = _book(tmp_path / "source.xlsx")
    graph = create_dependency_graph(source, ["Sheet1!A1"], load_values=True)
    bindings = validate_bindings_document(
        {
            "schema_version": "1.21.0",
            "series": [
                {
                    "id": "choice",
                    "sheet": "Sheet1",
                    "data_range": "Sheet1!A1",
                    "layout": "scalar",
                    "domain": {"enum": ["c", "d"]},
                    "input": {},
                    "structure": {
                        "measure": {
                            "concept": "OBS_VALUE",
                            "dtype": "string",
                            "bind": {"kind": "data_cell", "read": "string"},
                        },
                        "dimensions": [],
                    },
                    "key": [],
                }
            ],
        }
    )
    dest = tmp_path / "out.xlsx"
    write_workbook(graph, dest, series_bindings=bindings, bindings_workbook=source)
    listed = [row for row in _covers(_rows(dest), "A1") if row["type"] == "list"]
    assert len(listed) == 1
    assert listed[0]["formula1"] == '"c,d"'


def test_replace_node_formula_records_validations_for_new_cells(tmp_path: Path) -> None:
    source = _book(tmp_path / "source.xlsx")
    graph = create_dependency_graph(source, ["Sheet1!B1"], load_values=True)
    before = {item.address for item in graph.data_validations}
    assert "Sheet1!C1" not in before
    graph.replace_node_formula("Sheet1!B1", "=C1", "=Sheet1!C1", workbook=source)
    assert any(
        item.address == "Sheet1!C1" and item.domain == {"between": {"min": 1, "max": 9}}
        for item in graph.data_validations
    )
    # A validation the caller already dropped is not restored for an existing cell.
    graph.data_validations = tuple(
        item for item in graph.data_validations if item.address != "Sheet1!A1"
    )
    graph.replace_node_formula("Sheet1!B1", "=A1", "=Sheet1!A1", workbook=source)
    assert not any(item.address == "Sheet1!A1" for item in graph.data_validations)
