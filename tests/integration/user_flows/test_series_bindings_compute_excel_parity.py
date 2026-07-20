"""Excel parity for soft-error `compute_*` measure capture (issue #436).

A mixed numeric / `#DIV/0!` output series must return a full Records list with
`OBS_VALUE` matching Excel's displayed cell values (`XlError` for error years).

Cached path always runs in CI. Live Excel path is `@pytest.mark.slow` and
skips when automation is unavailable.
"""

from __future__ import annotations

from collections.abc import Callable
from copy import deepcopy
from pathlib import Path
from typing import Any, cast

import pytest
import xlsxwriter

from excel_grapher.core.types import XlError
from excel_grapher.exporter import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import validate_bindings_document
from tests.utils.excel_workbook_parity import (
    assert_workbook_parity,
    compare_cached_to_evaluator,
)
from tests.utils.modify_and_recalculate import (
    ExcelRecalculationError,
    modify_and_recalculate_workbook,
)

TARGETS = ("Sheet1!C2", "Sheet1!D2")


def _mixed_error_bindings(*, workbook: str) -> dict[str, Any]:
    return {
        "schema_version": "1.3.0",
        "workbook": workbook,
        "series": [
            {
                "id": "mixed_output",
                "sheet": "Sheet1",
                "data_range": "Sheet1!C2:D2",
                "layout": "series",
                "output": {"compute": {"name": "compute_mixed_output"}},
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": "float",
                        "bind": {"kind": "data_cell", "read": "float"},
                    },
                    "dimensions": [
                        {
                            "concept": "TIME_PERIOD",
                            "role": "key",
                            "scope": "cell",
                            "bind": {
                                "kind": "column_header",
                                "header_row": 1,
                                "read": "int",
                            },
                        }
                    ],
                },
                "key": ["TIME_PERIOD"],
            }
        ],
    }


def _write_mixed_error_workbook(path: Path, *, with_cached_values: bool) -> None:
    """Write C2 numeric and D2 `#DIV/0!` series leaves with TIME_PERIOD headers."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 2, 2020)
    ws.write_number(0, 3, 2021)
    ws.write_number("B2", 5.0)
    if with_cached_values:
        ws.write_formula("C2", "=B2*2", None, 10.0)
        ws.write_formula("D2", "=1/0", None, "#DIV/0!")
    else:
        ws.write_formula("C2", "=B2*2")
        ws.write_formula("D2", "=1/0")
    wb.close()


def _normalize_obs_value(value: object) -> object:
    """Map exported or package error tokens to package `XlError` for comparison."""
    if isinstance(value, XlError):
        return value
    err = XlError.from_text(str(value))
    return err if err is not None else value


def _compute_records(workbook: Path) -> list[dict[str, object]]:
    bindings = validate_bindings_document(deepcopy(_mixed_error_bindings(workbook=workbook.name)))
    graph = create_dependency_graph(workbook, list(TARGETS), load_values=True)
    assert_workbook_parity(graph, list(TARGETS))

    with CodeGenerator(graph) as gen:
        code = gen.generate(
            list(TARGETS),
            series_bindings=bindings,
            bindings_workbook=workbook,
        )

    ns: dict[str, object] = {}
    exec(code, ns)
    compute = cast(Callable[..., list[dict[str, object]]], ns["compute_mixed_output"])
    records = compute()
    assert len(records) == 2

    by_address = {
        "Sheet1!C2": next(r for r in records if r["TIME_PERIOD"] == 2020),
        "Sheet1!D2": next(r for r in records if r["TIME_PERIOD"] == 2021),
    }
    for address, record in by_address.items():
        node = graph.get_node(address)
        assert node is not None
        obs = _normalize_obs_value(record["OBS_VALUE"])
        mismatch = compare_cached_to_evaluator(node.value, obs)
        assert mismatch is None, (
            f"{address}: excel={node.value!r} compute={record['OBS_VALUE']!r} kind={mismatch}"
        )

    assert by_address["Sheet1!C2"]["OBS_VALUE"] == pytest.approx(10.0)
    assert _normalize_obs_value(by_address["Sheet1!D2"]["OBS_VALUE"]) == XlError.DIV
    assert str(by_address["Sheet1!D2"]["OBS_VALUE"]) == "#DIV/0!"
    return records


def test_compute_mixed_obs_value_matches_excel_cache(tmp_path: Path) -> None:
    """Cached Excel values ↔ evaluator ↔ soft-error `compute_*` OBS_VALUE."""
    workbook = tmp_path / "mixed_error_cache.xlsx"
    _write_mixed_error_workbook(workbook, with_cached_values=True)
    _compute_records(workbook)


@pytest.mark.slow
def test_compute_mixed_obs_value_matches_live_excel(tmp_path: Path) -> None:
    """Live Excel recalc ↔ soft-error `compute_*` OBS_VALUE (skip if unavailable)."""
    input_path = tmp_path / "mixed_error_live.xlsx"
    output_path = tmp_path / "mixed_error_live_recalc.xlsx"
    _write_mixed_error_workbook(input_path, with_cached_values=False)

    try:
        modify_and_recalculate_workbook(input_path, output_path, {})
    except (ExcelRecalculationError, RuntimeError, ImportError) as exc:
        pytest.skip(f"Excel recalculation not available: {exc}")

    _compute_records(output_path)
