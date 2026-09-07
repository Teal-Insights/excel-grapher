"""CHOOSE into a joined TIME_PERIOD row is a keyed lookup, not an N-way dual (#754).

A host member that indexes a producer year-row with `CHOOSE(k, p[t0], …, p[tn])`
(`n > 2`) already joins on the non-time keys. The CHOOSE arguments are one
instrument's cumulative schedule; `k` picks the year. `_is_keyed_multi_read`
used to require a host-or-literal pattern set that is identical across members,
so a matching host year bound as `host` (not `lit`) made the sets disagree and
fail-closed at `more than two positions`.

The same edge shape is a positional catalog along `TIME_PERIOD` when every
multi-read member of a joined partition hits the same year set: join
`INSTRUMENT`, select the year from the CHOOSE (or an explicit sum of that
row). Partitions may have different widths (#760): IMF `{2024…2051}` and
IDA's longer horizon are still one cumulative series. `INDEX(row, k)` is a
lookup-window spelling of the same access; this shape covers the identity-hit
form that LIC-DSF writes as `CHOOSE`.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    call_compute,
    generate_inverted,
    input_kwargs,
    inverted_graph_parts,
    load_package,
    write_workbook,
)

_TIME_DIM = {
    "id": "TIME_PERIOD",
    "concept": "TIME_PERIOD",
    "role": "key",
    "scope": "cell",
    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
}

_IMF_CUMULATIVE = (10.0, 20.0, 30.0, 40.0)
_IDA_CUMULATIVE = (100.0, 200.0, 300.0, 400.0)
_IDA_RAGGED_CUMULATIVE = (100.0, 200.0, 300.0, 400.0, 500.0)
_IMF_INDEX = (2, 3, 1, 4)
_IDA_INDEX = (1, 4, 2, 3)
_IDA_RAGGED_INDEX = (1, 5, 2, 3, 4)


def _measure() -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }


def _instrument(rows: list[int]) -> dict[str, Any]:
    return {
        "id": "INSTRUMENT",
        "concept": "INSTRUMENT",
        "role": "key",
        "scope": "cell",
        "bind": {
            "kind": "value_map",
            "values": {"IMF": rows[0], "IDA": rows[1]},
            "read": "string",
        },
    }


def _document(*series: dict[str, Any]) -> dict[str, Any]:
    doc = bindings_document(*series, schema_version="1.14.0")
    known = {item["id"] for item in doc["concept_scheme"]["concepts"]}
    if "INSTRUMENT" not in known:
        doc["concept_scheme"]["concepts"].append({"id": "INSTRUMENT", "dtype": "string"})
    return doc


def _choose_formula(index_cell: str, row: int, cols: str = "DEFG") -> str:
    args = ",".join(f"{col}{row}" for col in cols)
    return f"=IF({index_cell}=0,0,CHOOSE({index_cell},{args}))"


def _sum_formula(row: int) -> str:
    return f"=D{row}+E{row}+F{row}+G{row}"


def _mcve_sheets(
    *,
    host_formula: str = "choose",
) -> dict[str, dict[str, object]]:
    cells: dict[str, object] = {
        "D1": 2024,
        "E1": 2025,
        "F1": 2026,
        "G1": 2027,
        "D5": 10,
        "E5": 20,
        "F5": 30,
        "G5": 40,
        "D10": 100,
        "E10": 200,
        "F10": 300,
        "G10": 400,
        "A12": "=D4",
    }
    for col, idx in zip("DEFG", _IMF_INDEX, strict=True):
        cells[f"{col}3"] = idx
        if host_formula == "choose":
            cells[f"{col}4"] = _choose_formula(f"{col}3", 5)
        else:
            cells[f"{col}4"] = _sum_formula(5)
    for col, idx in zip("DEFG", _IDA_INDEX, strict=True):
        cells[f"{col}8"] = idx
        if host_formula == "choose":
            cells[f"{col}9"] = _choose_formula(f"{col}8", 10)
        else:
            cells[f"{col}9"] = _sum_formula(10)
    return {"PV": cells}


def _one_instrument_sheets() -> dict[str, dict[str, object]]:
    cells: dict[str, object] = {
        "D1": 2024,
        "E1": 2025,
        "F1": 2026,
        "G1": 2027,
        "D5": 10,
        "E5": 20,
        "F5": 30,
        "G5": 40,
        "A12": "=D4",
    }
    for col, idx in zip("DEFG", _IMF_INDEX, strict=True):
        cells[f"{col}3"] = idx
        cells[f"{col}4"] = _choose_formula(f"{col}3", 5)
    return {"PV": cells}


def _ragged_sheets() -> dict[str, dict[str, object]]:
    """IMF CHOOSE lists four years; IDA lists five (#760)."""
    cells: dict[str, object] = {
        "D1": 2024,
        "E1": 2025,
        "F1": 2026,
        "G1": 2027,
        "H1": 2028,
        "D5": 10,
        "E5": 20,
        "F5": 30,
        "G5": 40,
        "D10": 100,
        "E10": 200,
        "F10": 300,
        "G10": 400,
        "H10": 500,
        "A12": "=D4",
    }
    for col, idx in zip("DEFG", _IMF_INDEX, strict=True):
        cells[f"{col}3"] = idx
        cells[f"{col}4"] = _choose_formula(f"{col}3", 5, "DEFG")
    for col, idx in zip("DEFGH", _IDA_RAGGED_INDEX, strict=True):
        cells[f"{col}8"] = idx
        cells[f"{col}9"] = _choose_formula(f"{col}8", 10, "DEFGH")
    return {"PV": cells}


def _series(
    series_id: str,
    data_range: str | list[str],
    *,
    rows: list[int],
    direction: str,
) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "id": series_id,
        "sheet": "PV",
        "data_range": data_range,
        "layout": "series",
        "structure": {
            "measure": _measure(),
            "dimensions": [_instrument(rows), _TIME_DIM],
        },
        "key": ["INSTRUMENT", "TIME_PERIOD"],
    }
    if direction == "input":
        entry["input"] = {"setter": {"name": f"set_{series_id}"}}
    elif direction == "internal":
        entry["internal"] = {}
    else:
        entry["output"] = {"compute": {"name": f"compute_{series_id}"}}
    return entry


def _one_instrument_dim(row: int) -> dict[str, Any]:
    return {
        "id": "INSTRUMENT",
        "concept": "INSTRUMENT",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "value_map", "values": {"IMF": row}, "read": "string"},
    }


def _one_instrument_series(
    series_id: str,
    data_range: str,
    *,
    row: int,
    direction: str,
) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "id": series_id,
        "sheet": "PV",
        "data_range": data_range,
        "layout": "series",
        "structure": {
            "measure": _measure(),
            "dimensions": [_one_instrument_dim(row), _TIME_DIM],
        },
        "key": ["INSTRUMENT", "TIME_PERIOD"],
    }
    if direction == "input":
        entry["input"] = {"setter": {"name": f"set_{series_id}"}}
    else:
        entry["output"] = {"compute": {"name": f"compute_{series_id}"}}
    return entry


def _mcve_bindings(*, one_instrument: bool = False, ragged: bool = False) -> dict[str, Any]:
    result = {
        "id": "result",
        "sheet": "PV",
        "data_range": "PV!A12",
        "layout": "scalar",
        "output": {"compute": {"name": "compute_result"}},
        "structure": {"measure": _measure(), "dimensions": []},
        "key": [],
    }
    if one_instrument:
        return _document(
            _one_instrument_series("index", "PV!D3:G3", row=3, direction="input"),
            _one_instrument_series("cumulative", "PV!D5:G5", row=5, direction="input"),
            _one_instrument_series("post_grace", "PV!D4:G4", row=4, direction="output"),
            result,
        )
    if ragged:
        return _document(
            _series("index", ["PV!D3:G3", "PV!D8:H8"], rows=[3, 8], direction="input"),
            _series("cumulative", ["PV!D5:G5", "PV!D10:H10"], rows=[5, 10], direction="input"),
            _series("post_grace", ["PV!D4:G4", "PV!D9:H9"], rows=[4, 9], direction="output"),
            result,
        )
    return _document(
        _series("index", ["PV!D3:G3", "PV!D8:G8"], rows=[3, 8], direction="input"),
        _series("cumulative", ["PV!D5:G5", "PV!D10:G10"], rows=[5, 10], direction="input"),
        _series("post_grace", ["PV!D4:G4", "PV!D9:G9"], rows=[4, 9], direction="output"),
        result,
    )


def _choose_expected() -> tuple[float, ...]:
    imf = tuple(_IMF_CUMULATIVE[i - 1] for i in _IMF_INDEX)
    ida = tuple(_IDA_CUMULATIVE[i - 1] for i in _IDA_INDEX)
    return imf + ida


def _ragged_expected() -> tuple[float, ...]:
    imf = tuple(_IMF_CUMULATIVE[i - 1] for i in _IMF_INDEX)
    ida = tuple(_IDA_RAGGED_CUMULATIVE[i - 1] for i in _IDA_RAGGED_INDEX)
    return imf + ida


def _sum_expected() -> tuple[float, ...]:
    imf_total = float(sum(_IMF_CUMULATIVE))
    ida_total = float(sum(_IDA_CUMULATIVE))
    return (imf_total,) * 4 + (ida_total,) * 4


def _unwrap(value: object) -> object:
    """Return the sole member of a 1-tuple compute result."""
    if isinstance(value, tuple) and len(value) == 1:
        return value[0]
    return value


def _eval_cells(workbook: Path, cells: list[str]) -> dict[str, object]:
    return FormulaEvaluator(create_dependency_graph(workbook, cells, load_values=True)).evaluate(
        cells
    )


def test_choose_year_row_is_keyed(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a35.xlsx", _mcve_sheets())
    catalog, deps, _graph = inverted_graph_parts(workbook, _mcve_bindings())
    host = deps["post_grace"]
    assert "cumulative" in host.param_ids
    assert "cumulative" in host.keyed_ids
    assert "cumulative" not in host.lagged_ids
    assert "cumulative" not in host.aligned_ids
    assert catalog.get("cumulative").cells == (
        "PV!D5",
        "PV!E5",
        "PV!F5",
        "PV!G5",
        "PV!D10",
        "PV!E10",
        "PV!F10",
        "PV!G10",
    )


def test_choose_year_row_emits_and_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a35_eval.xlsx", _mcve_sheets())
    document = _mcve_bindings()
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    assert "cumulative" in deps["post_grace"].keyed_ids
    modules = generate_inverted(workbook, document)
    internals = modules["internals.py"]
    assert "cumulative[i + 1]" not in internals
    assert "cumulative[0]" in internals
    assert "cumulative[4]" in internals
    pkg = load_package(modules, tmp_path, name="a35_eval")
    cells = [f"PV!{col}{row}" for row in (4, 9) for col in "DEFG"]
    expected = _eval_cells(workbook, cells)
    kwargs = input_kwargs(catalog, graph)
    got = call_compute(pkg, "post_grace", kwargs)
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx(_choose_expected())
    assert _unwrap(call_compute(pkg, "result", kwargs)) == pytest.approx(20.0)


def test_choose_year_row_one_instrument_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a35_one.xlsx", _one_instrument_sheets())
    document = _mcve_bindings(one_instrument=True)
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    assert "cumulative" in deps["post_grace"].keyed_ids
    assert "cumulative" not in deps["post_grace"].lagged_ids
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a35_one")
    cells = [f"PV!{col}4" for col in "DEFG"]
    expected = _eval_cells(workbook, cells)
    got = call_compute(pkg, "post_grace", input_kwargs(catalog, graph))
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx(tuple(_IMF_CUMULATIVE[i - 1] for i in _IMF_INDEX))


def test_year_row_sum_is_keyed_and_matches_evaluator(tmp_path: Path) -> None:
    """An explicit sum of the joined year-row is the same N-way identity hit."""
    workbook = write_workbook(tmp_path / "a35_sum.xlsx", _mcve_sheets(host_formula="sum"))
    document = _mcve_bindings()
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    assert "cumulative" in deps["post_grace"].keyed_ids
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a35_sum")
    cells = [f"PV!{col}{row}" for row in (4, 9) for col in "DEFG"]
    expected = _eval_cells(workbook, cells)
    got = call_compute(pkg, "post_grace", input_kwargs(catalog, graph))
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx(_sum_expected())


def test_choose_year_row_ragged_widths_is_keyed(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a35_ragged.xlsx", _ragged_sheets())
    catalog, deps, _graph = inverted_graph_parts(workbook, _mcve_bindings(ragged=True))
    host = deps["post_grace"]
    assert "cumulative" in host.param_ids
    assert "cumulative" in host.keyed_ids
    assert "cumulative" not in host.lagged_ids
    assert "cumulative" not in host.aligned_ids
    assert catalog.get("cumulative").cells == (
        "PV!D5",
        "PV!E5",
        "PV!F5",
        "PV!G5",
        "PV!D10",
        "PV!E10",
        "PV!F10",
        "PV!G10",
        "PV!H10",
    )


def test_choose_year_row_ragged_widths_emits_and_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a35_ragged_eval.xlsx", _ragged_sheets())
    document = _mcve_bindings(ragged=True)
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    assert "cumulative" in deps["post_grace"].keyed_ids
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a35_ragged")
    cells = [f"PV!{col}4" for col in "DEFG"] + [f"PV!{col}9" for col in "DEFGH"]
    expected = _eval_cells(workbook, cells)
    kwargs = input_kwargs(catalog, graph)
    got = call_compute(pkg, "post_grace", kwargs)
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx(_ragged_expected())
    assert _unwrap(call_compute(pkg, "result", kwargs)) == pytest.approx(20.0)


def test_choose_year_row_fail_closed_when_year_sets_differ_within_instrument(
    tmp_path: Path,
) -> None:
    sheets = _mcve_sheets()
    # E4 lists three years while the rest of IMF lists four — not a lookup axis.
    sheets["PV"]["E4"] = "=IF(E3=0,0,CHOOSE(E3,D5,E5,F5))"
    workbook = write_workbook(tmp_path / "a35_mismatch.xlsx", sheets)
    with pytest.raises(InvertedTreeExportError, match="more than two positions"):
        generate_inverted(workbook, _mcve_bindings())
