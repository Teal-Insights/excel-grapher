"""MATCH of INDEX((range<>0),0) over sparse columns (#713).

Published tenors have formulas; intervening years are empty and off-graph.
Those holes must stay as positional `None` in MATCH/INDEX windows so the
next-non-blank index does not shift. Declaring them in `BLANK_RANGES` is
not required. On-graph formula cells without a series still fail closed.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any

import pytest

from excel_grapher.core.address_keys import as_canonical
from excel_grapher.core.types import XlError
from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.deps import resolve_positional_range
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import validate_bindings_document, validate_series_bindings
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    call_compute,
    generate_inverted,
    input_kwargs,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)

_NEXT_NONBLANK = "=INDEX(N11:N$13,MATCH(TRUE,INDEX((N11:N$13<>0),0),0))"


def _sparse_quoted_sheets() -> dict[str, dict[str, object]]:
    return {
        "Engine": {
            "M10": 1,
            "M11": 2,
            "M12": 3,
            "M13": 4,
            "N10": "=0.01",
            "N12": "=0.03",
            "O10": _NEXT_NONBLANK,
        }
    }


def _quoted_series(*, exclude_rows: Sequence[int] | None = None) -> dict[str, Any]:
    entry = series_entry(
        "quoted",
        "Engine!N10:N13",
        layout="series",
        direction="internal",
        label_column="M",
        key_read="int",
    )
    if exclude_rows is not None:
        entry["exclude_rows"] = list(exclude_rows)
    return entry


def _sparse_quoted_bindings(*, exclude_rows: Sequence[int] | None = (11, 13)) -> dict[str, Any]:
    return bindings_document(
        series_entry(
            "tenor",
            "Engine!M10:M13",
            layout="series",
            direction="constant",
            dtype="int",
            label_column="M",
            key_read="int",
        ),
        _quoted_series(exclude_rows=exclude_rows),
        series_entry("interp", "Engine!O10", layout="scalar", direction="output"),
    )


def _norm_measure(value: object) -> object:
    if isinstance(value, XlError):
        return value.value
    return value


def _export_interp(
    tmp_path: Path,
    *,
    stem: str,
    document: dict[str, Any] | None = None,
    sheets: Mapping[str, Mapping[str, object]] | None = None,
    blank_ranges: Sequence[str] | None = None,
) -> tuple[object, object, dict[str, str]]:
    bound = document or _sparse_quoted_bindings()
    workbook = write_workbook(tmp_path / f"{stem}.xlsx", sheets or _sparse_quoted_sheets())
    catalog, _deps, graph = inverted_graph_parts(workbook, bound, blank_ranges=blank_ranges)
    modules = generate_inverted(workbook, bound, blank_ranges=blank_ranges)
    pkg = load_package(modules, tmp_path, name=stem)
    with FormulaEvaluator(graph, blank_ranges=blank_ranges) as ev:
        expected = ev.evaluate(["Engine!O10"])["Engine!O10"]
    got = call_compute(pkg, "interp", input_kwargs(catalog, graph))
    if isinstance(got, tuple):
        assert len(got) == 1
        got = got[0]
    return _norm_measure(got), _norm_measure(expected), modules


def test_resolve_positional_range_keeps_off_graph_holes(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "positional.xlsx", _sparse_quoted_sheets())
    catalog, _deps, graph = inverted_graph_parts(workbook, _sparse_quoted_bindings())
    cells, missing = resolve_positional_range(
        [
            as_canonical("Engine!N11"),
            as_canonical("Engine!N12"),
            as_canonical("Engine!N13"),
        ],
        catalog,
        graph=graph,
    )
    assert missing == ()
    assert [cell.blank for cell in cells] == [True, False, True]
    assert [cell.series_id for cell in cells] == [None, "quoted", None]
    assert cells[1].catalog_index == 1


def test_coverage_passes_before_export(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "coverage.xlsx", _sparse_quoted_sheets())
    bindings = validate_bindings_document(_sparse_quoted_bindings())
    graph = create_dependency_graph(
        workbook,
        all_series_targets(bindings, workbook=workbook),
        load_values=True,
        use_cached_dynamic_refs=True,
        capture_dependency_provenance=True,
    )
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert report["ok"] is True


def test_mcve_exports_next_nonblank_without_blank_ranges(tmp_path: Path) -> None:
    got, want, modules = _export_interp(tmp_path, stem="sparse_nb")
    assert got == pytest.approx(want)
    assert got == pytest.approx(0.03)
    internals = modules["internals.py"]
    assert "None" in internals
    assert "xl_match(" in internals
    assert "xl_index(" in internals
    assert "xl_ne(" in internals


def test_declared_blanks_are_optional_for_off_graph_holes(tmp_path: Path) -> None:
    got, want, _modules = _export_interp(
        tmp_path,
        stem="sparse_declared",
        blank_ranges=("Engine!N11", "Engine!N13"),
    )
    assert got == pytest.approx(want)
    assert got == pytest.approx(0.03)


def test_omit_exclude_rows_still_keeps_hole_positions(tmp_path: Path) -> None:
    got, want, _modules = _export_interp(
        tmp_path,
        stem="sparse_no_excl",
        document=_sparse_quoted_bindings(exclude_rows=None),
    )
    assert got == pytest.approx(want)
    assert got == pytest.approx(0.03)


def test_shifted_holes_would_return_blank_not_next_quote(tmp_path: Path) -> None:
    """Dropping N11 from the MATCH vector would INDEX into the hole."""
    got, want, modules = _export_interp(tmp_path, stem="sparse_shift")
    assert got == pytest.approx(want)
    assert got == pytest.approx(0.03)
    internals = modules["internals.py"]
    none_before_quote = internals.find("None") < internals.find("quoted[")
    assert none_before_quote


def test_on_graph_unbound_formula_in_window_fails_closed(tmp_path: Path) -> None:
    sheets = _sparse_quoted_sheets()
    sheets["Engine"]["N11"] = "=0.02"
    with pytest.raises(InvertedTreeExportError) as exc:
        generate_inverted(
            write_workbook(tmp_path / "unbound_formula.xlsx", sheets),
            _sparse_quoted_bindings(),
        )
    message = str(exc.value)
    assert "Engine!O10" in message
    assert "Engine!N11" in message
