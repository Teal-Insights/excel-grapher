"""Tests for series-binding resolution audit helpers."""

from __future__ import annotations

from pathlib import Path
from typing import Any, cast

import fastpyxl
import pytest

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.audit import (
    DIRECTIONS,
    AuditFinding,
    audit_binding_resolutions,
    find_duplicate_internal_formula_cell_bindings,
    find_sparse_label_bind_issues,
    findings_from_resolution,
    format_audit_findings,
    unfilled_label_binds,
)
from excel_grapher.series_bindings.resolve import BindingDirection
from excel_grapher.series_bindings.workflow import validate_bindings_workbook
from tests.unit.series_bindings.authoring_helpers import (
    public_io_series,
    scalar_series,
    write_authoring_workbook,
    write_shards,
    years_internal_series,
)


def _load_context(workbook: Path, bindings_dir: Path):
    result = validate_bindings_workbook(workbook, bindings_dir)
    return result["graph"], result["bindings"]


def test_audit_directions_include_constant() -> None:
    assert DIRECTIONS == ("input", "output", "internal", "constant")


def test_findings_from_resolution_flags_partial_bind_and_empty_public() -> None:
    partial = findings_from_resolution(
        {
            "series_id": "gap_milestones",
            "ok": False,
            "requires_address": False,
            "leaves": [
                {
                    "address": "Sheet!I26",
                    "coordinates": {},
                    "key": {},
                    "record": {},
                }
            ],
            "issues": [
                {
                    "level": "error",
                    "code": "bind_resolution_failed",
                    "message": "column_header row 23: no source label",
                    "series_id": "gap_milestones",
                    "address": "Sheet!H26",
                }
            ],
        },
        direction="output",
        series={
            "id": "gap_milestones",
            "output": {"compute": {"name": "compute_gap"}},
        },
    )
    assert any(finding.code == "bind_resolution_failed" for finding in partial)
    assert any(finding.code == "partial_bind_failure" for finding in partial)

    empty = findings_from_resolution(
        {
            "series_id": "missing_output",
            "ok": True,
            "requires_address": False,
            "leaves": [],
            "issues": [
                {
                    "level": "warning",
                    "code": "no_resolved_cells",
                    "message": "No resolved output cells",
                    "series_id": "missing_output",
                    "address": None,
                }
            ],
        },
        direction="output",
        series={
            "id": "missing_output",
            "output": {"compute": {"name": "compute_missing"}},
        },
    )
    assert any(
        finding.code == "empty_public_series" and finding.severity == "warning" for finding in empty
    )

    empty_input = findings_from_resolution(
        {
            "series_id": "missing_input",
            "ok": True,
            "requires_address": False,
            "leaves": [],
            "issues": [
                {
                    "level": "warning",
                    "code": "no_resolved_cells",
                    "message": "No resolved input cells",
                    "series_id": "missing_input",
                    "address": None,
                }
            ],
        },
        direction="input",
        series={"id": "missing_input", "input": {}},
    )
    assert any(
        finding.code == "empty_public_series" and finding.severity == "warning"
        for finding in empty_input
    )


def test_unfilled_label_binds_skips_filled_and_non_label_binds() -> None:
    assert (
        unfilled_label_binds(
            {
                "structure": {
                    "dimensions": [
                        {
                            "bind": {
                                "kind": "column_header",
                                "header_row": 1,
                                "fill": True,
                            }
                        },
                        {"bind": {"kind": "data_cell", "read": "float"}},
                    ]
                }
            }
        )
        == []
    )
    binds = unfilled_label_binds(
        {
            "structure": {
                "dimensions": [
                    {"bind": {"kind": "column_header", "header_row": 1}},
                    {"bind": {"kind": "row_label", "label_column": "B"}},
                ]
            }
        }
    )
    assert len(binds) == 2


def test_find_sparse_label_bind_issues_short_circuits_without_workbook_io(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def _fail_expand(*_args: object, **_kwargs: object) -> list[str]:
        raise AssertionError("expand should not run when no unfilled label binds")

    def _fail_load(*_args: object, **_kwargs: object) -> object:
        raise AssertionError("workbook should not load when no unfilled label binds")

    monkeypatch.setattr(
        "excel_grapher.series_bindings.audit.expand_data_range_for_graph",
        _fail_expand,
    )
    monkeypatch.setattr(
        "excel_grapher.series_bindings.audit.fastpyxl.load_workbook",
        _fail_load,
    )
    findings = find_sparse_label_bind_issues(
        graph=cast(DependencyGraph, None),
        series={
            "id": "scalar_only",
            "data_range": "Engine!B2",
            "structure": {"dimensions": [{"bind": {"kind": "data_cell", "read": "float"}}]},
        },
        workbook_path="unused.xlsx",
        direction="internal",
    )
    assert findings == []


def test_sparse_years_without_fill_are_audit_errors(tmp_path: Path) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx", sparse_year=True)
    inputs, result_a, result_b = public_io_series()
    series = years_internal_series(fill=False)
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
        internals=[series],
    )
    graph, bindings = _load_context(workbook, bindings_dir)

    sparse = find_sparse_label_bind_issues(
        graph,
        series,
        workbook_path=workbook,
        direction="internal",
    )
    assert sparse
    assert sparse[0].code == "sparse_label_without_fill"
    assert "fill: true" in sparse[0].message

    report = audit_binding_resolutions(graph, bindings, workbook=workbook)
    assert not report.ok
    assert any(
        finding.series_id == "engine_years" and finding.code == "sparse_label_without_fill"
        for finding in report.findings
    )
    assert any(
        finding.series_id == "engine_years"
        and finding.code in {"bind_resolution_failed", "partial_bind_failure"}
        for finding in report.findings
    )
    rendered = "\n".join(format_audit_findings(report.findings))
    assert "engine_years" in rendered
    assert "sparse_label_without_fill" in rendered


def test_sparse_years_with_fill_are_audit_clean(tmp_path: Path) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx", sparse_year=True)
    inputs, result_a, result_b = public_io_series()
    # Gap-column pattern: data sits under a blank header; fill reads the group year.
    series = years_internal_series(data_range="Engine!C2", fill=True)
    series["layout"] = "scalar"
    series["key"] = ["TIME_PERIOD"]
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
        internals=[series],
    )
    graph, bindings = _load_context(workbook, bindings_dir)
    report = audit_binding_resolutions(graph, bindings, workbook=workbook)
    assert report.ok
    assert report.error_count == 0
    assert not any(finding.code == "sparse_label_without_fill" for finding in report.findings)


def test_audit_binding_resolutions_reuses_one_workbook_for_sparse_checks(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import excel_grapher.series_bindings.audit as audit_mod

    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx", sparse_year=True)
    inputs, result_a, result_b = public_io_series()
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
        internals=[years_internal_series(fill=False)],
    )
    graph, bindings = _load_context(workbook, bindings_dir)

    sparse_workbook_handles: list[fastpyxl.Workbook | None] = []
    real_sparse = audit_mod.find_sparse_label_bind_issues

    def tracking_sparse(
        graph: DependencyGraph,
        series: dict[str, Any],
        *,
        workbook_path: Path | str,
        direction: BindingDirection,
        workbook: fastpyxl.Workbook | None = None,
    ) -> list[AuditFinding]:
        sparse_workbook_handles.append(workbook)
        return real_sparse(
            graph,
            series,
            workbook_path=workbook_path,
            direction=direction,
            workbook=workbook,
        )

    monkeypatch.setattr(audit_mod, "find_sparse_label_bind_issues", tracking_sparse)
    report = audit_binding_resolutions(
        graph,
        bindings,
        workbook=workbook,
        directions=("internal",),
    )
    assert not report.ok
    assert sparse_workbook_handles
    assert all(handle is not None for handle in sparse_workbook_handles)
    assert len({id(handle) for handle in sparse_workbook_handles}) == 1


def test_duplicate_internal_audit_expands_list_data_range(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import excel_grapher.series_bindings.audit as audit_module

    class _FakeNode:
        is_leaf = False
        has_formula = True

    class _FakeGraph:
        def get_node(self, _address: str) -> _FakeNode:
            return _FakeNode()

    monkeypatch.setattr(
        audit_module,
        "expand_data_range_for_graph",
        lambda _graph, data_range, workbook=None: (
            ("Engine!B2",)
            if data_range == "Engine!B2"
            else (("Engine!C2",) if data_range == "Engine!C2" else ())
        ),
    )
    bindings = cast(
        Any,
        {
            "series": [
                {
                    "id": "left_span",
                    "data_range": ["Engine!B2", "Engine!C2"],
                    "internal": {},
                },
                {
                    "id": "right_cell",
                    "data_range": "Engine!B2",
                    "internal": {},
                },
            ]
        },
    )
    findings = find_duplicate_internal_formula_cell_bindings(
        cast(DependencyGraph, _FakeGraph()),
        bindings,
        workbook_path=Path("unused.xlsx"),
    )
    assert findings
    assert findings[0].code == "duplicate_internal_cell_binding"
    assert findings[0].address == "Engine!B2"
    assert "left_span" in findings[0].message
    assert "right_cell" in findings[0].message


def test_duplicate_internal_audit_respects_exclude_rows(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import excel_grapher.series_bindings.audit as audit_module

    class _FakeNode:
        is_leaf = False
        has_formula = True

    class _FakeGraph:
        def get_node(self, _address: str) -> _FakeNode:
            return _FakeNode()

    monkeypatch.setattr(
        audit_module,
        "expand_data_range_for_graph",
        lambda _graph, data_range, workbook=None: (
            ("Engine!B2", "Engine!B3", "Engine!B4", "Engine!B5")
            if data_range == "Engine!B2:B5"
            else (("Engine!B4",) if data_range == "Engine!B4" else ())
        ),
    )
    bindings = cast(
        Any,
        {
            "series": [
                {
                    "id": "row_series",
                    "data_range": "Engine!B4",
                    "internal": {},
                },
                {
                    "id": "column_series",
                    "data_range": "Engine!B2:B5",
                    "internal": {},
                    "exclude_rows": [4],
                },
            ]
        },
    )
    findings = find_duplicate_internal_formula_cell_bindings(
        cast(DependencyGraph, _FakeGraph()),
        bindings,
        workbook_path=Path("unused.xlsx"),
    )
    assert findings == []


def test_duplicate_internal_formula_ownership_is_audit_error(tmp_path: Path) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx")
    inputs, result_a, result_b = public_io_series()
    first = years_internal_series(series_id="engine_b2", data_range="Engine!B2")
    first["layout"] = "scalar"
    first["key"] = []
    first["structure"]["dimensions"] = []
    duplicate = dict(first)
    duplicate["id"] = "engine_b2_duplicate"
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
        internals=[first, duplicate],
    )
    graph, bindings = _load_context(workbook, bindings_dir)
    findings = find_duplicate_internal_formula_cell_bindings(
        graph, bindings, workbook_path=workbook
    )
    assert findings
    assert findings[0].code == "duplicate_internal_cell_binding"
    assert findings[0].address == "Engine!B2"
    assert "engine_b2" in findings[0].message
    assert "engine_b2_duplicate" in findings[0].message

    report = audit_binding_resolutions(graph, bindings, workbook=workbook)
    assert not report.ok
    assert any(finding.code == "duplicate_internal_cell_binding" for finding in report.findings)


def test_output_and_internal_formula_overlap_is_audit_error(tmp_path: Path) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx")
    inputs, result_a, result_b = public_io_series()
    overlapping_internal = scalar_series(
        "engine_as_internal",
        "Outputs!B1",
        direction="internal",
    )
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
        internals=[overlapping_internal],
    )
    graph, bindings = _load_context(workbook, bindings_dir)
    report = audit_binding_resolutions(graph, bindings, workbook=workbook)
    assert not report.ok
    assert any(
        finding.code == "duplicate_formula_cell_binding" and finding.address == "Outputs!B1"
        for finding in report.findings
    )


def test_audit_clean_for_filled_public_and_internal_shards(tmp_path: Path) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx")
    inputs, result_a, result_b = public_io_series()
    constant = scalar_series("input_bias", "Inputs!B1", direction="constant")
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
        internals=[years_internal_series(fill=False)],
        constants=[constant],
    )
    graph, bindings = _load_context(workbook, bindings_dir)
    report = audit_binding_resolutions(graph, bindings, workbook=workbook)
    assert report.ok
    assert report.error_count == 0


def test_empty_public_output_is_audit_warning(tmp_path: Path) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx")
    inputs, result_a, result_b = public_io_series()
    missing = scalar_series("missing_output", "Engine!B2", direction="output")
    missing["exclude_rows"] = [2]
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b, missing],
    )
    graph, bindings = _load_context(workbook, bindings_dir)
    report = audit_binding_resolutions(graph, bindings, workbook=workbook)
    assert report.ok
    assert any(
        finding.code == "empty_public_series"
        and finding.severity == "warning"
        and finding.series_id == "missing_output"
        for finding in report.findings
    )
