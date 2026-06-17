"""Reusable accuracy helpers for series-binding integration tests."""

from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal, cast

import pytest
from fastpyxl import load_workbook

from excel_grapher.core.address_keys import parse_address
from excel_grapher.exporter import CodeGenerator
from excel_grapher.grapher import DependencyGraph, create_dependency_graph
from excel_grapher.series_bindings import resolve_series_binding, validate_bindings_workbook
from excel_grapher.series_bindings.types import LeafResolution, WorkbookSeriesBindings
from excel_grapher.series_bindings.workflow import all_series_targets, compute_names, setter_names

BindingDirection = Literal["input", "output"]


@dataclass(frozen=True)
class SeriesSpotCheck:
    """Optional resolution spot-check for one binding series."""

    series_id: str
    direction: BindingDirection = "input"
    leaf_count: int | None = None
    sample_key: dict[str, str] | None = None
    sample_value: object | None = None
    unique_key_fields: tuple[str, ...] | None = None


@dataclass(frozen=True)
class DownstreamUpdateCase:
    """Setter write that should change a downstream compute observation."""

    setter_name: str
    setter_records: tuple[dict[str, object], ...]
    compute_name: str
    record_key: dict[str, str]
    expected_obs_value: float


@dataclass(frozen=True)
class BindingsAccuracyCase:
    """Workbook + bindings shard exercised by generic accuracy tests."""

    name: str
    workbook: Path
    bindings_path: Path
    setter_name_prefix: str | None = None
    compute_name_prefix: str | None = None
    expected_setter_count: int | None = None
    expected_compute_count: int | None = None
    series_checks: tuple[SeriesSpotCheck, ...] = ()
    downstream_update: DownstreamUpdateCase | None = None


def read_workbook_cell(workbook: Path, address: str) -> object:
    """Return a cached ``data_only`` cell value from ``workbook``."""
    sheet, coord = parse_address(address)
    wb = load_workbook(workbook, data_only=True, read_only=True)
    try:
        return wb[sheet][coord].value
    finally:
        wb.close()


def series_by_id(bindings: WorkbookSeriesBindings, series_id: str) -> dict[str, Any]:
    """Return the ``series[]`` entry with ``id == series_id``."""
    for series in bindings["series"]:
        if series["id"] == series_id:
            return series
    raise KeyError(series_id)


def resolve_leaves(
    graph: DependencyGraph,
    workbook: Path,
    bindings: WorkbookSeriesBindings,
    series_id: str,
    *,
    direction: BindingDirection,
) -> list[LeafResolution]:
    """Resolve one binding series and return its leaves."""
    series = series_by_id(bindings, series_id)
    resolved = resolve_series_binding(
        graph,
        workbook,
        series,
        concept_scheme=bindings.get("concept_scheme"),
        direction=direction,
    )
    if not resolved["ok"]:
        raise AssertionError(f"{series_id} resolution failed: {resolved['issues']}")
    return resolved["leaves"]


def assert_obs_values_match_workbook(
    workbook: Path,
    leaves: Sequence[LeafResolution],
    *,
    rel: float = 1e-9,
) -> None:
    """Assert each leaf ``OBS_VALUE`` matches the workbook cache at ``address``."""
    for leaf in leaves:
        expected = read_workbook_cell(workbook, leaf["address"])
        actual = leaf["record"]["OBS_VALUE"]
        if isinstance(expected, float) or isinstance(actual, float):
            assert actual == pytest.approx(expected, rel=rel, abs=rel)
        else:
            assert actual == expected


def assert_unique_keys(
    leaves: Sequence[LeafResolution],
    *key_fields: str,
) -> None:
    """Assert leaves have distinct composite keys over ``key_fields``."""
    keys = [tuple(leaf["key"][field] for field in key_fields) for leaf in leaves]
    assert len(keys) == len(set(keys))


def leaf_matching(leaves: Sequence[LeafResolution], key: dict[str, str]) -> LeafResolution:
    """Return the first leaf whose ``key`` equals ``key``."""
    return next(leaf for leaf in leaves if leaf["key"] == key)


def assert_bindings_validate(case: BindingsAccuracyCase) -> dict[str, Any]:
    """Validate bindings against the workbook and assert API name conventions."""
    result = validate_bindings_workbook(case.workbook, case.bindings_path)
    report = result["report"]
    assert report["ok"] is True
    assert not any(issue["level"] == "error" for issue in report["issues"])

    bindings = result["bindings"]
    setters = setter_names(bindings)
    computes = compute_names(bindings)

    if case.expected_setter_count is not None:
        assert len(setters) == case.expected_setter_count
    if case.expected_compute_count is not None:
        assert len(computes) == case.expected_compute_count
    if case.setter_name_prefix is not None:
        assert all(name.startswith(f"set_{case.setter_name_prefix}_") for name in setters)
    if case.compute_name_prefix is not None:
        assert all(name.startswith(f"compute_{case.compute_name_prefix}_") for name in computes)
    return result


def build_dependency_graph(
    workbook: Path,
    bindings: WorkbookSeriesBindings,
) -> DependencyGraph:
    """Build a dependency graph for all binding targets in ``bindings``."""
    targets = all_series_targets(bindings, workbook=workbook)
    return create_dependency_graph(workbook, targets, load_values=True)


def generate_bindings_namespace(
    graph: DependencyGraph,
    workbook: Path,
    bindings: WorkbookSeriesBindings,
) -> dict[str, object]:
    """Generate binding setters/computes and return an executable namespace."""
    targets = all_series_targets(bindings, workbook=workbook)
    with CodeGenerator(graph) as gen:
        code = gen.generate(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )
    namespace: dict[str, object] = {}
    exec(code, namespace)
    return namespace


def compute_records(
    namespace: dict[str, object],
    compute_name: str,
    *,
    ctx: Any | None = None,
) -> list[dict[str, object]]:
    """Call a generated ``compute_*`` function."""
    make_context = cast(Callable[[], Any], namespace["make_context"])
    compute = cast(Callable[..., list[dict[str, object]]], namespace[compute_name])
    if ctx is None:
        ctx = make_context()
    return compute(ctx=ctx)


def apply_setter(
    namespace: dict[str, object],
    setter_name: str,
    records: Sequence[dict[str, object]],
    *,
    ctx: Any | None = None,
) -> Any:
    """Call a generated ``set_*`` function and return the context used."""
    make_context = cast(Callable[[], Any], namespace["make_context"])
    setter = cast(Callable[[Any, list[dict[str, object]]], None], namespace[setter_name])
    if ctx is None:
        ctx = make_context()
    setter(ctx, list(records))
    return ctx


def assert_compute_records_match_workbook(
    graph: DependencyGraph,
    workbook: Path,
    bindings: WorkbookSeriesBindings,
    series_id: str,
    records: Sequence[dict[str, object]],
) -> None:
    """Assert compute records match cached workbook values for one output series."""
    leaves = resolve_leaves(graph, workbook, bindings, series_id, direction="output")
    assert len(records) == len(leaves)
    for leaf in leaves:
        record = next(
            record
            for record in records
            if all(record.get(key) == value for key, value in leaf["key"].items())
        )
        expected = read_workbook_cell(workbook, leaf["address"])
        actual = record["OBS_VALUE"]
        if isinstance(expected, float) or isinstance(actual, float):
            assert actual == pytest.approx(expected, rel=1e-9, abs=1e-9)
        else:
            assert actual == expected


def assert_all_compute_functions_match_workbook(
    graph: DependencyGraph,
    workbook: Path,
    bindings: WorkbookSeriesBindings,
    namespace: dict[str, object],
) -> None:
    """Assert every declared compute function matches the workbook cache."""
    for compute_name in compute_names(bindings):
        series_id = compute_name.removeprefix("compute_")
        records = compute_records(namespace, compute_name)
        assert_compute_records_match_workbook(
            graph,
            workbook,
            bindings,
            series_id,
            records,
        )


def assert_series_spot_check(
    graph: DependencyGraph,
    workbook: Path,
    bindings: WorkbookSeriesBindings,
    check: SeriesSpotCheck,
) -> None:
    """Run leaf-count, uniqueness, sample, and workbook parity checks for one series."""
    leaves = resolve_leaves(
        graph,
        workbook,
        bindings,
        check.series_id,
        direction=check.direction,
    )
    if check.leaf_count is not None:
        assert len(leaves) == check.leaf_count
    if check.unique_key_fields:
        assert_unique_keys(leaves, *check.unique_key_fields)
    if check.sample_key is not None:
        sample = leaf_matching(leaves, check.sample_key)
        if check.sample_value is not None:
            assert sample["record"]["OBS_VALUE"] == check.sample_value
    if check.direction == "input":
        assert_obs_values_match_workbook(workbook, leaves)


def assert_shared_game_log_columns(
    graph: DependencyGraph,
    workbook: Path,
    bindings: WorkbookSeriesBindings,
    *,
    date_series_id: str,
    result_series_id: str,
) -> None:
    """Assert standard week/date/result columns shared by ffv3 player sheets."""
    date_leaves = resolve_leaves(graph, workbook, bindings, date_series_id, direction="input")
    result_leaves = resolve_leaves(graph, workbook, bindings, result_series_id, direction="input")

    by_week = {leaf["key"]["GAME_WEEK"]: leaf for leaf in date_leaves if leaf["key"]["GAME_WEEK"]}
    assert by_week["W1"]["record"]["OBS_VALUE"] == "Sep 7"
    assert by_week["PO3"]["record"]["OBS_VALUE"] == "Jan 25"

    by_week_result = {
        leaf["key"]["GAME_WEEK"]: leaf for leaf in result_leaves if leaf["key"]["GAME_WEEK"]
    }
    assert by_week_result["W1"]["record"]["OBS_VALUE"] == "W 14-9"
    assert by_week_result["PO1"]["record"]["OBS_VALUE"] == "W 34-31 (WC)"

    assert_obs_values_match_workbook(
        workbook, [leaf for leaf in date_leaves if leaf["key"]["GAME_WEEK"]]
    )
    assert_obs_values_match_workbook(
        workbook, [leaf for leaf in result_leaves if leaf["key"]["GAME_WEEK"]]
    )


def assert_downstream_update(
    namespace: dict[str, object],
    update: DownstreamUpdateCase,
) -> None:
    """Apply a setter and assert a downstream compute observation changed."""
    ctx = apply_setter(namespace, update.setter_name, update.setter_records)
    records = compute_records(namespace, update.compute_name, ctx=ctx)
    record = next(
        record
        for record in records
        if all(record.get(key) == value for key, value in update.record_key.items())
    )
    assert record["OBS_VALUE"] == pytest.approx(update.expected_obs_value)


def run_bindings_accuracy_case(case: BindingsAccuracyCase) -> None:
    """Run the full generic accuracy suite for one ``BindingsAccuracyCase``."""
    result = assert_bindings_validate(case)
    bindings = result["bindings"]
    graph = build_dependency_graph(case.workbook, bindings)
    namespace = generate_bindings_namespace(graph, case.workbook, bindings)

    slug = case.setter_name_prefix or case.name
    assert_shared_game_log_columns(
        graph,
        case.workbook,
        bindings,
        date_series_id=f"{slug}_game_date",
        result_series_id=f"{slug}_game_result",
    )

    for check in case.series_checks:
        assert_series_spot_check(graph, case.workbook, bindings, check)

    assert_all_compute_functions_match_workbook(graph, case.workbook, bindings, namespace)

    if case.downstream_update is not None:
        assert_downstream_update(namespace, case.downstream_update)
