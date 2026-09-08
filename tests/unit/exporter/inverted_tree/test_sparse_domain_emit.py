"""Issue 799 — compact mostly regular sparse multidimensional domains."""

from __future__ import annotations

import timeit
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, KeyPoint, Statement
from excel_grapher.exporter.inverted_tree.domains import (
    DomainEmitPlan,
    domain_const_name,
    plan_domain_emission,
    series_domain_points,
)
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    input_kwargs,
    inverted_graph_parts,
    load_package,
    make_catalog,
    write_workbook,
)


def _country_dim() -> dict[str, Any]:
    return {
        "id": "COUNTRY",
        "concept": "COUNTRY",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "row_label", "label_column": "A", "read": "string"},
    }


def _time_dim(*, header_row: int = 1) -> dict[str, Any]:
    return {
        "id": "TIME_PERIOD",
        "concept": "TIME_PERIOD",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "column_header", "header_row": header_row, "read": "int"},
    }


def _measure() -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }


def _matrix_entry(
    series_id: str,
    data_range: str | list[str],
    *,
    direction: str,
    header_row: int = 1,
) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "id": series_id,
        "sheet": "Data",
        "data_range": data_range,
        "layout": "matrix",
        "key": ["COUNTRY", "TIME_PERIOD"],
        "structure": {
            "measure": _measure(),
            "dimensions": [_country_dim(), _time_dim(header_row=header_row)],
        },
    }
    if direction == "input":
        entry["input"] = {"setter": {"name": f"set_{series_id}"}}
    else:
        entry["output"] = {"compute": {"name": f"compute_{series_id}"}}
    return entry


def _result_entry(size: int, *, last_row: int | None = None) -> dict[str, Any]:
    end = last_row if last_row is not None else size + 1
    return {
        "id": "result",
        "sheet": "Data",
        "data_range": f"Data!E2:E{end}",
        "layout": "series",
        "output": {"compute": {"name": "compute_result"}},
        "key": ["COUNTRY"],
        "structure": {
            "measure": _measure(),
            "dimensions": [_country_dim()],
        },
    }


def _missing_corner_workbook(tmp_path: Path, size: int) -> Path:
    cells: dict[str, object] = {"B1": 2020, "C1": 2021}
    last = size + 1
    for row in range(2, last + 1):
        cells[f"A{row}"] = f"Country {row}"
        cells[f"B{row}"] = float(row)
        cells[f"C{row}"] = float(row * 2)
        cells[f"E{row}"] = f"=B{row}+C{row}"
    cells[f"E{last}"] = f"=B{last}"
    return write_workbook(tmp_path / f"missing_corner_{size}.xlsx", {"Data": cells})


def _missing_corner_bindings(size: int) -> dict[str, Any]:
    last = size + 1
    return bindings_document(
        _matrix_entry("values", [f"Data!B2:C{last - 1}", f"Data!B{last}"], direction="input"),
        _result_entry(size),
    )


def _eval_series_domain(plan: DomainEmitPlan, series_id: str) -> tuple[object, ...]:
    """Evaluate a planned `__domain__` expression against interned constants."""
    ns: dict[str, object] = {
        domain_const_name(field): values for field, values in plan.field_domains.items()
    }
    ns.update({"tuple": tuple, "enumerate": enumerate})
    for name, values in plan.interned:
        source = plan.interned_source.get(name)
        ns[name] = values if source is None else eval(source, ns)
    data = SimpleNamespace(**{name: ns[name] for name in ns})
    got = eval(plan.series_expr[series_id], {**ns, "data": data})
    if not isinstance(got, tuple):
        raise TypeError(f"expected a domain tuple, got {type(got).__name__}")
    return got


def _bound_matrix(
    series_id: str,
    points: tuple[tuple[object, ...], ...],
    *,
    keys: tuple[str, ...] = ("COUNTRY", "TIME_PERIOD"),
) -> BoundSeries:
    cells = tuple(f"Data!A{index + 1}" for index in range(len(points)))
    domain = tuple(KeyPoint(tuple(zip(keys, point, strict=True))) for point in points)
    return BoundSeries(
        series_id=series_id,
        layout="matrix",
        direction="input",
        cells=cells,
        key_fields=keys,
        dtype="float",
        compute_name=None,
        raw={},
        domain=domain,
        statements=(Statement(series_id, series_id, None, 0, len(cells), cells, domain),),
    )


def _catalog_for(*series: BoundSeries):
    mapping = {item.series_id: item for item in series}
    order = tuple(item.series_id for item in series)
    address_to_id = {cell: item.series_id for item in series for cell in item.cells}
    return make_catalog(mapping, order, address_to_id)


def _interned_source_for(plan: DomainEmitPlan, series_id: str) -> str | None:
    expr = plan.series_expr[series_id]
    prefix = "data."
    if not expr.startswith(prefix):
        return None
    name = expr[len(prefix) :]
    return plan.interned_source.get(name)


def test_full_product_still_uses_axis_comprehension() -> None:
    countries = ("France", "Kenya")
    years = (2020, 2021)
    points = tuple((country, year) for country in countries for year in years)
    plan = plan_domain_emission(_catalog_for(_bound_matrix("values", points)))
    expr = plan.series_expr["values"]
    assert "for country in data.COUNTRY_DOMAIN" in expr
    assert "for period in data.TIME_PERIOD_DOMAIN" in expr
    assert plan.interned == ()
    assert _eval_series_domain(plan, "values") == points
    assert (
        series_domain_points(_catalog_for(_bound_matrix("values", points)).get("values")) == points
    )


def test_missing_corner_reuses_axis_constants_and_preserves_order() -> None:
    countries = tuple(f"Country {i}" for i in range(6))
    years = (2020, 2021)
    full = tuple((country, year) for country in countries for year in years)
    points = full[:-1]
    plan = plan_domain_emission(_catalog_for(_bound_matrix("values", points)))
    source = _interned_source_for(plan, "values")
    assert source is not None
    assert "Country " not in source
    assert "COUNTRY_DOMAIN" in source
    assert "TIME_PERIOD_DOMAIN" in source
    assert _eval_series_domain(plan, "values") == points
    assert plan.series_key["values"] == ("COUNTRY", "TIME_PERIOD")


def test_triangular_domain_uses_enumerate_prefix() -> None:
    countries = ("France", "Kenya", "Norway", "Peru", "Spain", "Tunisia")
    years = (2020, 2021, 2022, 2023, 2024, 2025)
    points = tuple(
        (country, year) for i, country in enumerate(countries) for year in years[: i + 1]
    )
    plan = plan_domain_emission(_catalog_for(_bound_matrix("values", points)))
    source = _interned_source_for(plan, "values")
    assert source is not None
    assert "enumerate" in source
    assert _eval_series_domain(plan, "values") == points


def test_disjoint_rectangles_concatenate_products() -> None:
    countries = (
        "Austria",
        "Belgium",
        "Canada",
        "Denmark",
        "Estonia",
        "Finland",
        "Germany",
        "Hungary",
    )
    years = (2020, 2021, 2022, 2023)
    first = tuple((country, year) for country in countries[:4] for year in years[:2])
    second = tuple((country, year) for country in countries[4:] for year in years[2:])
    points = first + second
    plan = plan_domain_emission(_catalog_for(_bound_matrix("values", points)))
    source = _interned_source_for(plan, "values")
    assert source is not None
    assert source.count("tuple(") >= 2 or "*" in source
    assert _eval_series_domain(plan, "values") == points


def test_irregular_exceptions_do_not_invent_or_reorder_keys() -> None:
    points = (("A", 2020), ("C", 2023), ("B", 2021), ("A", 2022))
    plan = plan_domain_emission(_catalog_for(_bound_matrix("values", points)))
    got = _eval_series_domain(plan, "values")
    assert got == points
    assert ("A", 2021) not in got
    assert ("B", 2020) not in got


def test_three_key_missing_corner_preserves_product_order() -> None:
    keys = ("COUNTRY", "SCENARIO", "TIME_PERIOD")
    countries = ("France", "Kenya", "Norway")
    scenarios = ("Baseline", "Shock")
    years = (2020, 2021, 2022)
    full = tuple(
        (country, scenario, year)
        for country in countries
        for scenario in scenarios
        for year in years
    )
    points = full[:-1]
    plan = plan_domain_emission(_catalog_for(_bound_matrix("values", points, keys=keys)))
    source = _interned_source_for(plan, "values")
    assert source is not None
    assert "COUNTRY_DOMAIN" in source
    assert "SCENARIO_DOMAIN" in source
    assert "TIME_PERIOD_DOMAIN" in source
    assert _eval_series_domain(plan, "values") == points


def test_missing_corner_generated_domain_matches_public_output(tmp_path: Path) -> None:
    size = 8
    workbook = _missing_corner_workbook(tmp_path, size)
    document = _missing_corner_bindings(size)
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    modules = generate_inverted(workbook, document)
    interned = [line for line in modules["data.py"].splitlines() if line.startswith("_DOMAIN_")]
    assert interned
    assert all("Country " not in line for line in interned)
    assert "COUNTRY_DOMAIN" in interned[0]
    pkg = load_package(modules, tmp_path, name="sparse_corner")
    expected = series_domain_points(catalog.get("values"))
    assert expected == pkg.data._DOMAIN_0
    assert pkg.compute_result.__key__ == ("COUNTRY",)
    assert pkg.compute_result.__domain__ == tuple(f"Country {row}" for row in range(2, size + 2))
    cells = [f"Data!E{row}" for row in range(2, size + 2)]
    evaluated = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = pkg.compute_result(**input_kwargs(catalog, graph))
    assert got == pytest.approx(tuple(evaluated[cell] for cell in cells))
    records = pkg.as_records(pkg.compute_result, got)
    assert [row["COUNTRY"] for row in records] == list(pkg.compute_result.__domain__)


def test_sparse_domain_source_and_import_scale(tmp_path: Path) -> None:
    sizes = (10, 40)
    data_sizes: list[int] = []
    domain_sizes: list[int] = []
    compile_times: list[float] = []
    for size in sizes:
        workbook = _missing_corner_workbook(tmp_path, size)
        document = _missing_corner_bindings(size)
        modules = generate_inverted(workbook, document)
        data_py = modules["data.py"]
        interned_line = next(line for line in data_py.splitlines() if line.startswith("_DOMAIN_"))
        data_sizes.append(len(data_py))
        domain_sizes.append(len(interned_line))
        compile_times.append(
            timeit.timeit(lambda src=data_py: compile(src, "<data>", "exec"), number=20)
        )
        assert "Country " not in interned_line
    assert domain_sizes[-1] / domain_sizes[0] < 2.0
    assert data_sizes[-1] / data_sizes[0] < 3.0
    assert compile_times[-1] / compile_times[0] < 8.0
