"""Issue 797 — factor shared formula prefixes across different input closures.

Exact-input shared runners (A14) only merge outputs with identical required
leaves. Two outputs that share a long internal prefix but differ by a private
input still duplicated that prefix in each `compute_*`. Helpers share source
without evaluating another output's private tail or caching across calls.
"""

from __future__ import annotations

import inspect
from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from tests.unit.exporter.inverted_tree.helpers import (
    all_param_names,
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    required_param_names,
    series_entry,
    write_workbook,
)

_PREFIX_LEN = 6


def _source(pkg, name, values):
    return pkg.Tensor.from_records(
        domain=getattr(pkg.data, name.upper() + "_DOMAIN"),
        records=zip(((2020,), (2021,), (2022,)), values, strict=True),
    )


def _observations(tensor):
    return tuple(tensor[year] for year in (2020, 2021, 2022))


def _prefix_workbook(tmp_path: Path) -> Path:
    engine: dict[str, object] = {"B1": 2020, "C1": 2021, "D1": 2022}
    engine["B2"] = "=Inputs!B2+1"
    engine["C2"] = "=Inputs!C2+1"
    engine["D2"] = "=Inputs!D2+1"
    for step in range(1, _PREFIX_LEN):
        src = step + 1
        dst = step + 2
        engine[f"B{dst}"] = f"=Engine!B{src}+1"
        engine[f"C{dst}"] = f"=Engine!C{src}+1"
        engine[f"D{dst}"] = f"=Engine!D{src}+1"
    last = _PREFIX_LEN + 1
    engine[f"B{last + 1}"] = f"=Engine!B{last}+1"
    engine[f"C{last + 1}"] = f"=Engine!C{last}+1"
    engine[f"D{last + 1}"] = f"=Engine!D{last}+1"
    engine[f"B{last + 2}"] = f"=Engine!B{last}+Inputs!B3"
    engine[f"C{last + 2}"] = f"=Engine!C{last}+Inputs!C3"
    engine[f"D{last + 2}"] = f"=Engine!D{last}+Inputs!D3"
    return write_workbook(
        tmp_path / "shared_prefix.xlsx",
        {
            "Inputs": {
                "B1": 2020,
                "C1": 2021,
                "D1": 2022,
                "B2": 10.0,
                "C2": 20.0,
                "D2": 30.0,
                "B3": 1.0,
                "C3": 2.0,
                "D3": 3.0,
            },
            "Engine": engine,
            "Outputs": {
                "B1": 2020,
                "C1": 2021,
                "D1": 2022,
                "B2": f"=Engine!B{last + 1}",
                "C2": f"=Engine!C{last + 1}",
                "D2": f"=Engine!D{last + 1}",
                "B3": f"=Engine!B{last + 2}",
                "C3": f"=Engine!C{last + 2}",
                "D3": f"=Engine!D{last + 2}",
            },
        },
    )


def _prefix_bindings() -> dict:
    steps = [
        series_entry(
            f"step_{index}",
            f"Engine!B{index + 2}:D{index + 2}",
            layout="series",
            direction="internal",
            header_row=1,
        )
        for index in range(_PREFIX_LEN)
    ]
    last = _PREFIX_LEN + 1
    return bindings_document(
        series_entry("values", "Inputs!B2:D2", layout="series", direction="input", header_row=1),
        series_entry("extra", "Inputs!B3:D3", layout="series", direction="input", header_row=1),
        *steps,
        series_entry(
            "first_tail",
            f"Engine!B{last + 1}:D{last + 1}",
            layout="series",
            direction="internal",
            header_row=1,
        ),
        series_entry(
            "second_tail",
            f"Engine!B{last + 2}:D{last + 2}",
            layout="series",
            direction="internal",
            header_row=1,
        ),
        series_entry(
            "first",
            "Outputs!B2:D2",
            layout="series",
            direction="output",
            header_row=1,
        ),
        series_entry(
            "second",
            "Outputs!B3:D3",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )


def test_long_shared_prefix_is_factored_once(tmp_path: Path) -> None:
    modules = generate_inverted(_prefix_workbook(tmp_path), _prefix_bindings())
    api = modules["api.py"]
    for index in range(_PREFIX_LEN):
        assert api.count(f"internals.step_{index}(") == 1
    first_src = api[api.index("def compute_first") :]
    second_src = api[api.index("def compute_second") :]
    first_fn = first_src.split("\ndef ")[0]
    second_fn = second_src.split("\ndef ")[0]
    assert "extra" not in first_fn
    assert "second_tail" not in first_fn
    assert "extra" in second_fn
    assert "first_tail" not in second_fn


def test_shared_prefix_parity_and_isolation(tmp_path: Path) -> None:
    workbook = _prefix_workbook(tmp_path)
    document = _prefix_bindings()
    modules = generate_inverted(workbook, document)
    pkg = load_package(modules, tmp_path, name="shared_prefix")
    assert required_param_names(pkg.compute_first) == ("values",)
    assert "extra" not in all_param_names(pkg.compute_first)
    assert set(required_param_names(pkg.compute_second)) == {"values", "extra"}

    values = (10.0, 20.0, 30.0)
    extra = (1.0, 2.0, 3.0)
    first = pkg.compute_first(values=_source(pkg, "values", values))
    second = pkg.compute_second(
        values=_source(pkg, "values", values), extra=_source(pkg, "extra", extra)
    )
    bump = _PREFIX_LEN + 1
    assert _observations(first) == pytest.approx(tuple(v + bump for v in values))
    assert _observations(second) == pytest.approx(
        tuple(v + _PREFIX_LEN + e for v, e in zip(values, extra, strict=True))
    )

    changed = pkg.compute_first(values=_source(pkg, "values", (11.0, 20.0, 30.0)))
    assert changed[2020] == pytest.approx(first[2020] + 1.0)
    assert _observations(pkg.compute_first(values=_source(pkg, "values", values))) == pytest.approx(
        _observations(first)
    )
    only_extra = pkg.compute_second(
        values=_source(pkg, "values", values), extra=_source(pkg, "extra", (9.0, 2.0, 3.0))
    )
    assert only_extra[2020] == pytest.approx(second[2020] + 8.0)
    assert _observations(
        pkg.compute_second(
            values=_source(pkg, "values", values), extra=_source(pkg, "extra", extra)
        )
    ) == pytest.approx(_observations(second))

    _catalog, _deps, graph = inverted_graph_parts(workbook, document)
    expected = FormulaEvaluator(graph).evaluate(
        [f"Outputs!{col}{row}" for row in (2, 3) for col in "BCD"]
    )
    assert _observations(first) == pytest.approx(
        tuple(expected[f"Outputs!{col}2"] for col in "BCD")
    )
    assert _observations(second) == pytest.approx(
        tuple(expected[f"Outputs!{col}3"] for col in "BCD")
    )


def test_shared_helper_is_not_a_cross_call_cache(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_prefix_workbook(tmp_path), _prefix_bindings()),
        tmp_path,
        name="shared_no_cache",
    )
    api_src = inspect.getsource(pkg.api)
    assert "lru_cache" not in api_src
    first_a = pkg.compute_first(values=_source(pkg, "values", (1.0, 2.0, 3.0)))
    first_b = pkg.compute_first(values=_source(pkg, "values", (4.0, 5.0, 6.0)))
    assert _observations(first_a) != pytest.approx(_observations(first_b))
    assert _observations(
        pkg.compute_first(values=_source(pkg, "values", (1.0, 2.0, 3.0)))
    ) == pytest.approx(_observations(first_a))


def test_issue_mcve_partial_overlap_keeps_signatures(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "issue797.xlsx",
        {
            "Data": {
                "B1": 2020,
                "C1": 2021,
                "D1": 2022,
                "B2": 2.0,
                "C2": 3.0,
                "D2": 4.0,
                "B3": 1.0,
                "C3": 2.0,
                "D3": 3.0,
                "B4": "=B2+B2",
                "C4": "=C2+C2",
                "D4": "=D2+D2",
                "B5": "=B4+1",
                "C5": "=C4+1",
                "D5": "=D4+1",
                "B6": "=B4+B3",
                "C6": "=C4+C3",
                "D6": "=D4+D3",
            }
        },
    )
    document = bindings_document(
        series_entry("values", "Data!B2:D2", layout="series", direction="input", header_row=1),
        series_entry("extra", "Data!B3:D3", layout="series", direction="input", header_row=1),
        series_entry("result", "Data!B4:D4", layout="series", direction="internal", header_row=1),
        series_entry("first", "Data!B5:D5", layout="series", direction="output", header_row=1),
        series_entry("second", "Data!B6:D6", layout="series", direction="output", header_row=1),
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="issue797_mcve")
    assert required_param_names(pkg.compute_first) == ("values",)
    assert "extra" not in all_param_names(pkg.compute_first)
    assert set(required_param_names(pkg.compute_second)) == {"extra", "values"}
    values = (2.0, 3.0, 4.0)
    extra = (1.0, 2.0, 3.0)
    assert _observations(pkg.compute_first(values=_source(pkg, "values", values))) == pytest.approx(
        (5.0, 7.0, 9.0)
    )
    assert _observations(
        pkg.compute_second(
            values=_source(pkg, "values", values), extra=_source(pkg, "extra", extra)
        )
    ) == pytest.approx((5.0, 8.0, 11.0))
    first_src = inspect.getsource(pkg.compute_first)
    assert "extra" not in first_src
    assert "second" not in first_src
