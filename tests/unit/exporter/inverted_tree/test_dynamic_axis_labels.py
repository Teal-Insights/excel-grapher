"""Runtime axis labels keyed by the labeller all the way down (#841)."""

from __future__ import annotations

import ast
import re
from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from tests.unit.exporter.inverted_tree.helpers import (
    all_param_names,
    assert_package_matches_evaluator,
    bindings_document,
    generate_inverted,
    input_field_names,
    inverted_graph_parts,
    invoke_public_compute,
    load_package,
    named_input_kwargs,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a10_other_series_lag import (
    _lag_bindings,
    _lag_workbook,
)

_YEARS = (2024, 2025, 2026)


def _labelled_workbook(tmp_path: Path, *, first: int = 2024, shock_year: int = 2025) -> Path:
    years = (first, first + 1, first + 2)
    return write_workbook(
        tmp_path / f"labelled_{first}_{shock_year}.xlsx",
        {
            "Inputs": {
                "A1": first,
                "B1": shock_year,
                "A2": years[0],
                "B2": years[1],
                "C2": years[2],
                "A3": 3.5,
                "B3": 4.0,
                "C3": 4.5,
            },
            "Engine": {
                "A1": "=Inputs!A1",
                "B1": "=A1+1",
                "C1": "=B1+1",
                "A2": "=Inputs!A3",
                "B2": "=A2+1",
                "C2": "=B2+1",
                "A3": "=IF(A1>=Inputs!B1,1,0)",
                "B3": "=IF(B1>=Inputs!B1,1,0)",
                "C3": "=IF(C1>=Inputs!B1,1,0)",
            },
        },
    )


def _labelled_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("first_year", "Inputs!A1", direction="input", dtype="int"),
        series_entry("shock_year", "Inputs!B1", direction="input", dtype="int"),
        series_entry(
            "growth",
            "Inputs!A3:C3",
            layout="series",
            direction="input",
            header_row=2,
        ),
        series_entry(
            "year_labels",
            "Engine!A1:C1",
            layout="series",
            direction="internal",
            dtype="int",
            header_row=1,
            axis_labels="TIME_PERIOD",
        ),
        series_entry(
            "shock_active",
            "Engine!A3:C3",
            layout="series",
            direction="internal",
            dtype="int",
            header_row=1,
        ),
        series_entry(
            "path",
            "Engine!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
        schema_version="1.17.0",
    )


def _duplicate_header_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "duplicate_headers.xlsx",
        {
            "Inputs": {"A1": 2024, "C1": 2026},
            "Engine": {
                "A1": "=Inputs!A1",
                "B1": "=A1+1",
                "C1": "=Inputs!C1",
                "A2": "=A1",
                "B2": "=B1",
                "C2": "=C1",
            },
        },
    )


def _duplicate_header_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("first_year", "Inputs!A1", direction="input", dtype="int"),
        series_entry("third_year", "Inputs!C1", direction="input", dtype="int"),
        series_entry(
            "year_labels",
            "Engine!A1:C1",
            layout="series",
            direction="internal",
            dtype="int",
            header_row=1,
            axis_labels="TIME_PERIOD",
        ),
        series_entry(
            "path",
            "Engine!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
        schema_version="1.17.0",
    )


def _cycle_labeller_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry(
            "growth",
            "Inputs!A3:C3",
            layout="series",
            direction="input",
            header_row=2,
        ),
        series_entry(
            "year_labels",
            "Engine!A1:C1",
            layout="series",
            direction="internal",
            dtype="int",
            header_row=1,
            axis_labels="TIME_PERIOD",
        ),
        series_entry(
            "path",
            "Engine!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
        schema_version="1.17.0",
    )


def _cycle_labeller_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "cycle_labeller.xlsx",
        {
            "Inputs": {"A2": 2024, "B2": 2025, "C2": 2026, "A3": 1.0, "B3": 2.0, "C3": 3.0},
            "Engine": {
                "A1": "=Inputs!A3",
                "B1": "=Inputs!B3",
                "C1": "=Inputs!C3",
                "A2": "=A1",
                "B2": "=B1",
                "C2": "=C1",
            },
        },
    )


def _shift_tensor(tensor: Any, delta: int) -> Any:
    axis = tensor.domain.axes[0]
    domain = tensor.domain
    axis_cls = type(axis)
    domain_cls = type(domain)
    shifted = axis_cls(axis.name, tuple(key + delta for key in axis.keys), axis.key_type)
    domain = domain_cls.product(shifted)
    schema = getattr(tensor, "schema", None)
    if schema is None:
        return type(tensor).from_records(
            domain=domain,
            records=[((coord[0] + delta,), value) for coord, value in tensor.items()],
        )
    return type(tensor)(
        domain,
        tuple(value for _coord, value in tensor.items()),
        schema=schema,
        cells=getattr(tensor, "cells", None),
    )


def _snapshot_key_uses(source: str, keys: tuple[object, ...]) -> list[str]:
    """Return internals expressions that still mention a snapshot key of the axis."""
    tree = ast.parse(source)
    snapshot = set(keys)
    found: list[str] = []

    def add(node: ast.AST) -> None:
        found.append(ast.unparse(node))

    for node in ast.walk(tree):
        if isinstance(node, ast.Compare):
            for operand in [node.left, *node.comparators]:
                if isinstance(operand, ast.Constant) and operand.value in snapshot:
                    add(node)
        elif isinstance(node, ast.Subscript):
            slc = node.slice
            if isinstance(slc, ast.Constant) and slc.value in snapshot:
                add(node)
            if isinstance(slc, ast.Tuple):
                for elt in slc.elts:
                    if isinstance(elt, ast.Constant) and elt.value in snapshot:
                        add(node)
        elif isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
            if node.func.id not in {"span", "axis_step"}:
                continue
            for arg in node.args[1:]:
                if isinstance(arg, ast.Constant) and arg.value in snapshot:
                    add(node)
    return found


def test_static_packages_do_not_emit_runtime_axis_machinery(tmp_path: Path) -> None:
    modules = generate_inverted(_lag_workbook(tmp_path), _lag_bindings())
    joined = "\n".join(
        modules[name] for name in ("api.py", "model.py", "data.py", "internals.py", "validation.py")
    )
    assert "AxisTemplate" not in joined
    assert "SchemaTemplate" not in joined
    assert "LABELLED_AXES" not in joined
    assert "label_axis" not in modules["internals.py"]
    assert "def cells(self" not in modules["api.py"]
    assert "def cells(self" not in modules["model.py"]
    assert "if name in {" not in modules["api.py"]
    assert "if name in {" not in modules["model.py"]
    assert "Keys along" not in modules["api.py"]


def test_labeller_is_an_implicit_dependency(tmp_path: Path) -> None:
    workbook = _labelled_workbook(tmp_path)
    catalog, deps, _graph = inverted_graph_parts(workbook, _labelled_bindings())
    assert "year_labels" in deps["path"].param_ids
    assert "year_labels" in deps["shock_active"].param_ids
    pkg = load_package(
        generate_inverted(workbook, _labelled_bindings()), tmp_path, name="impl_edge"
    )
    assert "year_labels" in all_param_names(pkg.internals.path)
    assert "first_year" in input_field_names(pkg, pkg.compute_path)


def test_labeller_compiles_positionally_and_publishes_identity(tmp_path: Path) -> None:
    workbook = _labelled_workbook(tmp_path)
    modules = generate_inverted(workbook, _labelled_bindings())
    internals = modules["internals.py"]
    assert "data.YEAR_LABELS_POSITIONS" in internals
    assert "Series.from_labels(" in internals
    assert "label_axis('TIME_PERIOD'" in internals
    pkg = load_package(modules, tmp_path, name="labeller_body")
    catalog, _deps, graph = inverted_graph_parts(workbook, _labelled_bindings())
    kwargs = named_input_kwargs(pkg, catalog, graph)
    labels = pkg.internals.year_labels(first_year=kwargs["first_year"])
    assert tuple(labels.domain.axes[0].keys) == _YEARS
    assert list(labels.items()) == [((year,), year) for year in _YEARS]


def test_positional_emission_avoids_snapshot_key_literals(tmp_path: Path) -> None:
    modules = generate_inverted(_labelled_workbook(tmp_path), _labelled_bindings())
    internals = modules["internals.py"]
    assert "_p_time_period" in internals or "_p ==" in internals
    assert "axis_step(" in internals or "[_p - 1]" in internals or "[_p -1]" in internals
    uses = _snapshot_key_uses(internals, _YEARS)
    assert uses == [], uses


def test_shift_oracle_moves_keys_and_preserves_positional_values(tmp_path: Path) -> None:
    workbook = _labelled_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _labelled_bindings()), tmp_path, name="shift")
    catalog, _deps, graph = inverted_graph_parts(workbook, _labelled_bindings())
    kwargs = named_input_kwargs(pkg, catalog, graph)
    path_kwargs = {key: kwargs[key] for key in input_field_names(pkg, pkg.compute_path)}
    baseline = invoke_public_compute(pkg, pkg.compute_path, path_kwargs)
    assert tuple(baseline.domain.axes[0].keys) == _YEARS
    delta = 3
    shifted_growth = _shift_tensor(kwargs["growth"], delta)
    shifted = invoke_public_compute(
        pkg,
        pkg.compute_path,
        dict(
            first_year=kwargs["first_year"] + delta,
            growth=shifted_growth,
        ),
    )
    assert tuple(shifted.domain.axes[0].keys) == tuple(year + delta for year in _YEARS)
    for year in _YEARS:
        assert shifted[year + delta] == pytest.approx(baseline[year])
    snapshot_active = pkg.model.Model(**kwargs).shock_active
    moved_active = pkg.model.Model(
        first_year=kwargs["first_year"] + delta,
        shock_year=kwargs["shock_year"] + delta,
        growth=shifted_growth,
    ).shock_active
    assert snapshot_active[_YEARS[1]] == 1
    assert moved_active[_YEARS[1] + delta] == 1
    assert moved_active[_YEARS[0] + delta] == 0


def test_compute_result_is_a_series_and_rejects_snapshot_keys(tmp_path: Path) -> None:
    workbook = _labelled_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _labelled_bindings()), tmp_path, name="contract")
    catalog, _deps, graph = inverted_graph_parts(workbook, _labelled_bindings())
    kwargs = named_input_kwargs(pkg, catalog, graph)
    result = invoke_public_compute(
        pkg,
        pkg.compute_path,
        {key: kwargs[key] for key in input_field_names(pkg, pkg.compute_path)},
    )
    assert isinstance(result, pkg.Series)
    assert result.sel(TIME_PERIOD=2025) == pytest.approx(result[2025])
    records = pkg.as_records(pkg.compute_path, result)
    assert [row["TIME_PERIOD"] for row in records] == list(_YEARS)
    shifted = invoke_public_compute(
        pkg,
        pkg.compute_path,
        dict(
            first_year=2027,
            growth=_shift_tensor(kwargs["growth"], 3),
        ),
    )
    with pytest.raises(
        Exception, match="unknown labels.*2024.*accepted labels are \\(2027, 2028, 2029\\)"
    ):
        invoke_public_compute(
            pkg,
            pkg.compute_path,
            dict(
                first_year=2027,
                growth=kwargs["growth"],
            ),
        )
    assert tuple(shifted.domain.axes[0].keys) == (2027, 2028, 2029)


def test_duplicate_header_raises_axis_error(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_duplicate_header_workbook(tmp_path), _duplicate_header_bindings()),
        tmp_path,
        name="dup_headers",
    )
    with pytest.raises(Exception, match="TIME_PERIOD.*duplicate label 2024"):
        invoke_public_compute(pkg, pkg.compute_path, dict(first_year=2024, third_year=2024))


def test_package_matches_evaluator_at_snapshot_and_shifted_year(tmp_path: Path) -> None:
    document = _labelled_bindings()
    pkg = assert_package_matches_evaluator(
        _labelled_workbook(tmp_path), document, tmp_path, "labelled_parity"
    )
    catalog, _deps, graph = inverted_graph_parts(_labelled_workbook(tmp_path), document)
    kwargs = named_input_kwargs(pkg, catalog, graph)
    delta = 5
    shifted = _labelled_workbook(tmp_path, first=2024 + delta, shock_year=2025 + delta)
    shifted_catalog, _shifted_deps, shifted_graph = inverted_graph_parts(shifted, document)
    expected = FormulaEvaluator(shifted_graph).evaluate(list(shifted_catalog.get("path").cells))
    got = invoke_public_compute(
        pkg,
        pkg.compute_path,
        dict(
            first_year=kwargs["first_year"] + delta,
            growth=_shift_tensor(kwargs["growth"], delta),
        ),
    )
    for (runtime_coord, value), cell in zip(
        got.items(), shifted_catalog.get("path").cells, strict=True
    ):
        assert runtime_coord[0] == _YEARS[shifted_catalog.get("path").cells.index(cell)] + delta
        assert value == pytest.approx(expected[cell])


def test_internals_sel_uses_runtime_labels(tmp_path: Path) -> None:
    workbook = _labelled_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _labelled_bindings()), tmp_path, name="honest")
    catalog, _deps, graph = inverted_graph_parts(workbook, _labelled_bindings())
    kwargs = named_input_kwargs(pkg, catalog, graph)
    model = pkg.model.Model(
        first_year=kwargs["first_year"] + 2,
        shock_year=kwargs["shock_year"] + 2,
        growth=_shift_tensor(kwargs["growth"], 2),
    )
    year = 2026
    assert model.path.sel(TIME_PERIOD=year) == pytest.approx(model.path[year])
    cells = catalog.get("path").coordinate_cells
    snapshot_year = year - 2
    expected = FormulaEvaluator(graph).evaluate([cells[(snapshot_year,)]])[cells[(snapshot_year,)]]
    assert model.path.sel(TIME_PERIOD=year) == pytest.approx(expected)


def test_model_cells_bind_runtime_keys_and_key_note(tmp_path: Path) -> None:
    workbook = _labelled_workbook(tmp_path)
    modules = generate_inverted(workbook, _labelled_bindings())
    assert "Keys along `TIME_PERIOD` are determined by `first_year`" in modules["api.py"]
    pkg = load_package(modules, tmp_path, name="cells_note")
    catalog, _deps, graph = inverted_graph_parts(workbook, _labelled_bindings())
    kwargs = named_input_kwargs(pkg, catalog, graph)
    model = pkg.model.Model(
        first_year=2030,
        shock_year=2031,
        growth=_shift_tensor(kwargs["growth"], 6),
    )
    cells = model.cells("shock_active")
    assert dict(cells) == {
        (2030,): "Engine!A3",
        (2031,): "Engine!B3",
        (2032,): "Engine!C3",
    }


def test_labeller_input_on_labelled_axis_is_an_export_error(tmp_path: Path) -> None:
    with pytest.raises(InvertedTreeExportError, match="keyed on labelled axis"):
        generate_inverted(_cycle_labeller_workbook(tmp_path), _cycle_labeller_bindings())


def _horizon_sheets() -> dict[str, dict[str, object]]:
    """Five-year labelled horizon with a cumulative path and split history/projection."""
    return {
        "Inputs": {
            "B1": 2024,
            "A2": 2024,
            "B2": 2025,
            "C2": 2026,
            "D2": 2027,
            "E2": 2028,
            "A3": 1.0,
            "B3": 1.0,
            "C3": 1.0,
            "D3": 1.0,
            "E3": 1.0,
        },
        "Engine": {
            "A1": "=Inputs!B1",
            "B1": "=A1+1",
            "C1": "=B1+1",
            "D1": "=C1+1",
            "E1": "=D1+1",
            "A2": "=Inputs!A3",
            "B2": "=A2+1",
            "C2": "=B2+1",
            "D2": "=C2+1",
            "E2": "=D2+1",
            "A3": "=A2",
            "B3": "=A3+B2",
            "C3": "=B3+C2",
            "D3": "=C3+D2",
            "E3": "=D3+E2",
        },
    }


def _horizon_labeller_and_path() -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    return (
        series_entry("first_year", "Inputs!B1", direction="input", dtype="int"),
        series_entry(
            "year_labels",
            "Engine!A1:E1",
            layout="series",
            direction="internal",
            dtype="int",
            header_row=1,
            axis_labels="TIME_PERIOD",
        ),
        series_entry(
            "path",
            "Engine!A2:E2",
            layout="series",
            direction="internal",
            header_row=1,
        ),
    )


def _suffix_sheets() -> dict[str, dict[str, object]]:
    sheets = _horizon_sheets()
    sheets["Engine"]["C3"] = "=C2"
    sheets["Engine"]["D3"] = "=C3+D2"
    sheets["Engine"]["E3"] = "=D3+E2"
    del sheets["Engine"]["A3"]
    del sheets["Engine"]["B3"]
    return sheets


def _suffix_bindings() -> dict[str, Any]:
    first_year, year_labels, path = _horizon_labeller_and_path()
    return bindings_document(
        first_year,
        series_entry(
            "growth",
            "Inputs!A3:E3",
            layout="series",
            direction="input",
            header_row=2,
        ),
        year_labels,
        path,
        series_entry(
            "projection",
            "Engine!C3:E3",
            layout="series",
            direction="output",
            dtype="float",
            header_row=1,
        ),
        schema_version="1.17.0",
    )


def _prefix_bindings() -> dict[str, Any]:
    first_year, year_labels, path = _horizon_labeller_and_path()
    return bindings_document(
        first_year,
        series_entry(
            "growth",
            "Inputs!A3:E3",
            layout="series",
            direction="input",
            header_row=2,
        ),
        year_labels,
        path,
        series_entry(
            "history",
            "Engine!A3:C3",
            layout="series",
            direction="output",
            dtype="float",
            header_row=1,
        ),
        schema_version="1.17.0",
    )


def test_suffix_series_family_and_lag(tmp_path: Path) -> None:
    """A projection suffix seeds at labeller position 2, not series-local 0."""
    workbook = write_workbook(tmp_path / "suffix.xlsx", _suffix_sheets())
    document = _suffix_bindings()
    modules = generate_inverted(workbook, document)
    internals = modules["internals.py"]
    assert re.search(r"def projection\b[\s\S]*?_p_time_period == 2", internals)
    pkg = load_package(modules, tmp_path, name="suffix_axis")
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    kwargs = named_input_kwargs(pkg, catalog, graph)
    result = invoke_public_compute(
        pkg,
        pkg.compute_projection,
        {key: kwargs[key] for key in input_field_names(pkg, pkg.compute_projection)},
    )
    assert tuple(result.domain.axes[0].keys) == (2026, 2027, 2028)
    assert [value for _coord, value in result.items()] == pytest.approx([3.0, 7.0, 12.0])
    shifted_growth = _shift_tensor(kwargs["growth"], 10)
    shifted = invoke_public_compute(
        pkg,
        pkg.compute_projection,
        dict(first_year=kwargs["first_year"] + 10, growth=shifted_growth),
    )
    assert tuple(shifted.domain.axes[0].keys) == (2036, 2037, 2038)
    assert [value for _coord, value in shifted.items()] == pytest.approx([3.0, 7.0, 12.0])


def test_prefix_series_family_and_lag(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "prefix.xlsx", _horizon_sheets())
    document = _prefix_bindings()
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="prefix_axis")
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    kwargs = named_input_kwargs(pkg, catalog, graph)
    result = invoke_public_compute(
        pkg,
        pkg.compute_history,
        {key: kwargs[key] for key in input_field_names(pkg, pkg.compute_history)},
    )
    assert tuple(result.domain.axes[0].keys) == (2024, 2025, 2026)
    assert [value for _coord, value in result.items()] == pytest.approx([1.0, 3.0, 6.0])


def test_subset_input_default_import(tmp_path: Path) -> None:
    """A runtime-axis input covering only the projection suffix still imports."""
    workbook = write_workbook(
        tmp_path / "subset_input.xlsx",
        {
            "Inputs": {
                "B1": 2024,
                "C2": 2026,
                "D2": 2027,
                "E2": 2028,
                "C3": 1.0,
                "D3": 2.0,
                "E3": 3.0,
            },
            "Engine": {
                "A1": "=Inputs!B1",
                "B1": "=A1+1",
                "C1": "=B1+1",
                "D1": "=C1+1",
                "E1": "=D1+1",
                "C2": "=Inputs!C3",
                "D2": "=Inputs!D3",
                "E2": "=Inputs!E3",
            },
        },
    )
    document = bindings_document(
        series_entry("first_year", "Inputs!B1", direction="input", dtype="int"),
        series_entry(
            "growth",
            "Inputs!C3:E3",
            layout="series",
            direction="input",
            header_row=2,
        ),
        series_entry(
            "year_labels",
            "Engine!A1:E1",
            layout="series",
            direction="internal",
            dtype="int",
            header_row=1,
            axis_labels="TIME_PERIOD",
        ),
        series_entry(
            "path",
            "Engine!C2:E2",
            layout="series",
            direction="output",
            header_row=1,
        ),
        schema_version="1.17.0",
    )
    modules = generate_inverted(workbook, document)
    assert "source=(2024, 2025, 2026, 2027, 2028)" in modules["data.py"]
    pkg = load_package(modules, tmp_path, name="subset_input")
    assert tuple(pkg.data.GROWTH.domain.axes[0].keys) == (2026, 2027, 2028)
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    kwargs = named_input_kwargs(pkg, catalog, graph)
    result = invoke_public_compute(
        pkg,
        pkg.compute_path,
        {key: kwargs[key] for key in input_field_names(pkg, pkg.compute_path)},
    )
    assert tuple(result.domain.axes[0].keys) == (2026, 2027, 2028)
    assert [value for _coord, value in result.items()] == pytest.approx([1.0, 2.0, 3.0])


def test_labelled_offset_steps_on_the_labeller_axis(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "offset.xlsx",
        {
            "Inputs": {"A1": 2024},
            "Engine": {
                "A1": "=Inputs!A1",
                "B1": "=A1+1",
                "C1": "=B1+1",
                "A2": 1.0,
                "B2": "=OFFSET(B2,0,-1)+1",
                "C2": "=OFFSET(C2,0,-1)+1",
            },
        },
    )
    document = bindings_document(
        series_entry("first_year", "Inputs!A1", direction="input", dtype="int"),
        series_entry(
            "year_labels",
            "Engine!A1:C1",
            layout="series",
            direction="internal",
            dtype="int",
            header_row=1,
            axis_labels="TIME_PERIOD",
        ),
        series_entry(
            "path",
            "Engine!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
        schema_version="1.17.0",
    )
    modules = generate_inverted(workbook, document)
    assert "axis_step(" in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="offset_axis")
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    kwargs = named_input_kwargs(pkg, catalog, graph)
    result = invoke_public_compute(
        pkg,
        pkg.compute_path,
        {key: kwargs[key] for key in input_field_names(pkg, pkg.compute_path)},
    )
    assert tuple(result.domain.axes[0].keys) == (2024, 2025, 2026)
    assert [value for _coord, value in result.items()] == pytest.approx([1.0, 2.0, 3.0])


def test_labelled_suffix_offset_walks_onto_history(tmp_path: Path) -> None:
    """OFFSET from a projection suffix must land on the history series."""
    sheets = _horizon_sheets()
    sheets["Engine"]["C3"] = "=OFFSET(C3,0,-1)+C2"
    sheets["Engine"]["D3"] = "=C3+D2"
    sheets["Engine"]["E3"] = "=D3+E2"
    workbook = write_workbook(tmp_path / "suffix_offset.xlsx", sheets)
    first_year, year_labels, path = _horizon_labeller_and_path()
    document = bindings_document(
        first_year,
        series_entry(
            "growth",
            "Inputs!A3:E3",
            layout="series",
            direction="input",
            header_row=2,
        ),
        year_labels,
        path,
        series_entry(
            "history",
            "Engine!A3:B3",
            layout="series",
            direction="internal",
            dtype="float",
            header_row=1,
        ),
        series_entry(
            "projection",
            "Engine!C3:E3",
            layout="series",
            direction="output",
            dtype="float",
            header_row=1,
        ),
        schema_version="1.17.0",
    )
    modules = generate_inverted(workbook, document)
    internals = modules["internals.py"]
    assert re.search(
        r"def projection\b[\s\S]*?"
        r"history\[axis_step\(_ax_time_period, time_period, xl_neg\(1\)\)\]",
        internals,
    )
    pkg = load_package(modules, tmp_path, name="suffix_offset_axis")
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    assert "history" in deps["projection"].param_ids
    kwargs = named_input_kwargs(pkg, catalog, graph)
    result = invoke_public_compute(
        pkg,
        pkg.compute_projection,
        {key: kwargs[key] for key in input_field_names(pkg, pkg.compute_projection)},
    )
    assert tuple(result.domain.axes[0].keys) == (2026, 2027, 2028)
    assert [value for _coord, value in result.items()] == pytest.approx([6.0, 10.0, 15.0])


def test_labelled_tensor_constant_overrides_without_identity_bind(tmp_path: Path) -> None:
    """data.overrides must not bind AxisTemplate/SeriesSpec as a labeller tensor."""
    workbook = write_workbook(
        tmp_path / "const_override.xlsx",
        {
            "Inputs": {"A1": 2024},
            "Engine": {
                "A1": "=Inputs!A1",
                "B1": "=A1+1",
                "C1": "=B1+1",
                "A2": 1.0,
                "B2": 2.0,
                "C2": 3.0,
                "D2": "=SUM(A2:C2)+A1*0",
            },
        },
    )
    document = bindings_document(
        series_entry("first_year", "Inputs!A1", direction="input", dtype="int"),
        series_entry(
            "year_labels",
            "Engine!A1:C1",
            layout="series",
            direction="internal",
            dtype="int",
            header_row=1,
            axis_labels="TIME_PERIOD",
        ),
        series_entry(
            "rates",
            "Engine!A2:C2",
            layout="series",
            direction="constant",
            header_row=1,
        ),
        series_entry("total", "Engine!D2", direction="output", dtype="float"),
        schema_version="1.17.0",
    )
    modules = generate_inverted(workbook, document)
    data = modules["data.py"]
    assert "getattr(schema, 'bind'" not in data
    assert "namespace[labeller.upper()]" not in data
    pkg = load_package(modules, tmp_path, name="const_override")
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    kwargs = named_input_kwargs(pkg, catalog, graph)
    compute_kwargs = {key: kwargs[key] for key in input_field_names(pkg, pkg.compute_total)}
    assert invoke_public_compute(pkg, pkg.compute_total, compute_kwargs) == pytest.approx(6.0)
    original = pkg.data.RATES
    zeros = original.with_values((0.0, 0.0, 0.0))
    with pkg.data.overrides(RATES=zeros):
        assert invoke_public_compute(pkg, pkg.compute_total, compute_kwargs) == pytest.approx(0.0)
    assert pkg.data.RATES is original
    assert invoke_public_compute(pkg, pkg.compute_total, compute_kwargs) == pytest.approx(6.0)


def test_ragged_runtime_range_does_not_invent_keys(tmp_path: Path) -> None:
    """A non-contiguous labelled range must not span the interned hole."""
    workbook = write_workbook(
        tmp_path / "ragged.xlsx",
        {
            "Inputs": {"A1": 2024},
            "Engine": {
                "A1": "=Inputs!A1",
                "B1": "=A1+1",
                "C1": "=B1+1",
                "A2": 1.0,
                "C2": 3.0,
                "D2": "=SUM(A2:C2)+A1*0",
            },
        },
    )
    odd = series_entry(
        "odd_years",
        "Engine!A2:C2",
        layout="series",
        direction="input",
        header_row=1,
    )
    odd["exclude_columns"] = ["B"]
    document = bindings_document(
        series_entry("first_year", "Inputs!A1", direction="input", dtype="int"),
        series_entry(
            "year_labels",
            "Engine!A1:C1",
            layout="series",
            direction="internal",
            dtype="int",
            header_row=1,
            axis_labels="TIME_PERIOD",
        ),
        odd,
        series_entry(
            "total",
            "Engine!D2",
            direction="output",
            dtype="float",
        ),
        schema_version="1.17.0",
    )
    modules = generate_inverted(workbook, document, blank_ranges=["Engine!B2"])
    internals = modules["internals.py"]
    assert "span(" not in internals
    assert "_ax_time_period.keys[0]" in internals
    assert "_ax_time_period.keys[2]" in internals
    pkg = load_package(modules, tmp_path, name="ragged_axis")
    catalog, _deps, graph = inverted_graph_parts(workbook, document, blank_ranges=["Engine!B2"])
    kwargs = named_input_kwargs(pkg, catalog, graph)
    result = invoke_public_compute(
        pkg,
        pkg.compute_total,
        {key: kwargs[key] for key in input_field_names(pkg, pkg.compute_total)},
    )
    assert result == pytest.approx(4.0)


def test_runtime_string_remap_is_an_export_error(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "string_remap.xlsx",
        {
            "Data": {
                "A1": "name",
                "B1": "ifs",
                "C1": "code",
                "A2": "Alpha",
                "B2": 10.0,
                "C2": "=D2",
                "D2": "A",
                "E2": "=B2",
                "A3": "Beta",
                "B3": 20.0,
                "C3": "=D3",
                "D3": "B",
                "E3": "=B3",
            }
        },
    )
    document = bindings_document(
        series_entry("code_a", "Data!D2", direction="input", dtype="string", key_read="string"),
        series_entry("code_b", "Data!D3", direction="input", dtype="string", key_read="string"),
        series_entry(
            "codes",
            "Data!C2:C3",
            layout="series",
            direction="internal",
            dtype="string",
            label_column="C",
            key_concept="REF_AREA",
            key_read="string",
            axis_labels="REF_AREA",
        ),
        series_entry(
            "catalog_ifs",
            "Data!B2:B3",
            layout="series",
            direction="input",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
        series_entry(
            "trigger_ifs",
            "Data!E2:E3",
            layout="series",
            direction="output",
            label_column="C",
            key_concept="REF_AREA",
            key_read="string",
        ),
        schema_version="1.17.0",
    )
    document["concept_scheme"]["concepts"].append({"id": "REF_AREA", "dtype": "string"})
    with pytest.raises(InvertedTreeExportError, match="remap"):
        generate_inverted(workbook, document)


def test_two_runtime_labellers_for_one_axis_name_are_an_export_error(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "two_labellers.xlsx",
        {
            "Engine": {
                "A1": "=2023+1",
                "B1": "=A1+1",
                "C1": "=2025+1",
                "D1": "=C1+1",
                "A2": "=A1",
                "B2": "=B1",
                "C2": "=C1",
                "D2": "=D1",
            }
        },
    )
    document = bindings_document(
        series_entry(
            "history_labels",
            "Engine!A1:B1",
            layout="series",
            direction="internal",
            dtype="int",
            header_row=1,
            axis_labels="TIME_PERIOD",
        ),
        series_entry(
            "projection_labels",
            "Engine!C1:D1",
            layout="series",
            direction="internal",
            dtype="int",
            header_row=1,
            axis_labels="TIME_PERIOD",
        ),
        series_entry(
            "path",
            "Engine!A2:D2",
            layout="series",
            direction="output",
            header_row=1,
        ),
        schema_version="1.17.0",
    )
    with pytest.raises(InvertedTreeExportError, match="multiple runtime labellers"):
        generate_inverted(workbook, document)


def test_evaluator_override_on_same_graph_matches_shifted_package(tmp_path: Path) -> None:
    document = _labelled_bindings()
    workbook = _labelled_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="p5_eval")
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    kwargs = named_input_kwargs(pkg, catalog, graph)
    delta = 5
    graph.set_node_value("Inputs!A1", kwargs["first_year"] + delta)
    expected = FormulaEvaluator(graph).evaluate(list(catalog.get("path").cells))
    got = invoke_public_compute(
        pkg,
        pkg.compute_path,
        dict(
            first_year=kwargs["first_year"] + delta,
            growth=_shift_tensor(kwargs["growth"], delta),
        ),
    )
    for (_coord, value), cell in zip(got.items(), catalog.get("path").cells, strict=True):
        assert value == pytest.approx(expected[cell])
