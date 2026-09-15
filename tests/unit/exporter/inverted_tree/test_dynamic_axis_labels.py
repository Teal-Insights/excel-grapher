"""Runtime axis labels keyed by the labeller all the way down (#841)."""

from __future__ import annotations

import ast
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
    inverted_graph_parts,
    load_package,
    named_input_kwargs,
    required_param_names,
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
        modules[name] for name in ("api.py", "data.py", "internals.py", "validation.py")
    )
    assert "AxisTemplate" not in joined
    assert "SchemaTemplate" not in joined
    assert "LABELLED_AXES" not in joined
    assert "label_axis" not in modules["internals.py"]
    assert "def cells(self" not in modules["api.py"]
    assert "if name in {" not in modules["api.py"]
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
    assert "first_year" in required_param_names(pkg.compute_path)


def test_labeller_compiles_positionally_and_publishes_identity(tmp_path: Path) -> None:
    workbook = _labelled_workbook(tmp_path)
    modules = generate_inverted(workbook, _labelled_bindings())
    internals = modules["internals.py"]
    assert "data.YEAR_LABELS_POSITIONS" in internals
    assert "from_labels(" in internals
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
    path_kwargs = {key: kwargs[key] for key in required_param_names(pkg.compute_path)}
    baseline = pkg.compute_path(**path_kwargs)
    assert tuple(baseline.domain.axes[0].keys) == _YEARS
    delta = 3
    shifted_growth = _shift_tensor(kwargs["growth"], delta)
    shifted = pkg.compute_path(
        first_year=kwargs["first_year"] + delta,
        growth=shifted_growth,
    )
    assert tuple(shifted.domain.axes[0].keys) == tuple(year + delta for year in _YEARS)
    for year in _YEARS:
        assert shifted[year + delta] == pytest.approx(baseline[year])
    snapshot_active = pkg.api.Model(**kwargs).shock_active
    moved_active = pkg.api.Model(
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
    result = pkg.compute_path(
        **{key: kwargs[key] for key in required_param_names(pkg.compute_path)}
    )
    assert isinstance(result, pkg.Series)
    assert result.sel(TIME_PERIOD=2025) == pytest.approx(result[2025])
    records = pkg.as_records(pkg.compute_path, result)
    assert [row["TIME_PERIOD"] for row in records] == list(_YEARS)
    shifted = pkg.compute_path(
        first_year=2027,
        growth=_shift_tensor(kwargs["growth"], 3),
    )
    with pytest.raises(
        Exception, match="unknown labels.*2024.*accepted labels are \\(2027, 2028, 2029\\)"
    ):
        pkg.compute_path(
            first_year=2027,
            growth=kwargs["growth"],
        )
    assert tuple(shifted.domain.axes[0].keys) == (2027, 2028, 2029)


def test_duplicate_header_raises_axis_error(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_duplicate_header_workbook(tmp_path), _duplicate_header_bindings()),
        tmp_path,
        name="dup_headers",
    )
    with pytest.raises(Exception, match="TIME_PERIOD.*duplicate label 2024"):
        pkg.compute_path(first_year=2024, third_year=2024)


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
    got = pkg.compute_path(
        first_year=kwargs["first_year"] + delta,
        growth=_shift_tensor(kwargs["growth"], delta),
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
    model = pkg.api.Model(
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
    model = pkg.api.Model(
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
