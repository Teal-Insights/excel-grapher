"""Codegen tests for projected export with public address aliases."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import cast

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter import CodeGenerator, IdentityTransitCompression
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import Records, WorkbookSeriesBindings


def _exec_generated(code: str) -> dict[str, object]:
    namespace: dict[str, object] = {}
    exec(code, namespace)
    return namespace


def _write_identity_workbook(workbook_path: Path) -> None:
    import xlsxwriter

    wb = xlsxwriter.Workbook(workbook_path)
    ws = wb.add_worksheet("Engine")
    ws.write_number("C6", 10)
    out = wb.add_worksheet("Outputs")
    out.write_formula("B12", "=Engine!C6")
    out.write_formula("B14", "=Outputs!B12+1")
    wb.close()


def _write_identity_workbook_with_unrelated_component(workbook_path: Path) -> None:
    import xlsxwriter

    wb = xlsxwriter.Workbook(workbook_path)
    engine = wb.add_worksheet("Engine")
    engine.write_number("C6", 10)
    engine.write_number("C7", 20)
    outputs = wb.add_worksheet("Outputs")
    outputs.write_formula("B12", "=Engine!C6")
    outputs.write_formula("B14", "=Outputs!B12+1")
    other = wb.add_worksheet("Other")
    other.write_formula("B1", "=Engine!C7")
    other.write_formula("B2", "=Other!B1+1")
    wb.close()


def _baseline_bindings(workbook_path: Path) -> WorkbookSeriesBindings:
    return cast(
        WorkbookSeriesBindings,
        {
            "schema_version": "1.2.0",
            "workbook": str(workbook_path),
            "series": [
                {
                    "id": "baseline",
                    "data_range": "Outputs!B12",
                    "layout": "scalar",
                    "output": {"compute": {"name": "compute_baseline"}},
                    "structure": {
                        "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
                        "dimensions": [
                            {
                                "concept": "LABEL",
                                "role": "key",
                                "scope": "series",
                                "bind": {"kind": "constant", "value": "baseline"},
                            }
                        ],
                    },
                    "key": ["LABEL"],
                }
            ],
        },
    )


def test_projected_codegen_emits_compute_for_removed_public_mirror(tmp_path: Path) -> None:
    workbook_path = tmp_path / "identity_target.xlsx"
    _write_identity_workbook(workbook_path)

    graph = create_dependency_graph(
        workbook_path,
        ["Outputs!B12", "Outputs!B14"],
        load_values=True,
        capture_dependency_provenance=True,
    )
    bindings = _baseline_bindings(workbook_path)

    projection = IdentityTransitCompression().project(graph)
    with CodeGenerator(projection) as gen:
        code = gen.generate(
            ["Outputs!B12", "Outputs!B14"],
            series_bindings=bindings,
            bindings_workbook=workbook_path,
        )

    assert "def compute_baseline(" in code
    assert "# --- Projection public address aliases ---" in code
    assert "Outputs!B12" in code

    ns = _exec_generated(code)
    compute = cast(Callable[..., Records], ns["compute_baseline"])
    records = compute()
    assert len(records) == 1
    assert records[0]["OBS_VALUE"] == 10


def test_projected_generate_modules_emits_alias_resolver(tmp_path: Path) -> None:
    workbook_path = tmp_path / "identity_target.xlsx"
    _write_identity_workbook(workbook_path)

    graph = create_dependency_graph(
        workbook_path,
        ["Outputs!B12", "Outputs!B14"],
        load_values=True,
        capture_dependency_provenance=True,
    )
    bindings = _baseline_bindings(workbook_path)

    projection = IdentityTransitCompression().project(graph)
    modules = CodeGenerator(projection).generate_modules(
        ["Outputs!B12", "Outputs!B14"],
        series_bindings=bindings,
        bindings_workbook=workbook_path,
    )

    internals = modules["internals.py"]
    api = modules["api.py"]
    assert "def compute_baseline(" in api
    # Projected leaf aliases resolve through `xl_cell`; formula→formula `xl_eval`
    # is only imported when a body actually references another formula cell.
    assert "xl_cell" in internals
    assert "Outputs!B12" in internals or "outputs_b12" in internals.lower()


def test_projected_codegen_omits_unrelated_public_aliases_outside_export_closure(
    tmp_path: Path,
) -> None:
    workbook_path = tmp_path / "identity_target.xlsx"
    _write_identity_workbook_with_unrelated_component(workbook_path)

    graph = create_dependency_graph(
        workbook_path,
        ["Outputs!B12", "Outputs!B14", "Other!B2"],
        load_values=True,
        capture_dependency_provenance=True,
    )
    projection = IdentityTransitCompression().project(graph)
    assert projection.manifest.map_to_projected("Other!B1") == "Engine!C7"

    code = CodeGenerator(projection).generate(["Outputs!B12"])

    assert "cell_outputs_b12" in code
    assert "cell_other_b1" not in code
    assert "Engine!C7" not in code


def test_projected_codegen_omits_internal_aliases_inside_export_closure(tmp_path: Path) -> None:
    workbook_path = tmp_path / "identity_target.xlsx"
    _write_identity_workbook(workbook_path)

    graph = create_dependency_graph(
        workbook_path,
        ["Outputs!B14"],
        load_values=True,
        capture_dependency_provenance=True,
    )
    projection = IdentityTransitCompression().project(graph)
    assert projection.manifest.map_to_projected("Outputs!B12") == "Engine!C6"

    code = CodeGenerator(projection).generate(["Outputs!B14"])

    assert "cell_outputs_b14" in code
    assert "cell_outputs_b12" not in code


def test_projected_codegen_preserves_public_targets_for_removed_mirror(tmp_path: Path) -> None:
    workbook_path = tmp_path / "identity_target.xlsx"
    _write_identity_workbook(workbook_path)

    graph = create_dependency_graph(
        workbook_path,
        ["Outputs!B12", "Outputs!B14"],
        load_values=True,
        capture_dependency_provenance=True,
    )
    mirror = graph.get_node("Outputs!B12")
    assert mirror is not None
    graph._nodes["Outputs!B12"].is_target = True

    projection = IdentityTransitCompression().project(graph)
    with CodeGenerator(projection) as gen:
        code = gen.generate()

    assert "'Outputs!B12'" in code or '"Outputs!B12"' in code
    ns = _exec_generated(code)
    targets = cast(dict[str, object], ns["TARGETS"])
    assert "Outputs!B12" in targets


def test_projected_codegen_matches_evaluator_on_public_targets(tmp_path: Path) -> None:
    workbook_path = tmp_path / "identity_target.xlsx"
    _write_identity_workbook(workbook_path)

    graph = create_dependency_graph(
        workbook_path,
        ["Outputs!B12", "Outputs!B14"],
        load_values=True,
        capture_dependency_provenance=True,
    )
    targets = ["Outputs!B12", "Outputs!B14"]
    projection = IdentityTransitCompression().project(graph)

    with FormulaEvaluator(graph) as ev:
        evaluator_results = ev.evaluate(targets)

    code = CodeGenerator(projection).generate(targets)
    ns = _exec_generated(code)
    ctx_factory = cast(Callable[..., object], ns["make_context"])
    xl_cell = cast(Callable[..., object], ns["xl_cell"])
    ctx = ctx_factory()
    generated_results = {target: xl_cell(ctx, target) for target in targets}

    assert generated_results["Outputs!B12"] == evaluator_results["Outputs!B12"]
    assert generated_results["Outputs!B14"] == evaluator_results["Outputs!B14"]


def test_projected_codegen_matches_evaluator_with_unpack_return(tmp_path: Path) -> None:
    workbook_path = tmp_path / "identity_target.xlsx"
    _write_identity_workbook(workbook_path)

    graph = create_dependency_graph(
        workbook_path,
        ["Outputs!B12", "Outputs!B14"],
        load_values=True,
        capture_dependency_provenance=True,
    )
    targets = ["Outputs!B12", "Outputs!B14"]
    projection = IdentityTransitCompression().project(graph)

    with FormulaEvaluator(graph) as ev:
        evaluator_results = ev.evaluate(targets)

    code = CodeGenerator(projection, unpack_return=True).generate(targets)
    assert "# --- Projection public address aliases ---" in code
    ns = _exec_generated(code)
    ctx_factory = cast(Callable[..., object], ns["make_context"])
    xl_cell = cast(Callable[..., object], ns["xl_cell"])
    ctx = ctx_factory()
    generated_results = {target: xl_cell(ctx, target) for target in targets}

    assert generated_results == evaluator_results


def _write_identity_workbook_formula_target(workbook_path: Path) -> None:
    import xlsxwriter

    wb = xlsxwriter.Workbook(workbook_path)
    engine = wb.add_worksheet("Engine")
    engine.write_number("C5", 5)
    engine.write_formula("C6", "=Engine!C5*2", None, 10)
    out = wb.add_worksheet("Outputs")
    out.write_formula("B12", "=Engine!C6", None, 10)
    out.write_formula("B14", "=Outputs!B12+1", None, 11)
    wb.close()


def test_projected_codegen_alias_delegates_to_retained_formula(tmp_path: Path) -> None:
    workbook_path = tmp_path / "identity_formula_target.xlsx"
    _write_identity_workbook_formula_target(workbook_path)

    graph = create_dependency_graph(
        workbook_path,
        ["Outputs!B12", "Outputs!B14"],
        load_values=True,
        capture_dependency_provenance=True,
    )
    projection = IdentityTransitCompression().project(graph)
    assert projection.manifest.map_to_projected("Outputs!B12") == "Engine!C6"

    code = CodeGenerator(projection).generate(["Outputs!B12", "Outputs!B14"])
    assert "# --- Projection public address aliases ---" in code
    alias_start = code.index("# --- Projection public address aliases ---")
    alias_section = code[alias_start:]
    assert "def cell_outputs_b12(ctx):" in alias_section
    assert "xl_eval(ctx, 'Engine!C6', cell_engine_c6)" in alias_section
    assert "xl_number(xl_cell(ctx, 'Engine!C5'))" not in alias_section

    ns = _exec_generated(code)
    ctx_factory = cast(Callable[..., object], ns["make_context"])
    xl_cell = cast(Callable[..., object], ns["xl_cell"])
    ctx = ctx_factory()
    assert xl_cell(ctx, "Outputs!B12") == 10
