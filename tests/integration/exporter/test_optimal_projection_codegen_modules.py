"""Integration: optimal-compression projected packages import and run."""

from __future__ import annotations

import importlib
import sys
from collections.abc import Callable
from pathlib import Path
from typing import cast

import xlsxwriter

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter import CodeGenerator, OptimalCompression
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import Records, WorkbookSeriesBindings

_PACKAGE = "optimal_projected_export_pkg"


def _write_inline_workbook(workbook_path: Path) -> None:
    wb = xlsxwriter.Workbook(workbook_path)
    engine = wb.add_worksheet("Engine")
    engine.write_number("C6", 10)
    out = wb.add_worksheet("Outputs")
    out.write_formula("B12", "=Engine!C6")
    out.write_formula("B14", "=Outputs!B12+1")
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


def _clear_package_modules() -> None:
    for name in list(sys.modules):
        if name == _PACKAGE or name.startswith(f"{_PACKAGE}."):
            sys.modules.pop(name, None)


def test_optimal_projected_generate_modules_package_runs_and_matches_evaluator(
    tmp_path: Path,
) -> None:
    workbook_path = tmp_path / "optimal_target.xlsx"
    _write_inline_workbook(workbook_path)

    targets = ["Outputs!B12", "Outputs!B14"]
    graph = create_dependency_graph(
        workbook_path,
        targets,
        load_values=True,
        capture_dependency_provenance=True,
    )
    bindings = _baseline_bindings(workbook_path)

    projection = OptimalCompression().project(graph)
    # Target / series-bound public addresses must remain in the projection so
    # series helpers see the full published leaf set (no alias patch-up).
    assert "Outputs!B12" in projection

    files = CodeGenerator(projection).generate_modules(
        targets,
        series_bindings=bindings,
        bindings_workbook=workbook_path,
    )

    assert "xl_cell" in files["internals.py"]
    assert "def compute_baseline(" in files["api.py"]
    assert "# --- Projection public address aliases ---" not in files["internals.py"]

    pkg_dir = tmp_path / _PACKAGE
    for filename, content in files.items():
        pkg_dir.mkdir(parents=True, exist_ok=True)
        (pkg_dir / filename).write_text(content, encoding="utf-8")

    sys.path.insert(0, str(tmp_path))
    try:
        pkg = importlib.import_module(_PACKAGE)

        compute_baseline = cast(Callable[..., Records], pkg.compute_baseline)
        records = compute_baseline(ctx=pkg.make_context())
        assert len(records) == 1
        assert records[0]["OBS_VALUE"] == 10

        generated_targets = pkg.compute_all()
        with FormulaEvaluator(graph) as ev:
            evaluator_results = ev.evaluate(targets)
        assert generated_targets == evaluator_results
    finally:
        sys.path.remove(str(tmp_path))
        _clear_package_modules()
