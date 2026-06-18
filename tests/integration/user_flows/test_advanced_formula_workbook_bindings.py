"""Bindings validation and ``_xludf`` regression for ``advanced_formula_workbook``."""

from __future__ import annotations

import re
import tempfile
from pathlib import Path

import pytest
from fastpyxl import load_workbook

from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.series_bindings.workflow import run_binding_checks, validate_bindings_workbook
from tests.integration.user_flows.bindings_accuracy import (
    BindingsAccuracyCase,
    assert_bindings_validate,
)
from tests.integration.utils.rewrite_xludf_workbook import rewrite_formula_to_xludf

EXAMPLES = Path(__file__).resolve().parents[3] / "examples" / "micro_workbooks"
SANDBOX = Path(__file__).resolve().parents[3] / "sandbox" / "model"

WORKBOOK = EXAMPLES / "advanced_formula_workbook_xludf.xlsx"
BINDINGS_DIR = EXAMPLES / "advanced_formula_workbook_xludf.bindings"
SANDBOX_WORKBOOK = SANDBOX / "advanced_formula_workbook.xlsx"
SANDBOX_BINDINGS = SANDBOX / "advanced_formula_workbook.bindings"

_XLUDF_RUNTIME_PATTERN = re.compile(r"xl__xludf_[a-z0-9_]+")


def _skip_if_fixtures_missing() -> None:
    if not WORKBOOK.is_file():
        pytest.skip(f"Workbook fixture missing: {WORKBOOK}")
    if not BINDINGS_DIR.is_dir():
        pytest.skip(f"Bindings directory missing: {BINDINGS_DIR}")


def test_sandbox_bindings_validate() -> None:
    """Sandbox shard bindings validate against the source workbook."""
    if not SANDBOX_WORKBOOK.is_file() or not SANDBOX_BINDINGS.is_dir():
        pytest.skip("sandbox advanced_formula_workbook fixtures missing")
    result = validate_bindings_workbook(SANDBOX_WORKBOOK, SANDBOX_BINDINGS)
    assert result["report"]["ok"], result["report"]["issues"]


def test_xludf_fixture_bindings_validate() -> None:
    """Committed ``_xludf`` workbook fixture validates with its binding shards."""
    _skip_if_fixtures_missing()
    assert_bindings_validate(
        BindingsAccuracyCase(
            name="advanced_formula_workbook_xludf",
            workbook=WORKBOOK,
            bindings_path=BINDINGS_DIR,
        )
    )


def test_xludf_fixture_workbook_contains_xludf_formulas() -> None:
    """Fixture on disk uses ``_xludf.`` spelling for allowlisted built-ins."""
    _skip_if_fixtures_missing()
    wb = load_workbook(WORKBOOK, data_only=False)
    found = 0
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and "_xludf." in cell.value.lower():
                        found += 1
    finally:
        wb.close()
    assert found >= 4, "expected multiple _xludf formulas in xludf fixture"


def test_xludf_formula_rewrite_idempotent_on_sandbox_cells() -> None:
    """Rewrite helper produces ``_xludf`` spellings for known workbook formulas."""
    if not SANDBOX_WORKBOOK.is_file():
        pytest.skip("sandbox workbook missing")
    wb = load_workbook(SANDBOX_WORKBOOK, data_only=False)
    try:
        samples = [
            wb["Product Lookup"]["K12"].value,
            wb["Formula Toolkit"]["D12"].value,
        ]
    finally:
        wb.close()
    for formula in samples:
        assert isinstance(formula, str)
        rewritten = rewrite_formula_to_xludf(formula)
        assert "_xludf." in rewritten.lower()


def test_codegen_modules_never_emit_xludf_runtime_symbols(tmp_path: Path) -> None:
    """Modular export must not reference ``xl__xludf_*`` runtime names."""
    from excel_grapher import DependencyGraph, Node

    graph = DependencyGraph()
    graph.add_node(
        Node(
            sheet="S",
            column="A",
            row=1,
            formula=None,
            normalized_formula=None,
            value=1,
            is_leaf=True,
        )
    )
    graph.add_node(
        Node(
            sheet="S",
            column="B",
            row=1,
            formula='=_xludf.IFNA(_xludf.XLOOKUP(1,S!A1:S!A1,S!A1:S!A1),"x")',
            normalized_formula='=_xludf.IFNA(_xludf.XLOOKUP(1,S!A1:S!A1,S!A1:S!A1),"x")',
            value=None,
            is_leaf=False,
        )
    )
    files = CodeGenerator(graph).generate_modules(["S!B1"])
    combined = "\n".join(files.values())
    assert _XLUDF_RUNTIME_PATTERN.search(combined) is None
    assert "xl__xlfn_xlookup" in combined or "xl_ifna" in combined


def test_sandbox_advanced_formula_workbook_smoke_test() -> None:
    """Full binding smoke test for the sandbox advanced formula workbook."""
    if not SANDBOX_WORKBOOK.is_file() or not SANDBOX_BINDINGS.is_dir():
        pytest.skip("sandbox advanced_formula_workbook fixtures missing")

    with tempfile.TemporaryDirectory() as tmp:
        module_dir = Path(tmp) / "advanced_formula_workbook"
        run_binding_checks(
            SANDBOX_WORKBOOK,
            SANDBOX_BINDINGS,
            module_dir=module_dir,
            package_name="advanced_formula_workbook",
            smoke_test=True,
        )


def test_normdist_sigma_band_compute_smoke() -> None:
    """Targeted smoke for ``compute_normdist_sigma_band`` after ``xl_abs`` support."""
    if not SANDBOX_WORKBOOK.is_file() or not SANDBOX_BINDINGS.is_dir():
        pytest.skip("sandbox advanced_formula_workbook fixtures missing")

    with tempfile.TemporaryDirectory() as tmp:
        module_dir = Path(tmp) / "advanced_formula_workbook"
        result = run_binding_checks(
            SANDBOX_WORKBOOK,
            SANDBOX_BINDINGS,
            module_dir=module_dir,
            package_name="advanced_formula_workbook",
            smoke_test=False,
        )
        import importlib
        import sys

        sys.path.insert(0, str(module_dir.parent))
        try:
            pkg = importlib.import_module("advanced_formula_workbook")
            ctx = pkg.make_context()
            records = pkg.compute_normdist_sigma_band(ctx=ctx)
            assert len(records) == 10
            labels = {record["OBS_VALUE"] for record in records}
            assert labels.issubset({"Within 1σ", "Within 2σ", "Outlier >2σ"})
        finally:
            sys.path.remove(str(module_dir.parent))
            for name in list(sys.modules):
                if name.startswith("advanced_formula_workbook"):
                    sys.modules.pop(name, None)
        assert result["report"]["ok"]
