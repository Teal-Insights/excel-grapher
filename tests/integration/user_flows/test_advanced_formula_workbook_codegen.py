"""Codegen health and setter propagation for ``advanced_formula_workbook.xlsx``."""

from __future__ import annotations

import importlib
import re
import sys
import tempfile
import types
from pathlib import Path

import pytest

from excel_grapher.grapher import DependencyGraph
from excel_grapher.series_bindings.types import WorkbookSeriesBindings
from excel_grapher.series_bindings.workflow import (
    BindingsCheckResult,
    generate_bindings_modules,
    validate_bindings_workbook,
)
from tests.integration.user_flows.advanced_formula_workbook_cases import (
    BINDINGS_DIR,
    DOWNSTREAM_UPDATE_CASES,
    WORKBOOK,
)
from tests.integration.user_flows.bindings_accuracy import (
    DownstreamUpdateCase,
    assert_downstream_update,
    generate_bindings_namespace,
)

_XLUDF_RUNTIME_PATTERN = re.compile(r"xl__xludf_[a-z0-9_]+")
_EXPECTED_MODULE_FILES = frozenset(
    {"__init__.py", "api.py", "data.py", "internals.py", "runtime.py"},
)
_PACKAGE_NAME = "advanced_formula_workbook"


def _skip_if_fixtures_missing() -> None:
    if not WORKBOOK.is_file() or not BINDINGS_DIR.is_dir():
        pytest.skip("sandbox advanced_formula_workbook fixtures missing")


@pytest.fixture(scope="module")
def sandbox_validation() -> BindingsCheckResult:
    _skip_if_fixtures_missing()
    result = validate_bindings_workbook(WORKBOOK, BINDINGS_DIR)
    assert result["report"]["ok"], result["report"]["issues"]
    return result


@pytest.fixture(scope="module")
def sandbox_bindings(sandbox_validation: BindingsCheckResult) -> WorkbookSeriesBindings:
    return sandbox_validation["bindings"]


@pytest.fixture(scope="module")
def sandbox_graph(sandbox_validation: BindingsCheckResult) -> DependencyGraph:
    return sandbox_validation["graph"]


@pytest.fixture(scope="module")
def sandbox_namespace(
    sandbox_graph: DependencyGraph,
    sandbox_bindings: WorkbookSeriesBindings,
) -> dict[str, object]:
    return generate_bindings_namespace(sandbox_graph, WORKBOOK, sandbox_bindings)


@pytest.fixture(scope="module")
def sandbox_module_files(sandbox_validation: BindingsCheckResult) -> dict[str, str]:
    return generate_bindings_modules(
        sandbox_validation["graph"],
        targets=sandbox_validation["targets"],
        bindings=sandbox_validation["bindings"],
        workbook=WORKBOOK,
    )


def _import_generated_package(
    files: dict[str, str],
    *,
    parent_dir: Path,
    package_name: str,
) -> types.ModuleType:
    module_dir = parent_dir / package_name
    module_dir.mkdir(parents=True, exist_ok=True)
    for filename, content in files.items():
        (module_dir / filename).write_text(content, encoding="utf-8")

    sys.path.insert(0, str(parent_dir))
    for name in list(sys.modules):
        if name == package_name or name.startswith(f"{package_name}."):
            sys.modules.pop(name, None)
    return importlib.import_module(package_name)


def _cleanup_generated_package(*, parent_dir: Path, package_name: str) -> None:
    parent = str(parent_dir)
    if parent in sys.path:
        sys.path.remove(parent)
    for name in list(sys.modules):
        if name == package_name or name.startswith(f"{package_name}."):
            sys.modules.pop(name, None)


def test_sandbox_modular_export_writes_expected_files(
    sandbox_module_files: dict[str, str],
) -> None:
    """Modular codegen emits the five-file package layout."""
    assert set(sandbox_module_files) == _EXPECTED_MODULE_FILES


def test_sandbox_modular_export_has_no_xludf_runtime_symbols(
    sandbox_module_files: dict[str, str],
) -> None:
    """Full-workbook modular export must not reference ``xl__xludf_*`` names."""
    combined = "\n".join(sandbox_module_files.values())
    assert _XLUDF_RUNTIME_PATTERN.search(combined) is None


def test_sandbox_modular_export_imports_and_compute_all(
    sandbox_module_files: dict[str, str],
) -> None:
    """Generated package imports cleanly and ``compute_all`` runs without error."""
    with tempfile.TemporaryDirectory() as tmp:
        parent_dir = Path(tmp)
        pkg = _import_generated_package(
            sandbox_module_files,
            parent_dir=parent_dir,
            package_name=_PACKAGE_NAME,
        )
        try:
            make_context = pkg.make_context
            compute_all = pkg.compute_all
            ctx = make_context()
            results = compute_all(ctx=ctx)
            assert isinstance(results, dict)
            assert len(results) > 0
        finally:
            _cleanup_generated_package(parent_dir=parent_dir, package_name=_PACKAGE_NAME)


@pytest.mark.parametrize(
    "update_case",
    DOWNSTREAM_UPDATE_CASES,
    ids=[f"{case.setter_name}->{case.compute_name}" for case in DOWNSTREAM_UPDATE_CASES],
)
def test_sandbox_downstream_setter_propagation(
    sandbox_namespace: dict[str, object],
    update_case: DownstreamUpdateCase,
) -> None:
    """Priority setter writes change downstream compute observations."""
    assert_downstream_update(sandbox_namespace, update_case)
