"""Public named computation contract.

An exported package must compute its public outputs through the generated
named formula bodies. The private flat modules are not part of the package,
and no public entry point may execute code from them. The witness records
every generated function that runs while a public output is computed.
"""

from __future__ import annotations

import sys
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from types import FrameType, ModuleType

import pytest

from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from excel_grapher.series_bindings.load import load_series_bindings
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.paths import INVERTED_TREE_TINY_DSA
from tests.unit.exporter.inverted_tree.helpers import load_package
from tests.unit.exporter.inverted_tree.local_corpus import load_constraints_module

_WORKBOOK = INVERTED_TREE_TINY_DSA / "tiny-dsa.xlsx"
_BINDINGS_DIR = INVERTED_TREE_TINY_DSA / "bindings"

PRIVATE_FLAT_MODULES = frozenset({"_kernel.py", "_kernels.py", "_data.py", "_tensor_lowering.py"})

_EXPECTED_SHOCKED = (
    61.28985507246378,
    63.29945741415009,
    64.85855735045921,
    65.95605876303210,
    66.58059223010186,
)

# Every formula series on the path from the inputs to `output_shocked`.
_SHOCKED_CLOSURE = frozenset(
    {
        "initial_debt_resolved",
        "engine_initial_debt_shocked",
        "shock_magnitude_resolved",
        "shock_active",
        "shocked_growth",
        "shocked_interest",
        "shocked_primary_balance",
        "shocked_path_internal",
        "output_shocked",
    }
)


@contextmanager
def execution_witness(package_dir: Path) -> Iterator[list[tuple[str, str]]]:
    """Record `(module file, function)` for every generated call in `package_dir`."""
    calls: list[tuple[str, str]] = []
    root = str(package_dir.resolve())

    def tracer(frame: FrameType, event: str, _arg: object) -> object:
        if event == "call":
            filename = frame.f_code.co_filename
            if filename.startswith(root):
                calls.append((Path(filename).name, frame.f_code.co_name))
        return None

    previous = sys.gettrace()
    sys.settrace(tracer)
    try:
        yield calls
    finally:
        sys.settrace(previous)


@pytest.fixture(scope="module")
def tiny_dsa_modules() -> dict[str, str]:
    constraints = load_constraints_module(INVERTED_TREE_TINY_DSA / "constraints.py")
    assert constraints is not None
    bindings = load_series_bindings(_BINDINGS_DIR)
    targets = all_series_targets(bindings, workbook=_WORKBOOK)
    graph = create_dependency_graph(
        _WORKBOOK,
        targets,
        load_values=True,
        dynamic_refs=DynamicRefConfig.from_constraints(constraints.CONSTRAINTS, {}),
    )
    with CodeGenerator(graph) as generator:
        return generator.generate_modules(series_bindings=bindings, bindings_workbook=_WORKBOOK)


def _public_package(modules: dict[str, str], tmp_path: Path, name: str) -> ModuleType:
    public = {file: source for file, source in modules.items() if file not in PRIVATE_FLAT_MODULES}
    return load_package(public, tmp_path, name=name)


def test_package_has_no_private_flat_modules(tiny_dsa_modules: dict[str, str]) -> None:
    assert not (set(tiny_dsa_modules) & PRIVATE_FLAT_MODULES)
    for file, source in tiny_dsa_modules.items():
        assert "_kernel" not in source, file
        assert "CoordinateBuffer" not in source, file


def test_public_output_runs_named_formula_bodies_without_flat_modules(
    tiny_dsa_modules: dict[str, str], tmp_path: Path
) -> None:
    package = _public_package(tiny_dsa_modules, tmp_path, "tiny_dsa_public_only")
    data = package.data
    with execution_witness(tmp_path / "tiny_dsa_public_only") as calls:
        result = package.compute_output_shocked(
            country_name=data.COUNTRY_NAME_DEFAULT,
            country_initial_debt=data.COUNTRY_INITIAL_DEBT_DEFAULT,
            growth_baseline=data.GROWTH_BASELINE_DEFAULT,
            interest_baseline=data.INTEREST_BASELINE_DEFAULT,
            primary_balance_baseline=data.PRIMARY_BALANCE_BASELINE_DEFAULT,
            shock_year=data.SHOCK_YEAR_DEFAULT,
            shock_type=data.SHOCK_TYPE_DEFAULT,
            shock_magnitudes=data.SHOCK_MAGNITUDES_DEFAULT,
        )
    values = tuple(result[year] for year in (1, 2, 3, 4, 5))
    assert values == pytest.approx(_EXPECTED_SHOCKED)
    private_frames = sorted({file for file, _ in calls if file.startswith("_")})
    assert not private_frames, private_frames
    executed = {function for file, function in calls if file == "internals.py"}
    assert executed >= _SHOCKED_CLOSURE, sorted(_SHOCKED_CLOSURE - executed)


_STANDALONE_PROBE = """
import sys


class _Block:
    def find_spec(self, name, path=None, target=None):
        if name == "excel_grapher" or name.startswith("excel_grapher."):
            raise ImportError("excel_grapher is not installed here")
        return None


sys.meta_path.insert(0, _Block())
sys.path.insert(0, __import__("os").getcwd())
import tiny_dsa_standalone as package

data = package.data
result = package.compute_output_shocked(
    country_name=data.COUNTRY_NAME_DEFAULT,
    country_initial_debt=data.COUNTRY_INITIAL_DEBT_DEFAULT,
    growth_baseline=data.GROWTH_BASELINE_DEFAULT,
    interest_baseline=data.INTEREST_BASELINE_DEFAULT,
    primary_balance_baseline=data.PRIMARY_BALANCE_BASELINE_DEFAULT,
    shock_year=data.SHOCK_YEAR_DEFAULT,
    shock_type=data.SHOCK_TYPE_DEFAULT,
    shock_magnitudes=data.SHOCK_MAGNITUDES_DEFAULT,
)
print(repr(tuple(result[year] for year in (1, 2, 3, 4, 5))))
print(sorted(name for name in sys.modules if name.startswith("excel_grapher")))
"""


def test_package_runs_without_excel_grapher_installed(
    tiny_dsa_modules: dict[str, str], tmp_path: Path
) -> None:
    import ast
    import subprocess

    for file, source in tiny_dsa_modules.items():
        for node in ast.walk(ast.parse(source)):
            if isinstance(node, ast.ImportFrom):
                assert not (node.module or "").startswith("excel_grapher"), file
            elif isinstance(node, ast.Import):
                assert not any(alias.name.startswith("excel_grapher") for alias in node.names), file
    package_dir = tmp_path / "tiny_dsa_standalone"
    package_dir.mkdir()
    for file, source in tiny_dsa_modules.items():
        (package_dir / file).write_text(source, encoding="utf-8")
    completed = subprocess.run(
        [sys.executable, "-I", "-c", _STANDALONE_PROBE],
        cwd=tmp_path,
        capture_output=True,
        text=True,
        check=False,
    )
    assert completed.returncode == 0, completed.stderr
    values, loaded = completed.stdout.strip().splitlines()
    assert eval(values) == pytest.approx(_EXPECTED_SHOCKED)
    assert loaded == "[]"
