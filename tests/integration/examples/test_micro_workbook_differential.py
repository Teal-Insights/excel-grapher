"""Integration checks for micro_workbook Excel-vs-evaluator differential parity."""

from __future__ import annotations

import math
import shutil
import tempfile
from dataclasses import dataclass
from itertools import product
from pathlib import Path
from typing import Any, Literal, Protocol

import pytest

from excel_grapher import DynamicRefConfig, XlError, create_dependency_graph
from excel_grapher.evaluator import FormulaEvaluator

REPO_ROOT = Path(__file__).resolve().parents[3]
WORKBOOK_PATH = REPO_ROOT / "examples" / "micro_workbook" / "example_cases.xlsx"
SHEET = "Sheet1"
INPUT_VALUES: tuple[int, ...] = (-1, 0, 1, 2, 10)
ATOL = 1e-9


@dataclass(frozen=True)
class Case:
    label: str
    target: str
    leaves: tuple[str, ...]
    skip_reason: str | None = None
    oracle_graph_kwargs: dict[str, Any] | None = None
    leaf_domains: dict[str, tuple[Any, ...]] | None = None


@dataclass(frozen=True)
class Trial:
    case_label: str
    target: str
    inputs: dict[str, Any]
    golden: Any
    oracle: Any
    match: bool
    abs_diff: float | None
    rel_diff: float | None
    note: str = ""


def _sheet_addr(address: str) -> str:
    return address if "!" in address else f"{SHEET}!{address}"


def _case(
    label: str,
    target: str,
    *leaves: str,
    skip: str | None = None,
    oracle_graph_kwargs: dict[str, Any] | None = None,
    leaf_domains: dict[str, tuple[Any, ...]] | None = None,
) -> Case:
    qualified_leaves = tuple(_sheet_addr(leaf) for leaf in leaves)
    qualified_domains = (
        {_sheet_addr(key): values for key, values in leaf_domains.items()} if leaf_domains else None
    )
    return Case(
        label=label,
        target=_sheet_addr(target),
        leaves=qualified_leaves,
        skip_reason=skip,
        oracle_graph_kwargs=dict(oracle_graph_kwargs) if oracle_graph_kwargs else None,
        leaf_domains=qualified_domains,
    )


_CYCLE_SKIP = "circular reference; iterative calculation required on both sides"


_ROW10_CONSTRAINT_SCHEMA: dict[str, Any] = {f"{SHEET}!B10": Literal[0, 1]}
_ROW10_DYNAMIC_REFS = DynamicRefConfig.from_constraints(_ROW10_CONSTRAINT_SCHEMA, {})

CASES: tuple[Case, ...] = (
    _case("Formula with no dependencies", "B1", skip="no input leaves to vary"),
    _case("Linear dependency", "C2", "B2"),
    _case("Conditional branches", "E3", "B3", "C3", "D3"),
    _case("Nested conditional in a cell", "D4", "B4", "C4"),
    _case("Nested conditional across cells", "E5", "A5", "B5", "C5"),
    _case("Will cycle", "C6", skip=_CYCLE_SKIP),
    _case("Won't cycle", "C7", "B7", skip=_CYCLE_SKIP),
    _case("May cycle", "C8", "B8", skip=_CYCLE_SKIP),
    _case(
        "OFFSET/INDIRECT reference resolution with scalar arguments",
        "D9",
        "C9",
    ),
    _case(
        "OFFSET/INDIRECT reference resolution with dynamic arguments",
        "E10",
        "B10",
        "C10",
        "D10",
        oracle_graph_kwargs={"dynamic_refs": _ROW10_DYNAMIC_REFS},
        leaf_domains={"B10": (0, 1)},
    ),
)


def _coerce_excel_error(value: Any) -> Any:
    if isinstance(value, str):
        err = XlError.from_text(value)
        if err is not None:
            return err
    return value


def values_match(g: Any, o: Any, atol: float) -> tuple[bool, float | None, float | None, str]:
    g = _coerce_excel_error(g)
    o = _coerce_excel_error(o)

    if isinstance(g, XlError) or isinstance(o, XlError):
        if isinstance(g, XlError) and isinstance(o, XlError) and g == o:
            return True, None, None, f"both error: {g}"
        return False, None, None, f"error mismatch (golden={g!r}, oracle={o!r})"
    if g is None and o is None:
        return True, None, None, "both None"
    if g is None or o is None:
        return False, None, None, "one side None"
    if isinstance(g, bool) and isinstance(o, bool):
        return (g == o), None, None, "bool"
    if isinstance(g, int | float) and isinstance(o, int | float):
        gf, of = float(g), float(o)
        if math.isnan(gf) and math.isnan(of):
            return True, None, None, "both NaN"
        abs_d = abs(gf - of)
        rel_d = abs_d / max(abs(gf), abs(of), 1e-15)
        return (abs_d <= atol), abs_d, rel_d, ""
    return (g == o), None, None, "exact"


class GoldenDriver:
    def __init__(self, workbook_path: Path) -> None:
        import xlwings as xw

        self._tmpdir = Path(tempfile.mkdtemp(prefix="excel_grapher_diff_"))
        self._tmp_wb = self._tmpdir / workbook_path.name
        shutil.copy2(workbook_path, self._tmp_wb)
        self._app = xw.App(visible=False, add_book=False)
        self._app.display_alerts = False
        self._app.screen_updating = False
        self._book = self._app.books.open(str(self._tmp_wb))

    def set_inputs(self, inputs: dict[str, Any]) -> None:
        for key, val in inputs.items():
            sheet, addr = key.split("!", 1)
            self._book.sheets[sheet].range(addr).value = val
        self._app.calculate()

    def read(self, target: str) -> Any:
        sheet, addr = target.split("!", 1)
        return self._book.sheets[sheet].range(addr).value

    def close(self) -> None:
        try:
            self._book.close()
        finally:
            try:
                self._app.quit()
            finally:
                shutil.rmtree(self._tmpdir, ignore_errors=True)


class _GoldenForRunCase(Protocol):
    def set_inputs(self, inputs: dict[str, Any]) -> None: ...
    def read(self, target: str) -> Any: ...


class OracleDriver:
    def __init__(
        self, workbook_path: Path, target: str, graph_kwargs: dict[str, Any] | None = None
    ) -> None:
        graph_kwargs = graph_kwargs or {}
        self._graph = create_dependency_graph(
            workbook_path, [target], load_values=True, **graph_kwargs
        )
        self._evaluator = FormulaEvaluator(self._graph)
        self._target = target

    def set_inputs(self, inputs: dict[str, Any]) -> None:
        for key, val in inputs.items():
            self._graph.set_node_value(key, val)

    def read(self) -> Any:
        return self._evaluator.evaluate(self._target)


def run_case(case: Case, golden: _GoldenForRunCase) -> list[Trial]:
    oracle = OracleDriver(WORKBOOK_PATH, case.target, graph_kwargs=case.oracle_graph_kwargs)
    trials: list[Trial] = []
    leaf_domains = case.leaf_domains or {}
    domain_values = tuple(leaf_domains.get(leaf, INPUT_VALUES) for leaf in case.leaves)
    for combo in product(*domain_values):
        inputs = dict(zip(case.leaves, combo, strict=True))
        golden.set_inputs(inputs)
        g_val = golden.read(case.target)
        oracle.set_inputs(inputs)
        try:
            o_val: Any = oracle.read()
        except Exception as exc:  # pragma: no cover - defensive guard for reporting
            o_val = f"<{type(exc).__name__}: {exc}>"
            match, abs_d, rel_d, note = False, None, None, "oracle raised"
        else:
            match, abs_d, rel_d, note = values_match(g_val, o_val, ATOL)
        trials.append(
            Trial(
                case_label=case.label,
                target=case.target,
                inputs=inputs,
                golden=g_val,
                oracle=o_val,
                match=match,
                abs_diff=abs_d,
                rel_diff=rel_d,
                note=note,
            )
        )
    return trials


def run_differential_sweep() -> tuple[list[Trial], list[Case]]:
    runnable = [c for c in CASES if c.skip_reason is None and c.leaves]
    skipped = [c for c in CASES if c not in runnable]
    golden = GoldenDriver(WORKBOOK_PATH)
    trials: list[Trial] = []
    try:
        for case in runnable:
            trials.extend(run_case(case, golden))
    finally:
        golden.close()
    return trials, skipped


def test_case_factory_supports_graph_kwargs_and_leaf_domains() -> None:
    case = _case(
        "dynamic refs with constraints",
        "E10",
        "B10",
        "C10",
        oracle_graph_kwargs={"use_cached_dynamic_refs": True},
        leaf_domains={"B10": (0, 1), "C10": (5,)},
    )
    assert case.oracle_graph_kwargs == {"use_cached_dynamic_refs": True}
    assert case.leaf_domains == {"Sheet1!B10": (0, 1), "Sheet1!C10": (5,)}


def test_oracle_driver_forwards_graph_kwargs_to_graph_builder(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, Any] = {}

    class _DummyGraph:
        def set_node_value(self, _key: str, _val: Any) -> None:
            return None

    def _fake_graph_builder(
        workbook_path: Path,
        targets: list[str],
        *,
        load_values: bool,
        **kwargs: Any,
    ) -> _DummyGraph:
        captured["workbook_path"] = workbook_path
        captured["targets"] = targets
        captured["load_values"] = load_values
        captured["kwargs"] = kwargs
        return _DummyGraph()

    monkeypatch.setattr(
        "tests.integration.examples.test_micro_workbook_differential.create_dependency_graph",
        _fake_graph_builder,
    )
    monkeypatch.setattr(
        "tests.integration.examples.test_micro_workbook_differential.FormulaEvaluator",
        lambda _graph: object(),
    )
    _ = OracleDriver(
        Path("dummy.xlsx"),
        "Sheet1!E10",
        graph_kwargs={"use_cached_dynamic_refs": True},
    )
    assert captured["targets"] == ["Sheet1!E10"]
    assert captured["load_values"] is True
    assert captured["kwargs"] == {"use_cached_dynamic_refs": True}


def test_run_case_uses_per_leaf_domains_for_cartesian_sweep(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    seen: list[dict[str, Any]] = []

    class _FakeGolden:
        def __init__(self) -> None:
            self._last_inputs: dict[str, Any] = {}

        def set_inputs(self, inputs: dict[str, Any]) -> None:
            self._last_inputs = dict(inputs)

        def read(self, target: str) -> int:
            return 7

    class _FakeOracle:
        def __init__(
            self,
            _workbook_path: Path,
            _target: str,
            graph_kwargs: dict[str, Any] | None = None,
        ) -> None:
            self._last_inputs: dict[str, Any] = {}
            self.graph_kwargs = graph_kwargs

        def set_inputs(self, inputs: dict[str, Any]) -> None:
            self._last_inputs = dict(inputs)
            seen.append(dict(inputs))

        def read(self) -> int:
            return 7

    monkeypatch.setattr(
        "tests.integration.examples.test_micro_workbook_differential.OracleDriver",
        _FakeOracle,
    )
    golden = _FakeGolden()
    case = _case(
        "dynamic args domain sweep",
        "E10",
        "B10",
        "C10",
        leaf_domains={"B10": (0, 1), "C10": (5,)},
    )
    trials = run_case(case, golden)
    assert len(trials) == 2
    assert seen == [
        {"Sheet1!B10": 0, "Sheet1!C10": 5},
        {"Sheet1!B10": 1, "Sheet1!C10": 5},
    ]


def _format_trial(trial: Trial) -> str:
    target = trial.target.split("!", 1)[1] if "!" in trial.target else trial.target
    return (
        f"{trial.case_label} {target} inputs={trial.inputs!r} "
        f"golden={trial.golden!r} oracle={trial.oracle!r} note={trial.note!r}"
    )


@pytest.mark.slow
def test_micro_workbook_excel_parity_run_if_available() -> None:
    if not WORKBOOK_PATH.exists():
        pytest.skip(f"Test workbook not found at {WORKBOOK_PATH}")
    try:
        import xlwings  # noqa: F401
    except ImportError as exc:
        pytest.skip(f"xlwings not available: {exc}")
    try:
        trials, skipped = run_differential_sweep()
    except Exception as exc:  # pragma: no cover - environment dependent
        if any(msg in str(exc).lower() for msg in ("excel", "com", "active x")):
            pytest.skip(f"Excel automation not available: {exc}")
        raise
    assert trials, "No runnable differential trials were produced."
    assert skipped, "Expected at least one documented skipped case."
    failures = [trial for trial in trials if not trial.match]
    if failures:
        details = "\n".join(_format_trial(trial) for trial in failures[:10])
        pytest.fail(f"{len(failures)} differential mismatches:\n{details}")
