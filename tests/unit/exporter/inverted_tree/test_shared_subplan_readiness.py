"""Shared source factoring must preserve producer-before-consumer evaluation."""

from pathlib import Path
from typing import Any

from excel_grapher.exporter.inverted_tree.emit import _emit_evaluation_body, _SharedSubplan
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    inverted_graph_parts,
    series_entry,
    write_workbook,
)


def test_shared_subplan_waits_for_external_formula_inputs(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "readiness.xlsx",
        {
            "Model": {
                "A1": 10,
                "A2": "=A1+1",
                "A3": "=A1+2",
                "A4": "=A2+A3",
            }
        },
    )
    document = bindings_document(
        series_entry("source", "Model!A1", direction="input"),
        series_entry("early", "Model!A2", direction="internal"),
        series_entry("middle", "Model!A3", direction="internal"),
        series_entry("late", "Model!A4", direction="output"),
    )
    catalog, deps, _graph = inverted_graph_parts(workbook, document)
    plan = _SharedSubplan(
        "_shared", ("early", "late"), frozenset({"late"}), ("late",), ("source", "middle"), {}, {}
    )
    body, _, _ = _emit_evaluation_body(
        leaves=("source",),
        formula_ids=("early", "middle", "late"),
        result_indices={},
        call_indices={},
        catalog=catalog,
        deps=deps,
        scc_map={},
        subplans=(plan,),
    )
    source = "\n".join(body)
    assert "_shared(" not in source
    assert source.index("early =") < source.index("middle =") < source.index("late =")


def test_shared_subplan_does_not_gather_leaves_behind_supplied_formula(tmp_path: Path) -> None:
    from types import SimpleNamespace

    from excel_grapher.exporter.inverted_tree.emit import _emit_shared_subplan
    from excel_grapher.exporter.inverted_tree.runtime import take

    workbook = write_workbook(
        tmp_path / "boundary.xlsx",
        {
            "M": {
                "A1": 2020,
                "B1": 2021,
                "C1": 2022,
                "A2": 1,
                "B2": 2,
                "C2": 3,
                "A3": "=SUM(A2:C2)",
                "A4": "=A3+1",
            }
        },
    )
    document = bindings_document(
        series_entry("source", "M!A2:C2", layout="series", direction="input", header_row=1),
        series_entry("upstream", "M!A3", direction="internal"),
        series_entry("late", "M!A4", direction="output"),
    )
    catalog, deps, _ = inverted_graph_parts(workbook, document)
    plan = _SharedSubplan(
        "_shared", ("late",), frozenset({"late"}), ("late",), ("upstream",), {"source": (1, 2)}, {}
    )
    source, _, _ = _emit_shared_subplan(
        plan, catalog=catalog, deps=deps, scc_map={}, subplans=(plan,)
    )
    namespace: dict[str, Any] = {
        "internals": SimpleNamespace(late=lambda upstream: upstream + 1),
        "take": take,
    }
    exec(source, namespace)
    assert namespace["_shared"](upstream=6.0) == 7.0


def test_shared_subplan_requires_its_declared_producer_window(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "shared_windows.xlsx",
        {
            "M": {
                "A1": 2020,
                "B1": 2021,
                "C1": 2022,
                "D1": 2023,
                "A2": 1,
                "B2": 2,
                "C2": 3,
                "D2": 4,
                "A3": "=A2",
                "B3": "=B2",
                "C3": "=C2",
                "D3": "=D2",
                "B4": "=B3",
                "D4": "=D3",
            }
        },
    )
    document = bindings_document(
        series_entry("source", "M!A2:D2", layout="series", direction="input", header_row=1),
        series_entry("upstream", "M!A3:D3", layout="series", direction="internal", header_row=1),
        {
            **series_entry("late", "M!B4:D4", layout="series", direction="output", header_row=1),
            "data_range": ["M!B4", "M!D4"],
        },
    )
    catalog, deps, _ = inverted_graph_parts(workbook, document)
    plan = _SharedSubplan(
        "_shared",
        ("late",),
        frozenset({"late"}),
        ("late",),
        ("upstream",),
        {"upstream": (1, 3), "late": (0, 1)},
        {"late": (0, 1)},
    )
    body, _, _ = _emit_evaluation_body(
        leaves=(),
        formula_ids=("late",),
        result_indices=plan.result_indices,
        call_indices=plan.call_indices,
        catalog=catalog,
        deps=deps,
        scc_map={},
        subplans=(plan,),
        bound_windows={"upstream": (1, 3)},
    )
    assert "_shared(" not in "\n".join(body)
