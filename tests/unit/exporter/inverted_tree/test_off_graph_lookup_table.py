"""Issue 989 — INDEX/MATCH over a rectangle that includes an off-graph series.

`NamedAxes.plan` used to register only series with graph cells. Lookup
emission still walks the whole Excel range, so a constant with
`intersect_graph_leaves: false` raised `KeyError` from `NamedAxes.constant`.
"""

from __future__ import annotations

from pathlib import Path

from excel_grapher import (
    CodeGenerator,
    DynamicRefConfig,
    create_dependency_graph,
    load_series_bindings,
)
from excel_grapher.core.formula_ast import RangeNode
from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.ast_emit import EmitContext, _named_range_view
from excel_grapher.exporter.inverted_tree.catalog import build_catalog
from excel_grapher.exporter.inverted_tree.deps import SeriesDeps
from excel_grapher.exporter.inverted_tree.named_axes import NamedAxes
from tests.unit.exporter.inverted_tree.helpers import load_package

_BINDINGS = """
schema_version: 1.21.0
workbook: lookup.xlsx
concept_scheme:
  id: mcve
  concepts:
    - {id: OBS_VALUE, name: Observation value, dtype: number}
    - {id: INDICATOR, name: Indicator, dtype: string}
    - {id: PARAMETER, name: Parameter, dtype: string}
    - {id: TIME_PERIOD, name: Time period, dtype: string}
series:
  - id: lookup_key
    sheet: Calc
    data_range: Calc!A2
    layout: scalar
    input: {}
    domain: {from_workbook: true}
    structure:
      measure: {concept: OBS_VALUE, dtype: string, bind: {kind: data_cell, read: string}}
      dimensions: []
    key: []
  - id: column_key
    sheet: Calc
    data_range: Calc!B1
    layout: scalar
    input: {}
    domain: {from_workbook: true}
    structure:
      measure: {concept: OBS_VALUE, dtype: string, bind: {kind: data_cell, read: string}}
      dimensions: []
    key: []
  - id: looked_up
    sheet: Calc
    data_range: Calc!B2
    layout: scalar
    output:
      compute: {name: compute_looked_up, record_contract: records}
    structure:
      measure: {concept: OBS_VALUE, dtype: float, bind: {kind: data_cell, read: float}}
      dimensions: []
    key: []
  - id: codes
    sheet: Data
    data_range: Data!A2:A3
    layout: series
    constant: {}
    domain: {from_workbook: true}
    structure:
      measure: {concept: OBS_VALUE, dtype: string, bind: {kind: data_cell, read: string}}
      dimensions:
        - id: INDICATOR
          concept: INDICATOR
          role: key
          scope: cell
          bind: {kind: data_cell, read: string}
    key: [INDICATOR]
  - id: headers
    sheet: Data
    data_range: Data!A1:E1
    layout: series
    constant: {}
    domain: {from_workbook: true}
    structure:
      measure: {concept: OBS_VALUE, dtype: string, bind: {kind: data_cell, read: string}}
      dimensions:
        - id: PARAMETER
          concept: PARAMETER
          role: key
          scope: cell
          bind: {kind: data_cell, read: string}
    key: [PARAMETER]
  - id: values
    sheet: Data
    data_range: Data!D2:E3
    layout: matrix
    constant: {}
    domain: {from_workbook: true}
    structure:
      measure: {concept: OBS_VALUE, dtype: float, bind: {kind: data_cell, read: float}}
      dimensions:
        - id: INDICATOR
          concept: INDICATOR
          role: key
          scope: cell
          bind: {kind: row_label, label_column: A, read: string}
        - id: TIME_PERIOD
          concept: TIME_PERIOD
          role: key
          scope: cell
          bind: {kind: column_header, header_row: 1, read: string}
    key: [INDICATOR, TIME_PERIOD]
  - id: text_attributes
    sheet: Data
    data_range: Data!B2:C3
    layout: matrix
    constant: {}
    domain: {from_workbook: true}
    validation: {intersect_graph_leaves: false}
    structure:
      measure: {concept: OBS_VALUE, dtype: string, bind: {kind: data_cell, read: string}}
      dimensions:
        - id: INDICATOR
          concept: INDICATOR
          role: key
          scope: cell
          bind: {kind: row_label, label_column: A, read: string}
        - id: PARAMETER
          concept: PARAMETER
          role: key
          scope: cell
          bind: {kind: column_header, header_row: 1, read: string}
    key: [INDICATOR, PARAMETER]
""".lstrip()


def _write_mcve(tmp_path: Path) -> tuple[Path, Path]:
    import fastpyxl
    from fastpyxl.workbook.defined_name import DefinedName

    workbook_path = tmp_path / "lookup.xlsx"
    bindings_path = tmp_path / "model.bindings.yaml"
    workbook = fastpyxl.Workbook()
    data = workbook.active
    assert data is not None
    data.title = "Data"
    for col, value in enumerate(("Code", "Name", "Note", "Y2020", "Y2021"), start=1):
        data.cell(1, col, value)
    for row_index, row in enumerate(
        (("a", "alpha", "n1", 10, 20), ("b", "beta", "n2", 30, 40)),
        start=2,
    ):
        for col, value in enumerate(row, start=1):
            data.cell(row_index, col, value)
    calc = workbook.create_sheet("Calc")
    calc["A2"] = "a"
    calc["B1"] = "Y2020"
    calc["B2"] = "=INDEX(DATA,MATCH(A2,INDEX(DATA,,1),0),MATCH(B1,INDEX(DATA,1,),0))"
    workbook.defined_names.add(DefinedName(name="DATA", attr_text="Data!$A$1:$E$3"))
    workbook.save(workbook_path)
    bindings_path.write_text(_BINDINGS, encoding="utf-8")
    return workbook_path, bindings_path


def test_index_match_lowers_when_the_table_includes_an_off_graph_series(tmp_path: Path) -> None:
    workbook_path, bindings_path = _write_mcve(tmp_path)
    bindings = load_series_bindings(bindings_path)
    graph = create_dependency_graph(
        workbook_path,
        targets=["Calc!B2"],
        dynamic_refs=DynamicRefConfig.from_bindings(
            bindings, workbook_path, bindings_path=bindings_path
        ),
    )
    assert "Data!B2" not in graph and "Data!C2" not in graph
    assert "Data!D2" in graph

    with CodeGenerator(graph) as generator:
        modules = generator.generate_modules(
            series_bindings=bindings, bindings_workbook=workbook_path
        )

    internals = modules["internals.py"]
    assert "view(text_attributes" in internals or "text_attributes[" in internals
    assert "TEXT_ATTRIBUTES" in modules["data.py"]

    pkg = load_package(modules, tmp_path, name="off_graph_lookup")
    got = pkg.compute_looked_up(pkg.LookedUpInputs(lookup_key="a", column_key="Y2020"))
    expected = FormulaEvaluator(graph).evaluate(["Calc!B2"])["Calc!B2"]
    assert got == expected
    assert got == 10


def test_unplanned_lookup_axis_falls_back_instead_of_raising(tmp_path: Path) -> None:
    workbook_path, bindings_path = _write_mcve(tmp_path)
    bindings = load_series_bindings(bindings_path)
    graph = create_dependency_graph(
        workbook_path,
        targets=["Calc!B2"],
        dynamic_refs=DynamicRefConfig.from_bindings(
            bindings, workbook_path, bindings_path=bindings_path
        ),
    )
    catalog = build_catalog(bindings, workbook=workbook_path, graph=graph)
    host = catalog.get("looked_up")
    ctx = EmitContext(
        host=host,
        catalog=catalog,
        deps=SeriesDeps(
            host_id=host.series_id,
            param_ids=(),
            is_scan=False,
            seed_id=None,
            aligned_ids=frozenset(),
            lookup_ids=frozenset(),
            index_maps={},
            affine_maps={},
        ),
        host_index=0,
        host_cell=host.cells[0],
        coordinate_vars={},
        named_axes=NamedAxes.plan(()),
        graph=graph,
    )
    assert _named_range_view(RangeNode("Data!B2", "Data!C2"), ctx) is None
