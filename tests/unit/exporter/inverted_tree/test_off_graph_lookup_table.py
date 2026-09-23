"""Issue 989 — INDEX/MATCH over a rectangle that includes an off-graph series.

`NamedAxes.plan` used to register only series with graph cells. Lookup
emission still walks the whole Excel range, so a constant with
`intersect_graph_leaves: false` raised `KeyError` from `NamedAxes.constant`.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

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


_DISJOINT_TAIL = """
  - id: tags
    sheet: Tags
    data_range: Tags!A1:A2
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
  - id: tag_echo
    sheet: Calc
    data_range: Calc!D2
    layout: scalar
    output:
      compute: {name: compute_tag_echo, record_contract: records}
    structure:
      measure: {concept: OBS_VALUE, dtype: string, bind: {kind: data_cell, read: string}}
      dimensions: []
    key: []
"""


def _write_disjoint_mcve(tmp_path: Path) -> tuple[Path, Path]:
    """MCVE plus an on-graph PARAMETER series whose keys are not table headers."""
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
    tags = workbook.create_sheet("Tags")
    tags["A1"] = "EcDatabase"
    tags["A2"] = "Country"
    calc = workbook.create_sheet("Calc")
    calc["A2"] = "a"
    calc["B1"] = "Y2020"
    calc["B2"] = "=INDEX(DATA,MATCH(A2,INDEX(DATA,,1),0),MATCH(B1,INDEX(DATA,1,),0))"
    calc["D2"] = "=Tags!A1"
    workbook.defined_names.add(DefinedName(name="DATA", attr_text="Data!$A$1:$E$3"))
    workbook.save(workbook_path)
    bindings_path.write_text(_BINDINGS + _DISJOINT_TAIL, encoding="utf-8")
    return workbook_path, bindings_path


_INPUT_BINDINGS = """
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
  - id: text_attributes
    sheet: Data
    data_range: Data!B2:C3
    layout: matrix
    input: {}
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
""".lstrip()


def _write_input_mcve(tmp_path: Path) -> tuple[Path, Path]:
    """Same rectangle, with the off-graph body declared as an input first."""
    workbook_path, _bindings_path = _write_mcve(tmp_path)
    bindings_path = tmp_path / "model.bindings.yaml"
    bindings_path.write_text(_INPUT_BINDINGS, encoding="utf-8")
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
    table = _eval_lookup_table(pkg, internals)
    assert pkg.excel.xl_index(table, 2, 2) == "alpha"
    assert pkg.excel.xl_index(table, 2, 3) == "n1"
    assert pkg.excel.xl_index(table, 2, 4) == 10


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


def _eval_lookup_table(pkg: Any, internals: str, name: str = "_looked_up_table_0") -> object:
    """Evaluate one generated lookup table against the package's series."""
    marker = f"{name} = "
    expression = next(
        (
            line.strip()[len(marker) :]
            for line in internals.splitlines()
            if line.strip().startswith(marker)
        ),
        None,
    )
    assert expression is not None
    namespace = {
        "lazy_table": pkg.runtime.lazy_table,
        "view": pkg.runtime.view,
        "span": pkg.runtime.span,
        "data": pkg.data,
        "headers": pkg.data.HEADERS,
        "codes": pkg.data.CODES,
        "values": pkg.data.VALUES,
        "text_attributes": pkg.data.TEXT_ATTRIBUTES,
    }
    return eval(expression, namespace)  # noqa: S307 — expression emitted by this package


def _graph_and_modules(tmp_path: Path, workbook_path: Path, bindings_path: Path, *targets: str):
    bindings = load_series_bindings(bindings_path)
    graph = create_dependency_graph(
        workbook_path,
        targets=list(targets),
        dynamic_refs=DynamicRefConfig.from_bindings(
            bindings, workbook_path, bindings_path=bindings_path
        ),
    )
    with CodeGenerator(graph) as generator:
        modules = generator.generate_modules(
            series_bindings=bindings, bindings_workbook=workbook_path
        )
    return bindings, graph, modules


def test_disjoint_parameter_keys_keep_lookup_positions(tmp_path: Path) -> None:
    """An on-graph PARAMETER axis with foreign keys must not widen the table."""
    workbook_path, bindings_path = _write_disjoint_mcve(tmp_path)
    _bindings, graph, modules = _graph_and_modules(
        tmp_path, workbook_path, bindings_path, "Calc!B2", "Calc!D2"
    )
    assert "Data!B2" not in graph and "Data!C2" not in graph
    assert "Tags!A1" in graph
    internals = modules["internals.py"]
    assert "EcDatabase" in modules["data.py"]
    pkg = load_package(modules, tmp_path, name="off_graph_disjoint")
    assert pkg.compute_looked_up(pkg.LookedUpInputs(lookup_key="a", column_key="Y2020")) == 10
    assert pkg.compute_tag_echo(pkg.TagEchoInputs()) == "EcDatabase"
    table = _eval_lookup_table(pkg, internals)
    assert pkg.excel.xl_index(table, 2, 2) == "alpha"
    assert pkg.excel.xl_index(table, 2, 4) == 10
    assert pkg.excel.xl_index(table, 1, 1) == "Code"
    assert "EcDatabase" not in [pkg.excel.xl_index(table, 1, col) for col in range(1, 6)]


def test_unplanned_axis_callbacks_read_off_graph_values(tmp_path: Path) -> None:
    """Per-cell callbacks, not only `None`, lower an axis `plan` never registered."""
    from excel_grapher.exporter.inverted_tree.deps import collect_all_deps
    from excel_grapher.exporter.inverted_tree.named_emit import _retained, _semantic_body

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
    deps = collect_all_deps(catalog, graph)
    named_axes = NamedAxes.plan(
        axis
        for series in _retained(catalog)
        if not series.single_valued
        for axis in series.tensor_domain.axes
    )
    host = catalog.get("looked_up")
    lines, _used = _semantic_body(host, catalog, deps[host.series_id], graph, named_axes)
    source = "\n".join(lines)
    assert "lambda:" in source
    assert "text_attributes[" in source
    assert "view(text_attributes" not in source

    _bindings, _graph, modules = _graph_and_modules(
        tmp_path, workbook_path, bindings_path, "Calc!B2"
    )
    pkg = load_package(modules, tmp_path, name="off_graph_callbacks")
    namespace: dict[str, Any] = {
        "lazy_table": pkg.runtime.lazy_table,
        "view": pkg.runtime.view,
        "span": pkg.runtime.span,
        "xl_index": pkg.excel.xl_index,
        "xl_match": pkg.excel.xl_match,
        "as_measure": pkg.excel.as_measure,
        "data": pkg.data,
        "headers": pkg.data.HEADERS,
        "codes": pkg.data.CODES,
        "values": pkg.data.VALUES,
        "text_attributes": pkg.data.TEXT_ATTRIBUTES,
        "lookup_key": "a",
        "column_key": "Y2020",
    }
    exec("def run():\n" + source, namespace)  # noqa: S102 — generated function body
    assert namespace["run"]() == 10
    table = _eval_lookup_table(pkg, source, name="_looked_up_table_0")
    assert pkg.excel.xl_index(table, 2, 2) == "alpha"


def test_unplanned_runtime_axis_emits_a_literal_key(tmp_path: Path) -> None:
    from unittest.mock import patch

    from excel_grapher.exporter.inverted_tree.ast_emit import _named_keys

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
    owner = catalog.get("text_attributes")
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
    point = owner.domain[0]
    # SeriesCatalog is frozen, so the labeller is injected on the emit context.
    with patch.object(ctx, "runtime_labeller_for", return_value=host):
        keys = _named_keys(owner, point, ctx)
    assert keys == [repr(point[field]) for field in owner.key_fields]


def test_fingerprint_required_covers_emitted_off_graph_domain(tmp_path: Path) -> None:
    from unittest.mock import patch

    from excel_grapher.exporter.inverted_tree import named_emit
    from excel_grapher.exporter.inverted_tree.named_emit import named_codegen_fingerprint

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
    captured: dict[str, object] = {}
    real_dumps = named_emit.json.dumps

    def spy(payload: object, **kwargs: object) -> str:
        captured["payload"] = payload
        return real_dumps(payload, **kwargs)

    with patch.object(named_emit.json, "dumps", spy):
        named_codegen_fingerprint(catalog, include=("text_attributes",))
    payload = captured["payload"]
    assert isinstance(payload, dict)
    series = payload["series"]
    assert isinstance(series, list)
    entry = next(
        item
        for item in series
        if isinstance(item, dict) and item.get("series_id") == "text_attributes"
    )
    required = entry["required"]
    assert isinstance(required, list)
    assert len(required) == len(catalog.get("text_attributes").tensor_domain)


def test_off_graph_input_is_checked_in_catalog_order(tmp_path: Path) -> None:
    workbook_path, bindings_path = _write_input_mcve(tmp_path)
    _bindings, graph, modules = _graph_and_modules(
        tmp_path, workbook_path, bindings_path, "Calc!B2"
    )
    assert "Data!B2" not in graph
    assert (
        '_INPUT_IDS: tuple[str, ...] = ("text_attributes", "lookup_key", "column_key")'
        in modules["model.py"]
    )
    pkg = load_package(modules, tmp_path, name="off_graph_input")
    assert "text_attributes" in pkg.validation.CHECKS
    got = pkg.compute_looked_up(
        pkg.LookedUpInputs(
            text_attributes=pkg.data.TEXT_ATTRIBUTES_DEFAULT,
            lookup_key="a",
            column_key="Y2020",
        )
    )
    assert got == 10
