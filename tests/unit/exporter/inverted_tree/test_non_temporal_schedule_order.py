"""Scheduled categorical matrices publish authored coordinate order."""

from pathlib import Path

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a20_matrix_join import _matrix_entry


@pytest.mark.parametrize("force_rung", [None, 2, 3])
def test_categorical_self_references_do_not_reverse_output_columns(
    tmp_path: Path, force_rung
) -> None:
    cells = {"B1": "Latest actual", "C1": "1-year ahead", "D1": "% change"}
    for row, commodity in enumerate(("Brent", "Wheat", "Coffee"), 2):
        cells[f"A{row}"] = commodity
        cells[f"B{row}"] = float(row * 10)
        cells[f"C{row}"] = float(row * 8)
        cells[f"A{row + 4}"] = commodity
        cells[f"B{row + 4}"] = f"=B{row}"
        cells[f"C{row + 4}"] = f"=C{row}"
        cells[f"D{row + 4}"] = f"=(C{row + 4}-B{row + 4})/B{row + 4}"
    entries = [
        _matrix_entry("source", "M!B2:C4", header_row=1),
        _matrix_entry("result", "M!B6:D8", header_row=1, direction="output"),
    ]
    for entry in entries:
        entry["key"] = ["REF_AREA", "INDICATOR"]
        dimension = entry["structure"]["dimensions"][1]
        dimension["id"] = dimension["concept"] = "INDICATOR"
        dimension["bind"]["read"] = "string"
    workbook = write_workbook(tmp_path / "categorical.xlsx", {"M": cells})
    pkg = load_package(
        generate_inverted(workbook, bindings_document(*entries), force_rung=force_rung),
        tmp_path,
        name="categorical",
    )
    actual = pkg.compute_result()
    for row, commodity in enumerate(("Brent", "Wheat", "Coffee"), 2):
        assert actual[commodity, "Latest actual"] == row * 10.0
        assert actual[commodity, "1-year ahead"] == row * 8.0
        assert actual[commodity, "% change"] == pytest.approx(-0.2)


@pytest.mark.parametrize("force_rung", [None, 2, 3])
def test_row_uses_workbook_geometry_in_categorical_schedule(tmp_path: Path, force_rung) -> None:
    from tests.unit.exporter.inverted_tree.helpers import series_entry

    cells = {}
    for row, label in enumerate(("Cameroon", "Cabo Verde", "Cambodia"), 2):
        cells[f"A{row}"] = cells[f"A{row + 4}"] = label
        cells[f"B{row}"] = row * 10.0
        cells[f"B{row + 4}"] = "=INDEX($B$2:$B$4,ROW()-5)"
    entries = [
        series_entry(
            "source",
            "M!B2:B4",
            layout="series",
            direction="constant",
            label_column="A",
            key_concept="REF_AREA",
            key_read="string",
        ),
        series_entry(
            "result",
            "M!B6:B8",
            layout="series",
            direction="output",
            label_column="A",
            key_concept="REF_AREA",
            key_read="string",
        ),
    ]
    workbook = write_workbook(tmp_path / "categorical_row.xlsx", {"M": cells})
    pkg = load_package(
        generate_inverted(workbook, bindings_document(*entries), force_rung=force_rung),
        tmp_path,
        name="categorical_row",
    )
    result = pkg.compute_result()
    for row, label in enumerate(("Cameroon", "Cabo Verde", "Cambodia"), 2):
        assert result[label] == row * 10.0

    from excel_grapher.exporter.inverted_tree.ast_emit import EmitContext, _host_coord_expr
    from excel_grapher.exporter.inverted_tree.catalog import schedule_axis_coord
    from excel_grapher.exporter.inverted_tree.schedule import FusedPlan
    from tests.unit.exporter.inverted_tree.helpers import inverted_graph_parts

    catalog, deps, graph = inverted_graph_parts(workbook, bindings_document(*entries))
    host = catalog.get("result")
    schedule = tuple(sorted(schedule_axis_coord(cell, catalog) for cell in host.cells))
    plan = FusedPlan(scc=("result",), schedule=schedule, domain={"result": (0, 3)}, regions=())
    ctx = EmitContext(
        host,
        catalog,
        deps["result"],
        0,
        host.cells[0],
        "t",
        None,
        fused_mode=True,
        fused_plan=plan,
        graph=graph,
    )
    expression = _host_coord_expr(ctx, axis="row")
    for cell in host.cells:
        t = plan.coord_to_t[schedule_axis_coord(cell, catalog)]
        assert eval(expression, {"t": t}) == int(cell.rsplit("B", 1)[1])


@pytest.mark.parametrize("area,row", [("a", 2), ("b", 3)])
def test_row_geometry_respects_active_group(tmp_path: Path, area, row) -> None:
    from excel_grapher.exporter.inverted_tree.ast_emit import EmitContext, _host_coord_expr
    from excel_grapher.exporter.inverted_tree.catalog import schedule_axis_coord
    from excel_grapher.exporter.inverted_tree.schedule import FusedPlan
    from tests.unit.exporter.inverted_tree.helpers import inverted_graph_parts

    workbook = write_workbook(
        tmp_path / "group_rows.xlsx",
        {
            "M": {
                "B1": 2025,
                "C1": 2026,
                "A2": "a",
                "A3": "b",
                "B2": "=ROW()",
                "C2": "=ROW()",
                "B3": "=ROW()",
                "C3": "=ROW()",
            }
        },
    )
    document = bindings_document(
        _matrix_entry("result", "M!B2:C3", header_row=1, direction="output")
    )
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    host = catalog.get("result")
    schedule = tuple(dict.fromkeys(schedule_axis_coord(cell, catalog) for cell in host.cells))
    plan = FusedPlan(scc=("result",), schedule=schedule, domain={"result": (0, 2)}, regions=())
    index = (row - 2) * 2
    ctx = EmitContext(
        host,
        catalog,
        deps["result"],
        index,
        host.cells[index],
        "t",
        None,
        fused_mode=True,
        fused_plan=plan,
        fused_partition=(area,),
        graph=graph,
    )
    expression = _host_coord_expr(ctx, axis="row")
    assert [eval(expression, {"t": t}) for t in range(2)] == [row, row]
