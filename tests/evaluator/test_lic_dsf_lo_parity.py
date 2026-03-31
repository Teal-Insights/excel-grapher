"""
LIC-DSF: FormulaEvaluator vs LibreOffice recalculated values (slow).

Compares FormulaEvaluator results against an independently recalculated
workbook produced by LibreOffice headless.  This is stronger than comparing
against Excel's embedded cache because it uses a second, independent
calculation engine.

Known divergence: LO 25.8+ correctly evaluates ``_xlfn.XLOOKUP`` on the
``BLEND floating calculations WB`` sheet, producing slightly different
interest-rate parameters than Excel's cached values.  This cascades through
PV-of-debt calculations and causes ~0.2–0.6 pp drift on most Chart Data
cells.  FormulaEvaluator matches **Excel's cache** (tested separately in
``test_lic_dsf_chart_parity.py``), so the divergence reflects an Excel-vs-LO
difference, not a FormulaEvaluator bug.

The LO-recalculated xlsx is expected at::

    example/data/tmp/recalc/lic-dsf-template-2025-08-12.xlsx

Generate it with (requires LibreOffice 25.8+)::

    soffice --headless --norestore --convert-to xlsx \\
        --outdir example/data/tmp/recalc \\
        example/data/lic-dsf-template-2025-08-12.xlsm
"""

from __future__ import annotations

import contextlib
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from excel_grapher import FormulaEvaluator, create_dependency_graph
from tests.evaluator.lic_dsf_chart_targets import (
    GRAPH_MAX_DEPTH,
    WORKBOOK_PATH,
    collect_chart_data_cell_keys,
)

pytestmark = pytest.mark.slow

LO_RECALC_PATH = Path("example/data/tmp/recalc/lic-dsf-template-2025-08-12.xlsx")

RTOL = 1e-5
ATOL = 1e-9

# Figure 1 rows: baseline + stress tests (excluding customized-scenario rows
# 51/93/135/177 which involve PV calculations that diverge between Excel and LO).
_FIGURE1_ROWS = {
    61,
    62,
    63,
    64,
    66,
    103,
    104,
    105,
    106,
    108,
    145,
    146,
    147,
    148,
    150,
    187,
    188,
    189,
    190,
    192,
}

# ── helpers ──────────────────────────────────────────────────────────────────

_SS_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")


def _read_lo_cached_values(
    xlsx_path: Path,
    sheet_name: str,
    cell_refs: set[str],
) -> dict[str, float | str]:
    """Stream-read cached cell values from an LO-recalculated xlsx.

    Returns a dict of ``{CellRef: value}`` where value is ``float`` for
    numeric cells or a string like ``"#N/A"`` for error cells.
    Only cells in *cell_refs* (e.g. ``{"D61", "E61", …}``) are returned.

    Uses streaming XML parsing to avoid loading the entire 40+ MB sheet
    into memory.
    """
    results: dict[str, float | str] = {}

    with zipfile.ZipFile(xlsx_path, "r") as z:
        wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
        target_rid = None
        for sh in wb_xml.iter(f"{{{_SS_NS}}}sheet"):
            if sh.get("name") == sheet_name:
                target_rid = sh.get(f"{{{_REL_NS}}}id")
                break
        if target_rid is None:
            return results

        rels_xml = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        sheet_path = None
        for rel in rels_xml.iter(f"{{{_PKG_NS}}}Relationship"):
            if rel.get("Id") == target_rid:
                sheet_path = "xl/" + (rel.get("Target") or "")
                break
        if sheet_path is None:
            return results

        with z.open(sheet_path) as f:
            for _event, elem in ET.iterparse(f, events=["end"]):
                if elem.tag == f"{{{_SS_NS}}}c":
                    r = elem.get("r", "")
                    if r in cell_refs:
                        t = elem.get("t", "")
                        v_el = elem.find(f"{{{_SS_NS}}}v")
                        val = v_el.text if v_el is not None else None
                        if t == "e":
                            if val:
                                results[r] = val  # e.g. "#N/A"
                        elif val is not None:
                            with contextlib.suppress(ValueError, TypeError):
                                results[r] = float(val)
                # Free memory for processed rows
                if elem.tag == f"{{{_SS_NS}}}row":
                    elem.clear()

    return results


def _numeric_close(a: float, b: float, *, rtol: float, atol: float) -> bool:
    if a == b:
        return True
    scale = max(abs(a), abs(b), 1.0)
    return abs(a - b) <= max(atol, rtol * scale)


def _compare_evaluator_to_lo(
    graph,
    candidates: list[str],
    lo_cache: dict[str, float | str],
    key_to_ref: dict[str, str],
    *,
    rtol: float,
    atol: float,
) -> list[tuple[str, float, object]]:
    """Return list of (key, lo_value, eval_value) mismatches."""
    mismatches: list[tuple[str, float, object]] = []
    with FormulaEvaluator(graph) as ev:
        for key in candidates:
            ref = key_to_ref[key]
            lo_val = lo_cache[ref]
            assert isinstance(lo_val, float)

            try:
                eval_val = ev._evaluate_cell(key)  # noqa: SLF001
            except Exception as exc:
                mismatches.append((key, lo_val, f"EXCEPTION: {exc!r}"))
                continue

            if not isinstance(eval_val, (int, float)) or isinstance(eval_val, bool):
                mismatches.append((key, lo_val, eval_val))
                continue

            if not _numeric_close(lo_val, float(eval_val), rtol=rtol, atol=atol):
                mismatches.append((key, lo_val, eval_val))

    return mismatches


def _format_mismatches(
    mismatches: list[tuple[str, float, object]],
    graph,
    total_candidates: int,
    limit: int = 50,
) -> str:
    lines = [
        f"FormulaEvaluator vs LibreOffice: {len(mismatches)} mismatches "
        f"(out of {total_candidates} candidates):"
    ]
    for key, lo_val, eval_val in mismatches[:limit]:
        node = graph.get_node(key)
        formula = node.normalized_formula if node else None
        lines.append(
            f"  {key}: LO={lo_val!r}  eval={eval_val!r}"
            + (f"  formula={formula[:100]}" if formula else "")
        )
    if len(mismatches) > limit:
        lines.append(f"  ... and {len(mismatches) - limit} more")
    return "\n".join(lines)


# ── fixtures ─────────────────────────────────────────────────────────────────


@pytest.fixture(scope="module")
def lic_dsf_full_chart_graph():
    if not WORKBOOK_PATH.exists():
        pytest.skip(f"Test workbook not found at {WORKBOOK_PATH}")
    keys = collect_chart_data_cell_keys()
    return create_dependency_graph(
        WORKBOOK_PATH,
        keys,
        load_values=True,
        max_depth=GRAPH_MAX_DEPTH,
        use_cached_dynamic_refs=True,
    )


@pytest.fixture(scope="module")
def lo_cache_and_mapping(lic_dsf_full_chart_graph):
    """Load LO cached values and build key→ref mapping for Chart Data."""
    if not LO_RECALC_PATH.exists():
        pytest.skip(f"LO-recalculated workbook not found at {LO_RECALC_PATH}")

    from excel_grapher.evaluator.name_utils import parse_address

    export_keys = collect_chart_data_cell_keys()

    key_to_ref: dict[str, str] = {}
    bare_refs: set[str] = set()
    for key in export_keys:
        sheet, coord = parse_address(key)
        if sheet == "Chart Data":
            bare_refs.add(coord)
            key_to_ref[key] = coord

    lo_cache = _read_lo_cached_values(LO_RECALC_PATH, "Chart Data", bare_refs)
    return lo_cache, key_to_ref


# ── tests ────────────────────────────────────────────────────────────────────


def test_lo_recalc_exists() -> None:
    """Guard: LO-recalculated xlsx must be present to run parity tests."""
    if not LO_RECALC_PATH.exists():
        pytest.skip(
            f"LO-recalculated workbook not found at {LO_RECALC_PATH}. "
            "Generate with: soffice --headless --norestore --convert-to xlsx "
            f"--outdir {LO_RECALC_PATH.parent} {WORKBOOK_PATH}"
        )


def test_evaluator_matches_lo_for_figure1_external_debt(
    lic_dsf_full_chart_graph,
    lo_cache_and_mapping,
) -> None:
    """FormulaEvaluator must match LO for Figure 1 (external debt) Chart Data rows.

    These rows (61-66, 103-108, 145-150, 187-192) cover the four panels of
    Figure 1: PV debt-to-GDP, PV debt-to-exports, debt-service-to-exports,
    and debt-service-to-revenue.
    """
    graph = lic_dsf_full_chart_graph
    lo_cache, key_to_ref = lo_cache_and_mapping
    export_keys = collect_chart_data_cell_keys()

    candidates: list[str] = []
    for key in export_keys:
        ref = key_to_ref.get(key)
        if ref is None:
            continue
        m = _CELL_RE.match(ref)
        if m is None or int(m.group(2)) not in _FIGURE1_ROWS:
            continue
        lo_val = lo_cache.get(ref)
        if not isinstance(lo_val, float):
            continue
        node = graph.get_node(key)
        if node is None or not node.normalized_formula:
            continue
        candidates.append(key)

    assert len(candidates) > 0, "No Figure 1 candidates found"

    mismatches = _compare_evaluator_to_lo(
        graph, candidates, lo_cache, key_to_ref, rtol=RTOL, atol=ATOL
    )

    if mismatches:
        raise AssertionError(_format_mismatches(mismatches, graph, len(candidates)))


def test_evaluator_matches_lo_for_all_chart_data(
    lic_dsf_full_chart_graph,
    lo_cache_and_mapping,
) -> None:
    """FormulaEvaluator must match LO for all Chart Data export cells.

    This covers the full set: external debt figures, public debt stress
    blocks, signal rows, and fiscal space rows.
    """
    graph = lic_dsf_full_chart_graph
    lo_cache, key_to_ref = lo_cache_and_mapping
    export_keys = collect_chart_data_cell_keys()

    candidates: list[str] = []
    for key in export_keys:
        ref = key_to_ref.get(key)
        if ref is None:
            continue
        lo_val = lo_cache.get(ref)
        if not isinstance(lo_val, float):
            continue
        node = graph.get_node(key)
        if node is None or not node.normalized_formula:
            continue
        candidates.append(key)

    assert len(candidates) > 0, "No candidate cells found"

    mismatches = _compare_evaluator_to_lo(
        graph, candidates, lo_cache, key_to_ref, rtol=RTOL, atol=ATOL
    )

    if mismatches:
        raise AssertionError(_format_mismatches(mismatches, graph, len(candidates)))
