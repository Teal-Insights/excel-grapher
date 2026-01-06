from __future__ import annotations

import contextlib
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import NamedTuple
from xml.etree import ElementTree as ET

from .graph import DependencyGraph


class ValidationResult(NamedTuple):
    is_valid: bool
    in_graph_not_in_chain: set[str]
    in_chain_not_in_graph: set[str]
    messages: list[str]


def _parse_workbook_sheetid_map(excel_path: Path) -> dict[str, str]:
    """
    Return mapping {sheetId -> sheet_name} from xl/workbook.xml.
    """
    with zipfile.ZipFile(excel_path, "r") as zf:
        xml_bytes = zf.read("xl/workbook.xml")

    root = ET.fromstring(xml_bytes)
    out: dict[str, str] = {}
    for node in root.iter():
        if not node.tag.endswith("sheet"):
            continue
        sheet_id = node.attrib.get("sheetId")
        name = node.attrib.get("name")
        if sheet_id and name:
            out[str(sheet_id)] = str(name)
    return out


def _parse_calcchain_formula_cells(excel_path: Path, sheet_id_to_name: dict[str, str]) -> set[str] | None:
    """
    Return set of formula-cell keys like "SheetName!D35" as enumerated by calcChain.xml.

    Returns None if the workbook does not contain xl/calcChain.xml.
    """
    with zipfile.ZipFile(excel_path, "r") as zf:
        if "xl/calcChain.xml" not in set(zf.namelist()):
            return None
        xml_bytes = zf.read("xl/calcChain.xml")

    root = ET.fromstring(xml_bytes)
    keys: set[str] = set()

    for node in root.iter():
        if not node.tag.endswith("c"):
            continue
        cell_ref = node.attrib.get("r")
        sheet_id = node.attrib.get("i") or node.attrib.get("s")
        if not cell_ref or not sheet_id:
            continue
        sheet_name = sheet_id_to_name.get(str(sheet_id))
        if not sheet_name:
            continue
        keys.add(f"{sheet_name}!{cell_ref}")

    return keys


@dataclass(frozen=True)
class WorkbookCalcSettings:
    iterate_enabled: bool
    iterate_count: int
    iterate_delta: float


def get_calc_settings(workbook_path: Path) -> WorkbookCalcSettings:
    """
    Extract calculation settings from xl/workbook.xml.

    Defaults follow Excel's typical defaults when attributes are missing.
    """
    with zipfile.ZipFile(workbook_path, "r") as zf:
        xml_bytes = zf.read("xl/workbook.xml")

    root = ET.fromstring(xml_bytes)
    calc_pr = None
    for node in root.iter():
        if node.tag.endswith("calcPr"):
            calc_pr = node
            break

    # Common Excel defaults:
    # - iterate: disabled
    # - iterateCount: 100
    # - iterateDelta: 0.001
    iterate_enabled = False
    iterate_count = 100
    iterate_delta = 0.001

    if calc_pr is not None:
        it = (calc_pr.attrib.get("iterate") or "").strip()
        if it in {"1", "true", "TRUE"}:
            iterate_enabled = True
        elif it in {"0", "false", "FALSE"}:
            iterate_enabled = False

        ic = (calc_pr.attrib.get("iterateCount") or "").strip()
        if ic:
            with contextlib.suppress(ValueError):
                iterate_count = int(ic)

        idel = (calc_pr.attrib.get("iterateDelta") or "").strip()
        if idel:
            with contextlib.suppress(ValueError):
                iterate_delta = float(idel)

    return WorkbookCalcSettings(
        iterate_enabled=iterate_enabled,
        iterate_count=iterate_count,
        iterate_delta=iterate_delta,
    )


def validate_graph(
    graph: DependencyGraph,
    workbook_path: Path,
    *,
    scope: set[str] | None = None,
) -> ValidationResult:
    sheet_id_to_name = _parse_workbook_sheetid_map(workbook_path)
    calc = _parse_calcchain_formula_cells(workbook_path, sheet_id_to_name)

    graph_formulas: set[str] = set()
    for key in graph:
        node = graph.get_node(key)
        if node is None:
            continue
        if node.is_leaf:
            continue
        if scope is not None and node.sheet not in scope:
            continue
        graph_formulas.add(key)

    if calc is None:
        return ValidationResult(
            is_valid=True,
            in_graph_not_in_chain=set(),
            in_chain_not_in_graph=set(),
            messages=["calcChain.xml not found; skipping calcChain validation"],
        )

    if scope is not None:
        calc = {k for k in calc if k.split("!", 1)[0] in scope}

    in_graph_not_in_chain = graph_formulas - calc
    in_chain_not_in_graph = calc - graph_formulas

    messages: list[str] = []
    if in_graph_not_in_chain:
        messages.append(f"{len(in_graph_not_in_chain)} formula cells in graph but not in calcChain")
    if in_chain_not_in_graph:
        messages.append(f"{len(in_chain_not_in_graph)} cells in calcChain not reached by traversal")

    return ValidationResult(
        is_valid=(len(in_graph_not_in_chain) == 0),
        in_graph_not_in_chain=in_graph_not_in_chain,
        in_chain_not_in_graph=in_chain_not_in_graph,
        messages=messages,
    )

