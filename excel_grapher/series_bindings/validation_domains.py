"""Read Excel data validations as domain suggestions (#983).

Validations are evidence, not authored domains. `suggest_validation_domains`
reports one suggestion per validated cell without needing a sidecar;
`validation_cell_type_env` compiles `domain` suggestions through the same
path as authored series domains; `compare_series_validation_domains` checks
suggestions against an existing sidecar and never rewrites it.

In scope as domains: inline `list`, `list` over one area of constant cells
(directly or through a defined name), and `whole` / `decimal` with an
inclusive `between` over numeric literals. Everything else is reported
`unsupported` with a machine-readable `reason`.
"""

from __future__ import annotations

import re
from collections import defaultdict
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

import fastpyxl
from fastpyxl.utils.cell import get_column_letter, range_boundaries
from fastpyxl.workbook.workbook import Workbook
from fastpyxl.worksheet.datavalidation import DataValidation
from fastpyxl.worksheet.worksheet import Worksheet

from excel_grapher.core.cell_types import (
    CellType,
    constraints_to_cell_type_env,
    normalize_cell_type_env_key,
)
from excel_grapher.grapher.parser import DEFAULT_MAX_RANGE_CELLS
from excel_grapher.series_bindings.domains import (
    _input_domain_annotation,
    compile_domain_spec,
    declared_domain,
)
from excel_grapher.series_bindings.ranges import expand_bound_series_addresses

__all__ = [
    "DataValidationRule",
    "SeriesValidationFinding",
    "ValidationDomainSuggestion",
    "compare_series_validation_domains",
    "suggest_validation_domains",
    "validation_cell_type_env",
]

SuggestionStatus = Literal["domain", "unsupported"]
UnsupportedReason = Literal[
    "list_formula",
    "list_formula_cells",
    "list_multi_area",
    "empty_enum",
    "non_stop",
    "unsupported_type",
    "unsupported_operator",
    "non_literal_bounds",
    "unreadable",
]
FindingCode = Literal[
    "proposed_domain",
    "partial_coverage",
    "mixed_rules",
    "mixed_blank",
    "domain_validation_conflict",
    "matches_value_map",
    "value_map_side_mismatch",
    "value_map_validation_conflict",
]

_NUMBER = re.compile(r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?")
_INTEGER = re.compile(r"[-+]?\d+")
_NAME = re.compile(r"[A-Za-z_\\][\w.\\]*")
_CELL = r"\$?[A-Za-z]{1,3}\$?\d+"
_REF = re.compile(rf"(?:(?P<sheet>'(?:[^']|'')+'|[^'!:,()\s]+)!)?(?P<area>{_CELL}(?::{_CELL})?)")


@dataclass(frozen=True, slots=True)
class DataValidationRule:
    """A worksheet `dataValidation` as stored in the workbook."""

    type: str | None
    operator: str | None
    formula1: str | None
    formula2: str | None
    error_style: str | None


@dataclass(frozen=True, slots=True)
class ValidationDomainSuggestion:
    """One validated cell and the series domain its rule implies, if any.

    Attributes:
        address: `normalize_cell_type_env_key` of the cell (`Sheet!A1`).
        rule_id: `Sheet#index` of the rule; cells sharing one `sqref` share it.
        status: `domain` when `domain` is set, else `unsupported`.
        domain: `enum` / `between` / `real_between` mapping, or `None`.
        allow_blank: Excel `allowBlank`; recorded, not folded into `domain`.
        validation: The rule as stored.
        reason: Why the rule is `unsupported`, else `None`.
    """

    address: str
    rule_id: str
    status: SuggestionStatus
    domain: dict[str, Any] | None
    allow_blank: bool
    validation: DataValidationRule
    reason: UnsupportedReason | None = None


@dataclass(frozen=True, slots=True)
class SeriesValidationFinding:
    """How a series' bound cells compare with the workbook's validations.

    Attributes:
        series_id: Series the finding is about.
        code: Finding kind (see module docs for `compare_series_validation_domains`).
        message: Human-readable explanation.
        workbook_domain: Uniform domain read from the workbook, when there is one.
        declared_domain: Authored `domain` (or `value_map`-derived enum) compared against.
    """

    series_id: str
    code: FindingCode
    message: str
    workbook_domain: dict[str, Any] | None = None
    declared_domain: dict[str, Any] | None = None


class _Unsupported(Exception):
    def __init__(self, reason: UnsupportedReason) -> None:
        super().__init__(reason)
        self.reason: UnsupportedReason = reason


def suggest_validation_domains(
    workbook: Path | str,
    addresses: Iterable[str] | None = None,
    *,
    max_cells: int = DEFAULT_MAX_RANGE_CELLS,
) -> tuple[ValidationDomainSuggestion, ...]:
    """Report the domain each validated cell's data validation implies.

    Args:
        workbook: Workbook to read.
        addresses: Restrict the report to these cells (any quoting / `$`
            spelling). Cells without a validation are absent. `None` reports
            every validated cell on every sheet.
        max_cells: Upper bound on cells expanded from all `sqref`s when
            `addresses` is `None`.

    Returns:
        Suggestions ordered by sheet, row, column, then rule.

    Raises:
        ValueError: `addresses` is `None` and validations cover more than
            `max_cells` cells.
    """
    wb = fastpyxl.load_workbook(Path(workbook), data_only=False)
    try:
        return _suggest(wb, addresses, max_cells=max_cells)
    finally:
        wb.close()


def validation_cell_type_env(
    suggestions: Iterable[ValidationDomainSuggestion],
) -> dict[str, CellType]:
    """Compile `domain` suggestions into a caller-owned `CellTypeEnv`.

    Uses the same compiler as authored series domains; pass the result as
    `dynamic_refs.cell_type_env`. `unsupported` suggestions are skipped.

    Raises:
        ValueError: Two rules suggest different domains for one cell.
    """
    chosen: dict[str, dict[str, Any]] = {}
    for suggestion in suggestions:
        if suggestion.domain is None:
            continue
        previous = chosen.setdefault(suggestion.address, suggestion.domain)
        if _domain_key(previous) != _domain_key(suggestion.domain):
            raise ValueError(f"Conflicting validation domains at {suggestion.address}")
    schema = {
        address: _input_domain_annotation({"domain": domain}) for address, domain in chosen.items()
    }
    return constraints_to_cell_type_env(schema, {})


def compare_series_validation_domains(
    bindings: Mapping[str, Any],
    workbook: Path | str,
    *,
    max_cells: int = DEFAULT_MAX_RANGE_CELLS,
) -> tuple[SeriesValidationFinding, ...]:
    """Compare each series' bound cells with the workbook's validations.

    At most one finding per series; series with no validated domain cell,
    and `from_workbook` / `constant` series, produce none. The sidecar is
    never modified.

    * `partial_coverage` / `mixed_rules` / `mixed_blank`: validations do not
      cover the series uniformly; nothing is proposed.
    * `domain_validation_conflict`: the authored domain differs from the
      uniform workbook domain. Agreement is silent.
    * `matches_value_map` / `value_map_side_mismatch` /
      `value_map_validation_conflict`: an `input.value_map` series whose
      list equals the needles, equals the keys, or neither.
    * `proposed_domain`: no authored domain or map; the uniform workbook
      domain is offered for the author to copy.

    Args:
        bindings: Series bindings document.
        workbook: Workbook the sidecar describes.
        max_cells: Upper bound on cells per series `data_range`.
    """
    series_list = [s for s in bindings.get("series", []) if isinstance(s, dict)]
    bound: dict[str, list[str]] = {}
    for series in series_list:
        spec = compile_domain_spec(series)
        if spec is not None and spec.get("from_workbook") is True:
            continue
        cells = expand_bound_series_addresses(series, workbook=workbook, max_range_cells=max_cells)
        bound[str(series["id"])] = [normalize_cell_type_env_key(a) for a in cells]
    wanted = {address for cells in bound.values() for address in cells}
    wb = fastpyxl.load_workbook(Path(workbook), data_only=False)
    try:
        suggestions = _suggest(wb, wanted, max_cells=max_cells) if wanted else ()
    finally:
        wb.close()
    by_address: dict[str, list[ValidationDomainSuggestion]] = defaultdict(list)
    for suggestion in suggestions:
        by_address[suggestion.address].append(suggestion)

    findings: list[SeriesValidationFinding] = []
    for series in series_list:
        sid = str(series["id"])
        if sid not in bound:
            continue
        finding = _compare_series(series, sid, bound[sid], by_address)
        if finding is not None:
            findings.append(finding)
    return tuple(findings)


def _compare_series(
    series: Mapping[str, Any],
    sid: str,
    cells: Sequence[str],
    by_address: Mapping[str, list[ValidationDomainSuggestion]],
) -> SeriesValidationFinding | None:
    per_cell = [by_address.get(address, []) for address in cells]
    domain_cells = [s for found in per_cell for s in found if s.domain is not None]
    if not domain_cells:
        return None
    if any(len(found) != 1 or found[0].domain is None for found in per_cell):
        if any(len(found) > 1 for found in per_cell):
            return SeriesValidationFinding(
                sid, "mixed_rules", "More than one validation covers a bound cell."
            )
        return SeriesValidationFinding(
            sid, "partial_coverage", "Only some bound cells carry a supported validation."
        )
    domain = domain_cells[0].domain
    assert domain is not None
    if len({_domain_key(s.domain) for s in domain_cells if s.domain is not None}) > 1:
        return SeriesValidationFinding(
            sid, "mixed_rules", "Bound cells carry validations with different domains."
        )
    if len({s.allow_blank for s in domain_cells}) > 1:
        return SeriesValidationFinding(
            sid, "mixed_blank", "Bound cells disagree on allowBlank.", workbook_domain=domain
        )

    declared = declared_domain(series)
    if declared is not None:
        if _domain_key(declared) == _domain_key(domain):
            return None
        return SeriesValidationFinding(
            sid,
            "domain_validation_conflict",
            "Authored domain differs from the workbook validation; authored domain kept.",
            workbook_domain=domain,
            declared_domain=declared,
        )

    input_block = series.get("input")
    value_map = input_block.get("value_map") if isinstance(input_block, Mapping) else None
    if isinstance(value_map, Mapping) and value_map:
        needles = {"enum": list(value_map.values())}
        listed = _domain_key(domain)
        if listed == _domain_key(needles):
            code: FindingCode = "matches_value_map"
            message = "Validation list equals the value_map needles."
        elif listed == _domain_key({"enum": list(value_map.keys())}):
            code = "value_map_side_mismatch"
            message = "Validation list equals the value_map keys, not the needles; not imported."
        else:
            code = "value_map_validation_conflict"
            message = "Validation list matches neither side of value_map; map kept."
        return SeriesValidationFinding(
            sid, code, message, workbook_domain=domain, declared_domain=needles
        )

    return SeriesValidationFinding(
        sid,
        "proposed_domain",
        "Every bound cell carries the same validation domain.",
        workbook_domain=domain,
    )


# --- reading ---------------------------------------------------------------


def _suggest(
    wb: Workbook,
    addresses: Iterable[str] | None,
    *,
    max_cells: int,
) -> tuple[ValidationDomainSuggestion, ...]:
    wanted = None if addresses is None else {normalize_cell_type_env_key(a) for a in addresses}
    rows: list[tuple[tuple[int, int, int, int], ValidationDomainSuggestion]] = []
    expanded = 0
    for sheet_index, ws in enumerate(wb.worksheets):
        for rule_index, dv in enumerate(ws.data_validations.dataValidation):
            try:
                cells = _sqref_cells(dv)
            except ValueError:
                cells = []
            if wanted is None:
                expanded += len(cells)
                if expanded > max_cells:
                    raise ValueError(
                        f"Data validations cover more than {max_cells} cells; pass addresses"
                    )
            multi_cell = len(cells) != 1
            base = _rule(dv)
            for row, col in cells:
                address = normalize_cell_type_env_key(f"{ws.title}!{_coord(row, col)}")
                if wanted is not None and address not in wanted:
                    continue
                try:
                    domain = _domain_for(wb, ws, dv, multi_cell=multi_cell)
                    status: SuggestionStatus = "domain"
                    reason: UnsupportedReason | None = None
                except _Unsupported as exc:
                    domain, status, reason = None, "unsupported", exc.reason
                suggestion = ValidationDomainSuggestion(
                    address=address,
                    rule_id=f"{ws.title}#{rule_index}",
                    status=status,
                    domain=domain,
                    allow_blank=bool(dv.allowBlank),
                    validation=base,
                    reason=reason,
                )
                rows.append(((sheet_index, row, col, rule_index), suggestion))
    rows.sort(key=lambda item: item[0])
    return tuple(s for _, s in rows)


def _sqref_cells(dv: DataValidation) -> list[tuple[int, int]]:
    cells: list[tuple[int, int]] = []
    if dv.sqref is None:
        return cells
    for area in dv.sqref.ranges:
        for row in range(area.min_row, area.max_row + 1):
            for col in range(area.min_col, area.max_col + 1):
                cells.append((row, col))
    return cells


def _coord(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


def _rule(dv: DataValidation) -> DataValidationRule:
    return DataValidationRule(
        type=dv.type,
        operator=dv.operator,
        formula1=dv.formula1,
        formula2=dv.formula2,
        error_style=dv.errorStyle,
    )


def _domain_for(
    wb: Workbook, ws: Worksheet, dv: DataValidation, *, multi_cell: bool
) -> dict[str, Any]:
    kind = dv.type
    if kind not in {"list", "whole", "decimal"}:
        raise _Unsupported("unsupported_type")
    if dv.errorStyle not in (None, "stop"):
        raise _Unsupported("non_stop")
    if kind == "list":
        return {"enum": _list_members(wb, ws, dv.formula1, multi_cell=multi_cell)}
    if (dv.operator or "between") != "between":
        raise _Unsupported("unsupported_operator")
    if dv.formula1 is None or dv.formula2 is None:
        raise _Unsupported("unreadable")
    whole = kind == "whole"
    low = _bound(dv.formula1, whole=whole)
    high = _bound(dv.formula2, whole=whole)
    if low > high:
        raise _Unsupported("unreadable")
    return {"between" if whole else "real_between": {"min": low, "max": high}}


def _bound(formula: str, *, whole: bool) -> int | float:
    text = formula.strip().removeprefix("=")
    if whole:
        if not _INTEGER.fullmatch(text):
            raise _Unsupported("non_literal_bounds")
        return int(text)
    if not _NUMBER.fullmatch(text):
        raise _Unsupported("non_literal_bounds")
    return _number(text)


def _number(text: str) -> int | float:
    return int(text) if _INTEGER.fullmatch(text) else float(text)


def _list_members(
    wb: Workbook, ws: Worksheet, formula: str | None, *, multi_cell: bool
) -> list[object]:
    if formula is None:
        raise _Unsupported("unreadable")
    text = formula.strip().removeprefix("=").strip()
    if len(text) >= 2 and text.startswith('"') and text.endswith('"'):
        items: list[object] = [
            _number(item) if _NUMBER.fullmatch(item) else item
            for item in text[1:-1].replace('""', '"').split(",")
        ]
    else:
        sheet, area = _list_area(wb, ws, text, multi_cell=multi_cell)
        items = _area_values(wb, sheet, area)
    return _dedupe(items)


def _list_area(wb: Workbook, ws: Worksheet, text: str, *, multi_cell: bool) -> tuple[str, str]:
    match = _REF.fullmatch(text)
    if match is not None:
        area = match.group("area")
        if multi_cell and area.count("$") != 2 * (area.count(":") + 1):
            # Relative sources shift per cell across the sqref.
            raise _Unsupported("list_formula")
        sheet = match.group("sheet")
        if sheet is None:
            return ws.title, area
        if sheet.startswith("'"):
            sheet = sheet[1:-1].replace("''", "'")
        return sheet, area
    if not _NAME.fullmatch(text):
        raise _Unsupported("list_formula")
    defined = ws.defined_names.get(text) or wb.defined_names.get(text)
    if defined is None:
        raise _Unsupported("list_formula")
    try:
        destinations = list(defined.destinations)
    except Exception as exc:  # noqa: BLE001 - malformed defined names are evidence, not errors
        raise _Unsupported("list_formula") from exc
    if not destinations:
        raise _Unsupported("list_formula")
    if len(destinations) > 1:
        raise _Unsupported("list_multi_area")
    return destinations[0]


def _area_values(wb: Workbook, sheet: str, area: str) -> list[object]:
    if sheet not in wb.sheetnames:
        raise _Unsupported("unreadable")
    try:
        min_col, min_row, max_col, max_row = range_boundaries(area.replace("$", ""))
    except ValueError as exc:
        raise _Unsupported("unreadable") from exc
    assert min_col and min_row and max_col and max_row
    values: list[object] = []
    for row in wb[sheet].iter_rows(
        min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col
    ):
        for cell in row:
            value = cell.value
            if cell.data_type == "f" or (isinstance(value, str) and value.startswith("=")):
                raise _Unsupported("list_formula_cells")
            values.append(value)
    return values


def _dedupe(items: Iterable[object]) -> list[object]:
    seen: set[tuple[str, object]] = set()
    members: list[object] = []
    for item in items:
        if item is None or item == "":
            continue
        key = _member_key(item)
        if key not in seen:
            seen.add(key)
            members.append(item)
    if not members:
        raise _Unsupported("empty_enum")
    return members


# --- comparison ------------------------------------------------------------


def _member_key(value: object) -> tuple[str, object]:
    if isinstance(value, bool):
        return ("bool", value)
    if isinstance(value, (int, float)):
        return ("num", float(value))
    if isinstance(value, str):
        return ("str", value)
    return (type(value).__name__, repr(value))


def _domain_key(domain: Mapping[str, Any]) -> tuple[tuple[str, object], ...]:
    parts: list[tuple[str, object]] = []
    for kind, spec in domain.items():
        if kind == "enum":
            parts.append((kind, frozenset(_member_key(v) for v in spec)))
        elif kind in {"between", "real_between"} and isinstance(spec, Mapping):
            bounds = tuple(
                None if spec.get(side) is None else float(spec[side]) for side in ("min", "max")
            )
            parts.append((kind, bounds))
        else:
            parts.append((kind, repr(spec)))
    return tuple(sorted(parts, key=lambda part: part[0]))
