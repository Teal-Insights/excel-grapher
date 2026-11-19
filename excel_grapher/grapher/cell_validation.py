"""Excel data-validation rules for constrained inputs on write-back.

Input domains (`enum`, `between`, `real_between`, `from_workbook`,
an enum unioned with one interval, `value_map` needles, and series
relations) become worksheet data validations. When `series_bindings` is
set, input-series domains are
applied and constant/output/internal cells are skipped; extract-time
`cell_type_env` domains still apply to unbound value leaves, and matching
cells must agree. Without bindings, value cells in `cell_type_env` are
the sole constraint source.
"""

from __future__ import annotations

import math
from collections import defaultdict
from collections.abc import Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass
from pathlib import Path

from fastpyxl.utils.cell import column_index_from_string, coordinate_from_string
from fastpyxl.worksheet.cell_range import CellRange, MultiCellRange
from fastpyxl.worksheet.datavalidation import DataValidation
from fastpyxl.worksheet.worksheet import Worksheet

from excel_grapher.core.address_keys import parse_address, quote_sheet_if_needed
from excel_grapher.core.cell_types import (
    CellType,
    GreaterThanCell,
    NotEqualCell,
    is_union_domain,
    normalize_cell_type_env_key,
)
from excel_grapher.grapher.graph import DependencyGraph, GraphReadView
from excel_grapher.series_bindings.domains import SeriesDomainIndex, compile_domain_spec
from excel_grapher.series_bindings.normalize import has_input_direction
from excel_grapher.series_bindings.ranges import expand_bound_series_addresses
from excel_grapher.series_bindings.relations import iter_series_relations
from excel_grapher.series_bindings.types import WorkbookSeriesBindings
from excel_grapher.series_bindings.validation_domains import ValidationDomainSuggestion

_FORMULA_LIMIT = 255
_SQREF_LIMIT = 255
_EXCEL_MAX_COL = 16_384
_EXCEL_MAX_ROW = 1_048_576
_LIST_ERROR = "Value must be one of the allowed inputs."
_WHOLE_ERROR = "Value must be an integer in the declared range."
_DECIMAL_ERROR = "Value must be a number in the declared range."
_CUSTOM_ERROR = "Value is outside the declared input constraint."


@dataclass(frozen=True, slots=True)
class _Rule:
    type: str
    operator: str | None
    formula1: str
    formula2: str | None
    allow_blank: bool
    error: str


def apply_constrained_input_validations(
    sheets: Mapping[str, Worksheet],
    graph: GraphReadView,
    *,
    planned: Sequence[tuple[str, str, object]],
    series_bindings: WorkbookSeriesBindings | None,
    bindings_workbook: Path | str | None,
) -> None:
    """Attach Excel data validations for constrained inputs on `sheets`.

    Args:
        sheets: Worksheets already created for this write, keyed by name.
        graph: Read view being written.
        planned: Cells this write will persist, as `(sheet, coord, value)`.
        series_bindings: Sidecar whose input domains contribute constraints
            when set. Constants and non-input series are ignored; unbound
            value leaves still take `cell_type_env` domains.
        bindings_workbook: Workbook the sidecar describes. Required when an
            input series declares relations.

    Raises:
        ValueError: A constraint cannot be expressed as one Excel data
            validation, a relation partner is absent from `planned`, a
            `from_workbook` input has no cached value, or relations need
            `bindings_workbook`.
    """
    written = {
        normalize_cell_type_env_key(f"{sheet}!{coord}"): (sheet, coord)
        for sheet, coord, _value in planned
    }
    constraints = _constraint_cells(
        graph,
        written=set(written),
        series_bindings=series_bindings,
        bindings_workbook=bindings_workbook,
    )
    _emit_constraint_rules(sheets, constraints, written)


def apply_workbook_cell_validations(
    sheets: Mapping[str, Worksheet],
    graph: GraphReadView,
    *,
    planned: Sequence[tuple[str, str, object]],
    series_bindings: WorkbookSeriesBindings | None,
    bindings_workbook: Path | str | None,
    replace_extracted: bool,
) -> None:
    """Write stored worksheet rules, then authored input constraints.

    Authored series domains and `cell_type_env` constraints win on a cell
    that also has a stored worksheet rule. When `replace_extracted` is
    True, existing worksheet rules are removed from materialized cells and
    from addresses that still have a stored validation before the stored
    set is written, so a rule deleted from the graph is not copied back.

    Raises:
        ValueError: A constraint cannot be expressed as one Excel data
            validation, a relation partner is absent from `planned`, a
            `from_workbook` input has no cached value, or relations need
            `bindings_workbook`.
    """
    written = {
        normalize_cell_type_env_key(f"{sheet}!{coord}"): (sheet, coord)
        for sheet, coord, _value in planned
    }
    constraints = _constraint_cells(
        graph,
        written=set(written),
        series_bindings=series_bindings,
        bindings_workbook=bindings_workbook,
    )
    if replace_extracted:
        strip_validation_coords(sheets, _managed_validation_coords(graph))
    _apply_stored_data_validations(
        sheets,
        graph.data_validations,
        skip=set(constraints),
        limit=None if replace_extracted else set(written),
    )
    _emit_constraint_rules(sheets, constraints, written)


def strip_validation_coords(
    sheets: Mapping[str, Worksheet],
    coords_by_sheet: Mapping[str, AbstractSet[str]],
) -> None:
    """Remove `coords_by_sheet` from existing worksheet data-validation ranges.

    Splits rectangular ranges around removed coordinates so untouched cells
    keep their rules, and drops a rule whose range becomes empty.

    Args:
        sheets: Worksheets keyed by title.
        coords_by_sheet: A1 coordinates (such as `B2`) to strip per sheet.
    """
    for sheet_name, raw_coords in coords_by_sheet.items():
        if not raw_coords or sheet_name not in sheets:
            continue
        worksheet = sheets[sheet_name]
        existing = list(worksheet.data_validations.dataValidation)
        if not existing:
            continue
        points = {_coord_key(coord) for coord in raw_coords}
        kept: list[DataValidation] = []
        for dv in existing:
            sqref = dv.sqref
            if sqref is None:
                continue
            next_ranges: list[CellRange] = []
            changed = False
            for rng in sqref.ranges:
                c1 = int(rng.min_col) if rng.min_col is not None else 1
                r1 = int(rng.min_row) if rng.min_row is not None else 1
                c2 = int(rng.max_col) if rng.max_col is not None else _EXCEL_MAX_COL
                r2 = int(rng.max_row) if rng.max_row is not None else _EXCEL_MAX_ROW
                hits = [(col, row) for col, row in points if c1 <= col <= c2 and r1 <= row <= r2]
                if not hits:
                    next_ranges.append(rng)
                    continue
                changed = True
                for sub_c1, sub_r1, sub_c2, sub_r2 in _subtract_points_from_rect(
                    (c1, r1, c2, r2), hits
                ):
                    next_ranges.append(
                        CellRange(min_col=sub_c1, min_row=sub_r1, max_col=sub_c2, max_row=sub_r2)
                    )
            if not next_ranges:
                continue
            if changed:
                next_ranges.sort(key=lambda item: (item.min_row or 1, item.min_col or 1))
                dv.sqref = MultiCellRange(next_ranges)
            kept.append(dv)
        worksheet.data_validations.dataValidation = kept


def _subtract_points_from_rect(
    rect: tuple[int, int, int, int],
    points: Sequence[tuple[int, int]],
) -> list[tuple[int, int, int, int]]:
    rects = [rect]
    for pcol, prow in points:
        updated: list[tuple[int, int, int, int]] = []
        for c1, r1, c2, r2 in rects:
            if not (c1 <= pcol <= c2 and r1 <= prow <= r2):
                updated.append((c1, r1, c2, r2))
                continue
            if r1 <= prow - 1:
                updated.append((c1, r1, c2, prow - 1))
            if prow + 1 <= r2:
                updated.append((c1, prow + 1, c2, r2))
            if c1 <= pcol - 1:
                updated.append((c1, prow, pcol - 1, prow))
            if pcol + 1 <= c2:
                updated.append((pcol + 1, prow, c2, prow))
        rects = updated
        if not rects:
            break
    return rects


def _managed_validation_coords(graph: GraphReadView) -> dict[str, set[str]]:
    """Cells whose worksheet rules come from the graph, not the file."""
    by_sheet: dict[str, set[str]] = defaultdict(set)
    for key in graph:
        sheet, coord = parse_address(str(key))
        by_sheet[sheet].add(coord.replace("$", ""))
    for item in graph.data_validations:
        sheet, coord = parse_address(item.address)
        by_sheet[sheet].add(coord.replace("$", ""))
    return by_sheet


def _apply_stored_data_validations(
    sheets: Mapping[str, Worksheet],
    validations: Sequence[ValidationDomainSuggestion],
    *,
    skip: set[str],
    limit: set[str] | None,
) -> None:
    grouped: dict[tuple[str, tuple[object, ...]], list[str]] = defaultdict(list)
    for item in validations:
        if item.address in skip:
            continue
        if limit is not None and item.address not in limit:
            continue
        sheet, coord = parse_address(item.address)
        if sheet not in sheets:
            continue
        grouped[(sheet, _stored_payload(item))].append(coord.replace("$", ""))
    for (sheet_name, payload), coords in grouped.items():
        worksheet = sheets[sheet_name]
        for sqref in _chunk_sqref(coords):
            worksheet.add_data_validation(_validation_from_payload(payload, sqref))


def _stored_payload(item: ValidationDomainSuggestion) -> tuple[object, ...]:
    rule = item.validation
    return (
        rule.type,
        rule.operator,
        rule.formula1,
        rule.formula2,
        item.allow_blank,
        rule.error_style,
        rule.show_error_message,
        rule.show_input_message,
        rule.show_drop_down,
        rule.error,
        rule.error_title,
        rule.prompt,
        rule.prompt_title,
        rule.ime_mode,
    )


def _validation_from_payload(payload: tuple[object, ...], sqref: str) -> DataValidation:
    (
        kind,
        operator,
        formula1,
        formula2,
        allow_blank,
        error_style,
        show_error,
        show_input,
        show_drop_down,
        error,
        error_title,
        prompt,
        prompt_title,
        ime_mode,
    ) = payload
    return DataValidation(
        type=kind,
        operator=operator,
        formula1=formula1,
        formula2=formula2,
        allow_blank=allow_blank,
        errorStyle=error_style,
        showErrorMessage=show_error,
        showInputMessage=show_input,
        showDropDown=show_drop_down,
        error=error,
        errorTitle=error_title,
        prompt=prompt,
        promptTitle=prompt_title,
        imeMode=ime_mode,
        sqref=sqref,
    )


def _emit_constraint_rules(
    sheets: Mapping[str, Worksheet],
    constraints: Mapping[str, CellType],
    written: Mapping[str, tuple[str, str]],
) -> None:
    grouped: dict[tuple[str, _Rule], list[str]] = defaultdict(list)
    written_keys = set(written)
    for address, cell_type in constraints.items():
        rule = _rule_for(address, cell_type, written=written_keys)
        sheet, coord = written[address]
        grouped[(sheet, rule)].append(coord)

    for (sheet_name, rule), coords in grouped.items():
        worksheet = sheets[sheet_name]
        for sqref in _chunk_sqref(coords):
            # OOXML `showDropDown` true hides the in-cell dropdown.
            worksheet.add_data_validation(
                DataValidation(
                    type=rule.type,
                    operator=rule.operator,
                    formula1=rule.formula1,
                    formula2=rule.formula2,
                    allow_blank=rule.allow_blank,
                    showErrorMessage=True,
                    showInputMessage=False,
                    showDropDown=False,
                    errorStyle="stop",
                    error=rule.error,
                    sqref=sqref,
                )
            )


def _constraint_cells(
    graph: GraphReadView,
    *,
    written: set[str],
    series_bindings: WorkbookSeriesBindings | None,
    bindings_workbook: Path | str | None,
) -> dict[str, CellType]:
    from_env = _constraints_from_env(graph, written=written)
    if series_bindings is None:
        return from_env
    from_bindings = _constraints_from_bindings(
        graph,
        written=written,
        series_bindings=series_bindings,
        bindings_workbook=bindings_workbook,
    )
    # Bindings own series direction: do not decorate constant/output/internal
    # cells from extract-time env. Unbound value leaves still take env domains
    # (label-only sidecars must not drop CONSTRAINTS). Conflicts fail closed.
    blocked = _non_input_bound_addresses(
        graph,
        series_bindings=series_bindings,
        bindings_workbook=bindings_workbook,
        written=written,
    )
    if blocked:
        from_env = {address: cell for address, cell in from_env.items() if address not in blocked}
    return _merge_constraint_maps(from_env, from_bindings)


def _merge_constraint_maps(
    base: Mapping[str, CellType],
    overlay: Mapping[str, CellType],
) -> dict[str, CellType]:
    """Union two constraint maps; identical cells may overlap, conflicts fail."""
    merged = dict(base)
    for address, cell_type in overlay.items():
        previous = merged.get(address)
        if previous is not None and previous != cell_type:
            raise _fail(address, "conflicting input constraints")
        merged[address] = cell_type
    return merged


def _non_input_bound_addresses(
    graph: GraphReadView,
    *,
    series_bindings: WorkbookSeriesBindings,
    bindings_workbook: Path | str | None,
    written: set[str],
) -> set[str]:
    blocked: set[str] = set()
    for series in series_bindings["series"]:
        if not isinstance(series, dict) or has_input_direction(series):
            continue
        for address in expand_bound_series_addresses(
            series,
            workbook=bindings_workbook,
            named_ranges=graph.named_ranges,
            named_range_ranges=graph.named_range_ranges,
        ):
            norm = normalize_cell_type_env_key(address)
            if norm in written:
                blocked.add(norm)
    return blocked


def _constraints_from_bindings(
    graph: GraphReadView,
    *,
    written: set[str],
    series_bindings: WorkbookSeriesBindings,
    bindings_workbook: Path | str | None,
) -> dict[str, CellType]:
    if _bindings_declare_relations(series_bindings) and bindings_workbook is None:
        raise ValueError(
            "bindings_workbook is required to write cell validation for series relations"
        )
    index = SeriesDomainIndex.from_bindings(
        series_bindings,
        workbook=bindings_workbook,
        graph=_domain_graph(graph),
    )
    found: dict[str, CellType] = {}
    for series in series_bindings["series"]:
        if not isinstance(series, dict) or not has_input_direction(series):
            continue
        spec = compile_domain_spec(series)
        relations = iter_series_relations(series)
        if spec is None and not relations:
            continue
        addresses = expand_bound_series_addresses(
            series,
            workbook=bindings_workbook,
            named_ranges=graph.named_ranges,
            named_range_ranges=graph.named_range_ranges,
        )
        for address in addresses:
            norm = normalize_cell_type_env_key(address)
            if norm not in written:
                continue
            cell_type = index.domain_for(norm)
            if cell_type is None:
                if spec is not None and spec.get("from_workbook") is True:
                    raise _fail(address, "from_workbook cell has no cached value")
                continue
            if not _is_constraining(cell_type):
                continue
            previous = found.get(norm)
            if previous is not None and previous != cell_type:
                raise _fail(address, "conflicting input constraints")
            found[norm] = cell_type
    return found


def _constraints_from_env(graph: GraphReadView, *, written: set[str]) -> dict[str, CellType]:
    env = _cell_type_env(graph)
    if env is None:
        return {}
    found: dict[str, CellType] = {}
    for key in graph:
        node = graph.get_node(key)
        if node is None or node.has_formula:
            continue
        try:
            norm = normalize_cell_type_env_key(str(key))
        except (IndexError, ValueError):
            continue
        if norm not in written:
            continue
        cell_type = _lookup_cell_type(env, norm)
        if cell_type is None or not _is_constraining(cell_type):
            continue
        found[norm] = cell_type
    return found


def _bindings_declare_relations(series_bindings: WorkbookSeriesBindings) -> bool:
    for series in series_bindings["series"]:
        if (
            isinstance(series, dict)
            and has_input_direction(series)
            and iter_series_relations(series)
        ):
            return True
    return False


def _domain_graph(graph: GraphReadView) -> DependencyGraph | None:
    if isinstance(graph, DependencyGraph):
        return graph
    projected = getattr(graph, "projected_graph", None)
    if isinstance(projected, DependencyGraph):
        return projected
    return None


def _cell_type_env(graph: GraphReadView) -> Mapping[str, CellType] | None:
    env = getattr(graph, "cell_type_env", None)
    if env is None:
        projected = getattr(graph, "projected_graph", None)
        if projected is not None:
            env = getattr(projected, "cell_type_env", None)
    if env is None:
        return None
    if not isinstance(env, Mapping):
        raise ValueError("cell_type_env must map addresses to CellType values")
    return env


def _lookup_cell_type(env: Mapping[str, CellType], address: str) -> CellType | None:
    found = env.get(address)
    if found is None:
        return None
    if not isinstance(found, CellType):
        raise _fail(address, f"expected a CellType, got {type(found).__name__}")
    return found


def _is_constraining(cell_type: CellType) -> bool:
    if cell_type.enum is not None or cell_type.relations:
        return True
    interval = cell_type.interval
    if interval is not None and (interval.min is not None or interval.max is not None):
        return True
    real = cell_type.real_interval
    return real is not None and (real.min is not None or real.max is not None)


def _rule_for(address: str, cell_type: CellType, *, written: set[str]) -> _Rule:
    sheet, coord = parse_address(address)
    interval = cell_type.interval
    real = cell_type.real_interval
    has_interval = interval is not None and (interval.min is not None or interval.max is not None)
    has_real = real is not None and (real.min is not None or real.max is not None)
    if has_interval and has_real:
        raise _fail(address, "integer and real intervals cannot both be written")
    relation_clauses = _relation_clauses(
        address, coord, cell_type, host_sheet=sheet, written=written
    )
    if is_union_domain(cell_type):
        return _union_rule(
            address,
            coord,
            cell_type,
            relation_clauses=relation_clauses,
        )
    if cell_type.enum is not None and (has_interval or has_real):
        raise _fail(address, "enum and interval domains cannot both be written")
    if cell_type.enum is not None and not cell_type.enum.values:
        raise _fail(address, "enum domain is empty")
    if cell_type.enum is not None:
        present, allow_blank = _split_blanks(cell_type.enum.values)
        if not present:
            rule = _Rule("custom", None, f'{coord}=""', None, True, _CUSTOM_ERROR)
            return _checked(address, rule)
        terms = [_enum_term(address, coord, value) for value in present]
        membership = _or(terms)
        if relation_clauses:
            rule = _Rule(
                "custom",
                None,
                _and([membership, *relation_clauses]),
                None,
                allow_blank,
                _CUSTOM_ERROR,
            )
            return _checked(address, rule)
        if _value_class(address, present) == "bool" or _needs_custom_enum(present):
            rule = _Rule("custom", None, membership, None, allow_blank, _CUSTOM_ERROR)
            return _checked(address, rule)
        rule = _Rule("list", None, _inline_list(address, present), None, allow_blank, _LIST_ERROR)
        return _checked(address, rule)

    if relation_clauses:
        bound_clauses: list[str] = []
        if has_interval and interval is not None:
            bound_clauses = _bound_clauses(address, coord, interval.min, interval.max, whole=True)
        elif has_real and real is not None:
            bound_clauses = _bound_clauses(address, coord, real.min, real.max, whole=False)
        rule = _Rule(
            "custom",
            None,
            _and([*bound_clauses, *relation_clauses]),
            None,
            False,
            _CUSTOM_ERROR,
        )
        return _checked(address, rule)

    if has_interval and interval is not None:
        return _checked(
            address, _native_interval(address, interval.min, interval.max, decimal=False)
        )
    if has_real and real is not None:
        return _checked(address, _native_interval(address, real.min, real.max, decimal=True))
    raise _fail(address, "constraint has no Excel data-validation form")


def _union_rule(
    address: str,
    coord: str,
    cell_type: CellType,
    *,
    relation_clauses: list[str],
) -> _Rule:
    """Write a custom formula that accepts the enum or the numeric interval."""
    assert cell_type.enum is not None
    present, allow_blank = _split_blanks(cell_type.enum.values)
    terms = [_enum_term(address, coord, value) for value in present]
    if cell_type.interval is not None:
        bounds = _bound_clauses(
            address,
            coord,
            cell_type.interval.min,
            cell_type.interval.max,
            whole=True,
        )
    elif cell_type.real_interval is not None:
        bounds = _bound_clauses(
            address,
            coord,
            cell_type.real_interval.min,
            cell_type.real_interval.max,
            whole=False,
        )
    else:
        bounds = []
    terms.append(_and([f"ISNUMBER({coord})", *bounds]) if bounds else f"ISNUMBER({coord})")
    membership = _or(terms)
    if relation_clauses:
        membership = _and([membership, *relation_clauses])
    return _checked(
        address,
        _Rule("custom", None, membership, None, allow_blank, _CUSTOM_ERROR),
    )


def _relation_clauses(
    address: str,
    coord: str,
    cell_type: CellType,
    *,
    host_sheet: str,
    written: set[str],
) -> list[str]:
    clauses: list[str] = []
    for relation in cell_type.relations:
        if isinstance(relation, GreaterThanCell):
            operator = ">"
            other = relation.other
        elif isinstance(relation, NotEqualCell):
            operator = "<>"
            other = relation.other
        else:
            raise _fail(address, f"unsupported relation {type(relation).__name__}")
        partner = normalize_cell_type_env_key(other)
        if partner not in written:
            raise _fail(address, f"relation partner {other} is not in the workbook")
        clauses.append(f"{coord}{operator}{_formula_ref(partner, host_sheet=host_sheet)}")
    return clauses


def _native_interval(
    address: str,
    low: float | int | None,
    high: float | int | None,
    *,
    decimal: bool,
) -> _Rule:
    kind = "decimal" if decimal else "whole"
    error = _DECIMAL_ERROR if decimal else _WHOLE_ERROR
    if low is not None and high is not None:
        if low > high:
            raise _fail(address, "interval minimum is greater than its maximum")
        return _Rule(
            kind,
            "between",
            _format_number(address, low),
            _format_number(address, high),
            False,
            error,
        )
    if low is not None:
        return _Rule(kind, "greaterThanOrEqual", _format_number(address, low), None, False, error)
    if high is not None:
        return _Rule(kind, "lessThanOrEqual", _format_number(address, high), None, False, error)
    raise _fail(address, "interval has no bounds")


def _bound_clauses(
    address: str,
    coord: str,
    low: float | int | None,
    high: float | int | None,
    *,
    whole: bool,
) -> list[str]:
    if low is not None and high is not None and low > high:
        raise _fail(address, "interval minimum is greater than its maximum")
    clauses: list[str] = []
    if low is not None:
        clauses.append(f"{coord}>={_format_number(address, low)}")
    if high is not None:
        clauses.append(f"{coord}<={_format_number(address, high)}")
    if whole:
        # Native `whole` validation rejects non-integers; custom formulas that
        # only compare bounds do not, so require integrality explicitly.
        clauses.append(f"INT({coord})={coord}")
    return clauses


def _needs_custom_enum(values: list[object]) -> bool:
    return any(isinstance(value, str) and _has_list_delimiter(value) for value in values)


def _has_list_delimiter(value: str) -> bool:
    return "," in value or "\n" in value or "\r" in value


def _split_blanks(values: frozenset[object]) -> tuple[list[object], bool]:
    present: list[object] = []
    allow_blank = False
    for value in values:
        if value is None or value == "":
            allow_blank = True
        else:
            present.append(value)
    present.sort(key=_sort_key)
    return present, allow_blank


def _sort_key(value: object) -> tuple[object, ...]:
    if isinstance(value, bool):
        return (0, value)
    if isinstance(value, int):
        return (1, value)
    if isinstance(value, float):
        return (2, value)
    if isinstance(value, str):
        return (3, value)
    return (4, type(value).__name__, str(value))


def _value_class(address: str, values: list[object]) -> str:
    kinds = {_scalar_class(address, value) for value in values}
    if len(kinds) != 1:
        raise _fail(address, f"enum values mix types {sorted(kinds)}")
    return next(iter(kinds))


def _scalar_class(address: str, value: object) -> str:
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, int):
        return "int"
    if isinstance(value, float):
        return "float"
    if isinstance(value, str):
        return "str"
    raise _fail(address, f"unsupported enum value type {type(value).__name__}")


def _inline_list(address: str, values: list[object]) -> str:
    parts: list[str] = []
    for value in values:
        if isinstance(value, str):
            if _has_list_delimiter(value):
                raise _fail(
                    address,
                    f"enum value {value!r} contains a list delimiter and cannot be "
                    "an inline Excel list",
                )
            parts.append(value.replace('"', '""'))
        elif isinstance(value, bool):
            raise _fail(address, "boolean enums are written as custom formulas")
        elif isinstance(value, int):
            parts.append(str(value))
        elif isinstance(value, float):
            text = _format_number(address, value)
            if _has_list_delimiter(text):
                raise _fail(
                    address,
                    f"enum value {value!r} contains a list delimiter and cannot be "
                    "an inline Excel list",
                )
            parts.append(text)
        else:
            raise _fail(address, f"unsupported enum value type {type(value).__name__}")
    return '"' + ",".join(parts) + '"'


def _enum_term(address: str, coord: str, value: object) -> str:
    if isinstance(value, bool):
        literal = "TRUE" if value else "FALSE"
        return f"{coord}={literal}"
    if isinstance(value, str):
        if "\n" in value or "\r" in value:
            raise _fail(
                address,
                f"enum value {value!r} contains a newline and cannot be written",
            )
        return f'{coord}="{value.replace(chr(34), chr(34) * 2)}"'
    if isinstance(value, int):
        return f"{coord}={value}"
    if isinstance(value, float):
        return f"{coord}={_format_number(address, value)}"
    raise _fail(address, f"unsupported enum value type {type(value).__name__}")


def _or(terms: list[str]) -> str:
    if len(terms) == 1:
        return terms[0]
    return "OR(" + ",".join(terms) + ")"


def _and(clauses: list[str]) -> str:
    if len(clauses) == 1:
        return clauses[0]
    return "AND(" + ",".join(clauses) + ")"


def _formula_ref(address: str, *, host_sheet: str) -> str:
    sheet, coord = parse_address(address)
    if sheet == host_sheet:
        return coord
    return f"{quote_sheet_if_needed(sheet)}!{coord}"


def _format_number(address: str, value: float | int) -> str:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise _fail(address, f"unsupported numeric bound {value!r}")
    if isinstance(value, float):
        if not math.isfinite(value):
            raise _fail(address, f"numeric bound {value!r} is not finite")
        return format(value, ".12g")
    return str(value)


def _checked(address: str, rule: _Rule) -> _Rule:
    for label, formula in (("formula1", rule.formula1), ("formula2", rule.formula2)):
        if formula is not None and len(formula) > _FORMULA_LIMIT:
            raise _fail(
                address,
                f"Excel {label} is {len(formula)} characters (limit {_FORMULA_LIMIT})",
            )
    if len(rule.error) > _FORMULA_LIMIT:
        raise _fail(address, "Excel error message exceeds 255 characters")
    return rule


def _chunk_sqref(coords: list[str]) -> list[str]:
    ordered = sorted(set(coords), key=_coord_key)
    chunks: list[list[str]] = []
    current: list[str] = []
    length = 0
    for coord in ordered:
        extra = len(coord) if not current else len(coord) + 1
        if current and length + extra > _SQREF_LIMIT:
            chunks.append(current)
            current = [coord]
            length = len(coord)
        else:
            current.append(coord)
            length += extra
    if current:
        chunks.append(current)
    return [" ".join(chunk) for chunk in chunks]


def _coord_key(coord: str) -> tuple[int, int]:
    column, row = coordinate_from_string(coord)
    return (column_index_from_string(column), int(row))


def _fail(address: str, message: str) -> ValueError:
    if address:
        return ValueError(f"Cannot write cell validation at {address}: {message}")
    return ValueError(f"Cannot write cell validation: {message}")
