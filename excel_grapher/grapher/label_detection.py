"""Optional row/column label detection for dependency graph nodes.

Label detection is opt-in via :func:`~excel_grapher.grapher.builder.create_dependency_graph`.
When enabled, detected labels are stored on each node under ``metadata["row_labels"]`` and
``metadata["column_labels"]`` as lists of strings.
"""

from __future__ import annotations

import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field, fields, is_dataclass
from typing import Any, Literal, Protocol, runtime_checkable

import fastpyxl.utils.cell
from fastpyxl.worksheet.worksheet import Worksheet

from .blank_ranges import normalize_blank_range_specs

# ---------------------------------------------------------------------------
# Public config types
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class RegionSpec:
    """One sheet-bound rectangular region with inclusive numeric bounds."""

    sheet: str
    min_row: int
    max_row: int
    min_col: int
    max_col: int


@dataclass(frozen=True)
class RegionSelector:
    """Region membership: any ``include`` rectangle minus ``exclude`` rectangles."""

    include: tuple[RegionSpec, ...] = ()
    exclude: tuple[RegionSpec, ...] = ()


@dataclass(frozen=True)
class RegionLabelParams:
    """Optional parameters for region-scoped built-in behaviors on a :class:`BehaviorRule`."""

    header_rows: tuple[int, ...] = ()
    label_columns: tuple[str, ...] = ()
    min_row: int | None = None
    max_row: int | None = None
    no_hierarchy_columns: tuple[str, ...] = ()


@dataclass(frozen=True)
class BehaviorRule:
    name: str
    selector: RegionSelector
    behaviors: tuple[str, ...]
    stop_after_match: bool = False
    region_params: RegionLabelParams | None = None


@dataclass(frozen=True)
class LabelDetectionConfig:
    enabled: bool = False
    merge_policy: Literal["append_dedupe", "append_dedupe_reverse", "replace"] = "append_dedupe"
    fallback_behaviors: tuple[str, ...] = ("left_edge_scan", "top_edge_scan")
    rules: tuple[BehaviorRule, ...] = ()


@dataclass
class LabelDetectionState:
    """Caches shared across behaviors for one ``create_dependency_graph`` invocation."""

    offset_maps: dict[tuple[str, int], dict[int, int]] = field(default_factory=dict)
    hierarchies: dict[tuple[str, str], dict[int, tuple[str, ...]]] = field(default_factory=dict)


@dataclass(frozen=True)
class LabelDetectionContext:
    """Workbook coordinates and worksheet handles for one graph node."""

    key: str
    sheet: str
    row: int
    col: int
    """1-based column index (same convention as *fastpyxl* column indices)."""

    ws_values: Worksheet | None
    ws_formulas: Worksheet | None
    region_params: RegionLabelParams | None
    state: LabelDetectionState


@dataclass(frozen=True)
class LabelResult:
    row_labels: tuple[str, ...] = ()
    column_labels: tuple[str, ...] = ()


@runtime_checkable
class LabelDetectionBehavior(Protocol):
    name: str

    def detect(self, ctx: LabelDetectionContext) -> LabelResult: ...


def region_specs_from_ranges(ranges: Sequence[str]) -> tuple[RegionSpec, ...]:
    """Parse sheet-qualified A1 ranges into normalized :class:`RegionSpec` rectangles.

    Uses the same parsing rules as :func:`~excel_grapher.grapher.blank_ranges.normalize_blank_range_specs`.
    """
    rects = normalize_blank_range_specs(list(ranges))
    return tuple(
        RegionSpec(sheet=sh, min_row=r1, max_row=r2, min_col=c1, max_col=c2)
        for sh, r1, c1, r2, c2 in rects
    )


def label_detection_config_to_jsonable(cfg: LabelDetectionConfig | None) -> dict[str, Any] | None:
    """Return a JSON-serializable dict for ``GraphCacheMeta["extraction_params"]``."""

    if cfg is None:
        return None
    return _dataclass_to_jsonable(cfg)


# ---------------------------------------------------------------------------
# Label text helpers (shared by heuristics and region behaviors)
# ---------------------------------------------------------------------------

EXCEL_ERRORS = frozenset(
    {
        "#DIV/0!",
        "#REF!",
        "#NAME?",
        "#VALUE!",
        "#N/A",
        "#NULL!",
        "#NUM!",
        "#GETTING_DATA",
        "#SPILL!",
        "#CALC!",
    }
)

PLACEHOLDER_PATTERNS = frozenset(
    {
        "...",
        "…",
        "---",
        "--",
        "-",
        "n/a",
        "N/A",
        "n.a.",
        "N.A.",
        "TBD",
        "tbd",
        "[insert filepath]",
    }
)


def is_valid_label(text: str) -> bool:
    stripped = text.strip()
    if not stripped:
        return False
    if stripped in EXCEL_ERRORS:
        return False
    if stripped in PLACEHOLDER_PATTERNS:
        return False
    return not all(c in ".-…" for c in stripped)


def is_year_like(value: int | float) -> bool:
    if not isinstance(value, int) or isinstance(value, bool):
        return False
    return 1900 <= value <= 2100


def dedupe_preserve_order(labels: Sequence[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for label in labels:
        if label in seen:
            continue
        seen.add(label)
        out.append(label)
    return out


def _merged_anchor_value(ws: Worksheet, row: int, col: int) -> Any | None:
    merged_ranges = getattr(getattr(ws, "merged_cells", None), "ranges", ())
    for merged_range in merged_ranges:
        if (
            merged_range.min_row <= row <= merged_range.max_row
            and merged_range.min_col <= col <= merged_range.max_col
        ):
            return ws.cell(row=merged_range.min_row, column=merged_range.min_col).value
    return None


def _scan_cell_value(ws: Worksheet, row: int, col: int) -> Any | None:
    value = ws.cell(row=row, column=col).value
    if value is not None:
        return value
    return _merged_anchor_value(ws, row, col)


def _get_effective_indent(cell: Any) -> int:
    alignment_indent = int(cell.alignment.indent) if cell.alignment and cell.alignment.indent else 0
    text_indent = 0
    if isinstance(cell.value, str):
        stripped = cell.value.lstrip()
        leading_spaces = len(cell.value) - len(stripped)
        if leading_spaces > 0:
            text_indent = 1
    return alignment_indent + text_indent


def _style_rank(cell: Any) -> int:
    """Return style priority rank: normal < italic < bold."""
    if cell.font and cell.font.bold:
        return 2
    if cell.font and cell.font.italic:
        return 1
    return 0


_TRAILING_PARENS_RE = re.compile(r"\s*\([^)]*\)\s*$")


def _normalize_label_text(text: str) -> str:
    cleaned = _TRAILING_PARENS_RE.sub("", text).strip()
    return cleaned if cleaned else text.strip()


def _build_label_hierarchy(
    ws: Worksheet,
    label_col_idx: int,
    min_row: int | None,
    max_row: int | None,
) -> dict[int, list[str]]:
    start = min_row or 1
    end = max_row or (ws.max_row or 1)
    stack: list[tuple[int, str]] = []
    hierarchy: dict[int, list[str]] = {}

    for row in range(start, end + 1):
        cell = ws.cell(row=row, column=label_col_idx)
        if cell.value is None:
            continue
        if not isinstance(cell.value, str):
            continue
        text = cell.value.strip()
        if not text or not is_valid_label(text):
            continue

        indent = _get_effective_indent(cell)
        while stack and stack[-1][0] >= indent:
            stack.pop()
        hierarchy[row] = [label for _, label in stack]
        stack.append((indent, text))

    return hierarchy


_OFFSET_PREFIX = "offset:"


# ---------------------------------------------------------------------------
# Region matching
# ---------------------------------------------------------------------------


def _cell_in_region(sheet: str, row: int, col: int, spec: RegionSpec) -> bool:
    return (
        sheet == spec.sheet
        and spec.min_row <= row <= spec.max_row
        and spec.min_col <= col <= spec.max_col
    )


def selector_matches(sheet: str, row: int, col: int, selector: RegionSelector) -> bool:
    if not selector.include:
        return False
    for ex in selector.exclude:
        if _cell_in_region(sheet, row, col, ex):
            return False
    return any(_cell_in_region(sheet, row, col, inc) for inc in selector.include)


# ---------------------------------------------------------------------------
# Heuristic scans (values worksheet)
# ---------------------------------------------------------------------------


def _heuristic_row_labels(ws: Worksheet, row: int, col: int) -> list[str]:
    labels: list[str] = []
    current_col = col - 1
    while current_col >= 1:
        cell_value = _scan_cell_value(ws, row, current_col)
        if cell_value is None or (isinstance(cell_value, str) and cell_value.strip() == ""):
            break
        if isinstance(cell_value, str):
            text = cell_value.strip()
            if is_valid_label(text):
                labels.append(text)
        elif isinstance(cell_value, (int, float)) and not isinstance(cell_value, bool):
            if is_year_like(cell_value):
                labels.append(str(cell_value))
            elif labels:
                labels = []
        elif not isinstance(cell_value, bool):
            text = str(cell_value)
            if is_valid_label(text):
                labels.append(text)
        current_col -= 1
    return dedupe_preserve_order(labels)


def _heuristic_column_labels(ws: Worksheet, row: int, col: int) -> list[str]:
    labels: list[str] = []
    current_row = row - 1
    while current_row >= 1:
        cell_value = _scan_cell_value(ws, current_row, col)
        if cell_value is None or (isinstance(cell_value, str) and cell_value.strip() == ""):
            break
        if isinstance(cell_value, str):
            text = cell_value.strip()
            if is_valid_label(text):
                labels.append(text)
        elif isinstance(cell_value, (int, float)) and not isinstance(cell_value, bool):
            if is_year_like(cell_value):
                labels.append(str(cell_value))
            elif labels:
                labels = []
        elif not isinstance(cell_value, bool):
            text = str(cell_value)
            if is_valid_label(text):
                labels.append(text)
        current_row -= 1
    return dedupe_preserve_order(labels)


def _full_row_labels(ws: Worksheet, row: int, col: int) -> list[str]:
    labels_with_col: list[tuple[int, str]] = []
    max_col = ws.max_column or col

    current_col = col + 1
    while current_col <= max_col:
        cell_value = _scan_cell_value(ws, row, current_col)
        if cell_value is None or (isinstance(cell_value, str) and cell_value.strip() == ""):
            break
        if isinstance(cell_value, str):
            text = cell_value.strip()
            if is_valid_label(text):
                labels_with_col.append((current_col, text))
        elif not isinstance(cell_value, (int, float, bool)):
            text = str(cell_value)
            if is_valid_label(text):
                labels_with_col.append((current_col, text))
        current_col += 1

    current_col = col - 1
    while current_col >= 1:
        cell_value = _scan_cell_value(ws, row, current_col)
        if cell_value is None or (isinstance(cell_value, str) and cell_value.strip() == ""):
            break
        if isinstance(cell_value, str):
            text = cell_value.strip()
            if is_valid_label(text):
                labels_with_col.append((current_col, text))
        elif not isinstance(cell_value, (int, float, bool)):
            text = str(cell_value)
            if is_valid_label(text):
                labels_with_col.append((current_col, text))
        current_col -= 1

    ordered = [label for _, label in sorted(labels_with_col, key=lambda x: x[0], reverse=True)]
    return dedupe_preserve_order(ordered)


def _find_leftmost_label_column(ws: Worksheet, row: int, col: int) -> int | None:
    """Return the leftmost text-label column before ``col`` on ``row``; stops at blanks."""
    leftmost: int | None = None
    current_col = col - 1
    while current_col >= 1:
        cell_value = _scan_cell_value(ws, row, current_col)
        if cell_value is None or (isinstance(cell_value, str) and cell_value.strip() == ""):
            break
        if isinstance(cell_value, str):
            text = cell_value.strip()
            if text and is_valid_label(text):
                leftmost = current_col
        elif isinstance(cell_value, (int, float)) and not isinstance(cell_value, bool):
            if is_year_like(cell_value):
                leftmost = current_col
        current_col -= 1
    return leftmost


def _left_then_up_label_text(value: Any) -> str | None:
    if isinstance(value, str):
        text = _normalize_label_text(value.strip())
        if text and is_valid_label(text):
            return text
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool) and is_year_like(value):
        return str(int(value))
    return None


def _left_edge_then_up_row_labels(
    ws: Worksheet, row: int, col: int, rp: RegionLabelParams | None
) -> list[str]:
    label_col_idx = _find_leftmost_label_column(ws, row, col)
    if label_col_idx is None:
        return []

    current_cell = ws.cell(row=row, column=label_col_idx)
    current_text = _left_then_up_label_text(current_cell.value)
    if current_text is None:
        return []

    collected: list[str] = [current_text]
    last_indent = _get_effective_indent(current_cell)
    last_style = _style_rank(current_cell)

    min_row = (rp.min_row if rp is not None and rp.min_row is not None else 1) or 1
    current_row = row - 1
    while current_row >= min_row:
        cell = ws.cell(row=current_row, column=label_col_idx)
        text = _left_then_up_label_text(cell.value)
        if text is None:
            break

        indent = _get_effective_indent(cell)
        style = _style_rank(cell)

        if indent == last_indent:
            if style == last_style:
                current_row -= 1
                continue
            if style > last_style:
                collected.append(text)
                last_style = style
                current_row -= 1
                continue
            break

        if indent < last_indent:
            collected.append(text)
            last_indent = indent
            last_style = style
            current_row -= 1
            continue

        break

    return dedupe_preserve_order(list(reversed(collected)))


def _find_topmost_label_row(ws: Worksheet, row: int, col: int, min_row: int = 1) -> int | None:
    """Return the topmost text/year label row above ``row`` on ``col``; stops at blanks."""
    topmost: int | None = None
    current_row = row - 1
    while current_row >= min_row:
        cell_value = _scan_cell_value(ws, current_row, col)
        if cell_value is None or (isinstance(cell_value, str) and cell_value.strip() == ""):
            break
        text = _left_then_up_label_text(cell_value)
        if text is not None:
            topmost = current_row
        current_row -= 1
    return topmost


def _top_edge_then_left_column_labels(
    ws: Worksheet, row: int, col: int, rp: RegionLabelParams | None
) -> list[str]:
    min_row = (rp.min_row if rp is not None and rp.min_row is not None else 1) or 1
    label_row_idx = _find_topmost_label_row(ws, row, col, min_row=min_row)
    if label_row_idx is None:
        return []

    current_cell_value = _scan_cell_value(ws, label_row_idx, col)
    current_text = _left_then_up_label_text(current_cell_value)
    if current_text is None:
        return []

    collected: list[str] = [current_text]

    min_col = 1
    if rp is not None and rp.label_columns:
        min_col = min(fastpyxl.utils.cell.column_index_from_string(c) for c in rp.label_columns)

    current_col = col - 1
    while current_col >= min_col:
        cell_value = _scan_cell_value(ws, label_row_idx, current_col)
        text = _left_then_up_label_text(cell_value)
        if text is None:
            break
        collected.append(text)
        current_col -= 1

    return dedupe_preserve_order(list(reversed(collected)))


def _full_column_labels(ws: Worksheet, row: int, col: int) -> list[str]:
    labels_with_row: list[tuple[int, str]] = []
    max_row = ws.max_row or row

    current_row = row + 1
    while current_row <= max_row:
        cell_value = _scan_cell_value(ws, current_row, col)
        if cell_value is None or (isinstance(cell_value, str) and cell_value.strip() == ""):
            break
        if isinstance(cell_value, str):
            text = cell_value.strip()
            if is_valid_label(text):
                labels_with_row.append((current_row, text))
        elif not isinstance(cell_value, (int, float, bool)):
            text = str(cell_value)
            if is_valid_label(text):
                labels_with_row.append((current_row, text))
        current_row += 1

    current_row = row - 1
    while current_row >= 1:
        cell_value = _scan_cell_value(ws, current_row, col)
        if cell_value is None or (isinstance(cell_value, str) and cell_value.strip() == ""):
            break
        if isinstance(cell_value, str):
            text = cell_value.strip()
            if is_valid_label(text):
                labels_with_row.append((current_row, text))
        elif not isinstance(cell_value, (int, float, bool)):
            text = str(cell_value)
            if is_valid_label(text):
                labels_with_row.append((current_row, text))
        current_row -= 1

    ordered = [label for _, label in sorted(labels_with_row, key=lambda x: x[0], reverse=True)]
    return dedupe_preserve_order(ordered)


def _right_edge_row_labels(ws: Worksheet, row: int, col: int) -> list[str]:
    labels: list[str] = []
    max_col = ws.max_column or col
    current_col = col + 1
    while current_col <= max_col:
        cell_value = _scan_cell_value(ws, row, current_col)
        if cell_value is None or (isinstance(cell_value, str) and cell_value.strip() == ""):
            break
        if isinstance(cell_value, str):
            text = cell_value.strip()
            if is_valid_label(text):
                labels.append(text)
        elif not isinstance(cell_value, (int, float, bool)):
            text = str(cell_value)
            if is_valid_label(text):
                labels.append(text)
        current_col += 1
    return dedupe_preserve_order(labels)


def _bottom_edge_column_labels(ws: Worksheet, row: int, col: int) -> list[str]:
    labels: list[str] = []
    max_row = ws.max_row or row
    current_row = row + 1
    while current_row <= max_row:
        cell_value = _scan_cell_value(ws, current_row, col)
        if cell_value is None or (isinstance(cell_value, str) and cell_value.strip() == ""):
            break
        if isinstance(cell_value, str):
            text = cell_value.strip()
            if is_valid_label(text):
                labels.append(text)
        elif not isinstance(cell_value, (int, float, bool)):
            text = str(cell_value)
            if is_valid_label(text):
                labels.append(text)
        current_row += 1
    return dedupe_preserve_order(labels)


# ---------------------------------------------------------------------------
# Built-in behaviors
# ---------------------------------------------------------------------------


class _HeuristicRowScan:
    name = "left_edge_scan"

    def detect(self, ctx: LabelDetectionContext) -> LabelResult:
        if ctx.ws_values is None:
            return LabelResult()
        rows = _heuristic_row_labels(ctx.ws_values, ctx.row, ctx.col)
        return LabelResult(row_labels=tuple(rows))


class _HeuristicColumnScan:
    name = "top_edge_scan"

    def detect(self, ctx: LabelDetectionContext) -> LabelResult:
        if ctx.ws_values is None:
            return LabelResult()
        cols = _heuristic_column_labels(ctx.ws_values, ctx.row, ctx.col)
        return LabelResult(column_labels=tuple(cols))


class _FullRowScan:
    name = "full_row_scan"

    def detect(self, ctx: LabelDetectionContext) -> LabelResult:
        if ctx.ws_values is None:
            return LabelResult()
        rows = _full_row_labels(ctx.ws_values, ctx.row, ctx.col)
        return LabelResult(row_labels=tuple(rows))


class _FullColumnScan:
    name = "full_column_scan"

    def detect(self, ctx: LabelDetectionContext) -> LabelResult:
        if ctx.ws_values is None:
            return LabelResult()
        cols = _full_column_labels(ctx.ws_values, ctx.row, ctx.col)
        return LabelResult(column_labels=tuple(cols))


class _RightEdgeScan:
    name = "right_edge_scan"

    def detect(self, ctx: LabelDetectionContext) -> LabelResult:
        if ctx.ws_values is None:
            return LabelResult()
        rows = _right_edge_row_labels(ctx.ws_values, ctx.row, ctx.col)
        return LabelResult(row_labels=tuple(rows))


class _BottomEdgeScan:
    name = "bottom_edge_scan"

    def detect(self, ctx: LabelDetectionContext) -> LabelResult:
        if ctx.ws_values is None:
            return LabelResult()
        rows = _bottom_edge_column_labels(ctx.ws_values, ctx.row, ctx.col)
        return LabelResult(row_labels=tuple(rows))


class _RegionHeaderRows:
    name = "region_header_rows"

    def detect(self, ctx: LabelDetectionContext) -> LabelResult:
        rp = ctx.region_params
        if rp is None or not rp.header_rows or ctx.ws_values is None:
            return LabelResult()
        col_labels: list[str] = []
        for header_row in rp.header_rows:
            hr_key = (ctx.sheet, header_row)
            offsets = ctx.state.offset_maps.get(hr_key, {})
            if ctx.col in offsets:
                col_labels.append(f"{_OFFSET_PREFIX}{offsets[ctx.col]}")
                continue
            cell_value = ctx.ws_values.cell(row=header_row, column=ctx.col).value
            if cell_value is None:
                continue
            if isinstance(cell_value, str):
                text = cell_value.strip()
                if text and is_valid_label(text):
                    col_labels.append(text)
            elif is_year_like(cell_value):
                col_labels.append(str(cell_value))
        return LabelResult(column_labels=tuple(dedupe_preserve_order(col_labels)))


class _LeftEdgeThenUpScan:
    name = "left_edge_then_up_scan"

    def detect(self, ctx: LabelDetectionContext) -> LabelResult:
        if ctx.ws_values is None:
            return LabelResult()
        rows = _left_edge_then_up_row_labels(ctx.ws_values, ctx.row, ctx.col, ctx.region_params)
        return LabelResult(row_labels=tuple(rows))


class _TopEdgeThenLeftScan:
    name = "top_edge_then_left_scan"

    def detect(self, ctx: LabelDetectionContext) -> LabelResult:
        if ctx.ws_values is None:
            return LabelResult()
        cols = _top_edge_then_left_column_labels(ctx.ws_values, ctx.row, ctx.col, ctx.region_params)
        return LabelResult(column_labels=tuple(cols))


def default_label_behaviors() -> tuple[LabelDetectionBehavior, ...]:
    return (
        _HeuristicRowScan(),
        _HeuristicColumnScan(),
        _FullRowScan(),
        _FullColumnScan(),
        _RightEdgeScan(),
        _BottomEdgeScan(),
        _RegionHeaderRows(),
        _LeftEdgeThenUpScan(),
        _TopEdgeThenLeftScan(),
    )


def build_label_behavior_registry(
    extra: Sequence[LabelDetectionBehavior] | None,
) -> dict[str, LabelDetectionBehavior]:
    reg: dict[str, LabelDetectionBehavior] = {}
    for b in default_label_behaviors():
        reg[b.name] = b
    for b in extra or ():
        reg[b.name] = b
    return reg


# ---------------------------------------------------------------------------
# Merge + orchestration
# ---------------------------------------------------------------------------


def _merge_results(
    policy: Literal["append_dedupe", "append_dedupe_reverse", "replace"],
    current_row: list[str],
    current_col: list[str],
    new: LabelResult,
) -> tuple[list[str], list[str]]:
    if policy == "replace":
        if new.row_labels:
            current_row = list(new.row_labels)
        if new.column_labels:
            current_col = list(new.column_labels)
        return current_row, current_col
    if new.row_labels:
        current_row = dedupe_preserve_order([*current_row, *new.row_labels])
    if new.column_labels:
        current_col = dedupe_preserve_order([*current_col, *new.column_labels])
    return current_row, current_col


def collect_labels_for_node(
    *,
    key: str,
    sheet: str,
    row: int,
    col: int,
    cfg: LabelDetectionConfig,
    registry: Mapping[str, LabelDetectionBehavior],
    state: LabelDetectionState,
    ws_values: Worksheet | None,
    ws_formulas: Worksheet | None,
) -> tuple[list[str], list[str]]:
    """Run configured rules and fallbacks; return ``(row_labels, column_labels)``."""

    row_out: list[str] = []
    col_out: list[str] = []
    any_selector_matched = False

    for rule in cfg.rules:
        if not selector_matches(sheet, row, col, rule.selector):
            continue
        any_selector_matched = True
        ctx = LabelDetectionContext(
            key=key,
            sheet=sheet,
            row=row,
            col=col,
            ws_values=ws_values,
            ws_formulas=ws_formulas,
            region_params=rule.region_params,
            state=state,
        )
        for bname in rule.behaviors:
            beh = registry.get(bname)
            if beh is None:
                raise ValueError(f"Unknown label detection behavior: {bname!r}")
            res = beh.detect(ctx)
            row_out, col_out = _merge_results(cfg.merge_policy, row_out, col_out, res)
        if rule.stop_after_match:
            break

    if not any_selector_matched:
        ctx_fb = LabelDetectionContext(
            key=key,
            sheet=sheet,
            row=row,
            col=col,
            ws_values=ws_values,
            ws_formulas=ws_formulas,
            region_params=None,
            state=state,
        )
        for bname in cfg.fallback_behaviors:
            beh = registry.get(bname)
            if beh is None:
                raise ValueError(f"Unknown label detection behavior: {bname!r}")
            res = beh.detect(ctx_fb)
            row_out, col_out = _merge_results(cfg.merge_policy, row_out, col_out, res)

    if cfg.merge_policy == "append_dedupe_reverse":
        row_out = list(reversed(row_out))
        col_out = list(reversed(col_out))
    return row_out, col_out


# ---------------------------------------------------------------------------
# JSON / cache helpers
# ---------------------------------------------------------------------------


def _dataclass_to_jsonable(obj: Any) -> Any:
    if obj is None:
        return None
    if isinstance(obj, (str, int, float, bool)):
        return obj
    if isinstance(obj, Mapping):
        return {str(k): _dataclass_to_jsonable(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_dataclass_to_jsonable(x) for x in obj]
    if is_dataclass(obj):
        out: dict[str, Any] = {}
        for f in fields(obj):
            out[f.name] = _dataclass_to_jsonable(getattr(obj, f.name))
        return out
    return obj
