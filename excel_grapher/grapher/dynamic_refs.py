from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass
from itertools import product
from typing import Any

from openpyxl.utils.cell import coordinate_to_tuple

from excel_grapher.core.addressing import (
    WorkbookBoundsProtocol,
    indirect_text_to_range,
    offset_range,
)
from excel_grapher.core.cell_types import (
    CellKind,
    CellTypeEnv,
    IntIntervalDomain,
)
from excel_grapher.core.expr_eval import Unsupported, evaluate_expr
from excel_grapher.core.formula_ast import (
    AstNode,
    CellRefNode,
    FunctionCallNode,
    RangeNode,
)
from excel_grapher.core.formula_ast import (
    parse as parse_ast,
)
from excel_grapher.core.types import ExcelRange, XlError

from .parser import _find_function_calls_with_spans


class DynamicRefError(ValueError):
    """Raised when dynamic reference analysis cannot proceed."""


@dataclass(frozen=True)
class DynamicRefLimits:
    max_branches: int = 1024
    max_cells: int = 10_000
    max_depth: int = 10


@dataclass(frozen=True)
class GlobalWorkbookBounds(WorkbookBoundsProtocol):
    """Simple bounds implementation using Excel's global sheet limits."""

    sheet: str
    min_row: int = 1
    max_row: int = 1_048_576  # Excel row limit
    min_col: int = 1
    max_col: int = 16_384  # Excel column limit


def infer_dynamic_offset_targets(
    formula: str,
    *,
    current_sheet: str,
    cell_type_env: CellTypeEnv,
    limits: DynamicRefLimits | None = None,
    bounds: WorkbookBoundsProtocol | None = None,
) -> set[str]:
    """Infer the union of all possible OFFSET targets for a formula.

    This helper is intentionally focused and conservative:
    - Only OFFSET calls are analysed (INDIRECT is currently ignored).
    - Arguments may use a small Excel expression subset supported by
      ``core.expr_eval.evaluate_expr``.
    - All leaf cells referenced by OFFSET arguments must have a numeric
      domain in ``cell_type_env``; otherwise DynamicRefError is raised.
    - Integer interval domains must be finite and small enough to enumerate.
    """

    if not isinstance(formula, str) or not formula.startswith("="):
        return set()

    lim = limits or DynamicRefLimits()
    out: set[str] = set()

    calls = _find_function_calls_with_spans(formula, {"OFFSET"})
    for fn, inner, _span in calls:
        if fn != "OFFSET":
            continue
        targets = _infer_single_offset_call(
            inner,
            current_sheet=current_sheet,
            cell_type_env=cell_type_env,
            limits=lim,
            bounds=bounds,
        )
        out |= targets
        if len(out) > lim.max_cells:
            raise DynamicRefError(
                f"Dynamic ref cells exceed limit ({len(out)} > {lim.max_cells})"
            )

    return out


def _infer_single_offset_call(
    inner_args: str,
    *,
    current_sheet: str,
    cell_type_env: CellTypeEnv,
    limits: DynamicRefLimits,
    bounds: WorkbookBoundsProtocol | None,
) -> set[str]:
    """Infer targets for a single OFFSET(...) call body."""

    args = _split_top_level_args(inner_args)
    if args is None or len(args) < 3 or len(args) > 5:
        raise DynamicRefError("OFFSET expects 3 to 5 arguments")

    base_expr, rows_expr, cols_expr = args[0], args[1], args[2]
    height_expr = args[3] if len(args) >= 4 else ""
    width_expr = args[4] if len(args) >= 5 else ""

    base_ast = parse_ast("=" + base_expr)
    base_range = _base_to_range(base_ast, current_sheet=current_sheet)

    rows_ast = parse_ast("=" + rows_expr)
    cols_ast = parse_ast("=" + cols_expr)
    height_ast = parse_ast("=" + height_expr) if height_expr else None
    width_ast = parse_ast("=" + width_expr) if width_expr else None

    # Collect all leaf addresses that influence arguments.
    leaf_addrs: set[str] = set()
    leaf_addrs |= _collect_addresses(rows_ast)
    leaf_addrs |= _collect_addresses(cols_ast)
    if height_ast is not None:
        leaf_addrs |= _collect_addresses(height_ast)
    if width_ast is not None:
        leaf_addrs |= _collect_addresses(width_ast)

    domains = _build_domains(leaf_addrs, cell_type_env, limits)

    if bounds is None:
        bounds = GlobalWorkbookBounds(sheet=base_range.sheet)

    targets: set[str] = set()

    for assignment in _enumerate_assignments(domains.values(), limits):
        addr_to_value = dict(zip(domains.keys(), assignment, strict=False))

        def get_cell_value(addr: str, addr_to_value_map=addr_to_value) -> float:
            try:
                return addr_to_value_map[addr]
            except KeyError as exc:
                raise DynamicRefError(
                    f"OFFSET argument formula references cell without domain: {addr!r}"
                ) from exc

        rows_val = _eval_arg(rows_ast, get_cell_value, limits)
        cols_val = _eval_arg(cols_ast, get_cell_value, limits)
        height_val = _eval_arg(height_ast, get_cell_value, limits) if height_ast is not None else None
        width_val = _eval_arg(width_ast, get_cell_value, limits) if width_ast is not None else None

        if isinstance(rows_val, XlError) or isinstance(cols_val, XlError):
            continue
        if isinstance(height_val, XlError) or isinstance(width_val, XlError):
            continue

        result = offset_range(
            base_range,
            rows=rows_val,
            cols=cols_val,
            height=height_val,
            width=width_val,
            bounds=bounds,
        )
        if isinstance(result, ExcelRange):
            targets |= set(result.cell_addresses())

    return targets


def _eval_arg(
    node: AstNode | None,
    get_cell_value,
    limits: DynamicRefLimits,
) -> float | XlError:
    if node is None:
        return 0.0

    value = evaluate_expr(node, get_cell_value=get_cell_value, max_depth=limits.max_depth)
    if isinstance(value, Unsupported):
        raise DynamicRefError(f"Unsupported argument expression: {value.reason or ''}")
    if isinstance(value, XlError):
        return value
    if isinstance(value, (int, float)):
        return float(value)
    raise DynamicRefError(f"Non-numeric OFFSET argument result: {value!r}")


def _base_to_range(base_ast: AstNode, *, current_sheet: str) -> ExcelRange:
    if isinstance(base_ast, CellRefNode):
        sheet, coord = _split_address(base_ast.address, current_sheet=current_sheet)
        row, col = coordinate_to_tuple(coord)
        return ExcelRange(
            sheet=sheet,
            start_row=row,
            start_col=col,
            end_row=row,
            end_col=col,
        )
    if isinstance(base_ast, RangeNode):
        try:
            s1, coord1 = base_ast.start.split("!", 1)
            s2, coord2 = base_ast.end.split("!", 1)
        except ValueError as exc:
            raise DynamicRefError("OFFSET base range must be a single-sheet A1 range") from exc
        if s1 != s2:
            raise DynamicRefError("OFFSET base range must be on a single sheet")
        row1, col1 = coordinate_to_tuple(coord1)
        row2, col2 = coordinate_to_tuple(coord2)
        start_row, end_row = sorted((row1, row2))
        start_col, end_col = sorted((col1, col2))
        return ExcelRange(
            sheet=s1,
            start_row=start_row,
            start_col=start_col,
            end_row=end_row,
            end_col=end_col,
        )
    raise DynamicRefError("OFFSET base must be a cell or range reference")


def _split_address(addr: str, *, current_sheet: str) -> tuple[str, str]:
    if "!" in addr:
        sheet, coord = addr.split("!", 1)
        return sheet, coord
    return current_sheet, addr


def _collect_addresses(node: AstNode) -> set[str]:
    addrs: set[str] = set()

    def visit(n: AstNode) -> None:
        if isinstance(n, CellRefNode):
            addrs.add(n.address)
            return
        if isinstance(n, RangeNode):
            try:
                sheet, coord_start = n.start.split("!", 1)
                _sheet2, coord_end = n.end.split("!", 1)
            except ValueError:
                return
            row1, col1 = coordinate_to_tuple(coord_start)
            row2, col2 = coordinate_to_tuple(coord_end)
            rlo, rhi = sorted((row1, row2))
            clo, chi = sorted((col1, col2))
            for r in range(rlo, rhi + 1):
                for c in range(clo, chi + 1):
                    # coordinate_to_tuple gives (row, col); we need back to A1
                    from openpyxl.utils.cell import get_column_letter

                    col_letter = get_column_letter(c)
                    addrs.add(f"{sheet}!{col_letter}{r}")
            return
        if isinstance(n, FunctionCallNode):
            for arg in n.args:
                visit(arg)
            return
        # Binary/unary ops and other nodes: recurse into children where present.
        if hasattr(n, "left") and hasattr(n, "right"):
            visit(n.left)  # type: ignore[arg-type]
            visit(n.right)  # type: ignore[arg-type]
        if hasattr(n, "operand"):
            visit(n.operand)  # type: ignore[arg-type]

    visit(node)
    return addrs


def _build_domains(
    addrs: Iterable[str],
    env: CellTypeEnv,
    limits: DynamicRefLimits,
) -> dict[str, list[int]]:
    domains: dict[str, list[int]] = {}
    for addr in addrs:
        ct = env.get(addr)
        if ct is None:
            raise DynamicRefError(f"Missing CellType for {addr!r}")
        if ct.kind is not CellKind.NUMBER:
            raise DynamicRefError(f"CellType for {addr!r} must be numeric, got {ct.kind!r}")
        vals: list[int]
        if ct.enum is not None:
            vals = [int(v) for v in ct.enum.values]
        elif ct.interval is not None:
            vals = _interval_to_values(ct.interval, limits)
        else:
            raise DynamicRefError(f"CellType for {addr!r} must have an interval or enum domain")
        if not vals:
            raise DynamicRefError(f"Empty domain for {addr!r}")
        domains[addr] = sorted(vals)

    total = 1
    for vs in domains.values():
        total *= len(vs)
        if total > limits.max_branches:
            raise DynamicRefError(
                f"Dynamic ref branches exceed limit ({total} > {limits.max_branches})"
            )
    return domains


def _interval_to_values(interval: IntIntervalDomain, limits: DynamicRefLimits) -> list[int]:
    if interval.min is None or interval.max is None:
        raise DynamicRefError("Unbounded intervals are not supported for dynamic refs")
    lo, hi = int(interval.min), int(interval.max)
    if hi < lo:
        raise DynamicRefError(f"Invalid interval domain [{lo}, {hi}]")
    count = hi - lo + 1
    if count > limits.max_branches:
        raise DynamicRefError(
            f"Interval size {count} exceeds branch limit {limits.max_branches}"
        )
    return list(range(lo, hi + 1))


def _enumerate_assignments(
    domains: Iterable[list[int]],
    limits: DynamicRefLimits,
) -> Iterable[tuple[int, ...]]:
    # Domains size has already been checked in _build_domains; this is a thin wrapper.
    return product(*domains)


def infer_dynamic_indirect_targets(
    formula: str,
    *,
    current_sheet: str,
    cell_type_env: CellTypeEnv,
    limits: DynamicRefLimits | None = None,
    bounds: WorkbookBoundsProtocol | None = None,
) -> set[str]:
    """Infer the union of all possible INDIRECT targets for a formula."""

    if not isinstance(formula, str) or not formula.startswith("="):
        return set()

    lim = limits or DynamicRefLimits()
    out: set[str] = set()

    calls = _find_function_calls_with_spans(formula, {"INDIRECT"})
    for fn, inner, _span in calls:
        if fn != "INDIRECT":
            continue
        targets = _infer_single_indirect_call(
            inner,
            current_sheet=current_sheet,
            cell_type_env=cell_type_env,
            limits=lim,
            bounds=bounds,
        )
        out |= targets
        if len(out) > lim.max_cells:
            raise DynamicRefError(
                f"Dynamic ref cells exceed limit ({len(out)} > {lim.max_cells})"
            )

    return out


def _infer_single_indirect_call(
    inner_args: str,
    *,
    current_sheet: str,
    cell_type_env: CellTypeEnv,
    limits: DynamicRefLimits,
    bounds: WorkbookBoundsProtocol | None,
) -> set[str]:
    args = _split_top_level_args(inner_args)
    if args is None or len(args) < 1 or len(args) > 2:
        raise DynamicRefError("INDIRECT expects 1 or 2 arguments")

    text_expr = args[0]
    a1_expr = args[1] if len(args) == 2 else ""

    text_ast = parse_ast("=" + text_expr)
    a1_ast = parse_ast("=" + a1_expr) if a1_expr else None

    leaf_addrs: set[str] = set()
    leaf_addrs |= _collect_addresses(text_ast)
    if a1_ast is not None:
        leaf_addrs |= _collect_addresses(a1_ast)

    domains = _build_value_domains(leaf_addrs, cell_type_env, limits)

    targets: set[str] = set()

    for assignment in _enumerate_value_assignments(domains.values(), limits):
        addr_to_value = dict(zip(domains.keys(), assignment, strict=False))

        def get_cell_value(addr: str, addr_to_value_map=addr_to_value) -> Any:
            try:
                return addr_to_value_map[addr]
            except KeyError as exc:
                raise DynamicRefError(
                    f"INDIRECT argument formula references cell without domain: {addr!r}"
                ) from exc

        text_value = evaluate_expr(text_ast, get_cell_value=get_cell_value, max_depth=limits.max_depth)
        if isinstance(text_value, Unsupported):
            raise DynamicRefError(f"Unsupported INDIRECT text expression: {text_value.reason or ''}")
        if isinstance(text_value, XlError):
            continue
        if not isinstance(text_value, str):
            raise DynamicRefError(f"INDIRECT text argument must be a string, got {type(text_value).__name__}")

        if a1_ast is None:
            a1_flag = True
        else:
            a1_value = evaluate_expr(a1_ast, get_cell_value=get_cell_value, max_depth=limits.max_depth)
            if isinstance(a1_value, Unsupported):
                raise DynamicRefError(f"Unsupported INDIRECT A1/R1C1 flag expression: {a1_value.reason or ''}")
            if isinstance(a1_value, XlError):
                continue
            if isinstance(a1_value, bool):
                a1_flag = a1_value
            elif isinstance(a1_value, (int, float)):
                a1_flag = bool(a1_value)
            else:
                raise DynamicRefError(
                    f"INDIRECT A1/R1C1 flag must be boolean or numeric, got {type(a1_value).__name__}"
                )

        # Derive per-call bounds so sheet-qualified references are not rejected.
        sheet_for_bounds = _sheet_from_indirect_text(text_value, current_sheet=current_sheet)
        if bounds is None:
            local_bounds = GlobalWorkbookBounds(sheet=sheet_for_bounds)
        else:
            local_bounds = GlobalWorkbookBounds(
                sheet=sheet_for_bounds,
                min_row=bounds.min_row,
                max_row=bounds.max_row,
                min_col=bounds.min_col,
                max_col=bounds.max_col,
            )

        result = indirect_text_to_range(text_value, a1_flag, bounds=local_bounds)
        if isinstance(result, ExcelRange):
            targets |= set(result.cell_addresses())

    return targets


def _sheet_from_indirect_text(text: str, *, current_sheet: str) -> str:
    raw = text.strip()
    if not raw:
        return current_sheet
    if "!" in raw:
        sheet_text, _addr = raw.split("!", 1)
        return sheet_text or current_sheet
    return current_sheet


def _build_value_domains(
    addrs: Iterable[str],
    env: CellTypeEnv,
    limits: DynamicRefLimits,
) -> dict[str, list[Any]]:
    domains: dict[str, list[Any]] = {}
    for addr in addrs:
        ct = env.get(addr)
        if ct is None:
            raise DynamicRefError(f"Missing CellType for {addr!r}")
        values: list[Any]
        if ct.enum is not None:
            values = list(ct.enum.values)
        elif ct.interval is not None:
            values = _interval_to_values(ct.interval, limits)
        else:
            raise DynamicRefError(
                f"CellType for {addr!r} must have an interval or enum domain for INDIRECT analysis"
            )
        if not values:
            raise DynamicRefError(f"Empty domain for {addr!r}")
        domains[addr] = values

    total = 1
    for vs in domains.values():
        total *= len(vs)
        if total > limits.max_branches:
            raise DynamicRefError(
                f"Dynamic ref branches exceed limit ({total} > {limits.max_branches})"
            )
    return domains


def _enumerate_value_assignments(
    domains: Iterable[list[Any]],
    limits: DynamicRefLimits,
) -> Iterable[tuple[Any, ...]]:
    return product(*domains)


def _split_top_level_args(s: str) -> list[str] | None:
    """Minimal top-level argument splitter mirroring parser._split_top_level_args."""

    buf: list[str] = []
    args: list[str] = []
    depth = 0
    in_str = False
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == '"':
            in_str = not in_str
            buf.append(ch)
            i += 1
            continue
        if in_str:
            buf.append(ch)
            i += 1
            continue
        if ch == "(":
            depth += 1
            buf.append(ch)
            i += 1
            continue
        if ch == ")":
            if depth == 0:
                return None
            depth -= 1
            buf.append(ch)
            i += 1
            continue
        if ch == "," and depth == 0:
            args.append("".join(buf).strip())
            buf = []
            i += 1
            continue
        buf.append(ch)
        i += 1
    if in_str or depth != 0:
        return None
    args.append("".join(buf).strip())
    return [a for a in args if a != ""]

