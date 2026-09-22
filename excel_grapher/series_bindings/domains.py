"""Compile series bindings into the `CellTypeEnv` dynamic-ref inference consumes.

`domain` is series-level (`enum` / `between` / `real_between` / `from_workbook`).
`input.domain` is accepted and normalized to that key. `constant` implies
`from_workbook`. Series-level `relations` name a partner series; each declaring
cell compiles to `GreaterThanCell` or `NotEqualCell` at the same key.

Rules:

* Declared `domain` applies to every bound cell of the series.
* `input` series without `domain` use `input.value_map` workbook needles.
* `internal`, `output`, and `input.mode: override` contribute nothing unless
  `domain` is explicit.
* `relations` expand per bound cell. Missing partner cells at a key fail closed.
"""

from __future__ import annotations

from collections.abc import Iterator, Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Annotated, Any, Literal, cast, get_args, get_origin

from fastpyxl.utils.cell import column_index_from_string, coordinate_from_string

from excel_grapher.core.address_keys import parse_address, split_address_on_colon
from excel_grapher.core.cell_types import (
    Between,
    CellRelation,
    CellType,
    RealBetween,
    constraints_to_cell_type_env,
    normalize_cell_type_env_key,
)
from excel_grapher.series_bindings.canonical import bindings_canonical_sha256
from excel_grapher.series_bindings.geometry import expand_column_specs, expand_row_specs
from excel_grapher.series_bindings.normalize import (
    has_constant_direction,
    has_input_direction,
    input_mode,
)
from excel_grapher.series_bindings.ranges import (
    expand_bound_series_addresses,
    series_data_ranges,
)
from excel_grapher.series_bindings.relations import (
    RELATION_TYPES,
    SeriesRelationError,
    canonical_measure_dtype,
    iter_series_relations,
    missing_partner_cell_issues,
    raise_if_relation_errors,
    relation_cell_indexes,
    relation_declaration_issues,
    series_by_id,
)

if TYPE_CHECKING:
    from excel_grapher.grapher.graph import DependencyGraph

# `Literal[...]` subscripted with a runtime tuple is fine at runtime but not for type checkers.
_RuntimeLiteral: Any = cast(Any, Literal)

__all__ = [
    "SeriesDomainHandle",
    "SeriesDomainIndex",
    "SeriesRelationError",
    "cell_type_env_from_bindings",
    "compile_domain_spec",
    "declared_domain",
    "undomained_leaves",
]


@dataclass(frozen=True, slots=True)
class SeriesDomainHandle:
    """Pickle/JSON pointer to the bindings sidecar that owns cell domains."""

    workbook: str
    bindings: str
    canonical_sha256: str

    def to_dict(self) -> dict[str, str]:
        """Return a JSON-serializable mapping of handle fields."""
        return {
            "workbook": self.workbook,
            "bindings": self.bindings,
            "canonical_sha256": self.canonical_sha256,
        }

    @classmethod
    def from_dict(cls, payload: Mapping[str, Any]) -> SeriesDomainHandle:
        """Rebuild a handle from `to_dict` output."""
        return cls(
            workbook=str(payload["workbook"]),
            bindings=str(payload["bindings"]),
            canonical_sha256=str(payload["canonical_sha256"]),
        )


@dataclass(frozen=True, slots=True)
class _Cover:
    sheet: str
    min_row: int
    min_col: int
    max_row: int
    max_col: int
    exclude_rows: frozenset[int]
    exclude_cols: frozenset[int]
    series_id: str
    addresses: frozenset[str] | None = None

    def covers(self, sheet: str, row: int, col: int, address: str) -> bool:
        if self.addresses is not None:
            return address in self.addresses
        if sheet != self.sheet:
            return False
        if not (self.min_row <= row <= self.max_row and self.min_col <= col <= self.max_col):
            return False
        return row not in self.exclude_rows and col not in self.exclude_cols


def declared_domain(series: Mapping[str, Any]) -> dict[str, Any] | None:
    """Return the authored series-level or `input.domain` mapping."""
    domain = series.get("domain")
    if isinstance(domain, dict) and domain:
        return dict(domain)
    input_block = series.get("input")
    if isinstance(input_block, Mapping):
        nested = input_block.get("domain")
        if isinstance(nested, dict) and nested:
            return dict(nested)
    return None


def compile_domain_spec(series: dict[str, Any]) -> dict[str, Any] | None:
    """Return the domain spec compiled into the cell-type env, if any.

    `constant` implies `from_workbook`. `input` (leaf mode) falls back to
    `value_map` needles. `internal` / `output` / override contribute nothing
    unless `domain` is explicit.
    """
    declared = declared_domain(series)
    if declared is not None:
        return declared
    if has_constant_direction(series):
        return {"from_workbook": True}
    if has_input_direction(series) and input_mode(series) != "override":
        input_block = series.get("input")
        if isinstance(input_block, Mapping):
            value_map = input_block.get("value_map")
            if isinstance(value_map, Mapping) and value_map:
                return {"enum": list(value_map.values())}
    return None


def _input_domain_annotation(series: dict[str, Any]) -> Any | None:
    """Return the typing annotation equivalent to a series' compiled domain."""
    domain = compile_domain_spec(series)
    if domain is None or domain.get("from_workbook") is True:
        return None
    if "enum" in domain:
        return _RuntimeLiteral[tuple(domain["enum"])]
    if "between" in domain:
        bounds = domain["between"]
        return Annotated[int, Between(bounds.get("min"), bounds.get("max"))]
    if "real_between" in domain:
        bounds = domain["real_between"]
        return Annotated[float, RealBetween(bounds.get("min"), bounds.get("max"))]
    return None


def _python_type_for_series(series: Mapping[str, Any]) -> type:
    dtype = canonical_measure_dtype(series)
    if dtype in {"int", "integer"}:
        return int
    if dtype in {"string", "str"}:
        return str
    if dtype == "bool":
        return bool
    return float


def _with_relation_metadata(annotation: Any, relations: tuple[CellRelation, ...]) -> Any:
    if not relations:
        return annotation
    if get_origin(annotation) is Annotated:
        args = get_args(annotation)
        return Annotated[args[0], *args[1:], *relations]
    return Annotated[annotation, *relations]


def _cell_rc(address_or_coord: str, *, default_sheet: str | None = None) -> tuple[str, int, int]:
    if "!" in address_or_coord:
        sheet, coord = parse_address(address_or_coord)
    else:
        if default_sheet is None:
            raise ValueError(f"Cannot parse coordinate {address_or_coord!r} without a sheet")
        sheet = default_sheet
        coord = address_or_coord
    col_letters, row = coordinate_from_string(coord)
    return sheet, int(row), column_index_from_string(col_letters)


def _covers_for_series(
    series: Mapping[str, Any],
    *,
    workbook: Path | str | None,
    named_range_ranges: Mapping[str, tuple[str, str, str]] | None,
) -> list[_Cover]:
    series_id = str(series.get("id") or "")
    exclude_rows = frozenset(expand_row_specs(series.get("exclude_rows") or []))
    exclude_cols = frozenset(expand_column_specs(series.get("exclude_columns") or []))
    covers: list[_Cover] = []
    for data_range in series_data_ranges(series):
        resolved = data_range
        if "!" not in data_range and named_range_ranges and data_range in named_range_ranges:
            sheet, start, end = named_range_ranges[data_range]
            resolved = f"{sheet}!{start}:{end}"
        try:
            split = split_address_on_colon(resolved)
            if split is None:
                sheet, row, col = _cell_rc(resolved)
                covers.append(
                    _Cover(
                        sheet=sheet,
                        min_row=row,
                        min_col=col,
                        max_row=row,
                        max_col=col,
                        exclude_rows=exclude_rows,
                        exclude_cols=exclude_cols,
                        series_id=series_id,
                    )
                )
                continue
            start, end = split
            sheet, row1, col1 = _cell_rc(start)
            _, row2, col2 = _cell_rc(end, default_sheet=sheet)
            covers.append(
                _Cover(
                    sheet=sheet,
                    min_row=min(row1, row2),
                    min_col=min(col1, col2),
                    max_row=max(row1, row2),
                    max_col=max(col1, col2),
                    exclude_rows=exclude_rows,
                    exclude_cols=exclude_cols,
                    series_id=series_id,
                )
            )
        except (ValueError, KeyError):
            addresses = frozenset(
                normalize_cell_type_env_key(addr)
                for addr in expand_bound_series_addresses(series, workbook=workbook)
            )
            covers.append(
                _Cover(
                    sheet="",
                    min_row=0,
                    min_col=0,
                    max_row=0,
                    max_col=0,
                    exclude_rows=exclude_rows,
                    exclude_cols=exclude_cols,
                    series_id=series_id,
                    addresses=addresses,
                )
            )
            break
    return covers


def _relative_bindings_path(workbook: Path, bindings_path: Path) -> str:
    try:
        return str(bindings_path.resolve().relative_to(workbook.resolve().parent))
    except ValueError:
        return str(bindings_path.resolve())


class SeriesDomainIndex(Mapping[str, CellType]):
    """Lazy `Mapping` of bound addresses to `CellType`.

    Lookup scans per-sheet covering rectangles from the manifest (O(#series)
    per miss) and memoizes `CellType` per address. `from_workbook` values come
    from graph node values when attached, otherwise from a streaming workbook
    read. Truthiness does not compile those values. Iteration and `len`
    compile every bound address and prefetch each sheet's off-graph cells once.
    """

    def __init__(
        self,
        bindings: Mapping[str, Any],
        *,
        workbook: Path | str | None = None,
        graph: DependencyGraph | None = None,
        bindings_path: Path | str | None = None,
    ) -> None:
        raise_if_relation_errors(relation_declaration_issues(bindings))
        self._bindings = bindings
        self._indexed = series_by_id(bindings)
        self._workbook = None if workbook is None else Path(workbook)
        self._graph = graph
        self._reader: Any = None
        self._keyed: dict[str, dict[Any, str]] | None = None
        self._memo: dict[str, CellType] = {}
        self._expanded: dict[str, CellType] | None = None
        named_range_ranges = None if graph is None else graph.named_range_ranges
        self._covers: list[_Cover] = []
        needed: set[str] = set()
        for series_id, series in self._indexed.items():
            relations = iter_series_relations(series)
            if compile_domain_spec(series) is None and not relations:
                continue
            self._covers.extend(
                _covers_for_series(
                    series, workbook=self._workbook, named_range_ranges=named_range_ranges
                )
            )
            if relations:
                needed.add(series_id)
                needed.update(partner_id for _kind, partner_id in relations)
        if needed and self._workbook is not None:
            keyed, issues = relation_cell_indexes(
                bindings, workbook=self._workbook, series_ids=needed
            )
            issues.extend(missing_partner_cell_issues(bindings, keyed))
            raise_if_relation_errors(issues)
            self._keyed = keyed
        handle = None
        if self._workbook is not None and bindings_path is not None:
            handle = SeriesDomainHandle(
                workbook=str(self._workbook.resolve()),
                bindings=_relative_bindings_path(self._workbook, Path(bindings_path)),
                canonical_sha256=bindings_canonical_sha256(bindings),
            )
        self.handle = handle

    @classmethod
    def from_bindings(
        cls,
        bindings: Mapping[str, Any],
        *,
        workbook: Path | str | None = None,
        graph: DependencyGraph | None = None,
        bindings_path: Path | str | None = None,
    ) -> SeriesDomainIndex:
        """Build an index from a loaded (merged) binding manifest."""
        return cls(bindings, workbook=workbook, graph=graph, bindings_path=bindings_path)

    @classmethod
    def from_handle(
        cls,
        handle: SeriesDomainHandle,
        *,
        graph: DependencyGraph | None = None,
    ) -> SeriesDomainIndex | None:
        """Re-read the sidecar named by `handle`, or return None if it is stale."""
        from excel_grapher.series_bindings.load import load_series_bindings

        workbook = Path(handle.workbook)
        bindings_path = Path(handle.bindings)
        if not bindings_path.is_absolute():
            bindings_path = workbook.parent / bindings_path
        if not workbook.is_file() or not (bindings_path.is_file() or bindings_path.is_dir()):
            return None
        bindings = load_series_bindings(bindings_path)
        if bindings_canonical_sha256(bindings) != handle.canonical_sha256:
            return None
        return cls(bindings, workbook=workbook, graph=graph, bindings_path=bindings_path)

    def bind_graph(self, graph: DependencyGraph) -> SeriesDomainIndex:
        """Use `graph` node values for subsequent `from_workbook` lookups."""
        self._graph = graph
        return self

    def domain_for(self, key: str) -> CellType | None:
        """Return the compiled `CellType` covering `key`, if any."""
        norm = normalize_cell_type_env_key(key)
        cached = self._memo.get(norm)
        if cached is not None:
            return cached
        if self._expanded is not None:
            return self._expanded.get(norm)
        series = self._series_covering(norm)
        if series is None:
            return None
        cell_type = self._cell_type_for(series, norm)
        if cell_type is not None:
            self._memo[norm] = cell_type
        return cell_type

    def _series_covering(self, address: str) -> dict[str, Any] | None:
        try:
            sheet, row, col = _cell_rc(address)
        except (ValueError, KeyError):
            sheet, row, col = "", 0, 0
        for cover in self._covers:
            if cover.covers(sheet, row, col, address):
                series = self._indexed.get(cover.series_id)
                if series is not None:
                    return series
        return None

    def _value_for(self, address: str) -> object | None:
        if self._graph is not None:
            node = self._graph.get_node(address)
            if node is not None:
                return node.value
        if self._workbook is None:
            return None
        return self._workbook_reader().read(address)

    def _cell_type_for(self, series: dict[str, Any], address: str) -> CellType | None:
        domain = compile_domain_spec(series)
        relations = self._relations_for(series, address)
        if domain is not None and domain.get("from_workbook") is True:
            value = self._value_for(address)
            if value is None:
                if relations:
                    annotation = _with_relation_metadata(_python_type_for_series(series), relations)
                    return constraints_to_cell_type_env({address: annotation}, {})[address]
                return None
            env = constraints_to_cell_type_env(
                {address: _with_relation_metadata(_RuntimeLiteral[tuple([value])], relations)},
                {},
            )
            return env[address]
        base = _input_domain_annotation(series)
        if base is None and not relations:
            return None
        annotation = base if base is not None else _python_type_for_series(series)
        env = constraints_to_cell_type_env(
            {address: _with_relation_metadata(annotation, relations)},
            {},
        )
        return env[address]

    def _relations_for(self, series: Mapping[str, Any], address: str) -> tuple[CellRelation, ...]:
        relations = iter_series_relations(series)
        if not relations:
            return ()
        if self._keyed is None:
            return ()
        series_id = str(series.get("id") or "")
        declaring = self._keyed.get(series_id) or {}
        frozen = next((key for key, addr in declaring.items() if addr == address), None)
        if frozen is None:
            raise SeriesRelationError(
                f"series {series_id!r} has no bound cell at {address!r} for relation expansion"
            )
        meta: list[CellRelation] = []
        for kind, partner_id in relations:
            partner_addr = self._keyed[partner_id][frozen]
            meta.append(RELATION_TYPES[kind](partner_addr))
        return tuple(meta)

    def _workbook_reader(self) -> Any:
        if self._workbook is None:
            raise RuntimeError("from_workbook reads require a workbook path")
        if self._reader is None:
            from excel_grapher.series_bindings.resolve import _WorkbookValues

            self._reader = _WorkbookValues(self._workbook)
        return self._reader

    def _series_addresses(self) -> list[tuple[dict[str, Any], list[str]]]:
        compiled: list[tuple[dict[str, Any], list[str]]] = []
        seen_ids: set[str] = set()
        for cover in self._covers:
            if cover.series_id in seen_ids:
                continue
            seen_ids.add(cover.series_id)
            series = self._indexed[cover.series_id]
            compiled.append(
                (series, list(expand_bound_series_addresses(series, workbook=self._workbook)))
            )
        return compiled

    def _prefetch_off_graph(self, compiled: list[tuple[dict[str, Any], list[str]]]) -> None:
        """Read every off-graph `from_workbook` cell, one stream per sheet."""
        if self._workbook is None:
            return
        pending: list[str] = []
        for series, addresses in compiled:
            domain = compile_domain_spec(series)
            if domain is None or domain.get("from_workbook") is not True:
                continue
            for address in addresses:
                if self._graph is not None and address in self._graph:
                    continue
                pending.append(address)
        if pending:
            self._workbook_reader().prefetch(pending, graph=self._graph)

    def _materialize(self) -> dict[str, CellType]:
        if self._expanded is not None:
            return self._expanded
        compiled = self._series_addresses()
        self._prefetch_off_graph(compiled)
        env: dict[str, CellType] = {}
        for series, addresses in compiled:
            for address in addresses:
                norm = normalize_cell_type_env_key(address)
                cell_type = self._memo.get(norm) or self._cell_type_for(series, norm)
                if cell_type is not None:
                    self._memo[norm] = cell_type
                    env[norm] = cell_type
        self._expanded = env
        return env

    def __bool__(self) -> bool:
        """Return whether any domain is declared, without reading workbook values.

        Before the index is expanded, this is whether any series contributes a
        covering domain. After `len` or iteration, it matches the compiled mapping.
        """
        if self._expanded is not None:
            return bool(self._expanded)
        return bool(self._covers)

    def __getitem__(self, key: str) -> CellType:
        cell_type = self.domain_for(str(key))
        if cell_type is None:
            raise KeyError(key)
        return cell_type

    def __iter__(self) -> Iterator[str]:
        return iter(self._materialize())

    def __len__(self) -> int:
        return len(self._materialize())

    def __contains__(self, key: object) -> bool:
        if not isinstance(key, str):
            return False
        return self.domain_for(key) is not None


def cell_type_env_from_bindings(
    bindings: Mapping[str, Any], *, workbook: Path | str
) -> SeriesDomainIndex:
    """Compile a binding manifest into a lazy `CellTypeEnv`.

    Args:
        bindings: Loaded (merged) series binding manifest.
        workbook: Workbook path; read for `data_range` expansion, key binds,
            and `from_workbook` cached values.

    Returns:
        A `SeriesDomainIndex` keyed like `constraints_to_cell_type_env` output.
        Address lookup compiles one cell; iteration and `len` expand the whole
        domain, prefetching each sheet once. Callers that need a fully expanded
        dict can write `dict(index)`.

    Raises:
        SeriesRelationError: A relation partner is missing, incomparable,
            cyclic, reflexive, or has no cell at the declaring key.
    """
    return SeriesDomainIndex.from_bindings(bindings, workbook=workbook)


def undomained_leaves(
    graph: DependencyGraph,
    bindings: Mapping[str, Any],
    *,
    workbook: Path | str | None = None,
) -> list[str]:
    """Return sorted graph leaves that no binding domain covers.

    Args:
        graph: Extracted dependency graph.
        bindings: Loaded (merged) series binding manifest.
        workbook: Workbook path used when `graph.domains` is unset.

    Returns:
        Leaf addresses with no compiled domain.
    """
    index = graph.domains
    if not isinstance(index, SeriesDomainIndex):
        index = SeriesDomainIndex.from_bindings(bindings, workbook=workbook, graph=graph)
    return [key for key in graph.leaf_keys() if key not in index]
