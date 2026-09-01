"""Mechanical synthesis of parameterized cluster helper bodies (issue #45).

Consumes a cluster's fingerprint summary (structural skeletons, per-ref
relations, per-member ref addresses/keys) plus the unpacked exemplar
translations, and produces a parameterized helper body without any LLM
involvement. The synthesizer only rewrites *read sites* — ``xl_cell`` /
``xl_eval`` literals, ``read_*`` accessor arguments, semantic-helper call
arguments, in-cluster self-recurrence, and the ref-info tuple of the constant
range INDEX/MATCH family (issue #170) — leaving every operator, coercion
wrapper, literal, and lazy branch of the exemplar untouched. Cluster-constant
``xl_range`` reads pass through verbatim; ``xl_match`` is pure and never
rewritten.

Every rewrite is verified per member: the derived address or argument keys
must equal the recorded refs for all members of the fingerprint group. Any
unsupported shape or verification mismatch raises
:class:`MechanicalSynthesisError`; callers fall back to the legacy full-body
LLM contract.
"""

from __future__ import annotations

import ast
import re
from collections.abc import Callable, Iterator, Mapping, Sequence
from dataclasses import dataclass, field

from fastpyxl.utils.cell import column_index_from_string

from excel_grapher.exporter.pass1.bindings import BindingKeyValue, KeyConceptSpec
from excel_grapher.exporter.pass1.empty_if_rewrite import rewrite_empty_if_none_literals
from excel_grapher.exporter.pass1.fingerprints import (
    ClusterFingerprintSummary,
    FingerprintGroup,
    LookupKey,
    RefRelation,
    _subset_routing_lookup,
)

# Bump when ``synthesize_*_body`` / mechanical rewrite semantics change.
MECHANICAL_BODY_SCHEMA_VERSION = "1.0.0"

_MECHANICAL_TEMP_PATTERN = re.compile(r"^_t\d+$")
_ACCESSOR_PREFIX = "read_"
_CELL_FUNCTION_PREFIX = "cell_"
_POINT_READ_CALLEES = frozenset({"xl_cell", "xl_eval"})
# xl_range_rows is a public boundary handler and never appears in cluster
# bodies; xl_match is pure (consumes an already-read Range) and needs no
# rewrite. xl_range / xl_offset / xl_index_ref are handled shape-by-shape in
# _collect_read_sites (issue #170).
_UNSUPPORTED_READ_CALLEES = frozenset({"xl_range_rows"})


class MechanicalSynthesisError(Exception):
    """A cluster cannot be mechanically synthesized; fall back to the LLM body."""

    def __init__(self, reason: str) -> None:
        super().__init__(reason)
        self.reason = reason


@dataclass(frozen=True)
class MechanicalBodyDraft:
    """A verified, parameterized helper body with mechanical local names."""

    body: str
    renameable_locals: tuple[str, ...]
    lookup_table_names: tuple[str, ...]
    group_count: int


def _is_literal_expr(node: ast.expr) -> bool:
    if isinstance(node, ast.Constant):
        return node.value is None or isinstance(node.value, (str, int, float, bool))
    if isinstance(node, ast.Tuple):
        return all(_is_literal_expr(element) for element in node.elts)
    return False


def parse_inlinable_wrapper(function_def: ast.FunctionDef) -> str | None:
    """Return the replacement expression for a ``(ctx)``-only thin wrapper.

    A wrapper is inlinable when its executable body is a single ``return`` of a
    call that forwards the same ``ctx`` and otherwise passes only literals —
    e.g. ``return shock_active(ctx, time_period=1)`` or
    ``return xl_cell(ctx, 'Inputs!B5')``. Substituting that call for a
    ``wrapper(ctx)`` call site is exact: no locals, defaults, or evaluation
    order are involved.
    """
    args = function_def.args
    if (
        [arg.arg for arg in args.args] != ["ctx"]
        or args.posonlyargs
        or args.kwonlyargs
        or args.vararg is not None
        or args.kwarg is not None
        or args.defaults
    ):
        return None
    body = function_def.body
    if (
        body
        and isinstance(body[0], ast.Expr)
        and isinstance(body[0].value, ast.Constant)
        and isinstance(body[0].value.value, str)
    ):
        body = body[1:]
    if len(body) != 1:
        return None
    statement = body[0]
    if not isinstance(statement, ast.Return) or statement.value is None:
        return None
    call = statement.value
    if not (
        isinstance(call, ast.Call)
        and isinstance(call.func, ast.Name)
        and not call.func.id.startswith(_CELL_FUNCTION_PREFIX)
        and call.args
        and isinstance(call.args[0], ast.Name)
        and call.args[0].id == "ctx"
    ):
        return None
    if not all(_is_literal_expr(arg) for arg in call.args[1:]):
        return None
    for keyword in call.keywords:
        if keyword.arg is None or not _is_literal_expr(keyword.value):
            return None
    return ast.unparse(call)


def synthesize_singleton_body(
    python_source: str,
    *,
    inline_replacements: Mapping[str, str],
) -> MechanicalBodyDraft:
    """Synthesize a draft body for a singleton refactor unit.

    The unpacked translation is already statement-shaped, so the only rewrite
    is call-graph rewiring: every ``cell_*(ctx)`` dependency call is replaced
    with its thin wrapper's semantic call (from ``inline_replacements``, keyed
    by wrapper function name). Everything else — coercion wrappers, laziness,
    literals, runtime reads — is preserved verbatim.

    Raises :class:`MechanicalSynthesisError` on any ``cell_*`` reference that
    cannot be provably inlined; callers fall back to the full-body contract.
    """
    module = ast.parse(python_source)
    functions = [node for node in module.body if isinstance(node, ast.FunctionDef)]
    if len(functions) != 1:
        raise MechanicalSynthesisError("unsupported_translation_shape")
    function_def = functions[0]
    if [arg.arg for arg in function_def.args.args] != ["ctx"]:
        raise MechanicalSynthesisError("unsupported_translation_signature")
    statements = list(function_def.body)
    if (
        statements
        and isinstance(statements[0], ast.Expr)
        and isinstance(statements[0].value, ast.Constant)
        and isinstance(statements[0].value.value, str)
    ):
        statements = statements[1:]

    replacements: dict[int, ast.expr] = {}
    inlined_call_funcs: set[int] = set()
    for statement in statements:
        for call in _iter_calls(statement):
            func = call.func
            if not isinstance(func, ast.Name) or not func.id.startswith(_CELL_FUNCTION_PREFIX):
                continue
            if (
                len(call.args) != 1
                or not isinstance(call.args[0], ast.Name)
                or call.args[0].id != "ctx"
                or call.keywords
            ):
                raise MechanicalSynthesisError(f"cell_call_shape_unsupported:{func.id}")
            replacement_source = inline_replacements.get(func.id)
            if replacement_source is None:
                raise MechanicalSynthesisError(f"cell_dependency_not_inlinable:{func.id}")
            replacements[id(call)] = ast.parse(replacement_source, mode="eval").body
            inlined_call_funcs.add(id(func))
    for statement in statements:
        for node in ast.walk(statement):
            if (
                isinstance(node, ast.Name)
                and node.id.startswith(_CELL_FUNCTION_PREFIX)
                and id(node) not in inlined_call_funcs
            ):
                raise MechanicalSynthesisError(f"cell_reference_unsupported:{node.id}")

    rewritten = [_replace_nodes(statement, replacements) for statement in statements]
    renameable = sorted(
        {
            node.id
            for statement in rewritten
            for node in ast.walk(statement)
            if isinstance(node, ast.Name) and _MECHANICAL_TEMP_PATTERN.match(node.id)
        }
    )
    body = "\n".join(ast.unparse(statement) for statement in rewritten)
    body = rewrite_empty_if_none_literals(body)
    indented = "\n".join(f"    {line}" for line in body.splitlines())
    try:
        ast.parse(f"def _draft(ctx):\n{indented}\n")
    except SyntaxError as error:  # pragma: no cover - defensive
        raise MechanicalSynthesisError(f"draft_body_invalid:{error}") from error
    return MechanicalBodyDraft(
        body=body,
        renameable_locals=tuple(renameable),
        lookup_table_names=(),
        group_count=1,
    )


def _sort_key(value: LookupKey) -> tuple[int, object]:
    if isinstance(value, tuple):
        return (4, tuple(_sort_key(item) for item in value))
    if isinstance(value, bool):
        return (3, value)
    if isinstance(value, (int, float)):
        return (0, value)
    if isinstance(value, str):
        return (1, value)
    return (2, str(value))


def _table_key_expr(key: LookupKey) -> ast.expr:
    if isinstance(key, tuple):
        return ast.Tuple(elts=[ast.Constant(value=item) for item in key], ctx=ast.Load())
    return ast.Constant(value=key)


@dataclass
class _TableRegistry:
    """Allocates deterministic names for mechanical lookup-table dict literals."""

    tables: dict[str, dict[LookupKey, BindingKeyValue]] = field(default_factory=dict)
    order: list[str] = field(default_factory=list)

    def register(self, base_name: str, content: Mapping[LookupKey, BindingKeyValue]) -> str:
        frozen = dict(content)
        name = base_name
        suffix = 2
        while name in self.tables:
            if self.tables[name] == frozen:
                return name
            name = f"{base_name}_{suffix}"
            suffix += 1
        self.tables[name] = frozen
        self.order.append(name)
        return name

    def assignments(self, names: Sequence[str]) -> list[ast.stmt]:
        statements: list[ast.stmt] = []
        for name in names:
            content = self.tables[name]
            keys = sorted(content, key=_sort_key)
            statements.append(
                ast.Assign(
                    targets=[ast.Name(id=name, ctx=ast.Store())],
                    value=ast.Dict(
                        keys=[_table_key_expr(key) for key in keys],
                        values=[ast.Constant(value=content[key]) for key in keys],
                    ),
                )
            )
        return statements


@dataclass(frozen=True)
class _IndexRefInfo:
    """The recognized ``xl_offset(ctx, xl_index_ref((...), _tN, col), 0, 0)`` shape."""

    index_ref_call: ast.Call
    tuple_node: ast.Tuple
    sheet: str
    corners: tuple[int, int, int, int]
    """The ref-info tuple's ``(row_start, col_start, row_end, col_end)`` literals."""


@dataclass(frozen=True)
class _ReadSite:
    node: ast.Call
    callee: str
    address: str | None
    endpoints: tuple[str, str] | None = None
    """Recorded-format endpoint addresses of a literal ``xl_range`` read."""
    index_ref: _IndexRefInfo | None = None


def _iter_calls(node: ast.AST) -> Iterator[ast.Call]:
    if isinstance(node, ast.Call):
        yield node
    for child in ast.iter_child_nodes(node):
        yield from _iter_calls(child)


def _call_address_literal(node: ast.Call) -> str | None:
    if len(node.args) >= 2 and isinstance(node.args[1], ast.Constant):
        value = node.args[1].value
        if isinstance(value, str):
            return value
    return None


def _sheet_prefix(address: str) -> str | None:
    """Return the ``Sheet!`` prefix of a sheet-qualified address, quotes kept."""
    if address.startswith("'"):
        i = 1
        while i < len(address):
            if address[i] == "'":
                if address[i + 1 : i + 2] == "'":
                    i += 2
                    continue
                break
            i += 1
        if address[i + 1 : i + 2] != "!":
            return None
        return address[: i + 2]
    bang = address.rfind("!")
    if bang < 0:
        return None
    return address[: bang + 1]


def _range_literal_endpoints(address: str) -> tuple[str, str] | None:
    """Split a literal range address into recorded-format endpoint addresses."""
    if ":" not in address:
        return None
    start_text, end_text = address.split(":", 1)
    prefix = _sheet_prefix(start_text)
    if prefix is None:
        return None
    if _sheet_prefix(end_text) is not None:
        return start_text, end_text
    return start_text, prefix + end_text


def _parse_endpoint_address(address: str) -> tuple[str, int, int] | None:
    """Parse a recorded ref address into ``(sheet, row, column_index)``."""
    if "!" not in address:
        return None
    sheet, colrow = address.split("!", 1)
    if sheet.startswith("'") and sheet.endswith("'") and len(sheet) >= 2:
        sheet = sheet[1:-1].replace("''", "'")
    column = "".join(character for character in colrow if character.isalpha())
    digits = "".join(character for character in colrow if character.isdigit())
    if not column or not digits:
        return None
    try:
        column_index = column_index_from_string(column)
    except ValueError:
        return None
    return sheet, int(digits), column_index


def _int_corner_literal(node: ast.expr) -> int | None:
    if (
        isinstance(node, ast.Constant)
        and not isinstance(node.value, bool)
        and isinstance(node.value, int)
    ):
        return node.value
    return None


def _is_zero_literal(node: ast.expr) -> bool:
    return (
        isinstance(node, ast.Constant)
        and not isinstance(node.value, bool)
        and isinstance(node.value, (int, float))
        and node.value == 0
    )


def _parse_index_offset_call(call: ast.Call) -> _IndexRefInfo | None:
    """Recognize exactly the constant INDEX/MATCH reference shape (issue #170).

    ``xl_offset(ctx, xl_index_ref((sheet, r1, c1, r2, c2), _tN, col), 0.0, 0.0)``
    with literal zero offsets, no height/width, a literal ref-info tuple, a name
    row argument, and a literal column argument. Any other shape returns None.
    """
    if call.keywords or len(call.args) != 4:
        return None
    if not (isinstance(call.args[0], ast.Name) and call.args[0].id == "ctx"):
        return None
    if not (_is_zero_literal(call.args[2]) and _is_zero_literal(call.args[3])):
        return None
    inner = call.args[1]
    if not (
        isinstance(inner, ast.Call)
        and isinstance(inner.func, ast.Name)
        and inner.func.id == "xl_index_ref"
        and not inner.keywords
        and len(inner.args) == 3
    ):
        return None
    tuple_node, row_arg, col_arg = inner.args
    if not isinstance(row_arg, ast.Name) or row_arg.id == "ctx":
        return None
    if not (
        isinstance(col_arg, ast.Constant)
        and not isinstance(col_arg.value, bool)
        and isinstance(col_arg.value, (int, float))
    ):
        return None
    if not (isinstance(tuple_node, ast.Tuple) and len(tuple_node.elts) == 5):
        return None
    sheet_node = tuple_node.elts[0]
    if not (isinstance(sheet_node, ast.Constant) and isinstance(sheet_node.value, str)):
        return None
    row_start, col_start, row_end, col_end = (
        _int_corner_literal(element) for element in tuple_node.elts[1:]
    )
    if row_start is None or col_start is None or row_end is None or col_end is None:
        return None
    return _IndexRefInfo(
        index_ref_call=inner,
        tuple_node=tuple_node,
        sheet=sheet_node.value,
        corners=(row_start, col_start, row_end, col_end),
    )


def _collect_read_sites(
    statements: Sequence[ast.stmt],
    semantic_helper_names: frozenset[str],
) -> list[_ReadSite]:
    sites: list[_ReadSite] = []
    consumed_index_refs: set[int] = set()
    for statement in statements:
        for call in _iter_calls(statement):
            func = call.func
            if not isinstance(func, ast.Name):
                continue
            callee = func.id
            if callee in _UNSUPPORTED_READ_CALLEES:
                raise MechanicalSynthesisError(f"unsupported_read_callee:{callee}")
            if callee == "xl_range":
                address = _call_address_literal(call)
                if address is None:
                    raise MechanicalSynthesisError("non_literal_range_address")
                endpoints = _range_literal_endpoints(address)
                if endpoints is None:
                    raise MechanicalSynthesisError(f"unparseable_range_address:{address}")
                sites.append(
                    _ReadSite(node=call, callee=callee, address=address, endpoints=endpoints)
                )
            elif callee == "xl_offset":
                info = _parse_index_offset_call(call)
                if info is None:
                    raise MechanicalSynthesisError("unsupported_offset_shape")
                # _iter_calls walks outer-first, so the inner xl_index_ref is
                # visited after this call and must not be re-collected.
                consumed_index_refs.add(id(info.index_ref_call))
                sites.append(_ReadSite(node=call, callee=callee, address=None, index_ref=info))
            elif callee == "xl_index_ref":
                if id(call) not in consumed_index_refs:
                    raise MechanicalSynthesisError("unsupported_index_ref_shape")
            elif callee in _POINT_READ_CALLEES:
                sites.append(
                    _ReadSite(node=call, callee=callee, address=_call_address_literal(call))
                )
            elif (
                callee.startswith((_ACCESSOR_PREFIX, _CELL_FUNCTION_PREFIX))
                or callee in semantic_helper_names
            ):
                sites.append(_ReadSite(node=call, callee=callee, address=None))
    return sites


def _literal_keywords(call: ast.Call) -> dict[str, BindingKeyValue] | None:
    """Return ``{kwarg: literal}`` when every keyword is a plain literal."""
    literals: dict[str, BindingKeyValue] = {}
    for keyword in call.keywords:
        if keyword.arg is None or not isinstance(keyword.value, ast.Constant):
            return None
        value = keyword.value.value
        if isinstance(value, bytes) or not isinstance(value, (str, int, float, bool)):
            return None
        literals[keyword.arg] = value
    return literals


def _param_expr(name: str) -> ast.expr:
    return ast.Name(id=name, ctx=ast.Load())


def _offset_expr(param: str, delta: int) -> ast.expr:
    if delta < 0:
        return ast.BinOp(left=_param_expr(param), op=ast.Sub(), right=ast.Constant(value=-delta))
    return ast.BinOp(left=_param_expr(param), op=ast.Add(), right=ast.Constant(value=delta))


@dataclass
class _GroupSynthesizer:
    group: FingerprintGroup
    group_index: int
    multi_group: bool
    helper_name: str
    param_by_dim: dict[str, str]
    varying_dims: tuple[str, ...]
    expected_member_keys: Mapping[str, Mapping[str, BindingKeyValue]]
    tables: _TableRegistry

    def __post_init__(self) -> None:
        self.ref_addresses = dict(self.group.ref_addresses_by_member)
        self.ref_keys = dict(self.group.ref_keys_by_member)
        exemplar = self.group.exemplar
        self.exemplar_refs = self.ref_addresses.get(exemplar.address, ())
        self.exemplar_ref_keys = self.ref_keys.get(exemplar.address, ())
        self.table_names: list[str] = []
        self.replacements: dict[int, ast.expr] = {}
        self.claim_counts: dict[int, int] = {}
        self.index_verified_slots: set[int] = set()
        self.template_claimed_slots: set[int] = set()

    # -- relation-derived key expressions and their mirror evaluation --------

    def _dim_param(self, dim: str) -> str:
        param = self.param_by_dim.get(dim)
        if param is None:
            raise MechanicalSynthesisError(f"dimension_without_parameter:{dim}")
        return param

    def _register_table(self, base_name: str, content: Mapping[LookupKey, BindingKeyValue]) -> str:
        name = self.tables.register(base_name, content)
        if name not in self.table_names:
            self.table_names.append(name)
        return name

    def _derived_key_expr(self, relation: RefRelation, dim: str) -> ast.expr:
        if dim in relation.fixed_keys:
            return ast.Constant(value=relation.fixed_keys[dim])
        if dim in relation.identity_dims:
            return _param_expr(self._dim_param(dim))
        if dim in relation.offsets:
            return _offset_expr(self._dim_param(dim), relation.offsets[dim])
        if dim in relation.lookups:
            key_dim = relation.lookup_keys.get(dim)
            if key_dim is None:
                raise MechanicalSynthesisError(f"lookup_without_key_dim:{dim}")
            table = relation.lookups[dim]
            if isinstance(key_dim, tuple):
                key_params = [self._dim_param(key) for key in key_dim]
                table_name = self._register_table(f"{dim.lower()}_by_{'_'.join(key_params)}", table)
                return ast.Subscript(
                    value=_param_expr(table_name),
                    slice=ast.Tuple(
                        elts=[_param_expr(param) for param in key_params],
                        ctx=ast.Load(),
                    ),
                    ctx=ast.Load(),
                )
            key_param = self._dim_param(key_dim)
            if dim in relation.lookup_bases:
                table_name = self._register_table(
                    f"{self._dim_param(dim)}_lag_by_{key_param}", table
                )
                return ast.BinOp(
                    left=_param_expr(self._dim_param(dim)),
                    op=ast.Sub(),
                    right=ast.Subscript(
                        value=_param_expr(table_name),
                        slice=_param_expr(key_param),
                        ctx=ast.Load(),
                    ),
                )
            table_name = self._register_table(f"{dim.lower()}_by_{key_param}", table)
            return ast.Subscript(
                value=_param_expr(table_name),
                slice=_param_expr(key_param),
                ctx=ast.Load(),
            )
        raise MechanicalSynthesisError(f"underived_dimension:{dim}")

    def _derived_key_value(
        self,
        relation: RefRelation,
        dim: str,
        member_keys: Mapping[str, BindingKeyValue],
    ) -> BindingKeyValue:
        if dim in relation.fixed_keys:
            return relation.fixed_keys[dim]
        if dim in relation.identity_dims:
            return member_keys[dim]
        if dim in relation.offsets:
            value = member_keys[dim]
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                raise MechanicalSynthesisError(f"offset_on_non_numeric_key:{dim}")
            return value + relation.offsets[dim]
        if dim in relation.lookups:
            key_dim = relation.lookup_keys.get(dim)
            if key_dim is None:
                raise MechanicalSynthesisError(f"lookup_without_key_dim:{dim}")
            table = relation.lookups[dim]
            if isinstance(key_dim, tuple):
                if any(key not in member_keys for key in key_dim):
                    raise MechanicalSynthesisError(f"lookup_without_key_dim:{dim}")
                tuple_key: LookupKey = tuple(member_keys[key] for key in key_dim)
                if tuple_key not in table:
                    raise MechanicalSynthesisError(f"lookup_key_not_covered:{dim}")
                return table[tuple_key]
            if key_dim not in member_keys:
                raise MechanicalSynthesisError(f"lookup_without_key_dim:{dim}")
            key_value = member_keys[key_dim]
            if key_value not in table:
                raise MechanicalSynthesisError(f"lookup_key_not_covered:{dim}")
            if dim in relation.lookup_bases:
                base = member_keys[dim]
                lag = table[key_value]
                if (
                    isinstance(base, bool)
                    or isinstance(lag, bool)
                    or not isinstance(base, (int, float))
                    or not isinstance(lag, (int, float))
                ):
                    raise MechanicalSynthesisError(f"lag_on_non_numeric_key:{dim}")
                return base - lag
            return table[key_value]
        raise MechanicalSynthesisError(f"underived_dimension:{dim}")

    # -- slot matching --------------------------------------------------------

    def _keywords_match_slot(self, site: _ReadSite, slot: int) -> bool:
        literals = _literal_keywords(site.node)
        if literals is None:
            return False
        slot_keys = self.exemplar_ref_keys[slot]
        for kwarg, value in literals.items():
            dim = self._dim_for_kwarg(kwarg, slot_keys)
            if dim is None or slot_keys.get(dim) != value:
                return False
        return True

    def _dim_for_kwarg(self, kwarg: str, slot_keys: Mapping[str, BindingKeyValue]) -> str | None:
        for dim, param in self.param_by_dim.items():
            if param == kwarg:
                return dim
        for dim in slot_keys:
            if dim.lower() == kwarg:
                return dim
        return None

    def _site_matches_slot(self, site: _ReadSite, slot: int) -> bool:
        relation = self.group.ref_relations[slot]
        if slot >= len(self.exemplar_refs):
            return False
        slot_address = self.exemplar_refs[slot]
        if site.callee in _POINT_READ_CALLEES:
            return site.address == slot_address
        if (
            relation.resolution.kind == "semantic_helper"
            and site.callee == relation.resolution.helper_name
        ):
            return self._keywords_match_slot(site, slot)
        if (
            relation.series_id is not None
            and site.callee == f"{_ACCESSOR_PREFIX}{relation.series_id}"
        ):
            return self._keywords_match_slot(site, slot)
        return False

    # -- rewrites -------------------------------------------------------------

    def _rewrite_keyword_call(self, site: _ReadSite, slot: int) -> ast.expr | None:
        relation = self.group.ref_relations[slot]
        slot_keys = self.exemplar_ref_keys[slot]
        new_keywords: list[ast.keyword] = []
        changed = False
        for keyword in site.node.keywords:
            assert keyword.arg is not None  # _keywords_match_slot enforced literals
            dim = self._dim_for_kwarg(keyword.arg, slot_keys)
            assert dim is not None
            if dim in relation.fixed_keys:
                new_keywords.append(keyword)
                continue
            new_keywords.append(
                ast.keyword(arg=keyword.arg, value=self._derived_key_expr(relation, dim))
            )
            changed = True
        if not changed:
            return None
        return ast.Call(func=site.node.func, args=site.node.args, keywords=new_keywords)

    def _address_template_expr(self, slot: int) -> ast.expr | None:
        """Build the parameterized address expression for an ``xl_cell`` slot."""
        relation = self.group.ref_relations[slot]
        resolution = relation.resolution
        template = resolution.address_template
        if template is None:
            raise MechanicalSynthesisError(f"missing_address_template:slot_{slot}")
        needs_col = "{col}" in template
        needs_row = "{row}" in template
        if not needs_col and not needs_row:
            self._verify_constant_slot_address(slot, template)
            return None

        parts = re.split(r"(\{col\}|\{row\})", template)
        joined: list[ast.expr] = []
        for part in parts:
            if part == "{col}":
                joined.append(self._axis_lookup_expr(slot, resolution.col_by_dim, "column"))
            elif part == "{row}":
                joined.append(self._axis_lookup_expr(slot, resolution.row_by_dim, "row"))
            elif part:
                joined.append(ast.Constant(value=part))
        formatted: list[ast.expr] = [
            expr
            if isinstance(expr, ast.Constant)
            else ast.FormattedValue(value=expr, conversion=-1, format_spec=None)
            for expr in joined
        ]
        return ast.JoinedStr(values=formatted)

    def _axis_lookup_expr(
        self,
        slot: int,
        axis_tables: Sequence[tuple[str, tuple[tuple[BindingKeyValue, BindingKeyValue], ...]]],
        axis: str,
    ) -> ast.expr:
        for dim, pairs in axis_tables:
            if dim not in self.param_by_dim or dim not in self.varying_dims:
                continue
            param = self.param_by_dim[dim]
            table_name = self._register_table(f"{axis}_by_{param}", dict(pairs))
            return ast.Subscript(
                value=_param_expr(table_name),
                slice=_param_expr(param),
                ctx=ast.Load(),
            )
        raise MechanicalSynthesisError(f"no_routable_{axis}_dimension:slot_{slot}")

    def _self_recurrence_call(self, slot: int) -> ast.expr:
        relation = self.group.ref_relations[slot]
        keywords = [
            ast.keyword(arg=self.param_by_dim[dim], value=self._derived_key_expr(relation, dim))
            for dim in self.varying_dims
        ]
        return ast.Call(
            func=ast.Name(id=self.helper_name, ctx=ast.Load()),
            args=[ast.Name(id="ctx", ctx=ast.Load())],
            keywords=keywords,
        )

    # -- verification ----------------------------------------------------------

    def _verify_constant_slot_address(self, slot: int, address: str) -> None:
        for member in self.group.members:
            recorded = self.ref_addresses[member][slot]
            if recorded != address:
                raise MechanicalSynthesisError(
                    f"constant_slot_address_mismatch:slot_{slot}:{member}"
                )

    def _verify_templated_slot_address(self, slot: int) -> None:
        relation = self.group.ref_relations[slot]
        resolution = relation.resolution
        template = resolution.address_template
        assert template is not None
        col_tables = {dim: dict(pairs) for dim, pairs in resolution.col_by_dim}
        row_tables = {dim: dict(pairs) for dim, pairs in resolution.row_by_dim}
        col_dim = self._chosen_axis_dim(resolution.col_by_dim)
        row_dim = self._chosen_axis_dim(resolution.row_by_dim)
        for member in self.group.members:
            member_keys = self.expected_member_keys.get(member, {})
            resolved = template
            if "{col}" in template:
                assert col_dim is not None
                key = member_keys.get(col_dim)
                if key not in col_tables[col_dim]:
                    raise MechanicalSynthesisError(f"column_key_not_covered:slot_{slot}:{member}")
                resolved = resolved.replace("{col}", str(col_tables[col_dim][key]))
            if "{row}" in template:
                assert row_dim is not None
                key = member_keys.get(row_dim)
                if key not in row_tables[row_dim]:
                    raise MechanicalSynthesisError(f"row_key_not_covered:slot_{slot}:{member}")
                resolved = resolved.replace("{row}", str(row_tables[row_dim][key]))
            if resolved != self.ref_addresses[member][slot]:
                raise MechanicalSynthesisError(f"templated_address_mismatch:slot_{slot}:{member}")

    def _chosen_axis_dim(
        self,
        axis_tables: Sequence[tuple[str, tuple[tuple[BindingKeyValue, BindingKeyValue], ...]]],
    ) -> str | None:
        for dim, _pairs in axis_tables:
            if dim in self.param_by_dim and dim in self.varying_dims:
                return dim
        return None

    def _verify_derived_keys(self, slot: int) -> None:
        relation = self.group.ref_relations[slot]
        for member in self.group.members:
            member_keys = self.expected_member_keys.get(member, {})
            recorded = self.ref_keys[member][slot]
            for dim, expected_value in recorded.items():
                derived = self._derived_key_value(relation, dim, member_keys)
                if derived != expected_value:
                    raise MechanicalSynthesisError(
                        f"derived_key_mismatch:slot_{slot}:{member}:{dim}"
                    )

    def _verify_self_recurrence(self, slot: int) -> None:
        relation = self.group.ref_relations[slot]
        for member in self.group.members:
            member_keys = self.expected_member_keys.get(member, {})
            target = self.ref_addresses[member][slot]
            target_keys = self.expected_member_keys.get(target)
            if target_keys is None:
                raise MechanicalSynthesisError(
                    f"self_recurrence_target_outside_cluster:slot_{slot}:{member}"
                )
            for dim in self.varying_dims:
                derived = self._derived_key_value(relation, dim, member_keys)
                if derived != target_keys.get(dim):
                    raise MechanicalSynthesisError(
                        f"self_recurrence_key_mismatch:slot_{slot}:{member}:{dim}"
                    )

    # -- orchestration ---------------------------------------------------------

    def synthesize(self) -> tuple[list[ast.stmt], list[str]]:
        """Return the rewritten group statements and its renameable local names."""
        exemplar = self.group.exemplar
        module = ast.parse(exemplar.python_source)
        functions = [node for node in module.body if isinstance(node, ast.FunctionDef)]
        if len(functions) != 1:
            raise MechanicalSynthesisError("unsupported_exemplar_shape")
        statements = list(functions[0].body)
        if (
            statements
            and isinstance(statements[0], ast.Expr)
            and isinstance(statements[0].value, ast.Constant)
            and isinstance(statements[0].value.value, str)
        ):
            statements = statements[1:]
        if not self.exemplar_refs and self.group.ref_relations:
            raise MechanicalSynthesisError("missing_exemplar_ref_addresses")
        if len(self.exemplar_ref_keys) != len(self.group.ref_relations):
            raise MechanicalSynthesisError("missing_exemplar_ref_keys")

        semantic_helper_names = frozenset(
            relation.resolution.helper_name
            for relation in self.group.ref_relations
            if relation.resolution.kind == "semantic_helper"
            and relation.resolution.helper_name is not None
        )
        sites = _collect_read_sites(statements, semantic_helper_names)
        claimed: dict[int, int] = {}
        for site in sites:
            if site.endpoints is not None:
                self._claim_range_site(site)
                continue
            if site.index_ref is not None:
                self._claim_index_site(site)
                continue
            matching = [
                slot
                for slot in range(len(self.group.ref_relations))
                if self._site_matches_slot(site, slot)
            ]
            if not matching:
                raise MechanicalSynthesisError(f"unclaimed_read_site:{ast.unparse(site.node)}")
            slot = min(matching, key=lambda s: (self.claim_counts.get(s, 0), s))
            self.claim_counts[slot] = self.claim_counts.get(slot, 0) + 1
            self.template_claimed_slots.add(slot)
            claimed[id(site.node)] = slot
            self._rewrite_site(site, slot)

        unclaimed_slots = [
            slot for slot in range(len(self.group.ref_relations)) if slot not in self.claim_counts
        ]
        if unclaimed_slots:
            raise MechanicalSynthesisError(f"slots_without_read_sites:{unclaimed_slots}")

        for slot in range(len(self.group.ref_relations)):
            self._verify_slot(slot)

        rewritten = [_replace_nodes(statement, self.replacements) for statement in statements]
        rewritten, locals_renamed = self._prefix_group_temporaries(rewritten)
        prelude = self.tables.assignments(self.table_names)
        return prelude + rewritten, locals_renamed

    def _rewrite_site(self, site: _ReadSite, slot: int) -> None:
        relation = self.group.ref_relations[slot]
        resolution = relation.resolution
        if site.callee == "xl_cell":
            address_expr = self._address_template_expr(slot)
            if address_expr is not None:
                self.replacements[id(site.node)] = ast.Call(
                    func=site.node.func,
                    args=[site.node.args[0], address_expr],
                    keywords=site.node.keywords,
                )
            return
        if site.callee == "xl_eval":
            if resolution.kind == "self_recurrence":
                self.replacements[id(site.node)] = self._self_recurrence_call(slot)
                return
            if resolution.kind == "xl_cell":
                address_expr = self._address_template_expr(slot)
                if address_expr is not None:
                    # Varying dependency addresses already carry an
                    # address_template from fingerprint resolution. The
                    # exemplar's cell_* callback is member-specific and cannot
                    # be kept, so evaluate through xl_cell (resolver) with the
                    # templated address — verified against recorded ref
                    # addresses like xl_cell sites.
                    self.replacements[id(site.node)] = ast.Call(
                        func=ast.Name(id="xl_cell", ctx=ast.Load()),
                        args=[site.node.args[0], address_expr],
                        keywords=[],
                    )
                return
            slot_addresses = {self.ref_addresses[member][slot] for member in self.group.members}
            if len(slot_addresses) == 1:
                return
            raise MechanicalSynthesisError(f"xl_eval_dependency_not_parameterizable:slot_{slot}")
        if site.callee.startswith(_CELL_FUNCTION_PREFIX):
            raise MechanicalSynthesisError(f"direct_cell_call_unsupported:{site.callee}")
        replacement = self._rewrite_keyword_call(site, slot)
        if replacement is not None:
            self.replacements[id(site.node)] = replacement

    # -- range and INDEX ref-info sites (issue #170) ---------------------------

    def _claim_range_site(self, site: _ReadSite) -> None:
        """Claim both endpoint slots of a literal ``xl_range`` read.

        Constant endpoints keep the call verbatim (no rewrite); a member whose
        recorded endpoint differs makes the range member-varying, which is out
        of scope until a workbook exercises it.
        """
        assert site.endpoints is not None
        for address in site.endpoints:
            try:
                slot = self.exemplar_refs.index(address)
            except ValueError:
                raise MechanicalSynthesisError(
                    f"unclaimed_read_site:{ast.unparse(site.node)}"
                ) from None
            self.claim_counts[slot] = self.claim_counts.get(slot, 0) + 1
            for member in self.group.members:
                if self.ref_addresses[member][slot] != address:
                    raise MechanicalSynthesisError(f"range_endpoints_vary:slot_{slot}")

    def _parsed_endpoint_slot(self, sheet: str, row: int, column_index: int) -> int | None:
        for slot, address in enumerate(self.exemplar_refs):
            if _parse_endpoint_address(address) == (sheet, row, column_index):
                return slot
        return None

    def _claim_index_site(self, site: _ReadSite) -> None:
        """Parameterize the ref-info tuple of a recognized INDEX reference.

        The tuple's corners must equal the exemplar's own parsed endpoint slot
        addresses; each corner is then rederived per member from the recorded
        slot addresses — constant corners stay literal, varying corners become
        registered lookup tables keyed by routing member dimensions.
        """
        info = site.index_ref
        assert info is not None
        row_start, col_start, row_end, col_end = info.corners
        start_slot = self._parsed_endpoint_slot(info.sheet, row_start, col_start)
        end_slot = self._parsed_endpoint_slot(info.sheet, row_end, col_end)
        if start_slot is None or end_slot is None:
            raise MechanicalSynthesisError(f"unclaimed_read_site:{ast.unparse(site.node)}")
        for slot in (start_slot, end_slot):
            self.claim_counts[slot] = self.claim_counts.get(slot, 0) + 1
            self.index_verified_slots.add(slot)

        corners_by_member: dict[str, tuple[int, int, int, int]] = {}
        for member in self.group.members:
            start = _parse_endpoint_address(self.ref_addresses[member][start_slot])
            if start is None or start[0] != info.sheet:
                raise MechanicalSynthesisError(
                    f"index_ref_tuple_mismatch:slot_{start_slot}:{member}"
                )
            end = _parse_endpoint_address(self.ref_addresses[member][end_slot])
            if end is None or end[0] != info.sheet:
                raise MechanicalSynthesisError(f"index_ref_tuple_mismatch:slot_{end_slot}:{member}")
            corners_by_member[member] = (start[1], start[2], end[1], end[2])

        corner_slots = (start_slot, start_slot, end_slot, end_slot)
        corner_names = ("row_start", "col_start_index", "row_end", "col_end_index")
        elements: list[ast.expr] = [ast.Constant(value=info.sheet)]
        for position, (base_name, slot) in enumerate(zip(corner_names, corner_slots, strict=True)):
            values_by_member = {
                member: corners[position] for member, corners in corners_by_member.items()
            }
            expression, derive = self._derived_corner_field(base_name, values_by_member, slot)
            for member in self.group.members:
                if derive(member) != values_by_member[member]:
                    raise MechanicalSynthesisError(f"index_ref_tuple_mismatch:slot_{slot}:{member}")
            elements.append(expression)
        self.replacements[id(info.tuple_node)] = ast.Tuple(elts=elements, ctx=ast.Load())

    def _derived_corner_field(
        self,
        base_name: str,
        values_by_member: Mapping[str, int],
        slot: int,
    ) -> tuple[ast.expr, Callable[[str], BindingKeyValue | None]]:
        """Return the corner's expression and its per-member mirror evaluator."""
        values = set(values_by_member.values())
        if len(values) == 1:
            constant = next(iter(values))
            return ast.Constant(value=constant), lambda member: constant
        member_keys = {
            member: dict(self.expected_member_keys.get(member, {})) for member in self.group.members
        }
        routed = _subset_routing_lookup(
            base_name,
            self.group.members,
            member_keys,
            {member: {base_name: value} for member, value in values_by_member.items()},
            self.varying_dims,
        )
        if routed is None:
            raise MechanicalSynthesisError(f"index_ref_field_unroutable:slot_{slot}")
        key_dims, table = routed
        if isinstance(key_dims, tuple):
            key_params = [self._dim_param(dim) for dim in key_dims]
            table_name = self._register_table(f"{base_name}_by_{'_'.join(key_params)}", table)
            expression: ast.expr = ast.Subscript(
                value=_param_expr(table_name),
                slice=ast.Tuple(elts=[_param_expr(param) for param in key_params], ctx=ast.Load()),
                ctx=ast.Load(),
            )
            tuple_dims = key_dims

            def derive(member: str) -> BindingKeyValue | None:
                key: LookupKey = tuple(member_keys[member][dim] for dim in tuple_dims)
                return table.get(key)

            return expression, derive
        key_param = self._dim_param(key_dims)
        table_name = self._register_table(f"{base_name}_by_{key_param}", table)
        expression = ast.Subscript(
            value=_param_expr(table_name),
            slice=_param_expr(key_param),
            ctx=ast.Load(),
        )
        scalar_dim = key_dims
        return expression, lambda member: table.get(member_keys[member][scalar_dim])

    def _verify_slot(self, slot: int) -> None:
        relation = self.group.ref_relations[slot]
        resolution = relation.resolution
        if resolution.kind == "self_recurrence":
            self._verify_self_recurrence(slot)
            return
        template = resolution.address_template
        if (
            resolution.kind == "xl_cell"
            and template is not None
            and ("{col}" in template or "{row}" in template)
        ):
            # Address templates verify against recorded addresses; keyword call
            # rewrites verify against recorded ref keys. Both hold for accessor
            # reads whose resolution predicted an xl_cell template. Slots
            # claimed only by an INDEX ref-info tuple were already verified
            # against parsed endpoint coordinates and need no axis tables.
            slot_addresses = {self.ref_addresses[member][slot] for member in self.group.members}
            if len(slot_addresses) > 1 and (
                slot not in self.index_verified_slots or slot in self.template_claimed_slots
            ):
                self._verify_templated_slot_address(slot)
        self._verify_derived_keys(slot)

    def _prefix_group_temporaries(
        self, statements: list[ast.stmt]
    ) -> tuple[list[ast.stmt], list[str]]:
        renames: dict[str, str] = {}
        for statement in statements:
            for node in ast.walk(statement):
                if isinstance(node, ast.Name) and _MECHANICAL_TEMP_PATTERN.match(node.id):
                    if self.multi_group:
                        renames.setdefault(node.id, f"_f{self.group_index}{node.id}")
                    else:
                        renames.setdefault(node.id, node.id)
        if self.multi_group:
            renamer = _NameRenamer(renames)
            statements = [renamer.visit(statement) for statement in statements]
        return statements, sorted(set(renames.values()))


class _NameRenamer(ast.NodeTransformer):
    def __init__(self, renames: Mapping[str, str]) -> None:
        self._renames = dict(renames)

    def visit_Name(self, node: ast.Name) -> ast.Name:
        replacement = self._renames.get(node.id)
        if replacement is not None:
            return ast.Name(id=replacement, ctx=node.ctx)
        return node


class _NodeReplacer(ast.NodeTransformer):
    def __init__(self, replacements: Mapping[int, ast.expr]) -> None:
        self._replacements = dict(replacements)

    def generic_visit(self, node: ast.AST) -> ast.AST:
        replacement = self._replacements.get(id(node))
        if replacement is not None:
            return replacement
        return super().generic_visit(node)


def _replace_nodes(statement: ast.stmt, replacements: Mapping[int, ast.expr]) -> ast.stmt:
    replaced = _NodeReplacer(replacements).visit(statement)
    return ast.fix_missing_locations(replaced)


def _routing_condition(
    group_members: Sequence[str],
    expected_member_keys: Mapping[str, Mapping[str, BindingKeyValue]],
    varying_dims: Sequence[str],
    param_by_dim: Mapping[str, str],
    other_groups: Sequence[Sequence[str]],
) -> ast.expr:
    """Build a membership test separating one group from all later groups."""

    def _values(members: Sequence[str], dim: str) -> set[BindingKeyValue]:
        return {expected_member_keys[member][dim] for member in members}

    for dim in varying_dims:
        own = _values(group_members, dim)
        if all(own.isdisjoint(_values(other, dim)) for other in other_groups):
            param = param_by_dim[dim]
            if len(own) == 1:
                return ast.Compare(
                    left=_param_expr(param),
                    ops=[ast.Eq()],
                    comparators=[ast.Constant(value=next(iter(own)))],
                )
            return ast.Compare(
                left=_param_expr(param),
                ops=[ast.In()],
                comparators=[
                    ast.Set(
                        elts=[ast.Constant(value=value) for value in sorted(own, key=_sort_key)]
                    )
                ],
            )

    params = tuple(param_by_dim[dim] for dim in varying_dims)
    own_tuples = {
        tuple(expected_member_keys[member][dim] for dim in varying_dims) for member in group_members
    }
    return ast.Compare(
        left=ast.Tuple(elts=[_param_expr(param) for param in params], ctx=ast.Load()),
        ops=[ast.In()],
        comparators=[
            ast.Set(
                elts=[
                    ast.Tuple(
                        elts=[ast.Constant(value=value) for value in combo],
                        ctx=ast.Load(),
                    )
                    for combo in sorted(own_tuples, key=lambda t: tuple(_sort_key(v) for v in t))
                ]
            )
        ],
    )


def _verify_group_partition(
    groups: Sequence[FingerprintGroup],
    expected_member_keys: Mapping[str, Mapping[str, BindingKeyValue]],
    varying_dims: Sequence[str],
) -> None:
    seen: dict[tuple, str] = {}
    for group in groups:
        for member in group.members:
            keys = expected_member_keys.get(member)
            if keys is None:
                raise MechanicalSynthesisError(f"member_without_expected_keys:{member}")
            combo = tuple(keys.get(dim) for dim in varying_dims)
            if combo in seen and seen[combo] != member:
                raise MechanicalSynthesisError(f"routing_key_collision:{member}:{seen[combo]}")
            seen[combo] = member


def synthesize_cluster_body(
    summary: ClusterFingerprintSummary,
    *,
    key_vocabulary: Sequence[KeyConceptSpec],
    expected_member_keys: Mapping[str, Mapping[str, BindingKeyValue]],
    helper_name: str,
) -> MechanicalBodyDraft:
    """Synthesize a verified parameterized body for one cluster refactor unit.

    Raises :class:`MechanicalSynthesisError` when any read cannot be rewritten
    from the recorded relations or any per-member verification fails.
    """
    if summary.fallback_reason is not None:
        raise MechanicalSynthesisError(f"fingerprint_fallback:{summary.fallback_reason}")
    if not summary.groups:
        raise MechanicalSynthesisError("no_fingerprint_groups")

    dim_sets = {frozenset(keys) for keys in expected_member_keys.values()}
    if len(dim_sets) != 1:
        raise MechanicalSynthesisError("inconsistent_varying_dimensions")
    varying_dims = tuple(sorted(next(iter(dim_sets))))
    if not varying_dims:
        raise MechanicalSynthesisError("no_varying_dimensions")
    param_by_dim = {spec.dimension_id: spec.suggested_param_name for spec in key_vocabulary}
    missing = [dim for dim in varying_dims if dim not in param_by_dim]
    if missing:
        raise MechanicalSynthesisError(f"dimensions_without_vocabulary:{missing}")

    multi_group = len(summary.groups) > 1
    if multi_group:
        _verify_group_partition(summary.groups, expected_member_keys, varying_dims)

    tables = _TableRegistry()
    group_statements: list[list[ast.stmt]] = []
    renameable: list[str] = []
    for index, group in enumerate(summary.groups, start=1):
        synthesizer = _GroupSynthesizer(
            group=group,
            group_index=index,
            multi_group=multi_group,
            helper_name=helper_name,
            param_by_dim=param_by_dim,
            varying_dims=varying_dims,
            expected_member_keys=expected_member_keys,
            tables=tables,
        )
        statements, group_locals = synthesizer.synthesize()
        group_statements.append(statements)
        renameable.extend(group_locals)

    body_statements: list[ast.stmt] = []
    if multi_group:
        for position, (group, statements) in enumerate(
            zip(summary.groups, group_statements, strict=True)
        ):
            is_last = position == len(summary.groups) - 1
            if is_last:
                body_statements.extend(statements)
            else:
                condition = _routing_condition(
                    group.members,
                    expected_member_keys,
                    varying_dims,
                    param_by_dim,
                    [g.members for g in summary.groups[position + 1 :]],
                )
                body_statements.append(ast.If(test=condition, body=statements, orelse=[]))
    else:
        body_statements.extend(group_statements[0])

    module = ast.Module(body=body_statements, type_ignores=[])
    ast.fix_missing_locations(module)
    body = "\n".join(ast.unparse(statement) for statement in module.body)
    # Exemplars from excel-grapher < 3.15.3 still emit None for empty IF arms;
    # lower those to 0.0 so helpers match xl_cell's numeric-blank coercion.
    body = rewrite_empty_if_none_literals(body)

    params = ", ".join(sorted(param_by_dim[dim] for dim in varying_dims))
    indented = "\n".join(f"    {line}" for line in body.splitlines())
    try:
        ast.parse(f"def _draft(ctx, {params}):\n{indented}\n")
    except SyntaxError as error:  # pragma: no cover - defensive
        raise MechanicalSynthesisError(f"draft_body_invalid:{error}") from error

    return MechanicalBodyDraft(
        body=body,
        renameable_locals=tuple(sorted(set(renameable))),
        lookup_table_names=tuple(tables.order),
        group_count=len(summary.groups),
    )
