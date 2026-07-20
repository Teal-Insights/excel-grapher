# Sprint plan: Issue 2 — Evaluator + codegen

**Design:** [formula-group-nodes.md](./formula-group-nodes.md) (Issue 2 section)  
**Tracking:** [#392](https://github.com/Teal-Insights/excel-grapher/issues/392)
(parent [#390](https://github.com/Teal-Insights/excel-grapher/issues/390);
supersedes legacy [#377](https://github.com/Teal-Insights/excel-grapher/issues/377)
row-only wording)  
**Depends on:** Issue 1 merged or stacked
([#374](https://github.com/Teal-Insights/excel-grapher/issues/374) /
[PR #375](https://github.com/Teal-Insights/excel-grapher/pull/375))  
**Branch:** `issue-392-formula-groups-eval-codegen`  
**Follow-up:** Issue 3 detection/coalesce (#393) — do not pull into this sprint

---

## Goal

Make **hand-built Option B** multi-cell nodes (`RangeKey` / `UnionKey`)
executable end-to-end with **evaluator ↔ export parity** (lazy per member):

- Shape fingerprint + `specialize_group` (shared by eval and codegen)
- Template fields on the group `Node` (`skeleton`, `member_bindings`,
  `shape_fingerprint`)
- `evaluate("Sheet1!E63")` via `locate_cell` → specialize → scalar; cache under
  the **member** key
- Codegen: one `_group_*` helper + thin per-member wrappers in
  `_RESOLVED_FORMULAS`
- `ProjectedAddress` so projection forwards a public cell to the owning group
  **plus** parameters

**Out of scope:** detection / `coalesce_formula_groups` / builder
`formula_groups=`; member alias nodes; evaluating a multi-area group key as a
vector; varying literals/ops across members; mixing leaf kinds at one hole;
OptimalCompression features beyond not crashing; row/column-only specialize
APIs.

---

## Locked decisions (do not re-litigate)

| Topic | Decision |
| ----- | -------- |
| Occupancy | Option B — no member cell nodes; public API is the member address |
| Eval entry | `evaluate(member_cell)` → scalar; `evaluate(group_key)` → clear error |
| Laziness | Evaluating one member must **not** eagerly evaluate siblings |
| Shape | Fingerprint ignores concrete addresses; ops / fns / literals / leaf **kinds** fixed |
| Specialize | Fill holes in deterministic walk order; kind/arity mismatch → error |
| No rewrite | No column-letter / relative-ref rewrite — holes get full address leaves |
| Hole model | Programmatic `AddressHoleNode` in skeleton AST (not parseable formula text) |
| Leaf kinds | `CELL` / `RANGE` / `WHOLE_COLUMN` / `WHOLE_ROW` ↔ existing AST ref nodes |
| Multi-cell `value` | Stays `None`; results live in evaluator cache under member keys |
| Projection | `map_to_projected(cell) -> ProjectedAddress(address, parameters)` |
| Detection | Issue 3 only — Issue 2 fixtures set skeleton + bindings by hand |

---

## Starting point in tree

**Already present (Issue 1):**

- `excel_grapher/core/address_keys.py` — `CellKey` / `RangeKey` / `UnionKey`,
  `parse_node_key`, `members_to_node_key`, occupancy helpers
- `excel_grapher/grapher/node.py` — address-centric `Node`, `locate_cell`,
  `make_union_node`
- `excel_grapher/grapher/graph.py` — occupancy, mixed edges, eval-order owner
  resolution
- `excel_grapher/core/formula_ast.py` — `CellRefNode` / `RangeNode` /
  `WholeColumnNode` / `WholeRowNode`
- `excel_grapher/evaluator/evaluator.py` — cell path via `get_node` +
  `normalized_formula` (no group path yet)
- `excel_grapher/exporter/codegen.py` + `projection.py` —
  `map_to_projected` returns `str` today; `_RESOLVED_FORMULAS` wrappers exist
  for cells

**Missing (this issue):**

- `excel_grapher/grapher/formula_groups.py` (fingerprint + specialize)
- Template fields + validation on `Node`
- Evaluator `locate_cell` group dispatch + member-key cache
- `ProjectedAddress` + codegen `_group_*` helpers
- Hand-built Option B fixtures + cell-only twins

---

## Sprint breakdown (TDD)

Practice RED → GREEN → refactor each slice. Prefer stubs + failing tests first.
Run: `uv run pytest` on the touched test path after each sprint.

```text
Sprint 1 (fingerprint + specialize)
  → Sprint 2 (Node template fields + fixtures)
  → Sprint 3 (evaluator Option B path)
  → Sprint 4 (ProjectedAddress + codegen)
  → Sprint 5 (parity / errors / cell-only regression)
```

Sprints 3–4 may overlap after Sprint 2 fixtures exist; keep
`specialize_group` as the single shared specialization path.

---

### Sprint 1 — Fingerprint + `specialize_group` ✅

**Files:** `excel_grapher/grapher/formula_groups.py` (new),
`AddressHoleNode` / `AddressLeafKind` in `excel_grapher/core/formula_ast.py`.

| Task | Done when |
| ---- | --------- |
| `AddressLeafKind` | ✅ `cell`, `range`, `whole_column`, `whole_row` |
| `AddressHoleNode` | ✅ Slot index + kind; only in skeletons (not parser output) |
| `shape_fingerprint(ast) -> str` | ✅ Walk AST; address leaves → typed holes; bake everything else |
| Equal fingerprints | ✅ Same ops/fns/literals/leaf kinds/arity; different addresses OK |
| Distinct fingerprints | ✅ Differing literal, op, fn name, arity, or leaf kind at a slot |
| `specialize_group(skeleton, bindings) -> AstNode` | ✅ Fill holes in walk order with concrete address leaves |
| Reject | ✅ Wrong binding length; kind mismatch |
| Unit tests | ✅ `tests/unit/grapher/formula_groups/test_fingerprint.py`, `test_specialize.py` |

**Must-pass examples**

| Case | Expect |
| ---- | ------ |
| INDEX/MATCH pair differing only in one `CELL` | Same fingerprint |
| Same shape but `MATCH(...,0)` vs `MATCH(...,1)` | Different fingerprint |
| Skeleton with one `CELL` hole + binding `Sheet1!D35` | Specialize → `CellRefNode("Sheet1!D35")` |
| Binding `RANGE` into a `CELL` hole | Error |
| Two holes, one binding | Error |

**Do not** wire evaluator or codegen in this sprint.

---

### Sprint 2 — Template fields on `Node` + Option B fixtures

**Files:** `excel_grapher/grapher/node.py`, `make_union_node` / helpers,
`tests/fixtures/formula_groups/`, unit fixture builders

| Task | Done when |
| ---- | --------- |
| Fields on multi-cell `Node` | `shape_fingerprint`, `skeleton: AstNode \| None`, `member_bindings` |
| Cell nodes | Template fields stay `None` / empty; no behavior change |
| Attach validation | Binding arity == hole count; kinds align; every member has an entry |
| `make_union_node(..., skeleton=..., member_bindings=..., shape_fingerprint=...)` | Builds Option B group; `value=None` |
| Fixture: contiguous one-row stripe | e.g. `Sheet1!D63:Y63` with INDEX/MATCH template |
| Fixture: non-contiguous / cross-sheet union | Members on two sheets; unique occupancy |
| Cell-only twin | Same public formulas as discrete cell nodes (no multi-cell node) |
| Pickle / projection copy | Preserve template fields |
| Unit tests | `tests/unit/grapher/formula_groups/test_group_node_template.py` |

**Fixture invariant:** graph contains the group node **and** its precedents as
cells (or other groups), but **no** cell node whose key is a member of the
group.

**Do not** change `FormulaEvaluator.evaluate` dispatch yet beyond what tests
need as stubs.

---

### Sprint 3 — Evaluator Option B path

**Files:** `excel_grapher/evaluator/evaluator.py` (+ thin helpers if needed)

| Task | Done when |
| ---- | --------- |
| Member dispatch | `evaluate(member)` → `locate_cell` → owner with `shape != cell` |
| Specialize then eval | `specialize_group(owner.skeleton, owner.member_bindings[member])` → `_evaluate_ast` |
| Cache key | Result cached under **member** address, not the group key |
| Laziness | Evaluating member A does not populate cache for sibling B |
| Missing template | Clear error if group lacks skeleton/bindings for the member |
| Reject group key | `evaluate("Sheet1!D63:Y63")` / union key → dedicated error (not KeyError ambiguity) |
| Cell-only path | Unchanged when `locate_cell` returns a cell node |
| Unit tests | `tests/unit/evaluator/test_formula_group_eval.py` (or under `formula_groups/`) |

**Entrypoint (locked)**

```text
evaluate("Sheet1!E63")
  → locate_cell → shape != cell
  → specialize_group(skeleton, member_bindings[E63])
  → eval → cache["Sheet1!E63"] → scalar
```

**Note:** Today's `_evaluate_cell` uses `get_node(norm)` and raises if missing.
Issue 2 must resolve members via `locate_cell` / `cell_owner` before treating
absence as an error.

---

### Sprint 4 — `ProjectedAddress` + codegen `_group_*`

**Files:** `excel_grapher/exporter/projection.py`, `codegen.py`,
callers of `map_to_projected`

| Task | Done when |
| ---- | --------- |
| `ProjectedAddress` | `address: NodeKey`, `parameters: Mapping[str, Any] \| None` |
| Protocol update | `map_to_projected(cell) -> ProjectedAddress` (greenfield break OK) |
| Identity / cell maps | `parameters is None`; `address` is the cell or forwarded key |
| Group maps | Public member → owning group key + binding parameters for wrappers |
| Codegen helper | One `_group_<stable_id>(...)` per group node; body from specialized skeleton pattern |
| Wrappers | Per-member entries in `_RESOLVED_FORMULAS` calling the helper with that member’s bindings |
| Wrappers ≠ nodes | No member cell nodes added to the graph |
| Shared specialize | Codegen uses the same `specialize_group` (or emits equivalent closed-over constants) |
| Unit / smoke | `tests/unit/exporter/test_formula_group_codegen.py` |

**Parameter shape (suggested lock for MVP):**

```python
parameters = {
    "member": "Sheet1!E63",
    "bindings": (...),  # same tuple as member_bindings[member], serializable
}
```

Exact serialization may use address strings rather than live AST nodes in the
exported module — document the chosen encoding in the Sprint 4 PR.

**Compat:** Update every in-repo `map_to_projected` implementer and test that
assumes `str`. Prefer a short migration in the same PR over a dual API.

---

### Sprint 5 — Parity, errors, cell-only regression

**Files:** integration / parity tests under `tests/integration/` + fixtures

| Task | Done when |
| ---- | --------- |
| Twin parity | `evaluate(member)` equals cell-only twin for that address alone |
| Export ↔ evaluator | `assert_codegen_matches_evaluator` (values + error **codes**) |
| Error channels | Evaluator sentinels vs export `XlErrorException` — same codes |
| Non-contiguous / cross-sheet | Fixtures from Sprint 2 pass eval + export |
| Group-key eval | Rejected consistently in evaluator (and not silently codegen’d as a target) |
| Cell-only regression | Existing evaluator + codegen suites green with no multi-cell nodes |
| Occupancy | Fixtures still unique-occupancy; no member cell nodes |

Prefer cache-based / in-process parity on Linux CI; live Excel remains
run-if-available (`pytest.skip` when automation missing).

---

## Suggested file layout after Issue 2

```text
excel_grapher/grapher/formula_groups.py   # fingerprint, specialize, hole kinds
excel_grapher/core/formula_ast.py         # AddressHoleNode (if added to AstNode)
excel_grapher/grapher/node.py             # template fields + validation
excel_grapher/evaluator/evaluator.py      # locate_cell group path
excel_grapher/exporter/projection.py     # ProjectedAddress
excel_grapher/exporter/codegen.py         # _group_* + wrappers
tests/fixtures/formula_groups/            # hand-built Option B + twins
tests/unit/grapher/formula_groups/
  test_fingerprint.py
  test_specialize.py
  test_group_node_template.py
tests/unit/evaluator/test_formula_group_eval.py
tests/unit/exporter/test_formula_group_codegen.py
tests/integration/...                     # parity harness cases
```

---

## Test plan checklist

**Fingerprint / specialize**

- [ ] Fingerprint ignores concrete addresses; distinguishes literals / ops / fns
- [ ] Same leaf-kind holes at the same walk indices ⇒ same fingerprint
- [ ] Specialize fills holes in walk order
- [ ] Rejects kind mismatch and arity mismatch

**Node / fixtures**

- [ ] Option B fixture: no member cell nodes; occupancy unique
- [ ] Non-contiguous + cross-sheet members supported
- [ ] Cell-only twin exists for parity
- [ ] Template fields pickle / copy with the graph

**Evaluator**

- [ ] `evaluate(member)` matches twin for that address alone
- [ ] Sibling members not cached when one member is evaluated
- [ ] Group key evaluation rejected with a clear error
- [ ] Cell-only graphs unchanged

**Codegen / projection**

- [ ] One `_group_*` helper per group; wrappers pass correct bindings
- [ ] `map_to_projected(member)` → owning address + parameters
- [ ] Codegen ↔ evaluator parity (values + error codes)
- [ ] Cell-only codegen unchanged when no groups present

---

## Success criteria (merge gate)

- [ ] Member eval equals cell-only twin; lazy; non-contiguous / cross-sheet work
- [ ] Export matches evaluator; fixtures obey unique occupancy
- [ ] Shared `specialize_group` used by evaluator and codegen
- [ ] `ProjectedAddress` is the projection map result type
- [ ] No detection / coalesce / builder flag required
- [ ] No row/column-only specialize API

---

## PR / branch notes

- Prefer **one PR for Issue 2**, or a short stack: specialize → fixtures+eval →
  codegen/parity if reviewability needs a split.
- Stack on Issue 1 (`issue-374-…` / PR #375) until that merges; retarget onto
  `main` afterward.
- Close or comment on legacy [#377](https://github.com/Teal-Insights/excel-grapher/issues/377)
  pointing at #392 when this lands.
- Conventional commits, e.g.:
  - `feat(grapher): add shape fingerprint and specialize_group`
  - `feat(evaluator): evaluate formula-group members via locate_cell`
  - `feat(exporter)!: return ProjectedAddress from map_to_projected`
