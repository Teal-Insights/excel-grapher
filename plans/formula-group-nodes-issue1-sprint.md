# Sprint plan: Issue 1 — Address model + storage + edges

**Design:** [formula-group-nodes.md](./formula-group-nodes.md) (Issue 1 section)  
**Tracking:** [#374](https://github.com/Teal-Insights/excel-grapher/issues/374) /
retarget [PR #375](https://github.com/Teal-Insights/excel-grapher/pull/375)  
**Branch:** `issue-374-formula-group-nodes` (or retarget the #375 branch)  
**Follow-up:** Issue 2 eval/codegen (#377) — do not pull into this sprint

---

## Goal

Replace row-only graph nodes with the #375 address model so hand-built
**multi-cell** nodes work end-to-end in the graph (no eval/codegen yet):

- `CellKey` / `RangeKey` / `UnionKey` + `parse_node_key`
- Address-centric `Node`
- Unique cell occupancy + `locate_cell`
- Mixed edges, pickle, projection copy
- Cell-only graphs unchanged; builder still cell-expands

**Out of scope:** skeleton/bindings, fingerprint, evaluator, codegen,
`ProjectedAddress`, detection / `formula_groups=`, OptimalCompression features
beyond `shape != cell` guards.

---

## Locked decisions (do not re-litigate)

| Topic | Decision |
| ----- | -------- |
| Keys | `NodeKey = CellKey \| RangeKey \| UnionKey` (`str` subclasses) |
| Entry | `parse_node_key(value) -> NodeKey` always canonicalize |
| Cover | `members_to_node_key`: sheet → sort → H-runs → V-merge → format |
| One-member union | Collapse to member |
| Empty union / empty members | `ValueError` |
| Node | Store `address`; derive `shape` / sheet / col / row |
| Drop | `NodeKind.row`, `ParsedRowKey`, extent fields as source of truth |
| Row stripe | Just `RangeKey` with `shape=row` (e.g. `Sheet1!D63:Y63`) |
| Multi-cell `value` | Always `None` |
| Occupancy | Cell in at most one node; `remove_node(member)` while owned → error |
| Lookup | `get_node` exact key; members via `locate_cell` |

---

## Starting point in tree

Already present from #375 row work (migrate, do not layer forever):

- `excel_grapher/core/address_keys.py` — row key helpers
- `excel_grapher/grapher/node.py` — `NodeKind.cell/row`, `make_row_node`, locate helpers
- `excel_grapher/grapher/graph.py` — mixed edges for row keys
- `tests/unit/grapher/row_nodes/` — row key/node/graph/location/smoke tests

**Migration policy:** rewrite APIs to the address model; move tests under
`tests/unit/grapher/formula_groups/` (or rename in place); keep behavioral
coverage that still applies (one-row `RangeKey`, occupancy, locate). Delete or
shim `make_row_node` / `ParsedRowKey` once replacements pass.

---

## Sprint breakdown (TDD)

Practice RED → GREEN → refactor each slice. Prefer stubs + failing tests first.
Run: `uv run pytest` on the touched test path after each sprint.

```text
Sprint 1 (keys) → Sprint 2 (Node) → Sprint 3 (graph + occupancy)
        → Sprint 4 (harden + migrate row tests + guards)
```

### Sprint 1 — Key types + cover algorithm ✅

**Files:** `excel_grapher/core/address_keys.py` (and/or new `excel_grapher/grapher/node_key.py`
if splitting keeps modules clearer — prefer extending `address_keys.py` unless
it becomes unwieldy)

| Task | Done when |
| ---- | --------- |
| `NodeShape` enum | `cell`, `row`, `column`, `range`, `union` |
| `CellKey` / `RangeKey` / `UnionKey` | `str` subclasses; geometry properties |
| `parse_node_key` | Canonicalize cell, range, union; strip `$`; order corners |
| One-member union | Collapses to `CellKey` or `RangeKey` |
| Empty union | `ValueError` |
| Union sort + dedupe | Stable member order; duplicates removed |
| `members_to_node_key` | H-run + V-merge cover; order-independent |
| Sheet rules | One sheet → sheet once on key; cross-sheet → per-area qualify |
| Quoted sheets | `'My Sheet'!A1:B2,C3` round-trips |
| Former row case | `Sheet1!D63:Y63` → `RangeKey`, `shape=row` |
| Unit tests | `tests/unit/grapher/formula_groups/test_node_keys.py` |

**Must-pass examples**

| Input | Canonical key | Type / shape |
| ----- | ------------- | ------------ |
| `Sheet1!E63` | same | `CellKey` / cell |
| `Sheet1!Y63:D63` | `Sheet1!D63:Y63` | `RangeKey` / row |
| `Sheet1!E4:I18` (filled cells) | `Sheet1!E4:I18` | `RangeKey` / range |
| shuffled `{A1,B1,C1,D1,E5}` | `Sheet1!A1:D1,E5` | `UnionKey` / union |
| `Sheet1!A1,Sheet1!A1` | `Sheet1!A1` | collapsed `CellKey` |
| `Sheet1!A1,Sheet2!B2` | sheet-qualified union | `UnionKey` |

**Do not** implement occupancy or `Node` changes in this sprint beyond what keys need.

---

### Sprint 2 — Address-centric `Node` + factories ✅

**Files:** `excel_grapher/grapher/node.py`, exports in `excel_grapher/grapher/__init__.py`

| Task | Done when |
| ---- | --------- |
| `Node.address: NodeKey` | Source of truth; `key` / dict identity = `str(address)` |
| Derived props | `shape`, `sheet`, `column`, `row` from parsed address (`None` when N/A) |
| Cell construction | Existing call sites still work via helper (`make_cell_node` / compatible ctor) |
| `make_union_node(members, ...)` | Builds multi-cell node; `value=None`; key from `members_to_node_key` |
| Reject/collapse | Empty members error; single cell → prefer cell node / `CellKey` |
| Remove authoritative | `NodeKind.row`, `min_col`/`max_col`/`min_row`/`max_row` as required fields (shim temporarily if needed for one PR, delete by end of Sprint 4) |
| `NodeView` | Mirrors address-centric surface |
| Locate helpers retarget | Work off `RangeKey`/`UnionKey` expansion, not `kind=row` only |
| Unit tests | `tests/unit/grapher/formula_groups/test_union_node.py` |

**Compat note:** If a flood of call sites breaks on `Node(sheet=..., column=..., row=...)`,
keep a compatibility constructor that sets `address=CellKey(...)` — do not keep
row extent fields as the long-term model.

---

### Sprint 3 — Graph storage, occupancy, edges, locate ✅

**Files:** `excel_grapher/grapher/graph.py` (+ node locate helpers if still in `node.py`)

| Task | Done when |
| ---- | --------- |
| Normalize on add/get/contains | All keys through `parse_node_key` |
| Occupancy index | `cell → owner key`; maintained on `add_node` / `remove_node` |
| Overlap on add | `ValueError` if cell already owned |
| `remove_node(union/range)` | Clears occupancy for all expanded cells |
| `remove_node(member_cell)` while owned | Clear error (not silent no-op) |
| Mixed edges | cell↔union, union↔union, cell↔cell |
| `locate_cell` | Member → owning multi-cell node + member key; bare cell → cell node |
| Pickle | Round-trip preserves address / multi-cell nodes / edges |
| `_copy_for_projection` / subgraph | Copies multi-cell fields + edges |
| Unit tests | `tests/unit/grapher/formula_groups/test_union_graph.py` |

**Transaction reminder for later Issue 3:** occupancy must stay consistent if a
test removes then re-adds; Issue 1 only needs single-node add/remove correctness.

---

### Sprint 4 — Harden, migrate row tests, guards, docs ✅

| Task | Done when |
| ---- | --------- |
| Migrate `tests/unit/grapher/row_nodes/` | ✅ under `formula_groups/`; `make_row_node` remains a deprecated shim |
| Scenario A | ✅ Outside cell → union; dependents / `evaluation_order` |
| Scenario B | ✅ Cross-sheet members → same owner |
| Scenario C | ✅ Builder still emits only cells |
| Scenario D | ✅ TACO + `OptimalCompression().project` skip / no crash |
| Perf smoke | ✅ ~1000 members + `locate_cell` |
| Cell-only regression | ✅ grapher suite |
| Docs blurb | ✅ user guide + `get_node` / `locate_cell` / graph docstrings |
| Export shims | ✅ `make_row_node` deprecated wrapper; guards use `shape == cell` |

**Guard pattern for TACO / compression:** treat `node.shape == NodeShape.cell` as the
only compressible / groupable cell units.

---

## Suggested file layout after Issue 1

```text
excel_grapher/core/address_keys.py    # CellKey, RangeKey, UnionKey, parse, cover
excel_grapher/grapher/node.py         # address-centric Node, make_* , locate_*
excel_grapher/grapher/graph.py        # occupancy, normalize_key via parse_node_key
tests/unit/grapher/formula_groups/
  test_node_keys.py
  test_union_node.py
  test_union_graph.py
  test_union_smoke.py                 # scenarios A–D + perf
```

---

## Test plan checklist

**Keys**

- [x] Cell / range / union round-trip via `parse_node_key`
- [x] One-member union collapses; empty union errors
- [x] Order-independent cover `{A1,B1,C1,D1,E5}` → `Sheet1!A1:D1,E5`
- [x] Vertical merge filled block → single `RangeKey`
- [x] `Sheet1!D63:Y63` → `RangeKey`, `shape=row`
- [x] Quoted sheet + cross-sheet union formatting
- [x] `$` / inverted corners / spaces normalize

**Node / graph**

- [x] Address-centric cell construction back-compat
- [x] `make_union_node` from non-contiguous + cross-sheet members
- [x] Occupancy conflict on overlapping add
- [x] `remove_node(union)` then re-add former member as cell succeeds
- [x] `remove_node(member)` while owned → error
- [x] Mixed edge kinds; pickle; projection/subgraph copy
- [x] `locate_cell` vs `get_node` behavior

**Integration**

- [x] Scenarios A–D
- [x] Builder still cell-only
- [x] 1000-member locate smoke
- [x] Cell-only + migrated former row tests green

---

## Success criteria (merge gate)

- [x] Cell-only graphs behave as today with no caller changes for pure cell APIs
- [x] `#375` address model is the key source of truth (`CellKey`/`RangeKey`/`UnionKey`)
- [x] Non-adjacent members supported; `NodeKind.row` not required
- [x] Occupancy + remove rules enforced
- [x] Pickle and projection copy preserve mixed graphs
- [x] Lookup story documented
- [x] No eval / codegen / detection required

---

## PR / branch notes

- Prefer **one PR for Issue 1** (or stack: keys → node → graph if reviewability needs split).
- Update PR #375 description to point at this sprint + the combined design doc;
  do not merge row-only as the final #374 design.
- Conventional commits (e.g. `feat(grapher): add CellKey/RangeKey/UnionKey parse_node_key`).
