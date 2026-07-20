# Sprint plan: Issue 3 — Detection + coalesce

**Design:** [formula-group-nodes.md](./formula-group-nodes.md) (Issue 3 section)  
**Tracking:** [#393](https://github.com/Teal-Insights/excel-grapher/issues/393)
(parent [#390](https://github.com/Teal-Insights/excel-grapher/issues/390))  
**Depends on:** Issues 1–2 merged or stacked
([#391](https://github.com/Teal-Insights/excel-grapher/issues/391) /
[#392](https://github.com/Teal-Insights/excel-grapher/issues/392) /
[PR #445](https://github.com/Teal-Insights/excel-grapher/pull/445))  
**Branch:** `group-node-coalesce`  
**Follow-up:** none in this plan — do not pull OptimalCompression group
inlining, fuzzy matching, or default-on builder into these sprints

---

## Goal

Add an **opt-in post-pass** that finds same-shape formula families on a
**cell-only** graph and rewrites them into Issue 2–compatible formula-group
nodes (unique occupancy, skeleton + bindings), so evaluator and codegen need
no second entrypoint:

- Discover families via Issue 2 `shape_fingerprint`
- Build skeleton + `member_bindings` by aligned leaf walk (bake vs hole)
- `coalesce_formula_groups(graph)` mutates in place with edge rewrite
- `create_dependency_graph(..., formula_groups=False)` default unchanged
- `target_keys()` still enumerates **member addresses** for former targets
- Eval / codegen parity against the pre-coalesce cell-only twin

**Out of scope:** enabling groups by default; absorbing stragglers into an
existing group on a second pass; rewriting dependent formula text to name the
group key; coalescing intra-family-edge families (skip whole family); fuzzy /
literal-tolerant fingerprints; OptimalCompression inlining through groups;
vector / multi-area `evaluate(group_key)`; incremental / watch-mode coalesce.

---

## Locked decisions (do not re-litigate)

| Topic | Decision |
| ----- | -------- |
| API shape | Post-pass `coalesce_formula_groups(graph, *, min_family_size=2) -> CoalesceReport` |
| Builder | `formula_groups: bool = False`; when True, call coalesce at end of build |
| Candidates | `CellKey` nodes with non-empty `normalized_formula` only |
| Fingerprint | Reuse Issue 2 `shape_fingerprint` (no second fingerprint dialect) |
| Intra-family edges | Skip **whole** family (`intra_family_edge`); leave as cells |
| Cross-sheet | Allowed in one fingerprint bucket |
| Skeleton build | Aligned address-leaf walk; all-equal → bake; else hole + per-member binding |
| Group key | `members_to_node_key(M)` (Issue 1) |
| Occupancy | Remove member cell nodes; `locate_cell(m)` → group |
| Formula text | Do **not** rewrite dependents; they still name member cells |
| Targets | Store `target_members`; `target_keys()` lists member addresses, not group key |
| Metadata | MVP: drop conflicting per-cell metadata; preserve `target_members` explicitly |
| Idempotency | Second pass ignores existing groups; no merge of groups with leftover cells |
| Eval / export | Unchanged Issue 2 paths on the coalesced graph |

---

## Starting point in tree

**Already present (Issues 1–2):**

- `excel_grapher/core/address_keys.py` — `CellKey` / `RangeKey` / `UnionKey`,
  `members_to_node_key`
- `excel_grapher/grapher/node.py` — `make_union_node`, template fields,
  `locate_cell`, occupancy
- `excel_grapher/grapher/graph.py` — edges, guards, provenance merge,
  `target_keys()` (cell `is_target` only today)
- `excel_grapher/grapher/formula_groups.py` — `shape_fingerprint`,
  `specialize_group`, `validate_group_template`, `collect_holes`
- Evaluator + codegen group paths (PR #445)
- Hand-built fixtures / twins under `tests/fixtures/formula_groups/hand_built.py`

**Missing (this issue):**

- Family discovery + `build_group_template` (pure, no mutate)
- `coalesce_formula_groups` + `CoalesceReport` / `SkippedFamily`
- Edge snapshot → remove members → add group → rewrite edges
- `target_members` + `target_keys()` expansion
- Builder `formula_groups=` wiring
- Cell-only → coalesce fixtures + parity tests

---

## Sprint breakdown (TDD)

Practice RED → GREEN → refactor each slice. Prefer stubs + failing tests first.
Keep **pure detection** separate from **graph mutation**.
Run: `uv run pytest` on the touched test path after each sprint.

```text
Sprint 1 (discover + skeleton, no mutate)
  → Sprint 2 (coalesce transform + targets)
  → Sprint 3 (builder flag + fixtures + docs)
  → Sprint 4 (parity / skips / cell-only regression)
```

Sprints 1–2 may share fixtures built as cell-only twins of Issue 2 hand-built
groups; Sprint 4 must prove coalesced graphs behave like those hand-built
groups under evaluator + codegen.

---

### Sprint 1 — Discover + skeleton (no graph mutation) ✅

**Files:** extend `excel_grapher/grapher/formula_groups.py`
(or `formula_groups/detect.py` if the module grows too large).

| Task | Done when |
| ---- | --------- |
| Candidate scan | ✅ Only formula `CellKey` nodes; existing groups ignored |
| Cluster | ✅ `dict[fingerprint, list[NodeKey]]` via `shape_fingerprint` |
| Unparseable | ✅ Omitted from clusters (cell left alone) |
| `iter_formula_families` / equivalent | ✅ Yields `ReadyFamily` / `SkippedFamily` |
| Intra-family edge | ✅ Any `a→b` with both in `M` → skip (`intra_family_edge`) |
| `below_min_size` | ✅ `len(M) < min_family_size` (default 2) |
| `build_group_template` | ✅ Bake equal leaves; hole + bindings for differing leaves |
| Walk order | ✅ Same address-leaf order as fingerprint / `specialize_group` |
| Determinism | ✅ Fingerprints sorted; members workbook-order |
| Unit tests | ✅ `tests/unit/grapher/formula_groups/test_detect.py` |

**Must-pass examples**

- Two `=A1+1` cells with different `A1` refs → one family, one `CELL` hole
- `=A1+1` vs `=A1+2` → different fingerprints (no cluster)
- Shared range refs bake; differing cell refs become holes
- Cross-sheet / non-contiguous members allowed in one family
- Member→member edge → whole family skipped with `intra_family_edge`

---

### Sprint 2 — Coalesce transform + targets ✅

**Files:** `formula_groups.py` (coalesce), `graph.py` (`target_keys`),
`node.py` if `target_members` becomes a real field (else metadata convention —
pick one and stick to it).

| Task | Done when |
| ---- | --------- |
| `CoalesceReport` / `SkippedFamily` | ✅ Created group keys + skipped families with reasons |
| `coalesce_formula_groups(graph)` | ✅ In-place mutate; returns report |
| Transaction order | ✅ Snapshot → `remove_node(m)` for each member → `add_node(G)` → re-add edges |
| Group construction | ✅ `make_union_node` + Issue 2 template fields + `members_to_node_key` |
| Inbound edges | ✅ `d → m` rewritten to `d → G`; guards/provenance via existing merge |
| Outbound edges | ✅ `m → dep` rewritten to `G → dep`; union of member deps |
| Occupancy | ✅ No member `CellKey` nodes; `locate_cell(m)` → group for all `m ∈ M` |
| `target_members` | ✅ Stored in `metadata["target_members"]` |
| `target_keys()` | ✅ Lists those member addresses (not the multi-area group key) |
| Idempotency | ✅ Second call creates no duplicate groups |
| Unit tests | ✅ `tests/unit/grapher/formula_groups/test_coalesce.py` |

**Field merge (locked for this sprint)**

| Field | Rule |
| ----- | ---- |
| `value` | Always `None` on the group |
| `formula` / `normalized_formula` | Optional specialized string for first member, or `None` |
| `is_leaf` | `False` if group has deps after rewrite |
| `is_target` | `True` if any member was a target |
| `metadata` | Do not silently merge conflicting per-cell keys; keep `target_members` explicit |

---

### Sprint 3 — Builder wiring + fixtures + docs ✅

**Files:** `excel_grapher/grapher/builder.py`, fixtures, `user_guide/01-dependency-graphs.qmd`.

| Task | Done when |
| ---- | --------- |
| `formula_groups: bool = False` | ✅ Default off; existing builder tests green unchanged |
| Flag on | ✅ When True, run `coalesce_formula_groups` at end of build |
| Cell-only fixture | ✅ Cross-sheet same-shape pair (`tests/fixtures/formula_groups/cell_only.py`) |
| Docs | ✅ User-guide note: opt-in flag + unique occupancy / member-address API |
| Demo (optional) | ✅ `examples/micro_workbooks/demo_formula_groups.py` Issue 3 section |

CLI: mirror the flag only if a graph-build CLI already exists; otherwise
Python API + user guide is enough.

---

### Sprint 4 — Parity, skips, cell-only regression ✅

**Files:** `tests/integration/exporter/test_formula_group_coalesce_parity.py`.

| Task | Done when |
| ---- | --------- |
| Eval twin | ✅ `evaluate(member)` on coalesced graph equals pre-coalesce cell graph |
| Export twin | ✅ `assert_codegen_matches_evaluator` on coalesced graph |
| Hand-built parity | ✅ Coalesced result matches Issue 2 hand-built group for the same family |
| Intra-family skip | ✅ Family with member→member edge remains cells; report reason |
| Min size | ✅ Lone formula cell never becomes a 1-member group |
| Cell-only regression | ✅ `formula_groups=False` path identical to pre-Issue-3 |
| TACO / OptimalCompression | ✅ Coalesced graph: `shape != cell` guards / no crash |
| Dependent formulas | ✅ Still name member addresses after coalesce |

Prefer cache-based / in-process parity on Linux CI; live Excel remains
run-if-available (`pytest.skip` when automation missing).

---

## Suggested file layout after Issue 3

```text
excel_grapher/grapher/formula_groups.py   # fingerprint, specialize, detect, coalesce
excel_grapher/grapher/builder.py          # formula_groups= flag
excel_grapher/grapher/graph.py            # target_keys expands target_members
excel_grapher/grapher/node.py             # target_members (field or metadata)
tests/fixtures/formula_groups/            # cell-only families + expected groups
tests/unit/grapher/formula_groups/
  test_detect.py
  test_coalesce.py
tests/integration/exporter/
  test_formula_group_coalesce_parity.py   # Sprint 4 parity / skips / regression
user_guide/01-dependency-graphs.qmd       # opt-in note
```

---

## Test plan checklist

**Discovery**

- [x] Two identical-shape formulas cluster; `A1+1` vs `A1+2` do not
- [x] Cross-sheet / non-contiguous cells in one family
- [x] Aligned walk: shared refs baked; differing refs → holes + bindings
- [x] Intra-family edge → skip whole family + report `intra_family_edge`
- [x] `len < 2` → `below_min_size`
- [x] Unparseable formula omitted from clusters

**Coalesce**

- [x] After pass: no member `CellKey` nodes; `locate_cell` → group
- [x] Group key == `members_to_node_key(M)`
- [x] Inbound edges merged onto group (guards/provenance via existing merge)
- [x] Outbound deps = union of member deps
- [x] Dependent formulas **unchanged** (still name member addresses)
- [x] `target_keys()` lists former target **member** addresses
- [x] Second `coalesce_formula_groups` is idempotent (no duplicate groups)
- [x] Evaluation order on coalesced graph does not crash / includes group node

**Wiring + parity**

- [x] `create_dependency_graph(..., formula_groups=False)` unchanged
- [x] `formula_groups=True` produces formula groups on fixture
- [x] `evaluate(member)` equals pre-coalesce cell-only twin
- [x] Codegen ↔ evaluator parity on coalesced graph
- [x] OptimalCompression/TACO do not crash on coalesced graph

---

## Success criteria (merge gate)

- [x] Opt-in detection emits Issue 2–compatible formula groups
- [x] Default builder path remains cell-only
- [x] Intra-family families are skipped safely (no `G→G` footgun)
- [x] Targets remain addressable as member keys via `target_keys()`
- [x] No new eval entrypoint; no fuzzy matching; no default-on coalesce

---

## PR / branch notes

- Prefer **one PR for Issue 3**, or a short stack: detect → coalesce →
  builder/parity if reviewability needs a split.
- Stack on Issue 2 (`group-node-eval-codegen` / PR #445) until that merges;
  retarget onto `main` afterward.
- Conventional commits, e.g.:
  - `feat(grapher): discover formula families by shape fingerprint`
  - `feat(grapher): coalesce same-shape cells into formula-group nodes`
  - `feat(grapher): opt-in formula_groups= on create_dependency_graph`
