# LIC-DSF may-cycle SCC#0 (issue 533)

Investigation of the solver MCVE at
[`mcve-lic-dsf-may-cycle-scc0`](https://github.com/Teal-Insights/excel-grapher/releases/tag/mcve-lic-dsf-may-cycle-scc0).
Measurements used excel-grapher on that fragment (schema `1.0.0`,
`kind: may_cycle_solver_mcve`), not a full workbook re-extract.

## 1. Can current `GuardConstraints` prove edges dead under leaf domains?

**No, not enough to shrink or eliminate SCC#0.** Baseline
`_subgraph_has_feasible_cycle` still finds a guard-feasible cycle. The
published 79-node witness is jointly consistent under the old solver
because:

- `Compare(CellRef, CellRef)` was **opaque**. The witness mixes
  `Ext_Debt_Data!C120 = lookup!X4` and `NOT(C120 = lookup!X4)` on one
  path; those two strings did not contradict.
- Leaf domains were **not consulted**. Complementary inequalities
  `NOT(B=0) AND NOT(B=1)` stayed feasible without an enum `{0, 1}`.
- Guard CellRefs are almost never leaves: **1 / 4548** `guard_refs` is a
  leaf (`C78`, domain `Between(0, 1)`). `NOT(C78=0)` is still feasible
  when `C78=1`.

Intra-SCC guard shapes (139,459 edges):

| Count | Shape |
| ---: | --- |
| 81,144 | `AND(cell<>0, cell=int)` (CHOOSE-index match) |
| 50,214 | unguarded |
| 4,960 | `NOT(cell=cell)` |
| 2,727 | `cell=cell` |
| 353 | `NOT(C78=0)` |
| 61 | `OR(cell=cell, NOT(cell=cell))` (tautology from OR-merge) |

All comparison operators on those edges are `=`. Guard cone formulas
are regular: 2,898 `IF(x-y>0, x-y, 0)` index cells, ~396 identity
aliases, ~2,829 `+1` year chains. The cone is **disjoint from the SCC**.

**This PR** teaches `GuardConstraints` cell-cell unification, literal
operands on either side, and `CellTypeEnv` enum/interval checks, and
rewrites identity formulas before conjoining guards. That:

- Rejects the published witness (`P` and `NOT(P)` on `C120` vs `X4`).
- Makes the documented `{0,1}` IF may-cycle (extraction basics §08)
  disappear when `DynamicRefConfig` is passed.
- **Does not empty SCC#0.** A shorter residual cycle remains, using
  `NOT(C78=0)`, a CHOOSE-index `AND(E759<>0, E759=1)`, and a residency
  inequality. `C78=1` is in-domain; the index cell is a cone `IF`, not a
  leaf.

## 2. Additional solver capabilities

Needed to finish SCC#0 (and likely the other 19 may-SCCs):

1. **Value-based guard pruning** (issue 30) — evaluate `GuardExpr`
   against a concrete assignment (cached workbook values or user
   inputs) and drop edges whose guards are false. On this MCVE, cached
   values make **81,906** guarded edges false and **0 residual cyclic
   SCCs**. That matches Excel calculating with iterate off: the saved
   state has no live cycle. This is assignment-specific, not a domain
   proof.
2. **Guard-cone abstract interpretation** — propagate leaf domains
   through identity aliases (done), `+1` year chains, and
   `IF(x-y>0, x-y, 0)` so CHOOSE-index equalities `idx=N` can be proven
   dead for some `N` under the year/`Between` domains. This is the
   remaining work for *universal* infeasibility (every input in the
   constraint box).
3. **Not required for SCC#0:** OR-merged multi-edges (issue 34). The
   61 tautological `OR(P, NOT(P))` arcs are a smell, but the residual
   cycle does not depend on OR-merge. Cell-cell opacity, not storage
   shape, is what made the published witness look feasible.

`evaluate_guard` in this PR is the three-valued helper issue 30 would
use; it is not a `prune_inactive_guards` graph API.

## 3. Design for resolving SCC#0, and the other 19 may-SCCs

**Phase A (this PR).** Symbolic solver: cell-cell unification, leaf
domains, identity-alias rewrite, `cycle_report(cell_type_env=...)`,
solver-MCVE loader. Small, local to `GuardConstraints` / cycle DFS.
Unblocks the documented constraints API. Does not clear LIC-DSF SCC#0.

**Phase B (issue 30).** `graph.prune_inactive_guards(values=...)` or
equivalent, then re-run SCC. On the saved LIC-DSF cache this **clears
SCC#0** (and should clear the other 19 if they share the same C78 /
CHOOSE-index / residency cone — they are described as sharing a large
financing cone). Downstream `evaluation_order(strict=True)` then works
for that assignment. Document that other inputs may reactivate edges.

**Phase C.** Interval / enum propagation on the 6,203-node disjoint
guard cone (IF-diff-or-zero, year increments, identity). Needed only if
we must prove acyclicity for **all** inputs in the leaf domains, not
just the saved cache. The cone vocabulary is small enough that a
specialized interpreter is more appropriate than a general Excel SMT
solver.

Same approach should generalize to the other 19 may-SCCs: they share
the financing cone, `C78`, residency aliases (`B116` / `C120` /
`lookup!X4` → `translation!C90` singleton), and CHOOSE-index guards.

## 4. MCVE format gaps

Useful additions for follow-up fixtures:

- **Axis-preserving formulas or `formula_ast`.** `normalized_formula`
  strips `$`, so identity detection must resolve relative axes against
  the host cell (this PR does that). Shipping AST JSON would avoid that
  ambiguity.
- **Quoted vs unquoted keys.** Graph / `GuardExpr` keys quote sheet
  names; `CellTypeEnv` keys do not. Fragments should document both, or
  include a `cell_type_env` object keyed like `normalize_cell_type_env_key`.
- **The other 19 SCCs** as separate fragments (same schema). SCC#0 is
  the priority witness but not a stand-in for size-419 clusters without
  measurement.
- **Cached-value pruning summary** in `summary.json` (live vs dead
  guarded edges, residual SCC count) so Phase B can be checked without
  replaying 139k edges.
- **Intra-SCC formulas** are optional for feasibility; they help
  explain unguarded edges. Guard-cone formulas are the ones that matter
  for Phase C.
- Schema is otherwise sufficient: `intra_scc_edges` + `guard` JSON,
  disjoint cone, 80/80 leaf constraints, published path.

Loader: `excel_grapher.grapher.solver_mcve.load_solver_mcve`.
