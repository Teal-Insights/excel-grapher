---
name: author-bindings
description: >
  Author, draft, repair, or finish excel-grapher series-binding sidecars for an
  arbitrary workbook. Use when the user asks to write inputs/outputs/internals/
  constants YAML, fix bind resolution, or finish coverage of formula cells.
---

# Author series bindings

Teach an agent to **declare** series bindings for a workbook, then run fail-closed
checks until codegen would succeed. This is not discovery (`#724`) and not a
bulk YAML generator.

## Schema version

Do **not** hardcode `schema_version`. Read it from the installed package:

```bash
uv run python -c "from excel_grapher.series_bindings import CURRENT_SCHEMA_VERSION; print(CURRENT_SCHEMA_VERSION)"
```

Authoritative field shapes live in the bundled JSON Schema:

```python
from importlib.resources import files

schema_text = (
    files("excel_grapher.series_bindings")
    .joinpath("series_binding.schema.json")
    .read_text(encoding="utf-8")
)
```

`SUPPORTED_SCHEMA_VERSIONS` lists every accepted version. Use `CURRENT_SCHEMA_VERSION`
for new sidecars. Implemented layouts are `scalar`, `series`, and `matrix`
(`IMPLEMENTED_LAYOUTS`). Legacy `row_series` is accepted at load time and
normalized to `series`.

## Trigger

Use this skill when the user asks to author, draft, repair, or finish series
bindings for a workbook.

## Workflow

1. Read the optional human guide / public API table. If a series is ambiguous,
   explain the ambiguity; **do not guess**.
2. Bootstrap four shards (empty `series: []` is fine):
   `inputs.bindings.yaml`, `outputs.bindings.yaml`, `internals.bindings.yaml`,
   `constants.bindings.yaml`.
3. Cross-check proposed series against the extracted graph:
   - `input` / `constant` overlap **leaves**
   - `output` overlap **target / formula output** nodes
   - `internal` overlap **formula** nodes
4. Author **one series per semantic family** (a scalar parameter, a 1-D series,
   or a 2-D table), not one entry per cell. Prefer `layout: matrix` when rows
   share a label dimension and headers share a key (usually `TIME_PERIOD`).
5. **Do not** dump burndown rows, sheet geometry, or later `#724` candidates into
   YAML as a first pass. Walk sheets and tables with domain meaning, and write
   the intended series correctly once.
6. Run bundled checks (`validate` then `audit`) until resolution is clean.
   Use `burndown` only as a **coverage worklist** for remaining holes *inside
   the current bound graph closure* (the dependency graph built from bound
   `data_range` targets). It is **not** a full workbook walk. Empty shards
   yield an empty graph and a zero unbound count; that is not “done.” Author
   remaining holes with the same semantic standard (often another matrix).
7. Optional: `bindings upsert` writes **one** already-reflected series after
   fail-closed checks. A workbook-specific script that upserts many semantic
   families is allowed. A generic catalog → four-file replace is not.
8. Return the sidecars plus a short validation summary.

Assets under `assets/` are **pedagogical** (how a series looks). They are **not
an input to a generator**.

## Commands

Same workbook wiring as `bindings validate`: `WORKBOOK`, `--bindings`,
`--constraints`, `--use-cached-dynamic-refs`, `--blank-ranges`.

```bash
uv run excel-grapher bindings validate WORKBOOK --bindings BINDINGS_DIR
uv run excel-grapher bindings audit WORKBOOK --bindings BINDINGS_DIR
uv run excel-grapher bindings burndown WORKBOOK --bindings BINDINGS_DIR
uv run excel-grapher bindings upsert WORKBOOK --bindings BINDINGS_DIR --series one_series.yaml
```

`validate` can look fine while resolution `ok=False`. Codegen requires `ok=True`.
`audit` is the codegen-fatal gate. `burndown` is advisory coverage of the
**current bound graph closure** (exit 1 only with `--strict`). It is not a
sheet-by-sheet walk of every formula in the workbook.

This skill is **not** inside the `excel-grapher` wheel. Agents load a copied
folder, not a Python extra. From a git clone or the `excel-grapher` sdist:

```bash
mkdir -p .agents/skills
cp -R skills/author-bindings .agents/skills/author-bindings
```

Cursor also loads `.cursor/skills/author-bindings`. Claude Code uses
`.claude/skills/author-bindings`. User-level installs go under
`~/.agents/skills/author-bindings`.

Thin wrappers: `scripts/audit.sh`, `scripts/burndown.sh`, `scripts/upsert.sh`.

## Progressive disclosure

- [references/conventions.md](references/conventions.md) — directions, layouts, keys, dimension `id` vs `concept`
- [references/pitfalls.md](references/pitfalls.md) — fill, measure shards, shared names, blanks, unique-address ownership
- [references/validation-loop.md](references/validation-loop.md) — validate → derive_* → audit → burndown → upsert
