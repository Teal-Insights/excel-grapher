Use the `gh` CLI tool to manage issues and pull requests.

Always run Python code with `uv run`. Use `uv add` to add dependencies or `uvx` to use uninstalled Python command line tools.

We use `fastpyxl` as a drop-in replacement for `openpyxl`.

`.qmd` files can be rendered to markdown with `uv run quarto render path/to/file.qmd`.

Always practice test-driven development. Write a stub (if necessary), write a test, watch it fail for the right reason (RED), write the code to make it pass (GREEN), and then refactor to clean up the code.

This is a greenfield project with no users, so we are free to make design decisions that unconstrained by the legacy codebase.

## Docstrings

Use **Google-style** docstrings in `excel_grapher/` (see `.cursor/rules/docstrings.mdc`).
Use single backticks for inline code in docstrings (not reST double-backtick literals).
Series-binding codegen defaults to the `google` renderer.

## Documentation site (great-docs)

On Windows, set UTF-8 mode before building or previewing (great-docs prints Unicode in its
build log and post-render script):

```bash
uv run python scripts/great_docs_build.py
uv run python scripts/great_docs_preview.py
```

Equivalent: `PYTHONUTF8=1 uv run great-docs build`. The `pre_render` hook in `great-docs.yml`
also patches `post-render.py` for Quarto's subprocess.

## Parity

The project aims for **behavioral parity** across **Excel** (reference), **`FormulaEvaluator`**, and **exported standalone code**. Semantics are centralized in `excel_grapher/exporter/export_runtime/`; the evaluator and codegen must both use that runtime so **evaluator ↔ export** stays aligned.

**Excel-facing tests:** Prefer validating against **live Excel** (xlwings on Windows/macOS, or Excel via COM from WSL) when comparing to the real engine. **Run-if-available:** if automation is missing, **`pytest.skip`** with a clear reason—do not fail CI. Cache-based comparisons (`excel_workbook_parity`) remain useful for environments without Excel. See `.cursor/rules/parity.mdc` for the full contract.

## Cursor Cloud specific instructions

This is a pure Python **library + CLI** (`excel-grapher`) managed with `uv`; there are no runtime services (no DB/web/queue). "End-to-end" runs entirely in-process against `.xlsx` fixtures in `examples/` and `tests/fixtures/`. Dependencies are installed by the startup update script (`uv sync --all-extras --dev`), which also provisions the pinned Python (3.13). Run everything through `uv run`.

- Standard lint/format/type/test commands live in `.github/workflows/ci.yml`: `uv run ruff check .`, `uv run ruff format --check .`, `uv run ty check`, `uv run pytest`.
- `pytest` deselects the `slow` marker by default (see `[tool.pytest.ini_options]`); opt in with `uv run pytest -m slow`.
- Live-Excel parity tests (`xlwings`) `pytest.skip` on this Linux VM since Excel automation is unavailable — this is expected, not a failure.
- The `graphviz` Python package is installed, but the system Graphviz binary is not; rendering visualizations to images needs `apt-get install -y graphviz` (only for viz/docs, not for tests or core use).
- CLI smoke check: `uv run excel-grapher bindings validate examples/micro_workbooks/ffv2.xlsx --bindings examples/micro_workbooks/ff.bindings.yaml --smoke-test` (note the colocated sidecar declares `workbook: ff.xlsx`, so pass the workbook explicitly as shown).