# excel-grapher-author-bindings

Agent skill for authoring excel-grapher series-binding sidecar YAML.

This package is a **file source** for the skill tree. It is not how agents
discover skills. The `excel-grapher` library wheel does not include the skill
files.

## Install

Copy the skill bundle into an agent skill directory. The folder name must match
the skill (`author-bindings`) and must contain `SKILL.md`:

```bash
mkdir -p .agents/skills
cp -R skills/author-bindings .agents/skills/author-bindings
```

Cursor also loads `.cursor/skills/author-bindings`. Claude Code uses
`.claude/skills/author-bindings`. User-level installs go under
`~/.agents/skills/author-bindings`.

You can copy from a git clone, from the `excel-grapher` source distribution
(sdist), or from this package after installing it as a file source:

```bash
uv add "excel-grapher-author-bindings @ git+https://github.com/Teal-Insights/excel-grapher#subdirectory=packages/excel-grapher-author-bindings"
```

```python
from excel_grapher_author_bindings import author_bindings_skill_dir

with author_bindings_skill_dir() as src:
    # copy `src` to `.agents/skills/author-bindings`
    print(src)
```
