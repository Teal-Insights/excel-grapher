# excel-grapher-author-bindings

Agent skill for authoring excel-grapher series-binding sidecar YAML.

This package is a **separate distribution** from `excel-grapher`. The library
wheel does not include the skill files.

## Install

From this repository:

```bash
uv add "excel-grapher-author-bindings @ git+https://github.com/Teal-Insights/excel-grapher#subdirectory=packages/excel-grapher-author-bindings"
```

Or with the optional extra on a source checkout of excel-grapher:

```bash
uv add excel-grapher --extra skills
```

Copy the skill into an agent skill path:

```python
from excel_grapher_author_bindings import author_bindings_skill_dir

with author_bindings_skill_dir() as src:
    # copy `src` to `.cursor/skills/author-bindings` or Claude's skill path
    print(src)
```

You can also copy `skills/author-bindings` from a git clone or from the
`excel-grapher` source distribution (sdist).
