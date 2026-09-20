#!/usr/bin/env bash
set -euo pipefail
exec uv run excel-grapher bindings burndown "$@"
