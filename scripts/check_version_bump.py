"""Verify the PR version in pyproject.toml is greater than the base branch."""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
import tomllib
from pathlib import Path
from typing import TypeAlias

VersionTuple: TypeAlias = tuple[int, int, int]
_VERSION_RE = re.compile(r"^(\d+)\.(\d+)\.(\d+)")
_ZERO_SHA = "0" * 40
_RELEASE_PATHS = ("excel_grapher", "pyproject.toml")
_MAIN_REFS = ("origin/main", "main")


def parse_version(version: str) -> VersionTuple:
    """Parse a simple ``major.minor.patch`` version string.

    Args:
        version: Version text from ``pyproject.toml``.

    Returns:
        A ``(major, minor, patch)`` tuple.

    Raises:
        ValueError: If ``version`` is not a supported ``x.y.z`` value.
    """
    match = _VERSION_RE.match(version.strip())
    if match is None:
        raise ValueError(f"unsupported version: {version}")
    major, minor, patch = match.groups()
    return int(major), int(minor), int(patch)


def version_from_pyproject(path: Path) -> VersionTuple:
    """Read ``project.version`` from a ``pyproject.toml`` file.

    Args:
        path: Path to ``pyproject.toml``.

    Returns:
        Parsed project version.
    """
    data = tomllib.loads(path.read_text(encoding="utf-8"))
    return parse_version(str(data["project"]["version"]))


def version_from_git_ref(
    ref: str,
    path: Path = Path("pyproject.toml"),
    cwd: Path | None = None,
) -> VersionTuple:
    """Read ``project.version`` from ``pyproject.toml`` at a git ref.

    Args:
        ref: Git ref or commit, for example ``origin/main``.
        path: Repository-relative path to ``pyproject.toml``.
        cwd: Optional repository root for git commands.

    Returns:
        Parsed project version at ``ref``.

    Raises:
        SystemExit: If git cannot resolve ``ref`` or the file is missing.
    """
    result = subprocess.run(
        ["git", "show", f"{ref}:{path.as_posix()}"],
        check=False,
        capture_output=True,
        text=True,
        cwd=cwd,
    )
    if result.returncode != 0:
        print(
            f"failed to read {path} at {ref}: {result.stderr.strip()}",
            file=sys.stderr,
        )
        raise SystemExit(1)
    data = tomllib.loads(result.stdout)
    return parse_version(str(data["project"]["version"]))


def version_is_greater_than(head: VersionTuple, base: VersionTuple) -> bool:
    """Return whether ``head`` is strictly greater than ``base``."""
    return head > base


def format_version(version: VersionTuple) -> str:
    """Format a version tuple as ``major.minor.patch``."""
    return ".".join(str(part) for part in version)


def resolve_main_ref(cwd: Path) -> str | None:
    """Return the first resolvable main-branch ref.

    Args:
        cwd: Repository root for git commands.

    Returns:
        A ref such as ``origin/main``, or ``None`` if none resolve.
    """
    for ref in _MAIN_REFS:
        result = subprocess.run(
            ["git", "rev-parse", "--verify", ref],
            check=False,
            capture_output=True,
            text=True,
            cwd=cwd,
        )
        if result.returncode == 0:
            return ref
    return None


def push_diff_range(remote_sha: str, local_sha: str, cwd: Path) -> str:
    """Build a git revision range for commits being pushed.

    Args:
        remote_sha: Remote tip before the push, or the all-zero OID for a new ref.
        local_sha: Local tip being pushed.
        cwd: Repository root for git commands.

    Returns:
        A revision range suitable for ``git diff`` / ``git log``.
    """
    if remote_sha == _ZERO_SHA:
        main_ref = resolve_main_ref(cwd)
        if main_ref is None:
            return local_sha
        merge_base = subprocess.run(
            ["git", "merge-base", main_ref, local_sha],
            check=False,
            capture_output=True,
            text=True,
            cwd=cwd,
        )
        if merge_base.returncode == 0 and merge_base.stdout.strip():
            return f"{merge_base.stdout.strip()}..{local_sha}"
        return local_sha
    return f"{remote_sha}..{local_sha}"


def push_includes_release_paths(remote_sha: str, local_sha: str, cwd: Path) -> bool:
    """Return whether a push updates package source or ``pyproject.toml``.

    Args:
        remote_sha: Remote tip before the push, or the all-zero OID for a new ref.
        local_sha: Local tip being pushed.
        cwd: Repository root for git commands.

    Returns:
        ``True`` when the push range touches ``excel_grapher/`` or ``pyproject.toml``.
    """
    range_expr = push_diff_range(remote_sha, local_sha, cwd)
    result = subprocess.run(
        [
            "git",
            "diff",
            "--name-only",
            range_expr,
            "--",
            *_RELEASE_PATHS,
        ],
        check=False,
        capture_output=True,
        text=True,
        cwd=cwd,
    )
    if result.returncode != 0:
        return False
    return bool(result.stdout.strip())


def read_pre_push_updates() -> list[tuple[str, str, str, str]]:
    """Parse git pre-push stdin lines.

    Returns:
        Tuples of ``(local_ref, local_sha, remote_ref, remote_sha)``.
    """
    updates: list[tuple[str, str, str, str]] = []
    for line in sys.stdin:
        stripped = line.strip()
        if not stripped:
            continue
        parts = stripped.split()
        if len(parts) != 4:
            continue
        local_ref, local_sha, remote_ref, remote_sha = parts
        updates.append((local_ref, local_sha, remote_ref, remote_sha))
    return updates


def emit_version_bump_message(
    *,
    head_version: VersionTuple,
    base_version: VersionTuple,
    base_ref: str,
    warn_only: bool,
) -> None:
    """Print a version-bump failure or reminder message.

    Args:
        head_version: Version on the branch being checked.
        base_version: Version on the base ref.
        base_ref: Git ref used for the base version.
        warn_only: When ``True``, print a reminder to stdout instead of an error.
    """
    head_text = format_version(head_version)
    base_text = format_version(base_version)
    if warn_only:
        print(
            "version bump reminder: pushing changes under excel_grapher/ or "
            "pyproject.toml, but pyproject.toml version "
            f"({head_text}) is not greater than {base_ref} ({base_text}).\n"
            "Bump the version before opening a PR — CI will reject equal versions."
        )
        return

    print(
        "pyproject.toml version must be greater than the base branch.\n"
        f"  base ({base_ref}): {base_text}\n"
        f"  head:               {head_text}",
        file=sys.stderr,
    )


def run_pre_push(warn_only: bool, cwd: Path, pyproject: Path) -> None:
    """Check pushed commits for a required version bump.

    Args:
        warn_only: Print a reminder and always exit successfully.
        cwd: Repository root.
        pyproject: Repository-relative path to ``pyproject.toml``.

    Raises:
        SystemExit: When ``warn_only`` is false and the version was not bumped.
    """
    updates = read_pre_push_updates()
    if not updates:
        return

    qualifying = [
        (local_sha, remote_sha)
        for _, local_sha, _, remote_sha in updates
        if push_includes_release_paths(remote_sha, local_sha, cwd)
    ]
    if not qualifying:
        return

    main_ref = resolve_main_ref(cwd)
    if main_ref is None:
        return

    local_sha = qualifying[-1][0]
    head_version = version_from_git_ref(local_sha, pyproject, cwd=cwd)
    base_version = version_from_git_ref(main_ref, pyproject, cwd=cwd)

    if version_is_greater_than(head_version, base_version):
        return

    emit_version_bump_message(
        head_version=head_version,
        base_version=base_version,
        base_ref=main_ref,
        warn_only=warn_only,
    )
    if not warn_only:
        raise SystemExit(1)


def main(argv: list[str] | None = None, cwd: Path | None = None) -> None:
    """Compare versions for CI or git pre-push hooks.

    Args:
        argv: Optional CLI arguments. Defaults to ``sys.argv[1:]``.
        cwd: Optional repository root. Defaults to the current directory.

    Raises:
        SystemExit: If the head version is not greater than the base version.
    """
    parser = argparse.ArgumentParser(
        description="Require pyproject.toml version to be greater than a base git ref.",
    )
    parser.add_argument(
        "--base",
        help="Git ref for the base branch (for example origin/main).",
    )
    parser.add_argument(
        "--pre-push",
        action="store_true",
        help="Read refs from git pre-push stdin instead of using the working tree.",
    )
    parser.add_argument(
        "--warn-only",
        action="store_true",
        help="Print a reminder instead of failing (for local pre-push hooks).",
    )
    parser.add_argument(
        "--pyproject",
        type=Path,
        default=Path("pyproject.toml"),
        help="Repository-relative path to pyproject.toml.",
    )
    args = parser.parse_args(argv)

    repo_root = cwd or Path.cwd()

    if args.pre_push:
        run_pre_push(warn_only=args.warn_only, cwd=repo_root, pyproject=args.pyproject)
        return

    if args.base is None:
        parser.error("--base is required unless --pre-push is set")

    pyproject_path = repo_root / args.pyproject
    head_version = version_from_pyproject(pyproject_path)
    base_version = version_from_git_ref(args.base, args.pyproject, cwd=repo_root)

    if version_is_greater_than(head_version, base_version):
        return

    emit_version_bump_message(
        head_version=head_version,
        base_version=base_version,
        base_ref=args.base,
        warn_only=args.warn_only,
    )
    if not args.warn_only:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
