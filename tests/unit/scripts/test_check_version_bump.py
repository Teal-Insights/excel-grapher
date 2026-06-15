from __future__ import annotations

import subprocess
from pathlib import Path

import pytest

from scripts.check_version_bump import (
    main,
    parse_version,
    push_includes_release_paths,
    run_pre_push,
    version_from_pyproject,
    version_is_greater_than,
)


def _write_pyproject(path: Path, version: str) -> None:
    path.write_text(
        f'[project]\nname = "excel-grapher"\nversion = "{version}"\n',
        encoding="utf-8",
    )


def _init_git_repo(repo: Path, version: str, branch: str = "main") -> str:
    _write_pyproject(repo / "pyproject.toml", version)
    subprocess.run(["git", "init"], cwd=repo, check=True, capture_output=True)
    subprocess.run(["git", "config", "user.email", "test@example.com"], cwd=repo, check=True)
    subprocess.run(["git", "config", "user.name", "test"], cwd=repo, check=True)
    subprocess.run(["git", "add", "pyproject.toml"], cwd=repo, check=True)
    subprocess.run(["git", "commit", "-m", "init"], cwd=repo, check=True)
    subprocess.run(["git", "branch", "-M", branch], cwd=repo, check=True)
    return subprocess.run(
        ["git", "rev-parse", "HEAD"],
        cwd=repo,
        check=True,
        capture_output=True,
        text=True,
    ).stdout.strip()


def _commit_file(repo: Path, relative_path: str, content: str, message: str) -> str:
    path = repo / relative_path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
    subprocess.run(["git", "add", relative_path], cwd=repo, check=True)
    subprocess.run(["git", "commit", "-m", message], cwd=repo, check=True)
    return subprocess.run(
        ["git", "rev-parse", "HEAD"],
        cwd=repo,
        check=True,
        capture_output=True,
        text=True,
    ).stdout.strip()


def test_parse_version_parses_semver() -> None:
    assert parse_version("0.1.0") == (0, 1, 0)
    assert parse_version("1.2.3") == (1, 2, 3)


def test_parse_version_rejects_invalid() -> None:
    with pytest.raises(ValueError, match="unsupported version"):
        parse_version("not-a-version")


def test_version_from_pyproject_reads_project_version(tmp_path: Path) -> None:
    pyproject = tmp_path / "pyproject.toml"
    _write_pyproject(pyproject, "0.2.1")

    assert version_from_pyproject(pyproject) == (0, 2, 1)


def test_version_is_greater_than() -> None:
    assert version_is_greater_than((0, 2, 0), (0, 1, 9))
    assert not version_is_greater_than((0, 1, 0), (0, 1, 0))
    assert not version_is_greater_than((0, 1, 0), (0, 2, 0))


def test_main_passes_when_head_version_is_greater(tmp_path: Path) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    _init_git_repo(repo, "0.1.0")
    _write_pyproject(repo / "pyproject.toml", "0.2.0")

    main(["--base", "main"], cwd=repo)


def test_main_fails_when_head_version_is_equal(tmp_path: Path) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    _init_git_repo(repo, "0.1.0")

    with pytest.raises(SystemExit) as exc_info:
        main(["--base", "main"], cwd=repo)

    assert exc_info.value.code == 1


def test_main_fails_when_head_version_is_lower(tmp_path: Path) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    _init_git_repo(repo, "0.2.0")
    _write_pyproject(repo / "pyproject.toml", "0.1.0")

    with pytest.raises(SystemExit) as exc_info:
        main(["--base", "main"], cwd=repo)

    assert exc_info.value.code == 1


def test_main_warn_only_does_not_fail_when_version_is_equal(tmp_path: Path) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    _init_git_repo(repo, "0.1.0")

    main(["--base", "main", "--warn-only"], cwd=repo)


def test_push_includes_release_paths_detects_excel_grapher_changes(tmp_path: Path) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    base_sha = _init_git_repo(repo, "0.1.0")
    head_sha = _commit_file(
        repo,
        "excel_grapher/example.py",
        "VALUE = 1\n",
        "add module",
    )

    assert push_includes_release_paths(base_sha, head_sha, repo)
    assert not push_includes_release_paths(head_sha, head_sha, repo)


def test_run_pre_push_warn_only_prints_reminder(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    _init_git_repo(repo, "0.1.0")
    head_sha = _commit_file(
        repo,
        "excel_grapher/example.py",
        "VALUE = 2\n",
        "change module",
    )
    zero_sha = "0" * 40
    monkeypatch.setattr(
        "scripts.check_version_bump.read_pre_push_updates",
        lambda: [("refs/heads/feature", head_sha, "refs/heads/feature", zero_sha)],
    )

    run_pre_push(warn_only=True, cwd=repo, pyproject=Path("pyproject.toml"))


def test_run_pre_push_skips_when_push_has_no_release_paths(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    _init_git_repo(repo, "0.1.0")
    head_sha = _commit_file(repo, "README.md", "docs\n", "docs only")
    zero_sha = "0" * 40
    monkeypatch.setattr(
        "scripts.check_version_bump.read_pre_push_updates",
        lambda: [("refs/heads/feature", head_sha, "refs/heads/feature", zero_sha)],
    )

    run_pre_push(warn_only=True, cwd=repo, pyproject=Path("pyproject.toml"))
