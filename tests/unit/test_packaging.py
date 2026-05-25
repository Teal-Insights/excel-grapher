"""Guards for release artifacts: examples stay out of install bundles."""

from __future__ import annotations

import subprocess
import sys
import tarfile
import zipfile
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[2]


def _example_paths(member_names: list[str]) -> list[str]:
    return [
        name
        for name in member_names
        if name.startswith("examples/") or "/examples/" in name or name.endswith("/examples")
    ]


def _build_artifacts(out_dir: Path) -> tuple[Path, Path]:
    subprocess.run(
        ["uv", "build", "--out-dir", str(out_dir), "--clear"],
        cwd=REPO_ROOT,
        check=True,
        capture_output=True,
        text=True,
    )
    wheels = sorted(out_dir.glob("*.whl"))
    sdists = sorted(out_dir.glob("*.tar.gz"))
    assert len(wheels) == 1, wheels
    assert len(sdists) == 1, sdists
    return wheels[0], sdists[0]


@pytest.fixture(scope="module")
def built_artifacts(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], list[str]]:
    build_dir = tmp_path_factory.mktemp("packaging-build")
    wheel_path, sdist_path = _build_artifacts(build_dir)
    with zipfile.ZipFile(wheel_path) as zf:
        wheel_members = zf.namelist()
    with tarfile.open(sdist_path, "r:gz") as tf:
        sdist_members = [m.name for m in tf.getmembers() if m.isfile()]
    return wheel_members, sdist_members


def test_wheel_does_not_ship_examples(built_artifacts: tuple[list[str], list[str]]) -> None:
    wheel_members, _ = built_artifacts
    offenders = _example_paths(wheel_members)
    assert not offenders, f"wheel must not bundle examples/: {offenders[:20]}"


def test_sdist_does_not_ship_examples(built_artifacts: tuple[list[str], list[str]]) -> None:
    _, sdist_members = built_artifacts
    offenders = _example_paths(sdist_members)
    assert not offenders, f"sdist must not bundle examples/: {offenders[:20]}"


def test_installed_wheel_does_not_expose_examples_module(tmp_path: Path) -> None:
    build_dir = tmp_path / "dist"
    wheel_path, _ = _build_artifacts(build_dir)

    venv_dir = tmp_path / "venv"
    subprocess.run([sys.executable, "-m", "venv", str(venv_dir)], check=True)
    pip = venv_dir / "bin" / "pip"
    python = venv_dir / "bin" / "python"
    subprocess.run(
        [str(pip), "install", "--no-deps", str(wheel_path)],
        check=True,
        capture_output=True,
        text=True,
    )

    result = subprocess.run(
        [str(python), "-I", "-c", "import examples"],
        cwd=tmp_path,
        capture_output=True,
        text=True,
    )
    assert result.returncode != 0, result.stdout + result.stderr
