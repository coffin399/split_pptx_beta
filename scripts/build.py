#!/usr/bin/env python3
"""Build the GUI application into distributable packages using Nuitka."""

from __future__ import annotations

import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
APP_ENTRY = PROJECT_ROOT / "app.py"
# Ensure dist directory exists ahead of builds
DIST_ROOT = PROJECT_ROOT / "dist"
DIST_ROOT.mkdir(exist_ok=True)
OUTPUT_BASENAME = "pptx_script_slides"
MAC_APP_NAME = "PPTXScriptSlides"


def run(cmd: list[str]) -> None:
    print("Running:", " ".join(cmd))
    subprocess.run(cmd, check=True)


def build_with_nuitka(target_dir: Path) -> None:
    target_dir.mkdir(parents=True, exist_ok=True)
    system = platform.system()
    base_cmd: list[str] = [
        sys.executable,
        "-m",
        "nuitka",
        str(APP_ENTRY),
        "--enable-plugin=pyside6",
        "--include-qt-plugins=sensible",
        f"--output-dir={target_dir}",
        "--follow-imports",
        "--assume-yes-for-downloads",
    ]
    if system == "Windows":
        base_cmd.extend(
            [
                "--onefile",
                f"--output-filename={OUTPUT_BASENAME}",
                "--windows-disable-console",
            ]
        )
    elif system == "Darwin":
        base_cmd.extend(
            [
                "--standalone",
                f"--output-filename={OUTPUT_BASENAME}",
                "--macos-create-app-bundle",
                f"--macos-app-name={MAC_APP_NAME}",
            ]
        )
    else:
        base_cmd.extend([
            "--onefile",
            f"--output-filename={OUTPUT_BASENAME}",
        ])
    run(base_cmd)


def _find_macos_app(target_dir: Path) -> Path:
    preferred = target_dir / f"{MAC_APP_NAME}.app"
    if preferred.exists():
        return preferred
    candidates = sorted(target_dir.glob("**/*.app"))
    if not candidates:
        raise FileNotFoundError(preferred)
    # Prefer bundle matching requested name if present among candidates
    for candidate in candidates:
        if candidate.name == f"{MAC_APP_NAME}.app":
            return candidate
    return candidates[0]


def package_artifact(target_dir: Path) -> Path:
    system = platform.system()
    if system == "Windows":
        binary = target_dir / f"{OUTPUT_BASENAME}.exe"
        if not binary.exists():
            raise FileNotFoundError(binary)
        return binary
    elif system == "Darwin":
        app_bundle = _find_macos_app(target_dir)
        archive_base = target_dir / f"{OUTPUT_BASENAME}-macos"
        archive_path = Path(
            shutil.make_archive(
                str(archive_base),
                "zip",
                root_dir=app_bundle.parent,
                base_dir=app_bundle.name,
            )
        )
        return archive_path
    else:
        binary = target_dir / OUTPUT_BASENAME
        if not binary.exists():
            raise FileNotFoundError(binary)
        return binary


def main() -> None:
    system = platform.system().lower()
    target_dir = DIST_ROOT / system
    build_with_nuitka(target_dir)
    artifact = package_artifact(target_dir)
    print(f"Artifact created: {artifact}")
    github_output = os.getenv("GITHUB_OUTPUT")
    if github_output:
        with Path(github_output).open("a", encoding="utf-8") as fh:
            fh.write(f"artifact={artifact}\n")


if __name__ == "__main__":
    main()
