#!/usr/bin/env python3
"""Runs deck2pptx from a private virtual environment, creating it on first use.

Standard library only, so it works before anything is installed:

    python scripts/run.py deck.html --verify

--data-dir chooses where the environment lives. It has to be an argument because a
plugin's data path is expanded in the skill's text and is not in the shell's environment.
"""
from __future__ import annotations

import hashlib
import os
import subprocess
import sys
import venv
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
REQUIREMENTS = ROOT / "requirements.txt"
MIN_PYTHON = (3, 10)


def default_data_dir() -> Path:
    if sys.platform == "win32":
        return Path(os.environ.get("LOCALAPPDATA", Path.home())) / "deck2pptx"
    return Path(os.environ.get("XDG_CACHE_HOME", Path.home() / ".cache")) / "deck2pptx"


def interpreter(env: Path) -> Path:
    return env / ("Scripts/python.exe" if sys.platform == "win32" else "bin/python")


def ensure_env(data: Path) -> Path:
    stamp_value = hashlib.sha1(REQUIREMENTS.read_bytes()).hexdigest()
    env = data / "venv"
    stamp = env / ".requirements"
    if interpreter(env).exists() and stamp.exists() and stamp.read_text() == stamp_value:
        return interpreter(env)
    print("First run: installing deck2pptx's dependencies...", file=sys.stderr)
    venv.EnvBuilder(with_pip=True, clear=env.exists()).create(env)
    python = interpreter(env)
    subprocess.run([python, "-m", "pip", "install", "--quiet", "--disable-pip-version-check", "-r", str(REQUIREMENTS)], check=True)
    stamp.write_text(stamp_value)
    return python


def main() -> int:
    if sys.version_info < MIN_PYTHON:
        print(f"deck2pptx needs Python {MIN_PYTHON[0]}.{MIN_PYTHON[1]} or newer.", file=sys.stderr)
        return 2
    args = sys.argv[1:]
    data = default_data_dir()
    if args[:1] == ["--data-dir"]:
        # An unexpanded variable means this is not running as a plugin.
        if len(args) > 1 and args[1] and not args[1].startswith("${"):
            data = Path(args[1])
        args = args[2:]
    python = ensure_env(data)
    env = dict(os.environ, PYTHONPATH=str(ROOT / "src"), PYTHONIOENCODING="utf-8")
    return subprocess.run([python, "-m", "deck2pptx", *args], env=env).returncode


if __name__ == "__main__":
    raise SystemExit(main())
