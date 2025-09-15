#!/usr/bin/env python3
"""Basic environment check for macOS users."""
import platform
import shutil
from pathlib import Path

def main() -> int:
    if platform.system() != "Darwin":
        print("This script is intended for macOS environments.")
        return 0

    if not shutil.which("python3"):
        print("python3 not found. Install it with 'brew install python'.")
        return 1

    venv = Path('.venv')
    if not venv.exists():
        print("No virtual environment detected. Create one with:\\n  python3 -m venv .venv && source .venv/bin/activate")
    else:
        print("Virtual environment '.venv' detected.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
