"""Utility helpers for working with filesystem paths.
Ensures consistent normalization across platforms and attempts to
handle case-insensitive lookups on macOS.
"""
from pathlib import Path
from typing import Union


def normalize_path(path: Union[str, Path]) -> Path:
    """Return a normalized absolute Path.

    On case-sensitive filesystems this function will attempt a
    case-insensitive match so that users with mis-cased paths on
    macOS do not encounter file-not-found errors.
    """
    p = Path(path).expanduser()
    if p.exists():
        return p.resolve()

    # Attempt case-insensitive search within parent directory
    parent = p.parent.expanduser()
    if parent.exists():
        lower_name = p.name.lower()
        for child in parent.iterdir():
            if child.name.lower() == lower_name:
                return child.resolve()

    # Fall back to resolved path without checking existence
    try:
        return p.resolve()
    except FileNotFoundError:
        return p
