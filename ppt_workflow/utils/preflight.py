"""Pre-flight checks for macOS environments."""
import platform
import shutil
from pathlib import Path


class PreflightError(RuntimeError):
    """Raised when required dependencies are missing."""


def run_mac_checks() -> None:
    """Ensure required tools are available on macOS.

    Checks for Homebrew `python3`, Microsoft PowerPoint, and the
    `osascript` command-line tool used for AppleScript automation.
    Raises :class:`PreflightError` with a helpful message if something is
    missing. No-op on non-macOS platforms.
    """
    if platform.system() != "Darwin":
        return

    missing = []
    if not shutil.which("python3"):
        missing.append("python3 (install via Homebrew: brew install python)")

    if not shutil.which("osascript"):
        missing.append("osascript (built-in, required for AppleScript automation)")

    if not Path("/Applications/Microsoft PowerPoint.app").exists():
        missing.append("Microsoft PowerPoint (install via Microsoft 365)")

    if missing:
        raise PreflightError(
            "Missing macOS dependencies: " + ", ".join(missing)
        )
