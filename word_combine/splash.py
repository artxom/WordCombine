"""Safe runtime wrappers for PyInstaller's optional splash screen (`pyi_splash`)."""

from __future__ import annotations

import os
import sys


def _splash_available() -> bool:
    """Only import `pyi_splash` when PyInstaller's bootloader actually started an IPC server."""
    if "pyi_splash" in sys.modules:
        return True
    ipc_port = os.environ.get("_PYI_SPLASH_IPC", "")
    return bool(ipc_port) and ipc_port != "0"


def update_splash(text: str) -> None:
    """Update the PyInstaller splash screen status text if active; no-op otherwise."""
    if not _splash_available():
        return
    try:
        import pyi_splash  # type: ignore[import-not-found]

        if pyi_splash.is_alive():
            pyi_splash.update_text(text)
    except Exception:
        pass


def close_splash() -> None:
    """Close the PyInstaller splash screen if active; no-op otherwise."""
    if not _splash_available():
        return
    try:
        import pyi_splash  # type: ignore[import-not-found]

        pyi_splash.close()
    except Exception:
        pass
