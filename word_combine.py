#!/usr/bin/env python3
"""Entry point for WordCombine (launches GUI with no args, CLI with args)."""

from __future__ import annotations

import multiprocessing
import os
import sys


def _ensure_stdio_for_windowed_cli() -> None:
    """Guard against sys.stdout / sys.stderr being None in PyInstaller --windowed builds."""
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")


def main() -> int:
    multiprocessing.freeze_support()
    if len(sys.argv) <= 1:
        from word_combine.gui import launch_gui

        launch_gui()
        return 0

    _ensure_stdio_for_windowed_cli()
    from word_combine.cli import main as cli_main

    return cli_main(sys.argv[1:])


if __name__ == "__main__":
    raise SystemExit(main())
