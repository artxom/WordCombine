#!/usr/bin/env python3
"""Entry point for WordCombine (launches GUI with no args, CLI with args)."""

from __future__ import annotations

import os
import sys

from word_combine.splash import close_splash, update_splash


def _ensure_stdio_for_windowed_cli() -> None:
    """Guard against None or non-UTF8 stdio streams in PyInstaller --windowed builds."""
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    elif hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")
    elif hasattr(sys.stderr, "reconfigure"):
        try:
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def main() -> int:
    if len(sys.argv) <= 1:
        update_splash("Loading interface...")
        from word_combine.gui import launch_gui

        launch_gui()
        return 0

    close_splash()
    _ensure_stdio_for_windowed_cli()
    from word_combine.cli import main as cli_main

    try:
        return cli_main(sys.argv[1:])
    except SystemExit:
        raise
    except Exception as exc:
        print(f"错误: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
