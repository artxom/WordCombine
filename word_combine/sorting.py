"""Natural filename sorting and .docx discovery utilities."""

from __future__ import annotations

import os
import re
import unicodedata
from pathlib import Path
from typing import Iterable, List, Optional, Sequence, Tuple, Union

_SPLIT_RE = re.compile(r"(\d+)")


def natural_sort_key(path: Union[str, os.PathLike[str]]) -> Tuple[Tuple[Tuple[int, Union[int, str]], ...], str]:
    """Return a natural sort key based on the NFC-normalized basename.

    Numeric segments sort numerically `(0, int(tok))` and text segments sort
    case-insensitively `(1, tok.lower())`. Thus:
        3.1 3.2 ... < 3.3 ... < ... < 3.9 ... < 3.10 ... < 4. ...
    """
    basename = unicodedata.normalize("NFC", os.path.basename(os.fspath(path)))
    parts: List[Tuple[int, Union[int, str]]] = []
    for token in _SPLIT_RE.split(basename):
        if not token:
            continue
        if token.isdigit():
            parts.append((0, int(token)))
        else:
            parts.append((1, token.lower()))
    return (tuple(parts), basename)


def sort_docx_paths(paths: Iterable[Union[str, os.PathLike[str]]]) -> List[str]:
    """Return deduplicated paths (preserving order on duplicates) sorted naturally by filename."""
    seen = set()
    unique: List[str] = []
    for p in paths:
        s = os.fspath(p)
        rp = os.path.realpath(s)
        if rp in seen:
            continue
        seen.add(rp)
        unique.append(s)
    return sorted(unique, key=natural_sort_key)


def _is_valid_docx_file(path: Path) -> bool:
    name = path.name
    if not path.is_file():
        return False
    if name.startswith("~$") or name.startswith("."):
        return False
    return path.suffix.lower() == ".docx"


def collect_docx_files(
    paths_or_dir: Union[str, os.PathLike[str], Sequence[Union[str, os.PathLike[str]]]],
    *,
    exclude_paths: Optional[Iterable[Union[str, os.PathLike[str]]]] = None,
    sort: bool = True,
) -> List[str]:
    """Collect .docx files from a directory or a sequence of file/dir paths.

    Ignores Word temporary lock files (`~$*.docx`), hidden files, and any paths
    matching `exclude_paths` (compared via resolved realpath).
    """
    excluded = {os.path.realpath(os.fspath(p)) for p in (exclude_paths or ())}

    if isinstance(paths_or_dir, (str, os.PathLike)):
        inputs: Sequence[Union[str, os.PathLike[str]]] = [paths_or_dir]
    else:
        inputs = paths_or_dir

    collected: List[str] = []
    seen = set()

    def _add_file(p: Path) -> None:
        if not _is_valid_docx_file(p):
            return
        rp = os.path.realpath(p)
        if rp in excluded or rp in seen:
            return
        seen.add(rp)
        collected.append(str(p))

    for item in inputs:
        p = Path(os.fspath(item))
        if p.is_dir():
            for child in p.iterdir():
                _add_file(child)
        elif p.is_file():
            _add_file(p)
        else:
            raise FileNotFoundError(f"Path not found: {p}")

    if sort:
        return sorted(collected, key=natural_sort_key)
    return collected
