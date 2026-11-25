"""WordCombine: High-fidelity Word (.docx) merger with silent TOC refresh."""

from __future__ import annotations

from importlib import import_module
from typing import Any, Dict, Tuple

__version__ = "1.0.1"

_LAZY_EXPORTS: Dict[str, Tuple[str, str]] = {
    "collect_docx_files": (".sorting", "collect_docx_files"),
    "natural_sort_key": (".sorting", "natural_sort_key"),
    "sort_docx_paths": (".sorting", "sort_docx_paths"),
    "MergeResult": (".ooxml_merger", "MergeResult"),
    "TOCInfo": (".ooxml_merger", "TOCInfo"),
    "detect_tocs": (".ooxml_merger", "detect_tocs"),
    "merge_docx_files": (".ooxml_merger", "merge_docx_files"),
    "RefreshResult": (".toc_refresher", "RefreshResult"),
    "inspect_word_open_and_headings": (".toc_refresher", "inspect_word_open_and_headings"),
    "refresh_toc_silent": (".toc_refresher", "refresh_toc_silent"),
    "VerificationReport": (".verifier", "VerificationReport"),
    "verify_merged_docx": (".verifier", "verify_merged_docx"),
}

__all__ = [
    "__version__",
    "MergeResult",
    "RefreshResult",
    "TOCInfo",
    "VerificationReport",
    "collect_docx_files",
    "detect_tocs",
    "inspect_word_open_and_headings",
    "merge_docx_files",
    "natural_sort_key",
    "refresh_toc_silent",
    "sort_docx_paths",
    "verify_merged_docx",
]


def __getattr__(name: str) -> Any:
    target = _LAZY_EXPORTS.get(name)
    if target is None:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    mod_name, attr_name = target
    module = import_module(mod_name, __name__)
    value = getattr(module, attr_name)
    globals()[name] = value
    return value


def __dir__() -> list[str]:
    return sorted(__all__)
