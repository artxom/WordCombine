"""WordCombine: High-fidelity Word (.docx) merger with silent TOC refresh."""

__version__ = "1.0.0"

from .sorting import collect_docx_files, natural_sort_key, sort_docx_paths
from .ooxml_merger import (
    MergeResult,
    TOCInfo,
    detect_tocs,
    merge_docx_files,
)
from .toc_refresher import RefreshResult, inspect_word_open_and_headings, refresh_toc_silent
from .verifier import VerificationReport, verify_merged_docx

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
