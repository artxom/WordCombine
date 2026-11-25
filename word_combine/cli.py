"""Command-line interface for WordCombine."""

from __future__ import annotations

import argparse
import os
import sys
from typing import List, Optional, Sequence

from .ooxml_merger import merge_docx_files
from .sorting import collect_docx_files, sort_docx_paths
from .toc_refresher import refresh_toc_silent
from .verifier import verify_merged_docx


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="word_combine",
        description="按文件名自然顺序高保真拼接多个 Word (.docx) 文档，并静默刷新目录。",
    )
    parser.add_argument(
        "--files",
        nargs="+",
        help="要拼接的 .docx 文件列表（可指定多个）。",
    )
    parser.add_argument(
        "--dir",
        help="扫描该目录下的所有 .docx 文件（自动忽略 ~$ 临时文件及输出文件本身）。",
    )
    parser.add_argument(
        "-o",
        "--output",
        required=True,
        help="输出的合并 .docx 文件路径。",
    )
    parser.add_argument(
        "--section-break",
        choices=("auto", "page", "continuous"),
        default="auto",
        help="文档间分节策略：auto=布局/页眉一致则无缝首尾相连，不同则分节；page=每篇换页分节；continuous=连续分节。(默认: auto)",
    )
    parser.add_argument(
        "--no-sort",
        action="store_true",
        help="保持传入顺序，不按文件名自然排序。",
    )
    parser.add_argument(
        "--no-refresh-toc",
        action="store_true",
        help="拼接后不调用 Microsoft Word 静默刷新目录。",
    )
    parser.add_argument(
        "--insert-toc",
        action="store_true",
        help="在合并后的文档开头自动插入 Word 目录（TOC）及分页符。",
    )
    parser.add_argument(
        "--toc-levels",
        default="1-3",
        help="使用 --insert-toc 时目录包含的标题级别范围，默认 1-3。",
    )
    parser.add_argument(
        "--verify",
        action="store_true",
        help="合并后执行严格的 XML/样式/段落/表格/图片/OLE/批注/页眉及目录一致性校验。",
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if not args.files and not args.dir:
        parser.error("必须提供 --files 或 --dir 参数。")

    out_abs = os.path.abspath(args.output)
    raw_inputs: List[str] = []
    if args.files:
        raw_inputs.extend(
            collect_docx_files(args.files, exclude_paths=[out_abs], sort=False)
        )
    if args.dir:
        raw_inputs.extend(
            collect_docx_files(args.dir, exclude_paths=[out_abs], sort=False)
        )

    # Deduplicate preserving order, then sort unless --no-sort
    deduped: List[str] = []
    seen = set()
    for p in raw_inputs:
        rp = os.path.realpath(p)
        if rp not in seen:
            seen.add(rp)
            deduped.append(p)

    if not deduped:
        print("错误：未找到任何可拼接的 .docx 文件。", file=sys.stderr)
        return 2

    ordered = deduped if args.no_sort else sort_docx_paths(deduped)

    print(f"共选择 {len(ordered)} 个文档，拼接顺序如下：")
    for idx, p in enumerate(ordered, start=1):
        print(f"  {idx:2d}. {os.path.basename(p)} ({os.path.getsize(p):,} bytes)")

    def _on_progress(msg: str, cur: int, total: int) -> None:
        print(f"[{cur}/{total}] {msg}")

    merge_res = merge_docx_files(
        ordered,
        out_abs,
        sort_by_name=False,  # already sorted above if requested
        section_break_mode=args.section_break,
        insert_toc=args.insert_toc,
        toc_levels=args.toc_levels,
        progress_callback=_on_progress,
    )
    print(
        f"拼接完成: {merge_res.output_path} "
        f"(sections={merge_res.section_count}, comments={merge_res.comments_merged}, "
        f"footnotes={merge_res.footnotes_merged}, endnotes={merge_res.endnotes_merged}, "
        f"has_toc={merge_res.has_toc})"
    )

    refreshed = False
    if not args.no_refresh_toc:
        r_res = refresh_toc_silent(out_abs, log_callback=print)
        print(f"目录刷新状态 [{r_res.status}]: {r_res.message}")
        if r_res.status == "error":
            return 1
        refreshed = r_res.status == "updated"

    if args.verify:
        print("正在执行严格内容与格式完整性校验...")
        report = verify_merged_docx(
            ordered,
            out_abs,
            ignore_toc_content=bool(args.insert_toc or merge_res.has_toc),
            check_media_hashes=not refreshed,
            word_resaved=refreshed,
        )
        print(report.format_summary())
        if not report.ok:
            return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
