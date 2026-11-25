"""Strict OOXML integrity, ID uniqueness, content parity, and TOC verifier."""

from __future__ import annotations

import hashlib
import os
import posixpath
import re
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

from lxml import etree

from .ooxml_merger import (
    NS_A,
    NS_CT,
    NS_MC,
    NS_O,
    NS_R,
    NS_REL,
    NS_V,
    NS_W,
    NS_W14,
    NS_WP,
    O,
    R,
    V,
    W,
    W14,
    _resolve_target,
    _TOC_RE,
)


@dataclass
class TOCEntry:
    """Single parsed entry from a Word Table of Contents."""

    level: str
    title: str
    page: int
    raw_text: str


@dataclass
class VerificationReport:
    """Detailed verification result for a merged .docx against its inputs."""

    ok: bool
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    stats: Dict[str, int] = field(default_factory=dict)
    toc_entries: List[TOCEntry] = field(default_factory=list)

    def format_summary(self) -> str:
        lines = [f"验证结果: {'通过 (PASS)' if self.ok else '失败 (FAIL)'}"]
        for k, v in self.stats.items():
            lines.append(f"  - {k}: {v}")
        if self.toc_entries:
            lines.append(f"  - 目录条目数: {len(self.toc_entries)}")
        for w in self.warnings:
            lines.append(f"  [WARN] {w}")
        for e in self.errors:
            lines.append(f"  [ERROR] {e}")
        return "\n".join(lines)


@dataclass
class _DocContentDigest:
    body_paragraphs: List[str]
    tables: List[List[List[str]]]
    image_ref_count: int
    image_hashes: Counter
    ole_count: int
    ole_hashes: Counter
    comments: List[Tuple[str, str]]  # (author, text)
    footnotes: List[str]
    endnotes: List[str]
    section_headers: List[Dict[str, str]]  # per-section {type: normalized text}
    section_footers: List[Dict[str, str]]


def _extract_p_text(p_el: etree._Element) -> str:
    parts: List[str] = []
    for el in p_el.iter():
        if el.tag == f"{W}t" or el.tag == f"{W}delText":
            if el.text:
                parts.append(el.text)
        elif el.tag == f"{W}tab":
            parts.append("\t")
        elif el.tag in (f"{W}br", f"{W}cr"):
            parts.append("\n")
    return "".join(parts)


def _extract_tbl_rows(tbl_el: etree._Element) -> List[List[str]]:
    rows: List[List[str]] = []
    for tr in tbl_el.findall(f"{W}tr"):
        row_cells: List[str] = []
        for tc in tr.findall(f"{W}tc"):
            cell_texts = []
            for ch in tc:
                if ch.tag == f"{W}p":
                    cell_texts.append(_extract_p_text(ch))
                elif ch.tag == f"{W}tbl":
                    cell_texts.append(str(_extract_tbl_rows(ch)))
            row_cells.append("\n".join(cell_texts).strip())
        rows.append(row_cells)
    return rows


def _is_toc_sdt(sdt_el: etree._Element) -> bool:
    sdt_pr = sdt_el.find(f"{W}sdtPr")
    if sdt_pr is not None:
        for gal in sdt_pr.iter(f"{W}docPartGallery"):
            val = (gal.get(f"{W}val") or "").strip().lower()
            if "table of contents" in val or val == "toc":
                return True
    for instr in sdt_el.iter(f"{W}instrText"):
        if _TOC_RE.search("".join(instr.itertext())):
            return True
    for fs in sdt_el.iter(f"{W}fldSimple"):
        if _TOC_RE.search(fs.get(f"{W}instr") or ""):
            return True
    return False


def _is_toc_title_before_sdt(prev_el: Optional[etree._Element]) -> bool:
    if prev_el is None or prev_el.tag != f"{W}p":
        return False
    txt = _extract_p_text(prev_el).replace(" ", "").strip()
    return txt in ("目录", "目　录")


def _parse_toc_entries_from_doc(doc_root: etree._Element) -> Tuple[List[TOCEntry], List[str]]:
    """Extract TOC entries and detect any ghost empty TOC lines."""
    entries: List[TOCEntry] = []
    ghost_warnings: List[str] = []

    def _parse_toc_paragraph(p: etree._Element) -> None:
        ppr = p.find(f"{W}pPr")
        pstyle = ""
        if ppr is not None:
            ps = ppr.find(f"{W}pStyle")
            if ps is not None:
                pstyle = ps.get(f"{W}val") or ""

        raw = _extract_p_text(p).strip()
        # Ignore field-delimiter-only closing paragraph (e.g. just <w:fldChar w:fldCharType="end"/>)
        if not raw:
            flds = [fc.get(f"{W}fldCharType") for fc in p.iter(f"{W}fldChar")]
            if flds and all(f in ("end", "begin", "separate") for f in flds) and not pstyle.upper().startswith("TOC"):
                return
            # A styled TOC line with no text IS a ghost entry!
            if pstyle.upper().startswith("TOC"):
                ghost_warnings.append(f"TOC 存在空条目 (style={pstyle})")
            return

        # Placeholder check
        if "目录将在刷新后自动生成" in raw:
            ghost_warnings.append("TOC 仍为未刷新的占位文本")
            return

        # Parse title and trailing page number (Word puts tab or page digits at end)
        m = re.match(r"^(.*?)[\t\s]*(\d+)$", raw)
        if m:
            title = m.group(1).strip()
            page = int(m.group(2))
            if not title:
                ghost_warnings.append(f"TOC 存在无标题的空条目 (raw={raw!r})")
                return
            entries.append(TOCEntry(level=pstyle or "TOC", title=title, page=page, raw_text=raw))
        else:
            # Some custom TOC header line inside SDT like "目录" or unnumbered entry
            if raw.replace(" ", "") in ("目录", "Contents", "TableofContents"):
                return
            ghost_warnings.append(f"TOC 条目缺少有效页码: {raw!r}")

    for sdt in doc_root.iter(f"{W}sdt"):
        if not _is_toc_sdt(sdt):
            continue
        sdt_content = sdt.find(f"{W}sdtContent")
        if sdt_content is None:
            continue
        for p in sdt_content.iter(f"{W}p"):
            _parse_toc_paragraph(p)

    return entries, ghost_warnings


def _digest_docx(
    docx_path: str,
    *,
    ignore_toc_content: bool = False,
) -> _DocContentDigest:
    with zipfile.ZipFile(docx_path, "r") as zf:
        names = set(zf.namelist())
        doc_root = etree.fromstring(zf.read("word/document.xml"))
        body = doc_root.find(f"{W}body")

        rels_by_id: Dict[str, Tuple[str, str]] = {}
        if "word/_rels/document.xml.rels" in names:
            rels_root = etree.fromstring(zf.read("word/_rels/document.xml.rels"))
            for rel in rels_root:
                rid = rel.get("Id")
                if rid:
                    rels_by_id[rid] = (rel.get("Type") or "", rel.get("Target") or "")

        body_paragraphs: List[str] = []
        tables: List[List[List[str]]] = []
        section_headers: List[Dict[str, str]] = []
        section_footers: List[Dict[str, str]] = []

        children = list(body) if body is not None else []
        prev_el: Optional[etree._Element] = None
        for idx, ch in enumerate(children):
            if ch.tag == f"{W}p":
                if ignore_toc_content and idx + 1 < len(children) and children[idx + 1].tag == f"{W}sdt":
                    if _is_toc_sdt(children[idx + 1]) and _is_toc_title_before_sdt(ch):
                        prev_el = ch
                        continue
                # If ignore_toc_content and this is an inserted page break right after TOC SDT
                if (
                    ignore_toc_content
                    and prev_el is not None
                    and prev_el.tag == f"{W}sdt"
                    and _is_toc_sdt(prev_el)
                    and not _extract_p_text(ch).strip()
                    and ch.find(f".//{W}br[@{W}type='page']") is not None
                ):
                    prev_el = ch
                    continue

                txt = _extract_p_text(ch).strip()
                if txt:
                    body_paragraphs.append(txt)

                ppr = ch.find(f"{W}pPr")
                if ppr is not None:
                    sp = ppr.find(f"{W}sectPr")
                    if sp is not None:
                        hdrs, ftrs = _read_sect_hf(sp, rels_by_id, zf, names)
                        section_headers.append(hdrs)
                        section_footers.append(ftrs)
            elif ch.tag == f"{W}tbl":
                tables.append(_extract_tbl_rows(ch))
            elif ch.tag == f"{W}sdt":
                if ignore_toc_content and _is_toc_sdt(ch):
                    prev_el = ch
                    continue
                sdt_content = ch.find(f"{W}sdtContent")
                if sdt_content is not None:
                    for sub in sdt_content:
                        if sub.tag == f"{W}p":
                            txt = _extract_p_text(sub).strip()
                            if txt:
                                body_paragraphs.append(txt)
                        elif sub.tag == f"{W}tbl":
                            tables.append(_extract_tbl_rows(sub))
            elif ch.tag == f"{W}sectPr":
                hdrs, ftrs = _read_sect_hf(ch, rels_by_id, zf, names)
                section_headers.append(hdrs)
                section_footers.append(ftrs)
            prev_el = ch

        # Count image refs and OLE objects across document.xml (and headers/footers/comments)
        image_ref_count = 0
        image_hashes: Counter = Counter()
        ole_count = 0
        ole_hashes: Counter = Counter()

        def _scan_part_media(part_path: str, root: etree._Element) -> None:
            nonlocal image_ref_count, ole_count
            part_rels_path = (
                "word/_rels/document.xml.rels"
                if part_path == "word/document.xml"
                else posixpath.join(posixpath.dirname(part_path), "_rels", f"{posixpath.basename(part_path)}.rels")
            )
            part_rels: Dict[str, Tuple[str, str]] = {}
            if part_rels_path in names:
                pr_root = etree.fromstring(zf.read(part_rels_path))
                for rel in pr_root:
                    rid = rel.get("Id")
                    if rid:
                        part_rels[rid] = (rel.get("Type") or "", rel.get("Target") or "")

            for el in root.iter():
                if el.tag == f"{{{NS_A}}}blip":
                    rid = el.get(f"{R}embed")
                    if rid and rid in part_rels:
                        image_ref_count += 1
                        target = _resolve_target(part_path, part_rels[rid][1])
                        if target in names:
                            image_hashes[hashlib.sha256(zf.read(target)).hexdigest()] += 1
                elif el.tag == f"{V}imagedata":
                    rid = el.get(f"{R}id")
                    if rid and rid in part_rels:
                        image_ref_count += 1
                        target = _resolve_target(part_path, part_rels[rid][1])
                        if target in names:
                            image_hashes[hashlib.sha256(zf.read(target)).hexdigest()] += 1
                elif el.tag == f"{O}OLEObject":
                    rid = el.get(f"{R}id")
                    if rid and rid in part_rels:
                        ole_count += 1
                        target = _resolve_target(part_path, part_rels[rid][1])
                        if target in names:
                            ole_hashes[hashlib.sha256(zf.read(target)).hexdigest()] += 1

        _scan_part_media("word/document.xml", doc_root)
        # Also scan referenced header/footer parts for images
        seen_hf_targets = set()
        for hf_ref in doc_root.xpath(".//w:headerReference | .//w:footerReference", namespaces={"w": NS_W}):
            rid = hf_ref.get(f"{R}id")
            if rid and rid in rels_by_id:
                target = _resolve_target("word/document.xml", rels_by_id[rid][1])
                if target in names and target not in seen_hf_targets:
                    seen_hf_targets.add(target)
                    _scan_part_media(target, etree.fromstring(zf.read(target)))

        comments: List[Tuple[str, str]] = []
        if "word/comments.xml" in names:
            croot = etree.fromstring(zf.read("word/comments.xml"))
            for cm in croot.findall(f"{W}comment"):
                author = cm.get(f"{W}author") or ""
                ctext = " ".join(_extract_p_text(p).strip() for p in cm.findall(f"{W}p")).strip()
                comments.append((author, ctext))

        footnotes: List[str] = []
        if "word/footnotes.xml" in names:
            froot = etree.fromstring(zf.read("word/footnotes.xml"))
            for fn in froot.findall(f"{W}footnote"):
                if fn.get(f"{W}type") not in ("separator", "continuationSeparator"):
                    footnotes.append(" ".join(_extract_p_text(p).strip() for p in fn.findall(f"{W}p")).strip())

        endnotes: List[str] = []
        if "word/endnotes.xml" in names:
            eroot = etree.fromstring(zf.read("word/endnotes.xml"))
            for en in eroot.findall(f"{W}endnote"):
                if en.get(f"{W}type") not in ("separator", "continuationSeparator"):
                    endnotes.append(" ".join(_extract_p_text(p).strip() for p in en.findall(f"{W}p")).strip())

    return _DocContentDigest(
        body_paragraphs=body_paragraphs,
        tables=tables,
        image_ref_count=image_ref_count,
        image_hashes=image_hashes,
        ole_count=ole_count,
        ole_hashes=ole_hashes,
        comments=comments,
        footnotes=footnotes,
        endnotes=endnotes,
        section_headers=section_headers,
        section_footers=section_footers,
    )


def _read_sect_hf(
    sect_pr: etree._Element,
    rels_by_id: Dict[str, Tuple[str, str]],
    zf: zipfile.ZipFile,
    names: set,
) -> Tuple[Dict[str, str], Dict[str, str]]:
    hdrs: Dict[str, str] = {}
    ftrs: Dict[str, str] = {}
    for hr in sect_pr.findall(f"{W}headerReference"):
        ht = hr.get(f"{W}type") or "default"
        rid = hr.get(f"{R}id")
        if rid and rid in rels_by_id:
            target = _resolve_target("word/document.xml", rels_by_id[rid][1])
            if target in names:
                root = etree.fromstring(zf.read(target))
                hdrs[ht] = " ".join(_extract_p_text(p).strip() for p in root.iter(f"{W}p")).strip()
    for fr in sect_pr.findall(f"{W}footerReference"):
        ft = fr.get(f"{W}type") or "default"
        rid = fr.get(f"{R}id")
        if rid and rid in rels_by_id:
            target = _resolve_target("word/document.xml", rels_by_id[rid][1])
            if target in names:
                root = etree.fromstring(zf.read(target))
                ftrs[ft] = " ".join(_extract_p_text(p).strip() for p in root.iter(f"{W}p")).strip()
    return hdrs, ftrs


def verify_merged_docx(
    input_paths: Sequence[str],
    merged_path: str,
    *,
    ignore_toc_content: bool = False,
    check_media_hashes: bool = True,
    word_resaved: bool = False,
    expected_toc_substrings: Optional[Sequence[str]] = None,
) -> VerificationReport:
    """Verify structural integrity, ID uniqueness, content parity, and TOC validity."""
    errors: List[str] = []
    warnings: List[str] = []
    stats: Dict[str, int] = {}
    toc_entries: List[TOCEntry] = []

    if not os.path.isfile(merged_path):
        return VerificationReport(ok=False, errors=[f"输出文件不存在: {merged_path}"])

    # 1. ZIP & XML well-formedness + Content_Types + rel targets
    parsed_xml: Dict[str, etree._Element] = {}
    try:
        with zipfile.ZipFile(merged_path, "r") as zf:
            names = set(zf.namelist())
            bad_crc = zf.testzip()
            if bad_crc is not None:
                errors.append(f"ZIP CRC 错误: {bad_crc}")

            for name in zf.namelist():
                if name.endswith(".xml") or name.endswith(".rels"):
                    try:
                        parsed_xml[name] = etree.fromstring(zf.read(name))
                    except etree.XMLSyntaxError as exc:
                        errors.append(f"XML 解析失败 {name}: {exc}")

            # Check Content_Types coverage
            ct_root = parsed_xml.get("[Content_Types].xml")
            if ct_root is None:
                errors.append("缺少 [Content_Types].xml")
            else:
                default_exts = {
                    (el.get("Extension") or "").lower()
                    for el in ct_root
                    if el.tag == f"{{{NS_CT}}}Default"
                }
                override_parts = {
                    (el.get("PartName") or "").lstrip("/")
                    for el in ct_root
                    if el.tag == f"{{{NS_CT}}}Override"
                }
                for name in names:
                    if name == "[Content_Types].xml" or name.endswith(".rels"):
                        continue
                    ext = posixpath.splitext(name)[1].lstrip(".").lower()
                    if name not in override_parts and ext not in default_exts:
                        errors.append(f"[Content_Types].xml 缺少定义: {name}")

            # Check every internal relationship target exists
            for rel_name, rel_root in parsed_xml.items():
                if not rel_name.endswith(".rels"):
                    continue
                if rel_name == "_rels/.rels":
                    src_part = ""
                else:
                    # word/_rels/document.xml.rels -> word/document.xml
                    rel_dir, rel_base = posixpath.split(rel_name)
                    parent_dir = posixpath.dirname(rel_dir)
                    part_base = rel_base[:-5]  # strip .rels
                    src_part = posixpath.join(parent_dir, part_base) if parent_dir else part_base
                for rel in rel_root:
                    if rel.get("TargetMode") == "External":
                        continue
                    target = rel.get("Target") or ""
                    resolved = _resolve_target(src_part, target)
                    if resolved not in names:
                        errors.append(f"关系目标丢失 {rel_name} ({rel.get('Id')}): {resolved}")
    except zipfile.BadZipFile as exc:
        return VerificationReport(ok=False, errors=[f"无效的 ZIP/DOCX 文件: {exc}"])

    # 2. Uniqueness checks on merged XML parts
    doc_root = parsed_xml.get("word/document.xml")
    styles_root = parsed_xml.get("word/styles.xml")
    if styles_root is not None:
        sids = [s.get(f"{W}styleId") for s in styles_root.findall(f"{W}style") if s.get(f"{W}styleId")]
        dup_sids = [k for k, v in Counter(sids).items() if v > 1]
        if dup_sids:
            errors.append(f"styles.xml 存在重复 styleId: {dup_sids[:10]}")
        snames = []
        for s in styles_root.findall(f"{W}style"):
            nm = s.find(f"{W}name")
            if nm is not None and nm.get(f"{W}val"):
                snames.append(nm.get(f"{W}val").lower())
        dup_snames = [k for k, v in Counter(snames).items() if v > 1]
        if dup_snames:
            errors.append(f"styles.xml 存在重复样式名称: {dup_snames[:10]}")

    if doc_root is not None:
        docpr_ids: List[str] = []
        bm_ids: List[str] = []
        v_shape_ids: List[str] = []
        ole_shape_ids: List[str] = []
        para_ids: List[str] = []
        v_shape_els: Dict[str, List[etree._Element]] = {}
        para_id_els: Dict[str, List[etree._Element]] = {}
        fallback_tag = f"{{{NS_MC}}}Fallback"

        for el in doc_root.iter():
            tag = el.tag
            if tag == f"{{{NS_WP}}}docPr":
                did = el.get("id")
                if did is not None:
                    docpr_ids.append(did)
            elif tag == f"{W}bookmarkStart":
                bid = el.get(f"{W}id")
                if bid is not None:
                    bm_ids.append(bid)
            elif tag == f"{V}shape":
                vid = el.get("id")
                if vid is not None:
                    v_shape_ids.append(vid)
                    if word_resaved:
                        v_shape_els.setdefault(vid, []).append(el)
            elif tag == f"{O}OLEObject":
                shid = el.get("ShapeID")
                if shid is not None:
                    ole_shape_ids.append(shid)
            pid = el.get(f"{W14}paraId")
            if pid is not None:
                up_pid = pid.upper()
                para_ids.append(up_pid)
                if word_resaved:
                    para_id_els.setdefault(up_pid, []).append(el)

        def _non_fallback_dups(seq: List[str], el_map: Dict[str, List[etree._Element]]) -> List[str]:
            dups = [k for k, v in Counter(seq).items() if v > 1]
            if not word_resaved or not dups:
                return dups
            real_dups: List[str] = []
            for k in dups:
                non_fb = [
                    e for e in el_map.get(k, [])
                    if not any(True for _ in e.iterancestors(fallback_tag))
                ]
                if len(non_fb) > 1:
                    real_dups.append(k)
            return real_dups

        for label, seq, el_map in (
            ("wp:docPr/@id", docpr_ids, {}),
            ("w:bookmarkStart/@w:id", bm_ids, {}),
            ("v:shape/@id", v_shape_ids, v_shape_els),
            ("o:OLEObject/@ShapeID", ole_shape_ids, {}),
            ("w14:paraId", para_ids, para_id_els),
        ):
            dups = _non_fallback_dups(seq, el_map) if el_map else [k for k, v in Counter(seq).items() if v > 1]
            if dups:
                errors.append(f"存在重复 {label}: {dups[:5]}")

    comments_root = parsed_xml.get("word/comments.xml")
    if comments_root is not None:
        cids = [c.get(f"{W}id") for c in comments_root.findall(f"{W}comment") if c.get(f"{W}id")]
        dup_cids = [k for k, v in Counter(cids).items() if v > 1]
        if dup_cids:
            errors.append(f"存在重复批注 ID: {dup_cids[:5]}")

    # 3. Strict content parity against ordered inputs
    expected_paragraphs: List[str] = []
    expected_tables: List[List[List[str]]] = []
    expected_img_refs = 0
    expected_img_hashes: Counter = Counter()
    expected_ole_count = 0
    expected_ole_hashes: Counter = Counter()
    expected_comments: List[Tuple[str, str]] = []
    expected_footnotes: List[str] = []
    expected_endnotes: List[str] = []
    expected_nonempty_headers: List[Tuple[str, Dict[str, str]]] = []

    for inp in input_paths:
        d = _digest_docx(inp, ignore_toc_content=ignore_toc_content)
        expected_paragraphs.extend(d.body_paragraphs)
        expected_tables.extend(d.tables)
        expected_img_refs += d.image_ref_count
        expected_img_hashes += d.image_hashes
        expected_ole_count += d.ole_count
        expected_ole_hashes += d.ole_hashes
        expected_comments.extend(d.comments)
        expected_footnotes.extend(d.footnotes)
        expected_endnotes.extend(d.endnotes)
        for sec_h in d.section_headers:
            non_empty_h = {k: v for k, v in sec_h.items() if v}
            if non_empty_h:
                expected_nonempty_headers.append((os.path.basename(inp), non_empty_h))

    merged_digest = _digest_docx(merged_path, ignore_toc_content=ignore_toc_content)

    stats["源文档数"] = len(input_paths)
    stats["非空正文段落数"] = len(merged_digest.body_paragraphs)
    stats["表格数"] = len(merged_digest.tables)
    stats["图片引用数"] = merged_digest.image_ref_count
    stats["嵌入OLE对象数"] = merged_digest.ole_count
    stats["批注数"] = len(merged_digest.comments)
    stats["脚注数"] = len(merged_digest.footnotes)
    stats["尾注数"] = len(merged_digest.endnotes)
    stats["节(Section)数"] = len(merged_digest.section_headers)

    if merged_digest.body_paragraphs != expected_paragraphs:
        diff_idx = next(
            (i for i, (a, b) in enumerate(zip(expected_paragraphs, merged_digest.body_paragraphs)) if a != b),
            min(len(expected_paragraphs), len(merged_digest.body_paragraphs)),
        )
        exp_sample = expected_paragraphs[diff_idx][:60] if diff_idx < len(expected_paragraphs) else "<EOF>"
        got_sample = (
            merged_digest.body_paragraphs[diff_idx][:60]
            if diff_idx < len(merged_digest.body_paragraphs)
            else "<EOF>"
        )
        errors.append(
            f"正文段落内容不一致: 期望 {len(expected_paragraphs)} 段, 实际 {len(merged_digest.body_paragraphs)} 段 "
            f"(首个差异 idx={diff_idx}: 期望={exp_sample!r}, 实际={got_sample!r})"
        )

    if len(merged_digest.tables) != len(expected_tables):
        errors.append(f"表格数量不一致: 期望 {len(expected_tables)}, 实际 {len(merged_digest.tables)}")
    else:
        for ti, (et, gt) in enumerate(zip(expected_tables, merged_digest.tables)):
            if et != gt:
                errors.append(f"第 {ti + 1} 个表格内容不一致 (期望 {len(et)} 行, 实际 {len(gt)} 行)")
                break

    if merged_digest.image_ref_count != expected_img_refs:
        errors.append(f"图片引用数不一致: 期望 {expected_img_refs}, 实际 {merged_digest.image_ref_count}")
    elif merged_digest.image_hashes != expected_img_hashes:
        msg = "嵌入图片二进制哈希集合与原文档不完全一致"
        if check_media_hashes:
            errors.append(msg)
        else:
            warnings.append(msg + "（可能由 Word 重新保存引起）")

    if merged_digest.ole_count != expected_ole_count:
        errors.append(f"OLE 对象数不一致: 期望 {expected_ole_count}, 实际 {merged_digest.ole_count}")
    elif merged_digest.ole_hashes != expected_ole_hashes:
        msg = "嵌入 OLE 工作簿二进制哈希集合与原文档不完全一致"
        if check_media_hashes:
            errors.append(msg)
        else:
            warnings.append(msg + "（可能由 Word 重新保存引起）")

    if merged_digest.comments != expected_comments:
        errors.append(
            f"批注内容不一致: 期望 {len(expected_comments)} 条 {expected_comments}, 实际 {len(merged_digest.comments)} 条 {merged_digest.comments}"
        )
    if merged_digest.footnotes != expected_footnotes:
        errors.append(f"脚注不一致: 期望 {len(expected_footnotes)} 条, 实际 {len(merged_digest.footnotes)} 条")
    if merged_digest.endnotes != expected_endnotes:
        errors.append(f"尾注不一致: 期望 {len(expected_endnotes)} 条, 实际 {len(merged_digest.endnotes)} 条")

    # Header isolation check: count non-empty section headers matches expected
    actual_nonempty_headers = [
        {k: v for k, v in sec_h.items() if v}
        for sec_h in merged_digest.section_headers
        if any(v for v in sec_h.values())
    ]
    exp_hdr_dicts = [h for _, h in expected_nonempty_headers]
    if actual_nonempty_headers != exp_hdr_dicts:
        errors.append(
            f"页眉节隔离异常: 期望非空页眉={exp_hdr_dicts}, 实际={actual_nonempty_headers}"
        )

    # 4. TOC validation (when TOC is present)
    if doc_root is not None:
        toc_entries, toc_ghosts = _parse_toc_entries_from_doc(doc_root)
        for gw in toc_ghosts:
            errors.append(gw)
        if toc_entries:
            has_pg_restart = bool(
                doc_root.xpath(f".//w:pgNumType[@w:start]", namespaces={"w": NS_W})
            )
            prev_page = 0
            for te in toc_entries:
                if te.page <= 0:
                    errors.append(f"TOC 页码无效: {te.raw_text!r}")
                elif te.page < prev_page and not has_pg_restart:
                    errors.append(f"TOC 页码倒退 ({prev_page} -> {te.page}): {te.raw_text!r}")
                prev_page = te.page
            if expected_toc_substrings:
                search_pos = 0
                for sub in expected_toc_substrings:
                    found_idx = None
                    for i in range(search_pos, len(toc_entries)):
                        if sub in toc_entries[i].title or sub in toc_entries[i].raw_text:
                            found_idx = i
                            break
                    if found_idx is None:
                        errors.append(f"TOC 缺少期望章节或顺序错误: {sub!r}")
                    else:
                        search_pos = found_idx + 1

    return VerificationReport(
        ok=len(errors) == 0,
        errors=errors,
        warnings=warnings,
        stats=stats,
        toc_entries=toc_entries,
    )
