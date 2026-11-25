"""High-fidelity pure-Python OOXML (.docx) deep-merge engine.

Preserves original paragraph/table content, formatting, styles, isolated list
numbering, embedded images, OLE Excel workbooks, comments, footnotes/endnotes,
hyperlinks, bookmarks, and per-section headers/footers without leaking them across
section boundaries.
"""

from __future__ import annotations

import copy
import hashlib
import os
import posixpath
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass, field
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Set, Tuple
from xml.sax.saxutils import escape as xml_escape

from lxml import etree

from .sorting import sort_docx_paths

# OOXML Namespaces
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
NS_V = "urn:schemas-microsoft-com:vml"
NS_O = "urn:schemas-microsoft-com:office:office"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
NS_W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
NS_W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
NS_WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
NS_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
NS_WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
NS_WPC = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
NS_XML = "http://www.w3.org/XML/1998/namespace"

W = f"{{{NS_W}}}"
R = f"{{{NS_R}}}"
O = f"{{{NS_O}}}"
V = f"{{{NS_V}}}"
MC = f"{{{NS_MC}}}"
W14 = f"{{{NS_W14}}}"
W15 = f"{{{NS_W15}}}"
W16CID = f"{{{NS_W16CID}}}"
WP14 = f"{{{NS_WP14}}}"

REL_TYPE_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
REL_TYPE_NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
REL_TYPE_FONT_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
REL_TYPE_SETTINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
REL_TYPE_HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
REL_TYPE_FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
REL_TYPE_FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
REL_TYPE_ENDNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
REL_TYPE_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
REL_TYPE_COMMENTS_EXT = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
REL_TYPE_COMMENTS_IDS = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
REL_TYPE_PEOPLE = "http://schemas.microsoft.com/office/2011/relationships/people"
REL_TYPE_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

CT_NUMBERING = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
CT_HEADER = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
CT_FOOTER = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
CT_FOOTNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
CT_ENDNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
CT_COMMENTS_EXT = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
CT_COMMENTS_IDS = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"
CT_PEOPLE = "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"

_TOC_RE = re.compile(r"\bTOC\b", re.IGNORECASE)
_HEADING_NAME_RE = re.compile(r"^heading\s+([1-9])$", re.IGNORECASE)
_ANCHOR_FIELD_RE = re.compile(
    r'^(\s*(?:PAGEREF|REF|NOTEREF)\s+)("?)([^\s"\\]+)("?)(.*)$',
    re.IGNORECASE | re.DOTALL,
)
_HYPERLINK_L_FIELD_RE = re.compile(
    r'(\s+\\l\s+)("?)([^"\s\\]+|\S+)("?)',
    re.IGNORECASE,
)

_DRAWING_ID_TAGS = {
    f"{{{NS_WP}}}docPr",
    f"{{{NS_PIC}}}cNvPr",
    f"{{{NS_A}}}cNvPr",
    f"{{{NS_WPS}}}cNvPr",
    f"{{{NS_WPG}}}cNvPr",
    f"{{{NS_WPC}}}cNvPr",
    "cNvPr",
}
_VML_SHAPE_TAGS = {
    f"{V}shape",
    f"{V}shapetype",
    f"{V}rect",
    f"{V}oval",
    f"{V}line",
    f"{V}roundrect",
    f"{V}group",
    f"{V}polyline",
    f"{V}curve",
    f"{V}arc",
}
_IGNORED_STYLE_TAGS = {
    f"{W}rsid",
    f"{W}uiPriority",
    f"{W}qFormat",
    f"{W}unhideWhenUsed",
    f"{W}semiHidden",
    f"{W}personal",
    f"{W}personalCompose",
    f"{W}personalReply",
}

# OOXML w:sectPr element ordering for headerReference / footerReference insertion
_SECTPR_FIRST_TAGS = {f"{W}headerReference", f"{W}footerReference"}

ProgressCallback = Callable[[str, int, int], None]
FileStatusCallback = Callable[[int, str], None]


@dataclass
class TOCInfo:
    """Information about Table of Contents structures inside a .docx file."""

    count: int
    sdt_count: int
    field_count: int
    instructions: List[str] = field(default_factory=list)

    @property
    def has_toc(self) -> bool:
        return self.count > 0


@dataclass
class MergeResult:
    """Summary of an OOXML merge run."""

    output_path: str
    ordered_inputs: List[str]
    has_toc: bool
    toc_info: TOCInfo
    section_count: int
    comments_merged: int
    footnotes_merged: int
    endnotes_merged: int


def _xml_parser() -> etree.XMLParser:
    return etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)


def _parse_xml(data: bytes) -> etree._Element:
    return etree.fromstring(data, parser=_xml_parser())


def _serialize_xml(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _rels_path_for(part_path: str) -> str:
    dirname, basename = posixpath.split(part_path)
    if dirname:
        return posixpath.join(dirname, "_rels", f"{basename}.rels")
    return posixpath.join("_rels", f"{basename}.rels")


def _resolve_target(source_part: str, target: str) -> str:
    base_dir = posixpath.dirname(source_part)
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base_dir, target))


def _relative_target(source_part: str, target_part: str) -> str:
    base_dir = posixpath.dirname(source_part)
    return posixpath.relpath(target_part, base_dir or ".").replace(os.sep, "/")


def detect_tocs(docx_path: str) -> TOCInfo:
    """Detect whether `docx_path` contains Word Table of Contents (TOC) SDTs or fields."""
    with zipfile.ZipFile(docx_path, "r") as zf:
        if "word/document.xml" not in zf.namelist():
            return TOCInfo(count=0, sdt_count=0, field_count=0, instructions=[])
        doc_root = _parse_xml(zf.read("word/document.xml"))

    sdt_toc_elements: Set[int] = set()
    instructions: List[str] = []
    field_count = 0

    for sdt in doc_root.iter(f"{W}sdt"):
        is_toc_sdt = False
        sdt_pr = sdt.find(f"{W}sdtPr")
        if sdt_pr is not None:
            for gal in sdt_pr.iter(f"{W}docPartGallery"):
                val = (gal.get(f"{W}val") or "").strip().lower()
                if "table of contents" in val or val == "toc":
                    is_toc_sdt = True
                    break
        if is_toc_sdt:
            sdt_toc_elements.add(id(sdt))

    # Track TOC fields and whether each sits inside a known TOC SDT
    sdt_ancestors: Dict[int, bool] = {}

    def _inside_toc_sdt(el: etree._Element) -> bool:
        cur = el.getparent()
        while cur is not None:
            cid = id(cur)
            if cid in sdt_ancestors:
                return sdt_ancestors[cid]
            if cid in sdt_toc_elements:
                sdt_ancestors[id(el)] = True
                return True
            cur = cur.getparent()
        sdt_ancestors[id(el)] = False
        return False

    standalone_fields = 0
    for el in doc_root.iter():
        if el.tag == f"{W}instrText":
            txt = "".join(el.itertext())
            if _TOC_RE.search(txt):
                field_count += 1
                instructions.append(txt.strip())
                if not _inside_toc_sdt(el):
                    standalone_fields += 1
        elif el.tag == f"{W}fldSimple":
            instr = el.get(f"{W}instr") or ""
            if _TOC_RE.search(instr):
                field_count += 1
                instructions.append(instr.strip())
                if not _inside_toc_sdt(el):
                    standalone_fields += 1

    total = len(sdt_toc_elements) + standalone_fields
    return TOCInfo(
        count=total,
        sdt_count=len(sdt_toc_elements),
        field_count=field_count,
        instructions=instructions,
    )


class _IdAllocator:
    """Tracks allocated IDs and produces globally unique values across merged docs."""

    def __init__(self) -> None:
        self.next_int: Dict[str, int] = {}
        self.used_str: Dict[str, Set[str]] = {}
        self.used_para_ids: Set[str] = set()
        self.next_para_int: int = 0x10000001

    def seed_int(self, bucket: str, value: int) -> None:
        cur = self.next_int.get(bucket, 1)
        if value >= cur:
            self.next_int[bucket] = value + 1

    def alloc_int(self, bucket: str, start: int = 1) -> int:
        val = max(self.next_int.get(bucket, start), start)
        self.next_int[bucket] = val + 1
        return val

    def seed_str(self, bucket: str, value: str) -> None:
        self.used_str.setdefault(bucket, set()).add(value)

    def has_str(self, bucket: str, value: str) -> bool:
        return value in self.used_str.get(bucket, set())

    def alloc_prefixed(self, bucket: str, prefix: str, candidate: str) -> str:
        used = self.used_str.setdefault(bucket, set())
        if candidate not in used:
            used.add(candidate)
            return candidate
        base = f"{prefix}_{candidate}"
        if base not in used:
            used.add(base)
            return base
        idx = 2
        while True:
            cand = f"{base}_{idx}"
            if cand not in used:
                used.add(cand)
                return cand
            idx += 1

    def alloc_bookmark_name(self, doc_idx: int, old_name: str) -> str:
        used = self.used_str.setdefault("bm_name", set())
        if old_name not in used:
            used.add(old_name)
            return old_name
        hidden = old_name.startswith("_")
        core = old_name[1:] if hidden else old_name
        suffix = f"_d{doc_idx}"
        # Word bookmark names have a 40-char limit
        max_core = 40 - (1 if hidden else 0) - len(suffix)
        cand = f"{'_' if hidden else ''}{core[:max(1, max_core)]}{suffix}"
        if cand not in used:
            used.add(cand)
            return cand
        idx = 2
        while True:
            extra = f"_{idx}"
            max_core2 = 40 - (1 if hidden else 0) - len(suffix) - len(extra)
            cand2 = f"{'_' if hidden else ''}{core[:max(1, max_core2)]}{suffix}{extra}"
            if cand2 not in used:
                used.add(cand2)
                return cand2
            idx += 1

    def seed_para_id(self, hex_id: str) -> None:
        up = hex_id.upper()
        self.used_para_ids.add(up)
        try:
            v = int(up, 16)
            if v >= self.next_para_int and v < 0x7FFFFFFE:
                self.next_para_int = v + 1
        except ValueError:
            pass

    def alloc_para_id(self) -> str:
        while True:
            if self.next_para_int >= 0x7FFFFFFE:
                self.next_para_int = 0x00000100
            cand = f"{self.next_para_int:08X}"
            self.next_para_int += 1
            if cand not in self.used_para_ids:
                self.used_para_ids.add(cand)
                return cand


def _get_or_create_rels(parts: Dict[str, bytes], parsed_rels: Dict[str, etree._Element], rels_path: str) -> etree._Element:
    if rels_path in parsed_rels:
        return parsed_rels[rels_path]
    if rels_path in parts:
        root = _parse_xml(parts[rels_path])
    else:
        root = etree.Element(f"{{{NS_REL}}}Relationships", nsmap={None: NS_REL})
    parsed_rels[rels_path] = root
    return root


def _rels_map(rels_root: Optional[etree._Element]) -> Dict[str, etree._Element]:
    if rels_root is None:
        return {}
    return {el.get("Id"): el for el in rels_root if el.get("Id")}


def _alloc_rid(rels_root: etree._Element) -> str:
    existing = {el.get("Id") for el in rels_root if el.get("Id")}
    n = 1
    while True:
        cand = f"rId{n}"
        if cand not in existing:
            return cand
        n += 1


def _find_rel_by_type(rels_root: Optional[etree._Element], rel_type: str) -> Optional[etree._Element]:
    if rels_root is None:
        return None
    for el in rels_root:
        if el.get("Type") == rel_type:
            return el
    return None


def _ensure_content_type_override(ct_root: etree._Element, part_name: str, content_type: str) -> None:
    pn = "/" + part_name.lstrip("/")
    for el in ct_root:
        if el.tag == f"{{{NS_CT}}}Override" and el.get("PartName") == pn:
            return
    ov = etree.SubElement(ct_root, f"{{{NS_CT}}}Override")
    ov.set("PartName", pn)
    ov.set("ContentType", content_type)


def _merge_content_type_defaults(base_ct: etree._Element, src_ct: etree._Element) -> None:
    existing_ext = {
        (el.get("Extension") or "").lower()
        for el in base_ct
        if el.tag == f"{{{NS_CT}}}Default"
    }
    for el in src_ct:
        if el.tag == f"{{{NS_CT}}}Default":
            ext = (el.get("Extension") or "").lower()
            if ext and ext not in existing_ext:
                base_ct.insert(0, copy.deepcopy(el))
                existing_ext.add(ext)


def _alloc_unique_part_path(parts: Dict[str, bytes], desired_dir: str, doc_idx: int, src_basename: str) -> str:
    stem, ext = posixpath.splitext(src_basename)
    cand = posixpath.join(desired_dir, f"d{doc_idx}_{stem}{ext}")
    if cand not in parts:
        return cand
    i = 2
    while True:
        cand2 = posixpath.join(desired_dir, f"d{doc_idx}_{stem}_{i}{ext}")
        if cand2 not in parts:
            return cand2
        i += 1


def _clean_style_for_fingerprint(
    style_el: etree._Element,
    resolved_based_on: Optional[str],
) -> str:
    """Produce a semantic fingerprint string for a <w:style> ignoring rsid/uiPriority."""
    clone = copy.deepcopy(style_el)
    stype = clone.get(f"{W}type") or ""
    # Strip attributes that don't affect rendering
    for attr in (f"{W}default", f"{W}customStyle", f"{W}styleId"):
        clone.attrib.pop(attr, None)

    orig_name = ""
    for child in list(clone):
        if child.tag == f"{W}name":
            orig_name = child.get(f"{W}val") or ""
            clone.remove(child)
        elif child.tag in _IGNORED_STYLE_TAGS:
            clone.remove(child)
        elif child.tag in (f"{W}basedOn", f"{W}next", f"{W}link"):
            clone.remove(child)

    # For heading styles (heading 1..9), keep heading level role in fingerprint
    m = _HEADING_NAME_RE.match(orig_name.strip())
    role = f"heading:{m.group(1)}" if m else "other"

    body_xml = etree.tostring(clone, encoding="unicode")
    # Strip any remaining rsid attributes or redundant xmlns on clone
    body_xml = re.sub(r' xmlns:[a-zA-Z0-9_]+="[^"]+"', "", body_xml)
    body_xml = re.sub(r' w:rsid[A-Za-z]*="[^"]+"', "", body_xml)
    return f"{stype}|{role}|{resolved_based_on or ''}|{body_xml}"


def _clean_doc_defaults(styles_root: Optional[etree._Element]) -> str:
    if styles_root is None:
        return ""
    dd = styles_root.find(f"{W}docDefaults")
    if dd is None:
        return ""
    s = etree.tostring(dd, encoding="unicode")
    s = re.sub(r' xmlns:[a-zA-Z0-9_]+="[^"]+"', "", s)
    s = re.sub(r' w:rsid[A-Za-z]*="[^"]+"', "", s)
    return s


def _get_default_style(styles_root: Optional[etree._Element], stype: str) -> Optional[etree._Element]:
    if styles_root is None:
        return None
    for st in styles_root.findall(f"{W}style"):
        if st.get(f"{W}type") == stype and st.get(f"{W}default") == "1":
            return st
    return None


def _style_topological_order(styles_by_id: Dict[str, etree._Element], target_ids: Set[str]) -> List[str]:
    ordered: List[str] = []
    visiting: Set[str] = set()
    visited: Set[str] = set()

    def _visit(sid: str) -> None:
        if sid in visited or sid in visiting or sid not in styles_by_id:
            return
        visiting.add(sid)
        based = styles_by_id[sid].find(f"{W}basedOn")
        if based is not None:
            parent_id = based.get(f"{W}val")
            if parent_id and parent_id in target_ids:
                _visit(parent_id)
        visiting.remove(sid)
        visited.add(sid)
        ordered.append(sid)

    for sid in sorted(target_ids):
        _visit(sid)
    return ordered


def _collect_used_styles_closure(
    src_styles_by_id: Dict[str, etree._Element],
    xml_roots: Sequence[etree._Element],
) -> Set[str]:
    used: Set[str] = set()
    for root in xml_roots:
        for el in root.iter():
            t = el.tag
            if t == f"{W}pStyle" or t == f"{W}rStyle" or t == f"{W}tblStyle":
                val = el.get(f"{W}val")
                if val:
                    used.add(val)

    closure = set(used)
    queue = list(used)
    while queue:
        sid = queue.pop()
        st = src_styles_by_id.get(sid)
        if st is None:
            continue
        for ref_tag in (f"{W}basedOn", f"{W}next", f"{W}link"):
            ref = st.find(ref_tag)
            if ref is not None:
                val = ref.get(f"{W}val")
                if val and val not in closure:
                    closure.add(val)
                    queue.append(val)
    return closure


def _empty_header_footer_xml(is_header: bool) -> bytes:
    tag = "hdr" if is_header else "ftr"
    root = etree.Element(
        f"{W}{tag}",
        nsmap={
            "w": NS_W,
            "r": NS_R,
        },
    )
    etree.SubElement(root, f"{W}p")
    return _serialize_xml(root)


def _is_hf_empty(data: bytes) -> bool:
    if not data:
        return True
    try:
        root = _parse_xml(data)
    except etree.XMLSyntaxError:
        return False
    for el in root.iter():
        if el.tag in (f"{W}drawing", f"{W}pict", f"{V}shape", f"{W}tbl", f"{W}object"):
            return False
        if el.tag == f"{W}t" and (el.text or "").strip():
            return False
    return True


def _extract_sect_signature(
    sect_pr: Optional[etree._Element],
    rels_by_id: Dict[str, etree._Element],
    src_zip: Optional[zipfile.ZipFile],
    src_doc_path: str,
    parts: Optional[Dict[str, bytes]] = None,
) -> Tuple[str, Dict[str, bytes], Dict[str, bytes]]:
    """Build a canonical signature of section layout & header/footer contents."""
    if sect_pr is None:
        return ("", {}, {})

    hdr_contents: Dict[str, bytes] = {}
    ftr_contents: Dict[str, bytes] = {}
    clone = copy.deepcopy(sect_pr)

    for child in list(clone):
        if child.tag in (f"{W}headerReference", f"{W}footerReference"):
            htype = child.get(f"{W}type") or "default"
            rid = child.get(f"{R}id")
            if rid and rid in rels_by_id:
                target = _resolve_target(src_doc_path, rels_by_id[rid].get("Target") or "")
                data = b""
                if src_zip is not None and target in src_zip.namelist():
                    data = src_zip.read(target)
                elif parts is not None and target in parts:
                    data = parts[target]
                if not _is_hf_empty(data):
                    if child.tag == f"{W}headerReference":
                        hdr_contents[htype] = data
                    else:
                        ftr_contents[htype] = data
            clone.remove(child)

    s = etree.tostring(clone, encoding="unicode")
    s = re.sub(r' xmlns:[a-zA-Z0-9_]+="[^"]+"', "", s)
    s = re.sub(r' w:rsid[A-Za-z]*="[^"]+"', "", s)
    hdr_sig = tuple(sorted((k, hashlib.sha256(v).hexdigest()) for k, v in hdr_contents.items()))
    ftr_sig = tuple(sorted((k, hashlib.sha256(v).hexdigest()) for k, v in ftr_contents.items()))
    return (f"{s}|H:{hdr_sig}|F:{ftr_sig}", hdr_contents, ftr_contents)


def _set_sect_type(sect_pr: etree._Element, val: str) -> None:
    """Insert or update `<w:type w:val="..."/>` in schema-valid order inside `w:sectPr`."""
    t_el = sect_pr.find(f"{W}type")
    if t_el is None:
        t_el = etree.Element(f"{W}type")
        before_tags = {
            f"{W}headerReference",
            f"{W}footerReference",
            f"{W}footnotePr",
            f"{W}endnotePr",
        }
        insert_idx = 0
        for idx, ch in enumerate(list(sect_pr)):
            if ch.tag in before_tags:
                insert_idx = idx + 1
            else:
                break
        sect_pr.insert(insert_idx, t_el)
    t_el.set(f"{W}val", val)


def _make_section_break_paragraph(sect_pr: etree._Element) -> etree._Element:
    p = etree.Element(f"{W}p")
    ppr = etree.SubElement(p, f"{W}pPr")
    ppr.append(copy.deepcopy(sect_pr))
    return p


def _insert_sect_hf_ref(sect_pr: etree._Element, is_header: bool, htype: str, rid: str) -> None:
    tag = f"{W}headerReference" if is_header else f"{W}footerReference"
    ref = etree.Element(tag)
    ref.set(f"{W}type", htype)
    ref.set(f"{R}id", rid)
    # Place headerReferences before footerReferences, both at top of sectPr
    insert_idx = 0
    for idx, child in enumerate(list(sect_pr)):
        if child.tag == f"{W}headerReference":
            insert_idx = idx + 1
        elif not is_header and child.tag == f"{W}footerReference":
            insert_idx = idx + 1
        elif child.tag not in _SECTPR_FIRST_TAGS:
            break
    sect_pr.insert(insert_idx, ref)


def _enforce_hf_isolation(
    prev_hdr_types: Set[str],
    prev_ftr_types: Set[str],
    cur_sect_pr: etree._Element,
    ensure_empty_hf: Callable[[bool], str],
) -> Tuple[Set[str], Set[str]]:
    """Ensure cur_sect_pr doesn't silently inherit headers/footers from prev section."""
    cur_hdrs = {
        (el.get(f"{W}type") or "default")
        for el in cur_sect_pr.findall(f"{W}headerReference")
    }
    cur_ftrs = {
        (el.get(f"{W}type") or "default")
        for el in cur_sect_pr.findall(f"{W}footerReference")
    }

    for ht in sorted(prev_hdr_types - cur_hdrs):
        empty_rid = ensure_empty_hf(True)
        _insert_sect_hf_ref(cur_sect_pr, True, ht, empty_rid)
        cur_hdrs.add(ht)

    for ft in sorted(prev_ftr_types - cur_ftrs):
        empty_rid = ensure_empty_hf(False)
        _insert_sect_hf_ref(cur_sect_pr, False, ft, empty_rid)
        cur_ftrs.add(ft)

    return cur_hdrs, cur_ftrs


def _build_toc_sdt_and_page_break(toc_levels: str = "1-3", title: str = "目  录") -> List[etree._Element]:
    """Create a standard Word TOC Content Control SDT followed by a page break."""
    title_p = etree.Element(f"{W}p")
    tppr = etree.SubElement(title_p, f"{W}pPr")
    tjc = etree.SubElement(tppr, f"{W}jc")
    tjc.set(f"{W}val", "center")
    tr = etree.SubElement(title_p, f"{W}r")
    trpr = etree.SubElement(tr, f"{W}rPr")
    etree.SubElement(trpr, f"{W}b")
    tsz = etree.SubElement(trpr, f"{W}sz")
    tsz.set(f"{W}val", "32")
    tt = etree.SubElement(tr, f"{W}t")
    tt.text = title

    safe_levels = re.sub(r"[^0-9\-]", "", toc_levels) or "1-3"
    sdt_xml = f'''<w:sdt xmlns:w="{NS_W}">
  <w:sdtPr>
    <w:docPartObj>
      <w:docPartGallery w:val="Table of Contents"/>
      <w:docPartUnique/>
    </w:docPartObj>
  </w:sdtPr>
  <w:sdtContent>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>
      <w:r><w:instrText xml:space="preserve"> TOC \\o "{xml_escape(safe_levels)}" \\h \\z \\u </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>目录将在刷新后自动生成</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:sdtContent>
</w:sdt>'''
    sdt_el = etree.fromstring(sdt_xml)

    br_p = etree.Element(f"{W}p")
    br_r = etree.SubElement(br_p, f"{W}r")
    br = etree.SubElement(br_r, f"{W}br")
    br.set(f"{W}type", "page")

    return [title_p, sdt_el, br_p]


def _seed_base_identifiers(
    alloc: _IdAllocator,
    parts: Dict[str, bytes],
    base_doc_root: etree._Element,
    base_styles_root: Optional[etree._Element],
    base_num_root: Optional[etree._Element],
    base_comments_root: Optional[etree._Element],
    base_comments_ids_root: Optional[etree._Element],
    base_footnotes_root: Optional[etree._Element],
    base_endnotes_root: Optional[etree._Element],
) -> None:
    """Seed allocator with all existing IDs in the base document (doc 0)."""
    # Numbering IDs
    if base_num_root is not None:
        for an in base_num_root.findall(f"{W}abstractNum"):
            aid = an.get(f"{W}abstractNumId")
            if aid and aid.isdigit():
                alloc.seed_int("abstractNum", int(aid))
        for num in base_num_root.findall(f"{W}num"):
            nid = num.get(f"{W}numId")
            if nid and nid.isdigit():
                alloc.seed_int("numId", int(nid))

    # Styles
    if base_styles_root is not None:
        for st in base_styles_root.findall(f"{W}style"):
            sid = st.get(f"{W}styleId")
            if sid:
                alloc.seed_str("styleId", sid)
            nm = st.find(f"{W}name")
            if nm is not None:
                nval = nm.get(f"{W}val")
                if nval:
                    alloc.seed_str("styleName", nval.lower())

    # Comments / Footnotes / Endnotes
    if base_comments_root is not None:
        for cm in base_comments_root.findall(f"{W}comment"):
            cid = cm.get(f"{W}id")
            if cid and cid.lstrip("-").isdigit():
                alloc.seed_int("comment", int(cid))
    if base_comments_ids_root is not None:
        for cid in base_comments_ids_root.findall(f"{W16CID}commentId"):
            did = cid.get(f"{W16CID}durableId")
            if did:
                alloc.seed_str("durableId", did.upper())
    if base_footnotes_root is not None:
        for fn in base_footnotes_root.findall(f"{W}footnote"):
            fid = fn.get(f"{W}id")
            if fid and fid.lstrip("-").isdigit():
                alloc.seed_int("footnote", int(fid))
    if base_endnotes_root is not None:
        for en in base_endnotes_root.findall(f"{W}endnote"):
            eid = en.get(f"{W}id")
            if eid and eid.lstrip("-").isdigit():
                alloc.seed_int("endnote", int(eid))

    # Scan all XML roots in base doc once for DrawingML / VML / OLE / bookmark / paraIds
    xml_roots: List[etree._Element] = [base_doc_root]
    for pname, raw in parts.items():
        if pname == "word/document.xml" or not pname.endswith(".xml") or pname.endswith(".rels"):
            continue
        if pname.startswith("word/") and (
            pname.startswith("word/header")
            or pname.startswith("word/footer")
            or pname in (
                "word/footnotes.xml",
                "word/endnotes.xml",
                "word/comments.xml",
                "word/commentsExtended.xml",
            )
        ):
            try:
                xml_roots.append(_parse_xml(raw))
            except etree.XMLSyntaxError:
                pass

    for root in xml_roots:
        for el in root.iter():
            tag = el.tag
            if tag == f"{W}bookmarkStart":
                bid = el.get(f"{W}id")
                if bid and bid.lstrip("-").isdigit():
                    alloc.seed_int("bm_id", int(bid))
                bname = el.get(f"{W}name")
                if bname:
                    alloc.seed_str("bm_name", bname)
            elif tag == f"{W}bookmarkEnd":
                bid = el.get(f"{W}id")
                if bid and bid.lstrip("-").isdigit():
                    alloc.seed_int("bm_id", int(bid))
            elif tag in _DRAWING_ID_TAGS or etree.QName(tag).localname in ("docPr", "cNvPr"):
                did = el.get("id")
                if did and did.isdigit():
                    alloc.seed_int("drawing_id", int(did))
            elif tag in _VML_SHAPE_TAGS:
                vid = el.get("id")
                if vid:
                    alloc.seed_str("vml_id", vid)
                spid = el.get(f"{O}spid")
                if spid:
                    alloc.seed_str("vml_spid", spid)
            elif tag == f"{O}OLEObject":
                shid = el.get("ShapeID")
                if shid:
                    alloc.seed_str("vml_id", shid)
                oid = el.get("ObjectID")
                if oid:
                    alloc.seed_str("ole_obj_id", oid)

            pid = el.get(f"{W14}paraId") or el.get(f"{W15}paraId")
            if pid:
                alloc.seed_para_id(pid)
            for attr_key in (f"{WP14}anchorId", f"{WP14}editId"):
                aid = el.get(attr_key)
                if aid:
                    alloc.seed_para_id(aid)


def _remap_single_pass(
    root: etree._Element,
    *,
    doc_idx: int,
    alloc: _IdAllocator,
    style_map: Dict[str, str],
    num_map: Dict[str, str],
    rel_map: Dict[str, str],
    comment_id_map: Dict[str, str],
    footnote_id_map: Dict[str, str],
    endnote_id_map: Dict[str, str],
    bm_id_map: Dict[str, str],
    bm_name_map: Dict[str, str],
    vml_id_map: Dict[str, str],
    vml_spid_map: Dict[str, str],
    ole_obj_id_map: Dict[str, str],
    para_id_map: Dict[str, str],
    normal_style_override: Optional[str],
    styled_table_ids_no_normal: Optional[Set[int]] = None,
    strip_vanish_outline: bool = False,
) -> None:
    """Remap all styles, numIds, rIds, IDs, bookmarks, VML, OLE, and paraIds in one O(N) pass."""
    goback_ids: Set[str] = set()
    to_remove: List[Tuple[etree._Element, etree._Element]] = []
    anchor_hyperlinks: List[etree._Element] = []
    field_instr_nodes: List[etree._Element] = []
    field_simple_nodes: List[etree._Element] = []
    vml_type_refs: List[etree._Element] = []

    # Track table styling scopes cheaply using element ids if needed
    # Better: pre-identify <w:p> inside <w:tbl> that have a non-default <w:tblStyle>?
    # We do it cleanly right inside iter() by checking parent chain only for unstyled <w:p>
    # when normal_style_override is active!
    for el in root.iter():
        tag = el.tag

        # Remap ANY r:* relationship attribute
        for ak in tuple(el.attrib.keys()):
            if ak.startswith(R):
                old_rid = el.attrib[ak]
                if old_rid in rel_map:
                    el.attrib[ak] = rel_map[old_rid]

        # Remap w14:paraId / wp14:anchorId / wp14:editId
        pid = el.attrib.get(f"{W14}paraId")
        if pid:
            new_pid = alloc.alloc_para_id()
            para_id_map.setdefault(pid, new_pid)
            el.attrib[f"{W14}paraId"] = new_pid
        for wp_attr in (f"{WP14}anchorId", f"{WP14}editId"):
            if wp_attr in el.attrib:
                el.attrib[wp_attr] = alloc.alloc_para_id()

        if tag == f"{W}pStyle" or tag == f"{W}rStyle" or tag == f"{W}tblStyle":
            old_val = el.get(f"{W}val")
            if old_val and old_val in style_map:
                el.set(f"{W}val", style_map[old_val])

        elif tag == f"{W}numId":
            old_val = el.get(f"{W}val")
            if old_val and old_val != "0" and old_val in num_map:
                el.set(f"{W}val", num_map[old_val])

        elif tag == f"{W}p":
            if normal_style_override is not None:
                ppr = el.find(f"{W}pPr")
                has_pstyle = ppr is not None and ppr.find(f"{W}pStyle") is not None
                if not has_pstyle:
                    # Only apply cloned Normal if not inside a table with an explicit tblStyle
                    inside_styled_tbl = False
                    parent = el.getparent()
                    if parent is not None and parent.tag == f"{W}tc":
                        # Walk up to enclosing w:tbl
                        cur = parent.getparent()
                        while cur is not None and cur.tag != f"{W}tbl":
                            cur = cur.getparent()
                        if cur is not None:
                            tbl_pr = cur.find(f"{W}tblPr")
                            if tbl_pr is not None and tbl_pr.find(f"{W}tblStyle") is not None:
                                inside_styled_tbl = True
                    if not inside_styled_tbl:
                        if ppr is None:
                            ppr = etree.Element(f"{W}pPr")
                            el.insert(0, ppr)
                        ps_el = etree.Element(f"{W}pStyle")
                        ps_el.set(f"{W}val", normal_style_override)
                        ppr.insert(0, ps_el)

            if strip_vanish_outline:
                ppr = el.find(f"{W}pPr")
                if ppr is not None:
                    ol = ppr.find(f"{W}outlineLvl")
                    if ol is not None:
                        rpr = ppr.find(f"{W}rPr")
                        has_vanish = rpr is not None and rpr.find(f"{W}vanish") is not None
                        if has_vanish and not "".join(el.itertext()).strip():
                            ppr.remove(ol)

        elif tag == f"{W}bookmarkStart":
            bname = el.get(f"{W}name") or ""
            bid = el.get(f"{W}id")
            if bname == "_GoBack":
                if bid is not None:
                    goback_ids.add(bid)
                parent = el.getparent()
                if parent is not None:
                    to_remove.append((parent, el))
                continue
            if bid is not None:
                new_bid = bm_id_map.get(bid)
                if new_bid is None:
                    new_bid = str(alloc.alloc_int("bm_id", 0))
                    bm_id_map[bid] = new_bid
                el.set(f"{W}id", new_bid)
            if bname:
                new_bname = bm_name_map.get(bname)
                if new_bname is None:
                    new_bname = alloc.alloc_bookmark_name(doc_idx, bname)
                    bm_name_map[bname] = new_bname
                el.set(f"{W}name", new_bname)

        elif tag == f"{W}bookmarkEnd":
            bid = el.get(f"{W}id")
            if bid in goback_ids:
                parent = el.getparent()
                if parent is not None:
                    to_remove.append((parent, el))
                continue
            if bid is not None:
                new_bid = bm_id_map.get(bid)
                if new_bid is None:
                    new_bid = str(alloc.alloc_int("bm_id", 0))
                    bm_id_map[bid] = new_bid
                el.set(f"{W}id", new_bid)

        elif tag == f"{W}hyperlink":
            if el.get(f"{W}anchor"):
                anchor_hyperlinks.append(el)

        elif tag == f"{W}instrText":
            field_instr_nodes.append(el)

        elif tag == f"{W}fldSimple":
            if el.get(f"{W}instr"):
                field_simple_nodes.append(el)

        elif tag in (f"{W}commentRangeStart", f"{W}commentRangeEnd", f"{W}commentReference"):
            cid = el.get(f"{W}id")
            if cid and cid in comment_id_map:
                el.set(f"{W}id", comment_id_map[cid])

        elif tag == f"{W}footnoteReference":
            fid = el.get(f"{W}id")
            if fid and fid in footnote_id_map:
                el.set(f"{W}id", footnote_id_map[fid])

        elif tag == f"{W}endnoteReference":
            eid = el.get(f"{W}id")
            if eid and eid in endnote_id_map:
                el.set(f"{W}id", endnote_id_map[eid])

        elif tag in _DRAWING_ID_TAGS or etree.QName(tag).localname in ("docPr", "cNvPr"):
            did = el.get("id")
            if did and did.isdigit():
                el.set("id", str(alloc.alloc_int("drawing_id", 1)))

        elif tag in _VML_SHAPE_TAGS:
            vid = el.get("id")
            if vid:
                new_vid = vml_id_map.get(vid)
                if new_vid is None:
                    new_vid = alloc.alloc_prefixed("vml_id", f"d{doc_idx}", vid)
                    vml_id_map[vid] = new_vid
                el.set("id", new_vid)
            spid = el.get(f"{O}spid")
            if spid:
                new_spid = vml_spid_map.get(spid)
                if new_spid is None:
                    new_spid = alloc.alloc_prefixed("vml_spid", f"d{doc_idx}", spid)
                    vml_spid_map[spid] = new_spid
                el.set(f"{O}spid", new_spid)
            tref = el.get("type")
            if tref and tref.startswith("#"):
                vml_type_refs.append(el)

        elif tag == f"{O}OLEObject":
            shid = el.get("ShapeID")
            if shid:
                new_shid = vml_id_map.get(shid)
                if new_shid is None:
                    new_shid = alloc.alloc_prefixed("vml_id", f"d{doc_idx}", shid)
                    vml_id_map[shid] = new_shid
                el.set("ShapeID", new_shid)
            oid = el.get("ObjectID")
            if oid:
                new_oid = ole_obj_id_map.get(oid)
                if new_oid is None:
                    new_oid = alloc.alloc_prefixed("ole_obj_id", f"_d{doc_idx}", oid)
                    ole_obj_id_map[oid] = new_oid
                el.set("ObjectID", new_oid)

    # Deferred removals (_GoBack bookmarks)
    for parent, child in to_remove:
        try:
            parent.remove(child)
        except ValueError:
            pass

    # Update VML shape/@type references (#shapetype_id)
    for el in vml_type_refs:
        old_ref = el.get("type")[1:]
        if old_ref in vml_id_map:
            el.set("type", f"#{vml_id_map[old_ref]}")

    # Update hyperlink anchors & REF/PAGEREF/HYPERLINK \l bookmark targets
    for hl in anchor_hyperlinks:
        anch = hl.get(f"{W}anchor")
        if anch and anch in bm_name_map:
            hl.set(f"{W}anchor", bm_name_map[anch])

    for instr_el in field_instr_nodes:
        txt = instr_el.text or ""
        m = _ANCHOR_FIELD_RE.match(txt)
        if m and m.group(3) in bm_name_map:
            instr_el.text = f"{m.group(1)}{m.group(2)}{bm_name_map[m.group(3)]}{m.group(4)}{m.group(5)}"
        elif "HYPERLINK" in txt.upper() and "\\l" in txt.lower():
            def _repl_hl(hm: re.Match[str]) -> str:
                name = hm.group(3)
                mapped = bm_name_map.get(name, name)
                return f"{hm.group(1)}{hm.group(2)}{mapped}{hm.group(4)}"
            instr_el.text = _HYPERLINK_L_FIELD_RE.sub(_repl_hl, txt)

    for fs_el in field_simple_nodes:
        txt = fs_el.get(f"{W}instr") or ""
        m = _ANCHOR_FIELD_RE.match(txt)
        if m and m.group(3) in bm_name_map:
            fs_el.set(f"{W}instr", f"{m.group(1)}{m.group(2)}{bm_name_map[m.group(3)]}{m.group(4)}{m.group(5)}")


def merge_docx_files(
    input_paths: Sequence[str],
    output_path: str,
    *,
    sort_by_name: bool = True,
    section_break_mode: str = "auto",
    insert_toc: bool = False,
    toc_levels: str = "1-3",
    strip_vanish_outline: Optional[bool] = None,
    progress_callback: Optional[ProgressCallback] = None,
    file_status_callback: Optional[FileStatusCallback] = None,
) -> MergeResult:
    """Merge `.docx` files end-to-end with strict content & formatting fidelity."""
    if not input_paths:
        raise ValueError("No input .docx files provided.")
    if section_break_mode not in ("auto", "page", "continuous"):
        raise ValueError(f"Unsupported section_break_mode: {section_break_mode}")

    ordered_inputs = sort_docx_paths(input_paths) if sort_by_name else list(input_paths)
    total_docs = len(ordered_inputs)
    # Empirical check with Word showed Word already ignores empty <w:vanish/> paragraphs in TOCs.
    should_strip_vanish_outline = bool(strip_vanish_outline) if strip_vanish_outline is not None else False

    if file_status_callback:
        file_status_callback(0, "processing")
    if progress_callback:
        progress_callback(f"读取基础文档 (1/{total_docs}): {os.path.basename(ordered_inputs[0])}", 0, total_docs)

    # Read doc 0 as the base package
    with zipfile.ZipFile(ordered_inputs[0], "r") as z0:
        parts: Dict[str, bytes] = {name: z0.read(name) for name in z0.namelist()}

    parsed_rels: Dict[str, etree._Element] = {}
    ct_root = _parse_xml(parts["[Content_Types].xml"])
    base_doc_root = _parse_xml(parts["word/document.xml"])
    base_body = base_doc_root.find(f"{W}body")
    if base_body is None:
        raise ValueError(f"Invalid .docx (missing w:body): {ordered_inputs[0]}")

    base_styles_root = _parse_xml(parts["word/styles.xml"]) if "word/styles.xml" in parts else None
    base_num_root = _parse_xml(parts["word/numbering.xml"]) if "word/numbering.xml" in parts else None
    base_fonts_root = _parse_xml(parts["word/fontTable.xml"]) if "word/fontTable.xml" in parts else None
    base_comments_root = _parse_xml(parts["word/comments.xml"]) if "word/comments.xml" in parts else None
    base_comments_ext_root = _parse_xml(parts["word/commentsExtended.xml"]) if "word/commentsExtended.xml" in parts else None
    base_comments_ids_root = _parse_xml(parts["word/commentsIds.xml"]) if "word/commentsIds.xml" in parts else None
    base_people_root = _parse_xml(parts["word/people.xml"]) if "word/people.xml" in parts else None
    base_footnotes_root = _parse_xml(parts["word/footnotes.xml"]) if "word/footnotes.xml" in parts else None
    base_endnotes_root = _parse_xml(parts["word/endnotes.xml"]) if "word/endnotes.xml" in parts else None

    # Track union nsmap and mc:Ignorable across all input word/document.xml roots
    union_nsmap: Dict[Optional[str], str] = dict(base_doc_root.nsmap)
    ignorable_tokens: List[str] = []
    for tok in (base_doc_root.get(f"{MC}Ignorable") or "").split():
        if tok and tok not in ignorable_tokens:
            ignorable_tokens.append(tok)

    alloc = _IdAllocator()
    _seed_base_identifiers(
        alloc,
        parts,
        base_doc_root,
        base_styles_root,
        base_num_root,
        base_comments_root,
        base_comments_ids_root,
        base_footnotes_root,
        base_endnotes_root,
    )

    # Index merged styles for semantic deduplication
    merged_styles_by_id: Dict[str, etree._Element] = {}
    merged_style_fp_to_id: Dict[str, str] = {}
    if base_styles_root is not None:
        for st in base_styles_root.findall(f"{W}style"):
            sid = st.get(f"{W}styleId")
            if sid:
                merged_styles_by_id[sid] = st
        # Compute fingerprints in topological order
        for sid in _style_topological_order(merged_styles_by_id, set(merged_styles_by_id.keys())):
            st = merged_styles_by_id[sid]
            based_el = st.find(f"{W}basedOn")
            based_id = based_el.get(f"{W}val") if based_el is not None else None
            fp = _clean_style_for_fingerprint(st, based_id)
            merged_style_fp_to_id.setdefault(fp, sid)

    base_doc_defaults_sig = _clean_doc_defaults(base_styles_root)
    base_normal_el = _get_default_style(base_styles_root, "paragraph")
    base_normal_sig = (
        _clean_style_for_fingerprint(base_normal_el, None) if base_normal_el is not None else ""
    )

    existing_fonts: Set[str] = set()
    if base_fonts_root is not None:
        for fn in base_fonts_root.findall(f"{W}font"):
            fname = fn.get(f"{W}name")
            if fname:
                existing_fonts.add(fname)

    existing_people: Set[str] = set()
    if base_people_root is not None:
        for person in base_people_root.findall(f"{W15}person"):
            author = person.get(f"{W15}author")
            if author:
                existing_people.add(author)

    # Helper to create empty header/footer parts on demand (prevents leakage across sections)
    empty_hf_rids: Dict[bool, Optional[str]] = {True: None, False: None}

    def _ensure_empty_hf(is_header: bool) -> str:
        cached = empty_hf_rids[is_header]
        if cached is not None:
            return cached
        part_name = "word/header_empty.xml" if is_header else "word/footer_empty.xml"
        if part_name in parts:
            part_name = _alloc_unique_part_path(parts, "word", 0, "header_empty.xml" if is_header else "footer_empty.xml")
        parts[part_name] = _empty_header_footer_xml(is_header)
        _ensure_content_type_override(ct_root, part_name, CT_HEADER if is_header else CT_FOOTER)
        doc_rels = _get_or_create_rels(parts, parsed_rels, "word/_rels/document.xml.rels")
        rid = _alloc_rid(doc_rels)
        rel_el = etree.SubElement(doc_rels, f"{{{NS_REL}}}Relationship")
        rel_el.set("Id", rid)
        rel_el.set("Type", REL_TYPE_HEADER if is_header else REL_TYPE_FOOTER)
        rel_el.set("Target", _relative_target("word/document.xml", part_name))
        empty_hf_rids[is_header] = rid
        return rid

    # Extract doc 0 body elements and trailing sectPr
    doc0_children = list(base_body)
    base_body.clear()  # Keep attributes on base_body if any, remove children to re-append cleanly

    if insert_toc:
        for el in _build_toc_sdt_and_page_break(toc_levels=toc_levels):
            base_body.append(el)

    doc0_sect_pr: Optional[etree._Element] = None
    if doc0_children and doc0_children[-1].tag == f"{W}sectPr":
        doc0_sect_pr = doc0_children.pop()
    else:
        doc0_sect_pr = etree.Element(f"{W}sectPr")

    doc0_rels_root = _get_or_create_rels(parts, parsed_rels, "word/_rels/document.xml.rels")
    doc0_rels_map = _rels_map(doc0_rels_root)

    # Append doc 0 children while tracking section header/footer isolation across any internal sections
    prev_hdr_types: Set[str] = set()
    prev_ftr_types: Set[str] = set()
    section_count = 1

    for child in doc0_children:
        if child.tag == f"{W}p":
            sp = child.find(f"{W}pPr/{W}sectPr")
            if sp is not None:
                prev_hdr_types, prev_ftr_types = _enforce_hf_isolation(
                    prev_hdr_types, prev_ftr_types, sp, _ensure_empty_hf
                )
                section_count += 1
        base_body.append(child)

    prev_hdr_types, prev_ftr_types = _enforce_hf_isolation(
        prev_hdr_types, prev_ftr_types, doc0_sect_pr, _ensure_empty_hf
    )
    prev_final_sect_pr = doc0_sect_pr
    prev_final_sig, _, _ = _extract_sect_signature(
        doc0_sect_pr, doc0_rels_map, None, "word/document.xml", parts=parts
    )

    comments_merged = 0
    footnotes_merged = 0
    endnotes_merged = 0

    # Helper to copy/remap auxiliary XML parts (headers, footers, comments, footnotes, endnotes)
    def _copy_binary_part(src_zip: zipfile.ZipFile, src_part: str, target_dir: str, doc_idx: int) -> str:
        raw = src_zip.read(src_part)
        new_path = _alloc_unique_part_path(parts, target_dir, doc_idx, posixpath.basename(src_part))
        parts[new_path] = raw
        return new_path

    def _copy_rels_and_targets(
        src_zip: zipfile.ZipFile,
        src_part: str,
        dst_part: str,
        doc_idx: int,
        used_rids: Optional[Set[str]] = None,
        skip_types: Optional[Set[str]] = None,
    ) -> Dict[str, str]:
        """Copy rels for `src_part` -> `dst_part`, cloning referenced internal media/embeddings/parts."""
        src_rels_path = _rels_path_for(src_part)
        if src_rels_path not in src_zip.namelist():
            return {}
        src_rels_root = _parse_xml(src_zip.read(src_rels_path))
        dst_rels_path = _rels_path_for(dst_part)
        dst_rels_root = _get_or_create_rels(parts, parsed_rels, dst_rels_path)
        rmap: Dict[str, str] = {}

        for rel in src_rels_root:
            old_rid = rel.get("Id")
            rtype = rel.get("Type") or ""
            target = rel.get("Target") or ""
            tmode = rel.get("TargetMode")
            if not old_rid:
                continue
            if skip_types and rtype in skip_types:
                continue
            if used_rids is not None and old_rid not in used_rids:
                continue

            used_dst_rids = {el.get("Id") for el in dst_rels_root if el.get("Id")}
            new_rid = old_rid if old_rid not in used_dst_rids else _alloc_rid(dst_rels_root)
            new_rel = etree.SubElement(dst_rels_root, f"{{{NS_REL}}}Relationship")
            new_rel.set("Id", new_rid)
            new_rel.set("Type", rtype)

            if tmode == "External":
                new_rel.set("Target", target)
                new_rel.set("TargetMode", "External")
            else:
                abs_src = _resolve_target(src_part, target)
                if abs_src not in src_zip.namelist():
                    # Keep reference best-effort if missing
                    new_rel.set("Target", target)
                else:
                    src_dir = posixpath.dirname(abs_src)
                    new_target = _copy_binary_part(src_zip, abs_src, src_dir, doc_idx)
                    new_rel.set("Target", _relative_target(dst_part, new_target))
                    # Preserve content type override if source part had one
                    if abs_src.endswith(".xml") and not abs_src.endswith(".rels"):
                        src_ct = _parse_xml(src_zip.read("[Content_Types].xml"))
                        for ov in src_ct:
                            if ov.tag == f"{{{NS_CT}}}Override" and ov.get("PartName") == "/" + abs_src:
                                _ensure_content_type_override(ct_root, new_target, ov.get("ContentType") or "")
                    # Recursively copy sub-part rels if any (e.g. chart -> workbook)
                    if _rels_path_for(abs_src) in src_zip.namelist():
                        sub_rmap = _copy_rels_and_targets(src_zip, abs_src, new_target, doc_idx)
                        if abs_src.endswith(".xml") and any(k != v for k, v in sub_rmap.items()):
                            sub_root = _parse_xml(parts[new_target])
                            for sub_el in sub_root.iter():
                                for ak in tuple(sub_el.attrib.keys()):
                                    if ak.startswith(R) and sub_el.attrib[ak] in sub_rmap:
                                        sub_el.attrib[ak] = sub_rmap[sub_el.attrib[ak]]
                            parts[new_target] = _serialize_xml(sub_root)

            rmap[old_rid] = new_rid
        return rmap

    # Process subsequent documents (k = 1 .. N-1)
    if file_status_callback:
        file_status_callback(0, "done")
    if progress_callback:
        progress_callback(f"基础文档加载完成 (1/{total_docs}): {os.path.basename(ordered_inputs[0])}", 1, total_docs)

    for doc_idx, src_path in enumerate(ordered_inputs[1:], start=1):
        if file_status_callback:
            file_status_callback(doc_idx, "processing")
        if progress_callback:
            progress_callback(f"合并文档 ({doc_idx + 1}/{total_docs}): {os.path.basename(src_path)}", doc_idx, total_docs)

        with zipfile.ZipFile(src_path, "r") as zf:
            znames = set(zf.namelist())
            src_ct_root = _parse_xml(zf.read("[Content_Types].xml"))
            _merge_content_type_defaults(ct_root, src_ct_root)

            src_doc_root = _parse_xml(zf.read("word/document.xml"))
            ns_added = False
            for k, v in src_doc_root.nsmap.items():
                if k not in union_nsmap:
                    union_nsmap[k] = v
                    ns_added = True
            for tok in (src_doc_root.get(f"{MC}Ignorable") or "").split():
                if tok and tok not in ignorable_tokens:
                    ignorable_tokens.append(tok)
            if ns_added:
                new_doc_root = etree.Element(base_doc_root.tag, attrib=base_doc_root.attrib, nsmap=union_nsmap)
                for ch in list(base_doc_root):
                    new_doc_root.append(ch)
                base_doc_root = new_doc_root
                base_body = base_doc_root.find(f"{W}body")

            src_body = src_doc_root.find(f"{W}body")
            if src_body is None:
                raise ValueError(f"Invalid .docx (missing w:body): {src_path}")

            src_styles_root = _parse_xml(zf.read("word/styles.xml")) if "word/styles.xml" in znames else None
            src_num_root = _parse_xml(zf.read("word/numbering.xml")) if "word/numbering.xml" in znames else None
            src_fonts_root = _parse_xml(zf.read("word/fontTable.xml")) if "word/fontTable.xml" in znames else None
            src_comments_root = _parse_xml(zf.read("word/comments.xml")) if "word/comments.xml" in znames else None
            src_comments_ext_root = _parse_xml(zf.read("word/commentsExtended.xml")) if "word/commentsExtended.xml" in znames else None
            src_comments_ids_root = _parse_xml(zf.read("word/commentsIds.xml")) if "word/commentsIds.xml" in znames else None
            src_people_root = _parse_xml(zf.read("word/people.xml")) if "word/people.xml" in znames else None
            src_footnotes_root = _parse_xml(zf.read("word/footnotes.xml")) if "word/footnotes.xml" in znames else None
            src_endnotes_root = _parse_xml(zf.read("word/endnotes.xml")) if "word/endnotes.xml" in znames else None

            src_doc_rels_path = "word/_rels/document.xml.rels"
            src_doc_rels_root = _parse_xml(zf.read(src_doc_rels_path)) if src_doc_rels_path in znames else None
            src_doc_rels_by_id = _rels_map(src_doc_rels_root)

            # Collect referenced rIds and header/footer XML roots so we know all used styles/numIds
            src_doc_used_rids: Set[str] = set()
            for el in src_doc_root.iter():
                for ak, av in el.attrib.items():
                    if ak.startswith(R) and av:
                        src_doc_used_rids.add(av)

            # Preload referenced header/footer parts (only those referenced by sectPr in src_body)
            src_hf_parts: Dict[str, Tuple[bool, etree._Element]] = {}  # old_rid -> (is_header, root)
            for el in src_body.iter():
                if el.tag in (f"{W}headerReference", f"{W}footerReference"):
                    rid = el.get(f"{R}id")
                    if rid and rid in src_doc_rels_by_id and rid not in src_hf_parts:
                        t_path = _resolve_target("word/document.xml", src_doc_rels_by_id[rid].get("Target") or "")
                        if t_path in znames:
                            src_hf_parts[rid] = (el.tag == f"{W}headerReference", _parse_xml(zf.read(t_path)))

            # 1. Font table merge
            if src_fonts_root is not None:
                if base_fonts_root is None:
                    base_fonts_root = copy.deepcopy(src_fonts_root)
                else:
                    for fn in src_fonts_root.findall(f"{W}font"):
                        fname = fn.get(f"{W}name")
                        if fname and fname not in existing_fonts:
                            base_fonts_root.append(copy.deepcopy(fn))
                            existing_fonts.add(fname)

            # 2. Numbering isolation (allocate abstractNumId & numId FIRST)
            abs_num_map: Dict[str, str] = {}
            num_map: Dict[str, str] = {}
            copied_abstract_nums: List[etree._Element] = []
            if src_num_root is not None:
                if base_num_root is None:
                    base_num_root = etree.Element(f"{W}numbering", nsmap={"w": NS_W, "r": NS_R, "w14": NS_W14})
                    doc0_rels = _get_or_create_rels(parts, parsed_rels, "word/_rels/document.xml.rels")
                    if _find_rel_by_type(doc0_rels, REL_TYPE_NUMBERING) is None:
                        nrid = _alloc_rid(doc0_rels)
                        nrel = etree.SubElement(doc0_rels, f"{{{NS_REL}}}Relationship")
                        nrel.set("Id", nrid)
                        nrel.set("Type", REL_TYPE_NUMBERING)
                        nrel.set("Target", "numbering.xml")
                    _ensure_content_type_override(ct_root, "word/numbering.xml", CT_NUMBERING)

                # Copy numbering.xml.rels (for numPicBullet if any)
                num_rel_map = _copy_rels_and_targets(zf, "word/numbering.xml", "word/numbering.xml", doc_idx)

                # Remap numPicBulletIds if present
                pic_bullet_map: Dict[str, str] = {}
                existing_pb_ids = {
                    int(pb.get(f"{W}numPicBulletId"))
                    for pb in base_num_root.findall(f"{W}numPicBullet")
                    if (pb.get(f"{W}numPicBulletId") or "").isdigit()
                }
                next_pb_id = max(existing_pb_ids, default=-1) + 1
                first_abs_idx = 0
                for idx, ch in enumerate(list(base_num_root)):
                    if ch.tag != f"{W}numPicBullet":
                        first_abs_idx = idx
                        break
                    first_abs_idx = idx + 1

                for pb in src_num_root.findall(f"{W}numPicBullet"):
                    old_pb = pb.get(f"{W}numPicBulletId")
                    new_pb = str(next_pb_id)
                    next_pb_id += 1
                    if old_pb is not None:
                        pic_bullet_map[old_pb] = new_pb
                    pb_clone = copy.deepcopy(pb)
                    pb_clone.set(f"{W}numPicBulletId", new_pb)
                    for el in pb_clone.iter():
                        for ak in tuple(el.attrib.keys()):
                            if ak.startswith(R) and el.attrib[ak] in num_rel_map:
                                el.attrib[ak] = num_rel_map[el.attrib[ak]]
                    base_num_root.insert(first_abs_idx, pb_clone)
                    first_abs_idx += 1

                # Find insertion index for abstractNum (MUST precede all w:num!)
                abs_insert_idx = len(base_num_root)
                for idx, ch in enumerate(list(base_num_root)):
                    if ch.tag == f"{W}num":
                        abs_insert_idx = idx
                        break

                for an in src_num_root.findall(f"{W}abstractNum"):
                    old_aid = an.get(f"{W}abstractNumId")
                    if old_aid is None:
                        continue
                    new_aid = str(alloc.alloc_int("abstractNum", 1))
                    abs_num_map[old_aid] = new_aid
                    an_clone = copy.deepcopy(an)
                    an_clone.set(f"{W}abstractNumId", new_aid)
                    nsid = an_clone.find(f"{W}nsid")
                    if nsid is not None:
                        nsid.set(f"{W}val", alloc.alloc_para_id())
                    for pb_ref in an_clone.iter(f"{W}lvlPicBulletId"):
                        v = pb_ref.get(f"{W}val")
                        if v in pic_bullet_map:
                            pb_ref.set(f"{W}val", pic_bullet_map[v])
                    base_num_root.insert(abs_insert_idx, an_clone)
                    abs_insert_idx += 1
                    copied_abstract_nums.append(an_clone)

                for num in src_num_root.findall(f"{W}num"):
                    old_nid = num.get(f"{W}numId")
                    if not old_nid or old_nid == "0":
                        continue
                    new_nid = str(alloc.alloc_int("numId", 1))
                    num_map[old_nid] = new_nid
                    num_clone = copy.deepcopy(num)
                    num_clone.set(f"{W}numId", new_nid)
                    abs_ref = num_clone.find(f"{W}abstractNumId")
                    if abs_ref is not None:
                        old_abs_ref = abs_ref.get(f"{W}val")
                        if old_abs_ref in abs_num_map:
                            abs_ref.set(f"{W}val", abs_num_map[old_abs_ref])
                    # Insert after w:num entries (before numIdMacAtCleanup if any)
                    num_insert_idx = len(base_num_root)
                    for idx, ch in enumerate(list(base_num_root)):
                        if ch.tag == f"{W}numIdMacAtCleanup":
                            num_insert_idx = idx
                            break
                    base_num_root.insert(num_insert_idx, num_clone)

            # 3. Styles resolution (topological order + semantic deduplication + collision-safe renaming)
            style_map: Dict[str, str] = {}
            normal_style_override: Optional[str] = None
            if src_styles_root is not None:
                if base_styles_root is None:
                    base_styles_root = copy.deepcopy(src_styles_root)
                else:
                    src_styles_by_id: Dict[str, etree._Element] = {}
                    for st in src_styles_root.findall(f"{W}style"):
                        sid = st.get(f"{W}styleId")
                        if sid:
                            src_styles_by_id[sid] = st

                    xml_roots_for_styles: List[etree._Element] = [src_doc_root]
                    for _, hf_root in src_hf_parts.values():
                        xml_roots_for_styles.append(hf_root)
                    for aux in (src_comments_root, src_footnotes_root, src_endnotes_root):
                        if aux is not None:
                            xml_roots_for_styles.append(aux)
                    if src_num_root is not None:
                        xml_roots_for_styles.append(src_num_root)

                    used_style_ids = _collect_used_styles_closure(src_styles_by_id, xml_roots_for_styles)

                    # Check if Normal or docDefaults differs from base
                    src_dd_sig = _clean_doc_defaults(src_styles_root)
                    src_normal_el = _get_default_style(src_styles_root, "paragraph")
                    src_normal_sig = (
                        _clean_style_for_fingerprint(src_normal_el, None) if src_normal_el is not None else ""
                    )
                    if src_normal_el is not None and (
                        src_normal_sig != base_normal_sig or src_dd_sig != base_doc_defaults_sig
                    ):
                        normal_id = src_normal_el.get(f"{W}styleId")
                        if normal_id:
                            used_style_ids.add(normal_id)

                    added_styles: List[Tuple[str, etree._Element, bool]] = []  # (old_id, new_el, is_new)
                    for old_sid in _style_topological_order(src_styles_by_id, used_style_ids):
                        st = copy.deepcopy(src_styles_by_id[old_sid])
                        # Remap numIds inside style first!
                        for nid_el in st.iter(f"{W}numId"):
                            nv = nid_el.get(f"{W}val")
                            if nv and nv != "0" and nv in num_map:
                                nid_el.set(f"{W}val", num_map[nv])

                        based_el = st.find(f"{W}basedOn")
                        resolved_based = None
                        if based_el is not None:
                            bv = based_el.get(f"{W}val")
                            if bv:
                                resolved_based = style_map.get(bv, bv)
                                based_el.set(f"{W}val", resolved_based)

                        # If docDefaults differs from base and this is a root style (no basedOn),
                        # we leave explicit style properties intact; for Normal when differing, check override
                        is_src_default_normal = (
                            src_normal_el is not None and old_sid == src_normal_el.get(f"{W}styleId")
                        )

                        fp = _clean_style_for_fingerprint(st, resolved_based)
                        if fp in merged_style_fp_to_id and not (
                            is_src_default_normal and src_dd_sig != base_doc_defaults_sig
                        ):
                            mapped_id = merged_style_fp_to_id[fp]
                            style_map[old_sid] = mapped_id
                            if is_src_default_normal and (
                                src_normal_sig != base_normal_sig or src_dd_sig != base_doc_defaults_sig
                            ):
                                normal_style_override = mapped_id
                            continue

                        # Need to add a new style to base_styles_root
                        new_sid = alloc.alloc_prefixed("styleId", f"d{doc_idx}", old_sid)
                        st.set(f"{W}styleId", new_sid)
                        st.attrib.pop(f"{W}default", None)

                        nm_el = st.find(f"{W}name")
                        orig_name = (nm_el.get(f"{W}val") or "") if nm_el is not None else old_sid
                        heading_match = _HEADING_NAME_RE.match(orig_name.strip())

                        if nm_el is not None and orig_name:
                            if alloc.has_str("styleName", orig_name.lower()):
                                new_name = f"{orig_name}_d{doc_idx}"
                                while alloc.has_str("styleName", new_name.lower()):
                                    new_name = f"{new_name}_x"
                                nm_el.set(f"{W}val", new_name)
                                st.set(f"{W}customStyle", "1")
                                alloc.seed_str("styleName", new_name.lower())
                                # Preserve outlineLvl for renamed built-in heading styles!
                                if heading_match:
                                    lvl_val = str(int(heading_match.group(1)) - 1)
                                    ppr = st.find(f"{W}pPr")
                                    if ppr is None:
                                        ppr = etree.Element(f"{W}pPr")
                                        st.append(ppr)
                                    if ppr.find(f"{W}outlineLvl") is None:
                                        ol = etree.SubElement(ppr, f"{W}outlineLvl")
                                        ol.set(f"{W}val", lvl_val)
                            else:
                                alloc.seed_str("styleName", orig_name.lower())

                        style_map[old_sid] = new_sid
                        merged_styles_by_id[new_sid] = st
                        merged_style_fp_to_id.setdefault(fp, new_sid)
                        base_styles_root.append(st)
                        added_styles.append((old_sid, st, True))
                        if is_src_default_normal and (
                            src_normal_sig != base_normal_sig or src_dd_sig != base_doc_defaults_sig
                        ):
                            normal_style_override = new_sid

                    # Second pass on newly added styles to resolve next / link cleanly
                    newly_added_ids = {st.get(f"{W}styleId") for _, st, _ in added_styles}
                    for _, st, _ in added_styles:
                        nxt = st.find(f"{W}next")
                        if nxt is not None:
                            nv = nxt.get(f"{W}val")
                            if nv and nv in style_map:
                                nxt.set(f"{W}val", style_map[nv])
                        lnk = st.find(f"{W}link")
                        if lnk is not None:
                            lv = lnk.get(f"{W}val")
                            mapped_lv = style_map.get(lv, "") if lv else ""
                            if mapped_lv in newly_added_ids:
                                lnk.set(f"{W}val", mapped_lv)
                            else:
                                # Avoid asymmetric link to an existing base style
                                st.remove(lnk)

            # Now remap <w:pStyle> inside copied abstractNum levels using completed style_map!
            for an_clone in copied_abstract_nums:
                for ps in an_clone.iter(f"{W}pStyle"):
                    old_ps = ps.get(f"{W}val")
                    if old_ps and old_ps in style_map:
                        ps.set(f"{W}val", style_map[old_ps])

            # 4. Copy referenced binary/external document.xml relationships (images, OLE, hyperlinks, etc.)
            skip_doc_rel_types = {
                REL_TYPE_STYLES,
                REL_TYPE_NUMBERING,
                REL_TYPE_FONT_TABLE,
                REL_TYPE_SETTINGS,
                REL_TYPE_FOOTNOTES,
                REL_TYPE_ENDNOTES,
                REL_TYPE_COMMENTS,
                REL_TYPE_COMMENTS_EXT,
                REL_TYPE_COMMENTS_IDS,
                REL_TYPE_PEOPLE,
                REL_TYPE_HEADER,
                REL_TYPE_FOOTER,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
            }
            doc_rel_map = _copy_rels_and_targets(
                zf,
                "word/document.xml",
                "word/document.xml",
                doc_idx,
                used_rids=src_doc_used_rids,
                skip_types=skip_doc_rel_types,
            )

            # Per-doc ID maps shared between document.xml, comments, footnotes, endnotes, headers/footers
            bm_id_map: Dict[str, str] = {}
            bm_name_map: Dict[str, str] = {}
            vml_id_map: Dict[str, str] = {}
            vml_spid_map: Dict[str, str] = {}
            ole_obj_id_map: Dict[str, str] = {}
            para_id_map: Dict[str, str] = {}
            comment_id_map: Dict[str, str] = {}
            footnote_id_map: Dict[str, str] = {}
            endnote_id_map: Dict[str, str] = {}

            # 5. Merge comments (+ commentsExtended, commentsIds, people)
            if src_comments_root is not None:
                src_real_comments = src_comments_root.findall(f"{W}comment")
                if src_real_comments:
                    if base_comments_root is None:
                        base_comments_root = etree.Element(f"{W}comments", nsmap=src_comments_root.nsmap)
                        doc0_rels = _get_or_create_rels(parts, parsed_rels, "word/_rels/document.xml.rels")
                        if _find_rel_by_type(doc0_rels, REL_TYPE_COMMENTS) is None:
                            crid = _alloc_rid(doc0_rels)
                            crel = etree.SubElement(doc0_rels, f"{{{NS_REL}}}Relationship")
                            crel.set("Id", crid)
                            crel.set("Type", REL_TYPE_COMMENTS)
                            crel.set("Target", "comments.xml")
                        _ensure_content_type_override(ct_root, "word/comments.xml", CT_COMMENTS)

                    cm_rel_map = _copy_rels_and_targets(zf, "word/comments.xml", "word/comments.xml", doc_idx)
                    for cm in src_real_comments:
                        old_cid = cm.get(f"{W}id")
                        if old_cid is None:
                            continue
                        new_cid = str(alloc.alloc_int("comment", 0))
                        comment_id_map[old_cid] = new_cid
                        cm_clone = copy.deepcopy(cm)
                        cm_clone.set(f"{W}id", new_cid)
                        _remap_single_pass(
                            cm_clone,
                            doc_idx=doc_idx,
                            alloc=alloc,
                            style_map=style_map,
                            num_map=num_map,
                            rel_map=cm_rel_map,
                            comment_id_map={},
                            footnote_id_map={},
                            endnote_id_map={},
                            bm_id_map=bm_id_map,
                            bm_name_map=bm_name_map,
                            vml_id_map=vml_id_map,
                            vml_spid_map=vml_spid_map,
                            ole_obj_id_map=ole_obj_id_map,
                            para_id_map=para_id_map,
                            normal_style_override=None,
                        )
                        base_comments_root.append(cm_clone)
                        comments_merged += 1

                    if src_comments_ext_root is not None:
                        if base_comments_ext_root is None:
                            base_comments_ext_root = etree.Element(
                                f"{W15}commentsEx", nsmap=src_comments_ext_root.nsmap
                            )
                            doc0_rels = _get_or_create_rels(parts, parsed_rels, "word/_rels/document.xml.rels")
                            if _find_rel_by_type(doc0_rels, REL_TYPE_COMMENTS_EXT) is None:
                                r = etree.SubElement(doc0_rels, f"{{{NS_REL}}}Relationship")
                                r.set("Id", _alloc_rid(doc0_rels))
                                r.set("Type", REL_TYPE_COMMENTS_EXT)
                                r.set("Target", "commentsExtended.xml")
                            _ensure_content_type_override(ct_root, "word/commentsExtended.xml", CT_COMMENTS_EXT)
                        for cex in src_comments_ext_root.findall(f"{W15}commentEx"):
                            cex_clone = copy.deepcopy(cex)
                            pid = cex_clone.get(f"{W15}paraId")
                            if pid and pid in para_id_map:
                                cex_clone.set(f"{W15}paraId", para_id_map[pid])
                            ppid = cex_clone.get(f"{W15}paraIdParent")
                            if ppid and ppid in para_id_map:
                                cex_clone.set(f"{W15}paraIdParent", para_id_map[ppid])
                            base_comments_ext_root.append(cex_clone)

                    if src_comments_ids_root is not None:
                        if base_comments_ids_root is None:
                            base_comments_ids_root = etree.Element(
                                f"{W16CID}commentsIds", nsmap=src_comments_ids_root.nsmap
                            )
                            doc0_rels = _get_or_create_rels(parts, parsed_rels, "word/_rels/document.xml.rels")
                            if _find_rel_by_type(doc0_rels, REL_TYPE_COMMENTS_IDS) is None:
                                r = etree.SubElement(doc0_rels, f"{{{NS_REL}}}Relationship")
                                r.set("Id", _alloc_rid(doc0_rels))
                                r.set("Type", REL_TYPE_COMMENTS_IDS)
                                r.set("Target", "commentsIds.xml")
                            _ensure_content_type_override(ct_root, "word/commentsIds.xml", CT_COMMENTS_IDS)
                        for cid_el in src_comments_ids_root.findall(f"{W16CID}commentId"):
                            cid_clone = copy.deepcopy(cid_el)
                            pid = cid_clone.get(f"{W16CID}paraId")
                            if pid and pid in para_id_map:
                                cid_clone.set(f"{W16CID}paraId", para_id_map[pid])
                            did = cid_clone.get(f"{W16CID}durableId")
                            if did and alloc.has_str("durableId", did.upper()):
                                new_did = alloc.alloc_para_id()
                                alloc.seed_str("durableId", new_did)
                                cid_clone.set(f"{W16CID}durableId", new_did)
                            elif did:
                                alloc.seed_str("durableId", did.upper())
                            base_comments_ids_root.append(cid_clone)

                    if src_people_root is not None:
                        if base_people_root is None:
                            base_people_root = etree.Element(f"{W15}people", nsmap=src_people_root.nsmap)
                            doc0_rels = _get_or_create_rels(parts, parsed_rels, "word/_rels/document.xml.rels")
                            if _find_rel_by_type(doc0_rels, REL_TYPE_PEOPLE) is None:
                                r = etree.SubElement(doc0_rels, f"{{{NS_REL}}}Relationship")
                                r.set("Id", _alloc_rid(doc0_rels))
                                r.set("Type", REL_TYPE_PEOPLE)
                                r.set("Target", "people.xml")
                            _ensure_content_type_override(ct_root, "word/people.xml", CT_PEOPLE)
                        for person in src_people_root.findall(f"{W15}person"):
                            author = person.get(f"{W15}author")
                            if author and author not in existing_people:
                                base_people_root.append(copy.deepcopy(person))
                                existing_people.add(author)

            # 6. Merge real footnotes & endnotes
            if src_footnotes_root is not None:
                real_fns = [
                    fn
                    for fn in src_footnotes_root.findall(f"{W}footnote")
                    if fn.get(f"{W}type") not in ("separator", "continuationSeparator")
                ]
                if real_fns:
                    if base_footnotes_root is None:
                        base_footnotes_root = copy.deepcopy(src_footnotes_root)
                        for fn in list(base_footnotes_root.findall(f"{W}footnote")):
                            if fn.get(f"{W}type") not in ("separator", "continuationSeparator"):
                                base_footnotes_root.remove(fn)
                        doc0_rels = _get_or_create_rels(parts, parsed_rels, "word/_rels/document.xml.rels")
                        if _find_rel_by_type(doc0_rels, REL_TYPE_FOOTNOTES) is None:
                            r = etree.SubElement(doc0_rels, f"{{{NS_REL}}}Relationship")
                            r.set("Id", _alloc_rid(doc0_rels))
                            r.set("Type", REL_TYPE_FOOTNOTES)
                            r.set("Target", "footnotes.xml")
                        _ensure_content_type_override(ct_root, "word/footnotes.xml", CT_FOOTNOTES)
                    fn_rel_map = _copy_rels_and_targets(zf, "word/footnotes.xml", "word/footnotes.xml", doc_idx)
                    for fn in real_fns:
                        old_fid = fn.get(f"{W}id")
                        if old_fid is None:
                            continue
                        new_fid = str(alloc.alloc_int("footnote", 1))
                        footnote_id_map[old_fid] = new_fid
                        fn_clone = copy.deepcopy(fn)
                        fn_clone.set(f"{W}id", new_fid)
                        _remap_single_pass(
                            fn_clone,
                            doc_idx=doc_idx,
                            alloc=alloc,
                            style_map=style_map,
                            num_map=num_map,
                            rel_map=fn_rel_map,
                            comment_id_map=comment_id_map,
                            footnote_id_map={},
                            endnote_id_map={},
                            bm_id_map=bm_id_map,
                            bm_name_map=bm_name_map,
                            vml_id_map=vml_id_map,
                            vml_spid_map=vml_spid_map,
                            ole_obj_id_map=ole_obj_id_map,
                            para_id_map=para_id_map,
                            normal_style_override=None,
                        )
                        base_footnotes_root.append(fn_clone)
                        footnotes_merged += 1

            if src_endnotes_root is not None:
                real_ens = [
                    en
                    for en in src_endnotes_root.findall(f"{W}endnote")
                    if en.get(f"{W}type") not in ("separator", "continuationSeparator")
                ]
                if real_ens:
                    if base_endnotes_root is None:
                        base_endnotes_root = copy.deepcopy(src_endnotes_root)
                        for en in list(base_endnotes_root.findall(f"{W}endnote")):
                            if en.get(f"{W}type") not in ("separator", "continuationSeparator"):
                                base_endnotes_root.remove(en)
                        doc0_rels = _get_or_create_rels(parts, parsed_rels, "word/_rels/document.xml.rels")
                        if _find_rel_by_type(doc0_rels, REL_TYPE_ENDNOTES) is None:
                            r = etree.SubElement(doc0_rels, f"{{{NS_REL}}}Relationship")
                            r.set("Id", _alloc_rid(doc0_rels))
                            r.set("Type", REL_TYPE_ENDNOTES)
                            r.set("Target", "endnotes.xml")
                        _ensure_content_type_override(ct_root, "word/endnotes.xml", CT_ENDNOTES)
                    en_rel_map = _copy_rels_and_targets(zf, "word/endnotes.xml", "word/endnotes.xml", doc_idx)
                    for en in real_ens:
                        old_eid = en.get(f"{W}id")
                        if old_eid is None:
                            continue
                        new_eid = str(alloc.alloc_int("endnote", 1))
                        endnote_id_map[old_eid] = new_eid
                        en_clone = copy.deepcopy(en)
                        en_clone.set(f"{W}id", new_eid)
                        _remap_single_pass(
                            en_clone,
                            doc_idx=doc_idx,
                            alloc=alloc,
                            style_map=style_map,
                            num_map=num_map,
                            rel_map=en_rel_map,
                            comment_id_map=comment_id_map,
                            footnote_id_map={},
                            endnote_id_map={},
                            bm_id_map=bm_id_map,
                            bm_name_map=bm_name_map,
                            vml_id_map=vml_id_map,
                            vml_spid_map=vml_spid_map,
                            ole_obj_id_map=ole_obj_id_map,
                            para_id_map=para_id_map,
                            normal_style_override=None,
                        )
                        base_endnotes_root.append(en_clone)
                        endnotes_merged += 1

            # 7. Copy referenced headers & footers (before remapping document.xml rIds!)
            doc0_rels = _get_or_create_rels(parts, parsed_rels, "word/_rels/document.xml.rels")
            for old_hf_rid, (is_hdr, hf_root) in src_hf_parts.items():
                src_hf_part = _resolve_target("word/document.xml", src_doc_rels_by_id[old_hf_rid].get("Target") or "")
                prefix = "header" if is_hdr else "footer"
                dst_hf_part = _alloc_unique_part_path(parts, "word", doc_idx, posixpath.basename(src_hf_part))
                hf_rel_map = _copy_rels_and_targets(zf, src_hf_part, dst_hf_part, doc_idx)
                hf_clone = copy.deepcopy(hf_root)
                _remap_single_pass(
                    hf_clone,
                    doc_idx=doc_idx,
                    alloc=alloc,
                    style_map=style_map,
                    num_map=num_map,
                    rel_map=hf_rel_map,
                    comment_id_map=comment_id_map,
                    footnote_id_map=footnote_id_map,
                    endnote_id_map=endnote_id_map,
                    bm_id_map=bm_id_map,
                    bm_name_map=bm_name_map,
                    vml_id_map=vml_id_map,
                    vml_spid_map=vml_spid_map,
                    ole_obj_id_map=ole_obj_id_map,
                    para_id_map=para_id_map,
                    normal_style_override=None,
                )
                parts[dst_hf_part] = _serialize_xml(hf_clone)
                _ensure_content_type_override(ct_root, dst_hf_part, CT_HEADER if is_hdr else CT_FOOTER)
                new_hf_rid = _alloc_rid(doc0_rels)
                hf_rel_el = etree.SubElement(doc0_rels, f"{{{NS_REL}}}Relationship")
                hf_rel_el.set("Id", new_hf_rid)
                hf_rel_el.set("Type", REL_TYPE_HEADER if is_hdr else REL_TYPE_FOOTER)
                hf_rel_el.set("Target", _relative_target("word/document.xml", dst_hf_part))
                doc_rel_map[old_hf_rid] = new_hf_rid

            # 8. Extract src_body children & trailing sectPr, determine section break BEFORE remapping
            src_children = list(src_body)
            if src_children and src_children[-1].tag == f"{W}sectPr":
                src_final_sect_pr = src_children.pop()
            else:
                src_final_sect_pr = etree.Element(f"{W}sectPr")

            # Determine first section properties of src_doc (internal break if any, else trailing)
            src_first_sect_pr = src_final_sect_pr
            for ch in src_children:
                if ch.tag == f"{W}p":
                    sp = ch.find(f"{W}pPr/{W}sectPr")
                    if sp is not None:
                        src_first_sect_pr = sp
                        break

            src_first_sig, _, _ = _extract_sect_signature(
                src_first_sect_pr, src_doc_rels_by_id, zf, "word/document.xml"
            )

            need_section_break = False
            if section_break_mode == "auto":
                need_section_break = (src_first_sig != prev_final_sig)
            else:
                need_section_break = True

            # Remap all elements in src_body in a SINGLE linear pass!
            _remap_single_pass(
                src_body,
                doc_idx=doc_idx,
                alloc=alloc,
                style_map=style_map,
                num_map=num_map,
                rel_map=doc_rel_map,
                comment_id_map=comment_id_map,
                footnote_id_map=footnote_id_map,
                endnote_id_map=endnote_id_map,
                bm_id_map=bm_id_map,
                bm_name_map=bm_name_map,
                vml_id_map=vml_id_map,
                vml_spid_map=vml_spid_map,
                ole_obj_id_map=ole_obj_id_map,
                para_id_map=para_id_map,
                normal_style_override=normal_style_override,
                strip_vanish_outline=should_strip_vanish_outline,
            )

            # Re-read children after single pass (in case _GoBack bookmarks were removed)
            src_children = list(src_body)
            if src_children and src_children[-1].tag == f"{W}sectPr":
                src_final_sect_pr = src_children.pop()
            else:
                src_final_sect_pr = etree.Element(f"{W}sectPr")

            if need_section_break:
                # Close previous document's section via a section-break paragraph
                prev_sp_copy = copy.deepcopy(prev_final_sect_pr)
                base_body.append(_make_section_break_paragraph(prev_sp_copy))
                section_count += 1

                # Set section break type on the FOLLOWING section (src's first section)
                if section_break_mode in ("continuous", "page"):
                    target_first_sp = src_final_sect_pr
                    for ch in src_children:
                        if ch.tag == f"{W}p":
                            sp = ch.find(f"{W}pPr/{W}sectPr")
                            if sp is not None:
                                target_first_sp = sp
                                break
                    _set_sect_type(
                        target_first_sp,
                        "continuous" if section_break_mode == "continuous" else "nextPage",
                    )

                # Enforce header/footer non-leakage on any internal sections and final section of src
                for ch in src_children:
                    if ch.tag == f"{W}p":
                        sp = ch.find(f"{W}pPr/{W}sectPr")
                        if sp is not None:
                            prev_hdr_types, prev_ftr_types = _enforce_hf_isolation(
                                prev_hdr_types, prev_ftr_types, sp, _ensure_empty_hf
                            )
                            section_count += 1
                    base_body.append(ch)

                prev_hdr_types, prev_ftr_types = _enforce_hf_isolation(
                    prev_hdr_types, prev_ftr_types, src_final_sect_pr, _ensure_empty_hf
                )
                prev_final_sect_pr = src_final_sect_pr
                prev_final_sig, _, _ = _extract_sect_signature(
                    src_final_sect_pr, _rels_map(doc0_rels), None, "word/document.xml", parts=parts
                )
            else:
                # Seamless end-to-end append: still check if src had internal sectPrs
                for ch in src_children:
                    if ch.tag == f"{W}p":
                        sp = ch.find(f"{W}pPr/{W}sectPr")
                        if sp is not None:
                            prev_hdr_types, prev_ftr_types = _enforce_hf_isolation(
                                prev_hdr_types, prev_ftr_types, sp, _ensure_empty_hf
                            )
                            section_count += 1
                    base_body.append(ch)
                # Keep prev_final_sect_pr / prev_final_sig updated to src_final_sect_pr
                prev_hdr_types, prev_ftr_types = _enforce_hf_isolation(
                    prev_hdr_types, prev_ftr_types, src_final_sect_pr, _ensure_empty_hf
                )
                prev_final_sect_pr = src_final_sect_pr
                prev_final_sig, _, _ = _extract_sect_signature(
                    src_final_sect_pr, _rels_map(doc0_rels), None, "word/document.xml", parts=parts
                )

        if file_status_callback:
            file_status_callback(doc_idx, "done")
        if progress_callback:
            progress_callback(
                f"已合并 ({doc_idx + 1}/{total_docs}): {os.path.basename(src_path)}",
                doc_idx + 1,
                total_docs,
            )

    # Append final trailing w:sectPr to merged body
    base_body.append(prev_final_sect_pr)

    # Ensure union namespaces & mc:Ignorable on root <w:document>
    if set(union_nsmap.keys()) != set(base_doc_root.nsmap.keys()) or any(
        union_nsmap[k] != base_doc_root.nsmap.get(k) for k in union_nsmap
    ):
        new_doc_root = etree.Element(base_doc_root.tag, attrib=base_doc_root.attrib, nsmap=union_nsmap)
        for ch in list(base_doc_root):
            new_doc_root.append(ch)
        base_doc_root = new_doc_root
    if ignorable_tokens:
        base_doc_root.set(f"{MC}Ignorable", " ".join(ignorable_tokens))

    # Ensure settings.xml does NOT have <w:updateFields w:val="true"/> so Word never prompts on open
    if "word/settings.xml" in parts:
        st_root = _parse_xml(parts["word/settings.xml"])
        changed_st = False
        for uf in list(st_root.findall(f"{W}updateFields")):
            st_root.remove(uf)
            changed_st = True
        if changed_st:
            parts["word/settings.xml"] = _serialize_xml(st_root)

    # Serialize modified XML parts
    if progress_callback:
        progress_callback("写入合并后的 Word 文档...", total_docs, total_docs)

    parts["word/document.xml"] = _serialize_xml(base_doc_root)
    parts["[Content_Types].xml"] = _serialize_xml(ct_root)
    if base_styles_root is not None:
        parts["word/styles.xml"] = _serialize_xml(base_styles_root)
    if base_num_root is not None:
        parts["word/numbering.xml"] = _serialize_xml(base_num_root)
    if base_fonts_root is not None:
        parts["word/fontTable.xml"] = _serialize_xml(base_fonts_root)
    if base_comments_root is not None:
        parts["word/comments.xml"] = _serialize_xml(base_comments_root)
    if base_comments_ext_root is not None:
        parts["word/commentsExtended.xml"] = _serialize_xml(base_comments_ext_root)
    if base_comments_ids_root is not None:
        parts["word/commentsIds.xml"] = _serialize_xml(base_comments_ids_root)
    if base_people_root is not None:
        parts["word/people.xml"] = _serialize_xml(base_people_root)
    if base_footnotes_root is not None:
        parts["word/footnotes.xml"] = _serialize_xml(base_footnotes_root)
    if base_endnotes_root is not None:
        parts["word/endnotes.xml"] = _serialize_xml(base_endnotes_root)
    for rel_path, rel_root in parsed_rels.items():
        parts[rel_path] = _serialize_xml(rel_root)

    # Atomic write to output_path
    out_abs = os.path.abspath(output_path)
    out_dir = os.path.dirname(out_abs)
    os.makedirs(out_dir, exist_ok=True)
    fd, tmp_out = tempfile.mkstemp(prefix=".wc_merge_", suffix=".docx", dir=out_dir)
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for pname, pdata in parts.items():
                zout.writestr(pname, pdata)
        os.replace(tmp_out, out_abs)
    finally:
        if os.path.exists(tmp_out):
            try:
                os.remove(tmp_out)
            except OSError:
                pass

    toc_info = detect_tocs(out_abs)
    return MergeResult(
        output_path=out_abs,
        ordered_inputs=ordered_inputs,
        has_toc=toc_info.has_toc,
        toc_info=toc_info,
        section_count=section_count,
        comments_merged=comments_merged,
        footnotes_merged=footnotes_merged,
        endnotes_merged=endnotes_merged,
    )
