"""Comprehensive unit and integration tests for WordCombine."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import time
import types
import unicodedata
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from docx import Document
from docx.shared import Twips
from lxml import etree

from word_combine.ooxml_merger import (
    NS_CT,
    NS_O,
    NS_R,
    NS_REL,
    NS_V,
    NS_W,
    NS_W14,
    NS_W15,
    NS_W16CID,
    O,
    R,
    TOCInfo,
    V,
    W,
    W14,
    W15,
    W16CID,
    detect_tocs,
    merge_docx_files,
)
from word_combine.sorting import collect_docx_files, natural_sort_key, sort_docx_paths
from word_combine.toc_refresher import _refresh_toc_windows, refresh_toc_silent
from word_combine.verifier import verify_merged_docx

WORKSPACE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MERGED_REAL_PATH = os.path.join(WORKSPACE_DIR, "output", "合并文档.docx")


def _make_basic_docx(path: str, paragraphs: list[str], header_text: str | None = None) -> None:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Twips(11906)
    sec.page_height = Twips(16838)
    sec.top_margin = Twips(1440)
    sec.bottom_margin = Twips(1440)
    sec.left_margin = Twips(1800)
    sec.right_margin = Twips(1800)
    if header_text is not None:
        sec.header.paragraphs[0].text = header_text
    for txt in paragraphs:
        doc.add_paragraph(txt)
    doc.save(path)


def _rewrite_zip(path: str, updates: dict[str, bytes]) -> None:
    with zipfile.ZipFile(path, "r") as zin:
        parts = {n: zin.read(n) for n in zin.namelist()}
    parts.update(updates)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for k, v in parts.items():
            zout.writestr(k, v)


class SortingTests(unittest.TestCase):
    def test_natural_sort_order_and_nfc_nfd(self) -> None:
        nfd = unicodedata.normalize("NFD", "3.3 变量中心.docx")
        nfc = unicodedata.normalize("NFC", "3.3 变量中心.docx")
        self.assertEqual(natural_sort_key(nfd), natural_sort_key(nfc))

        names = [
            "/tmp/10. 附录.docx",
            "/tmp/3.10 系统管理.docx",
            "/tmp/3.9 视图中心.docx",
            "/tmp/4. 其他要求.docx",
            "/tmp/3.1 3.2 工作台、数据中心.docx",
            "/tmp/3.3 变量中心.docx",
        ]
        sorted_names = [os.path.basename(p) for p in sort_docx_paths(names)]
        self.assertEqual(
            sorted_names,
            [
                "3.1 3.2 工作台、数据中心.docx",
                "3.3 变量中心.docx",
                "3.9 视图中心.docx",
                "3.10 系统管理.docx",
                "4. 其他要求.docx",
                "10. 附录.docx",
            ],
        )

    def test_collect_docx_ignores_lock_hidden_and_output(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            p1 = os.path.join(td, "1. 第一章.docx")
            p2 = os.path.join(td, "2. 第二章.docx")
            lock = os.path.join(td, "~$1. 第一章.docx")
            hidden = os.path.join(td, ".hidden.docx")
            txt = os.path.join(td, "readme.txt")
            for p in (p1, p2, lock, hidden, txt):
                Path(p).write_bytes(b"x")
            got = collect_docx_files(td, exclude_paths=[p2])
            self.assertEqual(got, [p1])


class MergerEdgeCaseTests(unittest.TestCase):
    def test_header_non_leakage_in_auto_mode(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            d0 = os.path.join(td, "1. 有页眉.docx")
            d1 = os.path.join(td, "2. 无页眉一.docx")
            d2 = os.path.join(td, "3. 无页眉二.docx")
            out = os.path.join(td, "merged.docx")
            _make_basic_docx(d0, ["文档一正文"], header_text="机密页眉")
            _make_basic_docx(d1, ["文档二正文"], header_text=None)
            _make_basic_docx(d2, ["文档三正文"], header_text=None)

            res = merge_docx_files([d0, d1, d2], out, section_break_mode="auto")
            # d0 is section 1 (with header); d1 and d2 share empty header so they merge into section 2!
            self.assertEqual(res.section_count, 2)
            rep = verify_merged_docx([d0, d1, d2], out)
            self.assertTrue(rep.ok, rep.errors)

            with zipfile.ZipFile(out) as z:
                doc_root = etree.fromstring(z.read("word/document.xml"))
                rels = {
                    el.get("Id"): el.get("Target")
                    for el in etree.fromstring(z.read("word/_rels/document.xml.rels"))
                }
                sect_prs = doc_root.xpath(".//w:sectPr", namespaces={"w": NS_W})
                self.assertEqual(len(sect_prs), 2)
                # Section 1 has non-empty header; Section 2 must have an explicit empty header
                s1_hdr = sect_prs[0].find(f"{W}headerReference")
                s2_hdr = sect_prs[1].find(f"{W}headerReference")
                self.assertIsNotNone(s1_hdr)
                self.assertIsNotNone(s2_hdr)
                s1_txt = "".join(etree.fromstring(z.read("word/" + rels[s1_hdr.get(f"{R}id")])).itertext()).strip()
                s2_txt = "".join(etree.fromstring(z.read("word/" + rels[s2_hdr.get(f"{R}id")])).itertext()).strip()
                self.assertEqual(s1_txt, "机密页眉")
                self.assertEqual(s2_txt, "")

    def test_page_and_continuous_section_break_placement(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            d0 = os.path.join(td, "1. a.docx")
            d1 = os.path.join(td, "2. b.docx")
            _make_basic_docx(d0, ["第一部分"])
            _make_basic_docx(d1, ["第二部分"])

            for mode, expected_val in (("page", "nextPage"), ("continuous", "continuous")):
                out = os.path.join(td, f"out_{mode}.docx")
                res = merge_docx_files([d0, d1], out, section_break_mode=mode)
                self.assertEqual(res.section_count, 2)
                with zipfile.ZipFile(out) as z:
                    doc_root = etree.fromstring(z.read("word/document.xml"))
                    sect_prs = doc_root.xpath(".//w:sectPr", namespaces={"w": NS_W})
                    self.assertEqual(len(sect_prs), 2)
                    # <w:type> must be on the FOLLOWING section (sect_prs[1]) BEFORE <w:pgSz>
                    self.assertIsNone(sect_prs[0].find(f"{W}type"))
                    t_el = sect_prs[1].find(f"{W}type")
                    self.assertIsNotNone(t_el)
                    self.assertEqual(t_el.get(f"{W}val"), expected_val)
                    tags = [etree.QName(c).localname for c in sect_prs[1]]
                    self.assertLess(tags.index("type"), tags.index("pgSz"))

    def test_style_cross_type_collision_and_heading_outline_preserved(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            d0 = os.path.join(td, "1. base.docx")
            d1 = os.path.join(td, "2. collide.docx")
            out = os.path.join(td, "merged.docx")
            _make_basic_docx(d0, ["Base body"])
            _make_basic_docx(d1, ["Heading text", "Table cell"])

            # Inject styleId="af0" as paragraph style in d0, and as table style in d1;
            # Also inject styleId="20" (heading 2) with different spacing and no inline outlineLvl on paragraph
            def _remove_by_name(styles_root: etree._Element, names: set[str]) -> None:
                for st in list(styles_root.findall(f"{W}style")):
                    nm = st.find(f"{W}name")
                    if nm is not None and (nm.get(f"{W}val") or "").lower() in names:
                        styles_root.remove(st)

            with zipfile.ZipFile(d0) as z0:
                st0 = etree.fromstring(z0.read("word/styles.xml"))
                doc0 = etree.fromstring(z0.read("word/document.xml"))
            _remove_by_name(st0, {"footer", "heading 2"})
            etree.SubElement(st0, f"{W}style", {f"{W}type": "paragraph", f"{W}styleId": "af0"}).append(
                etree.Element(f"{W}name", {f"{W}val": "footer"})
            )
            h0 = etree.SubElement(st0, f"{W}style", {f"{W}type": "paragraph", f"{W}styleId": "20"})
            etree.SubElement(h0, f"{W}name", {f"{W}val": "heading 2"})
            ppr0 = etree.SubElement(h0, f"{W}pPr")
            etree.SubElement(ppr0, f"{W}outlineLvl", {f"{W}val": "1"})
            etree.SubElement(ppr0, f"{W}spacing", {f"{W}before": "100"})
            p0 = doc0.find(f"{W}body/{W}p")
            p0.insert(0, etree.fromstring(f'<w:pPr xmlns:w="{NS_W}"><w:pStyle w:val="20"/></w:pPr>'))
            _rewrite_zip(d0, {
                "word/styles.xml": etree.tostring(st0, xml_declaration=True, encoding="UTF-8", standalone=True),
                "word/document.xml": etree.tostring(doc0, xml_declaration=True, encoding="UTF-8", standalone=True),
            })

            with zipfile.ZipFile(d1) as z1:
                st1 = etree.fromstring(z1.read("word/styles.xml"))
                doc1 = etree.fromstring(z1.read("word/document.xml"))
            _remove_by_name(st1, {"footer", "heading 2"})
            t1 = etree.SubElement(st1, f"{W}style", {f"{W}type": "table", f"{W}styleId": "af0"})
            etree.SubElement(t1, f"{W}name", {f"{W}val": "CustomTableGrid"})
            h1 = etree.SubElement(st1, f"{W}style", {f"{W}type": "paragraph", f"{W}styleId": "20"})
            etree.SubElement(h1, f"{W}name", {f"{W}val": "heading 2"})
            # Intentionally omit outlineLvl on h1 to verify merger adds outlineLvl=1 when renaming!
            ppr1 = etree.SubElement(h1, f"{W}pPr")
            etree.SubElement(ppr1, f"{W}spacing", {f"{W}before": "240"})
            ps1 = doc1.findall(f"{W}body/{W}p")
            ps1[0].insert(0, etree.fromstring(f'<w:pPr xmlns:w="{NS_W}"><w:pStyle w:val="20"/></w:pPr>'))
            tbl_xml = etree.fromstring(
                f'<w:tbl xmlns:w="{NS_W}"><w:tblPr><w:tblStyle w:val="af0"/></w:tblPr>'
                f'<w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
            )
            doc1.find(f"{W}body").insert(1, tbl_xml)
            _rewrite_zip(d1, {
                "word/styles.xml": etree.tostring(st1, xml_declaration=True, encoding="UTF-8", standalone=True),
                "word/document.xml": etree.tostring(doc1, xml_declaration=True, encoding="UTF-8", standalone=True),
            })

            merge_docx_files([d0, d1], out)
            rep = verify_merged_docx([d0, d1], out)
            self.assertTrue(rep.ok, rep.errors)

            with zipfile.ZipFile(out) as zout:
                st_out = etree.fromstring(zout.read("word/styles.xml"))
                styles = {s.get(f"{W}styleId"): s for s in st_out.findall(f"{W}style")}
                self.assertEqual(styles["af0"].get(f"{W}type"), "paragraph")
                self.assertIn("d1_af0", styles)
                self.assertEqual(styles["d1_af0"].get(f"{W}type"), "table")
                self.assertIn("d1_20", styles)
                ol = styles["d1_20"].find(f"{W}pPr/{W}outlineLvl")
                self.assertIsNotNone(ol)
                self.assertEqual(ol.get(f"{W}val"), "1")

    def test_bookmark_vml_ole_and_comments_footnotes_endnotes(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            d0 = os.path.join(td, "1. first.docx")
            d1 = os.path.join(td, "2. second.docx")
            out = os.path.join(td, "merged.docx")
            _make_basic_docx(d0, ["Doc0 para"])
            _make_basic_docx(d1, ["Doc1 para"])

            # Inject matching bookmark _Ref1, VML shape + OLEObject ShapeID="_x0000_i1025" in BOTH d0 and d1,
            # plus comments, footnote, and endnote in d1.
            for idx, pth in enumerate((d0, d1)):
                with zipfile.ZipFile(pth) as z:
                    doc = etree.fromstring(z.read("word/document.xml"))
                    ct = etree.fromstring(z.read("[Content_Types].xml"))
                    rels = etree.fromstring(z.read("word/_rels/document.xml.rels"))
                body = doc.find(f"{W}body")
                extra_p = etree.fromstring(
                    f'<w:p xmlns:w="{NS_W}" xmlns:r="{NS_R}" xmlns:v="{NS_V}" xmlns:o="{NS_O}" xmlns:w14="{NS_W14}" w14:paraId="11223344">'
                    f'<w:hyperlink w:anchor="_Ref1"><w:r><w:t>Go</w:t></w:r></w:hyperlink>'
                    f'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                    f'<w:r><w:instrText xml:space="preserve"> PAGEREF _Ref1 \\h </w:instrText></w:r>'
                    f'<w:r><w:fldChar w:fldCharType="end"/></w:r>'
                    f'<w:bookmarkStart w:id="5" w:name="_Ref1"/>'
                    f'<w:r><w:t>Target{idx}</w:t></w:r>'
                    f'<w:bookmarkEnd w:id="5"/>'
                    f'<w:r><w:object><v:shape id="_x0000_i1025" style="width:10pt;height:10pt"/>'
                    f'<o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" ShapeID="_x0000_i1025" ObjectID="_1857001" r:id="rId99"/></w:object></w:r>'
                    f'</w:p>'
                )
                body.insert(0, extra_p)
                rel_el = etree.SubElement(rels, f"{{{NS_REL}}}Relationship")
                rel_el.set("Id", "rId99")
                rel_el.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package")
                rel_el.set("Target", "embeddings/book.xlsx")
                if not any(el.get("Extension") == "xlsx" for el in ct if el.tag == f"{{{NS_CT}}}Default"):
                    d = etree.SubElement(ct, f"{{{NS_CT}}}Default")
                    d.set("Extension", "xlsx")
                    d.set("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                updates = {
                    "word/document.xml": etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True),
                    "word/_rels/document.xml.rels": etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True),
                    "[Content_Types].xml": etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True),
                    "word/embeddings/book.xlsx": f"dummy-xlsx-{idx}".encode(),
                }
                if idx == 1:
                    # Add comment + footnote + endnote in d1
                    p_cm = body.findall(f"{W}p")[1]
                    p_cm.append(etree.fromstring(f'<w:commentRangeStart xmlns:w="{NS_W}" w:id="3"/>'))
                    p_cm.append(etree.fromstring(f'<w:commentRangeEnd xmlns:w="{NS_W}" w:id="3"/>'))
                    p_cm.append(etree.fromstring(f'<w:r xmlns:w="{NS_W}"><w:commentReference w:id="3"/></w:r>'))
                    p_cm.append(etree.fromstring(f'<w:r xmlns:w="{NS_W}"><w:footnoteReference w:id="2"/></w:r>'))
                    p_cm.append(etree.fromstring(f'<w:r xmlns:w="{NS_W}"><w:endnoteReference w:id="4"/></w:r>'))
                    updates["word/document.xml"] = etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)
                    updates["word/comments.xml"] = (
                        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        f'<w:comments xmlns:w="{NS_W}" xmlns:w14="{NS_W14}">'
                        f'<w:comment w:id="3" w:author="Tester"><w:p w14:paraId="AABBCCDD"><w:r><w:t>批注A</w:t></w:r></w:p></w:comment>'
                        f'</w:comments>'
                    ).encode()
                    updates["word/commentsExtended.xml"] = (
                        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        f'<w15:commentsEx xmlns:w15="{NS_W15}"><w15:commentEx w15:paraId="AABBCCDD" w15:done="0"/></w15:commentsEx>'
                    ).encode()
                    updates["word/commentsIds.xml"] = (
                        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        f'<w16cid:commentsIds xmlns:w16cid="{NS_W16CID}"><w16cid:commentId w16cid:paraId="AABBCCDD" w16cid:durableId="11112222"/></w16cid:commentsIds>'
                    ).encode()
                    updates["word/people.xml"] = (
                        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        f'<w15:people xmlns:w15="{NS_W15}"><w15:person w15:author="Tester"/></w15:people>'
                    ).encode()
                    updates["word/footnotes.xml"] = (
                        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        f'<w:footnotes xmlns:w="{NS_W}"><w:footnote w:id="2"><w:p><w:r><w:t>脚注B</w:t></w:r></w:p></w:footnote></w:footnotes>'
                    ).encode()
                    updates["word/endnotes.xml"] = (
                        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        f'<w:endnotes xmlns:w="{NS_W}"><w:endnote w:id="4"><w:p><w:r><w:t>尾注C</w:t></w:r></w:p></w:endnote></w:endnotes>'
                    ).encode()
                _rewrite_zip(pth, updates)

            res = merge_docx_files([d0, d1], out)
            self.assertEqual(res.comments_merged, 1)
            self.assertEqual(res.footnotes_merged, 1)
            self.assertEqual(res.endnotes_merged, 1)
            rep = verify_merged_docx([d0, d1], out)
            self.assertTrue(rep.ok, rep.errors)

            with zipfile.ZipFile(out) as zout:
                doc_out = etree.fromstring(zout.read("word/document.xml"))
                # Verify d1 hyperlink & PAGEREF updated to match d1's renamed bookmark
                bms = [b.get(f"{W}name") for b in doc_out.iter(f"{W}bookmarkStart")]
                self.assertEqual(len(set(bms)), 2)
                renamed_bm = [b for b in bms if b != "_Ref1"][0]
                anchors = [h.get(f"{W}anchor") for h in doc_out.iter(f"{W}hyperlink")]
                self.assertEqual(anchors, ["_Ref1", renamed_bm])
                instrs = ["".join(i.itertext()) for i in doc_out.iter(f"{W}instrText")]
                self.assertTrue(any(renamed_bm in it for it in instrs))
                # Verify VML shape & OLEObject ShapeID paired and unique
                shapes = [s.get("id") for s in doc_out.iter(f"{V}shape")]
                oles = [o.get("ShapeID") for o in doc_out.iter(f"{O}OLEObject")]
                self.assertEqual(shapes, oles)
                self.assertEqual(len(set(shapes)), 2)
                # Verify comment paraId matches commentsExtended & commentsIds
                cm_p = etree.fromstring(zout.read("word/comments.xml")).find(f"{W}comment/{W}p")
                new_cm_pid = cm_p.get(f"{W14}paraId")
                cex_p = etree.fromstring(zout.read("word/commentsExtended.xml")).find(f"{W15}commentEx")
                self.assertEqual(cex_p.get(f"{W15}paraId"), new_cm_pid)
                cid_p = etree.fromstring(zout.read("word/commentsIds.xml")).find(f"{W16CID}commentId")
                self.assertEqual(cid_p.get(f"{W16CID}paraId"), new_cm_pid)


class MergerProgressCallbackTests(unittest.TestCase):
    def test_progress_and_file_status_callbacks_fire_in_order(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            d0 = os.path.join(td, "1. a.docx")
            d1 = os.path.join(td, "2. b.docx")
            d2 = os.path.join(td, "3. c.docx")
            out = os.path.join(td, "merged.docx")
            _make_basic_docx(d0, ["Part A"])
            _make_basic_docx(d1, ["Part B"])
            _make_basic_docx(d2, ["Part C"])

            prog_events: list[tuple[str, int, int]] = []
            file_events: list[tuple[int, str]] = []

            merge_docx_files(
                [d0, d1, d2],
                out,
                progress_callback=lambda m, c, t: prog_events.append((m, c, t)),
                file_status_callback=lambda idx, st: file_events.append((idx, st)),
            )

            self.assertEqual(
                file_events,
                [(0, "processing"), (0, "done"), (1, "processing"), (1, "done"), (2, "processing"), (2, "done")],
            )
            self.assertGreaterEqual(len(prog_events), 4)
            self.assertEqual(prog_events[-1][1], 3)
            self.assertEqual(prog_events[-1][2], 3)


class WindowsTocRefresherTests(unittest.TestCase):
    def _make_docx_with_toc(self, path: str) -> None:
        _make_basic_docx(path, ["TOC_MARKER", "Body"])
        with zipfile.ZipFile(path) as z:
            doc_xml = etree.fromstring(z.read("word/document.xml"))
        body = doc_xml.find(f"{W}body")
        for p in list(body):
            if "TOC_MARKER" in "".join(p.itertext()):
                sdt = etree.fromstring(
                    f'<w:sdt xmlns:w="{NS_W}"><w:sdtPr><w:docPartObj>'
                    f'<w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/>'
                    f'</w:docPartObj></w:sdtPr><w:sdtContent><w:p>'
                    f'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                    f'<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
                    f'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                    f'<w:r><w:t>旧目录\t1</w:t></w:r>'
                    f'<w:r><w:fldChar w:fldCharType="end"/></w:r>'
                    f'</w:p></w:sdtContent></w:sdt>'
                )
                body.replace(p, sdt)
        _rewrite_zip(
            path,
            {"word/document.xml": etree.tostring(doc_xml, xml_declaration=True, encoding="UTF-8", standalone=True)},
        )

    def test_windows_com_word_update_flow(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            docx_path = os.path.join(td, "sample.docx")
            self._make_docx_with_toc(docx_path)
            calls: list[str] = []

            class FakeToc:
                def Update(self) -> None:
                    calls.append("toc.Update")

            class FakeDoc:
                TablesOfContents = type("Tocs", (), {"Count": 1, "__call__": lambda self, i: FakeToc()})()

                def Save(self) -> None:
                    calls.append("doc.Save")

                def Close(self, save_changes: bool = False) -> None:
                    calls.append(f"doc.Close({save_changes})")

            class FakeDocs:
                def Open(self, path: str, **kwargs: object) -> FakeDoc:
                    calls.append("docs.Open")
                    self.opened = path
                    return FakeDoc()

            class FakeApp:
                Visible = True
                DisplayAlerts = 1
                ScreenUpdating = True
                AutomationSecurity = 1
                Documents = FakeDocs()

                def Quit(self, save_changes: bool = False) -> None:
                    calls.append(f"app.Quit({save_changes})")

            fake_app = FakeApp()
            fake_pythoncom = types.SimpleNamespace(
                CoInitialize=lambda: calls.append("CoInitialize"),
                CoUninitialize=lambda: calls.append("CoUninitialize"),
            )

            def fake_dispatch_ex(prog_id: str) -> FakeApp:
                calls.append(f"DispatchEx({prog_id})")
                return fake_app

            fake_win32com_client = types.SimpleNamespace(DispatchEx=fake_dispatch_ex)
            fake_win32com = types.SimpleNamespace(client=fake_win32com_client)

            with mock.patch.dict(
                sys.modules,
                {"pythoncom": fake_pythoncom, "win32com": fake_win32com, "win32com.client": fake_win32com_client},
            ):
                res = _refresh_toc_windows(docx_path, TOCInfo(count=1, sdt_count=1, field_count=1))

            self.assertEqual(res.status, "updated", res.message)
            self.assertFalse(fake_app.Visible)
            self.assertEqual(fake_app.DisplayAlerts, 0)
            self.assertEqual(
                calls,
                [
                    "CoInitialize",
                    "DispatchEx(Word.Application)",
                    "docs.Open",
                    "toc.Update",
                    "doc.Save",
                    "doc.Close(False)",
                    "app.Quit(False)",
                    "CoUninitialize",
                ],
            )

    def test_windows_com_falls_back_to_wps_and_skips_when_absent(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            docx_path = os.path.join(td, "sample.docx")
            self._make_docx_with_toc(docx_path)

            tried: list[str] = []
            fake_pythoncom = types.SimpleNamespace(
                CoInitialize=lambda: None,
                CoUninitialize=lambda: None,
            )

            class FakeToc:
                def Update(self) -> None:
                    pass

            class FakeDoc:
                TablesOfContents = type("Tocs", (), {"Count": 1, "__call__": lambda self, i: FakeToc()})()

                def Save(self) -> None:
                    pass

                def Close(self, save_changes: bool = False) -> None:
                    pass

            class FakeWpsApp:
                Documents = type("Docs", (), {"Open": lambda self, *a, **kw: FakeDoc()})()

                def Quit(self, save_changes: bool = False) -> None:
                    pass

            def dispatch_wps(prog_id: str) -> FakeWpsApp:
                tried.append(prog_id)
                if prog_id == "Word.Application":
                    raise RuntimeError("No Word")
                return FakeWpsApp()

            fake_win32com_client = types.SimpleNamespace(DispatchEx=dispatch_wps)
            fake_win32com = types.SimpleNamespace(client=fake_win32com_client)

            with mock.patch.dict(
                sys.modules,
                {"pythoncom": fake_pythoncom, "win32com": fake_win32com, "win32com.client": fake_win32com_client},
            ):
                res = _refresh_toc_windows(docx_path, TOCInfo(count=1, sdt_count=1, field_count=1))
            self.assertEqual(res.status, "updated")
            self.assertIn("WPS Office", res.message)
            self.assertEqual(tried, ["Word.Application", "KWPS.Application"])

            def dispatch_none(prog_id: str) -> None:
                raise RuntimeError("No Office")

            fake_win32com_client_none = types.SimpleNamespace(DispatchEx=dispatch_none)
            fake_win32com_none = types.SimpleNamespace(client=fake_win32com_client_none)
            with mock.patch.dict(
                sys.modules,
                {"pythoncom": fake_pythoncom, "win32com": fake_win32com_none, "win32com.client": fake_win32com_client_none},
            ):
                res_skip = _refresh_toc_windows(docx_path, TOCInfo(count=1, sdt_count=1, field_count=1))
            self.assertEqual(res_skip.status, "skipped")


class DetectTocsTests(unittest.TestCase):
    def test_detect_tocs_fast_prefilter_and_all_toc_variants(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            plain = os.path.join(td, "plain.docx")
            _make_basic_docx(plain, ["Plain paragraph without TOC"])
            plain_info = detect_tocs(plain)
            self.assertFalse(plain_info.has_toc)
            self.assertEqual((plain_info.count, plain_info.sdt_count, plain_info.field_count), (0, 0, 0))

            sdt_doc = os.path.join(td, "sdt_toc.docx")
            _make_basic_docx(sdt_doc, ["TOC_MARKER"])
            with zipfile.ZipFile(sdt_doc) as z:
                doc_xml = etree.fromstring(z.read("word/document.xml"))
            body = doc_xml.find(f"{W}body")
            body[0] = etree.fromstring(
                f'<w:sdt xmlns:w="{NS_W}"><w:sdtPr><w:docPartObj>'
                f'<w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/>'
                f'</w:docPartObj></w:sdtPr><w:sdtContent><w:p>'
                f'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                f'<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r>'
                f'<w:r><w:fldChar w:fldCharType="end"/></w:r>'
                f'</w:p></w:sdtContent></w:sdt>'
            )
            _rewrite_zip(
                sdt_doc,
                {"word/document.xml": etree.tostring(doc_xml, xml_declaration=True, encoding="UTF-8", standalone=True)},
            )
            sdt_info = detect_tocs(sdt_doc)
            self.assertTrue(sdt_info.has_toc)
            self.assertEqual((sdt_info.count, sdt_info.sdt_count, sdt_info.field_count), (1, 1, 1))

            instr_doc = os.path.join(td, "instr_toc.docx")
            _make_basic_docx(instr_doc, ["MARKER"])
            with zipfile.ZipFile(instr_doc) as z:
                doc_xml = etree.fromstring(z.read("word/document.xml"))
            doc_xml.find(f"{W}body")[0] = etree.fromstring(
                f'<w:p xmlns:w="{NS_W}">'
                f'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                f'<w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h </w:instrText></w:r>'
                f'<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
            )
            _rewrite_zip(
                instr_doc,
                {"word/document.xml": etree.tostring(doc_xml, xml_declaration=True, encoding="UTF-8", standalone=True)},
            )
            instr_info = detect_tocs(instr_doc)
            self.assertTrue(instr_info.has_toc)
            self.assertEqual((instr_info.count, instr_info.sdt_count, instr_info.field_count), (1, 0, 1))

            simple_doc = os.path.join(td, "fldsimple_toc.docx")
            _make_basic_docx(simple_doc, ["MARKER"])
            with zipfile.ZipFile(simple_doc) as z:
                doc_xml = etree.fromstring(z.read("word/document.xml"))
            doc_xml.find(f"{W}body")[0] = etree.fromstring(
                f'<w:p xmlns:w="{NS_W}"><w:fldSimple w:instr=" TOC \\o &quot;1-3&quot; "/></w:p>'
            )
            _rewrite_zip(
                simple_doc,
                {"word/document.xml": etree.tostring(doc_xml, xml_declaration=True, encoding="UTF-8", standalone=True)},
            )
            simple_info = detect_tocs(simple_doc)
            self.assertTrue(simple_info.has_toc)
            self.assertEqual((simple_info.count, simple_info.sdt_count, simple_info.field_count), (1, 0, 1))


class PackageAndSplashTests(unittest.TestCase):
    def test_lazy_package_exports_and_no_eager_lxml_on_gui_import(self) -> None:
        import word_combine

        self.assertEqual(word_combine.__version__, "1.0.1")
        for attr in ("merge_docx_files", "detect_tocs", "verify_merged_docx", "refresh_toc_silent"):
            self.assertTrue(callable(getattr(word_combine, attr)))
        with self.assertRaises(AttributeError):
            _ = word_combine.non_existent_symbol  # type: ignore[attr-defined]

        proc = subprocess.run(
            [
                sys.executable,
                "-c",
                "import sys, word_combine, word_combine.gui; assert 'lxml' not in sys.modules, 'lxml imported eagerly'",
            ],
            cwd=WORKSPACE_DIR,
            capture_output=True,
            text=True,
        )
        self.assertEqual(proc.returncode, 0, proc.stderr or proc.stdout)

    def test_splash_helpers_safe_noop_without_bootloader(self) -> None:
        from word_combine.splash import close_splash, update_splash

        update_splash("Loading interface...")
        close_splash()


class GuiStateTests(unittest.TestCase):
    def test_gui_add_sort_cache_output_and_row_status(self) -> None:
        from word_combine.gui import create_app_window

        try:
            root, app = create_app_window(show=False)
        except Exception as exc:
            self.skipTest(f"Tk display unavailable: {exc}")
            return

        with tempfile.TemporaryDirectory() as td:
            try:
                f1 = os.path.join(td, "3.10 第二.docx")
                f2 = os.path.join(td, "3.9 第一.docx")
                _make_basic_docx(f1, ["B"])
                _make_basic_docx(f2, ["A"])

                app._merge_added_paths([f1, f2])
                self.assertEqual([os.path.basename(p) for p in app.file_paths], ["3.9 第一.docx", "3.10 第二.docx"])
                # Auto-updates default output path to first file's directory
                self.assertEqual(app.output_var.get(), os.path.join(td, "合并文档.docx"))

                # Immediate non-blocking insert shows "检测中..." before background scan completes
                self.assertEqual(app.tree.item("0", "values")[3], "检测中...")
                self.assertEqual(app.tree.item("1", "values")[3], "检测中...")

                # Pump UI event loop until async TOC scan populates cache and updates rows
                deadline = time.monotonic() + 5.0
                rp1 = os.path.realpath(f1)
                rp2 = os.path.realpath(f2)
                while time.monotonic() < deadline:
                    root.update()
                    if (
                        rp1 in app._toc_cache
                        and rp2 in app._toc_cache
                        and app.tree.item("0", "values")[3] == "无"
                        and app.tree.item("1", "values")[3] == "无"
                    ):
                        break
                    time.sleep(0.02)

                self.assertIn(rp1, app._toc_cache)
                self.assertIn(rp2, app._toc_cache)
                self.assertEqual(app.tree.item("0", "values")[3], "无")
                self.assertEqual(app.tree.item("1", "values")[3], "无")

                # Row status column updates live
                app._update_row_status(0, "合并中...")
                self.assertEqual(app.tree.item("0", "values")[4], "合并中...")
                app._update_row_status(0, "已完成 ✓")
                self.assertEqual(app.tree.item("0", "values")[4], "已完成 ✓")
                root.update_idletasks()
            finally:
                app.shutdown()
                root.destroy()


class RealWorkspaceIntegrationTests(unittest.TestCase):
    @unittest.skipUnless(os.path.isfile(MERGED_REAL_PATH), "Local real workspace merged file not present")
    def test_real_merged_output_matches_all_nine_originals(self) -> None:
        real_inputs = collect_docx_files(WORKSPACE_DIR, exclude_paths=[MERGED_REAL_PATH])
        self.assertEqual(len(real_inputs), 9)
        self.assertEqual(
            [os.path.basename(p) for p in real_inputs],
            [
                "3.1 3.2 工作台、数据中心.docx",
                "3.3 变量中心.docx",
                "3.4 预警中心.docx",
                "3.5 关系中心.docx",
                "3.6 3.7 策略中心 服务中心.docx",
                "3.8 运营中心.docx",
                "3.9 视图中心.docx",
                "3.10 系统管理.docx",
                "4. 其他要求.docx",
            ],
        )
        rep = verify_merged_docx(real_inputs, MERGED_REAL_PATH, check_media_hashes=True)
        self.assertTrue(rep.ok, rep.errors)
        self.assertEqual(rep.stats["非空正文段落数"], 8900)
        self.assertEqual(rep.stats["表格数"], 4621)
        self.assertEqual(rep.stats["图片引用数"], 644)
        self.assertEqual(rep.stats["嵌入OLE对象数"], 17)
        self.assertEqual(rep.stats["批注数"], 5)
        self.assertEqual(rep.stats["节(Section)数"], 2)

    @unittest.skipUnless(
        sys.platform == "darwin"
        and os.path.exists("/Applications/Microsoft Word.app")
        and all(
            os.path.isfile(os.path.join(WORKSPACE_DIR, name))
            for name in ("3.3 变量中心.docx", "3.5 关系中心.docx", "4. 其他要求.docx")
        ),
        "macOS Microsoft Word.app and local sample docs required",
    )
    def test_silent_toc_refresh_on_real_subset(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            toc_doc = os.path.join(td, "0. 目录.docx")
            out_doc = os.path.join(td, "merged_toc.docx")
            _make_basic_docx(toc_doc, ["目  录", "TOC_MARKER"])
            with zipfile.ZipFile(toc_doc) as z:
                doc_xml = etree.fromstring(z.read("word/document.xml"))
            body = doc_xml.find(f"{W}body")
            for p in list(body):
                if "TOC_MARKER" in "".join(p.itertext()):
                    sdt = etree.fromstring(
                        f'<w:sdt xmlns:w="{NS_W}"><w:sdtPr><w:docPartObj>'
                        f'<w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/>'
                        f'</w:docPartObj></w:sdtPr><w:sdtContent><w:p>'
                        f'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                        f'<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
                        f'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                        f'<w:r><w:t>旧目录\t1</w:t></w:r>'
                        f'<w:r><w:fldChar w:fldCharType="end"/></w:r>'
                        f'</w:p></w:sdtContent></w:sdt>'
                    )
                    body.replace(p, sdt)
            _rewrite_zip(
                toc_doc,
                {"word/document.xml": etree.tostring(doc_xml, xml_declaration=True, encoding="UTF-8", standalone=True)},
            )

            inputs = [
                toc_doc,
                os.path.join(WORKSPACE_DIR, "3.3 变量中心.docx"),
                os.path.join(WORKSPACE_DIR, "3.5 关系中心.docx"),
                os.path.join(WORKSPACE_DIR, "4. 其他要求.docx"),
            ]
            mres = merge_docx_files(inputs, out_doc)
            self.assertTrue(mres.has_toc)

            rres = refresh_toc_silent(out_doc, timeout=180)
            self.assertEqual(rres.status, "updated", rres.message)

            rep = verify_merged_docx(
                inputs,
                out_doc,
                ignore_toc_content=True,
                check_media_hashes=True,
                word_resaved=True,
                expected_toc_substrings=["3.3.\t变量中心", "3.5.\t关系中心", "4.\t其他要求"],
            )
            self.assertTrue(rep.ok, rep.errors)
            self.assertTrue(detect_tocs(out_doc).has_toc)


if __name__ == "__main__":
    unittest.main()
