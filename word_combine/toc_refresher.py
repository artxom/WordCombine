"""Silent Microsoft Word / WPS Table of Contents (目录) and field refresher (macOS & Windows)."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import uuid
import zipfile
from dataclasses import dataclass
from typing import Callable, List, Optional, Sequence, Tuple

from lxml import etree

from .ooxml_merger import NS_W, TOCInfo, W, detect_tocs

LogCallback = Callable[[str], None]

_PRIMARY_STAGE_ROOT = os.path.expanduser(
    "~/Library/Group Containers/UBF8T346G9.Office/WordCombineWork"
)
_FALLBACK_STAGE_ROOT = os.path.expanduser(
    "~/Library/Containers/com.microsoft.Word/Data/tmp/WordCombineWork"
)


@dataclass
class RefreshResult:
    """Result of a silent TOC refresh attempt."""

    status: str  # "updated" | "no_toc" | "skipped" | "error"
    toc_count: int
    field_count: int
    message: str


def _is_word_running() -> bool:
    res = subprocess.run(["pgrep", "-x", "Microsoft Word"], capture_output=True)
    return res.returncode == 0


def _ensure_no_update_fields_prompt(docx_path: str) -> None:
    """Strip `<w:updateFields w:val="true"/>` from word/settings.xml if present."""
    with zipfile.ZipFile(docx_path, "r") as zin:
        if "word/settings.xml" not in zin.namelist():
            return
        raw_settings = zin.read("word/settings.xml")
        st_root = etree.fromstring(raw_settings)
        changed = False
        for uf in list(st_root.findall(f"{W}updateFields")):
            st_root.remove(uf)
            changed = True
        if not changed:
            return
        parts = {name: zin.read(name) for name in zin.namelist()}

    parts["word/settings.xml"] = etree.tostring(
        st_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    out_dir = os.path.dirname(os.path.abspath(docx_path))
    fd, tmp_path = tempfile.mkstemp(prefix=".wc_st_", suffix=".docx", dir=out_dir)
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for k, v in parts.items():
                zout.writestr(k, v)
        os.replace(tmp_path, docx_path)
    finally:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


def _stage_copy_for_word(docx_path: str, root_dir: str) -> Tuple[str, str]:
    stage_dir = os.path.join(root_dir, str(uuid.uuid4()))
    os.makedirs(stage_dir, exist_ok=True)
    staged_path = os.path.join(stage_dir, "doc.docx")
    shutil.copy2(docx_path, staged_path)
    return stage_dir, staged_path


def _atomic_replace_from_staged(staged_path: str, target_path: str) -> None:
    target_abs = os.path.abspath(target_path)
    target_dir = os.path.dirname(target_abs)
    fd, tmp_dst = tempfile.mkstemp(prefix=".wc_toc_", suffix=".docx", dir=target_dir)
    os.close(fd)
    try:
        shutil.copy2(staged_path, tmp_dst)
        os.replace(tmp_dst, target_abs)
    finally:
        if os.path.exists(tmp_dst):
            try:
                os.remove(tmp_dst)
            except OSError:
                pass


_REFRESH_APPLESCRIPT = r'''
on run argv
    set posixPath to item 1 of argv
    set hideAndQuit to (item 2 of argv) is "true"
    tell application "Microsoft Word"
        if hideAndQuit then
            launch
            try
                tell application "System Events"
                    set visible of process "Microsoft Word" to false
                end tell
            end try
        end if
        set prevAlerts to display alerts
        set display alerts to alerts none
        set docObj to missing value
        try
            set docObj to open file name posixPath confirm conversions false read only false add to recent files false
            set tocCount to count of tables of contents of docObj
            repeat with i from 1 to tocCount
                update (table of contents i of docObj)
            end repeat
            set fldCount to 0
            if tocCount is 0 then
                try
                    set fldList to every field of docObj
                    set fldCount to count of fldList
                    repeat with f in fldList
                        try
                            set fcText to content of (field code of f)
                            if fcText contains "TOC" or fcText contains "PAGEREF" or fcText contains "NUMPAGES" or fcText contains "PAGE" then
                                update field f
                            end if
                        end try
                    end repeat
                end try
            end if
            save docObj
            close docObj saving no
            set docObj to missing value
            set display alerts to prevAlerts
            if hideAndQuit then
                quit saving no
            end if
            return "OK tocCount=" & tocCount & " fldCount=" & fldCount
        on error errMsg number errNum
            try
                if docObj is not missing value then close docObj saving no
            end try
            try
                set display alerts to prevAlerts
            end try
            if hideAndQuit then
                try
                    quit saving no
                end try
            end if
            error errMsg number errNum
        end try
    end tell
end run
'''


def _refresh_toc_macos(
    docx_path: str,
    toc_info: TOCInfo,
    *,
    timeout: int = 300,
    log_callback: Optional[LogCallback] = None,
) -> RefreshResult:
    """Silently refresh TOC on macOS using /Applications/Microsoft Word.app."""
    if not os.path.exists("/Applications/Microsoft Word.app"):
        return RefreshResult(
            status="skipped",
            toc_count=0,
            field_count=0,
            message="未检测到 /Applications/Microsoft Word.app，已跳过静默目录刷新。",
        )

    if log_callback:
        log_callback("正在通过 Microsoft Word 后台静默刷新目录...")

    last_err = ""
    for stage_root in (_PRIMARY_STAGE_ROOT, _FALLBACK_STAGE_ROOT):
        stage_dir = None
        try:
            stage_dir, staged_docx = _stage_copy_for_word(docx_path, stage_root)
            was_running = _is_word_running()
            hide_and_quit = "false" if was_running else "true"

            proc = subprocess.run(
                ["osascript", "-", staged_docx, hide_and_quit],
                input=_REFRESH_APPLESCRIPT,
                text=True,
                capture_output=True,
                timeout=timeout,
            )
            if proc.returncode != 0 and "-609" in (proc.stderr or ""):
                # Word may have just exited or been swapped by AutoUpdate; wait briefly and retry once
                time.sleep(2.5)
                was_running = _is_word_running()
                hide_and_quit = "false" if was_running else "true"
                proc = subprocess.run(
                    ["osascript", "-", staged_docx, hide_and_quit],
                    input=_REFRESH_APPLESCRIPT,
                    text=True,
                    capture_output=True,
                    timeout=timeout,
                )
            if proc.returncode == 0:
                _atomic_replace_from_staged(staged_docx, docx_path)
                _ensure_no_update_fields_prompt(docx_path)
                out = proc.stdout.strip()
                toc_cnt = toc_info.count
                fld_cnt = 0
                for tok in out.split():
                    if tok.startswith("tocCount="):
                        try:
                            toc_cnt = int(tok.split("=", 1)[1])
                        except ValueError:
                            pass
                    elif tok.startswith("fldCount="):
                        try:
                            fld_cnt = int(tok.split("=", 1)[1])
                        except ValueError:
                            pass
                msg = f"目录静默刷新完成（TOC={toc_cnt}, Fields={fld_cnt}）。"
                if log_callback:
                    log_callback(msg)
                return RefreshResult(
                    status="updated",
                    toc_count=toc_cnt,
                    field_count=fld_cnt,
                    message=msg,
                )

            last_err = (proc.stderr or proc.stdout or "").strip()
        except subprocess.TimeoutExpired:
            last_err = f"Microsoft Word 刷新超时（>{timeout}s）。"
            break
        except Exception as exc:
            last_err = str(exc)
        finally:
            if stage_dir and os.path.exists(stage_dir):
                shutil.rmtree(stage_dir, ignore_errors=True)

    return RefreshResult(
        status="error",
        toc_count=toc_info.count,
        field_count=0,
        message=f"静默刷新目录失败: {last_err}",
    )


def _refresh_toc_windows(
    docx_path: str,
    toc_info: TOCInfo,
    *,
    timeout: int = 300,
    log_callback: Optional[LogCallback] = None,
) -> RefreshResult:
    """Silently refresh TOC on Windows via Word or WPS Office COM automation."""
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError:
        return RefreshResult(
            status="skipped",
            toc_count=0,
            field_count=0,
            message="未安装 pywin32 COM 组件，已跳过自动目录刷新（文档已正常合并）。",
        )

    if log_callback:
        log_callback("正在通过 Windows Office (Word / WPS) 后台静默刷新目录...")

    stage_dir = tempfile.mkdtemp(prefix="wc_word_")
    staged_docx = os.path.abspath(os.path.join(stage_dir, "doc.docx"))
    shutil.copy2(docx_path, staged_docx)

    outcome: dict = {"status": "error", "toc_count": toc_info.count, "field_count": 0, "message": "未知错误"}

    def _com_worker() -> None:
        pythoncom.CoInitialize()
        app = None
        doc = None
        app_label = ""
        try:
            for prog_id, label in (
                ("Word.Application", "Microsoft Word"),
                ("KWPS.Application", "WPS Office"),
                ("WPS.Application", "WPS Office"),
            ):
                try:
                    app = win32com.client.DispatchEx(prog_id)
                    app_label = label
                    break
                except Exception:
                    app = None

            if app is None:
                outcome.update(
                    {
                        "status": "skipped",
                        "toc_count": 0,
                        "field_count": 0,
                        "message": "未检测到本机 Microsoft Word 或 WPS Office，已跳过自动目录刷新（文档已正常合并，用 Word/WPS 打开后更新目录即可）。",
                    }
                )
                return

            for attr, val in (
                ("Visible", False),
                ("DisplayAlerts", 0),
                ("ScreenUpdating", False),
                ("AutomationSecurity", 3),
            ):
                try:
                    setattr(app, attr, val)
                except Exception:
                    pass

            try:
                doc = app.Documents.Open(
                    staged_docx,
                    ConfirmConversions=False,
                    ReadOnly=False,
                    AddToRecentFiles=False,
                    Visible=False,
                )
            except Exception:
                doc = app.Documents.Open(staged_docx, False, False, False)

            toc_cnt = int(doc.TablesOfContents.Count)
            for idx in range(1, toc_cnt + 1):
                doc.TablesOfContents(idx).Update()

            fld_cnt = 0
            if toc_cnt == 0:
                try:
                    fld_cnt = int(doc.Fields.Count)
                    for idx in range(1, fld_cnt + 1):
                        try:
                            fld = doc.Fields(idx)
                            code_txt = str(fld.Code.Text or "")
                            if any(k in code_txt for k in ("TOC", "PAGEREF", "NUMPAGES", "PAGE")):
                                fld.Update()
                        except Exception:
                            pass
                except Exception:
                    pass

            doc.Save()
            doc.Close(False)
            doc = None
            app.Quit(False)
            app = None

            outcome.update(
                {
                    "status": "updated",
                    "toc_count": toc_cnt,
                    "field_count": fld_cnt,
                    "message": f"目录静默刷新完成（{app_label}, TOC={toc_cnt}, Fields={fld_cnt}）。",
                }
            )
        except Exception as exc:
            outcome.update(
                {
                    "status": "error",
                    "toc_count": toc_info.count,
                    "field_count": 0,
                    "message": f"Windows 静默刷新目录失败 ({app_label or 'COM'}): {exc}",
                }
            )
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass
            if app is not None:
                try:
                    app.Quit(False)
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    try:
        worker = threading.Thread(target=_com_worker, daemon=True)
        worker.start()
        worker.join(timeout=timeout)
        if worker.is_alive():
            return RefreshResult(
                status="error",
                toc_count=toc_info.count,
                field_count=0,
                message=f"Word/WPS 目录刷新超时（>{timeout}s）。",
            )

        if outcome["status"] == "updated":
            _atomic_replace_from_staged(staged_docx, docx_path)
            _ensure_no_update_fields_prompt(docx_path)
            if log_callback:
                log_callback(outcome["message"])

        return RefreshResult(
            status=outcome["status"],
            toc_count=outcome["toc_count"],
            field_count=outcome["field_count"],
            message=outcome["message"],
        )
    finally:
        shutil.rmtree(stage_dir, ignore_errors=True)


def refresh_toc_silent(
    docx_path: str,
    *,
    force: bool = False,
    timeout: int = 300,
    log_callback: Optional[LogCallback] = None,
) -> RefreshResult:
    """Silently refresh Table of Contents and page fields in `docx_path` on macOS or Windows."""
    toc_info = detect_tocs(docx_path)
    if not toc_info.has_toc and not force:
        return RefreshResult(
            status="no_toc",
            toc_count=0,
            field_count=0,
            message="文档中未检测到目录（TOC），无需刷新。",
        )

    if sys.platform == "darwin":
        return _refresh_toc_macos(docx_path, toc_info, timeout=timeout, log_callback=log_callback)
    if sys.platform == "win32":
        return _refresh_toc_windows(docx_path, toc_info, timeout=timeout, log_callback=log_callback)

    return RefreshResult(
        status="skipped",
        toc_count=0,
        field_count=0,
        message=f"当前平台 ({sys.platform}) 不支持后台调用 Word 刷新目录，已跳过。",
    )


_INSPECT_APPLESCRIPT = r'''
on run argv
    set posixPath to item 1 of argv
    set hideAndQuit to (item 2 of argv) is "true"
    set queries to items 3 thru -1 of argv
    tell application "Microsoft Word"
        if hideAndQuit then
            launch
            try
                tell application "System Events"
                    set visible of process "Microsoft Word" to false
                end tell
            end try
        end if
        set prevAlerts to display alerts
        set display alerts to alerts none
        set docObj to missing value
        try
            set docObj to open file name posixPath confirm conversions false read only true add to recent files false
            set outLines to {}
            set secCnt to count of sections of docObj
            set end of outLines to ("SECTIONS=" & secCnt)
            set tocCnt to count of tables of contents of docObj
            set end of outLines to ("TOCS=" & tocCnt)
            repeat with q in queries
                set qStr to q as text
                set matchedLine to "MISSING|" & qStr
                try
                    set startRng to create range docObj start 0 end 0
                    select startRng
                    set fObj to find object of selection
                    clear formatting fObj
                    repeat 40 times
                        execute find fObj find text (qStr & "^p") match case true match forward true wrap find find stop
                        if not (found of fObj) then exit repeat
                        set selRng to text object of selection
                        set pObj to paragraph 1 of selRng
                        set olStr to (outline level of pObj) as text
                        if (olStr contains "level1" or olStr contains "level2" or olStr contains "level 1" or olStr contains "level 2") then
                            set pTxt to content of (text object of pObj)
                            if (pTxt starts with qStr) and ((length of pTxt) <= (length of qStr) + 2) then
                                set lstStr to ""
                                try
                                    set lstStr to list string of (list format of selRng)
                                end try
                                set matchedLine to ("FOUND|" & qStr & "|" & lstStr & "|" & qStr)
                                exit repeat
                            end if
                        end if
                        set nextPos to end of content of selRng
                        set nextRng to create range docObj start nextPos end nextPos
                        select nextRng
                    end repeat
                on error errQ
                    set matchedLine to ("ERR|" & qStr & "|" & errQ)
                end try
                set end of outLines to matchedLine
            end repeat
            close docObj saving no
            set docObj to missing value
            set display alerts to prevAlerts
            if hideAndQuit then
                quit saving no
            end if
            set AppleScript's text item delimiters to linefeed
            return outLines as text
        on error errMsg number errNum
            try
                if docObj is not missing value then close docObj saving no
            end try
            try
                set display alerts to prevAlerts
            end try
            if hideAndQuit then
                try
                    quit saving no
                end try
            end if
            error errMsg number errNum
        end try
    end tell
end run
'''


def inspect_word_open_and_headings(
    docx_path: str,
    heading_queries: Optional[Sequence[str]] = None,
    *,
    timeout: int = 300,
) -> dict:
    """Open `docx_path` silently read-only in Word and inspect section count, TOC count, and heading list strings."""
    queries = list(heading_queries) if heading_queries else [
        "前言",
        "系统概述",
        "软件功能需求",
        "工作台",
        "数据中心",
        "变量中心",
        "预警中心",
        "关系中心",
        "策略中心",
        "服务中心",
        "运营中心",
        "视图中心",
        "系统管理",
        "其他要求",
    ]

    stage_dir = None
    try:
        stage_dir, staged_docx = _stage_copy_for_word(docx_path, _PRIMARY_STAGE_ROOT)
        was_running = _is_word_running()
        hide_and_quit = "false" if was_running else "true"
        proc = subprocess.run(
            ["osascript", "-", staged_docx, hide_and_quit, *queries],
            input=_INSPECT_APPLESCRIPT,
            text=True,
            capture_output=True,
            timeout=timeout,
        )
        if proc.returncode != 0:
            return {
                "ok": False,
                "error": (proc.stderr or proc.stdout).strip(),
                "sections": 0,
                "tocs": 0,
                "headings": [],
            }
        sections = 0
        tocs = 0
        headings = []
        for raw_line in proc.stdout.splitlines():
            line = raw_line.strip("\r\n")
            if line.startswith("SECTIONS="):
                sections = int(line.split("=", 1)[1])
            elif line.startswith("TOCS="):
                tocs = int(line.split("=", 1)[1])
            elif line.startswith("FOUND|") or line.startswith("MISSING|") or line.startswith("ERR|"):
                parts = line.split("|")
                status = parts[0]
                query = parts[1] if len(parts) > 1 else ""
                list_str = parts[2] if len(parts) > 2 else ""
                text = parts[3].strip("\r\n\x07") if len(parts) > 3 else ""
                headings.append(
                    {"status": status, "query": query, "list_string": list_str, "text": text}
                )
        return {
            "ok": True,
            "error": "",
            "sections": sections,
            "tocs": tocs,
            "headings": headings,
        }
    finally:
        if stage_dir and os.path.exists(stage_dir):
            shutil.rmtree(stage_dir, ignore_errors=True)
