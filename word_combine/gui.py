"""Cross-platform desktop GUI for WordCombine using tkinter/ttk."""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional, Set, Tuple

from . import __version__
from .sorting import collect_docx_files, sort_docx_paths
from .splash import close_splash, update_splash

_TOC_CACHE_TYPE = Dict[str, Tuple[int, int, str]]


def _resource_path(rel_path: str) -> str:
    """Resolve asset paths both in source checkouts and PyInstaller bundles."""
    base = getattr(sys, "_MEIPASS", None)
    if not base:
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel_path)


def _default_desktop_output() -> str:
    home = os.path.expanduser("~")
    desktop = os.path.join(home, "Desktop")
    target_dir = desktop if os.path.isdir(desktop) else home
    return os.path.join(target_dir, "合并文档.docx")


def _configure_windows_dpi_and_theme(root: tk.Tk) -> None:
    if sys.platform != "win32":
        return
    try:
        style = ttk.Style(root)
        available = style.theme_names()
        for name in ("vista", "winnative", "xpnative", "clam"):
            if name in available:
                style.theme_use(name)
                break
        default_font = ("Microsoft YaHei UI", 9)
        root.option_add("*Font", default_font)
        style.configure(".", font=default_font)
        style.configure("Treeview", rowheight=24, font=default_font)
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 9, "bold"))
    except Exception:
        pass


def _enable_hidpi_before_tk() -> None:
    if sys.platform == "win32":
        try:
            import ctypes  # type: ignore

            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                import ctypes  # type: ignore

                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass


class WordCombineApp:
    """Interactive desktop GUI for selecting, sorting, merging, and verifying .docx files."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"Word 文档高保真拼接工具 v{__version__}")
        self.root.geometry("980x720")
        self.root.minsize(820, 600)

        _configure_windows_dpi_and_theme(self.root)
        self._load_app_icon()

        self.file_paths: List[str] = []
        self.file_statuses: Dict[str, str] = {}
        self._toc_cache: _TOC_CACHE_TYPE = {}
        self._toc_pending: Set[str] = set()
        self._toc_scan_queue: "queue.Queue[Optional[Tuple[str, int, int]]]" = queue.Queue()
        self._toc_thread: Optional[threading.Thread] = None
        self._warmup_thread: Optional[threading.Thread] = None
        self._engine_ready = False
        self._stop_event = threading.Event()
        self.worker_thread: Optional[threading.Thread] = None
        self.msg_queue: "queue.Queue[tuple]" = queue.Queue()

        self._output_customized = False
        self._setting_output_programmatically = False
        self._is_running = False
        self._run_start_time: Optional[float] = None
        self._current_pct = 0
        self._current_stage_msg = "就绪：请添加要拼接的 Word (.docx) 文档"
        self._is_indeterminate = False

        self.sort_var = tk.BooleanVar(value=True)
        self.section_mode_var = tk.StringVar(value="auto")
        self.refresh_toc_var = tk.BooleanVar(value=True)
        self.insert_toc_var = tk.BooleanVar(value=False)
        self.toc_levels_var = tk.StringVar(value="1-3")
        self.verify_var = tk.BooleanVar(value=True)
        self.output_var = tk.StringVar(value=_default_desktop_output())
        self.output_var.trace_add("write", self._on_output_edited)
        self.status_var = tk.StringVar(value=self._current_stage_msg)
        self.pct_var = tk.StringVar(value="0%")
        self.elapsed_var = tk.StringVar(value="")

        self._editable_widgets: List[tk.Widget] = []
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.bind("<Destroy>", self._on_destroy, add=True)
        self.root.after(20, self._start_engine_warmup)
        self.root.after(100, self._poll_queue)
        self.root.after(500, self._tick_elapsed)

    def shutdown(self) -> None:
        """Stop background warm-up and TOC scan threads cleanly."""
        if not self._stop_event.is_set():
            self._stop_event.set()
            self._toc_scan_queue.put(None)
        if self._toc_thread and self._toc_thread.is_alive():
            self._toc_thread.join(timeout=1.0)

    def _on_close(self) -> None:
        self.shutdown()
        try:
            self.root.destroy()
        except tk.TclError:
            pass

    def _on_destroy(self, event: tk.Event) -> None:  # type: ignore[type-arg]
        if event.widget is self.root:
            self.shutdown()

    def _start_engine_warmup(self) -> None:
        if self._stop_event.is_set():
            return
        if self._warmup_thread and self._warmup_thread.is_alive():
            return

        def _warmup() -> None:
            try:
                from . import ooxml_merger, toc_refresher, verifier  # noqa: F401
            except Exception:
                pass
            if not self._stop_event.is_set():
                self.msg_queue.put(("engine_ready",))

        self._warmup_thread = threading.Thread(target=_warmup, daemon=True, name="wordcombine-warmup")
        self._warmup_thread.start()

    def _load_app_icon(self) -> None:
        try:
            if sys.platform == "win32":
                ico = _resource_path(os.path.join("assets", "icon.ico"))
                if os.path.isfile(ico):
                    self.root.iconbitmap(default=ico)
                    return
            png = _resource_path(os.path.join("assets", "icon.png"))
            if os.path.isfile(png):
                self._icon_img = tk.PhotoImage(file=png)
                self.root.iconphoto(True, self._icon_img)
        except Exception:
            pass

    def _on_output_edited(self, *_args: object) -> None:
        if not self._setting_output_programmatically:
            self._output_customized = True

    def _set_output_default(self, new_path: str) -> None:
        self._setting_output_programmatically = True
        try:
            self.output_var.set(new_path)
        finally:
            self._setting_output_programmatically = False

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        # Toolbar
        toolbar = ttk.Frame(main)
        toolbar.pack(fill=tk.X, pady=(0, 8))

        btn_add_files = ttk.Button(toolbar, text="添加文件...", command=self._add_files)
        btn_add_files.pack(side=tk.LEFT, padx=(0, 6))
        btn_add_folder = ttk.Button(toolbar, text="添加文件夹...", command=self._add_folder)
        btn_add_folder.pack(side=tk.LEFT, padx=(0, 6))
        btn_sort = ttk.Button(toolbar, text="按文件名自然排序", command=self._sort_files)
        btn_sort.pack(side=tk.LEFT, padx=(0, 12))
        btn_up = ttk.Button(toolbar, text="上移", command=lambda: self._move_selected(-1))
        btn_up.pack(side=tk.LEFT, padx=(0, 4))
        btn_down = ttk.Button(toolbar, text="下移", command=lambda: self._move_selected(1))
        btn_down.pack(side=tk.LEFT, padx=(0, 4))
        btn_rm = ttk.Button(toolbar, text="移除选中", command=self._remove_selected)
        btn_rm.pack(side=tk.LEFT, padx=(0, 6))
        btn_clear = ttk.Button(toolbar, text="清空", command=self._clear_files)
        btn_clear.pack(side=tk.LEFT)

        self._editable_widgets.extend([btn_add_files, btn_add_folder, btn_sort, btn_up, btn_down, btn_rm, btn_clear])

        # File table
        tree_frame = ttk.Frame(main)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        cols = ("order", "name", "size", "toc", "status", "path")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="extended", height=10)
        self.tree.heading("order", text="#")
        self.tree.heading("name", text="文件名")
        self.tree.heading("size", text="大小")
        self.tree.heading("toc", text="目录")
        self.tree.heading("status", text="状态")
        self.tree.heading("path", text="完整路径")
        self.tree.column("order", width=42, anchor="center", stretch=False)
        self.tree.column("name", width=240, anchor="w")
        self.tree.column("size", width=92, anchor="e", stretch=False)
        self.tree.column("toc", width=72, anchor="center", stretch=False)
        self.tree.column("status", width=96, anchor="center", stretch=False)
        self.tree.column("path", width=340, anchor="w")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # Options card
        opts = ttk.LabelFrame(main, text="拼接与目录选项", padding=10)
        opts.pack(fill=tk.X, pady=(0, 10))

        row1 = ttk.Frame(opts)
        row1.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(row1, text="输出文件:").pack(side=tk.LEFT, padx=(0, 6))
        out_entry = ttk.Entry(row1, textvariable=self.output_var)
        out_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        btn_browse = ttk.Button(row1, text="浏览...", command=self._choose_output)
        btn_browse.pack(side=tk.LEFT)
        self._editable_widgets.extend([out_entry, btn_browse])

        row2 = ttk.Frame(opts)
        row2.pack(fill=tk.X, pady=(0, 6))
        chk_sort = ttk.Checkbutton(
            row2,
            text="拼接前按文件名自然排序 (如 3.9 < 3.10)",
            variable=self.sort_var,
        )
        chk_sort.pack(side=tk.LEFT, padx=(0, 16))
        ttk.Label(row2, text="文档分节策略:").pack(side=tk.LEFT, padx=(0, 6))
        sec_cb = ttk.Combobox(
            row2,
            textvariable=self.section_mode_var,
            values=("auto", "page", "continuous"),
            state="readonly",
            width=12,
        )
        sec_cb.pack(side=tk.LEFT)
        ttk.Label(
            row2,
            text="  (auto: 版式一致时直接首尾相连，页眉/页面不同则自动分节)",
            foreground="#555555",
        ).pack(side=tk.LEFT)
        self._editable_widgets.extend([chk_sort, sec_cb])

        row3 = ttk.Frame(opts)
        row3.pack(fill=tk.X)
        chk_refresh = ttk.Checkbutton(
            row3,
            text="拼接后静默刷新目录（后台调用本机 Word / WPS 更新，无弹窗）",
            variable=self.refresh_toc_var,
        )
        chk_refresh.pack(side=tk.LEFT, padx=(0, 16))
        chk_insert = ttk.Checkbutton(
            row3,
            text="在文档开头插入新目录",
            variable=self.insert_toc_var,
        )
        chk_insert.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(row3, text="级别:").pack(side=tk.LEFT, padx=(0, 4))
        lvl_cb = ttk.Combobox(
            row3,
            textvariable=self.toc_levels_var,
            values=("1-2", "1-3", "1-4", "1-5"),
            state="readonly",
            width=5,
        )
        lvl_cb.pack(side=tk.LEFT, padx=(0, 14))
        chk_verify = ttk.Checkbutton(
            row3,
            text="严格校验内容一致性",
            variable=self.verify_var,
        )
        chk_verify.pack(side=tk.LEFT)
        self._editable_widgets.extend([chk_refresh, chk_insert, lvl_cb, chk_verify])

        # Progress & actions
        act_row = ttk.Frame(main)
        act_row.pack(fill=tk.X, pady=(0, 6))
        self.start_btn = ttk.Button(act_row, text="开始拼接", command=self._start_merge)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 8))

        reveal_label = "打开所在文件夹" if sys.platform == "win32" else "在 Finder 中显示"
        open_label = "打开文档" if sys.platform == "win32" else "用 Word 打开"
        self.finder_btn = ttk.Button(act_row, text=reveal_label, command=self._reveal_in_explorer)
        self.finder_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.word_btn = ttk.Button(act_row, text=open_label, command=self._open_document)
        self.word_btn.pack(side=tk.LEFT, padx=(0, 12))

        self.progress = ttk.Progressbar(act_row, mode="determinate", maximum=100)
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Label(act_row, textvariable=self.pct_var, width=5, anchor="e").pack(side=tk.LEFT)

        status_row = ttk.Frame(main)
        status_row.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(status_row, textvariable=self.status_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(status_row, textvariable=self.elapsed_var, foreground="#555555").pack(side=tk.RIGHT, padx=(8, 0))

        # Log area header + text box
        log_hdr = ttk.Frame(main)
        log_hdr.pack(fill=tk.X, pady=(2, 2))
        ttk.Label(log_hdr, text="执行日志:").pack(side=tk.LEFT)
        ttk.Button(log_hdr, text="清空日志", command=self._clear_log).pack(side=tk.RIGHT)

        log_frame = ttk.Frame(main)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD, state=tk.DISABLED)
        log_vsb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_vsb.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_vsb.pack(side=tk.RIGHT, fill=tk.Y)

    def _set_controls_enabled(self, enabled: bool) -> None:
        state_val = tk.NORMAL if enabled else tk.DISABLED
        self.start_btn.configure(state=state_val)
        for widget in self._editable_widgets:
            if isinstance(widget, ttk.Combobox):
                widget.configure(state="readonly" if enabled else tk.DISABLED)
            else:
                widget.configure(state=state_val)

    def _ensure_toc_scanner(self) -> None:
        if self._stop_event.is_set():
            return
        if self._toc_thread and self._toc_thread.is_alive():
            return
        self._toc_thread = threading.Thread(
            target=self._toc_scanner_loop,
            daemon=True,
            name="wordcombine-toc-scan",
        )
        self._toc_thread.start()

    def _toc_scanner_loop(self) -> None:
        while not self._stop_event.is_set():
            item = self._toc_scan_queue.get()
            if item is None or self._stop_event.is_set():
                break
            rp, mtime_ns, size = item
            label = "?"
            try:
                from .ooxml_merger import detect_tocs

                tinfo = detect_tocs(rp)
                label = f"有({tinfo.count})" if tinfo.has_toc else "无"
            except Exception:
                label = "?"
            if not self._stop_event.is_set():
                self.msg_queue.put(("toc_result", rp, mtime_ns, size, label))

    def _lookup_toc_label_nonblocking(self, path: str) -> str:
        try:
            rp = os.path.realpath(path)
            st = os.stat(rp)
        except OSError:
            return "?"
        cached = self._toc_cache.get(rp)
        if cached and cached[0] == st.st_mtime_ns and cached[1] == st.st_size:
            return cached[2]
        if rp not in self._toc_pending:
            self._toc_pending.add(rp)
            self._ensure_toc_scanner()
            self._toc_scan_queue.put((rp, st.st_mtime_ns, st.st_size))
        return "检测中..."

    def _apply_toc_label_to_rows(self, rp: str, label: str) -> None:
        for idx, p in enumerate(self.file_paths):
            if os.path.realpath(p) == rp:
                iid = str(idx)
                if self.tree.exists(iid):
                    vals = list(self.tree.item(iid, "values"))
                    if len(vals) >= 6 and vals[3] != label:
                        vals[3] = label
                        self.tree.item(iid, values=vals)

    def _append_log(self, text: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _clear_log(self) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _refresh_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, path in enumerate(self.file_paths, start=1):
            name = os.path.basename(path)
            size_str = f"{os.path.getsize(path) / 1024:,.1f} KB" if os.path.isfile(path) else "-"
            toc_str = self._lookup_toc_label_nonblocking(path)
            status_str = self.file_statuses.get(os.path.realpath(path), "待处理")
            self.tree.insert("", tk.END, iid=str(idx - 1), values=(idx, name, size_str, toc_str, status_str, path))

    def _update_row_status(self, idx: int, status_label: str) -> None:
        if 0 <= idx < len(self.file_paths):
            rp = os.path.realpath(self.file_paths[idx])
            self.file_statuses[rp] = status_label
            iid = str(idx)
            if self.tree.exists(iid):
                vals = list(self.tree.item(iid, "values"))
                if len(vals) >= 6:
                    vals[4] = status_label
                    self.tree.item(iid, values=vals)
                    self.tree.see(iid)

    def _merge_added_paths(self, new_paths: List[str]) -> None:
        out_abs = os.path.realpath(self.output_var.get()) if self.output_var.get() else ""
        existing = {os.path.realpath(p) for p in self.file_paths}
        for p in new_paths:
            rp = os.path.realpath(p)
            if rp == out_abs or rp in existing:
                continue
            existing.add(rp)
            self.file_paths.append(p)
            self.file_statuses[rp] = "待处理"
        if self.sort_var.get():
            self.file_paths = sort_docx_paths(self.file_paths)
        if self.file_paths and not self._output_customized:
            first_dir = os.path.dirname(os.path.abspath(self.file_paths[0]))
            self._set_output_default(os.path.join(first_dir, "合并文档.docx"))
        self._refresh_tree()
        if self.file_paths:
            self.status_var.set(f"已选 {len(self.file_paths)} 个文档，准备就绪")

    def _add_files(self) -> None:
        files = filedialog.askopenfilenames(
            title="选择要拼接的 Word 文档",
            filetypes=[("Word 文档", "*.docx")],
        )
        if files:
            valid = collect_docx_files(list(files), sort=False)
            self._merge_added_paths(valid)

    def _add_folder(self) -> None:
        folder = filedialog.askdirectory(title="选择包含 Word 文档的文件夹")
        if folder:
            out_abs = os.path.abspath(self.output_var.get()) if self.output_var.get() else ""
            valid = collect_docx_files(folder, exclude_paths=[out_abs] if out_abs else None, sort=True)
            self._merge_added_paths(valid)

    def _sort_files(self) -> None:
        self.file_paths = sort_docx_paths(self.file_paths)
        self._refresh_tree()

    def _move_selected(self, delta: int) -> None:
        sel = [int(i) for i in self.tree.selection()]
        if not sel:
            return
        if delta < 0:
            for idx in sorted(sel):
                if idx + delta >= 0 and (idx + delta) not in sel:
                    self.file_paths[idx], self.file_paths[idx + delta] = (
                        self.file_paths[idx + delta],
                        self.file_paths[idx],
                    )
        else:
            for idx in sorted(sel, reverse=True):
                if idx + delta < len(self.file_paths) and (idx + delta) not in sel:
                    self.file_paths[idx], self.file_paths[idx + delta] = (
                        self.file_paths[idx + delta],
                        self.file_paths[idx],
                    )
        self._refresh_tree()

    def _remove_selected(self) -> None:
        sel = {int(i) for i in self.tree.selection()}
        if not sel:
            return
        self.file_paths = [p for i, p in enumerate(self.file_paths) if i not in sel]
        self._refresh_tree()

    def _clear_files(self) -> None:
        self.file_paths.clear()
        self.file_statuses.clear()
        self._refresh_tree()
        self.status_var.set("已清空文件列表")
        self.pct_var.set("0%")
        self.elapsed_var.set("")
        self.progress["value"] = 0

    def _choose_output(self) -> None:
        initial = self.output_var.get().strip() or _default_desktop_output()
        path = filedialog.asksaveasfilename(
            title="保存拼接后的 Word 文档",
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")],
            initialdir=os.path.dirname(initial),
            initialfile=os.path.basename(initial),
        )
        if path:
            self._output_customized = True
            self.output_var.set(path)

    def _reveal_in_explorer(self) -> None:
        out = self.output_var.get().strip()
        if out and os.path.exists(out):
            if sys.platform == "win32":
                subprocess.Popen(["explorer", "/select,", os.path.normpath(out)])
            else:
                subprocess.Popen(["open", "-R", out])
        else:
            messagebox.showinfo("提示", "输出文件尚未生成。")

    def _open_document(self) -> None:
        out = self.output_var.get().strip()
        if out and os.path.exists(out):
            if sys.platform == "win32":
                os.startfile(os.path.normpath(out))  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["open", "-a", "Microsoft Word", out])
        else:
            messagebox.showinfo("提示", "输出文件尚未生成。")

    def _set_progress_mode(self, indeterminate: bool) -> None:
        if indeterminate and not self._is_indeterminate:
            self.progress.configure(mode="indeterminate")
            self.progress.start(15)
            self._is_indeterminate = True
        elif not indeterminate and self._is_indeterminate:
            self.progress.stop()
            self.progress.configure(mode="determinate")
            self._is_indeterminate = False

    def _format_elapsed(self) -> str:
        if self._run_start_time is None:
            return ""
        secs = max(0, int(time.monotonic() - self._run_start_time))
        mins, rem = divmod(secs, 60)
        return f"已用时 {mins:02d}:{rem:02d}"

    def _tick_elapsed(self) -> None:
        if self._stop_event.is_set():
            return
        try:
            if self._is_running and self._run_start_time is not None:
                self.elapsed_var.set(self._format_elapsed())
            if not self._stop_event.is_set():
                self.root.after(500, self._tick_elapsed)
        except tk.TclError:
            pass

    def _start_merge(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            return
        if not self.file_paths:
            messagebox.showwarning("提示", "请先添加至少一个 .docx 文档。")
            return
        out_path = self.output_var.get().strip()
        if not out_path:
            messagebox.showwarning("提示", "请指定输出文件路径。")
            return

        files = sort_docx_paths(self.file_paths) if self.sort_var.get() else list(self.file_paths)
        if self.sort_var.get():
            self.file_paths = files
        out_real = os.path.realpath(out_path)
        if any(os.path.realpath(p) == out_real for p in files):
            messagebox.showerror("错误", "输出路径不能与源文档相同，请另选输出路径。")
            return

        for p in self.file_paths:
            self.file_statuses[os.path.realpath(p)] = "待处理"
        self._refresh_tree()

        self._is_running = True
        self._run_start_time = time.monotonic()
        self._set_controls_enabled(False)
        self._set_progress_mode(False)
        self.progress["value"] = 0
        self.pct_var.set("0%")
        self.elapsed_var.set("已用时 00:00")
        self._append_log("=== 开始执行拼接 ===")
        for i, p in enumerate(files, 1):
            self._append_log(f"  {i}. {os.path.basename(p)}")

        sec_mode = self.section_mode_var.get()
        do_refresh = self.refresh_toc_var.get()
        do_insert_toc = self.insert_toc_var.get()
        toc_lvls = self.toc_levels_var.get()
        do_verify = self.verify_var.get()

        def _worker() -> None:
            try:
                from .ooxml_merger import merge_docx_files
                from .toc_refresher import refresh_toc_silent
                from .verifier import verify_merged_docx

                def _prog(msg: str, cur: int, total: int) -> None:
                    pct = int((cur / max(total, 1)) * 70)
                    self.msg_queue.put(("progress", pct, f"阶段 1/3：{msg}", False))

                def _file_status(idx: int, st: str) -> None:
                    label = "合并中..." if st == "processing" else "已完成 ✓"
                    self.msg_queue.put(("file_status", idx, label))

                res = merge_docx_files(
                    files,
                    out_path,
                    sort_by_name=False,
                    section_break_mode=sec_mode,
                    insert_toc=do_insert_toc,
                    toc_levels=toc_lvls,
                    progress_callback=_prog,
                    file_status_callback=_file_status,
                )
                self.msg_queue.put(
                    (
                        "log",
                        f"OOXML 拼接完成: {res.output_path} (sections={res.section_count}, comments={res.comments_merged})",
                    )
                )

                refreshed = False
                if do_refresh:
                    if res.has_toc:
                        self.msg_queue.put(("progress", 78, "阶段 2/3：正在后台静默刷新目录（请稍候）...", True))
                    else:
                        self.msg_queue.put(("progress", 82, "阶段 2/3：检查目录状态...", False))
                    r_res = refresh_toc_silent(
                        out_path,
                        log_callback=lambda m: self.msg_queue.put(("log", m)),
                    )
                    self.msg_queue.put(("progress", 88, f"阶段 2/3：{r_res.message}", False))
                    if r_res.status == "error":
                        self.msg_queue.put(("done", False, r_res.message))
                        return
                    if r_res.status == "skipped":
                        self.msg_queue.put(("log", f"[提示] {r_res.message}"))
                    refreshed = r_res.status == "updated"

                if do_verify:
                    self.msg_queue.put(("progress", 92, "阶段 3/3：正在执行严格内容一致性校验...", False))
                    rep = verify_merged_docx(
                        files,
                        out_path,
                        ignore_toc_content=bool(do_insert_toc or res.has_toc),
                        check_media_hashes=not refreshed,
                        word_resaved=refreshed,
                    )
                    self.msg_queue.put(("log", rep.format_summary()))
                    if not rep.ok:
                        self.msg_queue.put(("done", False, "内容校验发现差异，请查看日志。"))
                        return

                self.msg_queue.put(("progress", 100, "全部完成！", False))
                self.msg_queue.put(("done", True, f"拼接成功：{out_path}"))
            except Exception as exc:
                self.msg_queue.put(("done", False, f"执行异常: {exc}"))

        self.worker_thread = threading.Thread(target=_worker, daemon=True)
        self.worker_thread.start()

    def _poll_queue(self) -> None:
        if self._stop_event.is_set():
            return
        try:
            while True:
                try:
                    item = self.msg_queue.get_nowait()
                except queue.Empty:
                    break
                kind = item[0]
                if kind == "engine_ready":
                    self._engine_ready = True
                elif kind == "toc_result":
                    _, rp, mtime_ns, size, label = item
                    self._toc_pending.discard(rp)
                    self._toc_cache[rp] = (mtime_ns, size, label)
                    self._apply_toc_label_to_rows(rp, label)
                elif kind == "progress":
                    _, pct, msg, is_indet = item
                    self._set_progress_mode(bool(is_indet))
                    if not is_indet:
                        self.progress["value"] = pct
                    self.pct_var.set(f"{pct}%")
                    self.status_var.set(msg)
                    self._append_log(msg)
                elif kind == "file_status":
                    _, idx, status_label = item
                    self._update_row_status(idx, status_label)
                elif kind == "log":
                    self._append_log(item[1])
                elif kind == "done":
                    _, ok, msg = item
                    self._is_running = False
                    self._set_progress_mode(False)
                    self._set_controls_enabled(True)
                    elapsed_str = self._format_elapsed()
                    if elapsed_str:
                        self.elapsed_var.set(elapsed_str)
                    if ok:
                        self.progress["value"] = 100
                        self.pct_var.set("100%")
                    self.status_var.set(msg)
                    self._append_log(f"{msg} ({elapsed_str})" if elapsed_str else msg)
                    if ok:
                        messagebox.showinfo("完成", msg)
                    else:
                        messagebox.showerror("失败", msg)
            if not self._stop_event.is_set():
                self.root.after(100, self._poll_queue)
        except tk.TclError:
            pass


def _center_window(root: tk.Tk, width: int = 980, height: int = 720) -> None:
    root.update_idletasks()
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    x = max(0, (sw - width) // 2)
    y = max(0, (sh - height) // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")


def create_app_window(show: bool = True) -> tuple[tk.Tk, WordCombineApp]:
    _enable_hidpi_before_tk()
    root = tk.Tk()
    root.withdraw()
    update_splash("Preparing window...")
    app = WordCombineApp(root)
    _center_window(root, 980, 720)
    if show:
        root.deiconify()
        root.lift()
        root.update()
        close_splash()
    return root, app


def launch_gui() -> None:
    root, _ = create_app_window(show=True)
    root.mainloop()
