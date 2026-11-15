import ctypes
from datetime import datetime
import json
import os
import queue
import re
import sys
import threading
import time
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from ai_client import (
    API_MODES,
    ROTATION_MODES,
    SUPPORTED_MODELS,
    ApiKeyPool,
    test_all_api_keys,
    test_api_connection,
)
from fellowship_processor import (
    OUTPUT_MODES,
    QUOTE_MODES,
    build_fellowship_output_name,
    find_post_quote_citation_spans,
    inspect_fellowship_docx_quick,
    natural_sort_docx_paths,
    parse_title_and_part,
    run_fellowship_pipeline,
)
from lang_registry import (
    AUTO_COMBO_LABEL,
    DEFAULT_LANGUAGES,
    build_keywords_list,
    detect_lang_code_from_hint,
    find_lang_display_name,
    format_combo_label,
    get_combo_values,
    get_languages_file,
    get_tri_prefix,
    load_languages,
    resolve_combo_to_code,
    save_languages,
)
from pipeline import run_step1_pipeline, run_step2_pipeline
from security_utils import (
    dpapi_decrypt_str,
    dpapi_encrypt_str,
    load_api_key_from_env,
    parse_api_keys,
    redact_secrets,
)
from word_counter import (
    analyze_docx_word_count,
    apply_word_count_filename,
    format_counted_docx_name,
    parse_doc_filename,
)
from win_dnd import WinFileDropManager


def _enable_high_dpi():
    if sys.platform == "win32":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass


def _app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


SETTINGS_FILE = _app_base_dir() / "app_settings.json"


class ApiKeyPoolDialog(tk.Toplevel):
    """Modal dialog for managing an unlimited pool of API keys, batch health testing, and dead-key pruning."""

    def __init__(
        self,
        parent: tk.Tk,
        initial_raw_text: str,
        rotation_label_var: tk.StringVar,
        model_id: str,
        api_mode: str,
        on_save_cb,
    ):
        super().__init__(parent)
        self.title("API 密钥池批量管理（支持无限个 Key 轮换 + 超额/报错自动秒切）")
        self.geometry("760x580")
        self.minsize(680, 500)
        self.transient(parent)
        self.grab_set()

        self.initial_raw_text = initial_raw_text or ""
        self.rotation_var = tk.StringVar(value=rotation_label_var.get())
        self.model_id = model_id
        self.api_mode = api_mode
        self.on_save_cb = on_save_cb

        self.count_var = tk.StringVar(value="当前共识别到 0 个不重复的 API Key")
        self._testing = False
        self._failed_keys_set: set[str] = set()
        self._ui_queue: queue.Queue = queue.Queue()

        self._build_ui()
        if self.initial_raw_text.strip():
            self.editor.insert("1.0", self.initial_raw_text.strip() + "\n")
        self._update_count()
        self._poll_ui_queue()

    def _build_ui(self):
        wrap = ttk.Frame(self, padding=12)
        wrap.pack(fill=tk.BOTH, expand=True)

        tip = ttk.Label(
            wrap,
            text=(
                "说明：支持填入【无限个 API Key】（推荐每行 1 个，也支持逗号/分号/竖线分隔；可在 Key 后加 # 备注，自动去重）。\n"
                "• 支持混填 Vertex AI Key (AQ.开头) 与 Gemini Key (AIza开头)；本机勾选记住时全程采用 Windows DPAPI 系统级加密存储。\n"
                "• 运行时若任一 Key 额度耗尽、超速 (HTTP 429)、失效报错 (401/403) 或超时，将毫秒级自动切换下一个 Key，绝不中断任务！"
            ),
            foreground="#475569",
        )
        tip.pack(fill=tk.X, pady=(0, 8))

        ed_frame = ttk.LabelFrame(wrap, text=" 密钥列表（每行 1 个 Key，支持 # 账号备注，数量无上限） ", padding=8)
        ed_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        ed_inner = ttk.Frame(ed_frame)
        ed_inner.pack(fill=tk.BOTH, expand=True)

        self.editor = tk.Text(
            ed_inner,
            height=10,
            wrap=tk.NONE,
            font=("Consolas", 9),
            bg="#ffffff",
            fg="#0f172a",
            relief=tk.SOLID,
            borderwidth=1,
            padx=8,
            pady=6,
            undo=True,
        )
        vsb = ttk.Scrollbar(ed_inner, orient=tk.VERTICAL, command=self.editor.yview)
        hsb = ttk.Scrollbar(ed_frame, orient=tk.HORIZONTAL, command=self.editor.xview)
        self.editor.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.editor.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(fill=tk.X)

        self.editor.bind("<KeyRelease>", lambda _e: self._update_count())
        self.editor.bind("<<Paste>>", lambda _e: self.after(20, self._update_count))

        mid_row = ttk.Frame(wrap)
        mid_row.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(
            mid_row,
            textvariable=self.count_var,
            font=("Microsoft YaHei UI", 9, "bold"),
            foreground="#1d4ed8",
        ).pack(side=tk.LEFT)

        rot_box = ttk.Frame(mid_row)
        rot_box.pack(side=tk.RIGHT)
        ttk.Label(rot_box, text="多 Key 调度策略：").pack(side=tk.LEFT)
        ttk.Combobox(
            rot_box,
            textvariable=self.rotation_var,
            values=[m[0] for m in ROTATION_MODES],
            state="readonly",
            width=38,
        ).pack(side=tk.LEFT, padx=(4, 0))

        chk_frame = ttk.LabelFrame(wrap, text=f" 批量可用性检测结果（当前检测模型：{self.model_id}） ", padding=6)
        chk_frame.pack(fill=tk.BOTH, expand=False, pady=(0, 8))

        self.chk_text = tk.Text(
            chk_frame,
            height=6,
            wrap=tk.WORD,
            font=("Microsoft YaHei UI", 9),
            bg="#f8fafc",
            fg="#0f172a",
            relief=tk.SOLID,
            borderwidth=1,
            padx=6,
            pady=4,
        )
        self.chk_text.tag_configure("ok", foreground="#15803d")
        self.chk_text.tag_configure("bad", foreground="#b91c1c")
        self.chk_text.tag_configure("info", foreground="#1d4ed8", font=("Microsoft YaHei UI", 9, "bold"))
        chk_vsb = ttk.Scrollbar(chk_frame, orient=tk.VERTICAL, command=self.chk_text.yview)
        self.chk_text.configure(yscrollcommand=chk_vsb.set)
        self.chk_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        chk_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.chk_text.insert(tk.END, "点击下方「🧪 批量检测全部 Key」可逐一验证所有密钥可用性，并支持一键剔除失效 Key。\n")

        btns = ttk.Frame(wrap)
        btns.pack(fill=tk.X)

        self.btn_test_all = ttk.Button(btns, text="🧪 批量检测全部 Key", command=self._on_test_all)
        self.btn_test_all.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_prune = ttk.Button(
            btns, text="🧹 一键剔除失效 Key", state=tk.DISABLED, command=self._on_prune_failed
        )
        self.btn_prune.pack(side=tk.LEFT, padx=(0, 8))

        ttk.Button(btns, text="🗑️ 清空输入", command=self._on_clear).pack(side=tk.LEFT)

        ttk.Button(btns, text="取消", width=9, command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(
            btns, text="💾 保存并应用密钥池", style="Primary.TButton", command=self._on_save
        ).pack(side=tk.RIGHT, padx=(0, 8))

    def _get_raw_text(self) -> str:
        return self.editor.get("1.0", tk.END).strip()

    def _update_count(self):
        keys = parse_api_keys(self._get_raw_text())
        self.count_var.set(f"当前共识别到 {len(keys)} 个不重复的 API Key")

    def _on_clear(self):
        if self._get_raw_text() and not messagebox.askyesno("确认清空", "确定要清空输入框中的所有 API Key 吗？", parent=self):
            return
        self.editor.delete("1.0", tk.END)
        self._failed_keys_set.clear()
        self.btn_prune.configure(state=tk.DISABLED)
        self._update_count()

    def _poll_ui_queue(self):
        if not self.winfo_exists():
            return
        try:
            while True:
                try:
                    kind, payload = self._ui_queue.get_nowait()
                except queue.Empty:
                    break
                if kind == "line":
                    txt, tag = payload
                    if tag:
                        self.chk_text.insert(tk.END, txt + "\n", tag)
                    else:
                        self.chk_text.insert(tk.END, txt + "\n")
                    self.chk_text.see(tk.END)
                elif kind == "done":
                    ok_cnt, bad_keys = payload
                    self._testing = False
                    self.btn_test_all.configure(state=tk.NORMAL, text="🧪 批量检测全部 Key")
                    self._failed_keys_set = set(bad_keys)
                    self.btn_prune.configure(state=tk.NORMAL if bad_keys else tk.DISABLED)
                    summary = f"检测完毕：✅ 可用 {ok_cnt} 个  |  ❌ 不可用/受限 {len(bad_keys)} 个。"
                    if bad_keys:
                        summary += "（可点击「🧹 一键剔除失效 Key」自动清理）"
                    self.chk_text.insert(tk.END, summary + "\n", "info")
                    self.chk_text.see(tk.END)
        finally:
            if self.winfo_exists():
                self.after(80, self._poll_ui_queue)

    def _on_test_all(self):
        if self._testing:
            return
        keys = parse_api_keys(self._get_raw_text())
        if not keys:
            messagebox.showwarning("提示", "请先输入至少 1 个 API Key。", parent=self)
            return

        self._testing = True
        self._failed_keys_set.clear()
        self.btn_prune.configure(state=tk.DISABLED)
        self.btn_test_all.configure(state=tk.DISABLED, text="⏳ 正在逐个检测...")
        self.chk_text.delete("1.0", tk.END)
        self.chk_text.insert(tk.END, f"开始检测共 {len(keys)} 个 API Key（模型：{self.model_id}）...\n", "info")

        def _each(done_i: int, tot: int, item: dict):
            icon = "✅" if item["ok"] else "❌"
            tag = "ok" if item["ok"] else "bad"
            self._ui_queue.put(("line", (f"[{done_i}/{tot}] {icon} {item['masked']}：{item['detail']}", tag)))

        def worker():
            try:
                reps = test_all_api_keys(
                    keys,
                    model=self.model_id,
                    api_mode=self.api_mode,
                    on_each_cb=_each,
                )
                ok_cnt = sum(1 for r in reps if r["ok"])
                bad_keys = [r["key"] for r in reps if not r["ok"]]
                self._ui_queue.put(("done", (ok_cnt, bad_keys)))
            except Exception as ex:
                safe_err = redact_secrets(str(ex), extra_secret=keys)
                self._ui_queue.put(("line", (f"❌ 检测出错：{safe_err}", "bad")))
                self._ui_queue.put(("done", (0, [])))

        threading.Thread(target=worker, daemon=True).start()

    def _on_prune_failed(self):
        if not self._failed_keys_set:
            return
        raw_lines = self._get_raw_text().splitlines()
        kept_lines = []
        removed_cnt = 0
        seen_good: set[str] = set()

        for raw_line in raw_lines:
            stripped = raw_line.strip()
            if not stripped or stripped.startswith("#"):
                kept_lines.append(raw_line)
                continue
            comment_part = ""
            body = stripped
            if "#" in stripped:
                body, cmt = stripped.split("#", 1)
                comment_part = f"  # {cmt.strip()}"
            tokens = [t.strip().strip("'\"") for t in re.split(r"[\s,，;；|]+", body) if t.strip().strip("'\"")]
            good_tokens = []
            for t in tokens:
                if t in self._failed_keys_set:
                    removed_cnt += 1
                elif t not in seen_good:
                    seen_good.add(t)
                    good_tokens.append(t)
            if good_tokens:
                if len(good_tokens) == 1:
                    kept_lines.append(good_tokens[0] + comment_part)
                else:
                    kept_lines.append(" | ".join(good_tokens) + comment_part)

        self.editor.delete("1.0", tk.END)
        new_txt = "\n".join(kept_lines).strip()
        if new_txt:
            self.editor.insert("1.0", new_txt + "\n")
        self._failed_keys_set.clear()
        self.btn_prune.configure(state=tk.DISABLED)
        self._update_count()
        messagebox.showinfo("清理完成", f"已自动剔除 {removed_cnt} 个检测失败的 API Key！", parent=self)

    def _on_save(self):
        raw_txt = self._get_raw_text()
        self.on_save_cb(raw_txt, self.rotation_var.get())
        self.destroy()


class LanguageManagerDialog(tk.Toplevel):
    """Modal dialog for viewing, adding, editing, and deleting languages."""

    def __init__(self, parent: tk.Tk, entries: list, on_saved_callback):
        super().__init__(parent)
        self.title("多语种配置管理（支持自由新增、修改、删除各语种）")
        self.geometry("720x520")
        self.minsize(660, 460)
        self.transient(parent)
        self.grab_set()

        self.entries = [dict(x) for x in entries]
        self.on_saved_callback = on_saved_callback

        self.code_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.kw_var = tk.StringVar()

        self._build_ui()
        self._refresh_tree()

    def _build_ui(self):
        wrap = ttk.Frame(self, padding=12)
        wrap.pack(fill=tk.BOTH, expand=True)

        tip = ttk.Label(
            wrap,
            text=(
                "说明：所有语种实时保存在同目录的 languages.json 中。\n"
                "• 语种代码：用于生成最终定稿文件名前缀（如 GR_xxx.docx、TH_xxx.docx）。\n"
                "• 自动识别关键词：当选择“AUTO 自动识别”时，若总库文件夹名或文档名含任一关键词即自动匹配该语种。"
            ),
            foreground="#475569",
        )
        tip.pack(fill=tk.X, pady=(0, 8))

        tbl_frame = ttk.Frame(wrap)
        tbl_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        cols = ("code", "name", "keywords")
        self.tree = ttk.Treeview(tbl_frame, columns=cols, show="headings", selectmode="browse", height=11)
        self.tree.heading("code", text="语种代码 (定稿前缀)")
        self.tree.heading("name", text="语种名称")
        self.tree.heading("keywords", text="自动识别关键词（目录名 / 文件名包含任一词即可）")

        self.tree.column("code", width=130, anchor=tk.CENTER)
        self.tree.column("name", width=140, anchor=tk.CENTER)
        self.tree.column("keywords", width=390, anchor=tk.W)

        vsb = ttk.Scrollbar(tbl_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<<TreeviewSelect>>", self._on_select_row)

        form = ttk.LabelFrame(wrap, text=" 新增 / 修改语种（点击上方列表行可直接载入编辑） ", padding=10)
        form.pack(fill=tk.X, pady=(0, 8))
        form.columnconfigure(3, weight=1)

        ttk.Label(form, text="语种代码(英文缩写)：").grid(row=0, column=0, sticky=tk.W, pady=4)
        ttk.Entry(form, textvariable=self.code_var, width=12).grid(row=0, column=1, sticky=tk.W, pady=4, padx=(4, 14))

        ttk.Label(form, text="语种名称(中文)：").grid(row=0, column=2, sticky=tk.W, pady=4)
        ttk.Entry(form, textvariable=self.name_var, width=20).grid(row=0, column=3, sticky=tk.W, pady=4, padx=4)

        ttk.Label(form, text="自动识别关键词：").grid(row=1, column=0, sticky=tk.W, pady=4)
        ttk.Entry(form, textvariable=self.kw_var).grid(row=1, column=1, columnspan=3, sticky=tk.EW, pady=4, padx=4)
        ttk.Label(
            form,
            text="（多个关键词用逗号分隔，如：塞尔维亚语, 塞语, 中英塞；留空也会自动按名称生成）",
            foreground="#64748b",
        ).grid(row=2, column=1, columnspan=3, sticky=tk.W, padx=4)

        btns = ttk.Frame(wrap)
        btns.pack(fill=tk.X)

        ttk.Button(btns, text="➕ 新增 / 保存语种", style="Primary.TButton", command=self._on_add_or_update).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(btns, text="🗑️ 删除选中语种", command=self._on_delete).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btns, text="↺ 补全内置默认语种", command=self._on_merge_defaults).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btns, text="📂 打开 languages.json", command=self._on_open_json).pack(side=tk.LEFT)
        ttk.Button(btns, text="关闭", width=10, command=self.destroy).pack(side=tk.RIGHT)

    def _refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, e in enumerate(self.entries):
            self.tree.insert(
                "",
                tk.END,
                iid=str(idx),
                values=(e["code"], e["name"], ", ".join(e.get("keywords", []))),
            )

    def _on_select_row(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        if 0 <= idx < len(self.entries):
            e = self.entries[idx]
            self.code_var.set(e["code"])
            self.name_var.set(e["name"])
            self.kw_var.set(", ".join(e.get("keywords", [])))

    def _on_add_or_update(self):
        raw_code = re.sub(r"[^A-Za-z0-9_-]", "", self.code_var.get().strip()).upper()
        name = self.name_var.get().strip()
        if not raw_code or raw_code == "AUTO":
            messagebox.showwarning("提示", "请填写有效的英文语种代码（例如 TH、NL、SR，不能为 AUTO）。", parent=self)
            return
        if not name:
            messagebox.showwarning("提示", "请填写语种中文名称（例如 泰语、塞尔维亚语）。", parent=self)
            return

        kws = build_keywords_list(name, self.kw_var.get())
        new_entry = {"code": raw_code, "name": name, "keywords": kws}

        updated = False
        for i, e in enumerate(self.entries):
            if e["code"] == raw_code:
                self.entries[i] = new_entry
                updated = True
                break
        if not updated:
            self.entries.insert(0, new_entry)

        save_languages(self.entries)
        self._refresh_tree()
        self.on_saved_callback(self.entries, selected_code=raw_code)
        action_str = "已更新" if updated else "已新增"
        messagebox.showinfo("保存成功", f"{action_str}语种：{raw_code} - {name}\n识别关键词：{', '.join(kws)}", parent=self)

    def _on_delete(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在列表中选中要删除的语种。", parent=self)
            return
        idx = int(sel[0])
        e = self.entries[idx]
        if not messagebox.askyesno("确认删除", f"确定要从列表中移除【{e['code']} - {e['name']}】吗？", parent=self):
            return
        self.entries.pop(idx)
        save_languages(self.entries)
        self._refresh_tree()
        self.on_saved_callback(self.entries, selected_code=None)

    def _on_merge_defaults(self):
        existing_codes = {e["code"] for e in self.entries}
        added = 0
        for d in DEFAULT_LANGUAGES:
            if d["code"] not in existing_codes:
                self.entries.append(dict(d))
                added += 1
        save_languages(self.entries)
        self._refresh_tree()
        self.on_saved_callback(self.entries, selected_code=None)
        messagebox.showinfo("完成", f"已补全 {added} 个内置语种（您自定义的语种均完整保留）。", parent=self)

    def _on_open_json(self):
        p = get_languages_file().resolve()
        if not p.exists():
            save_languages(self.entries)
        if p.is_file() and p.suffix.lower() == ".json":
            os.startfile(str(p))


class ScatteredRevisionApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("零散改点自动整理工具 v1.5.1  (零散改点 · 字数统计 · 弟兄交通整理 · 支持文件/文件夹直接拖拽)")
        self.root.geometry("1120x825")
        self.root.minsize(980, 700)

        self.log_queue: queue.Queue = queue.Queue()
        self.is_running = False
        self.task_start_ts: float = 0.0
        self.current_pct: float = 0.0
        self.languages: list = load_languages()

        self._pool_multiline_text: str = ""
        self._syncing_key_var: bool = False
        self.wc_files: list[str] = []
        self.fel_files: list[str] = []
        self._fel_inspect_cache: dict[str, dict] = {}
        self.active_tab_key: str = "main"
        self._nav_items: dict = {}
        self.dnd = WinFileDropManager(self.root, fallback_cb=self._on_drop_fallback)

        self._setup_styles()
        self._init_vars()
        self._load_settings()
        self._build_ui()
        self._setup_drag_and_drop()
        self._refresh_key_pool_badge()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._poll_log_queue()
        self._tick_elapsed_timer()

    def _setup_styles(self):
        style = ttk.Style(self.root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")

        default_font = ("Microsoft YaHei UI", 9)
        bold_font = ("Microsoft YaHei UI", 9, "bold")
        btn_font = ("Microsoft YaHei UI", 10, "bold")
        tab_font = ("Microsoft YaHei UI", 10, "bold")

        self.root.option_add("*Font", default_font)
        style.configure("TLabelframe.Label", font=bold_font, foreground="#1a56db")
        style.configure("Primary.TButton", font=btn_font, padding=(12, 7))
        style.configure("Action.TButton", font=default_font, padding=(10, 6))
        style.configure("TNotebook.Tab", font=tab_font, padding=(16, 7))

    def _init_vars(self):
        # Tab 1 vars
        self.api_mode_label = tk.StringVar(value=API_MODES[0][0])
        self.rotation_label = tk.StringVar(value=ROTATION_MODES[0][0])
        self.api_key_var = tk.StringVar(value="")
        self.pool_btn_text_var = tk.StringVar(value="🔑 批量管理 Key (0 个)")
        self.show_key_var = tk.BooleanVar(value=False)
        self.remember_key_var = tk.BooleanVar(value=True)
        self.model_label = tk.StringVar(value=SUPPORTED_MODELS[0][0])

        self.input_docx_var = tk.StringVar(value="")
        self.library_dir_var = tk.StringVar(value="")
        self.output_dir_var = tk.StringVar(value="")
        self.draft_docx_var = tk.StringVar(value="")

        self.lang_var = tk.StringVar(value=AUTO_COMBO_LABEL)
        self.doc_name_var = tk.StringVar(value="")
        self.dual_outputs_var = tk.BooleanVar(value=True)

        self.progress_var = tk.DoubleVar(value=0.0)
        self.stage_var = tk.StringVar(value="状态：就绪（请选择文档与目录后点击上方按钮执行）")
        self.pct_timer_var = tk.StringVar(value="0%")

        # Tab 2 (Word Count & Filename Updater) vars
        self.wc_lang_var = tk.StringVar(value=AUTO_COMBO_LABEL)
        self.wc_keep_existing_lang_var = tk.BooleanVar(value=True)
        self.wc_date_var = tk.StringVar(value=datetime.now().strftime("%Y%m%d"))
        self.wc_keep_existing_date_var = tk.BooleanVar(value=False)
        self.wc_strip_draft_var = tk.BooleanVar(value=True)
        self.wc_save_mode_var = tk.StringVar(value="inplace")  # "inplace" or "copy"
        self.wc_copy_dir_var = tk.StringVar(value="")
        self.wc_summary_var = tk.StringVar(
            value="当前共 0 份文档  |  第 2 列英文合计：UP 0 words, RV 0 words"
        )

        # Tab 3 (Scenario 4: Brother Fellowship Manuscript Processor) vars
        self.fel_output_dir_var = tk.StringVar(value="")
        self.fel_lang_var = tk.StringVar(value=AUTO_COMBO_LABEL)
        self.fel_tri_prefix_var = tk.StringVar(value="")
        self.fel_date_var = tk.StringVar(value=datetime.now().strftime("%Y%m%d"))
        self.fel_quote_label_var = tk.StringVar(value=QUOTE_MODES[0][0])
        self.fel_output_mode_var = tk.StringVar(value="single_only")  # "single_only", "merged_only", "both"
        self.fel_enable_ai_var = tk.BooleanVar(value=True)
        self.fel_summary_var = tk.StringVar(
            value="当前共 0 份分篇文档  |  支持直接拖入整个分篇文件夹或多个 .docx 自动排序"
        )
        self.fel_progress_var = tk.DoubleVar(value=0.0)
        self.fel_stage_var = tk.StringVar(value="状态：就绪（拖入文件夹或多个 .docx 稿件后点击右侧按钮执行整理）")
        self.fel_pct_timer_var = tk.StringVar(value="0%")

        self.api_key_var.trace_add("write", self._on_key_entry_changed)

    def _on_key_entry_changed(self, *_args):
        if self._syncing_key_var:
            return
        keys = parse_api_keys(self.api_key_var.get())
        self._pool_multiline_text = "\n".join(keys)
        self._refresh_key_pool_badge()

    def _set_pool_multiline_and_sync_entry(self, raw_multiline: str):
        self._pool_multiline_text = (raw_multiline or "").strip()
        keys = parse_api_keys(self._pool_multiline_text)
        self._syncing_key_var = True
        try:
            self.api_key_var.set(" | ".join(keys))
        finally:
            self._syncing_key_var = False
        self._refresh_key_pool_badge()

    def _get_all_keys_list(self) -> list[str]:
        if self._pool_multiline_text.strip():
            parsed = parse_api_keys(self._pool_multiline_text)
            if parsed:
                return parsed
        return parse_api_keys(self.api_key_var.get())

    def _refresh_key_pool_badge(self):
        n = len(self._get_all_keys_list())
        self.pool_btn_text_var.set(f"🔑 批量管理 Key ({n} 个)")

    def _get_api_mode_code(self) -> str:
        lbl = self.api_mode_label.get().strip()
        for disp, code in API_MODES:
            if lbl == disp or lbl == code:
                return code
        return "auto"

    def _get_rotation_mode_code(self) -> str:
        lbl = self.rotation_label.get().strip()
        for disp, code in ROTATION_MODES:
            if lbl == disp or lbl == code:
                return code
        return "round_robin"

    def _get_model_id(self) -> str:
        lbl = self.model_label.get().strip()
        for disp, mid in SUPPORTED_MODELS:
            if lbl == disp or lbl == mid:
                return mid
        return lbl or "gemini-3.5-flash-lite"

    def _get_selected_lang_code(self) -> str:
        return resolve_combo_to_code(self.lang_var.get(), self.languages)

    def _get_wc_lang_code(self) -> str:
        return resolve_combo_to_code(self.wc_lang_var.get(), self.languages)

    def _get_fel_lang_code(self) -> str:
        return resolve_combo_to_code(self.fel_lang_var.get(), self.languages)

    def _get_fel_quote_mode_code(self) -> str:
        lbl = self.fel_quote_label_var.get().strip()
        for disp, code in QUOTE_MODES:
            if lbl == disp or lbl == code:
                return code
        return "auto"

    def _code_to_combo_display(self, code: str) -> str:
        c = (code or "").strip().upper()
        if not c or c == "AUTO":
            return AUTO_COMBO_LABEL
        for e in self.languages:
            if e["code"] == c:
                return format_combo_label(e)
        return c

    def _load_settings(self):
        migrated_plaintext = False
        loaded_key_raw = ""
        if SETTINGS_FILE.exists():
            try:
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                self.remember_key_var.set(bool(cfg.get("remember_key", True)))
                enc_token = str(cfg.get("api_key_dpapi", "")).strip()
                if enc_token:
                    loaded_key_raw = dpapi_decrypt_str(enc_token)
                legacy_plain = str(cfg.get("api_key", "")).strip()
                if legacy_plain:
                    if not loaded_key_raw:
                        loaded_key_raw = legacy_plain
                    migrated_plaintext = True

                saved_mode = cfg.get("api_mode_label", "")
                if saved_mode in {m[0] for m in API_MODES}:
                    self.api_mode_label.set(saved_mode)
                saved_rot = cfg.get("rotation_label", "")
                if saved_rot in {m[0] for m in ROTATION_MODES}:
                    self.rotation_label.set(saved_rot)
                saved_model = cfg.get("model_label", "")
                if saved_model:
                    self.model_label.set(saved_model)
                self.library_dir_var.set(cfg.get("library_dir", ""))
                self.output_dir_var.set(cfg.get("output_dir", ""))
                saved_lang = cfg.get("lang", "AUTO")
                self.lang_var.set(self._code_to_combo_display(saved_lang))
                saved_wc_lang = cfg.get("wc_lang", saved_lang)
                self.wc_lang_var.set(self._code_to_combo_display(saved_wc_lang))
                self.dual_outputs_var.set(bool(cfg.get("dual_outputs", True)))
                self.fel_output_dir_var.set(cfg.get("fel_output_dir", ""))
                saved_fel_lang = cfg.get("fel_lang", "AUTO")
                self.fel_lang_var.set(self._code_to_combo_display(saved_fel_lang))
                saved_fel_mode = cfg.get("fel_output_mode", "single_only")
                if saved_fel_mode in {m[1] for m in OUTPUT_MODES}:
                    self.fel_output_mode_var.set(saved_fel_mode)
                saved_fel_quote = cfg.get("fel_quote_label", "")
                if saved_fel_quote in {q[0] for q in QUOTE_MODES}:
                    self.fel_quote_label_var.set(saved_fel_quote)
                self.fel_enable_ai_var.set(bool(cfg.get("fel_enable_ai", True)))
            except Exception:
                pass

        if not loaded_key_raw:
            loaded_key_raw = load_api_key_from_env(_app_base_dir())
        self._set_pool_multiline_and_sync_entry(loaded_key_raw)

        if migrated_plaintext:
            self._save_settings()

    def _save_settings(self):
        raw_to_store = self._pool_multiline_text.strip() or "\n".join(self._get_all_keys_list())
        enc_token = ""
        if self.remember_key_var.get() and raw_to_store:
            try:
                enc_token = dpapi_encrypt_str(raw_to_store)
            except Exception:
                enc_token = ""

        cfg = {
            "remember_key": self.remember_key_var.get(),
            "api_key_dpapi": enc_token,
            "api_mode_label": self.api_mode_label.get(),
            "rotation_label": self.rotation_label.get(),
            "model_label": self.model_label.get(),
            "library_dir": self.library_dir_var.get().strip(),
            "output_dir": self.output_dir_var.get().strip(),
            "lang": self._get_selected_lang_code(),
            "wc_lang": self._get_wc_lang_code(),
            "dual_outputs": self.dual_outputs_var.get(),
            "fel_output_dir": self.fel_output_dir_var.get().strip(),
            "fel_lang": self._get_fel_lang_code(),
            "fel_output_mode": self.fel_output_mode_var.get(),
            "fel_quote_label": self.fel_quote_label_var.get(),
            "fel_enable_ai": self.fel_enable_ai_var.get(),
        }
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _build_ui(self):
        outer = tk.Frame(self.root)
        outer.pack(fill=tk.BOTH, expand=True)

        # Left Navigation Sidebar ("标签选项放到左边")
        self.sidebar = tk.Frame(
            outer,
            bg="#ffffff",
            width=204,
            highlightbackground="#cbd5e1",
            highlightthickness=1,
        )
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar.pack_propagate(False)

        brand_box = tk.Frame(self.sidebar, bg="#ffffff", padx=14, pady=14)
        brand_box.pack(fill=tk.X)
        tk.Label(
            brand_box,
            text="零散改点整理工具",
            font=("Microsoft YaHei UI", 11, "bold"),
            fg="#0f172a",
            bg="#ffffff",
            anchor=tk.W,
        ).pack(fill=tk.X)
        tk.Label(
            brand_box,
            text="v1.5.1 · 支持拖拽文件/文件夹",
            font=("Microsoft YaHei UI", 8),
            fg="#64748b",
            bg="#ffffff",
            anchor=tk.W,
        ).pack(fill=tk.X, pady=(3, 0))

        tk.Frame(self.sidebar, bg="#e2e8f0", height=1).pack(fill=tk.X, padx=12, pady=(0, 10))

        tk.Label(
            self.sidebar,
            text="功能标签页",
            font=("Microsoft YaHei UI", 8, "bold"),
            fg="#64748b",
            bg="#ffffff",
            anchor=tk.W,
            padx=14,
        ).pack(fill=tk.X, pady=(0, 6))

        self.nav_btn_main = self._create_sidebar_tab(
            key="main",
            title="📝 零散改点整理",
            subtitle="初筛 · 标色 · 三版定稿",
        )
        self.nav_btn_wc = self._create_sidebar_tab(
            key="wc",
            title="📊 计算文档字数",
            subtitle="算英文词数 · 改文件名",
        )
        self.nav_btn_fel = self._create_sidebar_tab(
            key="fel",
            title="📚 弟兄交通整理",
            subtitle="拆框清痕 · 蓝粗还原 · 可选合并",
        )

        # Bottom Drag & Drop quick tip inside left sidebar
        dnd_tip_card = tk.Frame(
            self.sidebar,
            bg="#f8fafc",
            highlightbackground="#cbd5e1",
            highlightthickness=1,
            padx=10,
            pady=10,
        )
        dnd_tip_card.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=12)
        tk.Label(
            dnd_tip_card,
            text="📥 快捷拖拽说明",
            font=("Microsoft YaHei UI", 9, "bold"),
            fg="#1d4ed8",
            bg="#f8fafc",
            anchor=tk.W,
        ).pack(fill=tk.X)
        tk.Label(
            dnd_tip_card,
            text=(
                "可直接从桌面或资源管理器把 .docx 文档或文件夹拖入：\n"
                "• 右侧各路径输入框\n"
                "• 右侧空白处/预览列表\n"
                "• 左侧对应功能标签"
            ),
            font=("Microsoft YaHei UI", 8),
            fg="#475569",
            bg="#f8fafc",
            justify=tk.LEFT,
            anchor=tk.W,
            wraplength=172,
        ).pack(fill=tk.X, pady=(4, 0))

        # Right Content Area
        self.content_area = ttk.Frame(outer, padding=(10, 8, 10, 8))
        self.content_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.content_area.rowconfigure(0, weight=1)
        self.content_area.columnconfigure(0, weight=1)

        self.tab_main = ttk.Frame(self.content_area)
        self.tab_wc = ttk.Frame(self.content_area)
        self.tab_fel = ttk.Frame(self.content_area)

        self.tab_main.grid(row=0, column=0, sticky=tk.NSEW)
        self.tab_wc.grid(row=0, column=0, sticky=tk.NSEW)
        self.tab_fel.grid(row=0, column=0, sticky=tk.NSEW)

        self._build_tab1_pipeline(self.tab_main)
        self._build_tab2_word_counter(self.tab_wc)
        self._build_tab3_fellowship(self.tab_fel)

        self._switch_tab("main")

    def _create_sidebar_tab(self, key: str, title: str, subtitle: str) -> tk.Frame:
        card = tk.Frame(
            self.sidebar,
            bg="#ffffff",
            highlightbackground="#e2e8f0",
            highlightthickness=1,
            cursor="hand2",
        )
        card.pack(fill=tk.X, padx=10, pady=4)

        bar = tk.Frame(card, bg="#ffffff", width=4)
        bar.pack(side=tk.LEFT, fill=tk.Y)

        inner = tk.Frame(card, bg="#ffffff", padx=10, pady=9, cursor="hand2")
        inner.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        lbl_title = tk.Label(
            inner,
            text=title,
            font=("Microsoft YaHei UI", 10, "bold"),
            fg="#1e293b",
            bg="#ffffff",
            anchor=tk.W,
            cursor="hand2",
        )
        lbl_title.pack(fill=tk.X)

        lbl_sub = tk.Label(
            inner,
            text=subtitle,
            font=("Microsoft YaHei UI", 8),
            fg="#64748b",
            bg="#ffffff",
            anchor=tk.W,
            cursor="hand2",
        )
        lbl_sub.pack(fill=tk.X, pady=(2, 0))

        self._nav_items[key] = {
            "card": card,
            "bar": bar,
            "inner": inner,
            "title": lbl_title,
            "sub": lbl_sub,
        }

        def _on_click(_e=None):
            self._switch_tab(key)

        def _on_enter(_e=None):
            if self.active_tab_key != key:
                for w in (card, inner, lbl_title, lbl_sub):
                    w.configure(bg="#f8fafc")
                bar.configure(bg="#cbd5e1")

        def _on_leave(_e=None):
            if self.active_tab_key != key:
                for w in (card, inner, lbl_title, lbl_sub):
                    w.configure(bg="#ffffff")
                bar.configure(bg="#ffffff")

        for w in (card, bar, inner, lbl_title, lbl_sub):
            w.bind("<Button-1>", _on_click)
        card.bind("<Enter>", _on_enter)
        card.bind("<Leave>", _on_leave)
        return card

    def _switch_tab(self, key: str):
        self.active_tab_key = key
        tab_map = {
            "main": self.tab_main,
            "wc": self.tab_wc,
            "fel": self.tab_fel,
        }
        for k, frame in tab_map.items():
            if k == key:
                frame.grid(row=0, column=0, sticky=tk.NSEW)
            else:
                frame.grid_remove()

        for k, parts in self._nav_items.items():
            is_act = k == key
            bg = "#eff6ff" if is_act else "#ffffff"
            border = "#93c5fd" if is_act else "#e2e8f0"
            bar_c = "#2563eb" if is_act else "#ffffff"
            fg_t = "#1d4ed8" if is_act else "#1e293b"
            fg_s = "#2563eb" if is_act else "#64748b"

            parts["card"].configure(bg=bg, highlightbackground=border)
            parts["bar"].configure(bg=bar_c)
            parts["inner"].configure(bg=bg)
            parts["title"].configure(bg=bg, fg=fg_t)
            parts["sub"].configure(bg=bg, fg=fg_s)

    def _build_tab1_pipeline(self, main: ttk.Frame):
        # 1. API & Model Frame
        api_frame = ttk.LabelFrame(main, text=" 1. API 密钥池与模型设置（支持无限个 Key 轮换 + 超额/报错秒切） ", padding=10)
        api_frame.pack(fill=tk.X, pady=(0, 8))
        api_frame.columnconfigure(1, weight=1)

        ttk.Label(api_frame, text="API 类型：").grid(row=0, column=0, sticky=tk.W, pady=4)
        mode_cb = ttk.Combobox(
            api_frame,
            textvariable=self.api_mode_label,
            values=[m[0] for m in API_MODES],
            state="readonly",
            width=36,
        )
        mode_cb.grid(row=0, column=1, sticky=tk.W, pady=4, padx=(4, 12))

        ttk.Label(api_frame, text="选择模型：").grid(row=0, column=2, sticky=tk.W, pady=4)
        model_cb = ttk.Combobox(
            api_frame,
            textvariable=self.model_label,
            values=[m[0] for m in SUPPORTED_MODELS],
            state="readonly",
            width=32,
        )
        model_cb.grid(row=0, column=3, columnspan=2, sticky=tk.EW, pady=4, padx=(4, 0))

        ttk.Label(api_frame, text="API Key 池：").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.key_entry = ttk.Entry(api_frame, textvariable=self.api_key_var, show="•")
        self.key_entry.grid(row=1, column=1, columnspan=2, sticky=tk.EW, pady=4, padx=(4, 8))
        self.key_entry.bind("<<Paste>>", self._on_paste_into_key_entry)

        self.pool_btn = ttk.Button(
            api_frame,
            textvariable=self.pool_btn_text_var,
            command=self._open_key_pool_dialog,
        )
        self.pool_btn.grid(row=1, column=3, sticky=tk.EW, padx=(0, 6))

        self.test_btn = ttk.Button(api_frame, text="测试连接", width=10, command=self._on_test_api)
        self.test_btn.grid(row=1, column=4, sticky=tk.E)

        ttk.Label(api_frame, text="轮换策略：").grid(row=2, column=0, sticky=tk.W, pady=(4, 0))
        rot_row = ttk.Frame(api_frame)
        rot_row.grid(row=2, column=1, columnspan=4, sticky=tk.EW, pady=(4, 0), padx=(4, 0))

        ttk.Combobox(
            rot_row,
            textvariable=self.rotation_label,
            values=[m[0] for m in ROTATION_MODES],
            state="readonly",
            width=38,
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Checkbutton(
            rot_row, text="显示明文", variable=self.show_key_var, command=self._toggle_key_visibility
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Checkbutton(
            rot_row, text="本机加密记住", variable=self.remember_key_var, command=self._save_settings
        ).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(
            rot_row,
            text="提示：可直接粘贴多行/逗号分隔的多个 Key，或点右侧「🔑 批量管理 Key」分行维护",
            foreground="#64748b",
        ).pack(side=tk.LEFT)

        # 2. Paths & Multi-Language Options Frame
        path_frame = ttk.LabelFrame(
            main,
            text=" 2. 文档、总库路径与多语种输出设置（📥 支持直接将文件/文件夹拖拽到对应输入框或页面任意空白处） ",
            padding=10,
        )
        path_frame.pack(fill=tk.X, pady=(0, 8))
        path_frame.columnconfigure(1, weight=1)

        self.lbl_input_docx = ttk.Label(path_frame, text="待整理文档 (.docx)：")
        self.lbl_input_docx.grid(row=0, column=0, sticky=tk.W, pady=4)
        self.ent_input_docx = ttk.Entry(path_frame, textvariable=self.input_docx_var)
        self.ent_input_docx.grid(row=0, column=1, sticky=tk.EW, pady=4, padx=4)
        ttk.Button(path_frame, text="浏览...", width=10, command=self._browse_input).grid(row=0, column=2, padx=2)

        self.lbl_library_dir = ttk.Label(path_frame, text="对应语种总库目录：")
        self.lbl_library_dir.grid(row=1, column=0, sticky=tk.W, pady=4)
        self.ent_library_dir = ttk.Entry(path_frame, textvariable=self.library_dir_var)
        self.ent_library_dir.grid(row=1, column=1, sticky=tk.EW, pady=4, padx=4)
        ttk.Button(path_frame, text="浏览...", width=10, command=self._browse_library).grid(row=1, column=2, padx=2)

        self.lbl_output_dir = ttk.Label(path_frame, text="输出保存目录：")
        self.lbl_output_dir.grid(row=2, column=0, sticky=tk.W, pady=4)
        self.ent_output_dir = ttk.Entry(path_frame, textvariable=self.output_dir_var)
        self.ent_output_dir.grid(row=2, column=1, sticky=tk.EW, pady=4, padx=4)
        out_btns = ttk.Frame(path_frame)
        out_btns.grid(row=2, column=2, columnspan=2, sticky=tk.W)
        ttk.Button(out_btns, text="浏览...", width=10, command=self._browse_output).pack(side=tk.LEFT, padx=2)
        ttk.Button(out_btns, text="打开目录", width=10, command=self._open_output_dir).pack(side=tk.LEFT, padx=2)

        self.lbl_draft_docx = ttk.Label(path_frame, text="初筛确认稿 (第二步用)：")
        self.lbl_draft_docx.grid(row=3, column=0, sticky=tk.W, pady=4)
        self.ent_draft_docx = ttk.Entry(path_frame, textvariable=self.draft_docx_var)
        self.ent_draft_docx.grid(row=3, column=1, sticky=tk.EW, pady=4, padx=4)
        draft_btns = ttk.Frame(path_frame)
        draft_btns.grid(row=3, column=2, columnspan=2, sticky=tk.W)
        ttk.Button(draft_btns, text="浏览...", width=10, command=self._browse_draft).pack(side=tk.LEFT, padx=2)
        ttk.Button(draft_btns, text="打开初筛稿", width=10, command=self._open_draft_file).pack(side=tk.LEFT, padx=2)

        opt_row1 = ttk.Frame(path_frame)
        opt_row1.grid(row=4, column=0, columnspan=4, sticky=tk.EW, pady=(6, 2))

        ttk.Label(opt_row1, text="目标语种：").pack(side=tk.LEFT)
        self.lang_cb = ttk.Combobox(
            opt_row1,
            textvariable=self.lang_var,
            values=get_combo_values(self.languages),
            width=28,
        )
        self.lang_cb.pack(side=tk.LEFT, padx=(4, 6))

        ttk.Button(
            opt_row1,
            text="➕ 管理 / 新增语种...",
            command=self._open_lang_manager,
        ).pack(side=tk.LEFT, padx=(0, 18))

        ttk.Label(opt_row1, text="自定义原文档名(选填，留空自动取原文档名)：").pack(side=tk.LEFT)
        ttk.Entry(opt_row1, textvariable=self.doc_name_var, width=22).pack(side=tk.LEFT, padx=(4, 0))

        opt_row2 = ttk.Frame(path_frame)
        opt_row2.grid(row=5, column=0, columnspan=4, sticky=tk.EW, pady=(4, 0))
        ttk.Checkbutton(
            opt_row2,
            text="同时输出三份定稿：【发译者版(仅改点)】+【完整留底版】+【无改点内容版】",
            variable=self.dual_outputs_var,
        ).pack(side=tk.LEFT)
        ttk.Label(
            opt_row2,
            text="  命名格式：GR_God’s Words_(UP X words, RV Y words)_YYYYMMDD.docx",
            foreground="#64748b",
        ).pack(side=tk.LEFT)

        # 3. Action Buttons & Determinate Progress Frame
        act_frame = ttk.LabelFrame(main, text=" 3. 执行操作（支持两步走复核 / 一键直出） ", padding=10)
        act_frame.pack(fill=tk.X, pady=(0, 8))

        btn_row = ttk.Frame(act_frame)
        btn_row.pack(fill=tk.X)
        btn_row.columnconfigure(0, weight=1)
        btn_row.columnconfigure(1, weight=1)
        btn_row.columnconfigure(2, weight=1)

        self.btn_step1 = ttk.Button(
            btn_row,
            text="① 第一步：生成初筛稿 (供人工复核)",
            style="Action.TButton",
            command=self._on_run_step1,
        )
        self.btn_step1.grid(row=0, column=0, sticky=tk.EW, padx=(0, 6))

        self.btn_step2 = ttk.Button(
            btn_row,
            text="② 第二步：由初筛稿生成定稿",
            style="Action.TButton",
            command=self._on_run_step2,
        )
        self.btn_step2.grid(row=0, column=1, sticky=tk.EW, padx=6)

        self.btn_all = ttk.Button(
            btn_row,
            text="⚡ 一键直出全部定稿 (初筛 + 三份定稿)",
            style="Primary.TButton",
            command=self._on_run_all,
        )
        self.btn_all.grid(row=0, column=2, sticky=tk.EW, padx=(6, 0))

        prog_info = ttk.Frame(act_frame)
        prog_info.pack(fill=tk.X, pady=(10, 3))

        self.stage_lbl = ttk.Label(
            prog_info,
            textvariable=self.stage_var,
            font=("Microsoft YaHei UI", 9, "bold"),
            foreground="#1d4ed8",
        )
        self.stage_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.pct_lbl = ttk.Label(
            prog_info,
            textvariable=self.pct_timer_var,
            font=("Consolas", 9, "bold"),
            foreground="#0f172a",
        )
        self.pct_lbl.pack(side=tk.RIGHT, padx=(8, 0))

        self.progress = ttk.Progressbar(
            act_frame,
            mode="determinate",
            maximum=100.0,
            variable=self.progress_var,
        )
        self.progress.pack(fill=tk.X, pady=(0, 2))

        # 4. Log Frame
        log_frame = ttk.LabelFrame(main, text=" 4. 实时运行日志 ", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)

        log_top = ttk.Frame(log_frame)
        log_top.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(
            log_top,
            text="所有步骤、多 Key 自动轮换/切换提示、语种识别及保存结果均在此实时滚动显示：",
            foreground="#64748b",
        ).pack(side=tk.LEFT)
        ttk.Button(log_top, text="清空日志", width=9, command=self._clear_log).pack(side=tk.RIGHT)

        log_body = ttk.Frame(log_frame)
        log_body.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(
            log_body,
            height=9,
            wrap=tk.WORD,
            font=("Microsoft YaHei UI", 9),
            bg="#f8fafc",
            fg="#0f172a",
            relief=tk.SOLID,
            borderwidth=1,
            padx=8,
            pady=6,
        )
        self.log_text.tag_configure("header", foreground="#1d4ed8", font=("Microsoft YaHei UI", 9, "bold"))
        self.log_text.tag_configure("ok", foreground="#15803d", font=("Microsoft YaHei UI", 9, "bold"))
        self.log_text.tag_configure("err", foreground="#b91c1c", font=("Microsoft YaHei UI", 9, "bold"))
        self.log_text.tag_configure("retry", foreground="#b45309")
        self.log_text.tag_configure("dim", foreground="#64748b")

        scroll = ttk.Scrollbar(log_body, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        key_cnt = len(self._get_all_keys_list())
        self._append_log(
            f"欢迎使用零散改点自动整理工具 v1.5.1！已载入 {key_cnt} 个 API Key、{len(self.languages)} 种语言；支持零散改点、字数统计与弟兄交通整理。"
        )

    def _build_tab2_word_counter(self, parent: ttk.Frame):
        # 1. Rules banner
        banner = ttk.LabelFrame(parent, text=" 统计与文件重命名规则说明（本地毫秒级精准解析，无需消耗 API） ", padding=10)
        banner.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(
            banner,
            text=(
                "• 统计口径：仅统计表格【第 2 列（英文列）】修订后正文的英文单词数（标点符号不算字数，算单词数而非字母数；已删除修订自动排除）。\n"
                "• 颜色定义：UP = 黄色高亮 (Yellow)  |  RV = 青色高亮 (Cyan)\n"
                "• 命名格式：语言代号_原文档名字_(UP X words, RV Y words)_日期.docx  （例：GR_God’s Words_(UP 128 words, RV 64 words)_20261115.docx）\n"
                "• 智能更新：若原文档名中已有字数统计会自动替换为最新统计结果；若没有则自动按规范新增！"
            ),
            foreground="#334155",
        ).pack(fill=tk.X)

        # 2. File picker & naming parameters
        cfg_frame = ttk.LabelFrame(parent, text=" 1. 拖拽 / 选择文档与命名参数 ", padding=10)
        cfg_frame.pack(fill=tk.X, pady=(0, 8))

        drop_zone = tk.Frame(
            cfg_frame,
            bg="#eff6ff",
            highlightbackground="#93c5fd",
            highlightthickness=1,
            padx=12,
            pady=7,
            cursor="hand2",
        )
        drop_zone.pack(fill=tk.X, pady=(0, 8))
        dz_lbl = tk.Label(
            drop_zone,
            text="📥 直接把一个或多个 .docx 文档、文件夹拖拽到此窗口任意位置即可自动载入（点击此处也可选择文件）",
            font=("Microsoft YaHei UI", 9, "bold"),
            fg="#1d4ed8",
            bg="#eff6ff",
            cursor="hand2",
        )
        dz_lbl.pack(fill=tk.X)
        drop_zone.bind("<Button-1>", lambda _e: self._wc_add_files())
        dz_lbl.bind("<Button-1>", lambda _e: self._wc_add_files())

        file_btns = ttk.Frame(cfg_frame)
        file_btns.pack(fill=tk.X, pady=(0, 8))

        ttk.Button(
            file_btns,
            text="📄 选择 Word 文档 (支持单选/多选)...",
            style="Action.TButton",
            command=self._wc_add_files,
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(
            file_btns,
            text="📁 选择文件夹批量载入...",
            style="Action.TButton",
            command=self._wc_add_folder,
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(file_btns, text="🗑️ 移除选中行", command=self._wc_remove_selected).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(file_btns, text="🧹 清空列表", command=self._wc_clear_all).pack(side=tk.LEFT)

        row_lang = ttk.Frame(cfg_frame)
        row_lang.pack(fill=tk.X, pady=4)

        ttk.Label(row_lang, text="语言代号：").pack(side=tk.LEFT)
        self.wc_lang_cb = ttk.Combobox(
            row_lang,
            textvariable=self.wc_lang_var,
            values=get_combo_values(self.languages),
            width=28,
        )
        self.wc_lang_cb.pack(side=tk.LEFT, padx=(4, 6))
        self.wc_lang_cb.bind("<<ComboboxSelected>>", lambda _e: self._wc_refresh_previews())

        ttk.Button(row_lang, text="➕ 管理语种...", command=self._open_lang_manager).pack(side=tk.LEFT, padx=(0, 14))

        ttk.Checkbutton(
            row_lang,
            text="原文件名已有语言代号（如 GR_）时优先保留",
            variable=self.wc_keep_existing_lang_var,
            command=self._wc_refresh_previews,
        ).pack(side=tk.LEFT)

        row_date = ttk.Frame(cfg_frame)
        row_date.pack(fill=tk.X, pady=4)

        ttk.Label(row_date, text="日期 (YYYYMMDD)：").pack(side=tk.LEFT)
        date_ent = ttk.Entry(row_date, textvariable=self.wc_date_var, width=12)
        date_ent.pack(side=tk.LEFT, padx=(4, 14))
        date_ent.bind("<KeyRelease>", lambda _e: self._wc_refresh_previews())

        ttk.Checkbutton(
            row_date,
            text="原文件名已有 8 位日期时保留原日期",
            variable=self.wc_keep_existing_date_var,
            command=self._wc_refresh_previews,
        ).pack(side=tk.LEFT, padx=(0, 16))

        ttk.Checkbutton(
            row_date,
            text="自动去除开头的“初筛_”前缀",
            variable=self.wc_strip_draft_var,
            command=self._wc_refresh_previews,
        ).pack(side=tk.LEFT)

        row_mode = ttk.Frame(cfg_frame)
        row_mode.pack(fill=tk.X, pady=(4, 0))

        ttk.Label(row_mode, text="保存方式：").pack(side=tk.LEFT)
        ttk.Radiobutton(
            row_mode,
            text="直接就地更新原文档文件名（推荐）",
            value="inplace",
            variable=self.wc_save_mode_var,
        ).pack(side=tk.LEFT, padx=(4, 16))
        self.wc_copy_radio = ttk.Radiobutton(
            row_mode,
            text="另存副本到指定文件夹：",
            value="copy",
            variable=self.wc_save_mode_var,
        )
        self.wc_copy_radio.pack(side=tk.LEFT)
        self.wc_copy_dir_entry = ttk.Entry(row_mode, textvariable=self.wc_copy_dir_var, width=30)
        self.wc_copy_dir_entry.pack(side=tk.LEFT, padx=4, fill=tk.X, expand=True)
        ttk.Button(row_mode, text="浏览...", width=8, command=self._wc_browse_copy_dir).pack(side=tk.LEFT)

        # 3. Preview & Results Treeview
        tbl_box = ttk.LabelFrame(
            parent,
            text=" 2. 第 2 列英文字数统计与新文件名预览（📥 支持直接将文件/文件夹拖入下表；双击行可打开目录） ",
            padding=8,
        )
        tbl_box.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        cols = ("old_name", "lang", "up_words", "rv_words", "new_name", "status")
        self.wc_tree = ttk.Treeview(tbl_box, columns=cols, show="headings", selectmode="extended", height=12)
        self.wc_tree.heading("old_name", text="原文件名")
        self.wc_tree.heading("lang", text="语言代号")
        self.wc_tree.heading("up_words", text="UP 单词数 (黄色)")
        self.wc_tree.heading("rv_words", text="RV 单词数 (青色)")
        self.wc_tree.heading("new_name", text="计算后目标文件名")
        self.wc_tree.heading("status", text="状态")

        self.wc_tree.column("old_name", width=220, anchor=tk.W)
        self.wc_tree.column("lang", width=70, anchor=tk.CENTER)
        self.wc_tree.column("up_words", width=110, anchor=tk.CENTER)
        self.wc_tree.column("rv_words", width=110, anchor=tk.CENTER)
        self.wc_tree.column("new_name", width=310, anchor=tk.W)
        self.wc_tree.column("status", width=115, anchor=tk.CENTER)

        vsb = ttk.Scrollbar(tbl_box, orient=tk.VERTICAL, command=self.wc_tree.yview)
        hsb = ttk.Scrollbar(tbl_box, orient=tk.HORIZONTAL, command=self.wc_tree.xview)
        self.wc_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.wc_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.wc_tree.bind("<Double-1>", self._wc_open_selected_folder)

        # 4. Bottom Action Bar
        bot = ttk.Frame(parent)
        bot.pack(fill=tk.X)

        ttk.Label(
            bot,
            textvariable=self.wc_summary_var,
            font=("Microsoft YaHei UI", 9, "bold"),
            foreground="#1d4ed8",
        ).pack(side=tk.LEFT)

        ttk.Button(
            bot,
            text="⚡ 计算字数并立即更新文件名",
            style="Primary.TButton",
            command=self._wc_execute_rename,
        ).pack(side=tk.RIGHT)
        ttk.Button(
            bot,
            text="🔄 重新计算并刷新预览",
            style="Action.TButton",
            command=self._wc_refresh_previews,
        ).pack(side=tk.RIGHT, padx=(0, 8))

    def _build_tab3_fellowship(self, parent: ttk.Frame):
        # 1. File / Folder Drop & Parameters Frame
        cfg_frame = ttk.LabelFrame(
            parent,
            text=" 1. 弟兄交通稿件载入与排版输出设置（拆内容控件 · 关修订清批注 · 第3列蓝粗还原 · 固定值20磅 · 均衡列宽） ",
            padding=10,
        )
        cfg_frame.pack(fill=tk.X, pady=(0, 7))

        drop_zone = tk.Frame(
            cfg_frame,
            bg="#eff6ff",
            highlightbackground="#93c5fd",
            highlightthickness=1,
            padx=12,
            pady=6,
            cursor="hand2",
        )
        drop_zone.pack(fill=tk.X, pady=(0, 7))
        dz_lbl = tk.Label(
            drop_zone,
            text="📥 直接拖入整个分篇文件夹或多个 .docx 稿件即可自动按序号排序载入（点击此处可选择文件夹）",
            font=("Microsoft YaHei UI", 9, "bold"),
            fg="#1d4ed8",
            bg="#eff6ff",
            cursor="hand2",
        )
        dz_lbl.pack(fill=tk.X)
        drop_zone.bind("<Button-1>", lambda _e: self._fel_add_folder())
        dz_lbl.bind("<Button-1>", lambda _e: self._fel_add_folder())

        btn_bar = ttk.Frame(cfg_frame)
        btn_bar.pack(fill=tk.X, pady=(0, 7))

        ttk.Button(
            btn_bar,
            text="📁 选择文件夹批量载入...",
            style="Action.TButton",
            command=self._fel_add_folder,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            btn_bar,
            text="📄 选择多个 Word 文档...",
            style="Action.TButton",
            command=self._fel_add_files,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_bar, text="🗑️ 移除选中", command=self._fel_remove_selected).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_bar, text="🧹 清空列表", command=self._fel_clear_all).pack(side=tk.LEFT)

        self.fel_pool_btn = ttk.Button(
            btn_bar,
            textvariable=self.pool_btn_text_var,
            command=self._open_key_pool_dialog,
        )
        self.fel_pool_btn.pack(side=tk.RIGHT)

        out_row = ttk.Frame(cfg_frame)
        out_row.pack(fill=tk.X, pady=3)
        self.lbl_fel_output_dir = ttk.Label(out_row, text="输出保存目录：")
        self.lbl_fel_output_dir.pack(side=tk.LEFT)
        self.ent_fel_output_dir = ttk.Entry(out_row, textvariable=self.fel_output_dir_var)
        self.ent_fel_output_dir.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        ttk.Button(out_row, text="浏览...", width=9, command=self._fel_browse_output).pack(side=tk.LEFT, padx=2)
        ttk.Button(out_row, text="打开目录", width=9, command=self._fel_open_output_dir).pack(side=tk.LEFT, padx=(2, 0))

        opt_row = ttk.Frame(cfg_frame)
        opt_row.pack(fill=tk.X, pady=3)

        ttk.Label(opt_row, text="目标语种：").pack(side=tk.LEFT)
        self.fel_lang_cb = ttk.Combobox(
            opt_row,
            textvariable=self.fel_lang_var,
            values=get_combo_values(self.languages),
            width=22,
        )
        self.fel_lang_cb.pack(side=tk.LEFT, padx=(4, 4))
        self.fel_lang_cb.bind("<<ComboboxSelected>>", self._fel_on_lang_changed)

        ttk.Button(opt_row, text="➕ 管理语种", command=self._open_lang_manager).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(opt_row, text="三语前缀(可改)：").pack(side=tk.LEFT)
        fel_pfx_ent = ttk.Entry(opt_row, textvariable=self.fel_tri_prefix_var, width=10)
        fel_pfx_ent.pack(side=tk.LEFT, padx=(3, 10))
        fel_pfx_ent.bind("<KeyRelease>", lambda _e: self._fel_refresh_previews())

        ttk.Label(opt_row, text="日期：").pack(side=tk.LEFT)
        fel_dt_ent = ttk.Entry(opt_row, textvariable=self.fel_date_var, width=10)
        fel_dt_ent.pack(side=tk.LEFT, padx=(3, 10))
        fel_dt_ent.bind("<KeyRelease>", lambda _e: self._fel_refresh_previews())

        ttk.Label(opt_row, text="引号规则：").pack(side=tk.LEFT)
        ttk.Combobox(
            opt_row,
            textvariable=self.fel_quote_label_var,
            values=[q[0] for q in QUOTE_MODES],
            state="readonly",
            width=26,
        ).pack(side=tk.LEFT, padx=(3, 0), fill=tk.X, expand=True)

        mode_row = ttk.Frame(cfg_frame)
        mode_row.pack(fill=tk.X, pady=(4, 2))
        ttk.Label(mode_row, text="输出方式：", font=("Microsoft YaHei UI", 9, "bold")).pack(side=tk.LEFT)
        for disp_txt, mode_val in OUTPUT_MODES:
            ttk.Radiobutton(
                mode_row,
                text=disp_txt,
                value=mode_val,
                variable=self.fel_output_mode_var,
                command=self._fel_refresh_previews,
            ).pack(side=tk.LEFT, padx=(6, 10))

        ai_opt_row = ttk.Frame(cfg_frame)
        ai_opt_row.pack(fill=tk.X, pady=(2, 0))
        ttk.Checkbutton(
            ai_opt_row,
            text="使用 AI 自动还原第 3 列段内局部蓝粗引文（原稿已有蓝粗自动跳过；取消勾选则 100% 纯本地离线秒出）",
            variable=self.fel_enable_ai_var,
            command=self._save_settings,
        ).pack(side=tk.LEFT)

        # 2. Sorted Part List & Output Filename Preview Treeview
        tbl_box = ttk.LabelFrame(
            parent,
            text=" 2. 分篇顺序与输出定稿命名预览（Part 1 自动统一转为 -01；主标题保留四号、正文小四、引文出处五号） ",
            padding=7,
        )
        tbl_box.pack(fill=tk.BOTH, expand=True, pady=(0, 7))

        cols = ("seq", "old_name", "part_tag", "rows", "preview_name")
        self.fel_tree = ttk.Treeview(tbl_box, columns=cols, show="headings", selectmode="extended", height=7)
        self.fel_tree.heading("seq", text="排序")
        self.fel_tree.heading("old_name", text="原文档名")
        self.fel_tree.heading("part_tag", text="分篇号")
        self.fel_tree.heading("rows", text="表格行数")
        self.fel_tree.heading("preview_name", text="预计输出文件名")

        self.fel_tree.column("seq", width=52, anchor=tk.CENTER)
        self.fel_tree.column("old_name", width=265, anchor=tk.W)
        self.fel_tree.column("part_tag", width=68, anchor=tk.CENTER)
        self.fel_tree.column("rows", width=72, anchor=tk.CENTER)
        self.fel_tree.column("preview_name", width=385, anchor=tk.W)

        fel_vsb = ttk.Scrollbar(tbl_box, orient=tk.VERTICAL, command=self.fel_tree.yview)
        self.fel_tree.configure(yscrollcommand=fel_vsb.set)
        self.fel_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        fel_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.fel_tree.bind("<Double-1>", self._fel_open_selected_folder)

        # 3. Execute & Progress Frame
        act_box = ttk.LabelFrame(parent, text=" 3. 一键执行与实时日志 ", padding=8)
        act_box.pack(fill=tk.BOTH, expand=True)

        run_row = ttk.Frame(act_box)
        run_row.pack(fill=tk.X)

        ttk.Label(
            run_row,
            textvariable=self.fel_summary_var,
            font=("Microsoft YaHei UI", 9, "bold"),
            foreground="#1e293b",
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.btn_fel_run = ttk.Button(
            run_row,
            text="⚡ 开始整理弟兄交通稿件",
            style="Primary.TButton",
            command=self._on_run_fellowship,
        )
        self.btn_fel_run.pack(side=tk.RIGHT)

        p_info = ttk.Frame(act_box)
        p_info.pack(fill=tk.X, pady=(6, 2))
        self.fel_stage_lbl = ttk.Label(
            p_info,
            textvariable=self.fel_stage_var,
            font=("Microsoft YaHei UI", 9, "bold"),
            foreground="#1d4ed8",
        )
        self.fel_stage_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(
            p_info,
            textvariable=self.fel_pct_timer_var,
            font=("Consolas", 9, "bold"),
            foreground="#0f172a",
        ).pack(side=tk.RIGHT)

        ttk.Progressbar(
            act_box,
            mode="determinate",
            maximum=100.0,
            variable=self.fel_progress_var,
        ).pack(fill=tk.X, pady=(0, 6))

        log_wrap = ttk.Frame(act_box)
        log_wrap.pack(fill=tk.BOTH, expand=True)
        self.fel_log_text = tk.Text(
            log_wrap,
            height=6,
            wrap=tk.WORD,
            font=("Microsoft YaHei UI", 9),
            bg="#f8fafc",
            fg="#0f172a",
            relief=tk.SOLID,
            borderwidth=1,
            padx=8,
            pady=5,
        )
        self.fel_log_text.tag_configure("header", foreground="#1d4ed8", font=("Microsoft YaHei UI", 9, "bold"))
        self.fel_log_text.tag_configure("ok", foreground="#15803d", font=("Microsoft YaHei UI", 9, "bold"))
        self.fel_log_text.tag_configure("err", foreground="#b91c1c", font=("Microsoft YaHei UI", 9, "bold"))
        self.fel_log_text.tag_configure("retry", foreground="#b45309")
        self.fel_log_text.tag_configure("dim", foreground="#64748b")
        fel_log_sb = ttk.Scrollbar(log_wrap, orient=tk.VERTICAL, command=self.fel_log_text.yview)
        self.fel_log_text.configure(yscrollcommand=fel_log_sb.set)
        self.fel_log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        fel_log_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.fel_log_text.insert(
            tk.END,
            "提示：直接将分篇文件夹或多个 .docx 拖入本页即可自动识别语种（如 中英西/中英缅/中英腊）并按 -01..-0N 排序。\n",
            "dim",
        )

    # ---------------- Drag & Drop Integration ----------------

    def _setup_drag_and_drop(self):
        # Left sidebar tab buttons: dropping onto a tab switches to it and applies the drop
        self.dnd.register_target(self.nav_btn_main, lambda ps: (self._switch_tab("main"), self._on_drop_tab1_general(ps)))
        self.dnd.register_target(self.nav_btn_wc, lambda ps: (self._switch_tab("wc"), self._wc_ingest_paths(ps)))
        self.dnd.register_target(self.nav_btn_fel, lambda ps: (self._switch_tab("fel"), self._fel_ingest_paths(ps)))

        # Tab 1 specific input rows
        for w in (self.lbl_input_docx, self.ent_input_docx):
            self.dnd.register_target(w, self._on_drop_input_docx)
        for w in (self.lbl_library_dir, self.ent_library_dir):
            self.dnd.register_target(w, self._on_drop_library_dir)
        for w in (self.lbl_output_dir, self.ent_output_dir):
            self.dnd.register_target(w, self._on_drop_output_dir)
        for w in (self.lbl_draft_docx, self.ent_draft_docx):
            self.dnd.register_target(w, self._on_drop_draft_docx)

        # Tab 2 specific copy-folder target
        for w in (self.wc_copy_radio, self.wc_copy_dir_entry):
            self.dnd.register_target(w, self._on_drop_wc_copy_dir)

        # Tab 3 specific output-folder target
        for w in (self.lbl_fel_output_dir, self.ent_fel_output_dir):
            self.dnd.register_target(w, self._on_drop_fel_output_dir)

        self.dnd.install()

    def _on_drop_fallback(self, paths: list[Path]):
        if self.active_tab_key == "wc":
            self._wc_ingest_paths(paths)
        elif self.active_tab_key == "fel":
            self._fel_ingest_paths(paths)
        else:
            self._on_drop_tab1_general(paths)

    @staticmethod
    def _extract_docx_from_paths(paths: list[Path]) -> list[Path]:
        out: list[Path] = []
        for p in paths:
            try:
                rp = p.resolve()
                if rp.is_file() and rp.suffix.lower() == ".docx" and not rp.name.startswith("~$"):
                    out.append(rp)
                elif rp.is_dir():
                    top = sorted(x for x in rp.glob("*.docx") if not x.name.startswith("~$"))
                    if top:
                        out.extend(top)
                    else:
                        out.extend(sorted(x for x in rp.rglob("*.docx") if not x.name.startswith("~$")))
            except Exception:
                continue
        return out

    def _on_drop_input_docx(self, paths: list[Path]):
        docxs = self._extract_docx_from_paths(paths)
        if not docxs:
            messagebox.showinfo("拖拽提示", "【待整理文档】请拖入 .docx 格式的 Word 文档。")
            return
        picked = docxs[0]
        self.input_docx_var.set(str(picked))
        if not self.output_dir_var.get().strip():
            self.output_dir_var.set(str(picked.parent))
        self._save_settings()
        self._append_log(f"📥 已拖入【待整理文档】：{picked}")
        self.stage_var.set(f"📥 已载入待整理文档：{picked.name}")

    def _on_drop_library_dir(self, paths: list[Path]):
        p = paths[0].resolve()
        target_dir = p if p.is_dir() else p.parent
        if not target_dir.is_dir():
            return
        self.library_dir_var.set(str(target_dir))
        self._save_settings()
        self._append_log(f"📥 已拖入【对应语种总库目录】：{target_dir}")
        self.stage_var.set(f"📥 已载入总库目录：{target_dir.name}")

    def _on_drop_output_dir(self, paths: list[Path]):
        p = paths[0].resolve()
        target_dir = p if p.is_dir() else p.parent
        if not target_dir.is_dir():
            return
        self.output_dir_var.set(str(target_dir))
        self._save_settings()
        self._append_log(f"📥 已拖入【输出保存目录】：{target_dir}")
        self.stage_var.set(f"📥 已设定输出目录：{target_dir.name}")

    def _on_drop_draft_docx(self, paths: list[Path]):
        docxs = self._extract_docx_from_paths(paths)
        if not docxs:
            messagebox.showinfo("拖拽提示", "【初筛确认稿】请拖入 .docx 格式的 Word 文档。")
            return
        picked = docxs[0]
        self.draft_docx_var.set(str(picked))
        if not self.output_dir_var.get().strip():
            self.output_dir_var.set(str(picked.parent))
        self._save_settings()
        self._append_log(f"📥 已拖入【初筛确认稿】：{picked}")
        self.stage_var.set(f"📥 已载入初筛确认稿：{picked.name}")

    def _on_drop_tab1_general(self, paths: list[Path]):
        docx_files: list[Path] = []
        dirs: list[Path] = []
        for p in paths:
            try:
                rp = p.resolve()
                if rp.is_dir():
                    dirs.append(rp)
                elif rp.is_file() and rp.suffix.lower() == ".docx" and not rp.name.startswith("~$"):
                    docx_files.append(rp)
            except Exception:
                continue

        applied = []
        for df in docx_files:
            if df.name.startswith("初筛_"):
                self.draft_docx_var.set(str(df))
                if not self.output_dir_var.get().strip():
                    self.output_dir_var.set(str(df.parent))
                applied.append(f"初筛稿={df.name}")
                self._append_log(f"📥 已自动识别并填入【初筛确认稿】：{df}")
            else:
                self.input_docx_var.set(str(df))
                if not self.output_dir_var.get().strip():
                    self.output_dir_var.set(str(df.parent))
                applied.append(f"待整理文档={df.name}")
                self._append_log(f"📥 已自动识别并填入【待整理文档】：{df}")

        lib_assigned = False
        for d in dirs:
            has_corpus_docx = any(
                not f.name.startswith("~$") for f in d.rglob("*.docx")
            )
            if has_corpus_docx and not lib_assigned:
                self.library_dir_var.set(str(d))
                lib_assigned = True
                applied.append(f"总库目录={d.name}")
                self._append_log(f"📥 已自动识别并填入【对应语种总库目录】：{d}")
            else:
                self.output_dir_var.set(str(d))
                applied.append(f"输出目录={d.name}")
                self._append_log(f"📥 已填入【输出保存目录】：{d}")

        if not applied:
            messagebox.showinfo("拖拽提示", "请拖入 .docx 文档或文件夹。")
            return

        self._save_settings()
        self.stage_var.set("📥 拖拽载入成功：" + "  |  ".join(applied))

    def _on_drop_wc_copy_dir(self, paths: list[Path]):
        p = paths[0].resolve()
        target_dir = p if p.is_dir() else p.parent
        if target_dir.is_dir():
            self.wc_copy_dir_var.set(str(target_dir))
            self.wc_save_mode_var.set("copy")

    def _on_drop_fel_output_dir(self, paths: list[Path]):
        p = paths[0].resolve()
        target_dir = p if p.is_dir() else p.parent
        if target_dir.is_dir():
            self.fel_output_dir_var.set(str(target_dir))
            self._save_settings()
            self._append_log(f"📥 已设定【弟兄交通输出目录】：{target_dir}")
            self.fel_stage_var.set(f"📥 已设定输出目录：{target_dir.name}")

    # ---------------- Tab 3 Brother Fellowship Manuscript Methods ----------------

    def _fel_ingest_paths(self, paths: list[Path]):
        found_docx = self._extract_docx_from_paths(paths)
        if not found_docx:
            messagebox.showinfo("提示", "拖入或选择的项目中未找到任何有效的 .docx 文档。")
            return
        existing_norm = {os.path.normcase(str(Path(x).resolve())) for x in self.fel_files}
        for rp in found_docx:
            key = os.path.normcase(str(rp))
            if key not in existing_norm:
                existing_norm.add(key)
                self.fel_files.append(str(rp))

        sorted_ps = natural_sort_docx_paths([Path(x) for x in self.fel_files])
        self.fel_files = [str(p) for p in sorted_ps]

        if not self.fel_output_dir_var.get().strip() and sorted_ps:
            self.fel_output_dir_var.set(str(sorted_ps[0].parent))

        # Auto-sync trilingual prefix if empty or in AUTO mode
        self._fel_auto_sync_prefix()
        self._fel_refresh_previews()
        self._append_log(f"📥 【弟兄交通整理】已载入并按序号排序共 {len(self.fel_files)} 份文档。")
        self.fel_stage_var.set(f"📥 已就绪：共载入 {len(self.fel_files)} 份分篇文档")

    def _fel_auto_sync_prefix(self, force: bool = False):
        if not force and self.fel_tri_prefix_var.get().strip():
            return
        sel_code = self._get_fel_lang_code()
        if sel_code == "AUTO" and self.fel_files:
            p0 = Path(self.fel_files[0])
            hint_blob = " ".join([p0.parent.name] + [Path(x).name for x in self.fel_files[:3]])
            det = detect_lang_code_from_hint(hint_blob, self.languages)
            if det:
                sel_code = det
        if sel_code and sel_code != "AUTO":
            self.fel_tri_prefix_var.set(get_tri_prefix(sel_code, self.languages))

    def _fel_on_lang_changed(self, _event=None):
        self._fel_auto_sync_prefix(force=True)
        self._fel_refresh_previews()

    def _fel_resolve_effective_prefix(self) -> str:
        manual = self.fel_tri_prefix_var.get().strip().rstrip("_")
        if manual:
            return manual
        sel_code = self._get_fel_lang_code()
        if sel_code == "AUTO" and self.fel_files:
            p0 = Path(self.fel_files[0])
            hint_blob = " ".join([p0.parent.name] + [Path(x).name for x in self.fel_files[:3]])
            det = detect_lang_code_from_hint(hint_blob, self.languages)
            if det:
                sel_code = det
        return get_tri_prefix(sel_code, self.languages)

    def _fel_add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择待整理的弟兄交通 Word 文档（支持多选分篇）",
            filetypes=[("Word 文档", "*.docx")],
        )
        if paths:
            self._fel_ingest_paths([Path(x) for x in paths])

    def _fel_add_folder(self):
        d = filedialog.askdirectory(title="选择包含弟兄交通各分篇 (.docx) 的文件夹")
        if d:
            self._fel_ingest_paths([Path(d)])

    def _fel_remove_selected(self):
        sel = self.fel_tree.selection()
        if not sel:
            return
        indices = sorted((int(iid) for iid in sel), reverse=True)
        for idx in indices:
            if 0 <= idx < len(self.fel_files):
                removed = self.fel_files.pop(idx)
                self._fel_inspect_cache.pop(removed, None)
        self._fel_refresh_previews()

    def _fel_clear_all(self):
        self.fel_files.clear()
        self._fel_inspect_cache.clear()
        self._fel_refresh_previews()

    def _fel_browse_output(self):
        d = filedialog.askdirectory(title="选择弟兄交通定稿输出保存目录")
        if d:
            self.fel_output_dir_var.set(d)
            self._save_settings()

    def _fel_open_output_dir(self):
        d = self.fel_output_dir_var.get().strip()
        p = Path(d).resolve() if d else None
        if p and p.exists() and p.is_dir():
            os.startfile(str(p))
        else:
            messagebox.showinfo("提示", "请先选择真实存在的输出保存目录。")

    def _fel_open_selected_folder(self, _event=None):
        sel = self.fel_tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        if 0 <= idx < len(self.fel_files):
            p = Path(self.fel_files[idx])
            if p.parent.exists():
                os.startfile(str(p.parent))

    def _fel_inspect_cached(self, fp_str: str) -> dict:
        cached = self._fel_inspect_cache.get(fp_str)
        if cached is not None:
            return cached
        info = inspect_fellowship_docx_quick(Path(fp_str))
        self._fel_inspect_cache[fp_str] = info
        return info

    def _fel_refresh_previews(self):
        self._save_settings()
        for item in self.fel_tree.get_children():
            self.fel_tree.delete(item)

        pfx = self._fel_resolve_effective_prefix()
        dt = self.fel_date_var.get().strip() or datetime.now().strftime("%Y%m%d")
        mode = self.fel_output_mode_var.get()
        total_files = len(self.fel_files)
        first_base_title = ""
        total_rows = 0

        for idx, fp_str in enumerate(self.fel_files):
            p = Path(fp_str)
            try:
                info = self._fel_inspect_cached(fp_str)
                base_title = info["base_title"]
                if idx == 0:
                    first_base_title = base_title
                pnum = info["part_num"]
                if pnum is None and total_files > 1:
                    pnum = idx + 1
                part_disp = f"-{pnum:02d}" if pnum is not None else "单篇"
                total_rows += info["row_count"]
                single_preview = build_fellowship_output_name(pfx, base_title, dt, part_num=pnum)
                if mode == "merged_only":
                    shown_preview = f"[合并源 #{idx + 1}] → 汇总至总文档"
                else:
                    shown_preview = single_preview
                self.fel_tree.insert(
                    "",
                    tk.END,
                    iid=str(idx),
                    values=(
                        f"#{idx + 1}",
                        p.name,
                        part_disp,
                        f"{info['row_count']} 行",
                        shown_preview,
                    ),
                )
            except Exception as ex:
                self.fel_tree.insert(
                    "",
                    tk.END,
                    iid=str(idx),
                    values=(f"#{idx + 1}", p.name, "-", "-", f"解析失败：{ex}"),
                )

        if not self.fel_files:
            self.fel_summary_var.set("当前共 0 份分篇文档  |  支持直接拖入整个分篇文件夹或多个 .docx 自动排序")
        elif mode == "single_only":
            self.fel_summary_var.set(f"共 {total_files} 份文档（合计 {total_rows} 行） | 模式：仅输出各单篇（{pfx}_*-01-三语稿-{dt}.docx）")
        elif mode == "merged_only":
            merged_fn = build_fellowship_output_name(pfx, first_base_title or "弟兄交通", dt, part_num=None)
            self.fel_summary_var.set(f"共 {total_files} 份分篇 → 合并输出 1 份总文档：{merged_fn}")
        else:
            merged_fn = build_fellowship_output_name(pfx, first_base_title or "弟兄交通", dt, part_num=None)
            self.fel_summary_var.set(f"共 {total_files} 份单篇 + 1 份合并总文档（{merged_fn}）")

    def _on_run_fellowship(self):
        self._save_settings()
        if not self.fel_files:
            messagebox.showwarning("提示", "请先拖入或选择需要整理的弟兄交通 .docx 文档或分篇文件夹。")
            return
        out_dir = self.fel_output_dir_var.get().strip()
        if not out_dir:
            messagebox.showwarning("提示", "请先指定【输出保存目录】。")
            return

        use_ai = self.fel_enable_ai_var.get()
        keys = self._get_all_keys_list()
        if use_ai and not keys:
            proceed = messagebox.askyesno(
                "未配置 API Key 提醒",
                "当前已勾选「使用 AI 自动还原第 3 列段内局部蓝粗引文」，但尚未填写 API Key。\n\n"
                "• 拆内容控件、关闭修订模式、清批注、弯引号替换、四号/小四/五号字、固定值20磅行距、动态列宽均衡及整格蓝粗标题/问答均可【本地直接完成】。\n"
                "• 仅段落内部的局部跨语种神话/经文引文蓝粗定位需要调用 AI 模型。\n\n"
                "是否直接以纯本地极速模式继续执行？",
            )
            if not proceed:
                return

        model = self._get_model_id()
        api_mode = self._get_api_mode_code()
        rot_mode = self._get_rotation_mode_code()
        lang_code = self._get_fel_lang_code()
        tri_pfx = self.fel_tri_prefix_var.get().strip()
        quote_mode = self._get_fel_quote_mode_code()
        out_mode = self.fel_output_mode_var.get()
        dt_str = self.fel_date_var.get().strip()
        pool = ApiKeyPool(keys, rotation_mode=rot_mode) if (use_ai and keys) else None
        files_snapshot = list(self.fel_files)

        ai_tag = "AI开启(含503退避重试+备用模型兜底)" if use_ai else "纯本地离线极速模式"
        self._set_running(True, final_stage=f"正在启动【弟兄交通整理】（共 {len(files_snapshot)} 份文档）...")
        self._log("=" * 64)
        self._log(
            f"⚡ 启动【弟兄交通稿件整理】 | 文档数：{len(files_snapshot)} | 输出模式：{out_mode} | 语种：{lang_code} | {ai_tag}"
        )

        def worker():
            try:
                res = run_fellowship_pipeline(
                    input_files=files_snapshot,
                    output_dir=out_dir,
                    output_mode=out_mode,
                    lang_code=lang_code,
                    tri_prefix_override=tri_pfx,
                    quote_mode=quote_mode,
                    date_yyyymmdd=dt_str,
                    enable_ai_col3=use_ai,
                    api_key=pool,
                    model=model,
                    api_mode=api_mode,
                    rotation_mode=rot_mode,
                    log_fn=self._log,
                    progress_fn=self._progress,
                )
                saved = res["saved_files"]
                failed_ai = res.get("ai_failed_items") or []
                files_str = "\n".join(saved)
                self._log(f"🎉 弟兄交通整理全部完成！共输出 {len(saved)} 份文档（修订状态：关闭）：\n{files_str}")
                warn_block = ""
                if failed_ai:
                    warn_list = "\n".join(f"• {x}" for x in failed_ai)
                    warn_block = f"\n\n⚠️ 以下段落因模型接口繁忙暂未自动定位段内蓝粗（其余全部已完成导出，建议在 Word 中核对该行）：\n{warn_list}"
                self.log_queue.put(("done", (True, f"整理完成！共输出 {len(saved)} 份文档（修订: 关闭）")))
                self.log_queue.put(
                    (
                        "popup_info",
                        (
                            "弟兄交通整理完成",
                            f"前缀：{res['tri_prefix']}\n已输出 {len(saved)} 份定稿文档（修订模式已关闭）：\n\n{files_str}{warn_block}",
                        ),
                    )
                )
            except Exception as ex:
                tb = traceback.format_exc()
                self._log(f"❌ 弟兄交通整理出错：{ex}\n{tb}")
                self.log_queue.put(("done", (False, f"整理出错：{ex}")))
                self.log_queue.put(("popup_error", ("弟兄交通整理出错", str(ex))))

        threading.Thread(target=worker, daemon=True).start()

    # ---------------- Tab 2 Word Counter Methods ----------------

    def _wc_ingest_paths(self, paths: list[Path]):
        found_docx = self._extract_docx_from_paths(paths)
        if not found_docx:
            messagebox.showinfo("提示", "拖入或选择的项目中未找到任何有效的 .docx 文档。")
            return
        existing_norm = {os.path.normcase(str(Path(x).resolve())) for x in self.wc_files}
        for rp in found_docx:
            key = os.path.normcase(str(rp))
            if key not in existing_norm:
                existing_norm.add(key)
                self.wc_files.append(str(rp))
        self._wc_refresh_previews()

    def _wc_add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择要统计英文字数的 Word 文档（可多选）",
            filetypes=[("Word 文档", "*.docx")],
        )
        if not paths:
            return
        self._wc_ingest_paths([Path(x) for x in paths])

    def _wc_add_folder(self):
        d = filedialog.askdirectory(title="选择包含 Word 文档的文件夹")
        if not d:
            return
        self._wc_ingest_paths([Path(d)])

    def _wc_remove_selected(self):
        sel = self.wc_tree.selection()
        if not sel:
            return
        indices = sorted((int(iid) for iid in sel), reverse=True)
        for idx in indices:
            if 0 <= idx < len(self.wc_files):
                self.wc_files.pop(idx)
        self._wc_refresh_previews()

    def _wc_clear_all(self):
        self.wc_files.clear()
        self._wc_refresh_previews()

    def _wc_browse_copy_dir(self):
        d = filedialog.askdirectory(title="选择另存副本的目标文件夹")
        if d:
            self.wc_copy_dir_var.set(d)
            self.wc_save_mode_var.set("copy")

    def _wc_open_selected_folder(self, _event=None):
        sel = self.wc_tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        if 0 <= idx < len(self.wc_files):
            p = Path(self.wc_files[idx])
            target_dir = p.parent if p.parent.exists() else None
            if target_dir:
                os.startfile(str(target_dir))

    def _wc_analyze_one(self, file_path: str) -> dict:
        return analyze_docx_word_count(
            docx_path=Path(file_path),
            default_lang_code=self._get_wc_lang_code(),
            keep_existing_lang=self.wc_keep_existing_lang_var.get(),
            date_yyyymmdd=self.wc_date_var.get().strip(),
            keep_existing_date=self.wc_keep_existing_date_var.get(),
            strip_draft_prefix=self.wc_strip_draft_var.get(),
            lang_entries=self.languages,
        )

    def _wc_refresh_previews(self):
        self._save_settings()
        for item in self.wc_tree.get_children():
            self.wc_tree.delete(item)

        sum_up = 0
        sum_rv = 0
        for idx, fp in enumerate(self.wc_files):
            p = Path(fp)
            try:
                info = self._wc_analyze_one(fp)
                sum_up += info["up_words"]
                sum_rv += info["rv_words"]
                if info["old_name"] == info["new_name"]:
                    st = "✅ 已是最新"
                elif info["had_word_count"]:
                    st = "🔄 待更新字数"
                else:
                    st = "➕ 待新增字数"
                self.wc_tree.insert(
                    "",
                    tk.END,
                    iid=str(idx),
                    values=(
                        info["old_name"],
                        info["lang_code"],
                        f"{info['up_words']} words",
                        f"{info['rv_words']} words",
                        info["new_name"],
                        st,
                    ),
                )
            except Exception as ex:
                self.wc_tree.insert(
                    "",
                    tk.END,
                    iid=str(idx),
                    values=(p.name, "-", "-", "-", f"解析失败：{ex}", "❌ 错误"),
                )

        self.wc_summary_var.set(
            f"当前共 {len(self.wc_files)} 份文档  |  第 2 列英文合计：UP {sum_up} words, RV {sum_rv} words"
        )

    def _wc_execute_rename(self):
        self._save_settings()
        if not self.wc_files:
            messagebox.showwarning("提示", "请先点击「📄 选择 Word 文档」添加需要计算字数的 .docx 文件。")
            return

        copy_dir = None
        if self.wc_save_mode_var.get() == "copy":
            raw_cd = self.wc_copy_dir_var.get().strip()
            if not raw_cd:
                messagebox.showwarning("提示", "已选择“另存副本”，请先指定目标文件夹路径。")
                return
            copy_dir = Path(raw_cd).resolve()

        for item in self.wc_tree.get_children():
            self.wc_tree.delete(item)

        ok_cnt = 0
        err_msgs = []
        sum_up = 0
        sum_rv = 0

        for idx, fp in enumerate(list(self.wc_files)):
            p = Path(fp)
            try:
                info = self._wc_analyze_one(fp)
                sum_up += info["up_words"]
                sum_rv += info["rv_words"]
                final_p, status_txt = apply_word_count_filename(
                    docx_path=p,
                    new_filename=info["new_name"],
                    copy_to_dir=copy_dir,
                )
                if copy_dir is None:
                    self.wc_files[idx] = str(final_p)
                ok_cnt += 1
                self.wc_tree.insert(
                    "",
                    tk.END,
                    iid=str(idx),
                    values=(
                        info["old_name"],
                        info["lang_code"],
                        f"{info['up_words']} words",
                        f"{info['rv_words']} words",
                        final_p.name,
                        status_txt,
                    ),
                )
            except Exception as ex:
                err_msgs.append(f"• {p.name}：{ex}")
                self.wc_tree.insert(
                    "",
                    tk.END,
                    iid=str(idx),
                    values=(p.name, "-", "-", "-", str(ex), "❌ 失败"),
                )

        self.wc_summary_var.set(
            f"处理完成：成功 {ok_cnt}/{len(self.wc_files)} 份  |  合计：UP {sum_up} words, RV {sum_rv} words"
        )
        if err_msgs:
            messagebox.showwarning(
                "部分文档未能重命名",
                f"成功处理 {ok_cnt} 份，失败 {len(err_msgs)} 份：\n\n" + "\n".join(err_msgs[:10]),
            )
        else:
            messagebox.showinfo(
                "字数统计与重命名完成",
                f"已全部处理完成（共 {ok_cnt} 份文档）！\n"
                f"第 2 列英文总计：UP {sum_up} words, RV {sum_rv} words。",
            )

    # ---------------- Tab 1 Shared & Pipeline Methods ----------------

    def _on_paste_into_key_entry(self, _event=None):
        try:
            clip = self.root.clipboard_get()
        except Exception:
            return None
        if "\n" not in clip and "\r" not in clip:
            return None

        pasted_keys = parse_api_keys(clip)
        if not pasted_keys:
            return "break"

        existing_keys = []
        try:
            sel_first = self.key_entry.index(tk.SEL_FIRST)
            sel_last = self.key_entry.index(tk.SEL_LAST)
            cur_txt = self.api_key_var.get()
            rem_txt = cur_txt[:sel_first] + cur_txt[sel_last:]
            existing_keys = parse_api_keys(rem_txt)
        except tk.TclError:
            existing_keys = parse_api_keys(self.api_key_var.get())

        merged = parse_api_keys(existing_keys + pasted_keys)
        if not existing_keys:
            self._set_pool_multiline_and_sync_entry(clip)
        else:
            self._set_pool_multiline_and_sync_entry("\n".join(merged))
        self._save_settings()
        self._append_log(f"🔑 已从剪贴板粘贴多行密钥，当前密钥池共 {len(merged)} 个不重复 Key。")
        return "break"

    def _open_key_pool_dialog(self):
        init_txt = self._pool_multiline_text.strip() or "\n".join(self._get_all_keys_list())

        def _on_saved(new_raw_txt: str, new_rot_label: str):
            self.rotation_label.set(new_rot_label)
            self._set_pool_multiline_and_sync_entry(new_raw_txt)
            self._save_settings()
            cnt = len(self._get_all_keys_list())
            self._append_log(f"✅ 密钥池已更新：当前共 {cnt} 个 API Key | 策略：{new_rot_label}")

        ApiKeyPoolDialog(
            self.root,
            initial_raw_text=init_txt,
            rotation_label_var=self.rotation_label,
            model_id=self._get_model_id(),
            api_mode=self._get_api_mode_code(),
            on_save_cb=_on_saved,
        )

    def _open_lang_manager(self):
        self.languages = load_languages()

        def _on_saved(updated_list: list, selected_code=None):
            self.languages = [dict(x) for x in updated_list]
            combo_vals = get_combo_values(self.languages)
            self.lang_cb.configure(values=combo_vals)
            self.wc_lang_cb.configure(values=combo_vals)
            self.fel_lang_cb.configure(values=combo_vals)
            if selected_code:
                disp = self._code_to_combo_display(selected_code)
                self.lang_var.set(disp)
                self.wc_lang_var.set(disp)
                self.fel_lang_var.set(disp)
                self._fel_auto_sync_prefix(force=True)
            self._wc_refresh_previews()
            self._fel_refresh_previews()
            self._append_log(f"✅ 语种库已更新，当前共 {len(self.languages)} 个语种。")

        LanguageManagerDialog(self.root, self.languages, _on_saved)

    def _clear_log(self):
        self.log_text.delete("1.0", tk.END)

    def _toggle_key_visibility(self):
        self.key_entry.configure(show="" if self.show_key_var.get() else "•")

    def _browse_input(self):
        p = filedialog.askopenfilename(
            title="选择待整理 Word 文档",
            filetypes=[("Word 文档", "*.docx")],
        )
        if p:
            self.input_docx_var.set(p)
            if not self.output_dir_var.get().strip():
                self.output_dir_var.set(str(Path(p).parent))

    def _browse_library(self):
        d = filedialog.askdirectory(title="选择对应语种《七卷书》总库文件夹")
        if d:
            self.library_dir_var.set(d)

    def _browse_output(self):
        d = filedialog.askdirectory(title="选择输出保存文件夹")
        if d:
            self.output_dir_var.set(d)

    def _browse_draft(self):
        p = filedialog.askopenfilename(
            title="选择已复核的初筛稿 (.docx)",
            filetypes=[("Word 文档", "*.docx")],
        )
        if p:
            self.draft_docx_var.set(p)
            if not self.output_dir_var.get().strip():
                self.output_dir_var.set(str(Path(p).parent))

    def _open_output_dir(self):
        d = self.output_dir_var.get().strip()
        p = Path(d).resolve() if d else None
        if p and p.exists() and p.is_dir():
            os.startfile(str(p))
        else:
            messagebox.showinfo("提示", "请先选择真实存在的输出文件夹目录。")

    def _open_draft_file(self):
        raw = self.draft_docx_var.get().strip()
        p = Path(raw).resolve() if raw else None
        if p and p.is_file() and p.suffix.lower() == ".docx":
            os.startfile(str(p))
        else:
            messagebox.showinfo("提示", "当前还没有有效的 .docx 初筛稿文件，请先执行第一步或手动选择初筛稿。")

    def _log(self, msg: str):
        self.log_queue.put(("log", msg))

    def _progress(self, pct: float, stage: str):
        self.log_queue.put(("progress", (pct, stage)))

    def _format_pct_timer(self) -> str:
        pct_i = int(round(max(0.0, min(100.0, self.current_pct))))
        if self.task_start_ts > 0:
            elapsed = max(0, int(time.time() - self.task_start_ts))
            mm, ss = divmod(elapsed, 60)
            return f"{pct_i:3d}%  |  已耗时 {mm:02d}:{ss:02d}"
        return f"{pct_i}%"

    def _tick_elapsed_timer(self):
        try:
            if self.is_running and self.task_start_ts > 0:
                fmt = self._format_pct_timer()
                self.pct_timer_var.set(fmt)
                self.fel_pct_timer_var.set(fmt)
        except Exception:
            pass
        finally:
            self.root.after(500, self._tick_elapsed_timer)

    def _append_log(self, msg: str):
        if not msg:
            return
        safe_msg = redact_secrets(msg, extra_secret=self._get_all_keys_list())
        stripped = safe_msg.lstrip("\n")
        widgets = [w for w in (getattr(self, "log_text", None), getattr(self, "fel_log_text", None)) if w is not None]
        if stripped.startswith("=" * 10):
            for tw in widgets:
                tw.insert(tk.END, "\n" + stripped + "\n", "dim")
                tw.see(tk.END)
            return

        ts = datetime.now().strftime("%H:%M:%S")
        tag = None
        if any(k in stripped for k in ("▶", "⚡")):
            tag = "header"
        elif any(k in stripped for k in ("✅", "🎉", "Successfully aligned", "Final Document Saved")):
            tag = "ok"
        elif any(k in stripped for k in ("❌", "ERROR", "失败", "出错", "中止")):
            tag = "err"
        elif any(k in stripped for k in ("↻", "⚠️", "⏳", "提示", "提醒")):
            tag = "retry"

        line = f"[{ts}] {stripped}\n"
        for tw in widgets:
            if tag:
                tw.insert(tk.END, line, tag)
            else:
                tw.insert(tk.END, line)
            tw.see(tk.END)

    def _poll_log_queue(self):
        try:
            while True:
                try:
                    item = self.log_queue.get_nowait()
                except queue.Empty:
                    break

                try:
                    kind = item[0]
                    payload = item[1] if len(item) == 2 else item[1:]
                    all_secs = self._get_all_keys_list()

                    if kind == "log":
                        self._append_log(str(payload))
                    elif kind == "progress":
                        pct, stage_text = payload
                        self.current_pct = max(0.0, min(100.0, float(pct)))
                        self.progress_var.set(self.current_pct)
                        self.fel_progress_var.set(self.current_pct)
                        disp_stage = (
                            f"⏳ {stage_text}" if self.is_running and self.current_pct < 100 else f"✅ {stage_text}"
                        )
                        self.stage_var.set(disp_stage)
                        self.fel_stage_var.set(disp_stage)
                        fmt_t = self._format_pct_timer()
                        self.pct_timer_var.set(fmt_t)
                        self.fel_pct_timer_var.set(fmt_t)
                    elif kind == "draft_ready":
                        self.draft_docx_var.set(str(payload))
                    elif kind == "done":
                        ok_flag, final_stage = payload
                        self._set_running(False, ok=ok_flag, final_stage=final_stage)
                    elif kind == "popup_info":
                        title, text = payload
                        messagebox.showinfo(title, redact_secrets(str(text), extra_secret=all_secs))
                    elif kind == "popup_error":
                        title, text = payload
                        messagebox.showerror(title, redact_secrets(str(text), extra_secret=all_secs))
                except Exception as inner_ex:
                    self._append_log(f"❌ 日志队列处理异常：{inner_ex}")
        finally:
            self.root.after(80, self._poll_log_queue)

    def _set_running(self, running: bool, ok: bool = True, final_stage: str = ""):
        self.is_running = running
        state = tk.DISABLED if running else tk.NORMAL
        for b in (
            self.btn_step1,
            self.btn_step2,
            self.btn_all,
            self.test_btn,
            self.pool_btn,
            self.btn_fel_run,
            self.fel_pool_btn,
        ):
            b.configure(state=state)

        if running:
            self.task_start_ts = time.time()
            self.current_pct = 2.0
            self.progress_var.set(2.0)
            self.fel_progress_var.set(2.0)
            self.stage_lbl.configure(foreground="#1d4ed8")
            self.fel_stage_lbl.configure(foreground="#1d4ed8")
            if final_stage:
                self.stage_var.set(f"⏳ {final_stage}")
                self.fel_stage_var.set(f"⏳ {final_stage}")
            fmt = self._format_pct_timer()
            self.pct_timer_var.set(fmt)
            self.fel_pct_timer_var.set(fmt)
        else:
            if ok:
                self.current_pct = 100.0
                self.progress_var.set(100.0)
                self.fel_progress_var.set(100.0)
                self.stage_lbl.configure(foreground="#15803d")
                self.fel_stage_lbl.configure(foreground="#15803d")
                if final_stage:
                    self.stage_var.set(f"✅ {final_stage}")
                    self.fel_stage_var.set(f"✅ {final_stage}")
            else:
                self.stage_lbl.configure(foreground="#b91c1c")
                self.fel_stage_lbl.configure(foreground="#b91c1c")
                if final_stage:
                    self.stage_var.set(f"❌ {final_stage}")
                    self.fel_stage_var.set(f"❌ {final_stage}")
            fmt = self._format_pct_timer()
            self.pct_timer_var.set(fmt)
            self.fel_pct_timer_var.set(fmt)

    def _on_test_api(self):
        self._save_settings()
        keys = self._get_all_keys_list()
        model = self._get_model_id()
        mode = self._get_api_mode_code()
        if not keys:
            messagebox.showwarning("提示", "请先输入至少 1 个 Gemini API Key 或 Vertex API Key。")
            return

        self._set_running(True, final_stage=f"正在测试 {len(keys)} 个 API Key 与模型 {model} ...")
        self._progress(25.0, f"正在验证密钥池（共 {len(keys)} 个 Key）与模型 {model} ...")
        self._log(f"▶ 开始测试 API 密钥池（共 {len(keys)} 个 Key | 模型：{model}）...")

        def worker():
            try:
                msg = test_api_connection(keys, model, mode, log_fn=self._log)
                self._log(f"✅ {msg}")
                self.log_queue.put(("done", (True, msg)))
                self.log_queue.put(("popup_info", ("连接测试通过", msg)))
            except Exception as ex:
                err = f"连接失败：{ex}"
                self._log(f"❌ {err}")
                self.log_queue.put(("done", (False, "API 连接失败，请检查密钥或网络")))
                self.log_queue.put(("popup_error", ("API 连接失败", err)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_run_step1(self):
        self._save_settings()
        inp = self.input_docx_var.get().strip()
        lib = self.library_dir_var.get().strip()
        out = self.output_dir_var.get().strip()
        keys = self._get_all_keys_list()
        if not keys:
            messagebox.showwarning("提示", "请先填写至少 1 个 API Key（支持填写无限个自动轮换）。")
            return
        if not inp or not lib or not out:
            messagebox.showwarning("提示", "请先选择【待整理文档】、【对应语种总库目录】和【输出保存目录】。")
            return

        model = self._get_model_id()
        mode = self._get_api_mode_code()
        rot_mode = self._get_rotation_mode_code()
        lang = self._get_selected_lang_code()
        pool = ApiKeyPool(keys, rotation_mode=rot_mode)

        self._set_running(True, final_stage="正在启动第一步：生成初筛稿...")
        self._log("=" * 64)
        self._log(
            f"⚡ 开始执行【第一步：生成初筛稿】 | 模型：{model} | 密钥池：{len(pool)} 个 Key | 语种设定：{lang}"
        )

        def worker():
            try:
                res = run_step1_pipeline(
                    input_docx=inp,
                    library_dir=lib,
                    output_dir=out,
                    api_key=pool,
                    model=model,
                    api_mode=mode,
                    rotation_mode=rot_mode,
                    lang=lang,
                    log_fn=self._log,
                    progress_fn=self._progress,
                )
                draft_p = res["draft_path"]
                det_code = res.get("lang_code", "")
                det_disp = find_lang_display_name(det_code, self.languages)
                self.log_queue.put(("draft_ready", draft_p))
                self._log(f"✅ 第一步完成！当前语种：{det_disp} | 初筛稿已保存至：{draft_p}")
                self._log("💡 可点击上方【打开初筛稿】在 Word 中检查/微调黄青高亮，保存关闭后点击【② 第二步：由初筛稿生成定稿】。")
                self.log_queue.put(("done", (True, f"第一步完成！（语种：{det_disp}，初筛稿已就绪）")))
                self.log_queue.put(
                    (
                        "popup_info",
                        ("第一步完成", f"语种：{det_disp}\n初筛稿已生成：\n{draft_p}\n\n可人工复核后点击【第二步】输出定稿。"),
                    )
                )
            except SystemExit as se:
                msg = f"处理中止 (退出码 {se.code})，请查看下方实时日志详情（若目标文档已在 Word 打开请先关闭）。"
                self._log(f"❌ {msg}")
                self.log_queue.put(("done", (False, "第一步中止，请查看日志")))
                self.log_queue.put(("popup_error", ("执行中止", msg)))
            except Exception as ex:
                tb = traceback.format_exc()
                self._log(f"❌ 第一步出错：{ex}\n{tb}")
                self.log_queue.put(("done", (False, f"第一步出错：{ex}")))
                self.log_queue.put(("popup_error", ("第一步出错", str(ex))))

        threading.Thread(target=worker, daemon=True).start()

    def _on_run_step2(self):
        self._save_settings()
        draft = self.draft_docx_var.get().strip()
        out = self.output_dir_var.get().strip()
        keys = self._get_all_keys_list()
        if not keys:
            messagebox.showwarning("提示", "请先填写至少 1 个 API Key（支持填写无限个自动轮换）。")
            return
        if not draft or not out:
            messagebox.showwarning("提示", "请先选择【初筛确认稿】和【输出保存目录】。")
            return

        model = self._get_model_id()
        mode = self._get_api_mode_code()
        rot_mode = self._get_rotation_mode_code()
        lang = self._get_selected_lang_code()
        doc_name = self.doc_name_var.get().strip()
        dual = self.dual_outputs_var.get()
        pool = ApiKeyPool(keys, rotation_mode=rot_mode)

        self._set_running(True, final_stage="正在启动第二步：外文对齐并生成定稿...")
        self._log("=" * 64)
        self._log(
            f"⚡ 开始执行【第二步：由初筛稿生成定稿】 | 模型：{model} | 密钥池：{len(pool)} 个 Key | 语种设定：{lang}"
        )

        def worker():
            try:
                res = run_step2_pipeline(
                    draft_docx=draft,
                    output_dir=out,
                    api_key=pool,
                    model=model,
                    api_mode=mode,
                    rotation_mode=rot_mode,
                    lang=lang,
                    doc_name=doc_name,
                    dual_outputs=dual,
                    log_fn=self._log,
                    progress_fn=self._progress,
                )
                det_disp = find_lang_display_name(res.get("lang_code", ""), self.languages)
                files_str = "\n".join(res["saved_files"])
                self._log(f"🎉 第二步完成！（语种：{det_disp}）定稿已输出：\n{files_str}")
                self.log_queue.put(("done", (True, f"定稿生成完成（{det_disp}，共 {len(res['saved_files'])} 个文件）")))
                self.log_queue.put(("popup_info", ("定稿生成完成", f"语种：{det_disp}\n已成功生成以下定稿文件：\n\n{files_str}")))
            except SystemExit as se:
                msg = f"第二步中止 (退出码 {se.code})，请查看日志详情（若文件已在 Word 中打开，请先关闭后重试）。"
                self._log(f"❌ {msg}")
                self.log_queue.put(("done", (False, "第二步中止，请查看日志")))
                self.log_queue.put(("popup_error", ("第二步中止", msg)))
            except Exception as ex:
                tb = traceback.format_exc()
                self._log(f"❌ 第二步出错：{ex}\n{tb}")
                self.log_queue.put(("done", (False, f"第二步出错：{ex}")))
                self.log_queue.put(("popup_error", ("第二步出错", str(ex))))

        threading.Thread(target=worker, daemon=True).start()

    def _on_run_all(self):
        self._save_settings()
        inp = self.input_docx_var.get().strip()
        lib = self.library_dir_var.get().strip()
        out = self.output_dir_var.get().strip()
        keys = self._get_all_keys_list()
        if not keys:
            messagebox.showwarning("提示", "请先填写至少 1 个 API Key（支持填写无限个自动轮换）。")
            return
        if not inp or not lib or not out:
            messagebox.showwarning("提示", "请先选择【待整理文档】、【对应语种总库目录】和【输出保存目录】。")
            return

        model = self._get_model_id()
        mode = self._get_api_mode_code()
        rot_mode = self._get_rotation_mode_code()
        lang = self._get_selected_lang_code()
        doc_name = self.doc_name_var.get().strip()
        dual = self.dual_outputs_var.get()
        pool = ApiKeyPool(keys, rotation_mode=rot_mode)

        self._set_running(True, final_stage="正在启动一键全自动直出（第一步 + 第二步）...")
        self._log("=" * 64)
        self._log(
            f"⚡ 启动【一键全自动直出】（初筛评估 + 三份定稿导出） | 模型：{model} | 密钥池：{len(pool)} 个 Key | 语种设定：{lang}"
        )

        def _prog_s1(p: float, msg: str):
            self._progress(p * 0.40, f"[总阶段 1/2] {msg}")

        def _prog_s2(p: float, msg: str):
            self._progress(40.0 + p * 0.60, f"[总阶段 2/2] {msg}")

        def worker():
            try:
                r1 = run_step1_pipeline(
                    input_docx=inp,
                    library_dir=lib,
                    output_dir=out,
                    api_key=pool,
                    model=model,
                    api_mode=mode,
                    rotation_mode=rot_mode,
                    lang=lang,
                    log_fn=self._log,
                    progress_fn=_prog_s1,
                )
                draft_p = r1["draft_path"]
                det_lang = r1.get("lang_code") or lang
                det_disp = find_lang_display_name(det_lang, self.languages)
                self.log_queue.put(("draft_ready", draft_p))

                r2 = run_step2_pipeline(
                    draft_docx=draft_p,
                    output_dir=out,
                    api_key=pool,
                    model=model,
                    api_mode=mode,
                    rotation_mode=rot_mode,
                    lang=det_lang,
                    doc_name=doc_name,
                    dual_outputs=dual,
                    log_fn=self._log,
                    progress_fn=_prog_s2,
                )
                all_files = [draft_p] + r2["saved_files"]
                files_str = "\n".join(all_files)
                self._log(f"🎉 一键全自动直出全部完成！（语种：{det_disp}）共生成 {len(all_files)} 个文件：\n{files_str}")
                self.log_queue.put(("done", (True, f"一键直出完成！（{det_disp}，共生成 {len(all_files)} 个文件）")))
                self.log_queue.put(("popup_info", ("全部完成", f"语种：{det_disp}\n已成功生成初筛稿与定稿：\n\n{files_str}")))
            except SystemExit as se:
                msg = f"执行中止 (退出码 {se.code})，请查看运行日志（若目标文件已在 Word 中打开，请先关闭）。"
                self._log(f"❌ {msg}")
                self.log_queue.put(("done", (False, "执行中止，请查看日志")))
                self.log_queue.put(("popup_error", ("执行中止", msg)))
            except Exception as ex:
                tb = traceback.format_exc()
                self._log(f"❌ 执行出错：{ex}\n{tb}")
                self.log_queue.put(("done", (False, f"执行出错：{ex}")))
                self.log_queue.put(("popup_error", ("执行出错", str(ex))))

        threading.Thread(target=worker, daemon=True).start()

    def _on_close(self):
        self._save_settings()
        try:
            self.dnd.unhook()
        except Exception:
            pass
        self.root.destroy()


def _run_frozen_self_test() -> int:
    import docx_xml_utils
    import step1_extract_and_diff  # noqa: F401
    import step1_render_draft  # noqa: F401
    import step2_parse_reviewed  # noqa: F401
    import step2_render_final  # noqa: F401
    import win_dnd  # noqa: F401

    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        bundled_cfg = Path(sys._MEIPASS) / "resources" / "config.json"
        if not bundled_cfg.is_file():
            raise FileNotFoundError(f"Bundled config.json missing inside EXE: {bundled_cfg}")

    cfg = docx_xml_utils.load_config()
    if cfg.get("col2_up_highlight") != "yellow" or not cfg.get("language_code_map"):
        raise RuntimeError(f"Invalid config loaded in self-test: {cfg}")
    langs = load_languages()
    if not langs:
        raise RuntimeError("Empty language registry in self-test")
    pool = ApiKeyPool(["self_test_placeholder_1", "self_test_placeholder_2"])
    if len(pool) != 2:
        raise RuntimeError("ApiKeyPool self-test failed")

    # Verify word_counter filename parser & formatter
    p1 = parse_doc_filename("GR_God’s Words_(UP 10 words, RV 20 words)_20261114", known_codes={"GR"})
    if p1["orig_name"] != "God’s Words" or p1["existing_lang"] != "GR" or not p1["had_word_count"]:
        raise RuntimeError(f"Filename parser self-test failed: {p1}")
    fn = format_counted_docx_name("GR", "God’s Words", 128, 64, "20261115")
    if fn != "GR_God’s Words_(UP 128 words, RV 64 words)_20261115.docx":
        raise RuntimeError(f"Filename formatter self-test failed: {fn}")

    # Verify Scenario 4 fellowship_processor self-test
    import fellowship_processor  # noqa: F401

    bt, pn = parse_title_and_part("怎样分辨各种人（二）Part 1")
    if bt != "怎样分辨各种人（二）" or pn != 1:
        raise RuntimeError(f"Fellowship title parser self-test failed: {(bt, pn)}")
    fel_fn = build_fellowship_output_name("中英缅", "《怎样实行真理（二）》-上面的交通 2026年11月2日", "20261115", part_num=1)
    if fel_fn != "中英缅_《怎样实行真理（二）》-上面的交通 2026年11月2日-01-三语稿-20261115.docx":
        raise RuntimeError(f"Fellowship filename self-test failed: {fel_fn}")
    cit_s = find_post_quote_citation_spans("„Защото който не е против нас“ (Лука 9:50).")
    if cit_s != [(31, 42)]:
        raise RuntimeError(f"Fellowship citation self-test failed: {cit_s}")
    if find_post_quote_citation_spans("到现在《话在肉身显现》已经有七卷书了") != []:
        raise RuntimeError("Fellowship narrative exclusion self-test failed")
    return 0


def main():
    if "--self-test" in sys.argv:
        sys.exit(_run_frozen_self_test())
    _enable_high_dpi()
    root = tk.Tk()
    ScatteredRevisionApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
