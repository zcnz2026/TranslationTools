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
from lang_registry import (
    AUTO_COMBO_LABEL,
    DEFAULT_LANGUAGES,
    build_keywords_list,
    find_lang_display_name,
    format_combo_label,
    get_combo_values,
    get_languages_file,
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

        # Multiline key editor
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

        # Count & Rotation Mode Row
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

        # Batch Health Check Output Box
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

        # Bottom Action Bar
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

        # Treeview table
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

        # Edit / Add form
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

        # Buttons bar
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
        self.root.title("零散改点自动整理工具 v1.3  (多 Key 无限轮换防限流 + 多语种自由扩展)")
        self.root.geometry("950x780")
        self.root.minsize(870, 700)

        self.log_queue: queue.Queue = queue.Queue()
        self.is_running = False
        self.task_start_ts: float = 0.0
        self.current_pct: float = 0.0
        self.languages: list = load_languages()

        self._pool_multiline_text: str = ""
        self._syncing_key_var: bool = False

        self._setup_styles()
        self._init_vars()
        self._load_settings()
        self._build_ui()
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

        self.root.option_add("*Font", default_font)
        style.configure("TLabelframe.Label", font=bold_font, foreground="#1a56db")
        style.configure("Primary.TButton", font=btn_font, padding=(12, 7))
        style.configure("Action.TButton", font=default_font, padding=(10, 6))

    def _init_vars(self):
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
                # Auto-migrate legacy plaintext api_key to Windows DPAPI and purge plaintext from disk
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
                self.dual_outputs_var.set(bool(cfg.get("dual_outputs", True)))
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
            "dual_outputs": self.dual_outputs_var.get(),
        }
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

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
        path_frame = ttk.LabelFrame(main, text=" 2. 文档、总库路径与多语种设置 ", padding=10)
        path_frame.pack(fill=tk.X, pady=(0, 8))
        path_frame.columnconfigure(1, weight=1)

        ttk.Label(path_frame, text="待整理文档 (.docx)：").grid(row=0, column=0, sticky=tk.W, pady=4)
        ttk.Entry(path_frame, textvariable=self.input_docx_var).grid(row=0, column=1, sticky=tk.EW, pady=4, padx=4)
        ttk.Button(path_frame, text="浏览...", width=10, command=self._browse_input).grid(row=0, column=2, padx=2)

        ttk.Label(path_frame, text="对应语种总库目录：").grid(row=1, column=0, sticky=tk.W, pady=4)
        ttk.Entry(path_frame, textvariable=self.library_dir_var).grid(row=1, column=1, sticky=tk.EW, pady=4, padx=4)
        ttk.Button(path_frame, text="浏览...", width=10, command=self._browse_library).grid(row=1, column=2, padx=2)

        ttk.Label(path_frame, text="输出保存目录：").grid(row=2, column=0, sticky=tk.W, pady=4)
        ttk.Entry(path_frame, textvariable=self.output_dir_var).grid(row=2, column=1, sticky=tk.EW, pady=4, padx=4)
        out_btns = ttk.Frame(path_frame)
        out_btns.grid(row=2, column=2, columnspan=2, sticky=tk.W)
        ttk.Button(out_btns, text="浏览...", width=10, command=self._browse_output).pack(side=tk.LEFT, padx=2)
        ttk.Button(out_btns, text="打开目录", width=10, command=self._open_output_dir).pack(side=tk.LEFT, padx=2)

        ttk.Label(path_frame, text="初筛确认稿 (第二步用)：").grid(row=3, column=0, sticky=tk.W, pady=4)
        ttk.Entry(path_frame, textvariable=self.draft_docx_var).grid(row=3, column=1, sticky=tk.EW, pady=4, padx=4)
        draft_btns = ttk.Frame(path_frame)
        draft_btns.grid(row=3, column=2, columnspan=2, sticky=tk.W)
        ttk.Button(draft_btns, text="浏览...", width=10, command=self._browse_draft).pack(side=tk.LEFT, padx=2)
        ttk.Button(draft_btns, text="打开初筛稿", width=10, command=self._open_draft_file).pack(side=tk.LEFT, padx=2)

        # Multi-language selector + Manager button + Title prefix
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

        ttk.Label(opt_row1, text="定稿标题前缀(选填)：").pack(side=tk.LEFT)
        ttk.Entry(opt_row1, textvariable=self.doc_name_var, width=22).pack(side=tk.LEFT, padx=(4, 0))

        opt_row2 = ttk.Frame(path_frame)
        opt_row2.grid(row=5, column=0, columnspan=4, sticky=tk.EW, pady=(4, 0))
        ttk.Checkbutton(
            opt_row2,
            text="同时输出【完整留底版】+【发译者版（自动剔除无改点表）】",
            variable=self.dual_outputs_var,
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
            text="⚡ 一键直出全部定稿 (初筛 + 定稿)",
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
            height=11,
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
            f"欢迎使用零散改点自动整理工具 v1.3！当前已载入 {key_cnt} 个 API Key、{len(self.languages)} 种语言（支持多 Key 无限轮换与报错自动切换）。"
        )

    def _on_paste_into_key_entry(self, _event=None):
        """Gracefully accept multi-line clipboard pastes directly in the single-line Entry without dropping lines."""
        try:
            clip = self.root.clipboard_get()
        except Exception:
            return None
        if "\n" not in clip and "\r" not in clip:
            return None

        # Merge existing keys (unless entire field is selected) with pasted multi-line keys
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
        # Preserve comments from clipboard if replacing entire field
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
            self.lang_cb.configure(values=get_combo_values(self.languages))
            if selected_code:
                self.lang_var.set(self._code_to_combo_display(selected_code))
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
                self.pct_timer_var.set(self._format_pct_timer())
        except Exception:
            pass
        finally:
            self.root.after(500, self._tick_elapsed_timer)

    def _append_log(self, msg: str):
        if not msg:
            return
        safe_msg = redact_secrets(msg, extra_secret=self._get_all_keys_list())
        stripped = safe_msg.lstrip("\n")
        if stripped.startswith("=" * 10):
            self.log_text.insert(tk.END, "\n" + stripped + "\n", "dim")
            self.log_text.see(tk.END)
            return

        ts = datetime.now().strftime("%H:%M:%S")
        tag = None
        if any(k in stripped for k in ("▶", "⚡")):
            tag = "header"
        elif any(k in stripped for k in ("✅", "🎉", "Successfully aligned", "Final Document Saved")):
            tag = "ok"
        elif any(k in stripped for k in ("❌", "ERROR", "失败", "出错", "中止")):
            tag = "err"
        elif any(k in stripped for k in ("↻", "⚠️", "⏳", "提示")):
            tag = "retry"

        line = f"[{ts}] {stripped}\n"
        if tag:
            self.log_text.insert(tk.END, line, tag)
        else:
            self.log_text.insert(tk.END, line)
        self.log_text.see(tk.END)

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
                        self.stage_var.set(
                            f"⏳ {stage_text}" if self.is_running and self.current_pct < 100 else f"✅ {stage_text}"
                        )
                        self.pct_timer_var.set(self._format_pct_timer())
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
        for b in (self.btn_step1, self.btn_step2, self.btn_all, self.test_btn, self.pool_btn):
            b.configure(state=state)

        if running:
            self.task_start_ts = time.time()
            self.current_pct = 2.0
            self.progress_var.set(2.0)
            self.stage_lbl.configure(foreground="#1d4ed8")
            if final_stage:
                self.stage_var.set(f"⏳ {final_stage}")
            self.pct_timer_var.set(self._format_pct_timer())
        else:
            if ok:
                self.current_pct = 100.0
                self.progress_var.set(100.0)
                self.stage_lbl.configure(foreground="#15803d")
                if final_stage:
                    self.stage_var.set(f"✅ {final_stage}")
            else:
                self.stage_lbl.configure(foreground="#b91c1c")
                if final_stage:
                    self.stage_var.set(f"❌ {final_stage}")
            self.pct_timer_var.set(self._format_pct_timer())

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
        # Shared pool across Step 1 + Step 2 so dead/rate-limited keys stay remembered seamlessly
        pool = ApiKeyPool(keys, rotation_mode=rot_mode)

        self._set_running(True, final_stage="正在启动一键全自动直出（第一步 + 第二步）...")
        self._log("=" * 64)
        self._log(
            f"⚡ 启动【一键全自动直出】（初筛评估 + 外文同步高亮定稿） | 模型：{model} | 密钥池：{len(pool)} 个 Key | 语种设定：{lang}"
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
        self.root.destroy()


def main():
    _enable_high_dpi()
    root = tk.Tk()
    ScatteredRevisionApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
