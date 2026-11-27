"""
Modular GUI Tab for Scenario 3: 「改点入库与通读定稿」
  - Sub-tab 3A: 改点返稿入库（发稿 + 回稿 + 总库 + 可选通读表）— 全栏位原生文件/文件夹拖拽
  - Sub-tab 3B: 通读定稿入库（Google Sheets / Excel + JSON授权导入 + 自定义列名）— 全栏位原生拖拽
"""

from __future__ import annotations

import ctypes
import os
from pathlib import Path
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import List, Optional, Sequence

from ai_client import ApiKeyPool
from library_ingester import run_library_ingestion_pipeline
from proofread_ingester import run_proofread_ingestion_pipeline
from security_utils import open_path_in_os
from sheets_reader import ColumnMappingConfig, GoogleOAuthManager
from win_dnd import WinFileDropManager

_UI_FONT = "PingFang SC" if sys.platform == "darwin" else "Microsoft YaHei UI"
_MONO_FONT = "Menlo" if sys.platform == "darwin" else "Consolas"
_PT_9 = 12 if sys.platform == "darwin" else 9


class _WinRect(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


class _MonitorInfo(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_ulong),
        ("rcMonitor", _WinRect),
        ("rcWork", _WinRect),
        ("dwFlags", ctypes.c_ulong),
    ]


def _get_parent_monitor_work_area(win: tk.Misc, parent: Optional[tk.Misc] = None) -> tuple[int, int, int, int]:
    """Return `(left, top, width, height)` of the monitor work area containing `parent` (multi-monitor aware)."""
    if sys.platform == "win32":
        try:
            ref = parent if (parent is not None and parent.winfo_exists()) else win
            ref.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(ref.winfo_id()) or ref.winfo_id()
            hmon = ctypes.windll.user32.MonitorFromWindow(hwnd, 2)  # MONITOR_DEFAULTTONEAREST
            if hmon:
                mi = _MonitorInfo()
                mi.cbSize = ctypes.sizeof(_MonitorInfo)
                if ctypes.windll.user32.GetMonitorInfoW(hmon, ctypes.byref(mi)):
                    rc = mi.rcWork
                    if rc.right > rc.left and rc.bottom > rc.top:
                        return int(rc.left), int(rc.top), int(rc.right - rc.left), int(rc.bottom - rc.top)
            rect = _WinRect()
            if ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0):
                if rect.right > rect.left and rect.bottom > rect.top:
                    return int(rect.left), int(rect.top), int(rect.right - rect.left), int(rect.bottom - rect.top)
        except Exception:
            pass
    sw = max(800, int(win.winfo_screenwidth()))
    sh = max(600, int(win.winfo_screenheight()) - 48)
    return 0, 0, sw, sh


def _center_and_fit_dialog(
    dlg: tk.Toplevel,
    parent: Optional[tk.Misc],
    base_w: int,
    base_h: int,
    min_w: int,
    min_h: int,
) -> None:
    """DPI-adaptively size `dlg` so no widgets clip, and center it squarely over `parent`."""
    dlg.update_idletasks()
    try:
        dpi = float(dlg.winfo_fpixels("1i"))
    except Exception:
        dpi = 96.0
    scale = max(1.0, dpi / 96.0)

    work_left, work_top, work_w, work_h = _get_parent_monitor_work_area(dlg, parent)
    avail_w = max(480, work_w - 32)
    avail_h = max(320, work_h - 48)

    req_w = int(dlg.winfo_reqwidth())
    req_h = int(dlg.winfo_reqheight())

    w = min(avail_w, max(int(round(base_w * scale)), req_w + 32))
    h = min(avail_h, max(int(round(base_h * scale)), req_h + 24))

    eff_min_w = min(w, max(int(round(min_w * scale)), min(avail_w, req_w + 16)))
    eff_min_h = min(h, max(int(round(min_h * scale)), min(avail_h, req_h + 12)))
    dlg.minsize(eff_min_w, eff_min_h)

    px = work_left + (work_w - w) // 2
    py = work_top + (work_h - h) // 2
    if parent is not None and parent.winfo_exists():
        try:
            top_p = parent.winfo_toplevel()
            top_p.update_idletasks()
            pw = int(top_p.winfo_width())
            ph = int(top_p.winfo_height())
            if pw > 120 and ph > 120:
                px = int(top_p.winfo_rootx()) + (pw - w) // 2
                py = int(top_p.winfo_rooty()) + (ph - h) // 2
        except Exception:
            pass

    px = max(work_left + 8, min(px, work_left + work_w - w - 8))
    py = max(work_top + 8, min(py, work_top + work_h - h - 8))
    dlg.geometry(f"{w}x{h}+{px}+{py}")
    dlg.update_idletasks()


def _make_drop_banner(parent: tk.Misc, text: str, on_click) -> tk.Frame:
    zone = tk.Frame(
        parent,
        bg="#eff6ff",
        highlightbackground="#93c5fd",
        highlightthickness=1,
        padx=12,
        pady=6,
        cursor="hand2",
    )
    lbl = tk.Label(
        zone,
        text=text,
        font=(_UI_FONT, _PT_9, "bold"),
        fg="#1d4ed8",
        bg="#eff6ff",
        anchor=tk.W,
        cursor="hand2",
    )
    lbl.pack(fill=tk.X)
    zone.bind("<Button-1>", lambda _e: on_click())
    lbl.bind("<Button-1>", lambda _e: on_click())
    zone.bind("<Enter>", lambda _e: (zone.configure(bg="#dbeafe"), lbl.configure(bg="#dbeafe")))
    zone.bind("<Leave>", lambda _e: (zone.configure(bg="#eff6ff"), lbl.configure(bg="#eff6ff")))
    return zone


class GoogleOAuthDialog(tk.Toplevel):
    """
    Modal dialog for two-step Google OAuth 2.0 authorization:
      Step 1: Import `.json` credential file (`credentials.json` / `client_secret_*.json`) via picker or drag-and-drop.
      Step 2: Authorize Google account in browser (tokens & client JSON encrypted at rest via Windows DPAPI).
    """

    def __init__(self, parent: tk.Misc, oauth: GoogleOAuthManager, on_updated):
        super().__init__(parent)
        self.withdraw()
        self.oauth = oauth
        self.on_updated = on_updated
        self.title("Google Sheets 授权（导入 JSON 授权文件 → 授权账户）")
        self.resizable(True, True)
        self.transient(parent)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        frm = ttk.Frame(self, padding=16)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            frm,
            text="说明：公开可读的 Google Sheets 链接无需授权即可直接读取。\n"
                 "读取受限表格只需两步：① 导入 Google 授权 JSON 文件 → ② 点击「授权 Google 账户」完成浏览器登录。",
            foreground="#475569",
            justify=tk.LEFT,
        ).pack(anchor=tk.W, fill=tk.X, pady=(0, 10))

        # Step 1 Frame: Import JSON file (supports click & native drag-and-drop)
        step1_box = ttk.LabelFrame(frm, text=" 第一步：导入 Google JSON 授权文件 ", padding=12)
        step1_box.pack(fill=tk.X, pady=(0, 10))

        drop_zone = _make_drop_banner(
            step1_box,
            "📂 点击此处选择（或直接把 .json 授权文件拖入此窗口）导入 credentials.json",
            self._browse_json,
        )
        drop_zone.pack(fill=tk.X, pady=(0, 8))

        self.json_info_var = tk.StringVar(value="")
        ttk.Label(step1_box, textvariable=self.json_info_var, foreground="#0f172a").pack(anchor=tk.W, fill=tk.X)

        # Step 2 Frame: Account authorization
        step2_box = ttk.LabelFrame(frm, text=" 第二步：登录并授权 Google 账户 ", padding=12)
        step2_box.pack(fill=tk.X, pady=(0, 12))

        self.status_var = tk.StringVar(value="")
        ttk.Label(
            step2_box,
            textvariable=self.status_var,
            foreground="#1d4ed8",
            font=(_UI_FONT, _PT_9, "bold"),
        ).pack(anchor=tk.W, fill=tk.X, pady=(0, 10))

        btn_row = ttk.Frame(step2_box)
        btn_row.pack(fill=tk.X)
        self.auth_btn = ttk.Button(btn_row, text="🌐 授权 Google 账户", command=self._on_authorize_account)
        self.auth_btn.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_row, text="⚡ 导入 JSON 并立即授权", command=self._import_and_authorize).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(btn_row, text="退出账户", command=self._on_logout).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_row, text="清除 JSON", command=self._on_clear_all).pack(side=tk.LEFT)

        foot = ttk.Frame(frm)
        foot.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(foot, text="完成并关闭", width=14, command=self._on_close).pack(side=tk.RIGHT)

        self._refresh_labels()
        _center_and_fit_dialog(self, parent, base_w=640, base_h=360, min_w=580, min_h=320)
        self.deiconify()
        self.lift()
        self.grab_set()
        self.focus_force()

        # Enable native Windows drag-and-drop directly onto this modal dialog HWND
        self._dlg_dnd = WinFileDropManager(self, fallback_cb=self._on_drop_paths)
        self.after(80, self._dlg_dnd.install)

    def _on_close(self):
        try:
            self._dlg_dnd.unhook()
        except Exception:
            pass
        self.destroy()

    def _refresh_labels(self):
        if self.oauth.has_client_config:
            fn = self.oauth.json_filename or "已保存的 OAuth 客户端配置"
            proj = f" | 项目：{self.oauth.project_id}" if self.oauth.project_id else ""
            cid_short = (
                f"{self.oauth.client_id[:22]}..." if len(self.oauth.client_id) > 25 else self.oauth.client_id
            )
            self.json_info_var.set(f"✅ 已导入 JSON：{fn}{proj}（Client ID: {cid_short}）")
        else:
            self.json_info_var.set("尚未导入 JSON 授权文件（请点击上方蓝色区域或拖入 .json 文件）")

        if self.oauth.is_authorized:
            self.status_var.set(f"当前账户状态：✅ 已授权（{self.oauth.user_email}）")
        elif self.oauth.has_client_config:
            self.status_var.set("当前账户状态：已导入 JSON，请点击下方「🌐 授权 Google 账户」完成浏览器登录")
        else:
            self.status_var.set("当前账户状态：未授权（仅支持链接公开可读的表格）")

    def _on_drop_paths(self, paths: List[Path]):
        json_files = [p for p in paths if p.is_file() and p.suffix.lower() == ".json"]
        if not json_files:
            messagebox.showwarning("提示", "请拖入 .json 格式的 Google 授权文件。", parent=self)
            return
        self._do_import_json(json_files[0], auto_auth=False)

    def _browse_json(self, auto_auth: bool = False):
        p = filedialog.askopenfilename(
            parent=self,
            title="选择 Google OAuth 2.0 授权 JSON 文件",
            filetypes=[("JSON 授权文件", "*.json"), ("所有文件", "*.*")],
        )
        if p:
            self._do_import_json(Path(p), auto_auth=auto_auth)

    def _import_and_authorize(self):
        self._browse_json(auto_auth=True)

    def _do_import_json(self, json_path: Path, auto_auth: bool = False):
        try:
            info = self.oauth.import_json_file(json_path)
            self._refresh_labels()
            if callable(self.on_updated):
                self.on_updated()
            if info.get("already_authorized"):
                messagebox.showinfo("导入成功", f"已导入 {info['filename']} 并识别到有效授权凭证！", parent=self)
            elif auto_auth:
                self._on_authorize_account()
            else:
                messagebox.showinfo(
                    "JSON 导入成功",
                    f"已成功导入：{info['filename']}\n\n接下来请点击「🌐 授权 Google 账户」在浏览器中登录即可！",
                    parent=self,
                )
        except Exception as ex:
            messagebox.showerror("导入 JSON 失败", str(ex), parent=self)

    def _on_authorize_account(self):
        if not self.oauth.has_client_config:
            messagebox.showwarning("请先导入 JSON", "请先在「第一步」导入 Google OAuth 2.0 的 .json 授权文件。", parent=self)
            return
        self.auth_btn.configure(state=tk.DISABLED)
        self.status_var.set("当前账户状态：⏳ 正在打开浏览器等待 Google 账户登录授权...")

        def _worker():
            try:
                email = self.oauth.authorize_interactive()
                self.after(0, lambda: self._finish_auth(True, email))
            except Exception as ex:
                self.after(0, lambda e=str(ex): self._finish_auth(False, e))

        threading.Thread(target=_worker, daemon=True).start()

    def _finish_auth(self, ok: bool, msg: str):
        if not self.winfo_exists():
            return
        self.auth_btn.configure(state=tk.NORMAL)
        self._refresh_labels()
        if ok:
            if callable(self.on_updated):
                self.on_updated()
            messagebox.showinfo("账户授权成功", f"✅ 已成功授权 Google 账户：{msg}", parent=self)
        else:
            messagebox.showerror("账户授权失败", msg, parent=self)

    def _on_logout(self):
        self.oauth.clear_tokens(clear_client_json=False)
        self._refresh_labels()
        if callable(self.on_updated):
            self.on_updated()

    def _on_clear_all(self):
        self.oauth.clear_tokens(clear_client_json=True)
        self._refresh_labels()
        if callable(self.on_updated):
            self.on_updated()


class ColumnMappingDialog(tk.Toplevel):
    """Modal dialog for customizing the 5 proofread spreadsheet column header aliases."""

    def __init__(self, parent: tk.Misc, col_cfg: ColumnMappingConfig, on_save):
        super().__init__(parent)
        self.withdraw()
        self.col_cfg = col_cfg
        self.on_save = on_save
        self.title("自定义通读表格列名识别规则")
        self.resizable(True, True)
        self.transient(parent)

        frm = ttk.Frame(self, padding=16)
        frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(
            frm,
            text="每列可填多个别名（用英文逗号分隔，支持中文/英文/任意语种表头，自动忽略大小写与空格）：",
            foreground="#475569",
        ).pack(anchor=tk.W, fill=tk.X, pady=(0, 10))

        grid = ttk.Frame(frm)
        grid.pack(fill=tk.X, expand=True)
        grid.columnconfigure(1, weight=1)

        fields = [
            ("cn", "中文原文列："),
            ("en", "英文列："),
            ("target", "目标语言列："),
            ("changes", "改动列："),
            ("final", "定稿列："),
        ]
        self.vars = {}
        for r_i, (k, lbl) in enumerate(fields):
            ttk.Label(grid, text=lbl).grid(row=r_i, column=0, sticky=tk.W, pady=4)
            v = tk.StringVar(value=getattr(col_cfg, k, ""))
            self.vars[k] = v
            ttk.Entry(grid, textvariable=v).grid(row=r_i, column=1, sticky=tk.EW, padx=(6, 0), pady=4)

        btns = ttk.Frame(frm)
        btns.pack(fill=tk.X, pady=(14, 0))
        ttk.Button(btns, text="恢复默认", command=self._reset_defaults).pack(side=tk.LEFT)
        ttk.Button(btns, text="取消", width=10, command=self.destroy).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(btns, text="保存列名设置", width=14, command=self._save).pack(side=tk.RIGHT)

        _center_and_fit_dialog(self, parent, base_w=620, base_h=310, min_w=540, min_h=280)
        self.deiconify()
        self.lift()
        self.grab_set()
        self.focus_force()

    def _reset_defaults(self):
        d = ColumnMappingConfig()
        for k, v in self.vars.items():
            v.set(getattr(d, k))

    def _save(self):
        for k, v in self.vars.items():
            setattr(self.col_cfg, k, v.get().strip())
        if callable(self.on_save):
            self.on_save()
        self.destroy()


class IngestTabPanel:
    """
    Self-contained Scenario 3 controller & UI panel mounted inside `app_gui.ScatteredRevisionApp`,
    with full native Windows Explorer drag-and-drop across all inputs.
    """

    def __init__(self, app, parent: ttk.Frame):
        self.app = app
        self.parent = parent
        self._running = False
        self._last_out_dir = ""

        cfg = getattr(app, "settings", {}) or {}
        self.oauth = GoogleOAuthManager.from_dpapi_token(str(cfg.get("ing_google_oauth_dpapi", "")))
        self.col_cfg = ColumnMappingConfig.from_dict(cfg.get("ing_col_mapping"))

        # 3A Vars — all task & library paths start empty on fresh launch (`都清空`)
        self.sent_path_var = tk.StringVar(value="")
        self.ret_path_var = tk.StringVar(value="")
        self.lib_dir_var = tk.StringVar(value="")
        self.out_dir_var = tk.StringVar(value="")
        self.conflict_policy_var = tk.StringVar(value=str(cfg.get("ing_conflict_policy", "keep_library")))
        self.rev_merge_mode_var = tk.StringVar(value=str(cfg.get("ing_rev_merge_mode", "single")))
        self.rev_batch_size_var = tk.StringVar(value=str(cfg.get("ing_rev_batch_size", "10")))

        # 3B Vars — all task & library paths start empty on fresh launch (`都清空`)
        self.pr_lib_dir_var = tk.StringVar(value="")
        self.pr_out_dir_var = tk.StringVar(value="")
        self.pr_enable_ai_var = tk.BooleanVar(value=bool(cfg.get("ing_pr_enable_ai", False)))

        # Status Vars
        self.oauth_status_var = tk.StringVar(value="")
        self.progress_var = tk.DoubleVar(value=0.0)
        self.stage_var = tk.StringVar(value="状态：就绪（支持直接拖入文件/文件夹，绝不改动原总库）")

        self._refresh_oauth_label()
        self._build_ui()

    def export_settings_dict(self) -> dict:
        enc_oauth = ""
        if self.oauth.client_id or self.oauth.refresh_token or self.oauth.access_token:
            try:
                enc_oauth = self.oauth.to_dpapi_token()
            except Exception:
                enc_oauth = ""
        return {
            "ing_conflict_policy": self.conflict_policy_var.get(),
            "ing_rev_merge_mode": self.rev_merge_mode_var.get(),
            "ing_rev_batch_size": self.rev_batch_size_var.get().strip() or "10",
            "ing_pr_enable_ai": self.pr_enable_ai_var.get(),
            "ing_col_mapping": self.col_cfg.to_dict(),
            "ing_google_oauth_dpapi": enc_oauth,
        }

    def _save_app_settings(self):
        if hasattr(self.app, "_save_settings"):
            self.app._save_settings()

    def _refresh_oauth_label(self):
        if self.oauth.is_authorized:
            self.oauth_status_var.set(f"Google 授权：✅ 已登录（{self.oauth.user_email}）")
        elif self.oauth.has_client_config:
            fn = self.oauth.json_filename or "JSON 已导入"
            self.oauth_status_var.set(f"Google 授权：已导入 {fn}（待登录账户）")
        else:
            self.oauth_status_var.set("Google 授权：未导入 JSON（公开链接免授权）")

    def _build_ui(self):
        pw = ttk.Panedwindow(self.parent, orient=tk.VERTICAL)
        pw.pack(fill=tk.BOTH, expand=True)

        top_box = ttk.Frame(pw)
        bot_box = ttk.Frame(pw)
        pw.add(top_box, weight=3)
        pw.add(bot_box, weight=2)

        # Shared utility bar (OAuth 2.0 JSON + Column Mapping)
        util_bar = ttk.Frame(top_box)
        util_bar.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(util_bar, textvariable=self.oauth_status_var, foreground="#1d4ed8").pack(side=tk.LEFT)
        ttk.Button(util_bar, text="🔑 导入 JSON 并授权 Google 账户", command=self._open_oauth_dialog).pack(
            side=tk.LEFT, padx=(10, 6)
        )
        ttk.Button(util_bar, text="⚙️ 自定义通读表列名", command=self._open_col_dialog).pack(side=tk.LEFT)

        # Sub-notebook for 3A & 3B
        self.sub_nb = ttk.Notebook(top_box)
        self.sub_nb.pack(fill=tk.BOTH, expand=True)

        tab_3a = ttk.Frame(self.sub_nb, padding=8)
        tab_3b = ttk.Frame(self.sub_nb, padding=8)
        self.sub_nb.add(tab_3a, text="  3A. 改点返稿入库（发稿 + 回稿）  ")
        self.sub_nb.add(tab_3b, text="  3B. 通读定稿入库（Google Sheets / Excel）  ")

        self._build_3a_tab(tab_3a)
        self._build_3b_tab(tab_3b)

        # Bottom Progress + Log Panel
        prog_row = ttk.Frame(bot_box)
        prog_row.pack(fill=tk.X, pady=(6, 4))
        ttk.Label(prog_row, textvariable=self.stage_var, foreground="#334155").pack(side=tk.LEFT)
        self.open_out_btn = ttk.Button(prog_row, text="📂 打开输出目录", command=self._open_last_output)
        self.open_out_btn.pack(side=tk.RIGHT)

        ttk.Progressbar(bot_box, variable=self.progress_var, maximum=100.0).pack(fill=tk.X, pady=(0, 4))

        self.log_box = scrolledtext.ScrolledText(
            bot_box,
            height=8,
            font=(_MONO_FONT, _PT_9),
            wrap=tk.WORD,
            state=tk.DISABLED,
        )
        self.log_box.pack(fill=tk.BOTH, expand=True)

    def _build_3a_tab(self, parent: ttk.Frame):
        parent.columnconfigure(1, weight=1)

        self.drop_banner_3a = _make_drop_banner(
            parent,
            "📥 支持直接将「发稿 / 回稿 / 总库文件夹 / 输出文件夹 / Excel 通读表」拖入下方任意对应输入框（或直接拖到此处智能识别）",
            lambda: self._pick_file(self.sent_path_var, [("Word 文档", "*.docx")]),
        )
        self.drop_banner_3a.grid(row=0, column=0, columnspan=3, sticky=tk.EW, pady=(0, 6))

        # Row 1: 发稿
        self.lbl_sent = ttk.Label(parent, text="发稿（拖入文件或文件夹）：")
        self.lbl_sent.grid(row=1, column=0, sticky=tk.W, pady=3)
        self.ent_sent = ttk.Entry(parent, textvariable=self.sent_path_var)
        self.ent_sent.grid(row=1, column=1, sticky=tk.EW, padx=4, pady=3)
        b_sent = ttk.Frame(parent)
        b_sent.grid(row=1, column=2, sticky=tk.W)
        ttk.Button(b_sent, text="选文件", width=7, command=lambda: self._pick_file(self.sent_path_var, [("Word", "*.docx")])).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(b_sent, text="选文件夹", width=8, command=lambda: self._pick_dir(self.sent_path_var)).pack(side=tk.LEFT)

        # Row 2: 回稿
        self.lbl_ret = ttk.Label(parent, text="回稿（拖入文件或文件夹）：")
        self.lbl_ret.grid(row=2, column=0, sticky=tk.W, pady=3)
        self.ent_ret = ttk.Entry(parent, textvariable=self.ret_path_var)
        self.ent_ret.grid(row=2, column=1, sticky=tk.EW, padx=4, pady=3)
        b_ret = ttk.Frame(parent)
        b_ret.grid(row=2, column=2, sticky=tk.W)
        ttk.Button(b_ret, text="选文件", width=7, command=lambda: self._pick_file(self.ret_path_var, [("Word", "*.docx")])).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(b_ret, text="选文件夹", width=8, command=lambda: self._pick_dir(self.ret_path_var)).pack(side=tk.LEFT)

        # Row 3: 总库目录
        self.lbl_lib = ttk.Label(parent, text="三语总库目录（只读，可拖入）：")
        self.lbl_lib.grid(row=3, column=0, sticky=tk.W, pady=3)
        self.ent_lib = ttk.Entry(parent, textvariable=self.lib_dir_var)
        self.ent_lib.grid(row=3, column=1, sticky=tk.EW, padx=4, pady=3)
        ttk.Button(parent, text="选择目录", width=16, command=lambda: self._pick_dir(self.lib_dir_var)).grid(row=3, column=2, sticky=tk.W)

        # Row 4: 输出目录
        self.lbl_out = ttk.Label(parent, text="输出目录（改前/改后，可拖入）：")
        self.lbl_out.grid(row=4, column=0, sticky=tk.W, pady=3)
        self.ent_out = ttk.Entry(parent, textvariable=self.out_dir_var)
        self.ent_out.grid(row=4, column=1, sticky=tk.EW, padx=4, pady=3)
        ttk.Button(parent, text="选择目录", width=16, command=lambda: self._pick_dir(self.out_dir_var)).grid(row=4, column=2, sticky=tk.W)

        # Row 5: 可选通读表
        self.pr_lbl_row_3a = ttk.Frame(parent)
        self.pr_lbl_row_3a.grid(row=5, column=0, columnspan=3, sticky=tk.EW, pady=(4, 2))
        ttk.Label(self.pr_lbl_row_3a, text="可选：通读定稿表（粘贴 Google Sheets 链接，或直接将本地 .xlsx/.csv 拖入下方框内）：").pack(side=tk.LEFT)
        ttk.Button(self.pr_lbl_row_3a, text="+ 添加本地 Excel", command=self._add_3a_xlsx).pack(side=tk.RIGHT)

        self.a3_proofread_txt = scrolledtext.ScrolledText(parent, height=3, font=(_MONO_FONT, _PT_9))
        self.a3_proofread_txt.grid(row=6, column=0, columnspan=3, sticky=tk.EW, pady=(0, 4))

        # Row 7: Conflict policy
        opt1 = ttk.Frame(parent)
        opt1.grid(row=7, column=0, columnspan=3, sticky=tk.EW, pady=2)
        ttk.Label(opt1, text="发稿与总库不一致时：").pack(side=tk.LEFT)
        ttk.Radiobutton(
            opt1, text="① 保留总库改动（仅入库其余一致部分）", value="keep_library", variable=self.conflict_policy_var, command=self._save_app_settings
        ).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Radiobutton(
            opt1, text="② 跳过该不一致段落（仅入库完全一致段落）", value="skip_paragraph", variable=self.conflict_policy_var, command=self._save_app_settings
        ).pack(side=tk.LEFT)

        # Row 8: Revision aggregation mode
        opt2 = ttk.Frame(parent)
        opt2.grid(row=8, column=0, columnspan=3, sticky=tk.EW, pady=2)
        ttk.Label(opt2, text="改前后修订稿汇总：").pack(side=tk.LEFT)
        ttk.Radiobutton(
            opt2, text="全部汇总为 1 份", value="single", variable=self.rev_merge_mode_var, command=self._save_app_settings
        ).pack(side=tk.LEFT, padx=(4, 10))
        ttk.Radiobutton(
            opt2, text="分批次汇总，每批", value="batch", variable=self.rev_merge_mode_var, command=self._save_app_settings
        ).pack(side=tk.LEFT)
        ttk.Spinbox(opt2, from_=2, to=100, width=4, textvariable=self.rev_batch_size_var).pack(side=tk.LEFT, padx=3)
        ttk.Label(opt2, text="篇").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(
            opt2, text="逐篇单独输出", value="per_article", variable=self.rev_merge_mode_var, command=self._save_app_settings
        ).pack(side=tk.LEFT)

        # Row 9: Action buttons
        act_row = ttk.Frame(parent)
        act_row.grid(row=9, column=0, columnspan=3, sticky=tk.EW, pady=(8, 0))
        self.btn_3a_check = ttk.Button(act_row, text="🔍 仅执行验收检查（输出报告与带批注痕迹稿）", command=lambda: self._run_3a(check_only=True))
        self.btn_3a_check.pack(side=tk.LEFT, padx=(0, 10))
        self.btn_3a_run = ttk.Button(act_row, text="🚀 一键验收并入库（生成 改前 / 改后 / 修订稿）", command=lambda: self._run_3a(check_only=False))
        self.btn_3a_run.pack(side=tk.LEFT)

    def _build_3b_tab(self, parent: ttk.Frame):
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(4, weight=1)

        self.drop_banner_3b = _make_drop_banner(
            parent,
            "📥 支持直接拖入「三语总库文件夹 / 输出文件夹 / 本地通读 .xlsx/.csv / 授权 .json 文件」自动载入",
            self._add_3b_local_sheet,
        )
        self.drop_banner_3b.grid(row=0, column=0, columnspan=3, sticky=tk.EW, pady=(0, 6))

        self.lbl_pr_lib = ttk.Label(parent, text="三语总库目录（只读，可拖入）：")
        self.lbl_pr_lib.grid(row=1, column=0, sticky=tk.W, pady=3)
        self.ent_pr_lib = ttk.Entry(parent, textvariable=self.pr_lib_dir_var)
        self.ent_pr_lib.grid(row=1, column=1, sticky=tk.EW, padx=4, pady=3)
        ttk.Button(parent, text="选择目录", width=14, command=lambda: self._pick_dir(self.pr_lib_dir_var)).grid(row=1, column=2, sticky=tk.W)

        self.lbl_pr_out = ttk.Label(parent, text="输出目录（改前/改后，可拖入）：")
        self.lbl_pr_out.grid(row=2, column=0, sticky=tk.W, pady=3)
        self.ent_pr_out = ttk.Entry(parent, textvariable=self.pr_out_dir_var)
        self.ent_pr_out.grid(row=2, column=1, sticky=tk.EW, padx=4, pady=3)
        ttk.Button(parent, text="选择目录", width=14, command=lambda: self._pick_dir(self.pr_out_dir_var)).grid(row=2, column=2, sticky=tk.W)

        self.lbl_row_3b = ttk.Frame(parent)
        self.lbl_row_3b.grid(row=3, column=0, columnspan=3, sticky=tk.EW, pady=(4, 2))
        ttk.Label(
            self.lbl_row_3b,
            text="粘贴「左列篇目标题 + 右列通读表格链接」（支持多行 Part 1..N，或直接将 .xlsx/.csv/.txt 拖入下方框内）：",
        ).pack(side=tk.LEFT)
        ttk.Button(self.lbl_row_3b, text="+ 添加本地 Excel/CSV", command=self._add_3b_local_sheet).pack(side=tk.RIGHT)

        self.b3_links_txt = scrolledtext.ScrolledText(parent, height=6, font=(_MONO_FONT, _PT_9))
        self.b3_links_txt.grid(row=4, column=0, columnspan=3, sticky=tk.NSEW, pady=(0, 6))

        act_row = ttk.Frame(parent)
        act_row.grid(row=5, column=0, columnspan=3, sticky=tk.EW)
        ttk.Checkbutton(
            act_row,
            text="篇幅/字词大幅变动时调用 AI 核对中文语义一致性（使用顶部 API Key）",
            variable=self.pr_enable_ai_var,
            command=self._save_app_settings,
        ).pack(side=tk.LEFT, padx=(0, 12))
        self.btn_3b_run = ttk.Button(
            act_row,
            text="🚀 开始通读定稿入库（提取 → 质检 → 精确替换 → 全篇核对）",
            command=self._run_3b,
        )
        self.btn_3b_run.pack(side=tk.RIGHT)

    # ---------------- Native Drag & Drop Registration & Smart Routing ----------------

    def register_dnd_targets(self, dnd: WinFileDropManager):
        """Register every input widget in Sub-tab 3A & 3B with the root window's WinFileDropManager."""
        # 3A specific rows
        for w in (self.lbl_sent, self.ent_sent):
            dnd.register_target(w, lambda ps: self._drop_file_or_dir_into_var(ps, self.sent_path_var, "发稿"))
        for w in (self.lbl_ret, self.ent_ret):
            dnd.register_target(w, lambda ps: self._drop_file_or_dir_into_var(ps, self.ret_path_var, "回稿"))
        for w in (self.lbl_lib, self.ent_lib):
            dnd.register_target(w, lambda ps: self._drop_dir_into_var(ps, self.lib_dir_var, "三语总库目录"))
        for w in (self.lbl_out, self.ent_out):
            dnd.register_target(w, lambda ps: self._drop_dir_into_var(ps, self.out_dir_var, "输出目录"))
        for w in (self.pr_lbl_row_3a, self.a3_proofread_txt):
            dnd.register_target(w, self._drop_into_3a_proofread)

        # 3B specific rows
        for w in (self.lbl_pr_lib, self.ent_pr_lib):
            dnd.register_target(w, lambda ps: self._drop_dir_into_var(ps, self.pr_lib_dir_var, "三语总库目录"))
        for w in (self.lbl_pr_out, self.ent_pr_out):
            dnd.register_target(w, lambda ps: self._drop_dir_into_var(ps, self.pr_out_dir_var, "输出目录"))
        for w in (self.lbl_row_3b, self.b3_links_txt):
            dnd.register_target(w, self._drop_into_3b_links)

        # Top banners
        dnd.register_target(self.drop_banner_3a, self._handle_3a_smart_drop)
        dnd.register_target(self.drop_banner_3b, self._handle_3b_smart_drop)

    def handle_fallback_drop(self, paths: Sequence[Path]):
        """Route drops anywhere else on the active Scenario 3 sub-tab."""
        # If user drops a Google OAuth `.json` anywhere on this tab, import it directly!
        json_files = [p for p in paths if p.is_file() and p.suffix.lower() == ".json"]
        if json_files:
            try:
                info = self.oauth.import_json_file(json_files[0])
                self._refresh_oauth_label()
                self._save_app_settings()
                self._append_log(f"🔑 已直接拖入并导入 Google 授权 JSON：{info['filename']}")
                self._open_oauth_dialog()
            except Exception as ex:
                messagebox.showerror("导入 JSON 授权文件失败", str(ex))
            return

        cur_sub = self.sub_nb.index(self.sub_nb.select())
        if cur_sub == 0:
            self._handle_3a_smart_drop(paths)
        else:
            self._handle_3b_smart_drop(paths)

    def _drop_file_or_dir_into_var(self, paths: Sequence[Path], var: tk.StringVar, label: str):
        for p in paths:
            if p.is_dir() or (p.is_file() and p.suffix.lower() == ".docx"):
                var.set(str(p.resolve()))
                self._save_app_settings()
                self._append_log(f"📥 已拖入{label}：{p.name}")
                return
        messagebox.showwarning("提示", f"请将 .docx 文档或文件夹拖入「{label}」栏位。")

    def _drop_dir_into_var(self, paths: Sequence[Path], var: tk.StringVar, label: str):
        for p in paths:
            target = p.resolve() if p.is_dir() else p.resolve().parent
            if target.is_dir():
                var.set(str(target))
                self._save_app_settings()
                self._append_log(f"📥 已拖入{label}：{target}")
                return

    def _drop_into_3a_proofread(self, paths: Sequence[Path]):
        added = 0
        for p in paths:
            if p.is_file() and p.suffix.lower() in (".xlsx", ".csv"):
                self.a3_proofread_txt.insert(tk.END, f"{p.resolve()}\n")
                added += 1
            elif p.is_dir():
                for sub in sorted(p.glob("*.xlsx")):
                    if not sub.name.startswith("~$"):
                        self.a3_proofread_txt.insert(tk.END, f"{sub.resolve()}\n")
                        added += 1
        if added:
            self._append_log(f"📥 已拖入 {added} 个通读表格文件至 3A 通读表栏位。")

    def _drop_into_3b_links(self, paths: Sequence[Path]):
        added = 0
        for p in paths:
            rp = p.resolve()
            if rp.is_file() and rp.suffix.lower() in (".xlsx", ".csv"):
                self.b3_links_txt.insert(tk.END, f"{rp.stem}\t{rp}\n")
                added += 1
            elif rp.is_file() and rp.suffix.lower() == ".txt":
                txt = rp.read_text(encoding="utf-8-sig", errors="replace").strip()
                if txt:
                    self.b3_links_txt.insert(tk.END, txt + "\n")
                    added += 1
            elif rp.is_dir():
                for sub in sorted(rp.glob("*.xlsx")):
                    if not sub.name.startswith("~$"):
                        self.b3_links_txt.insert(tk.END, f"{sub.stem}\t{sub.resolve()}\n")
                        added += 1
        if added:
            self._append_log(f"📥 已载入 {added} 项通读表格内容至 3B 列表。")

    def _handle_3a_smart_drop(self, paths: Sequence[Path]):
        sheets = [p for p in paths if p.is_file() and p.suffix.lower() in (".xlsx", ".csv")]
        if sheets:
            self._drop_into_3a_proofread(sheets)

        others = [p for p in paths if p not in sheets]
        for p in others:
            rp = p.resolve()
            nm = rp.name.lower()
            if "发稿" in nm:
                self.sent_path_var.set(str(rp))
                self._append_log(f"📥 智能识别并填入「发稿」：{rp.name}")
            elif any(k in nm for k in ("回稿", "净稿", "返稿")):
                self.ret_path_var.set(str(rp))
                self._append_log(f"📥 智能识别并填入「回稿」：{rp.name}")
            elif rp.is_dir() and any(k in nm for k in ("神话", "总库", "七卷", "定稿库")):
                self.lib_dir_var.set(str(rp))
                self._append_log(f"📥 智能识别并填入「三语总库目录」：{rp.name}")
            elif rp.is_dir() and any(k in nm for k in ("输出", "output", "结果")):
                self.out_dir_var.set(str(rp))
                self._append_log(f"📥 智能识别并填入「输出目录」：{rp.name}")
            else:
                # Fill first empty slot in order: 发稿 -> 回稿 -> 总库 -> 输出目录
                if not self.sent_path_var.get().strip():
                    self.sent_path_var.set(str(rp))
                    self._append_log(f"📥 已填入「发稿」：{rp.name}")
                elif not self.ret_path_var.get().strip():
                    self.ret_path_var.set(str(rp))
                    self._append_log(f"📥 已填入「回稿」：{rp.name}")
                elif rp.is_dir() and not self.lib_dir_var.get().strip():
                    self.lib_dir_var.set(str(rp))
                    self._append_log(f"📥 已填入「三语总库目录」：{rp.name}")
                elif rp.is_dir() and not self.out_dir_var.get().strip():
                    self.out_dir_var.set(str(rp))
                    self._append_log(f"📥 已填入「输出目录」：{rp.name}")
                else:
                    self.sent_path_var.set(str(rp))
                    self._append_log(f"📥 已更新「发稿」：{rp.name}")
        self._save_app_settings()

    def _handle_3b_smart_drop(self, paths: Sequence[Path]):
        for p in paths:
            rp = p.resolve()
            if rp.is_file() and rp.suffix.lower() in (".xlsx", ".csv", ".txt"):
                self._drop_into_3b_links([rp])
            elif rp.is_dir():
                nm = rp.name.lower()
                if any(k in nm for k in ("输出", "output", "结果")):
                    self.pr_out_dir_var.set(str(rp))
                    self._append_log(f"📥 已填入 3B「输出目录」：{rp.name}")
                elif not self.pr_lib_dir_var.get().strip() or any(k in nm for k in ("神话", "总库", "七卷")):
                    self.pr_lib_dir_var.set(str(rp))
                    self._append_log(f"📥 已填入 3B「三语总库目录」：{rp.name}")
                else:
                    self.pr_out_dir_var.set(str(rp))
                    self._append_log(f"📥 已填入 3B「输出目录」：{rp.name}")
        self._save_app_settings()

    # ---------------- UI Helpers ----------------

    def _pick_file(self, var: tk.StringVar, filetypes):
        p = filedialog.askopenfilename(filetypes=filetypes)
        if p:
            var.set(p)
            self._save_app_settings()

    def _pick_dir(self, var: tk.StringVar):
        p = filedialog.askdirectory()
        if p:
            var.set(p)
            self._save_app_settings()

    def _add_3a_xlsx(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel / CSV", "*.xlsx *.csv")])
        for p in paths:
            self.a3_proofread_txt.insert(tk.END, f"{p}\n")

    def _add_3b_local_sheet(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel / CSV", "*.xlsx *.csv")])
        for p in paths:
            stem = Path(p).stem
            self.b3_links_txt.insert(tk.END, f"{stem}\t{p}\n")

    def _open_oauth_dialog(self):
        def _after():
            self._refresh_oauth_label()
            self._save_app_settings()

        GoogleOAuthDialog(self.parent.winfo_toplevel(), self.oauth, on_updated=_after)

    def _open_col_dialog(self):
        ColumnMappingDialog(self.parent.winfo_toplevel(), self.col_cfg, on_save=self._save_app_settings)

    def _open_last_output(self):
        target = self._last_out_dir or self.out_dir_var.get().strip() or self.pr_out_dir_var.get().strip()
        if target and Path(target).is_dir():
            open_path_in_os(target)
        else:
            messagebox.showinfo("提示", "请先选择或生成有效的输出目录。")

    def _append_log(self, msg: str):
        def _do():
            self.log_box.configure(state=tk.NORMAL)
            self.log_box.insert(tk.END, msg + "\n")
            self.log_box.see(tk.END)
            self.log_box.configure(state=tk.DISABLED)

        self.parent.after(0, _do)

    def _set_progress(self, pct: float, text: str):
        def _do():
            self.progress_var.set(max(0.0, min(100.0, float(pct))))
            self.stage_var.set(f"状态：{text}")

        self.parent.after(0, _do)

    def _set_busy(self, busy: bool):
        self._running = busy
        st = tk.DISABLED if busy else tk.NORMAL
        for btn in (self.btn_3a_check, self.btn_3a_run, self.btn_3b_run):
            btn.configure(state=st)

    # ---------------- Execution Handlers ----------------

    def _run_3a(self, check_only: bool = False):
        if self._running:
            return
        sent_p = self.sent_path_var.get().strip()
        ret_p = self.ret_path_var.get().strip()
        lib_d = self.lib_dir_var.get().strip()
        out_d = self.out_dir_var.get().strip()
        if not sent_p or not Path(sent_p).exists():
            messagebox.showwarning("提示", "请选择或拖入有效的「发稿」文件或文件夹。")
            return
        if not ret_p or not Path(ret_p).exists():
            messagebox.showwarning("提示", "请选择或拖入有效的「回稿」文件或文件夹。")
            return
        if not lib_d or not Path(lib_d).is_dir():
            messagebox.showwarning("提示", "请选择或拖入有效的「三语总库目录」。")
            return
        if not out_d:
            messagebox.showwarning("提示", "请选择或拖入「输出目录」。")
            return

        self._save_app_settings()
        self._last_out_dir = out_d
        pr_sources = [ln.strip() for ln in self.a3_proofread_txt.get("1.0", tk.END).splitlines() if ln.strip()]
        keys = self.app._get_all_keys_list() if hasattr(self.app, "_get_all_keys_list") else []
        model = self.app._get_model_id() if hasattr(self.app, "_get_model_id") else "gemini-3.5-flash-lite"
        api_mode = self.app._get_api_mode_code() if hasattr(self.app, "_get_api_mode_code") else "auto"
        rot_mode = self.app._get_rotation_mode_code() if hasattr(self.app, "_get_rotation_mode_code") else "round_robin"
        pool = ApiKeyPool(keys, rotation_mode=rot_mode) if (pr_sources and keys) else ""
        try:
            bsz = max(1, int(self.rev_batch_size_var.get().strip() or "10"))
        except ValueError:
            bsz = 10

        self._set_busy(True)
        self._append_log("=" * 64)
        self._append_log(f"🚀 启动【场景 3A：{'仅验收检查' if check_only else '改点返稿入库'}】")

        def _worker():
            try:
                res = run_library_ingestion_pipeline(
                    sent_path=sent_p,
                    returned_path=ret_p,
                    library_dir=lib_d,
                    output_dir=out_d,
                    conflict_policy=self.conflict_policy_var.get(),
                    revision_merge_mode=self.rev_merge_mode_var.get(),
                    revision_batch_size=bsz,
                    check_only=check_only,
                    proofread_sources=pr_sources,
                    col_config=self.col_cfg,
                    oauth=self.oauth,
                    api_key=pool,
                    model=model,
                    api_mode=api_mode,
                    log_fn=self._append_log,
                    progress_fn=self._set_progress,
                )
                pr_extra = (
                    f"\n通读定稿表：精确替换 {res.get('proofread_applied', 0)} 条 | 需复核/未替换汇报 {res.get('proofread_report_count', 0)} 条"
                    if pr_sources
                    else ""
                )
                summary = (
                    f"完成！总库不一致 {res['mismatches_count']} 处，回稿异常 {res['anomalies_count']} 处。{pr_extra}"
                    if check_only
                    else f"入库完成！已更新 {res['updated_chapters']} 篇（{res['updated_rows']} 段），原总库未作任何改动。\n"
                         f"总库不一致：{res['mismatches_count']} 处 | 回稿异常：{res['anomalies_count']} 处{pr_extra}"
                )
                self.parent.after(0, lambda: messagebox.showinfo("改点入库完成", summary))
            except Exception as ex:
                self._append_log(f"❌ 执行失败：{ex}")
                self.parent.after(0, lambda e=str(ex): messagebox.showerror("执行失败", e))
            finally:
                self.parent.after(0, lambda: self._set_busy(False))

        threading.Thread(target=_worker, daemon=True).start()

    def _run_3b(self):
        if self._running:
            return
        lib_d = self.pr_lib_dir_var.get().strip()
        out_d = self.pr_out_dir_var.get().strip()
        links_txt = self.b3_links_txt.get("1.0", tk.END).strip()
        if not lib_d or not Path(lib_d).is_dir():
            messagebox.showwarning("提示", "请选择或拖入有效的「三语总库目录」。")
            return
        if not out_d:
            messagebox.showwarning("提示", "请选择或拖入「输出目录」。")
            return
        if not links_txt:
            messagebox.showwarning("提示", "请先粘贴「篇目标题 + 通读表格链接」（或拖入本地 .xlsx 文件）。")
            return

        self._save_app_settings()
        self._last_out_dir = out_d

        want_ai = bool(self.pr_enable_ai_var.get())
        keys = self.app._get_all_keys_list() if hasattr(self.app, "_get_all_keys_list") else []
        model = self.app._get_model_id() if hasattr(self.app, "_get_model_id") else "gemini-3.5-flash-lite"
        api_mode = self.app._get_api_mode_code() if hasattr(self.app, "_get_api_mode_code") else "auto"
        rot_mode = self.app._get_rotation_mode_code() if hasattr(self.app, "_get_rotation_mode_code") else "round_robin"
        pool = ApiKeyPool(keys, rotation_mode=rot_mode) if (want_ai and keys) else ""

        self._set_busy(True)
        self._append_log("=" * 64)
        self._append_log("🚀 启动【场景 3B：通读定稿入库】")

        def _worker():
            try:
                res = run_proofread_ingestion_pipeline(
                    title_links_text=links_txt,
                    library_dir=lib_d,
                    output_dir=out_d,
                    col_config=self.col_cfg,
                    oauth=self.oauth,
                    enable_ai=bool(want_ai and keys),
                    api_key=pool,
                    model=model,
                    api_mode=api_mode,
                    log_fn=self._append_log,
                    progress_fn=self._set_progress,
                )
                msg = (
                    f"通读定稿入库完成！\n"
                    f"成功处理：{res['processed_chapters']} 篇 | 未匹配：{len(res['unmatched_titles'])} 篇\n"
                    f"质检提醒：{res['qa_alerts_count']} 条 | 替换提醒：{res['replace_alerts_count']} 条\n"
                    f"全篇合并核对差异修订稿：{len(res['verify_revision_files'])} 份"
                )
                self.parent.after(0, lambda: messagebox.showinfo("通读定稿入库完成", msg))
            except Exception as ex:
                self._append_log(f"❌ 执行失败：{ex}")
                self.parent.after(0, lambda e=str(ex): messagebox.showerror("执行失败", e))
            finally:
                self.parent.after(0, lambda: self._set_busy(False))

        threading.Thread(target=_worker, daemon=True).start()
