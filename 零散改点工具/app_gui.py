from __future__ import annotations

import ctypes
from datetime import datetime
import inspect
import json
import os
import queue
import re
import sys
import threading
import time
import traceback
import webbrowser
from pathlib import Path

if "--self-test" in sys.argv:
    def _headless_self_test_excepthook(exc_type, exc_value, exc_tb):
        tb_txt = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        log_p = os.environ.get("SCATTERED_SELF_TEST_LOG", "").strip()
        if log_p:
            try:
                Path(log_p).write_text(tb_txt, encoding="utf-8")
            except OSError:
                pass
        try:
            sys.stderr.write(tb_txt)
            sys.stderr.flush()
        except Exception:
            pass
        os._exit(1)

    sys.excepthook = _headless_self_test_excepthook
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import xml.etree.ElementTree as ET
import zipfile

from ai_client import (
    API_MODES,
    ROTATION_MODES,
    SUPPORTED_MODELS,
    ApiKeyPool,
    test_all_api_keys,
    test_api_connection,
)
import ai_client as _ai_client_mod

fetch_available_models = getattr(_ai_client_mod, "fetch_available_models", None)

try:
    from ai_client import format_friendly_api_error
except ImportError:
    def format_friendly_api_error(ex, total_keys: int = 1) -> str:  # type: ignore[misc]
        lines = str(ex).strip().splitlines()
        return lines[0] if lines else str(ex)

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
import lang_registry as _lang_reg_mod
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

get_lang_typography = getattr(_lang_reg_mod, "get_lang_typography", None)
_ext_get_prioritized_combo_values = getattr(_lang_reg_mod, "get_prioritized_combo_values", None)
_ext_record_recent_language = getattr(_lang_reg_mod, "record_recent_language", None)

import pipeline as _pipeline_mod

run_oneclick_pipeline = getattr(_pipeline_mod, "run_oneclick_pipeline", None)
resolve_sc1_input_docx_list = getattr(_pipeline_mod, "resolve_sc1_input_docx_list", None)
ensure_library_index = getattr(_pipeline_mod, "ensure_library_index", None)

try:
    import master_sync as _master_sync_mod
except ImportError:
    _master_sync_mod = None

sync_pending_with_latest_master = getattr(_master_sync_mod, "sync_pending_with_latest_master", None)

try:
    import full_article_processor as _fa_mod
except ImportError:
    _fa_mod = None

parse_titles_text_or_file = getattr(_fa_mod, "parse_titles_text_or_file", None)
parse_titles_multiline = getattr(_fa_mod, "parse_titles_multiline", None)
refresh_master_index = getattr(_fa_mod, "refresh_master_index", None)
resolve_full_articles = getattr(_fa_mod, "resolve_full_articles", None)
resolve_direct_folder_articles = getattr(_fa_mod, "resolve_direct_folder_articles", None)
run_full_article_pipeline = getattr(_fa_mod, "run_full_article_pipeline", None)

try:
    import dedup_checker as _dd_mod
except ImportError:
    _dd_mod = None

check_and_export_duplicates = getattr(_dd_mod, "check_and_export_duplicates", None)

try:
    import ingest_tab_ui as _ingest_ui_mod
except ImportError:
    _ingest_ui_mod = None

IngestTabPanel = getattr(_ingest_ui_mod, "IngestTabPanel", None)

try:
    import updater as _updater_mod
except ImportError:
    _updater_mod = None

RAW_APP_VERSION = "1.9.7"
APP_VERSION = f"v{RAW_APP_VERSION}"
check_for_updates = getattr(_updater_mod, "check_for_updates", None)
download_and_apply_update = getattr(_updater_mod, "download_and_apply_update", None)
cleanup_stale_update_backups = getattr(_updater_mod, "cleanup_stale_update_backups", None)
run_finish_update_mode = getattr(_updater_mod, "run_finish_update_mode", None)
GITHUB_RELEASES_URL = "https://github.com/zcnz2026/TranslationTools/releases"

from security_utils import (
    cleanup_stale_ephemeral_workdirs,
    dpapi_decrypt_str,
    dpapi_encrypt_str,
    get_user_app_data_dir,
    load_api_key_from_env,
    open_path_in_os,
    parse_api_keys,
    redact_secrets,
)
from word_counter import (
    analyze_docx_word_count,
    apply_word_count_filename,
    format_counted_docx_name,
    parse_doc_filename,
)
from win_dnd import WinFileDropManager, create_tk_root

UI_FONT_FAMILY = "PingFang SC" if sys.platform == "darwin" else "Microsoft YaHei UI"
MONO_FONT_FAMILY = "Menlo" if sys.platform == "darwin" else "Consolas"


def _ui_pt(pt: int) -> int:
    return pt + 3 if sys.platform == "darwin" else pt


_BUILTIN_FONT_FALLBACKS: dict[str, str] = {
    "FR": "Georgia",
    "PT": "Georgia",
    "ES": "Georgia",
    "IT": "Georgia",
    "TL": "Georgia",
    "RO": "Georgia",
    "BG": "Georgia",
    "PL": "Georgia",
    "PO": "Georgia",
    "GR": "Georgia",
    "SR": "Georgia",
    "HR": "Georgia",
    "NL": "Georgia",
    "HMN": "Georgia",
    "SW": "Georgia",
    "RU": "Times New Roman",
    "ID": "Times New Roman",
    "VI": "Times New Roman",
    "HU": "Times New Roman",
    "MN": "Times New Roman",
    "DE": "Times New Roman",
    "SK": "Times New Roman",
    "UK": "Arial",
    "HI": "Mangal",
    "NE": "Mangal",
    "MY": "Myanmar Text",
    "AR": "Simplified Arabic",
    "TH": "Angsana New",
    "KO": "Malgun Gothic",
    "HE": "FrankRuehl",
    "KM": "Leelawadee UI",
    "LO": "DokChampa",
    "HY": "Noto Sans Armenian",
    "KA": "Sylfaen",
}

_COMMON_STANDARD_FONTS: list[str] = [
    "",
    "Georgia",
    "Times New Roman",
    "Arial",
    "Angsana New",
    "Mangal",
    "Myanmar Text",
    "Simplified Arabic",
    "Malgun Gothic",
    "FrankRuehl",
    "Leelawadee UI",
    "DokChampa",
    "Noto Sans Armenian",
    "Sylfaen",
]


def _call_compatible(fn, *args, **kwargs):
    """Call `fn` forwarding all supported keyword arguments cleanly across module versions."""
    try:
        sig = inspect.signature(fn)
        has_var_kw = any(p.kind == inspect.Parameter.VAR_KEYWORD for p in sig.parameters.values())
        if has_var_kw:
            return fn(*args, **kwargs)
        filtered_kw = {k: v for k, v in kwargs.items() if k in sig.parameters}
        return fn(*args, **filtered_kw)
    except (ValueError, TypeError):
        return fn(*args, **kwargs)


def _resolve_combo_code(text: str, entries: list[dict]) -> str:
    """Resolve combobox label (including `★ ` / `⏱ ` prefixes) cleanly to canonical language code."""
    cleaned = re.sub(r"^[★☆⏱\s]+", "", str(text or "")).strip()
    return resolve_combo_to_code(cleaned, entries)


def get_prioritized_combo_values(
    entries: list[dict],
    recent_codes: list[str] | None = None,
    favorite_codes: list[str] | None = None,
) -> list[str]:
    """Return language combobox items ordered: AUTO -> ★ Favorites -> ⏱ Recent -> Remaining."""
    if callable(_ext_get_prioritized_combo_values):
        return list(
            _call_compatible(
                _ext_get_prioritized_combo_values,
                entries,
                recent_codes=recent_codes,
                favorite_codes=favorite_codes,
            )
        )
    fav_set = {str(c).strip().upper() for c in (favorite_codes or []) if str(c).strip()}
    for e in entries:
        if e.get("favorite"):
            fav_set.add(str(e.get("code", "")).strip().upper())

    by_code = {str(e.get("code", "")).strip().upper(): e for e in entries if str(e.get("code", "")).strip()}
    out: list[str] = [AUTO_COMBO_LABEL]
    listed: set[str] = set()

    # 1. Favorites (preserving entries order)
    for e in entries:
        c = str(e.get("code", "")).strip().upper()
        if c and c in fav_set and c not in listed:
            out.append(f"★ {e['code']} - {e['name']}")
            listed.add(c)

    # 2. Recently used (MRU order, excluding favorites)
    for rc in recent_codes or []:
        c = str(rc or "").strip().upper()
        if not c or c in ("AUTO", "XX") or c in listed:
            continue
        ent = by_code.get(c)
        if ent is not None:
            out.append(f"⏱ {ent['code']} - {ent['name']}")
            listed.add(c)

    # 3. Remaining entries
    for e in entries:
        c = str(e.get("code", "")).strip().upper()
        if c and c not in listed:
            out.append(f"{e['code']} - {e['name']}")
            listed.add(c)

    return out


def record_recent_language(code: str, recent_codes: list[str] | None = None, max_items: int = 5) -> list[str]:
    """Record `code` at head of MRU list (ignoring AUTO/XX/empty), capped at `max_items`."""
    if callable(_ext_record_recent_language):
        return list(_call_compatible(_ext_record_recent_language, code, list(recent_codes or []), max_items=max_items))
    c = re.sub(r"^[★☆⏱\s]+", "", str(code or "")).strip().upper()
    existing = [
        str(x).strip().upper()
        for x in (recent_codes or [])
        if str(x).strip() and str(x).strip().upper() not in ("AUTO", "XX")
    ]
    dedup_existing: list[str] = []
    for x in existing:
        if x not in dedup_existing:
            dedup_existing.append(x)
    if not c or c in ("AUTO", "XX"):
        return dedup_existing[: max(1, int(max_items))]
    updated = [c] + [x for x in dedup_existing if x != c]
    return updated[: max(1, int(max_items))]


def _fallback_parse_titles_multiline(raw_text: str) -> list[str]:
    if callable(parse_titles_multiline):
        return list(parse_titles_multiline(raw_text))
    seen: set[str] = set()
    out: list[str] = []
    for raw_ln in (raw_text or "").splitlines():
        s = raw_ln.strip().strip("\ufeff")
        if not s:
            continue
        if s not in seen:
            seen.add(s)
            out.append(s)
    return out


def _fallback_parse_titles_file(file_path: str | Path) -> list[str]:
    if callable(parse_titles_text_or_file):
        return list(parse_titles_text_or_file(file_path))
    p = Path(file_path).resolve()
    if not p.is_file():
        raise FileNotFoundError(f"找不到标题文档：{p}")
    if p.suffix.lower() == ".docx":
        import docx_xml_utils as _dxu

        root, _, _ = _dxu.load_docx_xml(p)
        w_ns = f"{{{_dxu.W_URI}}}"
        lines: list[str] = []
        for p_el in root.iter(f"{w_ns}p"):
            txt = "".join(t.text or "" for t in p_el.iter(f"{w_ns}t")).strip()
            if txt:
                lines.append(txt)
        return _fallback_parse_titles_multiline("\n".join(lines))
    raw = p.read_text(encoding="utf-8-sig", errors="replace")
    return _fallback_parse_titles_multiline(raw)


def _clean_single_line_message(raw_obj) -> str:
    raw = str(raw_obj or "").strip()
    if not raw:
        return "未知错误"
    kept = []
    for ln in raw.splitlines():
        s = ln.strip()
        if not s:
            continue
        if s.startswith("Traceback (most recent call last)") or s.startswith('File "'):
            continue
        if s.startswith("{") or s.startswith("}") or s.startswith("]"):
            continue
        if re.match(r'^"[A-Za-z0-9_]+":', s):
            continue
        kept.append(s)
    first = kept[0] if kept else raw.splitlines()[0].strip()
    first = re.sub(r"[:：\s]*\{.*$", "", first).strip()
    return first or "处理异常，请检查文档或配置后重试"


def _format_exception_concise(
    ex: BaseException,
    keys: list[str] | None = None,
    api_related: bool = False,
) -> str:
    secret_list = keys or []
    total_keys = max(1, len(secret_list))
    if isinstance(ex, PermissionError):
        fname = getattr(ex, "filename", None)
        hint = f"（{Path(fname).name}）" if fname else ""
        return redact_secrets(
            f"文件写入受阻{hint}：目标文档可能正在 Word 中打开，请先关闭后重试。",
            extra_secret=secret_list,
        )
    if isinstance(ex, FileNotFoundError):
        return redact_secrets(_clean_single_line_message(ex), extra_secret=secret_list)

    raw = str(ex or "").strip()
    looks_api = (
        api_related
        or hasattr(ex, "http_status")
        or any(
            tok in raw
            for tok in (
                "HTTP ",
                "API Key",
                "API key",
                "api_key",
                "generativelanguage",
                "aiplatform",
                "OAuth2",
                "RESOURCE_EXHAUSTED",
                "UNAVAILABLE",
                "UNAUTHENTICATED",
                "PERMISSION_DENIED",
                "INVALID_ARGUMENT",
                "urlopen error",
                "timed out",
                "SSLError",
                "密钥池",
            )
        )
    )
    if looks_api:
        try:
            friendly = format_friendly_api_error(ex, total_keys=total_keys)
        except TypeError:
            friendly = format_friendly_api_error(ex, total_keys)
        return redact_secrets(_clean_single_line_message(friendly), extra_secret=secret_list)

    return redact_secrets(_clean_single_line_message(ex), extra_secret=secret_list)


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
        if sys.platform == "darwin":
            return get_user_app_data_dir()
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _env_search_dirs() -> list[Path]:
    dirs = [_app_base_dir()]
    if getattr(sys, "frozen", False) and sys.platform == "darwin":
        try:
            exe_p = Path(sys.executable).resolve()
            for p in exe_p.parents:
                if p.name.lower().endswith(".app") and p.parent not in dirs:
                    dirs.append(p.parent)
                    break
        except Exception:
            pass
    return dirs


SETTINGS_FILE = _app_base_dir() / "app_settings.json"


class _WinRect(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


def _get_work_area(win: tk.Misc) -> tuple[int, int, int, int]:
    """Return (left, top, width, height) of the usable desktop work area excluding the Windows taskbar."""
    if sys.platform == "win32":
        try:
            rect = _WinRect()
            ok = ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0)
            if ok and rect.right > rect.left and rect.bottom > rect.top:
                return (
                    int(rect.left),
                    int(rect.top),
                    int(rect.right - rect.left),
                    int(rect.bottom - rect.top),
                )
        except Exception:
            pass
    sw = max(800, int(win.winfo_screenwidth()))
    sh = max(600, int(win.winfo_screenheight()) - 48)
    return (0, 0, sw, sh)


def _fit_and_center_window(
    win: tk.Wm,
    base_w: int,
    base_h: int,
    min_w: int,
    min_h: int,
    parent: tk.Misc | None = None,
) -> tuple[int, int]:
    """Size `win` adaptively for DPI + required widget metrics and center within the usable work area."""
    win.update_idletasks()
    try:
        dpi = float(win.winfo_fpixels("1i"))
    except Exception:
        dpi = 96.0
    scale = max(1.0, dpi / 96.0)

    work_left, work_top, work_w, work_h = _get_work_area(win)
    avail_w = max(480, work_w - 32)
    avail_h = max(360, work_h - 48)

    req_w = int(win.winfo_reqwidth())
    req_h = int(win.winfo_reqheight())

    w = min(avail_w, max(int(round(base_w * scale)), req_w + 24))
    h = min(avail_h, max(int(round(base_h * scale)), req_h + 20))

    eff_min_w = min(w, min(avail_w, max(int(round(min_w * scale)), min(avail_w, req_w))))
    eff_min_h = min(h, min(avail_h, int(round(min_h * scale))))
    win.minsize(eff_min_w, eff_min_h)

    if parent is not None and parent.winfo_exists():
        try:
            parent.update_idletasks()
            pw = int(parent.winfo_width())
            ph = int(parent.winfo_height())
            if pw > 120 and ph > 120:
                px = int(parent.winfo_rootx()) + (pw - w) // 2
                py = int(parent.winfo_rooty()) + (ph - h) // 2
            else:
                px = work_left + (work_w - w) // 2
                py = work_top + max(8, (avail_h - h) // 2)
        except Exception:
            px = work_left + (work_w - w) // 2
            py = work_top + max(8, (avail_h - h) // 2)
    else:
        px = work_left + (work_w - w) // 2
        py = work_top + max(8, (avail_h - h) // 2)

    max_x = max(work_left + 8, work_left + work_w - w - 16)
    max_y = max(work_top + 8, work_top + work_h - h - 44)
    px = max(work_left + 8, min(px, max_x))
    py = max(work_top + 8, min(py, max_y))

    win.geometry(f"{w}x{h}+{px}+{py}")
    win.update_idletasks()
    return w, h


if sys.platform == "darwin":

    class PrimaryButton(tk.Label):
        """macOS Aqua-safe primary CTA button (tk.Label avoids Aqua white-bezel white-on-white text bug)."""

        _BG_NORMAL = "#2563eb"
        _BG_HOVER = "#1d4ed8"
        _BG_ACTIVE = "#1e40af"
        _BG_DISABLED = "#94a3b8"
        _FG_NORMAL = "#ffffff"
        _FG_DISABLED = "#f1f5f9"

        def __init__(
            self,
            master: tk.Misc,
            text: str = "",
            command=None,
            padx: int = 18,
            pady: int = 6,
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
            state: str = tk.NORMAL,
            **kwargs,
        ):
            self._command = command
            self._btn_state = tk.DISABLED if str(state).lower() == "disabled" else tk.NORMAL
            self._hovered = False
            self._pressed = False
            for unsupported_key in ("default", "overrelief", "repeatdelay", "repeatinterval"):
                kwargs.pop(unsupported_key, None)
            is_dis = self._btn_state == tk.DISABLED
            init_bg = self._BG_DISABLED if is_dis else self._BG_NORMAL
            init_fg = self._FG_DISABLED if is_dis else self._FG_NORMAL
            super().__init__(
                master,
                text=text,
                font=font,
                bg=init_bg,
                fg=init_fg,
                relief=tk.FLAT,
                bd=0,
                padx=padx,
                pady=pady,
                cursor="arrow" if is_dis else "pointinghand",
                highlightthickness=1,
                highlightbackground=init_bg,
                highlightcolor="#93c5fd",
                takefocus=0 if is_dis else 1,
                **kwargs,
            )
            self.bind("<Enter>", self._on_enter, add="+")
            self.bind("<Leave>", self._on_leave, add="+")
            self.bind("<ButtonPress-1>", self._on_press, add="+")
            self.bind("<ButtonRelease-1>", self._on_release, add="+")
            self.bind("<Return>", self._on_key_activate, add="+")
            self.bind("<space>", self._on_key_activate, add="+")

        def _is_disabled(self) -> bool:
            return str(self._btn_state).lower() == "disabled"

        def _sync_visual_state(self):
            if self._is_disabled():
                super().configure(
                    bg=self._BG_DISABLED,
                    fg=self._FG_DISABLED,
                    highlightbackground=self._BG_DISABLED,
                    cursor="arrow",
                    takefocus=0,
                )
            else:
                if self._pressed and self._hovered:
                    target_bg = self._BG_ACTIVE
                elif self._hovered:
                    target_bg = self._BG_HOVER
                else:
                    target_bg = self._BG_NORMAL
                super().configure(
                    bg=target_bg,
                    fg=self._FG_NORMAL,
                    highlightbackground=target_bg,
                    cursor="pointinghand",
                    takefocus=1,
                )

        def _on_enter(self, _event=None):
            self._hovered = True
            self._sync_visual_state()

        def _on_leave(self, _event=None):
            self._hovered = False
            self._pressed = False
            self._sync_visual_state()

        def _on_press(self, _event=None):
            if not self._is_disabled():
                self._pressed = True
                self._sync_visual_state()

        def _on_release(self, event=None):
            was_pressed = self._pressed
            self._pressed = False
            if self._is_disabled():
                self._sync_visual_state()
                return
            inside = self._hovered
            if event is not None:
                try:
                    inside = 0 <= int(event.x) <= int(self.winfo_width()) and 0 <= int(event.y) <= int(self.winfo_height())
                except Exception:
                    pass
            self._hovered = inside
            self._sync_visual_state()
            if was_pressed and inside and callable(self._command):
                self._command()

        def _on_key_activate(self, _event=None):
            if not self._is_disabled() and callable(self._command):
                self._command()

        def invoke(self):
            if not self._is_disabled() and callable(self._command):
                return self._command()
            return None

        def cget(self, key: str):
            if key == "state":
                return self._btn_state
            if key == "command":
                return self._command
            return super().cget(key)

        def configure(self, cnf=None, **kw):
            merged = {}
            if isinstance(cnf, dict):
                merged.update(cnf)
            merged.update(kw)
            state_changed = False
            if "state" in merged:
                st = merged.pop("state")
                self._btn_state = tk.DISABLED if str(st).lower() == "disabled" else tk.NORMAL
                state_changed = True
            if "command" in merged:
                self._command = merged.pop("command")
            res = super().configure(**merged) if merged else None
            if state_changed:
                self._sync_visual_state()
            return res

        config = configure

        def __setitem__(self, key: str, value):
            self.configure(**{key: value})

else:

    class PrimaryButton(tk.Button):
        """Consistent primary CTA button with crisp CJK width metrics and full hover/focus/active/disabled states."""

        _BG_NORMAL = "#2563eb"
        _BG_HOVER = "#1d4ed8"
        _BG_ACTIVE = "#1e40af"
        _BG_DISABLED = "#94a3b8"
        _FG_NORMAL = "#ffffff"
        _FG_DISABLED = "#f1f5f9"

        def __init__(
            self,
            master: tk.Misc,
            text: str = "",
            command=None,
            padx: int = 18,
            pady: int = 6,
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
            state: str = tk.NORMAL,
            **kwargs,
        ):
            is_disabled = str(state).lower() == "disabled"
            init_bg = self._BG_DISABLED if is_disabled else self._BG_NORMAL
            init_cursor = "arrow" if is_disabled else "hand2"
            super().__init__(
                master,
                text=text,
                command=command,
                font=font,
                bg=init_bg,
                fg=self._FG_NORMAL,
                activebackground=self._BG_ACTIVE,
                activeforeground=self._FG_NORMAL,
                disabledforeground=self._FG_DISABLED,
                relief=tk.FLAT,
                bd=0,
                padx=padx,
                pady=pady,
                cursor=init_cursor,
                highlightthickness=1,
                highlightbackground=init_bg,
                highlightcolor="#93c5fd",
                state=state,
                **kwargs,
            )
            self._hovered = False
            self.bind("<Enter>", self._on_enter, add="+")
            self.bind("<Leave>", self._on_leave, add="+")

        def _is_disabled(self) -> bool:
            try:
                return str(self.cget("state")).lower() == "disabled"
            except Exception:
                return False

        def _sync_visual_state(self):
            if self._is_disabled():
                super().configure(
                    bg=self._BG_DISABLED,
                    highlightbackground=self._BG_DISABLED,
                    cursor="arrow",
                )
            else:
                target_bg = self._BG_HOVER if self._hovered else self._BG_NORMAL
                super().configure(
                    bg=target_bg,
                    highlightbackground=target_bg,
                    cursor="hand2",
                )

        def _on_enter(self, _event=None):
            self._hovered = True
            if not self._is_disabled():
                super().configure(bg=self._BG_HOVER, highlightbackground=self._BG_HOVER)

        def _on_leave(self, _event=None):
            self._hovered = False
            if not self._is_disabled():
                super().configure(bg=self._BG_NORMAL, highlightbackground=self._BG_NORMAL)

        def configure(self, cnf=None, **kw):
            res = super().configure(cnf, **kw)
            if (isinstance(cnf, dict) and "state" in cnf) or "state" in kw:
                self._sync_visual_state()
            return res

        config = configure

        def __setitem__(self, key: str, value):
            super().__setitem__(key, value)
            if key == "state":
                self._sync_visual_state()


def _make_primary_btn(
    parent: tk.Misc,
    text: str,
    command=None,
    padx: int = 18,
    pady: int = 6,
    font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
    state: str = tk.NORMAL,
    **kwargs,
) -> PrimaryButton:
    return PrimaryButton(
        parent,
        text=text,
        command=command,
        padx=padx,
        pady=pady,
        font=font,
        state=state,
        **kwargs,
    )


def _make_vertical_paned(
    parent: tk.Misc,
    upper_weight: int = 0,
    lower_weight: int = 1,
    min_top: int = 140,
    min_bottom: int = 120,
) -> tuple[ttk.PanedWindow, ttk.Frame, ttk.Frame]:
    """Create a vertical `ttk.PanedWindow` with DPI-aware minimum top/bottom drag clamping."""
    try:
        scale = max(1.0, float(parent.winfo_fpixels("1i")) / 96.0)
    except Exception:
        scale = 1.0
    eff_min_top = int(round(min_top * scale))
    eff_min_bot = int(round(min_bottom * scale))

    pw = ttk.PanedWindow(parent, orient=tk.VERTICAL)
    pw.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    upper = ttk.Frame(pw)
    lower = ttk.Frame(pw)
    pw.add(upper, weight=upper_weight)
    pw.add(lower, weight=lower_weight)

    def _clamp_sash(_event=None):
        try:
            total_h = int(pw.winfo_height())
            if total_h <= (eff_min_top + eff_min_bot + 12):
                return
            pos = int(pw.sashpos(0))
            max_pos = total_h - eff_min_bot
            if pos < eff_min_top:
                pw.sashpos(0, eff_min_top)
            elif pos > max_pos:
                pw.sashpos(0, max_pos)
        except Exception:
            pass

    pw.bind("<B1-Motion>", lambda _e: pw.after_idle(_clamp_sash), add="+")
    pw.bind("<ButtonRelease-1>", lambda _e: pw.after_idle(_clamp_sash), add="+")
    return pw, upper, lower


class ApiKeyPoolDialog(tk.Toplevel):
    """Modal dialog for managing an unlimited pool of API keys, channel/encryption settings, and health testing."""

    def __init__(
        self,
        parent: tk.Tk,
        initial_raw_text: str = "",
        rotation_label_var: tk.StringVar | None = None,
        model_id: str = "gemini-3.5-flash-lite",
        api_mode: str = "auto",
        on_save_cb=None,
        api_mode_label_var: tk.StringVar | None = None,
        remember_key_var: tk.BooleanVar | None = None,
    ):
        super().__init__(parent)
        self.withdraw()
        self.title("API 密钥池管理（支持无限个 Key 轮换与报错自动切换）")
        self.transient(parent)

        self.initial_raw_text = initial_raw_text or ""
        init_rot = rotation_label_var.get() if rotation_label_var is not None else ROTATION_MODES[0][0]
        self.rotation_var = tk.StringVar(value=init_rot)
        self._ext_api_mode_var = api_mode_label_var
        self._ext_remember_var = remember_key_var

        init_mode_lbl = API_MODES[0][0]
        if api_mode_label_var is not None and api_mode_label_var.get().strip():
            init_mode_lbl = api_mode_label_var.get().strip()
        else:
            for disp, code in API_MODES:
                if api_mode == code or api_mode == disp:
                    init_mode_lbl = disp
                    break
        self.api_mode_var = tk.StringVar(value=init_mode_lbl)
        self.remember_var = tk.BooleanVar(
            value=bool(remember_key_var.get()) if remember_key_var is not None else True
        )

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
        _fit_and_center_window(self, base_w=780, base_h=580, min_w=680, min_h=480, parent=parent)
        self.deiconify()
        try:
            self.grab_set()
        except tk.TclError:
            pass
        self._poll_ui_queue()

    def _resolve_current_api_mode(self) -> str:
        lbl = self.api_mode_var.get().strip()
        for disp, code in API_MODES:
            if lbl == disp or lbl == code:
                return code
        return self.api_mode or "auto"

    def _build_ui(self):
        wrap = ttk.Frame(self, padding=14)
        wrap.pack(fill=tk.BOTH, expand=True)

        tip = ttk.Label(
            wrap,
            text=(
                "支持填入无限个 API Key（推荐每行 1 个，支持逗号/分号/竖线分隔及 # 备注，自动去重）。\n"
                "可混填 Vertex AI Key (AQ.开头) 与 Gemini Key (AIza开头)；运行中遇超额 (429) 或异常将自动切换下一个 Key。"
            ),
            foreground="#475569",
        )
        tip.pack(side=tk.TOP, fill=tk.X, pady=(0, 8))

        # Pack bottom action bar FIRST so bottom buttons NEVER clip at any window height
        btns = ttk.Frame(wrap)
        btns.pack(side=tk.BOTTOM, fill=tk.X)

        self.btn_cancel = ttk.Button(btns, text="取消", command=self.destroy)
        self.btn_cancel.pack(side=tk.RIGHT)
        self.btn_save = _make_primary_btn(btns, text="保存并应用密钥池", command=self._on_save)
        self.btn_save.pack(side=tk.RIGHT, padx=(0, 8))

        self.btn_test_all = ttk.Button(btns, text="批量检测全部 Key", command=self._on_test_all)
        self.btn_test_all.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_prune = ttk.Button(
            btns, text="一键剔除失效 Key", state=tk.DISABLED, command=self._on_prune_failed
        )
        self.btn_prune.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_clear = ttk.Button(btns, text="清空输入", command=self._on_clear)
        self.btn_clear.pack(side=tk.LEFT)

        tk.Frame(wrap, bg="#e2e8f0", height=1).pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 10))

        chk_frame = ttk.LabelFrame(wrap, text=f" 批量可用性检测结果（检测模型：{self.model_id}） ", padding=6)
        chk_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 4))

        self.chk_text = tk.Text(
            chk_frame,
            height=5,
            wrap=tk.WORD,
            font=(UI_FONT_FAMILY, _ui_pt(9)),
            bg="#f8fafc",
            fg="#0f172a",
            relief=tk.SOLID,
            borderwidth=1,
            padx=6,
            pady=4,
        )
        self.chk_text.tag_configure("ok", foreground="#15803d")
        self.chk_text.tag_configure("bad", foreground="#b91c1c")
        self.chk_text.tag_configure("info", foreground="#1d4ed8", font=(UI_FONT_FAMILY, _ui_pt(9), "bold"))
        chk_vsb = ttk.Scrollbar(chk_frame, orient=tk.VERTICAL, command=self.chk_text.yview)
        self.chk_text.configure(yscrollcommand=chk_vsb.set)
        self.chk_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        chk_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.chk_text.insert(tk.END, "点击下方「批量检测全部 Key」可逐一验证所有密钥可用性，并支持一键剔除失效 Key。\n")

        cfg_box = ttk.Frame(wrap)
        cfg_box.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 8))

        stat_row = ttk.Frame(cfg_box)
        stat_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Checkbutton(
            stat_row,
            text="本机加密记住密钥池 (Windows DPAPI)",
            variable=self.remember_var,
        ).pack(side=tk.RIGHT)
        ttk.Label(
            stat_row,
            textvariable=self.count_var,
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
            foreground="#1d4ed8",
        ).pack(side=tk.LEFT)

        chan_row = ttk.Frame(cfg_box)
        chan_row.pack(fill=tk.X)
        chan_row.columnconfigure(1, weight=1)
        chan_row.columnconfigure(3, weight=1)

        ttk.Label(chan_row, text="接口通道：").grid(row=0, column=0, sticky=tk.W)
        ttk.Combobox(
            chan_row,
            textvariable=self.api_mode_var,
            values=[m[0] for m in API_MODES],
            state="readonly",
            width=28,
        ).grid(row=0, column=1, sticky=tk.EW, padx=(4, 14))

        ttk.Label(chan_row, text="调度策略：").grid(row=0, column=2, sticky=tk.W)
        ttk.Combobox(
            chan_row,
            textvariable=self.rotation_var,
            values=[m[0] for m in ROTATION_MODES],
            state="readonly",
            width=30,
        ).grid(row=0, column=3, sticky=tk.EW, padx=(4, 0))

        # Pack flexible editor LAST so it smoothly absorbs vertical resize
        ed_frame = ttk.LabelFrame(wrap, text=" 密钥列表（每行 1 个 Key，支持 # 账号备注，数量无上限） ", padding=8)
        ed_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 8))

        hsb = ttk.Scrollbar(ed_frame, orient=tk.HORIZONTAL, command=lambda *a: self.editor.xview(*a))
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

        ed_inner = ttk.Frame(ed_frame)
        ed_inner.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.editor = tk.Text(
            ed_inner,
            height=8,
            wrap=tk.NONE,
            font=(MONO_FONT_FAMILY, _ui_pt(9)),
            bg="#ffffff",
            fg="#0f172a",
            relief=tk.SOLID,
            borderwidth=1,
            padx=8,
            pady=6,
            undo=True,
        )
        vsb = ttk.Scrollbar(ed_inner, orient=tk.VERTICAL, command=self.editor.yview)
        self.editor.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.editor.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.editor.bind("<KeyRelease>", lambda _e: self._update_count())
        self.editor.bind("<<Paste>>", lambda _e: self.after(20, self._update_count))

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
                    self.btn_test_all.configure(state=tk.NORMAL, text="批量检测全部 Key")
                    self._failed_keys_set = set(bad_keys)
                    self.btn_prune.configure(state=tk.NORMAL if bad_keys else tk.DISABLED)
                    summary = f"检测完毕：可用 {ok_cnt} 个  |  不可用/受限 {len(bad_keys)} 个。"
                    if bad_keys:
                        summary += "（可点击「一键剔除失效 Key」自动清理）"
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

        eff_mode = self._resolve_current_api_mode()
        self._testing = True
        self._failed_keys_set.clear()
        self.btn_prune.configure(state=tk.DISABLED)
        self.btn_test_all.configure(state=tk.DISABLED, text="正在逐个检测...")
        self.chk_text.delete("1.0", tk.END)
        self.chk_text.insert(tk.END, f"开始检测共 {len(keys)} 个 API Key（模型：{self.model_id}）...\n", "info")

        def _each(done_i: int, tot: int, item: dict):
            prefix = "[通过]" if item["ok"] else "[失败]"
            tag = "ok" if item["ok"] else "bad"
            raw_detail = str(item.get("detail", ""))
            if item["ok"]:
                detail = redact_secrets(_clean_single_line_message(raw_detail), extra_secret=keys)
            else:
                detail = _format_exception_concise(RuntimeError(raw_detail), keys=keys, api_related=True)
            self._ui_queue.put(("line", (f"[{done_i}/{tot}] {prefix} {item['masked']}：{detail}", tag)))

        def worker():
            try:
                reps = test_all_api_keys(
                    keys,
                    model=self.model_id,
                    api_mode=eff_mode,
                    on_each_cb=_each,
                )
                ok_cnt = sum(1 for r in reps if r["ok"])
                bad_keys = [r["key"] for r in reps if not r["ok"]]
                self._ui_queue.put(("done", (ok_cnt, bad_keys)))
            except Exception as ex:
                safe_err = _format_exception_concise(ex, keys=keys, api_related=True)
                self._ui_queue.put(("line", (f"[失败] 检测出错：{safe_err}", "bad")))
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
        if self._ext_api_mode_var is not None:
            self._ext_api_mode_var.set(self.api_mode_var.get())
        if self._ext_remember_var is not None:
            self._ext_remember_var.set(bool(self.remember_var.get()))
        raw_txt = self._get_raw_text()
        if callable(self.on_save_cb):
            self.on_save_cb(raw_txt, self.rotation_var.get())
        self.destroy()


class LanguageManagerDialog(tk.Toplevel):
    """Modal dialog for viewing, adding, editing, pinning favorite languages, and managing standard fonts."""

    def __init__(
        self,
        parent: tk.Tk,
        entries: list,
        on_saved_callback,
        favorite_codes: list[str] | None = None,
        initial_code: str | None = None,
    ):
        super().__init__(parent)
        self.withdraw()
        self.title("语种库、标准字体与特殊语种字号/行距存档管理")
        self.transient(parent)

        fav_set = {str(c).strip().upper() for c in (favorite_codes or []) if str(c).strip()}
        self.entries = []
        for x in entries:
            item = dict(x)
            code_u = str(item.get("code") or "").strip().upper()
            item["favorite"] = bool(item.get("favorite") or code_u in fav_set)
            self.entries.append(item)
        self.on_saved_callback = on_saved_callback

        self.code_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.font_var = tk.StringVar()
        self.title_pt_var = tk.StringVar()
        self.body_pt_var = tk.StringVar()
        self.citation_pt_var = tk.StringVar()
        self.line_pt_var = tk.StringVar()
        self.kw_var = tk.StringVar()
        self.fav_var = tk.BooleanVar(value=False)
        self.summary_var = tk.StringVar(value="")

        self._build_ui()
        self._refresh_tree(keep_code=initial_code if initial_code and initial_code != "AUTO" else None)
        if initial_code and initial_code != "AUTO":
            self._on_select_row()
        _fit_and_center_window(self, base_w=920, base_h=630, min_w=780, min_h=520, parent=parent)
        self.deiconify()
        try:
            self.grab_set()
        except tk.TclError:
            pass

    @staticmethod
    def _effective_entry_font(entry: dict) -> str:
        if "font" in entry:
            return str(entry.get("font") or "").strip()
        code = str(entry.get("code") or "").strip().upper()
        return _BUILTIN_FONT_FALLBACKS.get(code, "")

    @staticmethod
    def _is_entry_fav(entry: dict) -> bool:
        return bool(entry.get("favorite"))

    def _build_ui(self):
        wrap = ttk.Frame(self, padding=14)
        wrap.pack(fill=tk.BOTH, expand=True)

        tip = ttk.Label(
            wrap,
            text=(
                "支持将高频语种设为「★ 常用」置顶（与最近使用的语种自动排在所有板块下拉框最上方），配置实时保存在 languages.json 与系统设置中。\n"
                "• 语种代码：用于定稿文件名前缀（如 GR_、TH_、PL_）及自动识别；标准字体留空表示保留原稿第 3 列字体。"
            ),
            foreground="#475569",
        )
        tip.pack(side=tk.TOP, fill=tk.X, pady=(0, 8))

        # Pack bottom action bar FIRST so bottom buttons NEVER clip
        btns = ttk.Frame(wrap)
        btns.pack(side=tk.BOTTOM, fill=tk.X)

        ttk.Button(btns, text="关闭", command=self.destroy).pack(side=tk.RIGHT)
        _make_primary_btn(btns, text="新增 / 保存语种", command=self._on_add_or_update).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        self.btn_toggle_fav = ttk.Button(
            btns,
            text="切换选中语种常用置顶",
            command=self._on_toggle_favorite,
        )
        self.btn_toggle_fav.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btns, text="删除选中语种", command=self._on_delete).pack(side=tk.LEFT)

        tk.Frame(wrap, bg="#e2e8f0", height=1).pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 10))

        form = ttk.LabelFrame(wrap, text=" 新增 / 修改语种与标准字体（点击上方列表行可直接载入编辑） ", padding=10)
        form.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 4))
        form.columnconfigure(5, weight=1)

        ttk.Label(form, text="语种代码(缩写)：").grid(row=0, column=0, sticky=tk.W, pady=4)
        ttk.Entry(form, textvariable=self.code_var, width=10).grid(row=0, column=1, sticky=tk.W, pady=4, padx=(4, 12))

        ttk.Label(form, text="语种名称(中文)：").grid(row=0, column=2, sticky=tk.W, pady=4)
        ttk.Entry(form, textvariable=self.name_var, width=16).grid(row=0, column=3, sticky=tk.W, pady=4, padx=(4, 12))

        ttk.Label(form, text="标准字体(选填)：").grid(row=0, column=4, sticky=tk.W, pady=4)
        self.font_cb = ttk.Combobox(
            form,
            textvariable=self.font_var,
            values=_COMMON_STANDARD_FONTS,
            width=22,
        )
        self.font_cb.grid(row=0, column=5, sticky=tk.EW, pady=4, padx=4)

        ttk.Label(form, text="自动识别关键词：").grid(row=1, column=0, sticky=tk.W, pady=4)
        ttk.Entry(form, textvariable=self.kw_var).grid(row=1, column=1, columnspan=5, sticky=tk.EW, pady=4, padx=4)

        sz_row = ttk.Frame(form)
        sz_row.grid(row=2, column=0, columnspan=6, sticky=tk.EW, pady=4)
        ttk.Label(sz_row, text="特殊语种参数存档（留空用默认 14 / 12 / 10.5 / 20）：").pack(side=tk.LEFT)
        ttk.Label(sz_row, text="标题pt：").pack(side=tk.LEFT, padx=(6, 2))
        ttk.Entry(sz_row, textvariable=self.title_pt_var, width=6).pack(side=tk.LEFT)
        ttk.Label(sz_row, text="正文pt：").pack(side=tk.LEFT, padx=(10, 2))
        ttk.Entry(sz_row, textvariable=self.body_pt_var, width=6).pack(side=tk.LEFT)
        ttk.Label(sz_row, text="摘自pt：").pack(side=tk.LEFT, padx=(10, 2))
        ttk.Entry(sz_row, textvariable=self.citation_pt_var, width=6).pack(side=tk.LEFT)
        ttk.Label(sz_row, text="行距(磅)：").pack(side=tk.LEFT, padx=(10, 2))
        ttk.Entry(sz_row, textvariable=self.line_pt_var, width=6).pack(side=tk.LEFT)

        fav_row = ttk.Frame(form)
        fav_row.grid(row=3, column=0, columnspan=6, sticky=tk.EW, pady=(2, 0))
        self.chk_fav = ttk.Checkbutton(
            fav_row,
            text="设为常用语种（下拉框置顶）",
            variable=self.fav_var,
        )
        self.chk_fav.pack(side=tk.LEFT, padx=(0, 14))
        ttk.Label(
            fav_row,
            text="多个关键词用逗号分隔（如：波兰语, 波语, 中英波, PO）；标准字体留空则保留原字体",
            foreground="#475569",
        ).pack(side=tk.LEFT)

        tbl_toolbar = ttk.Frame(wrap)
        tbl_toolbar.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))
        ttk.Button(
            tbl_toolbar,
            text="打开 languages.json",
            style="Compact.TButton",
            command=self._on_open_json,
        ).pack(side=tk.RIGHT)
        ttk.Button(
            tbl_toolbar,
            text="补全内置默认语种",
            style="Compact.TButton",
            command=self._on_merge_defaults,
        ).pack(side=tk.RIGHT, padx=(0, 6))
        ttk.Label(
            tbl_toolbar,
            textvariable=self.summary_var,
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
            foreground="#1d4ed8",
        ).pack(side=tk.LEFT)

        tbl_frame = ttk.Frame(wrap)
        tbl_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 8))

        cols = ("fav", "code", "name", "font", "keywords")
        self.tree = ttk.Treeview(tbl_frame, columns=cols, show="headings", selectmode="browse", height=8)
        self.tree.heading("fav", text="常用置顶")
        self.tree.heading("code", text="语种代码 (前缀)")
        self.tree.heading("name", text="语种名称")
        self.tree.heading("font", text="标准字体 (如 Georgia / Angsana New)")
        self.tree.heading("keywords", text="自动识别关键词（目录名 / 文件名包含任一词即可）")

        self.tree.column("fav", width=80, anchor=tk.CENTER)
        self.tree.column("code", width=105, anchor=tk.CENTER)
        self.tree.column("name", width=110, anchor=tk.CENTER)
        self.tree.column("font", width=240, anchor=tk.CENTER)
        self.tree.column("keywords", width=295, anchor=tk.W)
        self.tree.tag_configure("fav_row", foreground="#1d4ed8")

        vsb = ttk.Scrollbar(tbl_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<<TreeviewSelect>>", self._on_select_row)
        self.tree.bind("<Double-1>", lambda _e: self._on_toggle_favorite())

    def _refresh_tree(self, keep_code: str | None = None):
        for item in self.tree.get_children():
            self.tree.delete(item)
        fav_cnt = 0
        select_iid: str | None = None
        for idx, e in enumerate(self.entries):
            is_fav = self._is_entry_fav(e)
            if is_fav:
                fav_cnt += 1
            fav_disp = "★ 常用" if is_fav else "-"
            f_val = self._effective_entry_font(e)
            has_custom_sz = any(e.get(k) is not None for k in ("title_pt", "body_pt", "citation_pt", "line_pt"))
            if has_custom_sz:
                t_pt = e.get("title_pt", 14)
                b_pt = e.get("body_pt", 12)
                c_pt = e.get("citation_pt", 10.5)
                l_pt = e.get("line_pt", 20)
                l_tag = f", 行距{l_pt}pt" if e.get("line_pt") is not None else ""
                base_f = f_val if f_val else "原字体"
                f_disp = f"{base_f} ({t_pt}/{b_pt}/{c_pt}pt{l_tag})"
            else:
                f_disp = f_val if f_val else "（保留原字体）"
            iid = str(idx)
            tags = ("fav_row",) if is_fav else ()
            self.tree.insert(
                "",
                tk.END,
                iid=iid,
                values=(fav_disp, e["code"], e["name"], f_disp, ", ".join(e.get("keywords", []))),
                tags=tags,
            )
            if keep_code and str(e.get("code", "")).upper() == keep_code.upper():
                select_iid = iid
        self.summary_var.set(f"当前共 {len(self.entries)} 个语种  |  已设常用置顶：{fav_cnt} 个（双击行也可快速切换）")
        if select_iid is not None:
            self.tree.selection_set(select_iid)
            self.tree.see(select_iid)

    def _on_select_row(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        if 0 <= idx < len(self.entries):
            e = self.entries[idx]
            self.code_var.set(e["code"])
            self.name_var.set(e["name"])
            self.font_var.set(self._effective_entry_font(e))
            self.title_pt_var.set("" if e.get("title_pt") is None else str(e.get("title_pt")))
            self.body_pt_var.set("" if e.get("body_pt") is None else str(e.get("body_pt")))
            self.citation_pt_var.set("" if e.get("citation_pt") is None else str(e.get("citation_pt")))
            self.line_pt_var.set("" if e.get("line_pt") is None else str(e.get("line_pt")))
            self.kw_var.set(", ".join(e.get("keywords", [])))
            self.fav_var.set(self._is_entry_fav(e))

    def _on_toggle_favorite(self):
        sel = self.tree.selection()
        target_idx = -1
        if sel:
            cand = int(sel[0])
            if 0 <= cand < len(self.entries):
                target_idx = cand
        elif self.code_var.get().strip():
            code_u = re.sub(r"[^A-Za-z0-9_-]", "", self.code_var.get().strip()).upper()
            for i, e in enumerate(self.entries):
                if str(e.get("code", "")).upper() == code_u:
                    target_idx = i
                    break
        if target_idx < 0:
            messagebox.showinfo("提示", "请先在列表中点击选中一个语种，再切换常用置顶。", parent=self)
            return

        entry = dict(self.entries[target_idx])
        new_fav = not self._is_entry_fav(entry)
        entry["favorite"] = new_fav
        self.entries[target_idx] = entry
        self.fav_var.set(new_fav)
        save_languages(self.entries)
        self._refresh_tree(keep_code=entry["code"])
        self.on_saved_callback(self.entries, selected_code=None)

    def _on_add_or_update(self):
        raw_code = re.sub(r"[^A-Za-z0-9_-]", "", self.code_var.get().strip()).upper()
        name = self.name_var.get().strip()
        font_val = self.font_var.get().strip()
        is_fav = bool(self.fav_var.get())
        if not raw_code or raw_code == "AUTO":
            messagebox.showwarning("提示", "请填写有效的英文语种代码（例如 TH、NL、SR，不能为 AUTO）。", parent=self)
            return
        if not name:
            messagebox.showwarning("提示", "请填写语种中文名称（例如 泰语、塞尔维亚语）。", parent=self)
            return

        kws = build_keywords_list(name, self.kw_var.get())

        custom_pts: dict = {}
        for pt_key, var_obj, lbl_cn in (
            ("title_pt", self.title_pt_var, "标题字号"),
            ("body_pt", self.body_pt_var, "正文字号"),
            ("citation_pt", self.citation_pt_var, "摘自字号"),
            ("line_pt", self.line_pt_var, "行距磅数"),
        ):
            raw_s = var_obj.get().strip()
            if raw_s:
                try:
                    val_f = float(raw_s)
                    if val_f <= 0 or val_f > 72:
                        raise ValueError()
                    custom_pts[pt_key] = int(val_f) if val_f.is_integer() else val_f
                except ValueError:
                    messagebox.showwarning("提示", f"「{lbl_cn}」请输入大于 0 的数字（或留空使用默认值）。", parent=self)
                    return

        updated = False
        for i, e in enumerate(self.entries):
            if e["code"] == raw_code:
                merged_entry = dict(e)
                for rm_k in ("title_pt", "body_pt", "citation_pt", "line_pt"):
                    merged_entry.pop(rm_k, None)
                merged_entry.update(
                    {
                        "code": raw_code,
                        "name": name,
                        "keywords": kws,
                        "font": font_val,
                        "favorite": is_fav,
                        **custom_pts,
                    }
                )
                self.entries[i] = merged_entry
                updated = True
                break
        if not updated:
            self.entries.insert(
                0,
                {
                    "code": raw_code,
                    "name": name,
                    "keywords": kws,
                    "font": font_val,
                    "favorite": is_fav,
                    **custom_pts,
                },
            )

        save_languages(self.entries)
        self._refresh_tree(keep_code=raw_code)
        self.on_saved_callback(self.entries, selected_code=raw_code)
        action_str = "已更新" if updated else "已新增"
        font_disp = font_val or "保留原稿字体"
        fav_disp = "★ 常用置顶" if is_fav else "普通排序"
        messagebox.showinfo(
            "保存成功",
            f"{action_str}语种：{raw_code} - {name}（{fav_disp}）\n标准字体：{font_disp}\n识别关键词：{', '.join(kws)}",
            parent=self,
        )

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
        existing_by_code = {e["code"]: e for e in self.entries}
        added = 0
        backfilled = 0
        for d in DEFAULT_LANGUAGES:
            code = d["code"]
            if code not in existing_by_code:
                entry_copy = dict(d)
                if "font" not in entry_copy:
                    entry_copy["font"] = _BUILTIN_FONT_FALLBACKS.get(code, "")
                entry_copy.setdefault("favorite", False)
                self.entries.append(entry_copy)
                added += 1
            else:
                cur = existing_by_code[code]
                def_font = d.get("font", _BUILTIN_FONT_FALLBACKS.get(code, ""))
                if not cur.get("font") and def_font:
                    cur["font"] = def_font
                    backfilled += 1
                for k_pt in ("title_pt", "body_pt", "citation_pt"):
                    if k_pt in d and k_pt not in cur:
                        cur[k_pt] = d[k_pt]
        save_languages(self.entries)
        self._refresh_tree()
        self.on_saved_callback(self.entries, selected_code=None)
        messagebox.showinfo(
            "完成",
            f"已补全 {added} 个内置语种，同步补齐 {backfilled} 项默认标准字体（您自定义的语种、常用置顶与字体均保留）。",
            parent=self,
        )

    def _on_open_json(self):
        p = get_languages_file().resolve()
        if not p.exists():
            save_languages(self.entries)
        if p.is_file() and p.suffix.lower() == ".json":
            open_path_in_os(p)


LangManagerDialog = LanguageManagerDialog


class AmbiguousArticleDialog(tk.Toplevel):
    """Modal disambiguation dialog when a full-article title query (e.g. '第一篇') matches multiple chapters."""

    def __init__(self, parent: tk.Tk, ambiguous_items: list[dict], resolved_count: int = 0):
        super().__init__(parent)
        self.withdraw()
        self.title("多篇同名标题选择确认（请勾选本次需要提取对比的具体篇目）")
        self.transient(parent)

        self.ambiguous_items = ambiguous_items or []
        self.resolved_count = resolved_count
        self.selected_candidates: list[dict] = []
        self.cancelled: bool = True

        self._cand_records: list[tuple[tk.BooleanVar, dict, str]] = []
        self.count_var = tk.StringVar(value="当前已勾选 0 篇候选篇目")

        self._build_ui()
        self._update_count()
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        _fit_and_center_window(self, base_w=820, base_h=560, min_w=700, min_h=440, parent=parent)
        self.deiconify()
        try:
            self.grab_set()
        except tk.TclError:
            pass

    @staticmethod
    def _format_candidate_label(cand: dict) -> str:
        disp = str(cand.get("display_label") or "").strip()
        if disp:
            return disp
        parts = []
        master_f = cand.get("master_rel_path") or cand.get("master_file") or ""
        if master_f:
            parts.append(f"[{Path(str(master_f)).name}]")
        vol = str(cand.get("volume_title") or cand.get("section_path") or "").strip()
        if vol:
            parts.append(vol)
        cn_t = str(cand.get("cn_title") or cand.get("title") or "").strip()
        if cn_t:
            parts.append(cn_t)
        en_t = str(cand.get("en_title") or "").strip()
        if en_t:
            parts.append(f"({en_t})")
        rc = cand.get("row_count")
        if rc is not None:
            parts.append(f"— 共 {rc} 行")
        return " / ".join(parts) if parts else str(cand)

    def _build_ui(self):
        wrap = ttk.Frame(self, padding=14)
        wrap.pack(fill=tk.BOTH, expand=True)

        extra_note = (
            f"（另有 {self.resolved_count} 篇唯一命中的标题已自动就绪）" if self.resolved_count > 0 else ""
        )
        ttk.Label(
            wrap,
            text=(
                f"检测到以下搜索标题在中英底稿目录中对应多篇不同卷/分部的同名篇目{extra_note}。\n"
                "请勾选您本次实际需要提取对比的篇目（支持单选或多选），确认后将逐篇独立导出完整版与仅改点版："
            ),
            foreground="#0f172a",
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
        ).pack(side=tk.TOP, fill=tk.X, pady=(0, 8))

        # Pack bottom action bar FIRST so bottom buttons NEVER clip
        bot = ttk.Frame(wrap)
        bot.pack(side=tk.BOTTOM, fill=tk.X)

        ttk.Button(bot, text="取消任务", command=self._on_cancel).pack(side=tk.RIGHT)
        _make_primary_btn(
            bot,
            text="确认选择并继续对比",
            command=self._on_confirm,
        ).pack(side=tk.RIGHT, padx=(0, 8))

        ttk.Button(bot, text="全选全部候选", command=lambda: self._set_all(True)).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(bot, text="清空全部勾选", command=lambda: self._set_all(False)).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(
            bot,
            textvariable=self.count_var,
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
            foreground="#1d4ed8",
        ).pack(side=tk.LEFT)

        tk.Frame(wrap, bg="#e2e8f0", height=1).pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 10))

        list_outer = ttk.Frame(wrap, relief=tk.SOLID, borderwidth=1)
        list_outer.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 4))

        self.canvas = tk.Canvas(list_outer, bg="#f8fafc", highlightthickness=0)
        vsb = ttk.Scrollbar(list_outer, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scroll_inner = tk.Frame(self.canvas, bg="#f8fafc", padx=10, pady=8)

        self.scroll_inner.bind(
            "<Configure>",
            lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self._canvas_win = self.canvas.create_window((0, 0), window=self.scroll_inner, anchor="nw")
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfigure(self._canvas_win, width=e.width),
        )
        self.canvas.configure(yscrollcommand=vsb.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        for g_idx, amb in enumerate(self.ambiguous_items):
            q_str = str(amb.get("query") or f"标题 #{g_idx + 1}")
            cands = list(amb.get("candidates") or [])
            group_key = f"g_{g_idx}"

            grp = ttk.LabelFrame(
                self.scroll_inner,
                text=f" 搜索标题：「{q_str}」（匹配到 {len(cands)} 篇候选） ",
                padding=8,
            )
            grp.pack(fill=tk.X, pady=(0, 8))

            grp_bar = ttk.Frame(grp)
            grp_bar.pack(fill=tk.X, pady=(0, 4))
            ttk.Button(
                grp_bar,
                text="全选本组",
                command=lambda gk=group_key: self._set_group(gk, True),
            ).pack(side=tk.LEFT, padx=(0, 6))
            ttk.Button(
                grp_bar,
                text="清空本组",
                command=lambda gk=group_key: self._set_group(gk, False),
            ).pack(side=tk.LEFT)

            for cand in cands:
                var = tk.BooleanVar(value=False)
                lbl = self._format_candidate_label(cand)
                cb = ttk.Checkbutton(
                    grp,
                    text=lbl,
                    variable=var,
                    command=self._update_count,
                )
                cb.pack(fill=tk.X, anchor=tk.W, pady=2)
                self._cand_records.append((var, cand, group_key))

    def _set_group(self, group_key: str, val: bool):
        for var, _cand, gk in self._cand_records:
            if gk == group_key:
                var.set(val)
        self._update_count()

    def _set_all(self, val: bool):
        for var, _cand, _gk in self._cand_records:
            var.set(val)
        self._update_count()

    def _update_count(self):
        n = sum(1 for var, _c, _gk in self._cand_records if var.get())
        self.count_var.set(f"当前已勾选 {n} 篇候选篇目")

    def _on_confirm(self):
        chosen = [cand for var, cand, _gk in self._cand_records if var.get()]
        if not chosen and self.resolved_count == 0:
            messagebox.showwarning("提示", "请至少勾选 1 篇需要提取对比的候选篇目（或点击「取消任务」）。", parent=self)
            return
        if not chosen and self.resolved_count > 0:
            if not messagebox.askyesno(
                "确认跳过同名标题",
                f"当前未勾选任何同名候选篇目，是否仅继续处理其余 {self.resolved_count} 篇已唯一命中的标题？",
                parent=self,
            ):
                return
        self.selected_candidates = chosen
        self.cancelled = False
        self.destroy()

    def _on_cancel(self):
        self.selected_candidates = []
        self.cancelled = True
        self.destroy()


class SoftwareUpdateDialog(tk.Toplevel):
    """Dialog for checking GitHub Releases, viewing release notes, and one-click SHA256-verified self-updating."""

    def __init__(self, parent: tk.Tk, app: "ScatteredRevisionApp", initial_info: dict | None = None):
        super().__init__(parent)
        self.withdraw()
        self.title(f"检查软件更新（当前版本：{APP_VERSION}）")
        self.transient(parent)

        self.app = app
        self.update_info: dict | None = initial_info
        self._busy = False
        self._ui_queue: queue.Queue = queue.Queue()

        self.header_var = tk.StringVar(value="正在检查 GitHub 最新版本...")
        self.status_var = tk.StringVar(value="准备就绪")
        self.prog_var = tk.DoubleVar(value=0.0)

        self._build_ui()
        _fit_and_center_window(self, base_w=680, base_h=480, min_w=580, min_h=400, parent=parent)
        self.deiconify()
        try:
            self.grab_set()
        except tk.TclError:
            pass
        self._poll_queue()

        if self.update_info is not None:
            self._render_update_info(self.update_info)
        else:
            self._start_check()

    def _build_ui(self):
        wrap = ttk.Frame(self, padding=14)
        wrap.pack(fill=tk.BOTH, expand=True)

        self.lbl_header = ttk.Label(
            wrap,
            textvariable=self.header_var,
            font=(UI_FONT_FAMILY, _ui_pt(10), "bold"),
            foreground="#1d4ed8",
        )
        self.lbl_header.pack(side=tk.TOP, fill=tk.X, pady=(0, 8))

        # Pack bottom action bar FIRST so bottom buttons NEVER clip
        btns = ttk.Frame(wrap)
        btns.pack(side=tk.BOTTOM, fill=tk.X)

        ttk.Button(btns, text="关闭", command=self.destroy).pack(side=tk.RIGHT)
        self.btn_apply = _make_primary_btn(
            btns,
            text="立即下载更新并自动重启",
            state=tk.DISABLED,
            command=self._start_download_and_apply,
        )
        self.btn_apply.pack(side=tk.RIGHT, padx=(0, 8))

        self.btn_recheck = ttk.Button(btns, text="重新检查", command=self._start_check)
        self.btn_recheck.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_open_web = ttk.Button(btns, text="打开 GitHub 发布页", command=self._open_release_webpage)
        self.btn_open_web.pack(side=tk.LEFT)

        tk.Frame(wrap, bg="#e2e8f0", height=1).pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 10))

        prog_box = ttk.Frame(wrap)
        prog_box.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 4))
        ttk.Label(prog_box, textvariable=self.status_var, foreground="#475569").pack(fill=tk.X, pady=(0, 4))
        self.prog_bar = ttk.Progressbar(prog_box, mode="determinate", maximum=100.0, variable=self.prog_var)
        self.prog_bar.pack(fill=tk.X)

        notes_frame = ttk.LabelFrame(wrap, text=" 版本更新说明 (Release Notes) ", padding=8)
        notes_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 8))

        self.notes_text = tk.Text(
            notes_frame,
            height=9,
            wrap=tk.WORD,
            font=(UI_FONT_FAMILY, _ui_pt(9)),
            bg="#f8fafc",
            fg="#0f172a",
            relief=tk.SOLID,
            borderwidth=1,
            padx=8,
            pady=6,
        )
        vsb = ttk.Scrollbar(notes_frame, orient=tk.VERTICAL, command=self.notes_text.yview)
        self.notes_text.configure(yscrollcommand=vsb.set)
        self.notes_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

    def _set_notes(self, content: str):
        self.notes_text.configure(state=tk.NORMAL)
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", (content or "").strip() or "暂无详细版本说明。")
        self.notes_text.configure(state=tk.DISABLED)

    def _open_release_webpage(self):
        url = GITHUB_RELEASES_URL
        if isinstance(self.update_info, dict):
            cand = str(self.update_info.get("html_url") or "").strip()
            if cand.startswith("https://github.com/"):
                url = cand
        try:
            webbrowser.open(url)
        except Exception:
            pass

    def _render_update_info(self, info: dict):
        self.update_info = info
        has_upd = bool(info.get("has_update"))
        latest_v = str(info.get("latest_version") or RAW_APP_VERSION).lstrip("vV")
        notes = str(info.get("release_notes") or "").strip()
        if has_upd:
            self.header_var.set(f"发现新版本 v{latest_v}（当前版本：{APP_VERSION}）")
            self.lbl_header.configure(foreground="#1d4ed8")
            self.status_var.set("更新时将自动校验 SHA-256 签名，并完整保留您本机的 API Key 池与语种设置。")
            self.prog_var.set(0.0)
            self._set_notes(notes)
            self.btn_apply.configure(state=tk.NORMAL, text="立即下载更新并自动重启")
            self.app._show_update_banner(info)
        else:
            self.header_var.set(f"当前已是最新版本（v{latest_v}）")
            self.lbl_header.configure(foreground="#15803d")
            self.status_var.set("无需更新。")
            self.prog_var.set(100.0)
            self._set_notes(notes or f"当前运行版本 {APP_VERSION} 已是 GitHub 最新正式版。")
            self.btn_apply.configure(state=tk.DISABLED, text="当前已是最新版")

    def _start_check(self):
        if self._busy:
            return
        if not callable(check_for_updates):
            self.header_var.set("更新组件未就绪")
            self.lbl_header.configure(foreground="#b45309")
            self._set_notes(f"当前环境未加载 updater 模块，请直接前往 GitHub Releases 页面查看：\n{GITHUB_RELEASES_URL}")
            return

        self._busy = True
        self.btn_recheck.configure(state=tk.DISABLED)
        self.btn_apply.configure(state=tk.DISABLED)
        self.prog_var.set(15.0)
        self.header_var.set("正在连接 GitHub 检查最新版本...")
        self.lbl_header.configure(foreground="#1d4ed8")
        self.status_var.set("正在请求 GitHub Releases 接口...")
        self._set_notes("正在获取最新版本信息，请稍候...")

        def worker():
            try:
                info = check_for_updates(current_version=RAW_APP_VERSION)
                self._ui_queue.put(("checked_ok", info))
            except Exception as ex:
                self._ui_queue.put(("checked_err", _clean_single_line_message(ex)))

        threading.Thread(target=worker, daemon=True).start()

    def _start_download_and_apply(self):
        if self._busy or not isinstance(self.update_info, dict):
            return
        if self.app.is_running:
            messagebox.showwarning(
                "任务运行中",
                "当前后台有文档整理任务正在运行，请等待任务完成后再执行自动更新重启。",
                parent=self,
            )
            return
        if not callable(download_and_apply_update):
            messagebox.showerror("组件缺失", "未找到 download_and_apply_update 函数，请点击左侧按钮前往网页下载。", parent=self)
            return

        self._busy = True
        self.btn_recheck.configure(state=tk.DISABLED)
        self.btn_apply.configure(state=tk.DISABLED, text="正在下载更新...")
        self.prog_var.set(2.0)
        self.status_var.set("正在准备下载更新包...")

        info_snapshot = dict(self.update_info)

        def _prog_cb(pct: float, msg: str):
            self._ui_queue.put(("dl_prog", (pct, msg)))

        def worker():
            try:
                res = download_and_apply_update(info_snapshot, progress_fn=_prog_cb)
                self._ui_queue.put(("dl_done", res))
            except Exception as ex:
                self._ui_queue.put(("dl_err", _clean_single_line_message(ex)))

        threading.Thread(target=worker, daemon=True).start()

    def _poll_queue(self):
        if not self.winfo_exists():
            return
        try:
            while True:
                try:
                    kind, payload = self._ui_queue.get_nowait()
                except queue.Empty:
                    break

                if kind == "checked_ok":
                    self._busy = False
                    self.btn_recheck.configure(state=tk.NORMAL)
                    self._render_update_info(payload)
                elif kind == "checked_err":
                    self._busy = False
                    self.btn_recheck.configure(state=tk.NORMAL)
                    self.prog_var.set(0.0)
                    self.header_var.set("检查更新失败（可点击下方按钮在浏览器查看）")
                    self.lbl_header.configure(foreground="#b91c1c")
                    self.status_var.set(str(payload))
                    self._set_notes(
                        f"无法连接 GitHub 检查更新：{payload}\n\n"
                        f"您可以检查网络连接后点击「重新检查」，或点击「打开 GitHub 发布页」手动下载最新版：\n{GITHUB_RELEASES_URL}"
                    )
                elif kind == "dl_prog":
                    pct, msg = payload
                    self.prog_var.set(max(0.0, min(100.0, float(pct))))
                    self.status_var.set(str(msg))
                elif kind == "dl_done":
                    self._busy = False
                    self.prog_var.set(100.0)
                    res = payload if isinstance(payload, dict) else {}
                    if res.get("ready_to_restart"):
                        self.status_var.set("更新包替换校验通过！正在保存配置并自动重启新版本...")
                        self.app._save_settings()
                        self.after(200, self.app._exit_for_update_restart)
                        return
                    else:
                        staged = str(res.get("staged_exe") or "")
                        self.btn_recheck.configure(state=tk.NORMAL)
                        self.status_var.set("已在源码模式下完成新版 EXE 下载校验。")
                        messagebox.showinfo(
                            "下载完成（源码运行模式）",
                            "当前以 Python 源码方式运行，已为您下载并校验最新独立版 EXE：\n\n"
                            f"{staged or GITHUB_RELEASES_URL}\n\n"
                            "（独立打包 .exe 运行时点击更新将自动热替换并重启）",
                            parent=self,
                        )
                elif kind == "dl_err":
                    self._busy = False
                    self.btn_recheck.configure(state=tk.NORMAL)
                    self.btn_apply.configure(state=tk.NORMAL, text="立即下载更新并自动重启")
                    self.status_var.set(f"自动更新失败：{payload}")
                    messagebox.showerror(
                        "自动更新失败",
                        f"{payload}\n\n您可点击「打开 GitHub 发布页」通过浏览器下载最新压缩包。",
                        parent=self,
                    )
        finally:
            if self.winfo_exists():
                self.after(80, self._poll_queue)


class ScatteredRevisionApp:
    _SECTION_TITLES: dict[str, str] = {
        "main": "零散改点整理",
        "fa": "全篇改点整理",
        "dd": "改点稿查重",
        "ingest": "改点入库与通读定稿",
        "wc": "英文字数统计",
        "fel": "分part稿整理/合并",
        "settings": "系统设置",
    }

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(
            f"多语种整理工作台 {APP_VERSION}  (零散改点 · 全篇对比 · 改点查重 · 改点入库/通读 · 英文字数统计 · 分part稿整理/合并 · 系统设置)"
        )

        self.log_queue: queue.Queue = queue.Queue()
        self.is_running = False
        self._running_tab_key: str = "main"
        self.task_start_ts: float = 0.0
        self.current_pct: float = 0.0
        self.languages: list = load_languages()
        self.favorite_langs: list[str] = [
            str(e.get("code", "")).strip().upper() for e in self.languages if e.get("favorite")
        ]
        self.recent_langs: list[str] = []

        self._pool_multiline_text: str = ""
        self._syncing_key_var: bool = False
        self._syncing_settings_editor: bool = False
        self.settings_key_editor: tk.Text | None = None
        self.dd_files: list[str] = []
        self.wc_files: list[str] = []
        self.fel_files: list[str] = []
        self._fel_inspect_cache: dict[str, dict] = {}
        self.active_tab_key: str = "main"
        self._nav_items: dict = {}
        self._nav_sb_visible: bool = False
        self._sidebar_hover: bool = False
        self.key_entry = None
        self.pool_btn = None
        self.test_btn = None
        self.fa_pool_btn = None
        self.fel_pool_btn = None
        self.btn_refresh_input = None
        self.btn_fa_refresh_master = None
        self.btn_settings_test = None
        self.btn_settings_test_all = None
        self._indexing_lock = threading.Lock()
        self._indexing_inflight: set[str] = set()
        self._last_typed_lib_norm: str = ""
        self._last_typed_fa_lib_norm: str = ""
        self._latest_update_info: dict | None = None
        self._closing: bool = False
        self._poll_after_id: str | None = None
        self._timer_after_id: str | None = None
        self._upd_after_id: str | None = None
        self.settings: dict = self._read_raw_settings_dict()
        self.dnd = WinFileDropManager(self.root, fallback_cb=self._on_drop_fallback)

        self._setup_styles()
        self._init_vars()
        self._load_settings()
        self._build_ui()
        self._setup_drag_and_drop()
        self._refresh_key_pool_badge()
        self._refresh_fav_summary()
        self._refresh_top_status_summary()
        _fit_and_center_window(self.root, base_w=1200, base_h=850, min_w=1050, min_h=700, parent=None)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._poll_log_queue()
        self._tick_elapsed_timer()

        init_lib = self.library_dir_var.get().strip()
        if init_lib and Path(init_lib).is_dir():
            self._warm_library_index_async(init_lib, target_tab="main")

        init_fa_lib = self.fa_library_dir_var.get().strip()
        if init_fa_lib and Path(init_fa_lib).is_dir():
            if not init_lib or os.path.normcase(str(Path(init_fa_lib).resolve())) != os.path.normcase(
                str(Path(init_lib).resolve()) if Path(init_lib).exists() else ""
            ):
                self._warm_library_index_async(init_fa_lib, target_tab="fa")

        self._upd_after_id = self.root.after(700, self._start_auto_update_check_async)

    @staticmethod
    def _read_raw_settings_dict() -> dict:
        if SETTINGS_FILE.exists():
            try:
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    return data
            except Exception:
                pass
        return {}

    @staticmethod
    def _sanitize_yellow_gap(raw_val) -> int:
        try:
            v = int(float(str(raw_val).strip()))
        except (ValueError, TypeError):
            return 3
        return max(0, min(20, v))

    def _get_yellow_gap_words(self) -> int:
        return self._sanitize_yellow_gap(self.yellow_gap_words_var.get())

    def _setup_styles(self):
        style = ttk.Style(self.root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")

        default_font = (UI_FONT_FAMILY, _ui_pt(9))
        bold_font = (UI_FONT_FAMILY, _ui_pt(9), "bold")
        btn_font = (UI_FONT_FAMILY, _ui_pt(10), "bold")
        tab_font = (UI_FONT_FAMILY, _ui_pt(10), "bold")

        self.root.option_add("*Font", default_font)
        style.configure("TLabelframe.Label", font=bold_font, foreground="#1d4ed8")
        style.configure("TButton", font=default_font, padding=(12, 5))
        style.configure("Action.TButton", font=default_font, padding=(12, 5))
        style.configure("Compact.TButton", font=default_font, padding=(10, 4))
        style.configure("Primary.TButton", font=btn_font, padding=(14, 6))
        style.configure("TNotebook.Tab", font=tab_font, padding=(16, 7))

    def _init_vars(self):
        # Global workbench, System Settings & shared API vars
        self.active_section_title_var = tk.StringVar(value=self._SECTION_TITLES["main"])
        self.top_status_summary_var = tk.StringVar(value="")
        self.api_mode_label = tk.StringVar(value=API_MODES[0][0])
        self.rotation_label = tk.StringVar(value=ROTATION_MODES[0][0])
        self.api_key_var = tk.StringVar(value="")
        self.pool_btn_text_var = tk.StringVar(value="API 密钥池 (0 个)")
        self.settings_key_count_var = tk.StringVar(value="当前共识别到 0 个不重复的 API Key")
        self.settings_save_status_var = tk.StringVar(
            value="提示：修改模型、通道或黄标连片间隔后自动同步至各工作台；点击「保存设置」可立即加密落盘"
        )
        self.fav_summary_var = tk.StringVar(value="")
        self.show_key_var = tk.BooleanVar(value=False)
        self.remember_key_var = tk.BooleanVar(value=True)
        self.dynamic_models: list[tuple[str, str]] = list(SUPPORTED_MODELS)
        self.model_label = tk.StringVar(value=SUPPORTED_MODELS[0][0])
        self.yellow_gap_words_var = tk.StringVar(
            value=str(self._sanitize_yellow_gap(self.settings.get("yellow_gap_words", 3)))
        )

        # Tab 1: Scattered Revision (零散改点整理) — all task paths start empty on fresh launch
        self.input_docx_var = tk.StringVar(value="")
        self.library_dir_var = tk.StringVar(value="")
        self.output_dir_var = tk.StringVar(value="")
        self.sc1_sync_master_var = tk.BooleanVar(value=False)
        self.sc1_master_path_var = tk.StringVar(value="")
        self.lang_var = tk.StringVar(value=AUTO_COMBO_LABEL)
        self.sc1_naming_mode_var = tk.StringVar(value="scattered")
        self.sc1_keyword_var = tk.StringVar(value="")
        self.doc_name_var = tk.StringVar(value="")
        self.dual_outputs_var = tk.BooleanVar(value=True)
        self.sc1_balance_cols_var = tk.BooleanVar(value=True)
        self.sc1_hl_base_var = tk.StringVar(value="mark_colors")
        self.sc1_col1_underline_var = tk.BooleanVar(value=True)
        self.sc1_highlight_mode_var = tk.StringVar(value="all_three")
        self.sc1_export_raw_var = tk.BooleanVar(value=True)
        self.sc1_enable_ai_var = tk.BooleanVar(value=True)
        self.sc1_enable_dedup_var = tk.BooleanVar(value=False)
        self.sc1_dedup_dir_var = tk.StringVar(value="")

        self.progress_var = tk.DoubleVar(value=0.0)
        self.stage_var = tk.StringVar(value="状态：就绪（请选择文档与目录后点击「一键输出定稿」）")
        self.pct_timer_var = tk.StringVar(value="0%")

        # Tab 2: Full-Article Revision (全篇改点整理)
        self.fa_use_titles_var = tk.BooleanVar(value=True)
        self.fa_master_dir_var = tk.StringVar(value="")
        self.fa_library_dir_var = tk.StringVar(value="")
        self.fa_output_dir_var = tk.StringVar(value="")
        self.fa_enable_dedup_var = tk.BooleanVar(value=False)
        self.fa_dedup_dir_var = tk.StringVar(value="")
        self.fa_lang_var = tk.StringVar(value=AUTO_COMBO_LABEL)
        self.fa_date_var = tk.StringVar(value=datetime.now().strftime("%Y%m%d"))
        self.fa_balance_cols_var = tk.BooleanVar(value=True)
        self.fa_hl_base_var = tk.StringVar(value="mark_colors")
        self.fa_col1_underline_var = tk.BooleanVar(value=True)
        self.fa_highlight_mode_var = tk.StringVar(value="all_three")
        self.fa_enable_ai_var = tk.BooleanVar(value=True)
        self.fa_title_count_var = tk.StringVar(value="当前共输入 0 个待提取篇目标题")

        self.fa_progress_var = tk.DoubleVar(value=0.0)
        self.fa_stage_var = tk.StringVar(value="状态：就绪（粘贴或导入篇目标题并设定目录后点击开始）")
        self.fa_pct_timer_var = tk.StringVar(value="0%")

        # Tab 3: Standalone Dedup Checker (改点稿查重) — all task paths start empty on fresh launch
        self.dd_pkg_dir_var = tk.StringVar(value="")
        self.dd_out_dir_var = tk.StringVar(value="")
        self.dd_export_deduped_var = tk.BooleanVar(value=True)
        self.dd_balance_cols_var = tk.BooleanVar(value=True)
        self.dd_date_var = tk.StringVar(value=datetime.now().strftime("%Y%m%d"))
        self.dd_summary_var = tk.StringVar(value="当前待查重列表共 0 份文档  |  支持拖入待查重 .docx 或文件夹")

        self.dd_progress_var = tk.DoubleVar(value=0.0)
        self.dd_stage_var = tk.StringVar(value="状态：就绪（拖入待查重文档并选择未入库改点包目录后开始查重）")
        self.dd_pct_timer_var = tk.StringVar(value="0%")

        # Tab 4: English Word Count & Renamer (英文字数统计) — all task inputs start empty on fresh launch
        self.wc_lang_var = tk.StringVar(value=AUTO_COMBO_LABEL)
        self.wc_keep_existing_lang_var = tk.BooleanVar(value=True)
        self.wc_date_var = tk.StringVar(value=datetime.now().strftime("%Y%m%d"))
        self.wc_keep_existing_date_var = tk.BooleanVar(value=False)
        self.wc_strip_draft_var = tk.BooleanVar(value=True)
        self.wc_naming_mode_var = tk.StringVar(value="keep")
        self.wc_keyword_var = tk.StringVar(value="")
        self.wc_update_headers_var = tk.BooleanVar(value=False)
        self.wc_save_mode_var = tk.StringVar(value="inplace")
        self.wc_copy_dir_var = tk.StringVar(value="")
        self.wc_summary_var = tk.StringVar(
            value="当前共 0 份文档  |  第 2 列英文合计：RV 0 words, UP 0 words"
        )

        # Tab 5: Part Manuscript Formatting & Merging (分part稿整理/合并)
        self._fel_batch_completed = False
        self._fel_output_was_auto = False
        self._dd_batch_completed = False
        self._wc_batch_completed = False
        self.fel_output_dir_var = tk.StringVar(value="")
        self.fel_lang_var = tk.StringVar(value=AUTO_COMBO_LABEL)
        self.fel_tri_prefix_var = tk.StringVar(value="")
        self.fel_date_var = tk.StringVar(value=datetime.now().strftime("%Y%m%d"))
        self.fel_quote_label_var = tk.StringVar(value=QUOTE_MODES[0][0])
        self.fel_output_mode_var = tk.StringVar(value="single_only")
        self.fel_enable_ai_var = tk.BooleanVar(value=True)
        self.fel_sync_red_var = tk.BooleanVar(value=True)
        self.fel_summary_var = tk.StringVar(
            value="当前共 0 份文档  |  支持多篇 Part 1..N 排序合并，或单文档内多 Part 表格一键合并"
        )
        self.fel_progress_var = tk.DoubleVar(value=0.0)
        self.fel_stage_var = tk.StringVar(value="状态：就绪（拖入分篇文件夹或含多 Part 表格的 .docx 稿件后点击执行）")
        self.fel_pct_timer_var = tk.StringVar(value="0%")

        self.api_key_var.trace_add("write", self._on_key_entry_changed)

    def _on_key_entry_changed(self, *_args):
        if self._syncing_key_var:
            return
        keys = parse_api_keys(self.api_key_var.get())
        self._pool_multiline_text = "\n".join(keys)
        self._sync_settings_editor_from_pool()
        self._refresh_key_pool_badge()

    def _sync_settings_editor_from_pool(self):
        if self._syncing_settings_editor or self.settings_key_editor is None:
            return
        try:
            if not self.settings_key_editor.winfo_exists():
                return
            self._syncing_settings_editor = True
            self.settings_key_editor.delete("1.0", tk.END)
            if self._pool_multiline_text.strip():
                self.settings_key_editor.insert("1.0", self._pool_multiline_text.strip() + "\n")
        except Exception:
            pass
        finally:
            self._syncing_settings_editor = False

    def _set_pool_multiline_and_sync_entry(self, raw_multiline: str):
        self._pool_multiline_text = (raw_multiline or "").strip()
        keys = parse_api_keys(self._pool_multiline_text)
        self._syncing_key_var = True
        try:
            self.api_key_var.set(" | ".join(keys))
        finally:
            self._syncing_key_var = False
        self._sync_settings_editor_from_pool()
        self._refresh_key_pool_badge()

    def _get_all_keys_list(self) -> list[str]:
        if self._pool_multiline_text.strip():
            parsed = parse_api_keys(self._pool_multiline_text)
            if parsed:
                return parsed
        return parse_api_keys(self.api_key_var.get())

    def _refresh_key_pool_badge(self):
        n = len(self._get_all_keys_list())
        self.pool_btn_text_var.set(f"API 密钥池 ({n} 个)")
        self.settings_key_count_var.set(f"当前共识别到 {n} 个不重复的 API Key")
        self._refresh_top_status_summary()

    def _refresh_top_status_summary(self):
        n = len(self._get_all_keys_list())
        gap_w = self._get_yellow_gap_words()
        mid = self._get_model_id()
        self.top_status_summary_var.set(f"模型: {mid} | Key: {n} 个 | 黄标连片: ≤{gap_w}词")

    def _refresh_fav_summary(self):
        fav_set = {c.upper() for c in self.favorite_langs if c}
        fav_labels = [
            f"★ {e['code']} - {e['name']}"
            for e in self.languages
            if e.get("favorite") or str(e.get("code", "")).upper() in fav_set
        ]
        rec_labels = []
        by_code = {str(e.get("code", "")).upper(): e for e in self.languages}
        for rc in self.recent_langs:
            rc_u = str(rc).upper()
            if rc_u not in fav_set and rc_u in by_code:
                ent = by_code[rc_u]
                rec_labels.append(f"⏱ {ent['code']} - {ent['name']}")

        if fav_labels:
            shown_fav = "  |  ".join(fav_labels[:6])
            if len(fav_labels) > 6:
                shown_fav += f" 等共 {len(fav_labels)} 个"
            fav_line = f"常用置顶（{len(fav_labels)} 个）：{shown_fav}"
        else:
            fav_line = "常用置顶：暂未设置（点击右侧按钮可将高频语种设为「★ 常用」置顶）"

        if rec_labels:
            fav_line += f"\n最近使用（自动排前）：{'  |  '.join(rec_labels[:5])}"
        self.fav_summary_var.set(fav_line)

    def _get_combo_values(self) -> list[str]:
        return get_prioritized_combo_values(
            self.languages,
            recent_codes=self.recent_langs,
            favorite_codes=self.favorite_langs,
        )

    def _refresh_all_lang_combos(self):
        vals = self._get_combo_values()
        pairs = [
            (getattr(self, "lang_cb", None), self.lang_var),
            (getattr(self, "fa_lang_cb", None), self.fa_lang_var),
            (getattr(self, "wc_lang_cb", None), self.wc_lang_var),
            (getattr(self, "fel_lang_cb", None), self.fel_lang_var),
        ]
        for cb, var in pairs:
            cur_code = _resolve_combo_code(var.get(), self.languages)
            if cb is not None:
                cb.configure(values=vals)
            var.set(self._code_to_combo_display(cur_code))
        self._refresh_fav_summary()

    def _record_recent_lang_and_refresh(self, code: str, save: bool = True):
        c = re.sub(r"^[★☆⏱\s]+", "", str(code or "")).strip().upper()
        if not c or c in ("AUTO", "XX"):
            return
        known = {str(e.get("code", "")).upper() for e in self.languages}
        if c not in known:
            return
        self.recent_langs = record_recent_language(c, self.recent_langs, max_items=5)
        self._refresh_all_lang_combos()
        if save:
            self._save_settings()

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
        for disp, mid in list(getattr(self, "dynamic_models", ())) + list(SUPPORTED_MODELS):
            if lbl == disp or lbl == mid:
                return mid
        m_br = re.search(r"\[([A-Za-z0-9][A-Za-z0-9._-]{0,63})\]\s*$", lbl)
        if m_br:
            return m_br.group(1)
        return lbl or "gemini-3.5-flash-lite"

    def _get_selected_lang_code(self) -> str:
        return _resolve_combo_code(self.lang_var.get(), self.languages)

    def _get_fa_lang_code(self) -> str:
        return _resolve_combo_code(self.fa_lang_var.get(), self.languages)

    def _get_wc_lang_code(self) -> str:
        return _resolve_combo_code(self.wc_lang_var.get(), self.languages)

    def _get_fel_lang_code(self) -> str:
        return _resolve_combo_code(self.fel_lang_var.get(), self.languages)

    def _get_fel_quote_mode_code(self) -> str:
        lbl = self.fel_quote_label_var.get().strip()
        for disp, code in QUOTE_MODES:
            if lbl == disp or lbl == code:
                return code
        return "auto"

    def _code_to_combo_display(self, code: str) -> str:
        c = re.sub(r"^[★☆⏱\s]+", "", str(code or "")).strip().upper()
        if not c or c == "AUTO":
            return AUTO_COMBO_LABEL
        for item in self._get_combo_values():
            if item != AUTO_COMBO_LABEL and _resolve_combo_code(item, self.languages) == c:
                return item
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
                if isinstance(cfg, dict):
                    self.settings = dict(cfg)
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
                elif "Vertex" in str(saved_mode):
                    self.api_mode_label.set(API_MODES[1][0])
                elif "Gemini" in str(saved_mode):
                    self.api_mode_label.set(API_MODES[2][0])
                saved_rot = cfg.get("rotation_label", "")
                if saved_rot in {m[0] for m in ROTATION_MODES}:
                    self.rotation_label.set(saved_rot)
                raw_cached_models = cfg.get("cached_models", [])
                if isinstance(raw_cached_models, list) and raw_cached_models:
                    restored_m: list[tuple[str, str]] = []
                    seen_m: set[str] = set()
                    for pair in raw_cached_models:
                        if isinstance(pair, (list, tuple)) and len(pair) == 2:
                            d_s, m_s = str(pair[0]).strip(), str(pair[1]).strip()
                            if d_s and m_s and m_s not in seen_m:
                                restored_m.append((d_s, m_s))
                                seen_m.add(m_s)
                    for d_b, m_b in SUPPORTED_MODELS:
                        if m_b not in seen_m:
                            restored_m.append((d_b, m_b))
                            seen_m.add(m_b)
                    if restored_m:
                        self.dynamic_models = restored_m
                saved_model = str(cfg.get("model_label", "")).strip()
                if saved_model:
                    # Map legacy label (e.g. 'Gemini 3.5 Flash Lite (推荐·均衡快速)') cleanly to updated label
                    mapped_lbl = saved_model
                    for d_cur, m_cur in self.dynamic_models:
                        if saved_model == m_cur or saved_model == d_cur or (
                            "3.5 Flash Lite" in saved_model and m_cur == "gemini-3.5-flash-lite"
                        ) or ("3.1 Flash Lite" in saved_model and m_cur == "gemini-3.1-flash-lite"):
                            mapped_lbl = d_cur
                            break
                    self.model_label.set(mapped_lbl)
                self.yellow_gap_words_var.set(str(self._sanitize_yellow_gap(cfg.get("yellow_gap_words", 3))))

                # Favorite & Recent Languages (Item #5)
                raw_favs = cfg.get("favorite_langs", [])
                merged_favs: list[str] = []
                for fc in list(raw_favs if isinstance(raw_favs, list) else []) + [
                    e.get("code", "") for e in self.languages if e.get("favorite")
                ]:
                    cu = str(fc or "").strip().upper()
                    if cu and cu not in ("AUTO", "XX") and cu not in merged_favs:
                        merged_favs.append(cu)
                self.favorite_langs = merged_favs
                fav_set = set(self.favorite_langs)
                for e in self.languages:
                    if str(e.get("code", "")).upper() in fav_set:
                        e["favorite"] = True

                raw_rec = cfg.get("recent_langs", [])
                clean_rec: list[str] = []
                for rc in raw_rec if isinstance(raw_rec, list) else []:
                    cu = str(rc or "").strip().upper()
                    if cu and cu not in ("AUTO", "XX") and cu not in clean_rec:
                        clean_rec.append(cu)
                self.recent_langs = clean_rec[:5]

                # Per-task paths, libraries, folders, keywords, and language selections are
                # intentionally NOT restored on fresh launch so every session starts clean (`都清空`).
                # Only general non-path workflow toggles default cleanly:
                self.dual_outputs_var.set(bool(cfg.get("dual_outputs", True)))
                self.sc1_balance_cols_var.set(bool(cfg.get("sc1_balance_cols", True)))
                self.sc1_hl_base_var.set("mark_colors")
                self.sc1_col1_underline_var.set(True)
                self._sync_sc1_hl_mode_vars()
                self.sc1_export_raw_var.set(bool(cfg.get("sc1_export_raw", True)))
                self.sc1_enable_ai_var.set(bool(cfg.get("sc1_enable_ai", True)))

                self.fa_use_titles_var.set(bool(cfg.get("fa_use_titles", True)))
                self.fa_balance_cols_var.set(bool(cfg.get("fa_balance_cols", True)))
                self.fa_hl_base_var.set("mark_colors")
                self.fa_col1_underline_var.set(True)
                self._sync_fa_hl_mode_vars()
                self.fa_enable_ai_var.set(bool(cfg.get("fa_enable_ai", True)))

                self.dd_export_deduped_var.set(bool(cfg.get("dd_export_deduped", True)))
                self.dd_balance_cols_var.set(bool(cfg.get("dd_balance_cols", True)))
                self.fel_enable_ai_var.set(bool(cfg.get("fel_enable_ai", True)))
            except Exception:
                pass

        if not loaded_key_raw:
            for env_d in _env_search_dirs():
                loaded_key_raw = load_api_key_from_env(env_d)
                if loaded_key_raw:
                    break
        self._set_pool_multiline_and_sync_entry(loaded_key_raw)

        if migrated_plaintext:
            self._save_settings()

    def _sync_sc1_hl_mode_vars(self):
        if self.sc1_hl_base_var.get() == "revisions_only":
            self.sc1_highlight_mode_var.set("revisions_only")
            if getattr(self, "sc1_col1_ul_chk", None) is not None:
                self.sc1_col1_ul_chk.configure(state=tk.DISABLED)
        else:
            self.sc1_highlight_mode_var.set("all_three" if self.sc1_col1_underline_var.get() else "en_foreign")
            if getattr(self, "sc1_col1_ul_chk", None) is not None:
                self.sc1_col1_ul_chk.configure(state=tk.NORMAL)

    def _sync_sc1_hl_mode_and_save(self):
        self._sync_sc1_hl_mode_vars()
        self._save_settings()

    def _sync_fa_hl_mode_vars(self):
        if self.fa_hl_base_var.get() == "revisions_only":
            self.fa_highlight_mode_var.set("revisions_only")
            if getattr(self, "fa_col1_ul_chk", None) is not None:
                self.fa_col1_ul_chk.configure(state=tk.DISABLED)
        else:
            self.fa_highlight_mode_var.set("all_three" if self.fa_col1_underline_var.get() else "en_foreign")
            if getattr(self, "fa_col1_ul_chk", None) is not None:
                self.fa_col1_ul_chk.configure(state=tk.NORMAL)

    def _sync_fa_hl_mode_and_save(self):
        self._sync_fa_hl_mode_vars()
        self._save_settings()

    def _save_settings(self):
        raw_to_store = self._pool_multiline_text.strip() or "\n".join(self._get_all_keys_list())
        enc_token = ""
        if self.remember_key_var.get() and raw_to_store:
            try:
                enc_token = dpapi_encrypt_str(raw_to_store)
            except Exception:
                enc_token = ""

        gap_val = self._get_yellow_gap_words()
        self._sync_sc1_hl_mode_vars()
        self._sync_fa_hl_mode_vars()
        self._refresh_top_status_summary()

        # Persist ONLY global settings (API keys, model/channel, favorites/recents, OAuth, column aliases);
        # strip all task paths/folders/libraries so every new launch opens completely clean.
        cfg = {
            "remember_key": self.remember_key_var.get(),
            "api_key_dpapi": enc_token,
            "api_mode_label": self.api_mode_label.get(),
            "rotation_label": self.rotation_label.get(),
            "model_label": self.model_label.get(),
            "cached_models": [[d, m] for d, m in getattr(self, "dynamic_models", SUPPORTED_MODELS)],
            "yellow_gap_words": gap_val,
            "favorite_langs": list(self.favorite_langs),
            "recent_langs": list(self.recent_langs),
            "dual_outputs": self.dual_outputs_var.get(),
            "sc1_balance_cols": self.sc1_balance_cols_var.get(),
            "sc1_export_raw": self.sc1_export_raw_var.get(),
            "sc1_enable_ai": self.sc1_enable_ai_var.get(),
            "fa_use_titles": self.fa_use_titles_var.get(),
            "fa_balance_cols": self.fa_balance_cols_var.get(),
            "fa_enable_ai": self.fa_enable_ai_var.get(),
            "dd_export_deduped": self.dd_export_deduped_var.get(),
            "dd_balance_cols": self.dd_balance_cols_var.get(),
            "fel_enable_ai": self.fel_enable_ai_var.get(),
        }
        for keep_key in ("ing_google_oauth_dpapi", "ing_col_mapping"):
            if keep_key in self.settings:
                cfg[keep_key] = self.settings[keep_key]
        if getattr(self, "ingest_panel", None) is not None:
            try:
                cfg.update(self.ingest_panel.export_settings_dict())
            except Exception:
                pass
        self.settings = dict(cfg)
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ---------------- UI Construction ----------------

    def _build_ui(self):
        outer = tk.Frame(self.root, bg="#ffffff")
        outer.pack(fill=tk.BOTH, expand=True)

        try:
            dpi_scale = max(1.0, float(self.root.winfo_fpixels("1i")) / 96.0)
        except Exception:
            dpi_scale = 1.0
        sidebar_w = int(round(212 * dpi_scale))

        # Left Scalable Sidebar (DPI-scaled 212pt logical width, #f8fafc surface + crisp 1px #cbd5e1 right border)
        self.sidebar = tk.Frame(outer, bg="#f8fafc", width=sidebar_w)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar.pack_propagate(False)

        tk.Frame(outer, bg="#cbd5e1", width=1).pack(side=tk.LEFT, fill=tk.Y)

        # Top Compact Brand Header
        brand_box = tk.Frame(self.sidebar, bg="#f8fafc", padx=12, pady=12)
        brand_box.pack(side=tk.TOP, fill=tk.X)
        tk.Label(
            brand_box,
            text=APP_VERSION,
            font=(MONO_FONT_FAMILY, _ui_pt(8), "bold"),
            fg="#1d4ed8",
            bg="#dbeafe",
            padx=6,
            pady=1,
        ).pack(side=tk.RIGHT)
        tk.Label(
            brand_box,
            text="多语种整理工作台",
            font=(UI_FONT_FAMILY, _ui_pt(10), "bold"),
            fg="#0f172a",
            bg="#f8fafc",
            anchor=tk.W,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Frame(self.sidebar, bg="#e2e8f0", height=1).pack(side=tk.TOP, fill=tk.X, padx=12, pady=(0, 4))

        # Bottom Anchored Footer (packed BEFORE middle scrollable nav so it NEVER clips)
        sidebar_footer = tk.Frame(self.sidebar, bg="#f8fafc", padx=12, pady=10)
        sidebar_footer.pack(side=tk.BOTTOM, fill=tk.X)
        tk.Frame(sidebar_footer, bg="#e2e8f0", height=1).pack(fill=tk.X, pady=(0, 8))
        tk.Label(
            sidebar_footer,
            text="支持拖拽文件或目录至窗口",
            font=(UI_FONT_FAMILY, _ui_pt(8)),
            fg="#475569",
            bg="#f8fafc",
            anchor=tk.W,
        ).pack(fill=tk.X, pady=(0, 6))
        self.btn_check_update = ttk.Button(
            sidebar_footer,
            text="检查软件更新",
            command=self._on_manual_check_update,
        )
        self.btn_check_update.pack(fill=tk.X)

        # Middle Scrollable Navigation List (auto-hiding scrollbar + mousewheel when hovered)
        nav_host = tk.Frame(self.sidebar, bg="#f8fafc")
        nav_host.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.nav_canvas = tk.Canvas(nav_host, bg="#f8fafc", highlightthickness=0, bd=0)
        self.nav_vsb = ttk.Scrollbar(nav_host, orient=tk.VERTICAL, command=self.nav_canvas.yview)
        self.nav_inner = tk.Frame(self.nav_canvas, bg="#f8fafc", padx=8, pady=2)
        self._nav_canvas_win = self.nav_canvas.create_window((0, 0), window=self.nav_inner, anchor="nw")
        self.nav_canvas.configure(yscrollcommand=self.nav_vsb.set)
        self.nav_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        def _sync_nav_scroll(_event=None):
            if not self.nav_canvas.winfo_exists():
                return
            self.nav_canvas.configure(scrollregion=self.nav_canvas.bbox("all"))
            req_h = int(self.nav_inner.winfo_reqheight())
            vis_h = int(self.nav_canvas.winfo_height())
            need_sb = vis_h > 24 and req_h > (vis_h + 2)
            if need_sb and not self._nav_sb_visible:
                self.nav_vsb.pack(side=tk.RIGHT, fill=tk.Y)
                self._nav_sb_visible = True
            elif not need_sb and self._nav_sb_visible:
                self.nav_vsb.pack_forget()
                self._nav_sb_visible = False

        def _on_nav_canvas_cfg(event):
            self.nav_canvas.itemconfigure(self._nav_canvas_win, width=event.width)
            _sync_nav_scroll()

        self.nav_inner.bind("<Configure>", _sync_nav_scroll)
        self.nav_canvas.bind("<Configure>", _on_nav_canvas_cfg)

        def _on_sidebar_wheel(event):
            if not self._nav_sb_visible or not self.nav_canvas.winfo_exists():
                return
            delta = int(getattr(event, "delta", 0))
            if delta:
                step = -1 if delta > 0 else 1
                self.nav_canvas.yview_scroll(step, "units")

        def _bind_sidebar_wheel(_e=None):
            self._sidebar_hover = True
            self.nav_canvas.bind_all("<MouseWheel>", _on_sidebar_wheel)

        def _unbind_sidebar_wheel(_e=None):
            self._sidebar_hover = False
            self.nav_canvas.unbind_all("<MouseWheel>")

        nav_host.bind("<Enter>", _bind_sidebar_wheel)
        nav_host.bind("<Leave>", _unbind_sidebar_wheel)

        self._create_sidebar_group_header("核心改点与查重", first=True)
        self.nav_btn_main = self._create_sidebar_tab(key="main", title="零散改点整理")
        self.nav_btn_fa = self._create_sidebar_tab(key="fa", title="全篇改点整理")
        self.nav_btn_dd = self._create_sidebar_tab(key="dd", title="改点稿查重")
        self.nav_btn_ingest = self._create_sidebar_tab(key="ingest", title="改点入库与通读定稿")

        self._create_sidebar_group_header("字数与分篇定稿", first=False)
        self.nav_btn_wc = self._create_sidebar_tab(key="wc", title="英文字数统计")
        self.nav_btn_fel = self._create_sidebar_tab(key="fel", title="分part稿整理/合并")

        self._create_sidebar_group_header("全局配置", first=False)
        self.nav_btn_settings = self._create_sidebar_tab(key="settings", title="系统设置")

        # Right Content Area
        self.content_area = ttk.Frame(outer, padding=(12, 8, 12, 10))
        self.content_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.content_area.rowconfigure(2, weight=1)
        self.content_area.columnconfigure(0, weight=1)

        # Row 0: Sleek Read-Only/Compact Global Top Bar — Active page title on left, compact summary + quick buttons on right
        self.top_toolbar = tk.Frame(
            self.content_area,
            bg="#f8fafc",
            highlightbackground="#e2e8f0",
            highlightthickness=1,
            padx=12,
            pady=5,
        )
        self.top_toolbar.grid(row=0, column=0, sticky=tk.EW, pady=(0, 8))

        # Pack right-side compact buttons FIRST so they never clip horizontally
        self.top_lang_btn = ttk.Button(
            self.top_toolbar,
            text="语种库",
            style="Compact.TButton",
            command=self._open_lang_manager,
        )
        self.top_lang_btn.pack(side=tk.RIGHT)

        self.top_settings_btn = ttk.Button(
            self.top_toolbar,
            text="系统设置",
            style="Compact.TButton",
            command=lambda: self._switch_tab("settings"),
        )
        self.top_settings_btn.pack(side=tk.RIGHT, padx=(0, 6))

        self.top_clear_cache_btn = ttk.Button(
            self.top_toolbar,
            text="🧹 清理缓存",
            style="Compact.TButton",
            command=self._on_clear_all_caches,
        )
        self.top_clear_cache_btn.pack(side=tk.RIGHT, padx=(0, 6))

        self.top_status_lbl = tk.Label(
            self.top_toolbar,
            textvariable=self.top_status_summary_var,
            font=(UI_FONT_FAMILY, _ui_pt(9)),
            fg="#475569",
            bg="#f8fafc",
            cursor="hand2",
        )
        self.top_status_lbl.pack(side=tk.RIGHT, padx=(0, 10))
        self.top_status_lbl.bind("<Button-1>", lambda _e: self._switch_tab("settings"))

        tk.Label(
            self.top_toolbar,
            textvariable=self.active_section_title_var,
            font=(UI_FONT_FAMILY, _ui_pt(10), "bold"),
            fg="#0f172a",
            bg="#f8fafc",
            anchor=tk.W,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Row 1: Sleek Top Update Notification Bar (hidden unless new version found)
        self.update_banner = tk.Frame(
            self.content_area,
            bg="#eff6ff",
            highlightbackground="#60a5fa",
            highlightthickness=1,
            padx=12,
            pady=6,
        )
        ttk.Button(
            self.update_banner,
            text="关闭",
            style="Compact.TButton",
            command=self.update_banner.grid_remove,
        ).pack(side=tk.RIGHT)

        _make_primary_btn(
            self.update_banner,
            text="查看说明并立即自动更新",
            padx=12,
            pady=3,
            command=lambda: self._open_update_dialog(self._latest_update_info),
        ).pack(side=tk.RIGHT, padx=(8, 8))

        self.update_banner_lbl = tk.Label(
            self.update_banner,
            text="发现新版本",
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
            fg="#1d4ed8",
            bg="#eff6ff",
            anchor=tk.W,
        )
        self.update_banner_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Row 2: 7 Tab Frames
        self.tab_main = ttk.Frame(self.content_area)
        self.tab_fa = ttk.Frame(self.content_area)
        self.tab_dd = ttk.Frame(self.content_area)
        self.tab_ingest = ttk.Frame(self.content_area)
        self.tab_wc = ttk.Frame(self.content_area)
        self.tab_fel = ttk.Frame(self.content_area)
        self.tab_settings = ttk.Frame(self.content_area)

        for tf in (self.tab_main, self.tab_fa, self.tab_dd, self.tab_ingest, self.tab_wc, self.tab_fel, self.tab_settings):
            tf.grid(row=2, column=0, sticky=tk.NSEW)

        self._build_tab1_pipeline(self.tab_main)
        self._build_tab2_full_article(self.tab_fa)
        self._build_tab3_dedup(self.tab_dd)
        self.ingest_panel = IngestTabPanel(self, self.tab_ingest) if IngestTabPanel is not None else None
        self._build_tab4_word_counter(self.tab_wc)
        self._build_tab5_part_merge(self.tab_fel)
        self._build_tab6_settings(self.tab_settings)

        self._switch_tab("main")

    def _create_sidebar_group_header(self, text: str, first: bool = False):
        top_pad = 6 if first else 14
        tk.Label(
            self.nav_inner,
            text=text,
            font=("Microsoft YaHei UI", 8, "bold"),
            fg="#475569",
            bg="#f8fafc",
            anchor=tk.W,
            padx=10,
        ).pack(fill=tk.X, pady=(top_pad, 4))

    def _create_sidebar_tab(self, key: str, title: str, subtitle: str = "") -> tk.Frame:
        item = tk.Frame(self.nav_inner, bg="#f8fafc", cursor="hand2")
        item.pack(fill=tk.X, pady=1)

        bar = tk.Frame(item, bg="#f8fafc", width=3)
        bar.pack(side=tk.LEFT, fill=tk.Y)

        lbl_title = tk.Label(
            item,
            text=title,
            font=(UI_FONT_FAMILY, _ui_pt(9)),
            fg="#334155",
            bg="#f8fafc",
            anchor=tk.W,
            padx=12,
            pady=8,
            cursor="hand2",
        )
        lbl_title.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self._nav_items[key] = {
            "card": item,
            "bar": bar,
            "title": lbl_title,
        }

        def _on_click(_e=None):
            self._switch_tab(key)

        def _on_enter(_e=None):
            if self.active_tab_key != key:
                item.configure(bg="#f1f5f9")
                lbl_title.configure(bg="#f1f5f9", fg="#0f172a")
                bar.configure(bg="#cbd5e1")

        def _on_leave(_e=None):
            if self.active_tab_key != key:
                item.configure(bg="#f8fafc")
                lbl_title.configure(bg="#f8fafc", fg="#334155")
                bar.configure(bg="#f8fafc")

        for w in (item, bar, lbl_title):
            w.bind("<Button-1>", _on_click)
        item.bind("<Enter>", _on_enter)
        item.bind("<Leave>", _on_leave)
        return item

    def _switch_tab(self, key: str):
        self.active_tab_key = key
        self.active_section_title_var.set(self._SECTION_TITLES.get(key, "多语种整理工作台"))
        tab_map = {
            "main": self.tab_main,
            "fa": self.tab_fa,
            "dd": self.tab_dd,
            "ingest": self.tab_ingest,
            "wc": self.tab_wc,
            "fel": self.tab_fel,
            "settings": self.tab_settings,
        }
        for k, frame in tab_map.items():
            if k == key:
                frame.grid(row=2, column=0, sticky=tk.NSEW)
            else:
                frame.grid_remove()

        for k, parts in self._nav_items.items():
            is_act = k == key
            bg = "#eff6ff" if is_act else "#f8fafc"
            bar_c = "#2563eb" if is_act else "#f8fafc"
            fg_t = "#1d4ed8" if is_act else "#334155"
            fnt = (UI_FONT_FAMILY, _ui_pt(9), "bold") if is_act else (UI_FONT_FAMILY, _ui_pt(9))

            parts["card"].configure(bg=bg)
            parts["bar"].configure(bg=bar_c)
            parts["title"].configure(bg=bg, fg=fg_t, font=fnt)

    # ---------------- Auto-Update Banner & Dialog Triggers ----------------

    def _show_update_banner(self, info: dict):
        self._latest_update_info = info
        v_str = str(info.get("latest_version") or "").lstrip("vV")
        if not v_str:
            return
        self.update_banner_lbl.configure(text=f"发现新版本 v{v_str}（支持一键校验 SHA-256 自动更新）")
        self.update_banner.grid(row=1, column=0, sticky=tk.EW, pady=(0, 8))

    def _start_auto_update_check_async(self):
        if not callable(check_for_updates):
            return

        def _bg():
            try:
                info = check_for_updates(current_version=RAW_APP_VERSION)
                if isinstance(info, dict) and info.get("has_update"):
                    self.log_queue.put(("update_available", info))
            except Exception:
                pass

        threading.Thread(target=_bg, daemon=True).start()

    def _open_update_dialog(self, initial_info: dict | None = None):
        SoftwareUpdateDialog(self.root, app=self, initial_info=initial_info)

    def _on_manual_check_update(self):
        self._open_update_dialog(initial_info=None)

    @staticmethod
    def _make_interactive_drop_zone(parent: tk.Misc, text: str, on_click) -> tk.Frame:
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
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
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

    # ---------------- Tab 1: 零散改点整理 ----------------

    def _build_tab1_pipeline(self, main: ttk.Frame):
        _pw, upper, lower = _make_vertical_paned(
            main, upper_weight=0, lower_weight=1, min_top=250, min_bottom=100
        )

        # 1. Paths, Naming, Column Balance & Dedup Options Frame
        path_frame = ttk.LabelFrame(
            upper,
            text=" 1. 文档、总库路径、最新中英底稿核对与命名/查重设置（支持拖拽文件或文件夹到对应输入框） ",
            padding=8,
        )
        path_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))
        path_frame.columnconfigure(1, weight=1)

        self.lbl_input_docx = ttk.Label(path_frame, text="待整理文档/目录：")
        self.lbl_input_docx.grid(row=0, column=0, sticky=tk.W, pady=2)
        self.ent_input_docx = ttk.Entry(path_frame, textvariable=self.input_docx_var)
        self.ent_input_docx.grid(row=0, column=1, sticky=tk.EW, pady=2, padx=4)
        inp_btns = ttk.Frame(path_frame)
        inp_btns.grid(row=0, column=2, columnspan=2, sticky=tk.W)
        ttk.Button(inp_btns, text="选文件...", style="Compact.TButton", command=self._browse_input).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(inp_btns, text="选文件夹...", style="Compact.TButton", command=self._browse_input_dir).pack(
            side=tk.LEFT, padx=2
        )
        self.btn_refresh_input = ttk.Button(
            inp_btns,
            text="刷新",
            style="Compact.TButton",
            command=self._refresh_input_docx_status,
        )
        self.btn_refresh_input.pack(side=tk.LEFT, padx=2)

        self.lbl_library_dir = ttk.Label(path_frame, text="对应语种总库目录：")
        self.lbl_library_dir.grid(row=1, column=0, sticky=tk.W, pady=2)
        self.ent_library_dir = ttk.Entry(path_frame, textvariable=self.library_dir_var)
        self.ent_library_dir.grid(row=1, column=1, sticky=tk.EW, pady=2, padx=4)
        self.ent_library_dir.bind("<Return>", self._on_library_entry_committed)
        self.ent_library_dir.bind("<FocusOut>", self._on_library_entry_committed)
        lib_btns = ttk.Frame(path_frame)
        lib_btns.grid(row=1, column=2, columnspan=2, sticky=tk.W)
        ttk.Button(lib_btns, text="浏览...", style="Compact.TButton", command=self._browse_library).pack(
            side=tk.LEFT, padx=2
        )
        self.btn_refresh_lib = ttk.Button(
            lib_btns,
            text="刷新总库",
            style="Compact.TButton",
            command=lambda: self._refresh_library_index_manual(self.library_dir_var.get(), target_tab="main"),
        )
        self.btn_refresh_lib.pack(side=tk.LEFT, padx=2)

        self.lbl_output_dir = ttk.Label(path_frame, text="输出保存目录：")
        self.lbl_output_dir.grid(row=2, column=0, sticky=tk.W, pady=2)
        self.ent_output_dir = ttk.Entry(path_frame, textvariable=self.output_dir_var)
        self.ent_output_dir.grid(row=2, column=1, sticky=tk.EW, pady=2, padx=4)
        out_btns = ttk.Frame(path_frame)
        out_btns.grid(row=2, column=2, columnspan=2, sticky=tk.W)
        ttk.Button(out_btns, text="浏览...", style="Compact.TButton", command=self._browse_output).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(out_btns, text="打开目录", style="Compact.TButton", command=self._open_output_dir).pack(
            side=tk.LEFT, padx=2
        )

        # Item #6: Sync Pending Manuscript with Latest CN-EN Master (File or Folder)
        self.chk_sc1_master = ttk.Checkbutton(
            path_frame,
            text="核对最新中英底稿：",
            variable=self.sc1_sync_master_var,
            command=self._save_settings,
        )
        self.chk_sc1_master.grid(row=3, column=0, sticky=tk.W, pady=2)
        self.ent_sc1_master_path = ttk.Entry(path_frame, textvariable=self.sc1_master_path_var)
        self.ent_sc1_master_path.grid(row=3, column=1, sticky=tk.EW, pady=2, padx=4)
        master_btns = ttk.Frame(path_frame)
        master_btns.grid(row=3, column=2, columnspan=2, sticky=tk.W)
        ttk.Button(
            master_btns,
            text="选文件...",
            style="Compact.TButton",
            command=self._browse_sc1_master_file,
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            master_btns,
            text="选文件夹...",
            style="Compact.TButton",
            command=self._browse_sc1_master_dir,
        ).pack(side=tk.LEFT, padx=2)

        # Optional Dedup Package Row
        self.chk_sc1_dedup = ttk.Checkbutton(
            path_frame,
            text="关联「未入库改点包」查重：",
            variable=self.sc1_enable_dedup_var,
            command=self._save_settings,
        )
        self.chk_sc1_dedup.grid(row=4, column=0, sticky=tk.W, pady=2)
        self.ent_sc1_dedup_dir = ttk.Entry(path_frame, textvariable=self.sc1_dedup_dir_var)
        self.ent_sc1_dedup_dir.grid(row=4, column=1, sticky=tk.EW, pady=2, padx=4)
        dedup_btns = ttk.Frame(path_frame)
        dedup_btns.grid(row=4, column=2, columnspan=2, sticky=tk.W)
        ttk.Button(dedup_btns, text="浏览...", style="Compact.TButton", command=self._browse_sc1_dedup_dir).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Label(dedup_btns, text="自动剔除发译者版重复段", foreground="#475569").pack(side=tk.LEFT, padx=4)

        # Target Language + Core Output Options Row
        opt_row1 = ttk.Frame(path_frame)
        opt_row1.grid(row=5, column=0, columnspan=4, sticky=tk.EW, pady=(4, 1))

        ttk.Label(opt_row1, text="目标语种：").pack(side=tk.LEFT)
        self.lang_cb = ttk.Combobox(
            opt_row1,
            textvariable=self.lang_var,
            values=self._get_combo_values(),
            width=28,
        )
        self.lang_cb.pack(side=tk.LEFT, padx=(4, 6))
        self.lang_cb.bind("<<ComboboxSelected>>", self._on_tab1_lang_selected)
        ttk.Button(
            opt_row1,
            text="管理语种...",
            style="Compact.TButton",
            command=self._open_lang_manager,
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Checkbutton(
            opt_row1,
            text="三语列宽自动拉齐",
            variable=self.sc1_balance_cols_var,
            command=self._save_settings,
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Checkbutton(
            opt_row1,
            text="同时输出完整留底版与无改点版",
            variable=self.dual_outputs_var,
            command=self._save_settings,
        ).pack(side=tk.LEFT)

        # Highlight Mode + Default Col 1 Chinese Underline Checkbox + Raw Trilingual Export Checkbox
        opt_row_hl = ttk.Frame(path_frame)
        opt_row_hl.grid(row=6, column=0, columnspan=4, sticky=tk.EW, pady=1)
        ttk.Label(opt_row_hl, text="标色选项：").pack(side=tk.LEFT)
        ttk.Radiobutton(
            opt_row_hl,
            text="英外标色（英黄/青 ｜ 外浅蓝底纹/浅灰）",
            value="mark_colors",
            variable=self.sc1_hl_base_var,
            command=self._sync_sc1_hl_mode_and_save,
        ).pack(side=tk.LEFT, padx=(2, 6))
        self.sc1_col1_ul_chk = ttk.Checkbutton(
            opt_row_hl,
            text="中文列对应改点标记下划线",
            variable=self.sc1_col1_underline_var,
            command=self._sync_sc1_hl_mode_and_save,
        )
        self.sc1_col1_ul_chk.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(
            opt_row_hl,
            text="仅输出对比修订",
            value="revisions_only",
            variable=self.sc1_hl_base_var,
            command=self._sync_sc1_hl_mode_and_save,
        ).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(
            opt_row_hl,
            text="导出对比前的「提取三语原稿」供核对",
            variable=self.sc1_export_raw_var,
            command=self._save_settings,
        ).pack(side=tk.LEFT)

        # Naming Mode Selector Row
        opt_row_name = ttk.Frame(path_frame)
        opt_row_name.grid(row=7, column=0, columnspan=4, sticky=tk.EW, pady=1)

        ttk.Label(opt_row_name, text="命名模式：").pack(side=tk.LEFT)
        ttk.Radiobutton(
            opt_row_name,
            text="零散标准命名 (God's words [in 关键词])",
            value="scattered",
            variable=self.sc1_naming_mode_var,
            command=self._save_settings,
        ).pack(side=tk.LEFT, padx=(4, 6))
        ttk.Label(opt_row_name, text="关键词(可留空)：", foreground="#475569").pack(side=tk.LEFT)
        ent_kw = ttk.Entry(opt_row_name, textvariable=self.sc1_keyword_var, width=12)
        ent_kw.pack(side=tk.LEFT, padx=(2, 14))
        ent_kw.bind("<KeyRelease>", self._on_sc1_keyword_typed)

        ttk.Radiobutton(
            opt_row_name,
            text="自定义 / 原文档名：",
            value="keep",
            variable=self.sc1_naming_mode_var,
            command=self._save_settings,
        ).pack(side=tk.LEFT, padx=(0, 4))
        ent_custom = ttk.Entry(opt_row_name, textvariable=self.doc_name_var, width=18)
        ent_custom.pack(side=tk.LEFT)
        ent_custom.bind("<KeyRelease>", self._on_sc1_custom_name_typed)

        opt_row3 = ttk.Frame(path_frame)
        opt_row3.grid(row=8, column=0, columnspan=4, sticky=tk.EW, pady=(1, 0))
        ttk.Checkbutton(
            opt_row3,
            text="使用 AI 辅助分级与中/外双列精准对齐（若未填 Key 或额度受限，自动切换本地位置估算标色）",
            variable=self.sc1_enable_ai_var,
            command=self._save_settings,
        ).pack(side=tk.LEFT)

        # 2. One-Click Primary Action & Determinate Progress Frame
        act_frame = ttk.LabelFrame(
            upper,
            text=" 2. 一键执行与实时进度（最新底稿无损核对 · 三列同步标色 · 原稿蓝/红字逐字保留 · 单文件平铺输出） ",
            padding=8,
        )
        act_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 4))

        btn_row = ttk.Frame(act_frame)
        btn_row.pack(fill=tk.X)
        btn_row.columnconfigure(0, weight=1)

        self.btn_oneclick = _make_primary_btn(
            btn_row,
            text="一键输出定稿",
            font=(UI_FONT_FAMILY, _ui_pt(10), "bold"),
            pady=7,
            command=self._on_run_oneclick,
        )
        self.btn_oneclick.grid(row=0, column=0, sticky=tk.EW)
        self.btn_all = self.btn_oneclick

        prog_info = ttk.Frame(act_frame)
        prog_info.pack(fill=tk.X, pady=(6, 2))

        self.pct_lbl = ttk.Label(
            prog_info,
            textvariable=self.pct_timer_var,
            font=("Consolas", 9, "bold"),
            foreground="#0f172a",
        )
        self.pct_lbl.pack(side=tk.RIGHT, padx=(8, 0))

        self.stage_lbl = ttk.Label(
            prog_info,
            textvariable=self.stage_var,
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
            foreground="#1d4ed8",
        )
        self.stage_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.progress = ttk.Progressbar(
            act_frame,
            mode="determinate",
            maximum=100.0,
            variable=self.progress_var,
        )
        self.progress.pack(fill=tk.X)

        # 3. Log Frame (Lower Resizable Pane)
        log_frame = ttk.LabelFrame(lower, text=" 3. 实时运行日志（可拖动上方分隔条调节高度） ", padding=6)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(2, 0))

        log_top = ttk.Frame(log_frame)
        log_top.pack(fill=tk.X, pady=(0, 4))
        ttk.Button(
            log_top,
            text="清空日志",
            style="Compact.TButton",
            command=lambda: self._clear_log_widget(self.log_text),
        ).pack(side=tk.RIGHT)
        ttk.Label(
            log_top,
            text="后台总库索引、最新中英底稿同步、未入库查重、多 Key 轮换及输出路径实时显示：",
            foreground="#475569",
        ).pack(side=tk.LEFT)

        log_body = ttk.Frame(log_frame)
        log_body.pack(fill=tk.BOTH, expand=True)

        self.log_text = self._create_styled_log_text(log_body, height=6)

        key_cnt = len(self._get_all_keys_list())
        self._append_log(
            f"欢迎使用多语种整理工作台 {APP_VERSION}！已载入 {key_cnt} 个 API Key、{len(self.languages)} 种语言；可在左侧「系统设置」配置密钥池、模型与黄标连片间隔。",
            target_tab="main",
        )

    # ---------------- Tab 2: 全篇改点整理 ----------------

    def _build_tab2_full_article(self, parent: ttk.Frame):
        _pw, upper, lower = _make_vertical_paned(
            parent, upper_weight=0, lower_weight=1, min_top=300, min_bottom=95
        )

        # Anchor Step 3 & Step 2 at bottom of `upper` so dragging sash down expands `fa_titles_text` cleanly
        act_frame = ttk.LabelFrame(
            upper,
            text=" 3. 全篇对比执行（单篇平铺输出 / 多篇自动分包 · 无改点篇目自动汇总报告） ",
            padding=8,
        )
        act_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 4))

        btn_row = ttk.Frame(act_frame)
        btn_row.pack(fill=tk.X)
        btn_row.columnconfigure(0, weight=1)
        self.btn_fa_run = _make_primary_btn(
            btn_row,
            text="开始全篇提取对比（逐篇输出完整版 + 仅改点版）",
            font=(UI_FONT_FAMILY, _ui_pt(10), "bold"),
            pady=7,
            command=self._on_run_full_article,
        )
        self.btn_fa_run.grid(row=0, column=0, sticky=tk.EW)

        p_info = ttk.Frame(act_frame)
        p_info.pack(fill=tk.X, pady=(6, 2))
        ttk.Label(
            p_info,
            textvariable=self.fa_pct_timer_var,
            font=("Consolas", 9, "bold"),
            foreground="#0f172a",
        ).pack(side=tk.RIGHT, padx=(8, 0))
        self.fa_stage_lbl = ttk.Label(
            p_info,
            textvariable=self.fa_stage_var,
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
            foreground="#1d4ed8",
        )
        self.fa_stage_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Progressbar(
            act_frame,
            mode="determinate",
            maximum=100.0,
            variable=self.fa_progress_var,
        ).pack(fill=tk.X)

        # 2. Directories & Options Frame
        cfg_frame = ttk.LabelFrame(
            upper,
            text=" 2. 最新中英底稿库、外文三语总库与输出设置（支持直接拖拽文件夹到各输入框） ",
            padding=8,
        )
        cfg_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 6))
        cfg_frame.columnconfigure(1, weight=1)

        self.lbl_fa_master_dir = ttk.Label(cfg_frame, text="最新中英底稿目录：")
        self.lbl_fa_master_dir.grid(row=0, column=0, sticky=tk.W, pady=2)
        self.ent_fa_master_dir = ttk.Entry(cfg_frame, textvariable=self.fa_master_dir_var)
        self.ent_fa_master_dir.grid(row=0, column=1, sticky=tk.EW, pady=2, padx=4)
        fa_master_btns = ttk.Frame(cfg_frame)
        fa_master_btns.grid(row=0, column=2, columnspan=2, sticky=tk.W)
        ttk.Button(fa_master_btns, text="浏览...", style="Compact.TButton", command=self._fa_browse_master_dir).pack(
            side=tk.LEFT, padx=2
        )
        self.btn_fa_refresh_master = ttk.Button(
            fa_master_btns,
            text="刷新底稿",
            style="Compact.TButton",
            command=self._refresh_master_index_manual,
        )
        self.btn_fa_refresh_master.pack(side=tk.LEFT, padx=2)

        self.lbl_fa_library_dir = ttk.Label(cfg_frame, text="外文三语总库目录：")
        self.lbl_fa_library_dir.grid(row=1, column=0, sticky=tk.W, pady=2)
        self.ent_fa_library_dir = ttk.Entry(cfg_frame, textvariable=self.fa_library_dir_var)
        self.ent_fa_library_dir.grid(row=1, column=1, sticky=tk.EW, pady=2, padx=4)
        self.ent_fa_library_dir.bind("<Return>", self._on_fa_library_entry_committed)
        self.ent_fa_library_dir.bind("<FocusOut>", self._on_fa_library_entry_committed)
        fa_lib_btns = ttk.Frame(cfg_frame)
        fa_lib_btns.grid(row=1, column=2, columnspan=2, sticky=tk.W)
        ttk.Button(fa_lib_btns, text="浏览...", style="Compact.TButton", command=self._fa_browse_library_dir).pack(
            side=tk.LEFT, padx=2
        )
        self.btn_fa_refresh_lib = ttk.Button(
            fa_lib_btns,
            text="刷新总库",
            style="Compact.TButton",
            command=lambda: self._refresh_library_index_manual(self.fa_library_dir_var.get(), target_tab="fa"),
        )
        self.btn_fa_refresh_lib.pack(side=tk.LEFT, padx=2)

        self.lbl_fa_output_dir = ttk.Label(cfg_frame, text="输出保存目录：")
        self.lbl_fa_output_dir.grid(row=2, column=0, sticky=tk.W, pady=2)
        self.ent_fa_output_dir = ttk.Entry(cfg_frame, textvariable=self.fa_output_dir_var)
        self.ent_fa_output_dir.grid(row=2, column=1, sticky=tk.EW, pady=2, padx=4)
        fa_out_btns = ttk.Frame(cfg_frame)
        fa_out_btns.grid(row=2, column=2, columnspan=2, sticky=tk.W)
        ttk.Button(fa_out_btns, text="浏览...", style="Compact.TButton", command=self._fa_browse_output_dir).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(fa_out_btns, text="打开目录", style="Compact.TButton", command=self._fa_open_output_dir).pack(
            side=tk.LEFT, padx=2
        )

        self.chk_fa_dedup = ttk.Checkbutton(
            cfg_frame,
            text="关联「未入库改点包」自动查重剔除：",
            variable=self.fa_enable_dedup_var,
            command=self._save_settings,
        )
        self.chk_fa_dedup.grid(row=3, column=0, sticky=tk.W, pady=2)
        self.ent_fa_dedup_dir = ttk.Entry(cfg_frame, textvariable=self.fa_dedup_dir_var)
        self.ent_fa_dedup_dir.grid(row=3, column=1, sticky=tk.EW, pady=2, padx=4)
        fa_dd_btns = ttk.Frame(cfg_frame)
        fa_dd_btns.grid(row=3, column=2, columnspan=2, sticky=tk.W)
        ttk.Button(fa_dd_btns, text="浏览...", style="Compact.TButton", command=self._fa_browse_dedup_dir).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Label(fa_dd_btns, text="仅改点版自动去重", foreground="#475569").pack(side=tk.LEFT, padx=4)

        opt_row1 = ttk.Frame(cfg_frame)
        opt_row1.grid(row=4, column=0, columnspan=4, sticky=tk.EW, pady=(4, 1))
        ttk.Label(opt_row1, text="目标语种：").pack(side=tk.LEFT)
        self.fa_lang_cb = ttk.Combobox(
            opt_row1,
            textvariable=self.fa_lang_var,
            values=self._get_combo_values(),
            width=28,
        )
        self.fa_lang_cb.pack(side=tk.LEFT, padx=(4, 6))
        self.fa_lang_cb.bind("<<ComboboxSelected>>", self._on_fa_lang_selected)
        ttk.Button(opt_row1, text="管理语种...", style="Compact.TButton", command=self._open_lang_manager).pack(
            side=tk.LEFT, padx=(0, 14)
        )

        ttk.Label(opt_row1, text="日期(YYYYMMDD)：").pack(side=tk.LEFT)
        ttk.Entry(opt_row1, textvariable=self.fa_date_var, width=10).pack(side=tk.LEFT, padx=(4, 14))

        ttk.Checkbutton(
            opt_row1,
            text="三语列宽自动拉齐",
            variable=self.fa_balance_cols_var,
            command=self._save_settings,
        ).pack(side=tk.LEFT)

        opt_row2 = ttk.Frame(cfg_frame)
        opt_row2.grid(row=5, column=0, columnspan=4, sticky=tk.EW, pady=(1, 0))
        ttk.Label(opt_row2, text="标色选项：").pack(side=tk.LEFT)
        ttk.Radiobutton(
            opt_row2,
            text="英外标色（英黄/青 ｜ 外浅蓝底纹/浅灰）",
            value="mark_colors",
            variable=self.fa_hl_base_var,
            command=self._sync_fa_hl_mode_and_save,
        ).pack(side=tk.LEFT, padx=(2, 6))
        self.fa_col1_ul_chk = ttk.Checkbutton(
            opt_row2,
            text="中文列对应改点标记下划线",
            variable=self.fa_col1_underline_var,
            command=self._sync_fa_hl_mode_and_save,
        )
        self.fa_col1_ul_chk.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(
            opt_row2,
            text="仅输出对比修订",
            value="revisions_only",
            variable=self.fa_hl_base_var,
            command=self._sync_fa_hl_mode_and_save,
        ).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(
            opt_row2,
            text="使用 AI 辅助分级与精准对齐（无 Key 自动本地标色）",
            variable=self.fa_enable_ai_var,
            command=self._save_settings,
        ).pack(side=tk.LEFT)

        # 1. Mode Switch & Multiline Titles Input Frame (expands when sash is dragged down)
        title_frame = ttk.LabelFrame(
            upper,
            text=" 1. 全篇对比模式与待提取篇目标题（支持粘贴/导入标题清单，或取消勾选直接对比两个现成文件包） ",
            padding=8,
        )
        title_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 6))

        mode_bar = ttk.Frame(title_frame)
        mode_bar.pack(side=tk.TOP, fill=tk.X, pady=(0, 4))
        self.chk_fa_use_titles = ttk.Checkbutton(
            mode_bar,
            text="按标题列表提取篇目（取消勾选则直接将「最新中英底稿包」与「外文三语文件包」按顺序逐篇对比）",
            variable=self.fa_use_titles_var,
            command=self._on_fa_mode_toggled,
        )
        self.chk_fa_use_titles.pack(side=tk.LEFT)

        t_bar = ttk.Frame(title_frame)
        t_bar.pack(side=tk.TOP, fill=tk.X, pady=(0, 4))
        self.btn_fa_import_titles = ttk.Button(
            t_bar,
            text="导入标题文档 (.docx / .txt)...",
            style="Action.TButton",
            command=self._fa_import_titles_file,
        )
        self.btn_fa_import_titles.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_fa_clear_titles = ttk.Button(t_bar, text="清空", command=self._fa_clear_titles)
        self.btn_fa_clear_titles.pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(
            t_bar,
            textvariable=self.fa_title_count_var,
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
            foreground="#1d4ed8",
        ).pack(side=tk.LEFT)

        ed_wrap = ttk.Frame(title_frame)
        ed_wrap.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.fa_titles_text = tk.Text(
            ed_wrap,
            height=3,
            wrap=tk.WORD,
            font=(UI_FONT_FAMILY, _ui_pt(9)),
            bg="#ffffff",
            fg="#0f172a",
            relief=tk.SOLID,
            borderwidth=1,
            padx=8,
            pady=5,
            undo=True,
        )
        fa_tvsb = ttk.Scrollbar(ed_wrap, orient=tk.VERTICAL, command=self.fa_titles_text.yview)
        self.fa_titles_text.configure(yscrollcommand=fa_tvsb.set)
        self.fa_titles_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        fa_tvsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.fa_titles_text.bind("<KeyRelease>", lambda _e: self._fa_update_title_count())
        self.fa_titles_text.bind("<<Paste>>", lambda _e: self.root.after(20, self._fa_update_title_count))

        # 4. Real-Time Log (Lower Resizable Pane)
        log_frame = ttk.LabelFrame(lower, text=" 4. 全篇提取对比实时日志（可拖动上方分隔条调节高度） ", padding=6)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(2, 0))
        log_top = ttk.Frame(log_frame)
        log_top.pack(fill=tk.X, pady=(0, 4))
        ttk.Button(
            log_top,
            text="清空日志",
            style="Compact.TButton",
            command=lambda: self._clear_log_widget(self.fa_log_text),
        ).pack(side=tk.RIGHT)
        ttk.Label(
            log_top,
            text="底稿定位、同名消歧、逐篇全段对比、无改点篇目汇总报告与分包导出详情：",
            foreground="#475569",
        ).pack(side=tk.LEFT)
        log_body = ttk.Frame(log_frame)
        log_body.pack(fill=tk.BOTH, expand=True)
        self.fa_log_text = self._create_styled_log_text(log_body, height=5)
        self.fa_log_text.insert(
            tk.END,
            "提示：支持粘贴带编号标题（如“3-1-39 032 七雷巨响——预言国度的福音将扩展全宇（审判）”）提取对比，或取消勾选上方模式直接对比两个现成文件夹。\n",
            "dim",
        )
        self._apply_fa_mode_ui()

    # ---------------- Tab 3: 改点稿查重 ----------------

    def _build_tab3_dedup(self, parent: ttk.Frame):
        _pw, upper, lower = _make_vertical_paned(
            parent, upper_weight=1, lower_weight=1, min_top=220, min_bottom=150
        )

        cfg_frame = ttk.LabelFrame(
            upper,
            text=" 1. 待查重文档列表与「未入库改点包」路径设置（依据第 1 列中文：整段 -> 段首 -> 段尾 -> 着色片段四级极速查重） ",
            padding=8,
        )
        cfg_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))

        drop_zone = self._make_interactive_drop_zone(
            cfg_frame,
            "直接拖入待查重的 .docx 文档或文件夹即可载入列表（点击此处也可选择待查重文件）",
            self._dd_add_files,
        )
        drop_zone.pack(fill=tk.X, pady=(0, 6))

        btn_bar = ttk.Frame(cfg_frame)
        btn_bar.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(
            btn_bar,
            text="选择待查重文档...",
            style="Action.TButton",
            command=self._dd_add_files,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            btn_bar,
            text="选择待查重文件夹...",
            style="Action.TButton",
            command=self._dd_add_folder,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_bar, text="移除选中", command=self._dd_remove_selected).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_bar, text="清空列表", command=self._dd_clear_all).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(
            btn_bar,
            textvariable=self.dd_summary_var,
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
            foreground="#1d4ed8",
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        paths_grid = ttk.Frame(cfg_frame)
        paths_grid.pack(fill=tk.X)
        paths_grid.columnconfigure(1, weight=1)

        self.lbl_dd_pkg_dir = ttk.Label(paths_grid, text="未入库改点包目录：")
        self.lbl_dd_pkg_dir.grid(row=0, column=0, sticky=tk.W, pady=3)
        self.ent_dd_pkg_dir = ttk.Entry(paths_grid, textvariable=self.dd_pkg_dir_var)
        self.ent_dd_pkg_dir.grid(row=0, column=1, sticky=tk.EW, pady=3, padx=4)
        ttk.Button(paths_grid, text="浏览...", style="Compact.TButton", command=self._dd_browse_pkg_dir).grid(
            row=0, column=2, sticky=tk.W, padx=2
        )

        self.lbl_dd_out_dir = ttk.Label(paths_grid, text="查重报告与去重稿输出目录：")
        self.lbl_dd_out_dir.grid(row=1, column=0, sticky=tk.W, pady=3)
        self.ent_dd_out_dir = ttk.Entry(paths_grid, textvariable=self.dd_out_dir_var)
        self.ent_dd_out_dir.grid(row=1, column=1, sticky=tk.EW, pady=3, padx=4)
        dd_out_btns = ttk.Frame(paths_grid)
        dd_out_btns.grid(row=1, column=2, columnspan=2, sticky=tk.W)
        ttk.Button(dd_out_btns, text="浏览...", style="Compact.TButton", command=self._dd_browse_out_dir).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(dd_out_btns, text="打开目录", style="Compact.TButton", command=self._dd_open_out_dir).pack(
            side=tk.LEFT, padx=2
        )

        opt_row = ttk.Frame(cfg_frame)
        opt_row.pack(fill=tk.X, pady=(5, 0))
        ttk.Checkbutton(
            opt_row,
            text="同时导出剔除重复段落后的发译者定稿（自动重算字数并更新文件名）",
            variable=self.dd_export_deduped_var,
            command=self._save_settings,
        ).pack(side=tk.LEFT, padx=(0, 16))
        ttk.Checkbutton(
            opt_row,
            text="三语列宽自动拉齐",
            variable=self.dd_balance_cols_var,
            command=self._save_settings,
        ).pack(side=tk.LEFT, padx=(0, 16))
        ttk.Label(opt_row, text="日期：").pack(side=tk.LEFT)
        ttk.Entry(opt_row, textvariable=self.dd_date_var, width=10).pack(side=tk.LEFT, padx=(3, 0))

        # 2. Pending Files Treeview
        tbl_box = ttk.LabelFrame(
            upper,
            text=" 2. 待查重 Word 文档队列（双击行可打开所在文件夹） ",
            padding=6,
        )
        tbl_box.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 4))

        cols = ("seq", "filename", "folder", "status")
        self.dd_tree = ttk.Treeview(tbl_box, columns=cols, show="headings", selectmode="extended", height=5)
        self.dd_tree.heading("seq", text="序号")
        self.dd_tree.heading("filename", text="待查重文档名")
        self.dd_tree.heading("folder", text="所在目录")
        self.dd_tree.heading("status", text="状态")

        self.dd_tree.column("seq", width=55, anchor=tk.CENTER)
        self.dd_tree.column("filename", width=340, anchor=tk.W)
        self.dd_tree.column("folder", width=380, anchor=tk.W)
        self.dd_tree.column("status", width=140, anchor=tk.CENTER)

        dd_vsb = ttk.Scrollbar(tbl_box, orient=tk.VERTICAL, command=self.dd_tree.yview)
        self.dd_tree.configure(yscrollcommand=dd_vsb.set)
        self.dd_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        dd_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.dd_tree.bind("<Double-1>", self._dd_open_selected_folder)

        # 3. Execute & Log Frame (Lower Resizable Pane)
        act_box = ttk.LabelFrame(lower, text=" 3. 四级中文查重执行与日志汇总（可拖动上方分隔条调节高度） ", padding=8)
        act_box.pack(fill=tk.BOTH, expand=True, pady=(2, 0))

        run_row = ttk.Frame(act_box)
        run_row.pack(side=tk.TOP, fill=tk.X)
        run_row.columnconfigure(0, weight=1)
        self.btn_dd_run = _make_primary_btn(
            run_row,
            text="开始第一列中文四级查重并导出报告",
            font=(UI_FONT_FAMILY, _ui_pt(10), "bold"),
            pady=7,
            command=self._on_run_dedup,
        )
        self.btn_dd_run.grid(row=0, column=0, sticky=tk.EW)

        p_info = ttk.Frame(act_box)
        p_info.pack(side=tk.TOP, fill=tk.X, pady=(6, 2))
        ttk.Label(
            p_info,
            textvariable=self.dd_pct_timer_var,
            font=("Consolas", 9, "bold"),
            foreground="#0f172a",
        ).pack(side=tk.RIGHT, padx=(8, 0))
        self.dd_stage_lbl = ttk.Label(
            p_info,
            textvariable=self.dd_stage_var,
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
            foreground="#1d4ed8",
        )
        self.dd_stage_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Progressbar(
            act_box,
            mode="determinate",
            maximum=100.0,
            variable=self.dd_progress_var,
        ).pack(side=tk.TOP, fill=tk.X, pady=(0, 5))

        dd_log_bar = ttk.Frame(act_box)
        dd_log_bar.pack(side=tk.TOP, fill=tk.X, pady=(0, 3))
        ttk.Button(
            dd_log_bar,
            text="清空日志",
            style="Compact.TButton",
            command=lambda: self._clear_log_widget(self.dd_log_text),
        ).pack(side=tk.RIGHT)
        ttk.Label(dd_log_bar, text="查重匹配详情与导出结果：", foreground="#475569").pack(side=tk.LEFT)

        log_wrap = ttk.Frame(act_box)
        log_wrap.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.dd_log_text = self._create_styled_log_text(log_wrap, height=4)
        self.dd_log_text.insert(
            tk.END,
            "提示：本板块完全离线运行（无需 API Key），自动输出《查重结果明细报告_YYYYMMDD.docx》及去重重算字数后的定稿。\n",
            "dim",
        )

    # ---------------- Tab 4: 英文字数统计 ----------------

    def _build_tab4_word_counter(self, parent: ttk.Frame):
        # Pack bottom action bar FIRST so primary CTA NEVER clips at any height
        bot = ttk.Frame(parent)
        bot.pack(side=tk.BOTTOM, fill=tk.X)

        _make_primary_btn(
            bot,
            text="计算字数并立即更新文件名",
            command=self._wc_execute_rename,
        ).pack(side=tk.RIGHT)
        ttk.Button(
            bot,
            text="重新计算并刷新预览",
            style="Action.TButton",
            command=self._wc_refresh_previews,
        ).pack(side=tk.RIGHT, padx=(0, 8))

        ttk.Label(
            bot,
            textvariable=self.wc_summary_var,
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
            foreground="#1d4ed8",
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Frame(parent, bg="#e2e8f0", height=1).pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 8))

        _pw, upper, lower = _make_vertical_paned(
            parent, upper_weight=0, lower_weight=1, min_top=220, min_bottom=120
        )

        # 1. Rules banner
        banner = ttk.LabelFrame(upper, text=" 统计与文件重命名规则说明（本地毫秒级解析 · 0 词自动省略 · 可选同步顶部说明行） ", padding=8)
        banner.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))

        ttk.Label(
            banner,
            text=(
                "• 统计口径：仅统计表格【第 2 列（英文列）】修订后正文的英文单词数（标点不计，已删除修订自动排除；UP = 黄色高亮，RV = 青色高亮）。\n"
                "• 0 词省略与命名：RV=0 或 UP=0 时自动省略 0 词项；支持保留原名、取首行加粗英文标题、或零散标准命名 God's words [in 关键词]。"
            ),
            foreground="#334155",
        ).pack(fill=tk.X)

        # 2. File picker & naming parameters
        cfg_frame = ttk.LabelFrame(upper, text=" 1. 拖拽 / 选择文档与命名参数 ", padding=8)
        cfg_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 4))

        drop_zone = self._make_interactive_drop_zone(
            cfg_frame,
            "直接把一个或多个 .docx 文档、文件夹拖拽到此窗口任意位置即可自动载入（点击此处也可选择文件）",
            self._wc_add_files,
        )
        drop_zone.pack(fill=tk.X, pady=(0, 6))

        file_btns = ttk.Frame(cfg_frame)
        file_btns.pack(fill=tk.X, pady=(0, 6))

        ttk.Button(
            file_btns,
            text="选择 Word 文档...",
            style="Action.TButton",
            command=self._wc_add_files,
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(
            file_btns,
            text="选择文件夹批量载入...",
            style="Action.TButton",
            command=self._wc_add_folder,
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(file_btns, text="移除选中行", command=self._wc_remove_selected).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(file_btns, text="清空列表", command=self._wc_clear_all).pack(side=tk.LEFT)

        row_lang = ttk.Frame(cfg_frame)
        row_lang.pack(fill=tk.X, pady=3)

        ttk.Label(row_lang, text="语言代号：").pack(side=tk.LEFT)
        self.wc_lang_cb = ttk.Combobox(
            row_lang,
            textvariable=self.wc_lang_var,
            values=self._get_combo_values(),
            width=28,
        )
        self.wc_lang_cb.pack(side=tk.LEFT, padx=(4, 6))
        self.wc_lang_cb.bind("<<ComboboxSelected>>", self._on_wc_lang_selected)

        ttk.Button(row_lang, text="管理语种...", style="Compact.TButton", command=self._open_lang_manager).pack(
            side=tk.LEFT, padx=(0, 14)
        )

        ttk.Checkbutton(
            row_lang,
            text="原文件名已有语言代号（如 GR_）时优先保留",
            variable=self.wc_keep_existing_lang_var,
            command=self._wc_refresh_previews,
        ).pack(side=tk.LEFT)

        # Naming Mode Radio Group Row
        row_naming = ttk.Frame(cfg_frame)
        row_naming.pack(fill=tk.X, pady=3)
        ttk.Label(row_naming, text="文档名规则：").pack(side=tk.LEFT)
        ttk.Radiobutton(
            row_naming,
            text="保留原文档名",
            value="keep",
            variable=self.wc_naming_mode_var,
            command=self._wc_refresh_previews,
        ).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Radiobutton(
            row_naming,
            text="全篇改点：取首行第2列加粗英文标题",
            value="full",
            variable=self.wc_naming_mode_var,
            command=self._wc_refresh_previews,
        ).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Radiobutton(
            row_naming,
            text="零散改点：God's words [in 关键词]",
            value="scattered",
            variable=self.wc_naming_mode_var,
            command=self._wc_refresh_previews,
        ).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Label(row_naming, text="关键词：", foreground="#475569").pack(side=tk.LEFT)
        wc_kw_ent = ttk.Entry(row_naming, textvariable=self.wc_keyword_var, width=14)
        wc_kw_ent.pack(side=tk.LEFT, padx=(2, 0))
        wc_kw_ent.bind("<KeyRelease>", self._on_wc_keyword_typed)

        row_date = ttk.Frame(cfg_frame)
        row_date.pack(fill=tk.X, pady=3)

        ttk.Label(row_date, text="日期 (YYYYMMDD)：").pack(side=tk.LEFT)
        date_ent = ttk.Entry(row_date, textvariable=self.wc_date_var, width=10)
        date_ent.pack(side=tk.LEFT, padx=(4, 12))
        date_ent.bind("<KeyRelease>", lambda _e: self._wc_refresh_previews())

        ttk.Checkbutton(
            row_date,
            text="保留原 8 位日期",
            variable=self.wc_keep_existing_date_var,
            command=self._wc_refresh_previews,
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Checkbutton(
            row_date,
            text="去除“初筛_”前缀",
            variable=self.wc_strip_draft_var,
            command=self._wc_refresh_previews,
        ).pack(side=tk.LEFT, padx=(0, 14))

        ttk.Checkbutton(
            row_date,
            text="同步更新开头高亮说明行（Blue / Yellow highlighted parts...）",
            variable=self.wc_update_headers_var,
            command=self._wc_refresh_previews,
        ).pack(side=tk.LEFT)

        row_mode = ttk.Frame(cfg_frame)
        row_mode.pack(fill=tk.X, pady=(3, 0))

        ttk.Button(row_mode, text="浏览...", style="Compact.TButton", command=self._wc_browse_copy_dir).pack(
            side=tk.RIGHT
        )
        ttk.Label(row_mode, text="保存方式：").pack(side=tk.LEFT)
        ttk.Radiobutton(
            row_mode,
            text="直接就地更新原文档（推荐）",
            value="inplace",
            variable=self.wc_save_mode_var,
        ).pack(side=tk.LEFT, padx=(4, 14))
        self.wc_copy_radio = ttk.Radiobutton(
            row_mode,
            text="另存副本到指定文件夹：",
            value="copy",
            variable=self.wc_save_mode_var,
        )
        self.wc_copy_radio.pack(side=tk.LEFT)
        self.wc_copy_dir_entry = ttk.Entry(row_mode, textvariable=self.wc_copy_dir_var, width=28)
        self.wc_copy_dir_entry.pack(side=tk.LEFT, padx=4, fill=tk.X, expand=True)

        # 3. Preview & Results Treeview (Lower Resizable Pane)
        tbl_box = ttk.LabelFrame(
            lower,
            text=" 2. 第 2 列英文字数统计与新文件名预览（支持拖入文件/文件夹；双击行可打开目录） ",
            padding=8,
        )
        tbl_box.pack(fill=tk.BOTH, expand=True, pady=(2, 0))

        cols = ("old_name", "lang", "up_words", "rv_words", "new_name", "status")
        self.wc_tree = ttk.Treeview(tbl_box, columns=cols, show="headings", selectmode="extended", height=7)
        self.wc_tree.heading("old_name", text="原文件名")
        self.wc_tree.heading("lang", text="语言代号")
        self.wc_tree.heading("up_words", text="UP 单词数 (黄色)")
        self.wc_tree.heading("rv_words", text="RV 单词数 (青色)")
        self.wc_tree.heading("new_name", text="计算后目标文件名")
        self.wc_tree.heading("status", text="状态")

        self.wc_tree.column("old_name", width=215, anchor=tk.W)
        self.wc_tree.column("lang", width=70, anchor=tk.CENTER)
        self.wc_tree.column("up_words", width=105, anchor=tk.CENTER)
        self.wc_tree.column("rv_words", width=105, anchor=tk.CENTER)
        self.wc_tree.column("new_name", width=320, anchor=tk.W)
        self.wc_tree.column("status", width=120, anchor=tk.CENTER)

        hsb = ttk.Scrollbar(tbl_box, orient=tk.HORIZONTAL, command=self.wc_tree.xview)
        vsb = ttk.Scrollbar(tbl_box, orient=tk.VERTICAL, command=self.wc_tree.yview)
        self.wc_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.wc_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.wc_tree.bind("<Double-1>", self._wc_open_selected_folder)

    # ---------------- Tab 5: 分part稿整理/合并 ----------------

    def _build_tab5_part_merge(self, parent: ttk.Frame):
        _pw, upper, lower = _make_vertical_paned(
            parent, upper_weight=1, lower_weight=1, min_top=240, min_bottom=145
        )

        cfg_frame = ttk.LabelFrame(
            upper,
            text=" 1. 分part稿整理/合并设置（支持多文件 Part 1..N 合并 & 单文件多 Part 表格合并 · 蓝/红/加粗精准同步 · 语系标准字体） ",
            padding=8,
        )
        cfg_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))

        drop_zone = self._make_interactive_drop_zone(
            cfg_frame,
            "直接拖入分篇文件夹、多个 Part .docx，或含多段 Part 表格的单个 .docx（自动套用语系字体，如泰语 Angsana New 18/16/14pt）",
            self._fel_add_folder,
        )
        drop_zone.pack(fill=tk.X, pady=(0, 6))

        btn_bar = ttk.Frame(cfg_frame)
        btn_bar.pack(fill=tk.X, pady=(0, 6))

        ttk.Button(
            btn_bar,
            text="选择文件夹批量载入...",
            style="Action.TButton",
            command=self._fel_add_folder,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            btn_bar,
            text="选择 Word 文档...",
            style="Action.TButton",
            command=self._fel_add_files,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_bar, text="移除选中", command=self._fel_remove_selected).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_bar, text="清空列表", command=self._fel_clear_all).pack(side=tk.LEFT)

        out_row = ttk.Frame(cfg_frame)
        out_row.pack(fill=tk.X, pady=3)
        ttk.Button(out_row, text="打开目录", style="Compact.TButton", command=self._fel_open_output_dir).pack(
            side=tk.RIGHT, padx=(2, 0)
        )
        ttk.Button(out_row, text="浏览...", style="Compact.TButton", command=self._fel_browse_output).pack(
            side=tk.RIGHT, padx=2
        )
        self.lbl_fel_output_dir = ttk.Label(out_row, text="输出保存目录：")
        self.lbl_fel_output_dir.pack(side=tk.LEFT)
        self.ent_fel_output_dir = ttk.Entry(out_row, textvariable=self.fel_output_dir_var)
        self.ent_fel_output_dir.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

        opt_row = ttk.Frame(cfg_frame)
        opt_row.pack(fill=tk.X, pady=3)

        ttk.Label(opt_row, text="目标语种：").pack(side=tk.LEFT)
        self.fel_lang_cb = ttk.Combobox(
            opt_row,
            textvariable=self.fel_lang_var,
            values=self._get_combo_values(),
            width=20,
        )
        self.fel_lang_cb.pack(side=tk.LEFT, padx=(4, 4))
        self.fel_lang_cb.bind("<<ComboboxSelected>>", self._fel_on_lang_changed)

        ttk.Button(opt_row, text="管理语种/字体...", style="Compact.TButton", command=self._open_lang_manager).pack(
            side=tk.LEFT, padx=(0, 10)
        )

        ttk.Label(opt_row, text="三语前缀：").pack(side=tk.LEFT)
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
            width=24,
        ).pack(side=tk.LEFT, padx=(3, 0), fill=tk.X, expand=True)

        mode_row = ttk.Frame(cfg_frame)
        mode_row.pack(fill=tk.X, pady=(4, 2))
        ttk.Label(mode_row, text="输出方式：", font=(UI_FONT_FAMILY, _ui_pt(9), "bold")).pack(side=tk.LEFT)
        for disp_txt, mode_val in OUTPUT_MODES:
            ttk.Radiobutton(
                mode_row,
                text=disp_txt,
                value=mode_val,
                variable=self.fel_output_mode_var,
                command=self._fel_refresh_previews,
            ).pack(side=tk.LEFT, padx=(6, 10))

        ai_opt_row = ttk.Frame(cfg_frame)
        ai_opt_row.pack(fill=tk.X, pady=(2, 1))
        ttk.Checkbutton(
            ai_opt_row,
            text="使用 AI 自动同步第 3 列段内局部引文蓝字/加粗（严格镜像第 2 列，不误加粗普通叙述；无 Key 则本地秒出）",
            variable=self.fel_enable_ai_var,
            command=self._save_settings,
        ).pack(side=tk.LEFT)

        red_opt_row = ttk.Frame(cfg_frame)
        red_opt_row.pack(fill=tk.X, pady=(1, 0))
        ttk.Checkbutton(
            red_opt_row,
            text="同步标记红色字体到第 3 列（适用于神话定稿；若外文已有红色则自动跳过）",
            variable=self.fel_sync_red_var,
            command=self._save_settings,
        ).pack(side=tk.LEFT)

        # 2. Sorted Part List & Output Filename Preview Treeview
        tbl_box = ttk.LabelFrame(
            upper,
            text=" 2. 分篇顺序与输出定稿命名预览（自动合并格线防错位 · 横向 A4 · 清理句尾空格与格底空行 · 标准语系字号） ",
            padding=6,
        )
        tbl_box.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 4))

        cols = ("seq", "old_name", "part_tag", "rows", "preview_name")
        self.fel_tree = ttk.Treeview(tbl_box, columns=cols, show="headings", selectmode="extended", height=5)
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

        # 3. Execute & Progress Frame (Lower Resizable Pane)
        act_box = ttk.LabelFrame(lower, text=" 3. 一键执行与实时日志（可拖动上方分隔条调节高度） ", padding=8)
        act_box.pack(fill=tk.BOTH, expand=True, pady=(2, 0))

        run_row = ttk.Frame(act_box)
        run_row.pack(side=tk.TOP, fill=tk.X)

        # Pack primary button FIRST so long preview filenames never squeeze the button
        self.btn_fel_run = _make_primary_btn(
            run_row,
            text="开始分part稿整理/合并",
            command=self._on_run_fellowship,
        )
        self.btn_fel_run.pack(side=tk.RIGHT, padx=(10, 0))

        ttk.Label(
            run_row,
            textvariable=self.fel_summary_var,
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
            foreground="#0f172a",
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        p_info = ttk.Frame(act_box)
        p_info.pack(side=tk.TOP, fill=tk.X, pady=(6, 2))
        ttk.Label(
            p_info,
            textvariable=self.fel_pct_timer_var,
            font=("Consolas", 9, "bold"),
            foreground="#0f172a",
        ).pack(side=tk.RIGHT, padx=(8, 0))
        self.fel_stage_lbl = ttk.Label(
            p_info,
            textvariable=self.fel_stage_var,
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
            foreground="#1d4ed8",
        )
        self.fel_stage_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Progressbar(
            act_box,
            mode="determinate",
            maximum=100.0,
            variable=self.fel_progress_var,
        ).pack(side=tk.TOP, fill=tk.X, pady=(0, 5))

        fel_log_bar = ttk.Frame(act_box)
        fel_log_bar.pack(side=tk.TOP, fill=tk.X, pady=(0, 3))
        ttk.Button(
            fel_log_bar,
            text="清空日志",
            style="Compact.TButton",
            command=lambda: self._clear_log_widget(self.fel_log_text),
        ).pack(side=tk.RIGHT)
        ttk.Label(fel_log_bar, text="合并排版、样式同步与导出日志：", foreground="#475569").pack(side=tk.LEFT)

        log_wrap = ttk.Frame(act_box)
        log_wrap.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.fel_log_text = self._create_styled_log_text(log_wrap, height=4)
        self.fel_log_text.insert(
            tk.END,
            "提示：支持多份分篇文档（Part 1..N）按序整理/合并，也支持单个文档内包含多段 Part 标题与表格自动合并统一列宽。\n",
            "dim",
        )

    # ---------------- Tab 6: 系统设置 ----------------

    def _build_tab6_settings(self, parent: ttk.Frame):
        # Pack Area 3 & Area 2 at BOTTOM first so fixed sections NEVER clip, while Area 1's key editor expands smoothly
        lang_box = ttk.LabelFrame(
            parent,
            text=" 3. 常用语种置顶与标准字体库（常用与最近使用语种自动排在所有板块下拉框顶部） ",
            padding=10,
        )
        lang_box.pack(side=tk.BOTTOM, fill=tk.X)

        lang_row = ttk.Frame(lang_box)
        lang_row.pack(fill=tk.X)
        ttk.Button(
            lang_row,
            text="管理常用语种与标准字体库...",
            style="Action.TButton",
            command=self._open_lang_manager,
        ).pack(side=tk.RIGHT, padx=(12, 0))
        self.lbl_fav_summary = ttk.Label(
            lang_row,
            textvariable=self.fav_summary_var,
            foreground="#334155",
            justify=tk.LEFT,
            wraplength=640,
        )
        self.lbl_fav_summary.pack(side=tk.LEFT, fill=tk.X, expand=True)
        lang_row.bind(
            "<Configure>",
            lambda e: self.lbl_fav_summary.configure(wraplength=max(260, int(e.width) - 230)),
        )

        rule_box = ttk.LabelFrame(
            parent,
            text=" 2. 改点高亮与排版规则参数 ",
            padding=10,
        )
        rule_box.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 8))

        gap_row = ttk.Frame(rule_box)
        gap_row.pack(fill=tk.X)
        ttk.Label(
            gap_row,
            text="两处黄色高亮自动连片最大间隔词数：",
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
            foreground="#0f172a",
        ).pack(side=tk.LEFT)
        self.spin_yellow_gap = ttk.Spinbox(
            gap_row,
            from_=0,
            to=20,
            textvariable=self.yellow_gap_words_var,
            width=8,
            command=self._on_yellow_gap_changed,
        )
        self.spin_yellow_gap.pack(side=tk.LEFT, padx=(6, 12))
        self.spin_yellow_gap.bind("<KeyRelease>", lambda _e: self._refresh_top_status_summary())
        self.spin_yellow_gap.bind("<FocusOut>", self._on_yellow_gap_changed)
        self.spin_yellow_gap.bind("<Return>", self._on_yellow_gap_changed)

        ttk.Label(
            rule_box,
            text="同一段内两处黄色高亮之间未标色英文词数 ≤ 此数值时自动连起来标记高亮（默认 3，填 0 则不自动连片）",
            foreground="#475569",
        ).pack(fill=tk.X, pady=(5, 0))

        # Area 1: API Key Pool, Channel & Model Settings
        api_box = ttk.LabelFrame(
            parent,
            text=" 1. API 密钥池、通道与模型设置（跨所有工作台共享生效） ",
            padding=10,
        )
        api_box.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 8))

        combo_grid = ttk.Frame(api_box)
        combo_grid.pack(side=tk.TOP, fill=tk.X, pady=(0, 8))
        combo_grid.columnconfigure(1, weight=1)
        combo_grid.columnconfigure(3, weight=1)

        ttk.Label(combo_grid, text="选择模型：").grid(row=0, column=0, sticky=tk.W, pady=3)
        model_row_box = ttk.Frame(combo_grid)
        model_row_box.grid(row=0, column=1, sticky=tk.EW, pady=3, padx=(4, 16))
        model_row_box.columnconfigure(0, weight=1)
        self.settings_model_cb = ttk.Combobox(
            model_row_box,
            textvariable=self.model_label,
            values=[m[0] for m in self.dynamic_models],
            state="normal",
            width=26,
        )
        self.settings_model_cb.grid(row=0, column=0, sticky=tk.EW)
        self.settings_model_cb.bind("<<ComboboxSelected>>", lambda _e: self._save_settings())
        self.settings_model_cb.bind("<FocusOut>", lambda _e: self._save_settings())
        self.settings_model_cb.bind("<Return>", lambda _e: self._save_settings())
        self.btn_refresh_models = ttk.Button(
            model_row_box,
            text="🔄 加载最新模型",
            style="Compact.TButton",
            command=self._on_refresh_models_clicked,
        )
        self.btn_refresh_models.grid(row=0, column=1, padx=(4, 0))

        ttk.Label(combo_grid, text="接口通道：").grid(row=0, column=2, sticky=tk.W, pady=3)
        self.settings_api_mode_cb = ttk.Combobox(
            combo_grid,
            textvariable=self.api_mode_label,
            values=[m[0] for m in API_MODES],
            state="readonly",
            width=26,
        )
        self.settings_api_mode_cb.grid(row=0, column=3, sticky=tk.EW, pady=3, padx=(4, 0))
        self.settings_api_mode_cb.bind("<<ComboboxSelected>>", lambda _e: self._save_settings())

        ttk.Label(combo_grid, text="多 Key 调度策略：").grid(row=1, column=0, sticky=tk.W, pady=3)
        self.settings_rotation_cb = ttk.Combobox(
            combo_grid,
            textvariable=self.rotation_label,
            values=[m[0] for m in ROTATION_MODES],
            state="readonly",
            width=28,
        )
        self.settings_rotation_cb.grid(row=1, column=1, sticky=tk.EW, pady=3, padx=(4, 16))
        self.settings_rotation_cb.bind("<<ComboboxSelected>>", lambda _e: self._save_settings())

        ttk.Checkbutton(
            combo_grid,
            text="本机加密记住密钥池 (Windows DPAPI)",
            variable=self.remember_key_var,
            command=self._save_settings,
        ).grid(row=1, column=2, columnspan=2, sticky=tk.W, pady=3)

        # Pack bottom buttons & status inside Area 1 BEFORE flexible Text editor so buttons NEVER clip
        api_btns = ttk.Frame(api_box)
        api_btns.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))

        self.btn_settings_save = _make_primary_btn(
            api_btns,
            text="保存设置",
            command=self._on_save_settings_clicked,
        )
        self.btn_settings_save.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_settings_test = ttk.Button(
            api_btns,
            text="测试连接",
            style="Action.TButton",
            command=self._on_test_api,
        )
        self.btn_settings_test.pack(side=tk.LEFT, padx=(0, 8))
        self.test_btn = self.btn_settings_test

        self.btn_settings_test_all = ttk.Button(
            api_btns,
            text="批量检测全部 Key",
            style="Action.TButton",
            command=self._open_key_pool_dialog,
        )
        self.btn_settings_test_all.pack(side=tk.LEFT, padx=(0, 8))
        self.pool_btn = self.btn_settings_test_all

        self.btn_settings_clear_keys = ttk.Button(
            api_btns,
            text="清空输入",
            command=self._on_clear_settings_keys,
        )
        self.btn_settings_clear_keys.pack(side=tk.LEFT)

        self.settings_status_lbl = ttk.Label(
            api_box,
            textvariable=self.settings_save_status_var,
            foreground="#475569",
        )
        self.settings_status_lbl.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 0))

        key_hdr = ttk.Frame(api_box)
        key_hdr.pack(side=tk.TOP, fill=tk.X, pady=(0, 4))
        ttk.Label(
            key_hdr,
            textvariable=self.settings_key_count_var,
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
            foreground="#1d4ed8",
        ).pack(side=tk.RIGHT)
        ttk.Label(
            key_hdr,
            text="API Key 池列表（每行 1 个 Key，支持 # 备注，与弹窗实时同步）：",
            font=(UI_FONT_FAMILY, _ui_pt(9), "bold"),
            foreground="#0f172a",
        ).pack(side=tk.LEFT)

        ed_outer = ttk.Frame(api_box)
        ed_outer.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        ed_hsb = ttk.Scrollbar(
            ed_outer,
            orient=tk.HORIZONTAL,
            command=lambda *a: self.settings_key_editor.xview(*a) if self.settings_key_editor else None,
        )
        ed_hsb.pack(side=tk.BOTTOM, fill=tk.X)

        ed_inner = ttk.Frame(ed_outer)
        ed_inner.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.settings_key_editor = tk.Text(
            ed_inner,
            height=5,
            wrap=tk.NONE,
            font=(MONO_FONT_FAMILY, _ui_pt(9)),
            bg="#ffffff",
            fg="#0f172a",
            relief=tk.SOLID,
            borderwidth=1,
            padx=8,
            pady=6,
            undo=True,
        )
        ed_vsb = ttk.Scrollbar(ed_inner, orient=tk.VERTICAL, command=self.settings_key_editor.yview)
        self.settings_key_editor.configure(yscrollcommand=ed_vsb.set, xscrollcommand=ed_hsb.set)
        self.settings_key_editor.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ed_vsb.pack(side=tk.RIGHT, fill=tk.Y)

        if self._pool_multiline_text.strip():
            self.settings_key_editor.insert("1.0", self._pool_multiline_text.strip() + "\n")

        self.settings_key_editor.bind("<KeyRelease>", self._on_settings_key_editor_modified)
        self.settings_key_editor.bind("<<Paste>>", lambda _e: self.root.after(20, self._on_settings_key_editor_modified))

    @staticmethod
    def _create_styled_log_text(parent: tk.Misc, height: int = 8) -> tk.Text:
        tw = tk.Text(
            parent,
            height=height,
            wrap=tk.WORD,
            font=(UI_FONT_FAMILY, _ui_pt(9)),
            bg="#f8fafc",
            fg="#0f172a",
            relief=tk.SOLID,
            borderwidth=1,
            padx=8,
            pady=5,
        )
        tw.tag_configure("header", foreground="#1d4ed8", font=(UI_FONT_FAMILY, _ui_pt(9), "bold"))
        tw.tag_configure("ok", foreground="#15803d", font=(UI_FONT_FAMILY, _ui_pt(9), "bold"))
        tw.tag_configure("err", foreground="#b91c1c", font=(UI_FONT_FAMILY, _ui_pt(9), "bold"))
        tw.tag_configure("retry", foreground="#b45309")
        tw.tag_configure("dim", foreground="#475569")
        sb = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=tw.yview)
        tw.configure(yscrollcommand=sb.set)
        tw.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        return tw

    @staticmethod
    def _clear_log_widget(tw: tk.Text):
        tw.delete("1.0", tk.END)

    # ---------------- Drag & Drop & Library Indexing ----------------

    def _warm_library_index_async(self, lib_dir: str, force_rebuild: bool = False, target_tab: str = "main"):
        raw = (lib_dir or "").strip()
        if not raw or ensure_library_index is None:
            return
        try:
            rp = Path(raw).resolve()
        except Exception:
            return
        if not rp.is_dir():
            if force_rebuild:
                messagebox.showwarning("提示", "请先选择真实存在的语种总库文件夹目录。")
            return

        norm_key = os.path.normcase(str(rp))
        if target_tab == "fa":
            self._last_typed_fa_lib_norm = norm_key
        else:
            self._last_typed_lib_norm = norm_key

        with self._indexing_lock:
            if norm_key in self._indexing_inflight:
                if force_rebuild:
                    self._log("⏳ 该总库目录正在后台扫描建库中，请稍候...", target_tab=target_tab)
                return
            self._indexing_inflight.add(norm_key)

        if force_rebuild:
            self._log(f"🔄 正在强制重新扫描并刷新语种总库索引：{rp} ...", target_tab=target_tab)

        def _bg_worker():
            t0 = time.perf_counter()
            try:
                info = _call_compatible(ensure_library_index, str(rp), force_rebuild=force_rebuild)
                if isinstance(info, dict):
                    total_files = info.get("total_files", 0)
                    total_rows = info.get("total_rows", 0)
                    rebuilt = info.get("rebuilt_files", 0)
                    elapsed_sec = float(info.get("elapsed_sec", time.perf_counter() - t0))
                    if force_rebuild:
                        self._log(
                            f"✅ 总库索引已强制刷新至最新状态：共 {total_files} 篇文档（{total_rows} 段，重解析 {rebuilt} 篇，耗时 {elapsed_sec:.1f}s）",
                            target_tab=target_tab,
                        )
                    else:
                        self._log(
                            f"📚 总库后台索引已就绪：共 {total_files} 篇文档（{total_rows} 段），支持无标题极速检索（耗时 {elapsed_sec:.1f}s）",
                            target_tab=target_tab,
                        )
            except Exception as ex:
                clean_msg = _format_exception_concise(ex)
                self._log(f"⚠️ 总库索引刷新跳过：{clean_msg}", target_tab=target_tab)
            finally:
                with self._indexing_lock:
                    self._indexing_inflight.discard(norm_key)

        threading.Thread(target=_bg_worker, daemon=True).start()

    def _refresh_library_index_manual(self, lib_dir: str, target_tab: str = "main"):
        raw = (lib_dir or "").strip()
        if not raw:
            messagebox.showinfo("提示", "请先选择【语种总库目录】，再点击「刷新总库」。")
            return
        self._save_settings()
        self._warm_library_index_async(raw, force_rebuild=True, target_tab=target_tab)

    def _on_library_entry_committed(self, _event=None):
        raw = self.library_dir_var.get().strip()
        if not raw:
            return
        try:
            rp = Path(raw).resolve()
        except Exception:
            return
        if not rp.is_dir():
            return
        norm_key = os.path.normcase(str(rp))
        if norm_key != self._last_typed_lib_norm:
            self._save_settings()
            self._warm_library_index_async(str(rp), target_tab="main")

    def _on_fa_library_entry_committed(self, _event=None):
        raw = self.fa_library_dir_var.get().strip()
        if not raw:
            return
        try:
            rp = Path(raw).resolve()
        except Exception:
            return
        if not rp.is_dir():
            return
        norm_key = os.path.normcase(str(rp))
        if norm_key != self._last_typed_fa_lib_norm:
            self._save_settings()
            self._warm_library_index_async(str(rp), target_tab="fa")

    def _setup_drag_and_drop(self):
        # Left sidebar navigation buttons
        self.dnd.register_target(self.nav_btn_main, lambda ps: (self._switch_tab("main"), self._handle_tab1_drop(ps)))
        self.dnd.register_target(self.nav_btn_fa, lambda ps: (self._switch_tab("fa"), self._handle_tab2_drop(ps)))
        self.dnd.register_target(self.nav_btn_dd, lambda ps: (self._switch_tab("dd"), self._handle_tab3_drop(ps)))
        if getattr(self, "ingest_panel", None) is not None:
            self.dnd.register_target(
                self.nav_btn_ingest,
                lambda ps: (self._switch_tab("ingest"), self.ingest_panel.handle_fallback_drop(ps)),
            )
            self.ingest_panel.register_dnd_targets(self.dnd)
        self.dnd.register_target(self.nav_btn_wc, lambda ps: (self._switch_tab("wc"), self._wc_ingest_paths(ps)))
        self.dnd.register_target(self.nav_btn_fel, lambda ps: (self._switch_tab("fel"), self._fel_ingest_paths(ps)))

        # Tab 1 specific input rows
        for w in (self.lbl_input_docx, self.ent_input_docx):
            self.dnd.register_target(w, self._on_drop_input_docx)
        for w in (self.lbl_library_dir, self.ent_library_dir):
            self.dnd.register_target(w, self._on_drop_library_dir)
        for w in (self.lbl_output_dir, self.ent_output_dir):
            self.dnd.register_target(w, self._on_drop_output_dir)
        for w in (self.chk_sc1_master, self.ent_sc1_master_path):
            self.dnd.register_target(w, self._on_drop_sc1_master_path)
        for w in (self.chk_sc1_dedup, self.ent_sc1_dedup_dir):
            self.dnd.register_target(w, self._on_drop_sc1_dedup_dir)

        # Tab 2 specific input rows
        self.dnd.register_target(self.fa_titles_text, self._on_drop_fa_titles_file)
        for w in (self.lbl_fa_master_dir, self.ent_fa_master_dir):
            self.dnd.register_target(w, self._on_drop_fa_master_dir)
        for w in (self.lbl_fa_library_dir, self.ent_fa_library_dir):
            self.dnd.register_target(w, self._on_drop_fa_library_dir)
        for w in (self.lbl_fa_output_dir, self.ent_fa_output_dir):
            self.dnd.register_target(w, self._on_drop_fa_output_dir)
        for w in (self.chk_fa_dedup, self.ent_fa_dedup_dir):
            self.dnd.register_target(w, self._on_drop_fa_dedup_dir)

        # Tab 3 specific input rows
        for w in (self.lbl_dd_pkg_dir, self.ent_dd_pkg_dir):
            self.dnd.register_target(w, self._on_drop_dd_pkg_dir)
        for w in (self.lbl_dd_out_dir, self.ent_dd_out_dir):
            self.dnd.register_target(w, self._on_drop_dd_out_dir)

        # Tab 4 specific copy-folder target
        for w in (self.wc_copy_radio, self.wc_copy_dir_entry):
            self.dnd.register_target(w, self._on_drop_wc_copy_dir)

        # Tab 5 specific output-folder target
        for w in (self.lbl_fel_output_dir, self.ent_fel_output_dir):
            self.dnd.register_target(w, self._on_drop_fel_output_dir)

        self.dnd.install()

    def _on_drop_fallback(self, paths: list[Path]):
        if self.active_tab_key == "fa":
            self._handle_tab2_drop(paths)
        elif self.active_tab_key == "dd":
            self._handle_tab3_drop(paths)
        elif self.active_tab_key == "ingest":
            if getattr(self, "ingest_panel", None) is not None:
                self.ingest_panel.handle_fallback_drop(paths)
        elif self.active_tab_key == "wc":
            self._wc_ingest_paths(paths)
        elif self.active_tab_key == "fel":
            self._fel_ingest_paths(paths)
        elif self.active_tab_key == "settings":
            return
        else:
            self._handle_tab1_drop(paths)

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

    # ---- Tab 1 Drop & Helpers ----

    def _on_tab1_lang_selected(self, _event=None):
        code = self._get_selected_lang_code()
        if code and code != "AUTO":
            self._record_recent_lang_and_refresh(code, save=True)
        else:
            self._save_settings()

    def _refresh_input_docx_status(self):
        raw = self.input_docx_var.get().strip()
        if not raw:
            messagebox.showinfo("提示", "请先选择或拖入【待整理文档 / 文件夹】。")
            return
        files = resolve_sc1_input_docx_list(raw) if callable(resolve_sc1_input_docx_list) else []
        if not files:
            msg = f"未在指定路径中找到有效的 .docx 文档：{raw}"
            self._append_log(f"⚠️ {msg}", target_tab="main")
            messagebox.showwarning("未找到文档", msg)
            return
        if len(files) == 1:
            p = files[0]
            st = p.stat()
            mtime_str = datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            size_kb = max(1, int(round(st.st_size / 1024.0)))
            log_msg = (
                f"🔄 已刷新待整理文档状态：{p.name}（最后修改：{mtime_str}，大小：{size_kb} KB；每次点击运行均自动读取磁盘最新内容）"
            )
            self._append_log(log_msg, target_tab="main")
            self.stage_var.set(f"已刷新待整理文档状态：{p.name}（最后修改 {mtime_str}）")
        else:
            latest_m = max(f.stat().st_mtime for f in files)
            mtime_str = datetime.fromtimestamp(latest_m).strftime("%Y-%m-%d %H:%M:%S")
            preview_names = "、".join(f.name for f in files[:3]) + (f" 等 {len(files)} 份" if len(files) > 3 else "")
            log_msg = f"🔄 已刷新待整理批量列表：共检测到 {len(files)} 份 .docx 文档（{preview_names} | 最近修改：{mtime_str}）"
            self._append_log(log_msg, target_tab="main")
            self.stage_var.set(f"已载入 {len(files)} 份待整理文档（最近修改 {mtime_str}）")

    def _browse_sc1_master_file(self):
        p = filedialog.askopenfilename(
            title="选择最新中英底稿 Word 文档 (.docx)",
            filetypes=[("Word 文档", "*.docx")],
        )
        if p:
            self.sc1_master_path_var.set(p)
            self.sc1_sync_master_var.set(True)
            self._save_settings()
            self._append_log(f"📥 已选择【最新中英底稿文件】：{p}", target_tab="main")

    def _browse_sc1_master_dir(self):
        d = filedialog.askdirectory(title="选择包含最新中英底稿的文件夹")
        if d:
            self.sc1_master_path_var.set(d)
            self.sc1_sync_master_var.set(True)
            self._save_settings()
            self._append_log(f"📥 已选择【最新中英底稿文件夹】：{d}", target_tab="main")

    def _on_drop_sc1_master_path(self, paths: list[Path]):
        for raw_p in paths:
            try:
                rp = raw_p.resolve()
            except Exception:
                continue
            if rp.is_dir() or (rp.is_file() and rp.suffix.lower() == ".docx" and not rp.name.startswith("~$")):
                self.sc1_master_path_var.set(str(rp))
                self.sc1_sync_master_var.set(True)
                self._save_settings()
                kind = "文件夹" if rp.is_dir() else "文件"
                self._append_log(f"📥 已拖入【最新中英底稿{kind}】：{rp}", target_tab="main")
                self.stage_var.set(f"📥 已设定最新中英底稿：{rp.name}")
                return
        messagebox.showinfo("拖拽提示", "【核对最新中英底稿】请拖入最新中英底稿 .docx 文件或所在文件夹。")

    def _on_sc1_keyword_typed(self, _event=None):
        if self.sc1_keyword_var.get().strip():
            self.sc1_naming_mode_var.set("scattered")
        self._save_settings()

    def _on_sc1_custom_name_typed(self, _event=None):
        if self.doc_name_var.get().strip():
            self.sc1_naming_mode_var.set("keep")
        self._save_settings()

    def _browse_sc1_dedup_dir(self):
        d = filedialog.askdirectory(title="选择「未入库改点包」文件夹")
        if d:
            self.sc1_dedup_dir_var.set(d)
            self.sc1_enable_dedup_var.set(True)
            self._save_settings()

    def _on_drop_sc1_dedup_dir(self, paths: list[Path]):
        p = paths[0].resolve()
        target_dir = p if p.is_dir() else p.parent
        if target_dir.is_dir():
            self.sc1_dedup_dir_var.set(str(target_dir))
            self.sc1_enable_dedup_var.set(True)
            self._save_settings()
            self._append_log(f"📥 已设定【未入库改点包目录】：{target_dir}", target_tab="main")

    def _on_drop_input_docx(self, paths: list[Path]):
        if len(paths) == 1 and paths[0].resolve().is_dir():
            d_resolved = paths[0].resolve()
            docxs = resolve_sc1_input_docx_list(d_resolved) if callable(resolve_sc1_input_docx_list) else self._extract_docx_from_paths([d_resolved])
            if not docxs:
                messagebox.showinfo("拖拽提示", f"文件夹中未找到待整理的 .docx 文档：\n{d_resolved}")
                return
            self.input_docx_var.set(str(d_resolved))
            if not self.output_dir_var.get().strip():
                self.output_dir_var.set(str(d_resolved.parent / f"{d_resolved.name}_整理输出"))
            self._save_settings()
            self._append_log(
                f"📥 已拖入【待整理文件夹】：{d_resolved}（共识别到 {len(docxs)} 份 .docx 文档）",
                target_tab="main",
            )
            self.stage_var.set(f"📥 已载入待整理文件夹：{d_resolved.name}（共 {len(docxs)} 份文档）")
            return

        docxs = self._extract_docx_from_paths(paths)
        if not docxs:
            messagebox.showinfo("拖拽提示", "【待整理文档/目录】请拖入 .docx 文档或包含多个 .docx 的文件夹。")
            return
        if len(docxs) == 1:
            picked = docxs[0]
            self.input_docx_var.set(str(picked))
            if not self.output_dir_var.get().strip():
                self.output_dir_var.set(str(picked.parent))
            self._save_settings()
            self._append_log(f"📥 已拖入【待整理文档】：{picked}", target_tab="main")
            self.stage_var.set(f"📥 已载入待整理文档：{picked.name}")
        else:
            joined = " ; ".join(str(p) for p in docxs)
            self.input_docx_var.set(joined)
            if not self.output_dir_var.get().strip():
                self.output_dir_var.set(str(docxs[0].parent / "零散改点整理输出"))
            self._save_settings()
            self._append_log(
                f"📥 已批量拖入 {len(docxs)} 份【待整理文档】：{docxs[0].name} 等",
                target_tab="main",
            )
            self.stage_var.set(f"📥 已批量载入 {len(docxs)} 份待整理文档")

    def _on_drop_library_dir(self, paths: list[Path]):
        p = paths[0].resolve()
        target_dir = p if p.is_dir() else p.parent
        if not target_dir.is_dir():
            return
        self.library_dir_var.set(str(target_dir))
        self._save_settings()
        self._append_log(f"📥 已拖入【对应语种总库目录】：{target_dir}", target_tab="main")
        self.stage_var.set(f"📥 已载入总库目录：{target_dir.name}")
        self._warm_library_index_async(str(target_dir), target_tab="main")

    def _on_drop_output_dir(self, paths: list[Path]):
        p = paths[0].resolve()
        target_dir = p if p.is_dir() else p.parent
        if not target_dir.is_dir():
            return
        self.output_dir_var.set(str(target_dir))
        self._save_settings()
        self._append_log(f"📥 已拖入【输出保存目录】：{target_dir}", target_tab="main")
        self.stage_var.set(f"📥 已设定输出目录：{target_dir.name}")

    def _handle_tab1_drop(self, paths: list[Path]):
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
        pending_inp_docs: list[Path] = []
        for df in docx_files:
            if any(tok in df.name for tok in ("最新中英", "标you")) and (
                self.input_docx_var.get().strip() or len(docx_files) >= 2
            ):
                self.sc1_master_path_var.set(str(df))
                self.sc1_sync_master_var.set(True)
                applied.append(f"最新中英底稿={df.name}")
                self._append_log(f"📥 已识别并载入【最新中英底稿】：{df}", target_tab="main")
            else:
                pending_inp_docs.append(df)

        if len(pending_inp_docs) == 1:
            df = pending_inp_docs[0]
            self.input_docx_var.set(str(df))
            if not self.output_dir_var.get().strip():
                self.output_dir_var.set(str(df.parent))
            applied.append(f"待整理文档={df.name}")
            self._append_log(f"📥 已载入【待整理文档】：{df}", target_tab="main")
        elif len(pending_inp_docs) > 1:
            self.input_docx_var.set(" ; ".join(str(x) for x in pending_inp_docs))
            if not self.output_dir_var.get().strip():
                self.output_dir_var.set(str(pending_inp_docs[0].parent / "零散改点整理输出"))
            applied.append(f"待整理文档={len(pending_inp_docs)}份")
            self._append_log(
                f"📥 已批量载入 {len(pending_inp_docs)} 份【待整理文档】：{pending_inp_docs[0].name} 等",
                target_tab="main",
            )

        def _looks_like_library_dir(dir_p: Path) -> bool:
            return any(tok in dir_p.name for tok in ("神话", "总库", "七卷", "1-1-1"))

        if len(dirs) == 1:
            d = dirs[0]
            if "未入库" in d.name:
                self.sc1_dedup_dir_var.set(str(d))
                self.sc1_enable_dedup_var.set(True)
                applied.append(f"未入库改点包={d.name}")
                self._append_log(f"📥 已识别并载入【未入库改点包目录】：{d}", target_tab="main")
            elif any(tok in d.name for tok in ("最新中英", "中英底稿", "标you")):
                self.sc1_master_path_var.set(str(d))
                self.sc1_sync_master_var.set(True)
                applied.append(f"最新中英底稿目录={d.name}")
                self._append_log(f"📥 已识别并载入【最新中英底稿目录】：{d}", target_tab="main")
            elif any(tok in d.name.lower() for tok in ("输出", "output")):
                self.output_dir_var.set(str(d))
                applied.append(f"输出目录={d.name}")
                self._append_log(f"📥 已填入【输出保存目录】：{d}", target_tab="main")
            elif _looks_like_library_dir(d):
                self.library_dir_var.set(str(d))
                applied.append(f"总库目录={d.name}")
                self._append_log(f"📥 已载入【对应语种总库目录】：{d}", target_tab="main")
                self._warm_library_index_async(str(d), target_tab="main")
            elif not self.input_docx_var.get().strip() or any(tok in d.name for tok in ("待整理", "改点", "选段")):
                cnt_d = len(resolve_sc1_input_docx_list(d)) if callable(resolve_sc1_input_docx_list) else len(self._extract_docx_from_paths([d]))
                self.input_docx_var.set(str(d))
                if not self.output_dir_var.get().strip():
                    self.output_dir_var.set(str(d.parent / f"{d.name}_整理输出"))
                applied.append(f"待整理文件夹={d.name}({cnt_d}份)")
                self._append_log(f"📥 已载入【待整理文件夹】：{d}（共 {cnt_d} 份 .docx 文档）", target_tab="main")
            else:
                self.library_dir_var.set(str(d))
                applied.append(f"总库目录={d.name}")
                self._append_log(f"📥 已载入【对应语种总库目录】：{d}", target_tab="main")
                self._warm_library_index_async(str(d), target_tab="main")
        elif len(dirs) >= 2:
            lib_assigned = bool(self.library_dir_var.get().strip())
            inp_assigned = bool(self.input_docx_var.get().strip())
            for d in dirs:
                if "未入库" in d.name:
                    self.sc1_dedup_dir_var.set(str(d))
                    self.sc1_enable_dedup_var.set(True)
                    applied.append(f"未入库改点包={d.name}")
                    continue
                if any(tok in d.name for tok in ("最新中英", "中英底稿", "标you")):
                    self.sc1_master_path_var.set(str(d))
                    self.sc1_sync_master_var.set(True)
                    applied.append(f"最新中英底稿目录={d.name}")
                    continue
                if _looks_like_library_dir(d) or (inp_assigned and not lib_assigned):
                    self.library_dir_var.set(str(d))
                    lib_assigned = True
                    applied.append(f"总库目录={d.name}")
                    self._append_log(f"📥 已载入【对应语种总库目录】：{d}", target_tab="main")
                    self._warm_library_index_async(str(d), target_tab="main")
                elif not inp_assigned:
                    cnt_d = len(resolve_sc1_input_docx_list(d)) if callable(resolve_sc1_input_docx_list) else len(self._extract_docx_from_paths([d]))
                    self.input_docx_var.set(str(d))
                    inp_assigned = True
                    applied.append(f"待整理文件夹={d.name}({cnt_d}份)")
                    self._append_log(f"📥 已载入【待整理文件夹】：{d}（共 {cnt_d} 份 .docx 文档）", target_tab="main")
                elif not lib_assigned:
                    self.library_dir_var.set(str(d))
                    lib_assigned = True
                    applied.append(f"总库目录={d.name}")
                    self._append_log(f"📥 已载入【对应语种总库目录】：{d}", target_tab="main")
                    self._warm_library_index_async(str(d), target_tab="main")
                else:
                    self.output_dir_var.set(str(d))
                    applied.append(f"输出目录={d.name}")
                    self._append_log(f"📥 已填入【输出保存目录】：{d}", target_tab="main")

        if not applied:
            messagebox.showinfo("拖拽提示", "请拖入 .docx 文档或文件夹。")
            return

        self._save_settings()
        self.stage_var.set("📥 拖拽载入成功：" + "  |  ".join(applied))

    _on_drop_tab1_general = _handle_tab1_drop

    # ---- Tab 2 (Full-Article) Methods & Drop Handlers ----

    def _on_fa_lang_selected(self, _event=None):
        code = self._get_fa_lang_code()
        if code and code != "AUTO":
            self._record_recent_lang_and_refresh(code, save=True)
        else:
            self._save_settings()

    def _apply_fa_mode_ui(self):
        use_titles = bool(self.fa_use_titles_var.get())
        if use_titles:
            self.btn_fa_import_titles.configure(state=tk.NORMAL)
            self.btn_fa_clear_titles.configure(state=tk.NORMAL)
            self.fa_titles_text.configure(state=tk.NORMAL, bg="#ffffff", fg="#0f172a")
            self._fa_update_title_count()
            if not self.is_running:
                self.fa_stage_var.set("状态：就绪（粘贴或导入篇目标题并设定目录后点击开始）")
        else:
            self.btn_fa_import_titles.configure(state=tk.DISABLED)
            self.btn_fa_clear_titles.configure(state=tk.DISABLED)
            self.fa_titles_text.configure(state=tk.DISABLED, bg="#f1f5f9", fg="#64748b")
            self.fa_title_count_var.set(
                "当前为【现成双文件夹直连对比模式】：无需输入标题，点击开始即按顺序逐篇对比并以中英底稿命名"
            )
            if not self.is_running:
                self.fa_stage_var.set(
                    "状态：现成双文件夹直连对比模式（无需输入标题，点击开始即按顺序逐篇对比并以中英底稿命名）"
                )

    def _on_fa_mode_toggled(self):
        self._apply_fa_mode_ui()
        self._save_settings()

    def _refresh_master_index_manual(self):
        raw = self.fa_master_dir_var.get().strip()
        if not raw:
            messagebox.showinfo("提示", "请先选择【最新中英底稿目录】，再点击「刷新底稿」。")
            return
        try:
            rp = Path(raw).resolve()
        except Exception:
            messagebox.showwarning("提示", "【最新中英底稿目录】路径格式无效。")
            return
        if not rp.is_dir():
            messagebox.showwarning("提示", "请先选择真实存在的【最新中英底稿目录】。")
            return

        self._save_settings()
        norm_key = f"master::{os.path.normcase(str(rp))}"
        with self._indexing_lock:
            if norm_key in self._indexing_inflight:
                self._log("⏳ 中英底稿目录正在后台刷新中，请稍候...", target_tab="fa")
                return
            self._indexing_inflight.add(norm_key)

        self._log(f"🔄 正在强制重新扫描并刷新最新中英底稿索引：{rp} ...", target_tab="fa")

        def _bg():
            t0 = time.perf_counter()
            try:
                if callable(refresh_master_index):
                    info = _call_compatible(
                        refresh_master_index,
                        str(rp),
                        log_fn=lambda m: self._log(m, target_tab="fa"),
                        force_rebuild=True,
                    )
                    elapsed_sec = float(
                        info.get("elapsed_sec", time.perf_counter() - t0) if isinstance(info, dict) else (time.perf_counter() - t0)
                    )
                    total_files = info.get("total_files", 0) if isinstance(info, dict) else 0
                    total_arts = info.get("total_articles", info.get("article_count", 0)) if isinstance(info, dict) else 0
                    extra = f"，解析篇目 {total_arts} 篇" if total_arts else ""
                    self._log(
                        f"✅ 最新中英底稿索引已强制刷新：共扫描 {total_files} 份底稿文件{extra}（耗时 {elapsed_sec:.1f}s）",
                        target_tab="fa",
                    )
                else:
                    docxs = [x for x in rp.rglob("*.docx") if not x.name.startswith("~$")]
                    for dx in docxs:
                        dx.stat()
                    elapsed_sec = time.perf_counter() - t0
                    self._log(
                        f"✅ 最新中英底稿目录已刷新校验：共发现 {len(docxs)} 份底稿文档（耗时 {elapsed_sec:.1f}s，每次运行均实时校验磁盘最新内容）",
                        target_tab="fa",
                    )
            except Exception as ex:
                self._log(f"⚠️ 刷新中英底稿索引出错：{_format_exception_concise(ex)}", target_tab="fa")
            finally:
                with self._indexing_lock:
                    self._indexing_inflight.discard(norm_key)

        threading.Thread(target=_bg, daemon=True).start()

    def _fa_get_titles_list(self) -> list[str]:
        raw = self.fa_titles_text.get("1.0", tk.END)
        return _fallback_parse_titles_multiline(raw)

    def _fa_update_title_count(self):
        if not self.fa_use_titles_var.get():
            self.fa_title_count_var.set(
                "当前为【现成双文件夹直连对比模式】：无需输入标题，点击开始即按顺序逐篇对比并以中英底稿命名"
            )
            return
        titles = self._fa_get_titles_list()
        self.fa_title_count_var.set(f"当前共输入 {len(titles)} 个待提取篇目标题")

    def _fa_append_titles(self, new_titles: list[str], source_label: str = ""):
        if not new_titles:
            messagebox.showinfo("提示", "未在所选文件中识别到非空标题行。")
            return
        if not self.fa_use_titles_var.get():
            self.fa_use_titles_var.set(True)
            self._apply_fa_mode_ui()
        existing = self._fa_get_titles_list()
        seen = set(existing)
        merged = list(existing)
        added = 0
        for t in new_titles:
            if t not in seen:
                seen.add(t)
                merged.append(t)
                added += 1
        self.fa_titles_text.configure(state=tk.NORMAL)
        self.fa_titles_text.delete("1.0", tk.END)
        self.fa_titles_text.insert("1.0", "\n".join(merged) + "\n")
        self._fa_update_title_count()
        if source_label:
            self._append_log(f"📥 已从《{source_label}》导入 {added} 个篇目标题（当前共 {len(merged)} 个）。", target_tab="fa")
            self.fa_stage_var.set(f"📥 已导入标题清单：共 {len(merged)} 个篇目")

    def _fa_import_titles_file(self):
        p = filedialog.askopenfilename(
            title="选择包含待对比篇目标题的 .docx 或 .txt 文档",
            filetypes=[("标题清单文档", "*.docx *.txt"), ("Word 文档", "*.docx"), ("文本文档", "*.txt")],
        )
        if not p:
            return
        try:
            titles = _fallback_parse_titles_file(p)
            self._fa_append_titles(titles, source_label=Path(p).name)
            if not self.fa_output_dir_var.get().strip():
                self.fa_output_dir_var.set(str(Path(p).resolve().parent))
                self._save_settings()
        except Exception as ex:
            messagebox.showerror("导入失败", _format_exception_concise(ex))

    def _fa_clear_titles(self):
        if not self.fa_use_titles_var.get():
            return
        self.fa_titles_text.delete("1.0", tk.END)
        self._fa_update_title_count()

    def _fa_browse_master_dir(self):
        d = filedialog.askdirectory(title="选择最新中英底稿目录（支持单篇或合集大文档目录）")
        if d:
            self.fa_master_dir_var.set(d)
            self._save_settings()

    def _fa_browse_library_dir(self):
        d = filedialog.askdirectory(title="选择外文三语总库目录")
        if d:
            self.fa_library_dir_var.set(d)
            self._save_settings()
            self._warm_library_index_async(d, target_tab="fa")

    def _fa_browse_output_dir(self):
        d = filedialog.askdirectory(title="选择全篇对比输出保存目录")
        if d:
            self.fa_output_dir_var.set(d)
            self._save_settings()

    def _fa_open_output_dir(self):
        d = self.fa_output_dir_var.get().strip()
        p = Path(d).resolve() if d else None
        if p and p.exists() and p.is_dir():
            open_path_in_os(p)
        else:
            messagebox.showinfo("提示", "请先选择真实存在的输出保存目录。")

    def _fa_browse_dedup_dir(self):
        d = filedialog.askdirectory(title="选择「未入库改点包」文件夹")
        if d:
            self.fa_dedup_dir_var.set(d)
            self.fa_enable_dedup_var.set(True)
            self._save_settings()

    def _on_drop_fa_titles_file(self, paths: list[Path]):
        picked_files = [
            p.resolve()
            for p in paths
            if p.exists() and p.is_file() and p.suffix.lower() in (".docx", ".txt") and not p.name.startswith("~$")
        ]
        if not picked_files:
            messagebox.showinfo("拖拽提示", "标题输入框请拖入包含篇目标题的 .docx 或 .txt 文件。")
            return
        all_t: list[str] = []
        for pf in picked_files:
            all_t.extend(_fallback_parse_titles_file(pf))
        self._fa_append_titles(all_t, source_label=picked_files[0].name)

    def _on_drop_fa_master_dir(self, paths: list[Path]):
        p = paths[0].resolve()
        target = p if p.is_dir() else p.parent
        if target.is_dir():
            self.fa_master_dir_var.set(str(target))
            self._save_settings()
            self._append_log(f"📥 已设定【最新中英底稿目录】：{target}", target_tab="fa")

    def _on_drop_fa_library_dir(self, paths: list[Path]):
        p = paths[0].resolve()
        target = p if p.is_dir() else p.parent
        if target.is_dir():
            self.fa_library_dir_var.set(str(target))
            self._save_settings()
            self._append_log(f"📥 已设定【外文三语总库目录】：{target}", target_tab="fa")
            self._warm_library_index_async(str(target), target_tab="fa")

    def _on_drop_fa_output_dir(self, paths: list[Path]):
        p = paths[0].resolve()
        target = p if p.is_dir() else p.parent
        if target.is_dir():
            self.fa_output_dir_var.set(str(target))
            self._save_settings()
            self._append_log(f"📥 已设定【输出保存目录】：{target}", target_tab="fa")

    def _on_drop_fa_dedup_dir(self, paths: list[Path]):
        p = paths[0].resolve()
        target = p if p.is_dir() else p.parent
        if target.is_dir():
            self.fa_dedup_dir_var.set(str(target))
            self.fa_enable_dedup_var.set(True)
            self._save_settings()
            self._append_log(f"📥 已设定【未入库改点包目录】：{target}", target_tab="fa")

    def _handle_tab2_drop(self, paths: list[Path]):
        title_files: list[Path] = []
        dirs: list[Path] = []
        for p in paths:
            try:
                rp = p.resolve()
                if rp.is_dir():
                    dirs.append(rp)
                elif rp.is_file() and rp.suffix.lower() in (".docx", ".txt") and not rp.name.startswith("~$"):
                    title_files.append(rp)
            except Exception:
                continue

        if title_files:
            merged_t: list[str] = []
            for tf in title_files:
                merged_t.extend(_fallback_parse_titles_file(tf))
            self._fa_append_titles(merged_t, source_label=title_files[0].name)
            if not self.fa_output_dir_var.get().strip():
                self.fa_output_dir_var.set(str(title_files[0].parent))

        for d in dirs:
            dname = d.name
            if "未入库" in dname:
                self.fa_dedup_dir_var.set(str(d))
                self.fa_enable_dedup_var.set(True)
                self._append_log(f"📥 已识别并填入【未入库改点包目录】：{d}", target_tab="fa")
            elif detect_lang_code_from_hint(dname, self.languages) and (
                self.fa_master_dir_var.get().strip() or "中英" not in dname or len(dname) >= 2
            ) and not any(tok in dname for tok in ("标you", "底稿", "中文", "中英底稿")):
                self.fa_library_dir_var.set(str(d))
                self._append_log(f"📥 已识别并填入【外文三语总库目录】：{d}", target_tab="fa")
                self._warm_library_index_async(str(d), target_tab="fa")
            elif not self.fa_master_dir_var.get().strip() or any(tok in dname for tok in ("标you", "底稿", "出书")):
                self.fa_master_dir_var.set(str(d))
                self._append_log(f"📥 已填入【最新中英底稿目录】：{d}", target_tab="fa")
            elif not self.fa_library_dir_var.get().strip():
                self.fa_library_dir_var.set(str(d))
                self._append_log(f"📥 已填入【外文三语总库目录】：{d}", target_tab="fa")
                self._warm_library_index_async(str(d), target_tab="fa")
            else:
                self.fa_output_dir_var.set(str(d))
                self._append_log(f"📥 已填入【输出保存目录】：{d}", target_tab="fa")

        self._save_settings()

    def _on_run_full_article(self):
        self._save_settings()
        use_titles = bool(self.fa_use_titles_var.get())
        titles = self._fa_get_titles_list() if use_titles else []
        if use_titles and not titles:
            messagebox.showwarning("提示", "请先输入或导入至少 1 个待提取对比的篇目标题（或取消勾选切换为双文件夹直连对比模式）。")
            return
        master_dir = self.fa_master_dir_var.get().strip()
        foreign_lib = self.fa_library_dir_var.get().strip()
        out_dir = self.fa_output_dir_var.get().strip()
        if not master_dir or not foreign_lib or not out_dir:
            messagebox.showwarning("提示", "请完整指定【最新中英底稿目录】、【外文三语总库目录】和【输出保存目录】。")
            return
        if not Path(master_dir).is_dir() or not Path(foreign_lib).is_dir():
            messagebox.showwarning("提示", "请确保【最新中英底稿目录】与【外文三语总库目录】均为真实存在的文件夹。")
            return

        dedup_dir = self.fa_dedup_dir_var.get().strip() if self.fa_enable_dedup_var.get() else ""
        if self.fa_enable_dedup_var.get() and (not dedup_dir or not Path(dedup_dir).is_dir()):
            messagebox.showwarning("提示", "已勾选关联「未入库改点包」查重，请选择有效的未入库改点包文件夹（或取消勾选）。")
            return

        if not callable(run_full_article_pipeline) or (use_titles and not callable(resolve_full_articles)):
            messagebox.showerror("组件未就绪", "未找到 full_article_processor 模块，请确认程序文件完整。")
            return
        if not use_titles and not callable(resolve_direct_folder_articles):
            messagebox.showerror("组件未就绪", "未找到 full_article_processor.resolve_direct_folder_articles，请确认已合并 v1.8.0 引擎模块。")
            return

        self._sync_fa_hl_mode_vars()
        highlight_mode = self.fa_highlight_mode_var.get() or "all_three"
        highlight_colors = highlight_mode != "revisions_only"
        hl_label = (
            "英外标色 + 中文改点下划线"
            if highlight_mode == "all_three"
            else ("仅英外标色（中文不标下划线）" if highlight_mode == "en_foreign" else "仅输出对比修订")
        )
        yellow_gap_words = self._get_yellow_gap_words()
        want_ai = bool(self.fa_enable_ai_var.get())
        keys = self._get_all_keys_list()
        effective_ai = bool(want_ai and keys and highlight_colors)
        model = self._get_model_id()
        api_mode = self._get_api_mode_code()
        rot_mode = self._get_rotation_mode_code()
        lang_code = self._get_fa_lang_code()
        if lang_code and lang_code != "AUTO":
            self._record_recent_lang_and_refresh(lang_code, save=True)
        dt_str = self.fa_date_var.get().strip()
        balance_cols = bool(self.fa_balance_cols_var.get())
        pool = ApiKeyPool(keys, rotation_mode=rot_mode) if effective_ai else ""

        init_stage = (
            f"正在扫描中英底稿库并解析 {len(titles)} 个篇目标题..."
            if use_titles
            else "正在按现成双文件夹直连模式扫描并配对全部篇目..."
        )
        self._set_running(True, final_stage=init_stage, tab_key="fa")
        self._log("=" * 64, target_tab="fa")
        if not highlight_colors:
            mode_desc = f"{hl_label}（不标黄/青高亮，免调 AI）"
        elif effective_ai:
            mode_desc = f"{hl_label} | AI 辅助（{model} | {len(keys)} 个 Key | 黄标连片 ≤{yellow_gap_words}词）"
        else:
            mode_desc = f"{hl_label} | 本地位置估算模式（免 API Key | 黄标连片 ≤{yellow_gap_words}词）"
        src_desc = f"按标题提取 {len(titles)} 篇" if use_titles else "现成双文件夹直连对比"
        self._log(
            f"⚡ 启动【全篇改点整理】 | {src_desc} | {mode_desc} | 语种：{lang_code}",
            target_tab="fa",
        )
        if highlight_colors and want_ai and not keys:
            self._log("ℹ️ 未填写 API Key，本次自动使用「本地位置估算模式」完成分级与标色", target_tab="fa")

        def worker():
            try:
                if use_titles:
                    self._progress(5.0, f"正在中英底稿目录中检索 {len(titles)} 个篇目标题...")
                    res_lookup = _call_compatible(
                        resolve_full_articles,
                        titles=titles,
                        cn_en_master_dir=master_dir,
                        foreign_lib_dir=foreign_lib,
                        log_fn=lambda m: self._log(m, target_tab="fa"),
                    )
                else:
                    self._progress(5.0, "正在按顺序配对「最新中英底稿包」与「外文三语文件包」...")
                    res_lookup = _call_compatible(
                        resolve_direct_folder_articles,
                        cn_en_master_dir=master_dir,
                        foreign_lib_dir=foreign_lib,
                        log_fn=lambda m: self._log(m, target_tab="fa"),
                    )
                resolved_list = list(res_lookup.get("resolved") or res_lookup.get("articles") or [])
                ambiguous_list = list(res_lookup.get("ambiguous") or [])
                unmatched_list = list(res_lookup.get("unmatched") or [])

                if unmatched_list:
                    self._log(
                        f"⚠️ 以下 {len(unmatched_list)} 项未匹配到对应篇目：{', '.join(str(x) for x in unmatched_list)}",
                        target_tab="fa",
                    )

                chosen_amb: list[dict] = []
                if ambiguous_list:
                    self._log(
                        f"🔎 检测到 {len(ambiguous_list)} 个同名多篇标题，正在等待弹窗确认具体卷/分部...",
                        target_tab="fa",
                    )
                    holder: dict = {"selected": [], "cancelled": True}
                    ev = threading.Event()
                    self.log_queue.put(("fa_disambiguate", (ambiguous_list, len(resolved_list), holder, ev)))
                    ev.wait()
                    if holder.get("cancelled"):
                        msg = "已取消多篇同名标题选择，任务中止。"
                        self._log(f"ℹ️ {msg}", target_tab="fa")
                        self.log_queue.put(("done", (False, msg, "fa")))
                        return
                    chosen_amb = list(holder.get("selected") or [])
                    self._log(f"✅ 已确认选择 {len(chosen_amb)} 篇同名候选篇目。", target_tab="fa")

                selected_articles = resolved_list + chosen_amb
                if not selected_articles:
                    unm_hint = f"（未命中：{', '.join(str(x) for x in unmatched_list)}）" if unmatched_list else ""
                    msg = f"未找到可对比的篇目{unm_hint}"
                    self._log(f"❌ {msg}", target_tab="fa")
                    self.log_queue.put(("done", (False, msg, "fa")))
                    self.log_queue.put(("popup_error", ("未匹配到篇目", msg)))
                    return

                pipe_res = _call_compatible(
                    run_full_article_pipeline,
                    selected_articles=selected_articles,
                    foreign_lib_dir=foreign_lib,
                    output_dir=out_dir,
                    api_key=pool,
                    model=model,
                    api_mode=api_mode,
                    rotation_mode=rot_mode,
                    lang=lang_code,
                    enable_ai=effective_ai,
                    balance_columns=balance_cols,
                    dedup_pkg_dir=dedup_dir,
                    date_str=dt_str,
                    use_title_list=use_titles,
                    highlight_colors=highlight_colors,
                    highlight_mode=highlight_mode,
                    yellow_gap_words=yellow_gap_words,
                    log_fn=lambda m: self._log(m, target_tab="fa"),
                    progress_fn=self._progress,
                )
                saved_files = list(pipe_res.get("saved_files") or [])
                dup_reports = list(pipe_res.get("duplicate_reports") or [])
                unchanged_articles = list(pipe_res.get("unchanged_articles") or [])
                unchanged_report = str(
                    pipe_res.get("unchanged_report") or pipe_res.get("unchanged_report_path") or ""
                )
                if not unchanged_report:
                    for sf in saved_files:
                        if "无改点篇目汇总报告" in str(sf):
                            unchanged_report = str(sf)
                            break

                det_lang = pipe_res.get("lang_code") or lang_code
                if det_lang and str(det_lang).upper() not in ("AUTO", "XX"):
                    self.log_queue.put(("record_recent_lang", str(det_lang)))
                det_disp = find_lang_display_name(det_lang, self.languages)

                files_str = "\n".join(saved_files)
                extra_lines: list[str] = []
                if unchanged_articles or unchanged_report:
                    unch_cnt = len(unchanged_articles)
                    rep_note = f"\n   汇总报告：{unchanged_report}" if unchanged_report else ""
                    unch_msg = f"ℹ️ 英文对比无改点篇目：共 {unch_cnt} 篇（已自动跳过单篇导出并归入汇总报告）{rep_note}"
                    self._log(unch_msg, target_tab="fa")
                    extra_lines.append(unch_msg)
                if dup_reports:
                    extra_lines.append(f"🔍 已同步输出未入库查重明细报告：{len(dup_reports)} 份")
                if unmatched_list:
                    extra_lines.append(f"⚠️ 未命中项（{len(unmatched_list)} 个）：{', '.join(str(x) for x in unmatched_list)}")
                extra_txt = ("\n\n" + "\n".join(extra_lines)) if extra_lines else ""

                unch_tag = f"（含无改点 {len(unchanged_articles)} 篇）" if unchanged_articles else ""
                self._log(
                    f"🎉 全篇改点整理完成！（{det_disp}）共处理 {len(selected_articles)} 篇{unch_tag}，导出 {len(saved_files)} 份文档：\n{files_str}",
                    target_tab="fa",
                )
                self.log_queue.put(
                    (
                        "done",
                        (
                            True,
                            f"全篇对比完成！共处理 {len(selected_articles)} 篇{unch_tag}，输出 {len(saved_files)} 份文档",
                            "fa",
                        ),
                    )
                )
                self.log_queue.put(
                    (
                        "popup_info",
                        (
                            "全篇改点整理完成",
                            f"语种：{det_disp}  |  完成篇数：{len(selected_articles)} 篇{unch_tag}\n"
                            f"共输出 {len(saved_files)} 份文档：\n\n{files_str}{extra_txt}",
                        ),
                    )
                )
            except Exception as ex:
                err_msg = _format_exception_concise(ex, keys=keys)
                self._log(f"❌ 全篇改点整理出错：{err_msg}", target_tab="fa")
                self.log_queue.put(("done", (False, f"全篇对比出错：{err_msg}", "fa")))
                self.log_queue.put(("popup_error", ("全篇改点整理出错", err_msg)))

        threading.Thread(target=worker, daemon=True).start()

    # ---- Tab 3 (Standalone Dedup Checker) Methods & Drop Handlers ----

    def _dd_refresh_tree(self, status_map: dict[str, str] | None = None):
        for item in self.dd_tree.get_children():
            self.dd_tree.delete(item)
        smap = status_map or {}
        for idx, fp in enumerate(self.dd_files):
            p = Path(fp)
            st = smap.get(os.path.normcase(str(p)), "待查重")
            self.dd_tree.insert(
                "",
                tk.END,
                iid=str(idx),
                values=(f"#{idx + 1}", p.name, str(p.parent), st),
            )
        self.dd_summary_var.set(f"当前待查重列表共 {len(self.dd_files)} 份文档  |  支持拖入待查重 .docx 或文件夹")

    def _dd_ingest_paths(self, paths: list[Path]):
        found = [
            p
            for p in self._extract_docx_from_paths(paths)
            if not p.name.startswith("查重结果")
        ]
        if found and self._dd_batch_completed:
            self.dd_files.clear()
            self._dd_batch_completed = False
            self._append_log("🧹 已自动清空上一轮已完成查重列表，载入新文档。", target_tab="dd")
        if not found:
            messagebox.showinfo("提示", "所选项目中未找到待查重的有效 .docx 文档。")
            return
        existing_norm = {os.path.normcase(str(Path(x).resolve())) for x in self.dd_files}
        for rp in found:
            k = os.path.normcase(str(rp))
            if k not in existing_norm:
                existing_norm.add(k)
                self.dd_files.append(str(rp))
        if not self.dd_out_dir_var.get().strip() and self.dd_files:
            self.dd_out_dir_var.set(str(Path(self.dd_files[0]).parent))
            self._save_settings()
        self._dd_refresh_tree()
        self._append_log(f"📥 已载入待查重文档：当前队列共 {len(self.dd_files)} 份。", target_tab="dd")

    def _dd_add_files(self):
        ps = filedialog.askopenfilenames(
            title="选择待查重的改点 Word 文档（支持多选）",
            filetypes=[("Word 文档", "*.docx")],
        )
        if ps:
            self._dd_ingest_paths([Path(x) for x in ps])

    def _dd_add_folder(self):
        d = filedialog.askdirectory(title="选择包含待查重 .docx 文档的文件夹")
        if d:
            self._dd_ingest_paths([Path(d)])

    def _dd_remove_selected(self):
        sel = self.dd_tree.selection()
        if not sel:
            return
        for idx in sorted((int(i) for i in sel), reverse=True):
            if 0 <= idx < len(self.dd_files):
                self.dd_files.pop(idx)
        self._dd_refresh_tree()

    def _dd_clear_all(self):
        self.dd_files.clear()
        self._dd_refresh_tree()

    def _dd_browse_pkg_dir(self):
        d = filedialog.askdirectory(title="选择「未入库改点包」文件夹")
        if d:
            self.dd_pkg_dir_var.set(d)
            self._save_settings()

    def _dd_browse_out_dir(self):
        d = filedialog.askdirectory(title="选择查重报告与去重稿输出保存目录")
        if d:
            self.dd_out_dir_var.set(d)
            self._save_settings()

    def _dd_open_out_dir(self):
        d = self.dd_out_dir_var.get().strip()
        p = Path(d).resolve() if d else None
        if p and p.exists() and p.is_dir():
            open_path_in_os(p)
        else:
            messagebox.showinfo("提示", "请先选择真实存在的输出目录。")

    def _dd_open_selected_folder(self, _event=None):
        sel = self.dd_tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        if 0 <= idx < len(self.dd_files):
            p = Path(self.dd_files[idx])
            if p.parent.exists():
                open_path_in_os(p.parent)

    def _on_drop_dd_pkg_dir(self, paths: list[Path]):
        p = paths[0].resolve()
        target = p if p.is_dir() else p.parent
        if target.is_dir():
            self.dd_pkg_dir_var.set(str(target))
            self._save_settings()
            self._append_log(f"📥 已设定【未入库改点包目录】：{target}", target_tab="dd")

    def _on_drop_dd_out_dir(self, paths: list[Path]):
        p = paths[0].resolve()
        target = p if p.is_dir() else p.parent
        if target.is_dir():
            self.dd_out_dir_var.set(str(target))
            self._save_settings()
            self._append_log(f"📥 已设定【查重输出目录】：{target}", target_tab="dd")

    def _handle_tab3_drop(self, paths: list[Path]):
        pkg_dirs = [p.resolve() for p in paths if p.exists() and p.is_dir() and "未入库" in p.name]
        other_paths = [p for p in paths if not (p.exists() and p.is_dir() and "未入库" in p.name)]
        if pkg_dirs:
            self.dd_pkg_dir_var.set(str(pkg_dirs[0]))
            self._save_settings()
            self._append_log(f"📥 已识别并填入【未入库改点包目录】：{pkg_dirs[0]}", target_tab="dd")
        if other_paths:
            self._dd_ingest_paths(other_paths)

    def _on_run_dedup(self):
        self._save_settings()
        if not self.dd_files:
            messagebox.showwarning("提示", "请先选择或拖入需要查重的改点 .docx 文档。")
            return
        pkg_dir = self.dd_pkg_dir_var.get().strip()
        out_dir = self.dd_out_dir_var.get().strip()
        if not pkg_dir or not Path(pkg_dir).is_dir():
            messagebox.showwarning("提示", "请选择真实存在的【未入库改点包目录】。")
            return
        if not out_dir:
            messagebox.showwarning("提示", "请指定【查重报告与去重稿输出目录】。")
            return
        if not callable(check_and_export_duplicates):
            messagebox.showerror("组件未就绪", "未找到 dedup_checker.check_and_export_duplicates，请检查程序文件完整性。")
            return

        files_snapshot = list(self.dd_files)
        export_dedup = bool(self.dd_export_deduped_var.get())
        balance_cols = bool(self.dd_balance_cols_var.get())
        dt_str = self.dd_date_var.get().strip()

        self._set_running(True, final_stage=f"正在执行第 1 列中文四级查重（共 {len(files_snapshot)} 份文档）...", tab_key="dd")
        self._log("=" * 64, target_tab="dd")
        self._log(
            f"🔍 启动【改点稿查重】 | 待查重文档：{len(files_snapshot)} 份 | 未入库包：{pkg_dir}",
            target_tab="dd",
        )

        def worker():
            try:
                res = _call_compatible(
                    check_and_export_duplicates,
                    input_files=files_snapshot,
                    unmerged_pkg_dir=pkg_dir,
                    output_dir=out_dir,
                    export_deduped_files=export_dedup,
                    balance_columns=balance_cols,
                    date_str=dt_str,
                    log_fn=lambda m: self._log(m, target_tab="dd"),
                    progress_fn=self._progress,
                )
                rep_path = str(res.get("report_path") or "")
                deduped = list(res.get("deduped_files") or [])
                tot_rows = int(res.get("total_checked_rows") or 0)
                dup_cnt = int(res.get("duplicate_count") or 0)

                summary_line = (
                    f"查重完成！共检查 {tot_rows} 段，发现重复段落 {dup_cnt} 处；报告已生成：{Path(rep_path).name if rep_path else out_dir}"
                )
                self._log(f"🎉 {summary_line}", target_tab="dd")
                if rep_path:
                    self._log(f"📄 查重明细报告：{rep_path}", target_tab="dd")
                for df in deduped:
                    self._log(f"✅ 去重定稿：{df}", target_tab="dd")

                self.log_queue.put(("dd_mark_done", dup_cnt))
                self.log_queue.put(("done", (True, summary_line, "dd")))
                dedup_block = ("\n\n已导出去重定稿：\n" + "\n".join(deduped)) if deduped else ""
                self.log_queue.put(
                    (
                        "popup_info",
                        (
                            "改点稿查重完成",
                            f"共检查文档：{len(files_snapshot)} 份（合计 {tot_rows} 个段落）\n"
                            f"命中未入库重复段落：{dup_cnt} 处\n\n"
                            f"查重结果明细报告：\n{rep_path}{dedup_block}",
                        ),
                    )
                )
            except Exception as ex:
                err_msg = _format_exception_concise(ex)
                self._log(f"❌ 查重执行出错：{err_msg}", target_tab="dd")
                self.log_queue.put(("done", (False, f"查重出错：{err_msg}", "dd")))
                self.log_queue.put(("popup_error", ("查重执行出错", err_msg)))

        threading.Thread(target=worker, daemon=True).start()

    # ---- Tab 4 & Tab 5 Drop Helpers ----

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
            self._append_log(f"📥 已设定【分part稿输出目录】：{target_dir}", target_tab="fel")
            self.fel_stage_var.set(f"📥 已设定输出目录：{target_dir.name}")

    # ---------------- Tab 5: Part Manuscript Formatting & Merging Methods ----------------

    def _fel_ingest_paths(self, paths: list[Path]):
        found_docx = self._extract_docx_from_paths(paths)
        if not found_docx:
            messagebox.showinfo("提示", "拖入或选择的项目中未找到任何有效的 .docx 文档。")
            return

        # Anti-mixing (`防串稿`): if the previous batch already finished running, auto-clear it before loading new files!
        if self._fel_batch_completed and self.fel_files:
            self.fel_files.clear()
            self._fel_inspect_cache.clear()
            if self._get_fel_lang_code() == "AUTO":
                self.fel_tri_prefix_var.set("")
            if self._fel_output_was_auto:
                self.fel_output_dir_var.set("")
            self._fel_batch_completed = False
            self._append_log("🧹 已自动清空上一轮已完成的分part列表，载入新一批文稿（防串稿）。", target_tab="fel")

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
            self._fel_output_was_auto = True

        self._fel_auto_sync_prefix()
        self._fel_refresh_previews()
        self._append_log(f"📥 【分part稿整理/合并】已载入并按序号排序共 {len(self.fel_files)} 份文档。", target_tab="fel")
        self.fel_stage_var.set(f"📥 已就绪：共载入 {len(self.fel_files)} 份文档")

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
        code = self._get_fel_lang_code()
        if code and code != "AUTO":
            self._record_recent_lang_and_refresh(code, save=False)
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
            title="选择待整理/合并的分part Word 文档（支持单篇多Part或多文件）",
            filetypes=[("Word 文档", "*.docx")],
        )
        if paths:
            self._fel_ingest_paths([Path(x) for x in paths])

    def _fel_add_folder(self):
        d = filedialog.askdirectory(title="选择包含各分篇 (.docx) 的文件夹")
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
        self._fel_batch_completed = False
        if self._get_fel_lang_code() == "AUTO":
            self.fel_tri_prefix_var.set("")
        self._fel_refresh_previews()

    def _fel_browse_output(self):
        d = filedialog.askdirectory(title="选择分part稿整理/合并输出保存目录")
        if d:
            self.fel_output_dir_var.set(d)
            self._fel_output_was_auto = False
            self._save_settings()

    def _fel_open_output_dir(self):
        d = self.fel_output_dir_var.get().strip()
        p = Path(d).resolve() if d else None
        if p and p.exists() and p.is_dir():
            open_path_in_os(p)
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
                open_path_in_os(p.parent)

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
                part_disp = f"-{pnum:02d}" if pnum is not None else "单篇/多表"
                total_rows += info["row_count"]
                single_preview = build_fellowship_output_name(pfx, base_title, dt, part_num=pnum)
                if mode == "merged_only":
                    shown_preview = f"[合并源 #{idx + 1}] → 汇总至合并总文档"
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
                clean_ex = _format_exception_concise(ex)
                self.fel_tree.insert(
                    "",
                    tk.END,
                    iid=str(idx),
                    values=(f"#{idx + 1}", p.name, "-", "-", f"解析失败：{clean_ex}"),
                )

        if not self.fel_files:
            self.fel_summary_var.set("当前共 0 份文档  |  支持多篇 Part 1..N 排序合并，或单文档内多 Part 表格一键合并")
        elif mode == "single_only":
            self.fel_summary_var.set(f"共 {total_files} 份文档（合计 {total_rows} 行） | 模式：输出各篇整理稿（单篇内多 Part 表格自动合并）")
        elif mode == "merged_only":
            merged_fn = build_fellowship_output_name(pfx, first_base_title or "合并定稿", dt, part_num=None)
            self.fel_summary_var.set(f"共 {total_files} 份 → 合并输出 1 份总文档：{merged_fn}")
        else:
            merged_fn = build_fellowship_output_name(pfx, first_base_title or "合并定稿", dt, part_num=None)
            self.fel_summary_var.set(f"共 {total_files} 份单篇 + 1 份合并总文档（{merged_fn}）")

    def _on_run_fellowship(self):
        self._save_settings()
        if not self.fel_files:
            messagebox.showwarning("提示", "请先拖入或选择需要整理/合并的分part .docx 文档或文件夹。")
            return
        out_dir = self.fel_output_dir_var.get().strip()
        if not out_dir:
            messagebox.showwarning("提示", "请先指定【输出保存目录】。")
            return

        use_ai = self.fel_enable_ai_var.get()
        sync_red = bool(self.fel_sync_red_var.get())
        keys = self._get_all_keys_list()
        if use_ai and not keys:
            proceed = messagebox.askyesno(
                "未配置 API Key 提醒",
                "当前已勾选「使用 AI 自动同步第 3 列段内局部引文蓝字/加粗」，但尚未填写 API Key。\n\n"
                "• 拆内容控件、关修订清批注、单篇/多篇 Part 表格合并、防错位三列拉齐、弯引号规范、语系标准字体（如泰语 Angsana New 18/16/14pt）及整格/冒号前缀（如 Question 10:）蓝粗均可【本地直接完成】。\n"
                "• 段内局部引文或红色文字对齐在无 Key 时将采用本地估算或跳过 AI。\n\n"
                "是否直接以纯本地极速模式继续执行？",
            )
            if not proceed:
                return

        model = self._get_model_id()
        api_mode = self._get_api_mode_code()
        rot_mode = self._get_rotation_mode_code()
        lang_code = self._get_fel_lang_code()
        if lang_code and lang_code != "AUTO":
            self._record_recent_lang_and_refresh(lang_code, save=True)
        tri_pfx = self.fel_tri_prefix_var.get().strip()
        quote_mode = self._get_fel_quote_mode_code()
        out_mode = self.fel_output_mode_var.get()
        dt_str = self.fel_date_var.get().strip()
        pool = ApiKeyPool(keys, rotation_mode=rot_mode) if (use_ai and keys) else None
        self._fel_inspect_cache.clear()
        files_snapshot = list(self.fel_files)

        ai_tag = "AI开启" if (use_ai and keys) else "纯本地离线极速模式"
        red_tag = " | 同步红字：开启" if sync_red else ""
        self._set_running(True, final_stage=f"正在启动【分part稿整理/合并】（共 {len(files_snapshot)} 份文档）...", tab_key="fel")
        self._log("=" * 64, target_tab="fel")
        self._log(
            f"⚡ 启动【分part稿整理/合并】 | 文档数：{len(files_snapshot)} | 模式：{out_mode} | 语种：{lang_code} | {ai_tag}{red_tag}",
            target_tab="fel",
        )

        def worker():
            try:
                res = _call_compatible(
                    run_fellowship_pipeline,
                    input_files=files_snapshot,
                    output_dir=out_dir,
                    output_mode=out_mode,
                    lang_code=lang_code,
                    tri_prefix_override=tri_pfx,
                    quote_mode=quote_mode,
                    date_yyyymmdd=dt_str,
                    enable_ai_col3=bool(use_ai and keys),
                    sync_red=sync_red,
                    api_key=pool,
                    model=model,
                    api_mode=api_mode,
                    rotation_mode=rot_mode,
                    log_fn=lambda m: self._log(m, target_tab="fel"),
                    progress_fn=self._progress,
                )
                saved = res["saved_files"]
                failed_ai = res.get("ai_failed_items") or []
                retained_cmts = res.get("retained_comments") or []
                files_str = "\n".join(saved)
                self._log(f"🎉 分part稿整理/合并全部完成！共输出 {len(saved)} 份文档（修订状态：关闭）：\n{files_str}", target_tab="fel")
                warn_block = ""
                if retained_cmts:
                    rc_list = "\n".join(f"• {x}" for x in retained_cmts)
                    warn_block += f"\n\n📌 批注核对提醒（以下行批注与正文/颜色不一致，已原样保留供复核）：\n{rc_list}"
                if failed_ai:
                    warn_list = "\n".join(f"• {x}" for x in failed_ai)
                    warn_block += f"\n\n⚠️ 以下段落因模型接口繁忙暂未调用 AI 对齐段内局部样式（其余全部已完成导出，建议在 Word 中核对该行）：\n{warn_list}"
                self.log_queue.put(("done", (True, f"整理/合并完成！共输出 {len(saved)} 份文档（修订: 关闭）", "fel")))
                self.log_queue.put(
                    (
                        "popup_info",
                        (
                            "分part稿整理/合并完成",
                            f"前缀：{res['tri_prefix']}\n已输出 {len(saved)} 份定稿文档（横向 A4 · 修订已关闭）：\n\n{files_str}{warn_block}",
                        ),
                    )
                )
            except SystemExit as se:
                msg = f"整理/合并中止（退出码 {se.code}）：若目标文档已在 Word 中打开，请先关闭后重试。"
                self._log(f"❌ {msg}", target_tab="fel")
                self.log_queue.put(("done", (False, msg, "fel")))
                self.log_queue.put(("popup_error", ("分part稿整理/合并中止", msg)))
            except Exception as ex:
                err_msg = _format_exception_concise(ex, keys=keys)
                self._log(f"❌ 分part稿整理/合并出错：{err_msg}", target_tab="fel")
                self.log_queue.put(("done", (False, f"整理出错：{err_msg}", "fel")))
                self.log_queue.put(("popup_error", ("分part稿整理/合并出错", err_msg)))

        threading.Thread(target=worker, daemon=True).start()

    # ---------------- Tab 4: Word Counter Methods ----------------

    def _on_wc_lang_selected(self, _event=None):
        code = self._get_wc_lang_code()
        if code and code != "AUTO":
            self._record_recent_lang_and_refresh(code, save=False)
        self._wc_refresh_previews()

    def _on_wc_keyword_typed(self, _event=None):
        if self.wc_keyword_var.get().strip() and self.wc_naming_mode_var.get() != "scattered":
            self.wc_naming_mode_var.set("scattered")
        self._wc_refresh_previews()

    def _wc_ingest_paths(self, paths: list[Path]):
        found_docx = self._extract_docx_from_paths(paths)
        if not found_docx:
            messagebox.showinfo("提示", "拖入或选择的项目中未找到任何有效的 .docx 文档。")
            return
        if self._wc_batch_completed and self.wc_files:
            self.wc_files.clear()
            self._wc_batch_completed = False
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
                open_path_in_os(target_dir)

    def _wc_analyze_one(self, file_path: str) -> dict:
        return _call_compatible(
            analyze_docx_word_count,
            docx_path=Path(file_path),
            default_lang_code=self._get_wc_lang_code(),
            keep_existing_lang=self.wc_keep_existing_lang_var.get(),
            date_yyyymmdd=self.wc_date_var.get().strip(),
            keep_existing_date=self.wc_keep_existing_date_var.get(),
            strip_draft_prefix=self.wc_strip_draft_var.get(),
            lang_entries=self.languages,
            naming_mode=self.wc_naming_mode_var.get(),
            scattered_keyword=self.wc_keyword_var.get().strip(),
            update_header_lines=self.wc_update_headers_var.get(),
        )

    def _wc_refresh_previews(self):
        self._save_settings()
        for item in self.wc_tree.get_children():
            self.wc_tree.delete(item)

        sum_up = 0
        sum_rv = 0
        upd_hdr = bool(self.wc_update_headers_var.get())
        for idx, fp in enumerate(self.wc_files):
            p = Path(fp)
            try:
                info = self._wc_analyze_one(fp)
                sum_up += info["up_words"]
                sum_rv += info["rv_words"]
                if info["old_name"] == info["new_name"]:
                    st = "📝 待同步说明行" if upd_hdr else "✅ 已是最新"
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
                clean_ex = _format_exception_concise(ex)
                self.wc_tree.insert(
                    "",
                    tk.END,
                    iid=str(idx),
                    values=(p.name, "-", "-", "-", f"解析失败：{clean_ex}", "❌ 错误"),
                )

        self.wc_summary_var.set(
            f"当前共 {len(self.wc_files)} 份文档  |  第 2 列英文合计：RV {sum_rv} words, UP {sum_up} words"
        )

    def _wc_execute_rename(self):
        self._save_settings()
        if not self.wc_files:
            messagebox.showwarning("提示", "请先点击「选择 Word 文档...」添加需要计算字数的 .docx 文件。")
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
        upd_hdr = bool(self.wc_update_headers_var.get())

        for idx, fp in enumerate(list(self.wc_files)):
            p = Path(fp)
            try:
                info = self._wc_analyze_one(fp)
                sum_up += info["up_words"]
                sum_rv += info["rv_words"]
                final_p, status_txt = _call_compatible(
                    apply_word_count_filename,
                    docx_path=p,
                    new_filename=info["new_name"],
                    copy_to_dir=copy_dir,
                    update_header_lines=upd_hdr,
                )
                if copy_dir is None:
                    self.wc_files[idx] = str(final_p)
                if info.get("lang_code") and str(info["lang_code"]).upper() not in ("AUTO", "XX"):
                    self._record_recent_lang_and_refresh(str(info["lang_code"]), save=False)
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
                clean_ex = _format_exception_concise(ex)
                err_msgs.append(f"• {p.name}：{clean_ex}")
                self.wc_tree.insert(
                    "",
                    tk.END,
                    iid=str(idx),
                    values=(p.name, "-", "-", "-", clean_ex, "❌ 失败"),
                )

        self._wc_batch_completed = True
        self._save_settings()
        self.wc_summary_var.set(
            f"处理完成：成功 {ok_cnt}/{len(self.wc_files)} 份  |  合计：RV {sum_rv} words, UP {sum_up} words"
        )
        if err_msgs:
            messagebox.showwarning(
                "部分文档未能处理",
                f"成功处理 {ok_cnt} 份，失败 {len(err_msgs)} 份：\n\n" + "\n".join(err_msgs[:10]),
            )
        else:
            hdr_note = "\n（已同步更新文档开头高亮说明行）" if upd_hdr else ""
            messagebox.showinfo(
                "字数统计与重命名完成",
                f"已全部处理完成（共 {ok_cnt} 份文档）！\n"
                f"第 2 列英文总计：RV {sum_rv} words, UP {sum_up} words。{hdr_note}",
            )

    # ---------------- System Settings & Shared Methods ----------------

    def _on_settings_key_editor_modified(self, _event=None):
        if self._syncing_settings_editor or self.settings_key_editor is None:
            return
        raw_txt = self.settings_key_editor.get("1.0", tk.END).strip()
        self._pool_multiline_text = raw_txt
        keys = parse_api_keys(raw_txt)
        self._syncing_key_var = True
        try:
            self.api_key_var.set(" | ".join(keys))
        finally:
            self._syncing_key_var = False
        self._refresh_key_pool_badge()

    def _on_yellow_gap_changed(self, _event=None):
        val = self._get_yellow_gap_words()
        self.yellow_gap_words_var.set(str(val))
        self._save_settings()

    def _on_save_settings_clicked(self):
        if self.settings_key_editor is not None and self.settings_key_editor.winfo_exists():
            self._on_settings_key_editor_modified()
        val = self._get_yellow_gap_words()
        self.yellow_gap_words_var.set(str(val))
        self._save_settings()
        key_cnt = len(self._get_all_keys_list())
        ts = datetime.now().strftime("%H:%M:%S")
        msg = f"✅ 全局设置已保存生效（{ts}  |  Key: {key_cnt} 个  |  黄标连片: ≤{val}词）"
        self.settings_save_status_var.set(msg)
        if hasattr(self, "settings_status_lbl"):
            self.settings_status_lbl.configure(foreground="#15803d")
        self._append_log(msg, target_tab="main")

    def _on_clear_settings_keys(self):
        if self._get_all_keys_list() and not messagebox.askyesno(
            "确认清空",
            "确定要清空当前输入的所有 API Key 吗？",
            parent=self.root,
        ):
            return
        self._set_pool_multiline_and_sync_entry("")
        self._save_settings()
        self.settings_save_status_var.set("已清空输入框中的所有 API Key 并保存")
        if hasattr(self, "settings_status_lbl"):
            self.settings_status_lbl.configure(foreground="#475569")

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
        if self.key_entry is not None:
            try:
                sel_first = self.key_entry.index(tk.SEL_FIRST)
                sel_last = self.key_entry.index(tk.SEL_LAST)
                cur_txt = self.api_key_var.get()
                rem_txt = cur_txt[:sel_first] + cur_txt[sel_last:]
                existing_keys = parse_api_keys(rem_txt)
            except tk.TclError:
                existing_keys = parse_api_keys(self.api_key_var.get())
        else:
            existing_keys = parse_api_keys(self.api_key_var.get())

        merged = parse_api_keys(existing_keys + pasted_keys)
        if not existing_keys:
            self._set_pool_multiline_and_sync_entry(clip)
        else:
            self._set_pool_multiline_and_sync_entry("\n".join(merged))
        self._save_settings()
        self._append_log(f"已从剪贴板粘贴多行密钥，当前密钥池共 {len(merged)} 个不重复 Key。", target_tab="main")
        return "break"

    def _open_key_pool_dialog(self):
        if self.settings_key_editor is not None and self.settings_key_editor.winfo_exists():
            self._on_settings_key_editor_modified()
        init_txt = self._pool_multiline_text.strip() or "\n".join(self._get_all_keys_list())

        def _on_saved(new_raw_txt: str, new_rot_label: str):
            self.rotation_label.set(new_rot_label)
            self._set_pool_multiline_and_sync_entry(new_raw_txt)
            self._save_settings()
            cnt = len(self._get_all_keys_list())
            status_txt = f"✅ 密钥池已同步保存：当前共 {cnt} 个 API Key | 通道：{self.api_mode_label.get()} | 策略：{new_rot_label}"
            self.settings_save_status_var.set(status_txt)
            if hasattr(self, "settings_status_lbl"):
                self.settings_status_lbl.configure(foreground="#15803d")
            self._append_log(status_txt, target_tab=self.active_tab_key)

        ApiKeyPoolDialog(
            self.root,
            initial_raw_text=init_txt,
            rotation_label_var=self.rotation_label,
            model_id=self._get_model_id(),
            api_mode=self._get_api_mode_code(),
            on_save_cb=_on_saved,
            api_mode_label_var=self.api_mode_label,
            remember_key_var=self.remember_key_var,
        )

    def _open_lang_manager(self):
        loaded = load_languages()
        fav_set = {c.upper() for c in self.favorite_langs if c}
        for e in loaded:
            if e.get("favorite"):
                fav_set.add(str(e.get("code", "")).upper())
            elif str(e.get("code", "")).upper() in fav_set:
                e["favorite"] = True
        self.languages = loaded

        def _on_saved(updated_list: list, selected_code=None):
            self.languages = [dict(x) for x in updated_list]
            self.favorite_langs = [
                str(e.get("code", "")).strip().upper()
                for e in self.languages
                if e.get("favorite") and str(e.get("code", "")).strip()
            ]
            if selected_code:
                self.recent_langs = record_recent_language(selected_code, self.recent_langs, max_items=5)
            self._refresh_all_lang_combos()
            if selected_code:
                disp = self._code_to_combo_display(selected_code)
                self.lang_var.set(disp)
                self.fa_lang_var.set(disp)
                self.wc_lang_var.set(disp)
                self.fel_lang_var.set(disp)
                self._fel_auto_sync_prefix(force=True)
            self._save_settings()
            self._wc_refresh_previews()
            self._fel_refresh_previews()
            self._append_log(
                f"✅ 语种与标准字体库已更新：共 {len(self.languages)} 个语种（常用置顶 {len(self.favorite_langs)} 个）。",
                target_tab=self.active_tab_key,
            )

        cur_code = "AUTO"
        if self.active_tab_key == "fel":
            cur_code = self._get_fel_lang_code()
        elif self.active_tab_key == "fa":
            cur_code = self._get_fa_lang_code()
        elif self.active_tab_key == "wc":
            cur_code = self._get_wc_lang_code()
        else:
            cur_code = self._get_selected_lang_code()

        LanguageManagerDialog(
            self.root,
            self.languages,
            _on_saved,
            favorite_codes=self.favorite_langs,
            initial_code=cur_code,
        )

    def _clear_log(self):
        self.log_text.delete("1.0", tk.END)

    def _on_clear_all_caches(self):
        if self.is_running or (getattr(self, "ingest_panel", None) is not None and getattr(self.ingest_panel, "_running", False)):
            messagebox.showwarning("任务执行中", "当前正在后台处理文稿，请等待任务完成后再点击「清理缓存」。")
            return

        cleared_items: list[str] = []
        try:
            import library_index as _li_mod
            mem_cnt = len(getattr(_li_mod, "_MEM_INDEX_CACHE", {}))
            if hasattr(_li_mod, "_MEM_INDEX_CACHE"):
                _li_mod._MEM_INDEX_CACHE.clear()
            cleared_items.append(f"总库内存索引 ({mem_cnt} 项)")
        except Exception:
            pass

        try:
            import full_article_processor as _fap_mod
            for cache_attr in ("_MASTER_MEM_CACHE", "_MEM_MASTER_CACHE"):
                if hasattr(_fap_mod, cache_attr) and isinstance(getattr(_fap_mod, cache_attr), dict):
                    getattr(_fap_mod, cache_attr).clear()
        except Exception:
            pass

        try:
            cleanup_stale_ephemeral_workdirs(max_age_seconds=0)
            cleared_items.append("临时工作目录")
        except Exception:
            pass

        # Clear inspect cache & residual completed document lists across tabs to prevent cross-document mixing
        self._fel_inspect_cache.clear()
        had_files = len(self.fel_files) + len(self.dd_files) + len(self.wc_files)
        self._fel_clear_all()
        self.dd_files.clear()
        self._dd_batch_completed = False
        if hasattr(self, "dd_tree"):
            self._dd_refresh_tree()
        self._wc_clear_all()
        self._wc_batch_completed = False
        cleared_items.append(f"各板块文档预览/列表缓存 ({had_files} 份)")

        msg = "🧹 已一键清理：" + "、".join(cleared_items) + "。"
        self._append_log(msg, target_tab=self.active_tab_key)
        messagebox.showinfo("缓存清理完成", msg + "\n\n（API 密钥池与系统语种设置已安全保留）")

    def _toggle_key_visibility(self):
        if self.key_entry is not None:
            self.key_entry.configure(show="" if self.show_key_var.get() else "•")

    def _browse_input(self):
        ps = filedialog.askopenfilenames(
            title="选择待整理 Word 文档（可按住 Ctrl/Shift 多选）",
            filetypes=[("Word 文档", "*.docx")],
        )
        if ps:
            if len(ps) == 1:
                self.input_docx_var.set(ps[0])
                if not self.output_dir_var.get().strip():
                    self.output_dir_var.set(str(Path(ps[0]).parent))
                self._append_log(f"📥 已选择【待整理文档】：{ps[0]}", target_tab="main")
            else:
                self.input_docx_var.set(" ; ".join(ps))
                if not self.output_dir_var.get().strip():
                    self.output_dir_var.set(str(Path(ps[0]).parent / "零散改点整理输出"))
                self._append_log(f"📥 已批量选择 {len(ps)} 份【待整理文档】", target_tab="main")
            self._save_settings()

    def _browse_input_dir(self):
        d = filedialog.askdirectory(title="选择包含待整理 .docx 文档的文件夹")
        if d:
            d_p = Path(d).resolve()
            cnt = len(resolve_sc1_input_docx_list(d_p)) if callable(resolve_sc1_input_docx_list) else len(self._extract_docx_from_paths([d_p]))
            self.input_docx_var.set(str(d_p))
            if not self.output_dir_var.get().strip():
                self.output_dir_var.set(str(d_p.parent / f"{d_p.name}_整理输出"))
            self._save_settings()
            self._append_log(f"📥 已选择【待整理文件夹】：{d_p}（共识别到 {cnt} 份 .docx 文档）", target_tab="main")
            self.stage_var.set(f"📥 已载入待整理文件夹：{d_p.name}（共 {cnt} 份文档）")

    def _browse_library(self):
        d = filedialog.askdirectory(title="选择对应语种《七卷书》总库文件夹")
        if d:
            self.library_dir_var.set(d)
            self._save_settings()
            self._warm_library_index_async(d, target_tab="main")

    def _browse_output(self):
        d = filedialog.askdirectory(title="选择输出保存文件夹")
        if d:
            self.output_dir_var.set(d)
            self._save_settings()

    def _open_output_dir(self):
        d = self.output_dir_var.get().strip()
        p = Path(d).resolve() if d else None
        if p and p.exists() and p.is_dir():
            open_path_in_os(p)
        else:
            messagebox.showinfo("提示", "请先选择真实存在的输出文件夹目录。")

    def _log(self, msg: str, target_tab: str | None = None):
        tab = target_tab or (self._running_tab_key if self.is_running else self.active_tab_key)
        self.log_queue.put(("log", (msg, tab)))

    def _progress(self, pct: float, stage: str):
        self.log_queue.put(("progress", (pct, stage, self._running_tab_key)))

    def _format_pct_timer(self) -> str:
        pct_i = int(round(max(0.0, min(100.0, self.current_pct))))
        if self.task_start_ts > 0:
            elapsed = max(0, int(time.time() - self.task_start_ts))
            mm, ss = divmod(elapsed, 60)
            return f"{pct_i:3d}%  |  已耗时 {mm:02d}:{ss:02d}"
        return f"{pct_i}%"

    def _get_tab_status_widgets(self, tab_key: str):
        mapping = {
            "main": (self.progress_var, self.stage_var, self.stage_lbl, self.pct_timer_var),
            "fa": (self.fa_progress_var, self.fa_stage_var, self.fa_stage_lbl, self.fa_pct_timer_var),
            "dd": (self.dd_progress_var, self.dd_stage_var, self.dd_stage_lbl, self.dd_pct_timer_var),
            "fel": (self.fel_progress_var, self.fel_stage_var, self.fel_stage_lbl, self.fel_pct_timer_var),
        }
        return mapping.get(tab_key, mapping["main"])

    def _tick_elapsed_timer(self):
        if self._closing or not self.root.winfo_exists():
            return
        try:
            if self.is_running and self.task_start_ts > 0:
                fmt = self._format_pct_timer()
                _pvar, _svar, _slbl, tvar = self._get_tab_status_widgets(self._running_tab_key)
                tvar.set(fmt)
        except Exception:
            pass
        finally:
            if not self._closing and self.root.winfo_exists():
                self._timer_after_id = self.root.after(500, self._tick_elapsed_timer)

    def _append_log(self, msg: str, target_tab: str | None = None):
        if not msg:
            return
        safe_msg = redact_secrets(msg, extra_secret=self._get_all_keys_list())
        stripped = safe_msg.lstrip("\n")

        tab_log_map = {
            "main": getattr(self, "log_text", None),
            "fa": getattr(self, "fa_log_text", None),
            "dd": getattr(self, "dd_log_text", None),
            "fel": getattr(self, "fel_log_text", None),
        }
        chosen = tab_log_map.get(target_tab or "main") or getattr(self, "log_text", None)
        widgets = [w for w in [chosen] if w is not None]

        if stripped.startswith("=" * 10):
            for tw in widgets:
                tw.insert(tk.END, "\n" + stripped + "\n", "dim")
                tw.see(tk.END)
            return

        ts = datetime.now().strftime("%H:%M:%S")
        tag = None
        lead = stripped.lstrip()
        if any(k in stripped for k in ("▶", "⚡", "ℹ️", "🔎")):
            tag = "header"
        elif any(k in stripped for k in ("✅", "🎉", "📚", "Successfully aligned", "Final Document Saved")):
            tag = "ok"
        elif lead.startswith("⚠️") or any(k in stripped for k in ("↻", "⏳")):
            tag = "retry"
        elif any(k in stripped for k in ("❌", "ERROR", "失败", "出错", "中止")):
            tag = "err"
        elif any(k in stripped for k in ("⚠️", "提示", "提醒", "跳过")):
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
                        if isinstance(payload, tuple) and len(payload) == 2:
                            m_txt, t_key = payload
                            self._append_log(str(m_txt), target_tab=t_key)
                        else:
                            self._append_log(str(payload))
                    elif kind == "progress":
                        if len(payload) == 3:
                            pct, stage_text, t_key = payload
                        else:
                            pct, stage_text = payload
                            t_key = self._running_tab_key
                        self.current_pct = max(0.0, min(100.0, float(pct)))
                        pvar, svar, _slbl, tvar = self._get_tab_status_widgets(t_key)
                        pvar.set(self.current_pct)
                        disp_stage = (
                            f"⏳ {stage_text}" if self.is_running and self.current_pct < 100 else f"✅ {stage_text}"
                        )
                        svar.set(disp_stage)
                        tvar.set(self._format_pct_timer())
                    elif kind == "done":
                        if len(payload) == 3:
                            ok_flag, final_stage, t_key = payload
                        else:
                            ok_flag, final_stage = payload
                            t_key = self._running_tab_key
                        self._set_running(False, ok=ok_flag, final_stage=final_stage, tab_key=t_key)
                    elif kind == "fa_disambiguate":
                        ambiguous_list, resolved_cnt, holder, done_ev = payload
                        try:
                            dlg = AmbiguousArticleDialog(self.root, ambiguous_list, resolved_count=resolved_cnt)
                            self.root.wait_window(dlg)
                            holder["selected"] = dlg.selected_candidates
                            holder["cancelled"] = dlg.cancelled
                        finally:
                            done_ev.set()
                    elif kind == "dd_mark_done":
                        dup_cnt = int(payload)
                        st_lbl = "✅ 已完成查重" if dup_cnt == 0 else f"✅ 已查重(共 {dup_cnt} 处重复)"
                        smap = {os.path.normcase(str(Path(x))): st_lbl for x in self.dd_files}
                        self._dd_refresh_tree(status_map=smap)
                    elif kind == "record_recent_lang":
                        self._record_recent_lang_and_refresh(str(payload), save=True)
                    elif kind == "settings_status":
                        st_txt, st_color = payload
                        self.settings_save_status_var.set(str(st_txt))
                        if hasattr(self, "settings_status_lbl"):
                            self.settings_status_lbl.configure(foreground=str(st_color))
                    elif kind == "models_refreshed":
                        new_models, status_msg = payload
                        self.dynamic_models = list(new_models)
                        if hasattr(self, "settings_model_cb") and self.settings_model_cb is not None:
                            self.settings_model_cb.configure(values=[m[0] for m in self.dynamic_models])
                        cur_mid = self._get_model_id()
                        matched_disp = next((d for d, m in self.dynamic_models if m == cur_mid), "")
                        if matched_disp:
                            self.model_label.set(matched_disp)
                        elif self.dynamic_models:
                            self.model_label.set(self.dynamic_models[0][0])
                        self._save_settings()
                        self.settings_save_status_var.set(f"✅ {status_msg}")
                        if hasattr(self, "settings_status_lbl"):
                            self.settings_status_lbl.configure(foreground="#15803d")
                    elif kind == "update_available":
                        if isinstance(payload, dict) and payload.get("has_update"):
                            self._show_update_banner(payload)
                    elif kind == "popup_info":
                        title, text = payload
                        messagebox.showinfo(title, redact_secrets(str(text), extra_secret=all_secs))
                    elif kind == "popup_error":
                        title, text = payload
                        messagebox.showerror(title, redact_secrets(str(text), extra_secret=all_secs))
                except Exception as inner_ex:
                    self._append_log(f"❌ 界面队列处理异常：{_format_exception_concise(inner_ex)}")
        finally:
            if not self._closing and self.root.winfo_exists():
                self._poll_after_id = self.root.after(80, self._poll_log_queue)

    def _set_running(self, running: bool, ok: bool = True, final_stage: str = "", tab_key: str = "main"):
        self.is_running = running
        if running:
            self._running_tab_key = tab_key
        elif ok:
            if tab_key == "fel":
                self._fel_batch_completed = True
            elif tab_key == "dd":
                self._dd_batch_completed = True
        state = tk.DISABLED if running else tk.NORMAL
        for b in (
            self.btn_oneclick,
            self.test_btn,
            self.pool_btn,
            self.btn_refresh_input,
            self.btn_refresh_lib,
            self.btn_fa_run,
            self.fa_pool_btn,
            self.btn_fa_refresh_master,
            self.btn_fa_refresh_lib,
            self.btn_dd_run,
            self.btn_fel_run,
            self.fel_pool_btn,
            self.btn_settings_test,
            self.btn_settings_test_all,
        ):
            if b is not None:
                b.configure(state=state)

        pvar, svar, slbl, tvar = self._get_tab_status_widgets(tab_key)
        if running:
            self.task_start_ts = time.time()
            self.current_pct = 2.0
            pvar.set(2.0)
            slbl.configure(foreground="#1d4ed8")
            if final_stage:
                svar.set(f"⏳ {final_stage}")
            tvar.set(self._format_pct_timer())
        else:
            if ok:
                self.current_pct = 100.0
                pvar.set(100.0)
                slbl.configure(foreground="#15803d")
                if final_stage:
                    svar.set(f"✅ {final_stage}")
            else:
                slbl.configure(foreground="#b91c1c")
                if final_stage:
                    svar.set(f"❌ {final_stage}")
            tvar.set(self._format_pct_timer())

    def _on_refresh_models_clicked(self):
        if self.settings_key_editor is not None and self.settings_key_editor.winfo_exists():
            self._on_settings_key_editor_modified()
        self._save_settings()
        keys = self._get_all_keys_list()
        mode = self._get_api_mode_code()
        if not keys:
            messagebox.showwarning("提示", "请先在下方 API Key 池列表中粘贴至少 1 个 Gemini 或 Vertex 收费版 API Key，再加载最新模型列表。")
            return
        if not callable(fetch_available_models):
            messagebox.showerror("组件未就绪", "未找到 fetch_available_models 函数。")
            return

        if hasattr(self, "btn_refresh_models") and self.btn_refresh_models is not None:
            self.btn_refresh_models.configure(state=tk.DISABLED)
        self.settings_save_status_var.set("⏳ 正在从官方接口拉取最新可用模型列表...")
        if hasattr(self, "settings_status_lbl"):
            self.settings_status_lbl.configure(foreground="#1d4ed8")

        def _worker():
            try:
                models_list, msg = fetch_available_models(keys, api_mode=mode)
                self.log_queue.put(("models_refreshed", (models_list, msg)))
                self.log_queue.put(("popup_info", ("模型列表已更新", msg + "\n\n现可在「选择模型」下拉框中自由选择最新模型。")))
            except Exception as ex:
                friendly = _format_exception_concise(ex, keys=keys, api_related=True)
                self.log_queue.put(("settings_status", (f"❌ 加载模型列表失败：{friendly}", "#b91c1c")))
                self.log_queue.put(("popup_error", ("加载模型列表失败", friendly)))
            finally:
                if not self._closing and self.root.winfo_exists():
                    self.root.after(
                        0,
                        lambda: self.btn_refresh_models.configure(state=tk.NORMAL)
                        if getattr(self, "btn_refresh_models", None) is not None
                        else None,
                    )

        threading.Thread(target=_worker, daemon=True).start()

    def _on_test_api(self):
        if self.settings_key_editor is not None and self.settings_key_editor.winfo_exists():
            self._on_settings_key_editor_modified()
        self._save_settings()
        keys = self._get_all_keys_list()
        model = self._get_model_id()
        mode = self._get_api_mode_code()
        if not keys:
            self._switch_tab("settings")
            messagebox.showwarning("提示", "请先在「系统设置」的 API Key 池列表中填入至少 1 个 Gemini 或 Vertex API Key。")
            return

        log_tab = self.active_tab_key if self.active_tab_key in ("main", "fa", "dd", "fel") else "main"
        self.settings_save_status_var.set(f"⏳ 正在测试 {len(keys)} 个 API Key 与模型 {model} ...")
        if hasattr(self, "settings_status_lbl"):
            self.settings_status_lbl.configure(foreground="#1d4ed8")
        self._set_running(True, final_stage=f"正在测试 {len(keys)} 个 API Key 与模型 {model} ...", tab_key=log_tab)
        self._progress(25.0, f"正在验证密钥池（共 {len(keys)} 个 Key）与模型 {model} ...")
        self._log(f"▶ 开始测试 API 密钥池（共 {len(keys)} 个 Key | 模型：{model}）...", target_tab=log_tab)

        def worker():
            try:
                msg = test_api_connection(keys, model, mode, log_fn=lambda m: self._log(m, target_tab=log_tab))
                self._log(f"✅ {msg}", target_tab=log_tab)
                self.log_queue.put(("settings_status", (f"✅ 连接测试通过：{msg}", "#15803d")))
                self.log_queue.put(("done", (True, msg, log_tab)))
                self.log_queue.put(("popup_info", ("连接测试通过", msg)))
            except Exception as ex:
                friendly = _format_exception_concise(ex, keys=keys, api_related=True)
                self._log(f"❌ {friendly}", target_tab=log_tab)
                self.log_queue.put(("settings_status", (f"❌ 连接测试失败：{friendly}", "#b91c1c")))
                self.log_queue.put(("done", (False, friendly, log_tab)))
                self.log_queue.put(("popup_error", ("API 连接失败", friendly)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_run_oneclick(self):
        self._save_settings()
        inp = self.input_docx_var.get().strip()
        lib = self.library_dir_var.get().strip()
        out = self.output_dir_var.get().strip()
        if not inp or not lib or not out:
            messagebox.showwarning("提示", "请先选择【待整理文档/目录】、【对应语种总库目录】和【输出保存目录】。")
            return
        resolved_inp_files = (
            resolve_sc1_input_docx_list(inp)
            if callable(resolve_sc1_input_docx_list)
            else ([Path(inp).resolve()] if Path(inp).is_file() else [])
        )
        if not resolved_inp_files:
            messagebox.showwarning("提示", f"未在指定路径或文件夹中找到待整理的 .docx 文档：\n{inp}")
            return
        if run_oneclick_pipeline is None:
            messagebox.showerror("组件未就绪", "未找到 pipeline.run_oneclick_pipeline，请检查程序文件完整性。")
            return

        sync_master = bool(self.sc1_sync_master_var.get())
        latest_master_path = self.sc1_master_path_var.get().strip() if sync_master else ""
        if sync_master:
            mp = Path(latest_master_path) if latest_master_path else None
            if mp is None or not (mp.is_dir() or (mp.is_file() and mp.suffix.lower() == ".docx")):
                messagebox.showwarning(
                    "提示",
                    "已勾选「核对最新中英底稿」，请选择真实存在的最新中英底稿 .docx 文件或所在文件夹（或取消勾选）。",
                )
                return

        dedup_dir = self.sc1_dedup_dir_var.get().strip() if self.sc1_enable_dedup_var.get() else ""
        if self.sc1_enable_dedup_var.get() and (not dedup_dir or not Path(dedup_dir).is_dir()):
            messagebox.showwarning("提示", "已勾选关联「未入库改点包」查重，请先选择有效的未入库改点包文件夹（或取消勾选）。")
            return

        self._sync_sc1_hl_mode_vars()
        highlight_mode = self.sc1_highlight_mode_var.get() or "all_three"
        highlight_colors = highlight_mode != "revisions_only"
        hl_label = (
            "英外标色 + 中文改点下划线"
            if highlight_mode == "all_three"
            else ("仅英外标色（中文不标下划线）" if highlight_mode == "en_foreign" else "仅输出对比修订")
        )
        export_raw_trilingual = bool(self.sc1_export_raw_var.get())
        yellow_gap_words = self._get_yellow_gap_words()
        want_ai = bool(self.sc1_enable_ai_var.get())
        keys = self._get_all_keys_list()
        model = self._get_model_id()
        mode = self._get_api_mode_code()
        rot_mode = self._get_rotation_mode_code()
        lang = self._get_selected_lang_code()
        if lang and lang != "AUTO":
            self._record_recent_lang_and_refresh(lang, save=True)
        naming_mode = self.sc1_naming_mode_var.get()
        sc_kw = self.sc1_keyword_var.get().strip()
        doc_name = self.doc_name_var.get().strip() if naming_mode == "keep" else ""
        dual = self.dual_outputs_var.get()
        balance_cols = bool(self.sc1_balance_cols_var.get())

        auto_no_key_local = False
        if not highlight_colors:
            effective_enable_ai = False
        elif want_ai and not keys:
            effective_enable_ai = False
            auto_no_key_local = True
        else:
            effective_enable_ai = want_ai

        pool = ApiKeyPool(keys, rotation_mode=rot_mode) if (effective_enable_ai and keys) else ""

        self._set_running(True, final_stage="正在启动一键输出定稿...", tab_key="main")
        self._log("=" * 64, target_tab="main")
        if not highlight_colors:
            mode_summary = f"{hl_label}（不标黄/青高亮，免调 AI）"
        elif effective_enable_ai:
            mode_summary = f"{hl_label} | AI 辅助（模型：{model} | 密钥池：{len(keys)} 个 Key | 黄标连片 ≤{yellow_gap_words}词）"
        else:
            mode_summary = f"{hl_label} | 本地位置估算模式（免 API Key | 黄标连片 ≤{yellow_gap_words}词）"
        name_desc = (
            f"God's words in {sc_kw}"
            if (naming_mode == "scattered" and sc_kw)
            else ("God's words" if naming_mode == "scattered" else (doc_name or "原文档名"))
        )
        master_tag = f" | 核对最新底稿：{Path(latest_master_path).name}" if latest_master_path else ""
        self._log(
            f"⚡ 启动【一键输出定稿】 | {mode_summary} | 语种：{lang} | 命名：{name_desc}{master_tag}",
            target_tab="main",
        )
        if not highlight_colors:
            self._log("ℹ️ 当前为「仅输出对比修订」模式：本次仅保留第 2 列修订痕迹，不输出黄/青高亮颜色", target_tab="main")
        elif auto_no_key_local:
            self._log("ℹ️ 未填写 API Key，本次自动使用「本地位置估算模式」完成标色并导出定稿", target_tab="main")
        elif not want_ai:
            self._log("ℹ️ 已关闭 AI 辅助开关，本次使用「本地位置估算模式」离线完成标色并导出定稿", target_tab="main")

        def worker():
            try:
                res = _call_compatible(
                    run_oneclick_pipeline,
                    input_docx=inp,
                    library_dir=lib,
                    output_dir=out,
                    api_key=pool,
                    model=model,
                    api_mode=mode,
                    rotation_mode=rot_mode,
                    lang=lang,
                    doc_name=doc_name,
                    dual_outputs=dual,
                    enable_ai=effective_enable_ai,
                    naming_mode=naming_mode,
                    scattered_keyword=sc_kw,
                    balance_columns=balance_cols,
                    dedup_pkg_dir=dedup_dir,
                    date_str="",
                    latest_master_path=latest_master_path,
                    highlight_colors=highlight_colors,
                    highlight_mode=highlight_mode,
                    yellow_gap_words=yellow_gap_words,
                    export_raw_trilingual=export_raw_trilingual,
                    log_fn=lambda m: self._log(m, target_tab="main"),
                    progress_fn=self._progress,
                )
                saved_files = list(res.get("saved_files") or [])
                proc_docs = int(res.get("processed_docs_count") or 1)
                total_inp_docs = int(res.get("total_input_docs") or proc_docs)
                failed_docs = list(res.get("failed_docs") or [])
                det_lang = res.get("lang_code") or lang
                if det_lang and str(det_lang).upper() not in ("AUTO", "XX"):
                    self.log_queue.put(("record_recent_lang", str(det_lang)))
                det_disp = find_lang_display_name(det_lang, self.languages)
                skipped_count = int(res.get("skipped_count") or 0)
                used_local_fallback = bool(res.get("used_local_fallback", False))
                up_words = res.get("up_words")
                rv_words = res.get("rv_words")
                dup_cnt = int(res.get("duplicate_count") or res.get("dedup_removed_count") or 0)
                dup_rep = str(res.get("duplicate_report") or res.get("dedup_report") or "")

                files_str = "\n".join(saved_files)
                popup_files_str = (
                    "\n".join(saved_files[:14]) + f"\n...（其余 {len(saved_files) - 14} 份文件已保存在输出目录）"
                    if len(saved_files) > 14
                    else files_str
                )
                extra_notes: list[str] = []
                if failed_docs:
                    extra_notes.append("⚠️ 部分文档处理失败：\n" + "\n".join(failed_docs[:6]))
                if not highlight_colors:
                    extra_notes.append("ℹ️ 输出模式：仅保留第 2 列英文对比修订痕迹（未标记黄/青高亮）")
                elif effective_enable_ai and used_local_fallback:
                    fb_note = (
                        "⚠️ 标色说明：因 API Key 失效、额度受限或网络异常，已自动切换「本地位置估算模式」完成多列标色。"
                    )
                    self._log(fb_note, target_tab="main")
                    extra_notes.append(fb_note)
                elif not effective_enable_ai:
                    extra_notes.append("ℹ️ 标色模式：本地位置估算模式（离线大概位置标色）")

                if skipped_count > 0:
                    skip_note = (
                        f"⚠️ 总库匹配提醒：共有 {skipped_count} 行未在总库匹配到原文，已优雅跳过"
                        "（发译者版与无改点版已自动剔除，完整留底版已红字标注）。"
                    )
                    self._log(skip_note, target_tab="main")
                    extra_notes.append(skip_note)

                if dup_cnt > 0 or dup_rep:
                    dup_note = f"🔍 未入库包查重：已从发译者版自动剔除 {dup_cnt} 处重复段落" + (
                        f"（明细报告：{dup_rep}）" if dup_rep else ""
                    )
                    self._log(dup_note, target_tab="main")
                    extra_notes.append(dup_note)

                wc_line = ""
                if up_words is not None and rv_words is not None and highlight_colors:
                    wc_line = f"\n第 2 列合计统计：RV {rv_words} words, UP {up_words} words" if total_inp_docs > 1 else f"\n第 2 列统计：RV {rv_words} words, UP {up_words} words"

                batch_prefix = f"已整理 {proc_docs}/{total_inp_docs} 份待整理文档，" if total_inp_docs > 1 else ""
                self._log(
                    f"🎉 一键输出定稿完成！（语种：{det_disp}）{batch_prefix}共输出 {len(saved_files)} 份文件：\n{files_str}",
                    target_tab="main",
                )

                stage_flags = []
                if not highlight_colors:
                    stage_flags.append("仅修订痕迹")
                elif effective_enable_ai and used_local_fallback:
                    stage_flags.append("已自动切换本地估算标色")
                elif not effective_enable_ai:
                    stage_flags.append("本地估算标色")
                if skipped_count > 0:
                    stage_flags.append(f"跳过 {skipped_count} 行未匹配")
                if dup_cnt > 0:
                    stage_flags.append(f"剔除 {dup_cnt} 处重复")
                if failed_docs:
                    stage_flags.append(f"失败 {len(failed_docs)} 份")
                flag_suffix = f"；{'，'.join(stage_flags)}" if stage_flags else ""

                self.log_queue.put(
                    ("done", (True, f"一键输出定稿完成！（{det_disp}，{batch_prefix}共 {len(saved_files)} 份文件{flag_suffix}）", "main"))
                )
                notes_block = ("\n\n" + "\n".join(extra_notes)) if extra_notes else ""
                self.log_queue.put(
                    (
                        "popup_info",
                        (
                            "一键输出定稿完成",
                            f"语种：{det_disp}{wc_line}\n{batch_prefix}已成功生成 {len(saved_files)} 份定稿文档：\n\n{popup_files_str}{notes_block}",
                        ),
                    )
                )
            except SystemExit as se:
                msg = f"执行中止（退出码 {se.code}）：请查看运行日志；若目标文档已在 Word 中打开请先关闭后重试。"
                self._log(f"❌ {msg}", target_tab="main")
                self.log_queue.put(("done", (False, msg, "main")))
                self.log_queue.put(("popup_error", ("执行中止", msg)))
            except Exception as ex:
                err_msg = _format_exception_concise(ex, keys=keys)
                self._log(f"❌ 执行出错：{err_msg}", target_tab="main")
                self.log_queue.put(("done", (False, f"执行出错：{err_msg}", "main")))
                self.log_queue.put(("popup_error", ("执行出错", err_msg)))

        threading.Thread(target=worker, daemon=True).start()

    _on_run_all = _on_run_oneclick

    def _on_close(self):
        self._closing = True
        self._save_settings()
        for aid in (self._poll_after_id, self._timer_after_id, self._upd_after_id):
            if aid is not None:
                try:
                    self.root.after_cancel(aid)
                except Exception:
                    pass
        try:
            self.dnd.unhook()
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass

    def _exit_for_update_restart(self):
        self._on_close()
        sys.exit(0)


def _run_frozen_self_test() -> int:
    import docx_xml_utils
    import library_index
    import step1_extract_and_diff  # noqa: F401
    import step1_render_draft  # noqa: F401
    import step2_parse_reviewed  # noqa: F401
    import step2_render_final  # noqa: F401
    import win_dnd  # noqa: F401
    from ai_client import format_friendly_api_error as _ai_friendly_err
    from ai_client import local_align_col3_spans

    if RAW_APP_VERSION != "1.9.7":
        raise RuntimeError(f"RAW_APP_VERSION expected '1.9.7', got {RAW_APP_VERSION!r}")

    is_frozen = bool(getattr(sys, "frozen", False))
    if is_frozen and hasattr(sys, "_MEIPASS"):
        bundled_cfg = Path(sys._MEIPASS) / "resources" / "config.json"
        if not bundled_cfg.is_file():
            raise FileNotFoundError(f"Bundled config.json missing inside EXE: {bundled_cfg}")

    if not callable(getattr(library_index, "ensure_library_index", None)):
        raise RuntimeError("library_index.ensure_library_index missing in self-test")

    cfg = docx_xml_utils.load_config()
    if cfg.get("col2_up_highlight") != "yellow" or not cfg.get("language_code_map"):
        raise RuntimeError(f"Invalid config loaded in self-test: {cfg}")
    langs = load_languages()
    if not langs:
        raise RuntimeError("Empty language registry in self-test")

    # Verify v1.8.0 favorite + recent language prioritization & code resolution
    prio_vals = get_prioritized_combo_values(langs, recent_codes=["BG", "PL"], favorite_codes=["PL"])
    if (
        len(prio_vals) < 4
        or prio_vals[0] != AUTO_COMBO_LABEL
        or not prio_vals[1].startswith("★ PL - ")
        or not prio_vals[2].startswith("⏱ BG - ")
    ):
        raise RuntimeError(f"get_prioritized_combo_values self-test failed: {prio_vals[:5]!r}")
    if _resolve_combo_code(prio_vals[1], langs) != "PL" or _resolve_combo_code(prio_vals[2], langs) != "BG":
        raise RuntimeError("Prioritized combo code resolution self-test failed")
    rec_updated = record_recent_language("TH", ["BG", "TH", "PL"], max_items=5)
    if rec_updated[:3] != ["TH", "BG", "PL"] or record_recent_language("AUTO", rec_updated, max_items=5) != rec_updated:
        raise RuntimeError(f"record_recent_language self-test failed: {rec_updated!r}")

    pool = ApiKeyPool(["self_test_placeholder_1", "self_test_placeholder_2"])
    if len(pool) != 2:
        raise RuntimeError("ApiKeyPool self-test failed")

    # Verify format_friendly_api_error strips raw JSON & produces concise Chinese message
    friendly_sample = _ai_friendly_err(
        'HTTP 401: {"error": {"code": 401, "message": "API key not valid"}}',
        total_keys=3,
    )
    if not isinstance(friendly_sample, str) or not friendly_sample.strip() or "\n" in friendly_sample or "{" in friendly_sample:
        raise RuntimeError(f"format_friendly_api_error self-test failed: {friendly_sample!r}")

    # Verify local_align_col3_spans offline approximate Col 3 alignment
    sample_old_en = "First sentence remains untouched. Second sentence contains a local phrase update here."
    sample_col3 = "Première phrase reste intacte. Deuxième phrase contient une mise à jour locale ici."
    phrase = "local phrase update"
    c_start = sample_old_en.index(phrase)
    c_end = c_start + len(phrase)
    local_res = local_align_col3_spans(
        sample_col3,
        sample_old_en,
        [
            {
                "span_id": "r0_s0",
                "type": "up",
                "old_en_in_col3": phrase,
                "old_en_char_start": c_start,
                "old_en_char_end": c_end,
                "left_old_context": sample_old_en[:c_start],
                "right_old_context": sample_old_en[c_end:],
            }
        ],
        docx_xml_utils.find_fuzzy_substring_span,
    )
    aligned_sub = local_res.get("r0_s0", "") if isinstance(local_res, dict) else ""
    if not aligned_sub or docx_xml_utils.find_fuzzy_substring_span(sample_col3, aligned_sub) is None:
        raise RuntimeError(f"local_align_col3_spans self-test failed: {local_res!r}")

    # Verify word_counter filename parser & formatter (RV first, UP second, cleaning `( words)` & `_完整版`)
    p1 = parse_doc_filename("GR_God’s Words_(UP 10 words, RV 20 words)_20261114", known_codes={"GR"})
    if p1["orig_name"] != "God’s Words" or p1["existing_lang"] != "GR" or not p1["had_word_count"]:
        raise RuntimeError(f"Filename parser self-test failed: {p1}")
    p2 = parse_doc_filename(
        "UA_3-4-09 100 Prepare Sufficient Good Deeds for Your Destination_( words)_20261025_完整版_(UP 440 words, RV 708 words)_20261118",
        known_codes={"UA"},
    )
    if p2["orig_name"] != "3-4-09 100 Prepare Sufficient Good Deeds for Your Destination" or p2["existing_lang"] != "UA":
        raise RuntimeError(f"Legacy suffix cleanup self-test failed: {p2}")
    fn_both = format_counted_docx_name("GR", "God’s Words", 128, 64, "20261115")
    if fn_both != "GR_God’s Words_(RV 64 words, UP 128 words)_20261115.docx":
        raise RuntimeError(f"Filename formatter (both>0, RV first) self-test failed: {fn_both}")

    # Determine whether full v1.7.0 / v1.8.0 sibling modules are present (always True in frozen EXE & merged repo)
    base_d = _app_base_dir()
    require_v17_siblings = is_frozen or (base_d / "full_article_processor.py").is_file()
    require_v18_siblings = is_frozen or (base_d / "master_sync.py").is_file()

    if require_v17_siblings:
        fn_up_only = format_counted_docx_name("GR", "God’s Words", 96, 0, "20261115")
        if fn_up_only != "GR_God’s Words_(UP 96 words)_20261115.docx":
            raise RuntimeError(f"0-word RV omission self-test failed: {fn_up_only}")
        fn_rv_only = format_counted_docx_name("GR", "God’s Words", 0, 55, "20261115")
        if fn_rv_only != "GR_God’s Words_(RV 55 words)_20261115.docx":
            raise RuntimeError(f"0-word UP omission self-test failed: {fn_rv_only}")
        fn_zero = format_counted_docx_name("GR", "God’s Words", 0, 0, "20261115")
        if fn_zero != "GR_God’s Words_20261115.docx":
            raise RuntimeError(f"0-word both-zero omission self-test failed: {fn_zero}")

        import full_article_processor as _fap_test
        import dedup_checker as _dd_test
        import updater as _upd_test
        import sheets_reader as _sr_test
        import library_ingester as _li_test
        import proofread_ingester as _pi_test
        import ingest_tab_ui as _iui_test

        if not callable(getattr(_sr_test, "parse_google_sheet_url", None)):
            raise RuntimeError("sheets_reader.parse_google_sheet_url missing in self-test")
        if not callable(getattr(_li_test, "run_library_ingestion_pipeline", None)):
            raise RuntimeError("library_ingester.run_library_ingestion_pipeline missing in self-test")
        if not callable(getattr(_pi_test, "run_proofread_ingestion_pipeline", None)):
            raise RuntimeError("proofread_ingester.run_proofread_ingestion_pipeline missing in self-test")
        if getattr(_iui_test, "IngestTabPanel", None) is None:
            raise RuntimeError("ingest_tab_ui.IngestTabPanel missing in self-test")

        # Self-test Scenario 3A comment attachment (`MarkupIdAllocator.next_id()`) & v1.9.3 core helpers
        _test_doc = ET.Element(f"{docx_xml_utils.W}document")
        _test_body = ET.SubElement(_test_doc, f"{docx_xml_utils.W}body")
        _test_tbl = ET.SubElement(_test_body, f"{docx_xml_utils.W}tbl")
        _test_tr = ET.SubElement(_test_tbl, f"{docx_xml_utils.W}tr")
        for _sample_txt in ("神的话语", "God's words", "Słowa Boga"):
            _c = ET.SubElement(_test_tr, f"{docx_xml_utils.W}tc")
            _p = ET.SubElement(_c, f"{docx_xml_utils.W}p")
            _p.append(docx_xml_utils.make_run(_sample_txt))
        _test_tc = _test_tr.findall(f"{docx_xml_utils.W}tc")[0]
        _test_cmts = docx_xml_utils.new_comments_root()
        _test_alloc = docx_xml_utils.CommentIdAllocator(doc_root=_test_doc, comments_root=_test_cmts)
        if _test_alloc.allocate() < 1:
            raise RuntimeError("MarkupIdAllocator.allocate() backward-compat alias failed")
        _li_test.add_comment_to_cell(_test_tc, "测试验收批注", _test_cmts, _test_alloc)
        if len(_test_cmts.findall(f"{docx_xml_utils.W}comment")) != 1:
            raise RuntimeError("library_ingester.add_comment_to_cell self-test failed")
        if docx_xml_utils.balance_trilingual_table_widths(_test_doc) != 1:
            raise RuntimeError("balance_trilingual_table_widths self-test failed")
        docx_xml_utils.apply_highlights_to_col3_cell(_test_tc, [(0, 2, "u:single")])
        if _test_tc.find(f".//{docx_xml_utils.W}u") is None or _test_tc.find(f".//{docx_xml_utils.W}highlight") is not None:
            raise RuntimeError("Col 1 u:single underline self-test failed")
        # Self-test v1.9.4 universal OOXML sanitizer & fellowship comment verification
        _hl_bad = ET.SubElement(_test_tc.find(f"{docx_xml_utils.W}p"), f"{docx_xml_utils.W}hyperlink", {f"{{{docx_xml_utils.R_URI}}}id": "rId_missing"})
        _hl_bad.append(docx_xml_utils.make_run("bad_link"))
        docx_xml_utils.sanitize_document_ooxml(_test_doc, _test_cmts, valid_hyperlink_rids=set(), declared_ns_prefixes={"w", "r"})
        if _test_doc.find(f".//{docx_xml_utils.W}hyperlink") is not None:
            raise RuntimeError("sanitize_document_ooxml did not unwrap broken hyperlink")
        import fellowship_processor as _fp_test
        if not callable(getattr(_fp_test, "verify_and_retain_fellowship_comments", None)):
            raise RuntimeError("fellowship_processor.verify_and_retain_fellowship_comments missing in self-test")
        if not callable(getattr(_li_test, "analyze_sent_col2_vs_library", None)) or not callable(
            getattr(_li_test, "evaluate_proofread_sheet_rows", None)
        ):
            raise RuntimeError("library_ingester 3A/proofread validators missing in self-test")

        for req_fn in ("parse_titles_text_or_file", "parse_titles_multiline", "resolve_full_articles", "run_full_article_pipeline"):
            if not callable(getattr(_fap_test, req_fn, None)):
                raise RuntimeError(f"full_article_processor.{req_fn} missing in self-test")
        parsed_t = _fap_test.parse_titles_multiline("第一篇\n\n1-01 追求真理的重要性与路途\n")
        if len(parsed_t) != 2 or "1-01 追求真理的重要性与路途" not in parsed_t[1]:
            raise RuntimeError(f"parse_titles_multiline self-test failed: {parsed_t!r}")

        if not callable(getattr(_dd_test, "check_and_export_duplicates", None)):
            raise RuntimeError("dedup_checker.check_and_export_duplicates missing in self-test")

        if not callable(resolve_sc1_input_docx_list):
            raise RuntimeError("pipeline.resolve_sc1_input_docx_list missing in self-test")

        upd_ver = str(getattr(_upd_test, "APP_VERSION", "")).lstrip("vV")
        expected_vers = ("1.9.7",) if require_v18_siblings else ("1.7.0", "1.8.0", "1.8.1", "1.9.0", "1.9.1", "1.9.2", "1.9.3", "1.9.4", "1.9.5", "1.9.6", "1.9.7")
        if upd_ver not in expected_vers:
            raise RuntimeError(f"updater.APP_VERSION mismatch: {getattr(_upd_test, 'APP_VERSION', None)!r}")
        plat_fn = getattr(_upd_test, "get_platform_zip_asset_name", None)
        if not callable(plat_fn):
            raise RuntimeError("updater.get_platform_zip_asset_name missing in self-test")
        if (
            plat_fn("win32", "AMD64") != "ScatteredRevisionTool-Windows-x64.zip"
            or plat_fn("darwin", "arm64") != "ScatteredRevisionTool-macOS-arm64.zip"
            or plat_fn("darwin", "x86_64") != "ScatteredRevisionTool-macOS-x64.zip"
        ):
            raise RuntimeError("updater.get_platform_zip_asset_name platform matrix failed")
        if not callable(getattr(_upd_test, "check_for_updates", None)) or not callable(
            getattr(_upd_test, "download_and_apply_update", None)
        ):
            raise RuntimeError("updater functions missing in self-test")
        clean_env_fn = getattr(_upd_test, "_build_clean_relaunch_env", None)
        if not callable(clean_env_fn):
            raise RuntimeError("updater._build_clean_relaunch_env missing in self-test")
        sample_clean_env = clean_env_fn()
        if sample_clean_env.get("PYINSTALLER_RESET_ENVIRONMENT") != "1" or any(
            k.upper().startswith(("_PYI_", "_MEI")) for k in sample_clean_env
        ):
            raise RuntimeError(f"_build_clean_relaunch_env did not sanitize _PYI_/_MEI vars: {sample_clean_env!r}")

    if require_v18_siblings:
        import master_sync as _ms_test
        import full_article_processor as _fap_v18

        if not callable(getattr(_ms_test, "sync_pending_with_latest_master", None)):
            raise RuntimeError("master_sync.sync_pending_with_latest_master missing in self-test")
        for req_v18_fn in ("refresh_master_index", "resolve_direct_folder_articles"):
            if not callable(getattr(_fap_v18, req_v18_fn, None)):
                raise RuntimeError(f"full_article_processor.{req_v18_fn} missing in self-test")
        if is_frozen:
            import ai_client as _ai_v18

            if not callable(getattr(_lang_reg_mod, "get_prioritized_combo_values", None)):
                raise RuntimeError("lang_registry.get_prioritized_combo_values missing in frozen self-test")
            if not callable(getattr(_ai_v18, "local_align_col1_spans", None)):
                raise RuntimeError("ai_client.local_align_col1_spans missing in frozen self-test")

    if require_v17_siblings or callable(get_lang_typography):
        from lang_registry import get_lang_typography as _get_typo

        th_typo = _get_typo("TH")
        if (
            not isinstance(th_typo, dict)
            or th_typo.get("font") != "Angsana New"
            or (th_typo.get("title_sz"), th_typo.get("body_sz"), th_typo.get("citation_sz")) != (36, 32, 28)
        ):
            raise RuntimeError(f"get_lang_typography('TH') self-test failed: {th_typo!r}")

    # Verify Scenario 4 / Part-Merge fellowship_processor self-test
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
    if len(sys.argv) >= 3 and sys.argv[1] == "--finish-update":
        if callable(run_finish_update_mode):
            sys.exit(run_finish_update_mode(sys.argv[2]))
        sys.exit(1)
    if callable(cleanup_stale_update_backups):
        try:
            cleanup_stale_update_backups(wait_first=1.5 if "--cleanup-helper" in sys.argv else 0.0)
        except Exception:
            pass
    _enable_high_dpi()
    root = create_tk_root()
    ScatteredRevisionApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
