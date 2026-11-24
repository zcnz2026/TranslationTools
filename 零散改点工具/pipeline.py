from concurrent.futures import ThreadPoolExecutor
import copy
import importlib
import inspect
import io
import json
import os
import re
import shutil
import sys
import tempfile
import time
from contextlib import redirect_stdout
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple, Union
from unittest.mock import patch
import xml.etree.ElementTree as ET


def _ensure_skill_scripts_on_path():
    candidates = []
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        candidates.append(Path(sys._MEIPASS) / "skill_scripts")
        candidates.append(Path(sys._MEIPASS))
    here = Path(__file__).resolve().parent
    candidates.append(here.parent / ".agents" / "skills" / "scattered-revision" / "scripts")
    candidates.append(here / "skill_scripts")
    for p in candidates:
        if p.exists() and str(p) not in sys.path:
            sys.path.insert(0, str(p))


_ensure_skill_scripts_on_path()

import docx_xml_utils  # noqa: E402
import step1_extract_and_diff  # noqa: E402
import step1_render_draft  # noqa: E402
import step2_parse_reviewed  # noqa: E402
import step2_render_final  # noqa: E402
from ai_client import ApiKeyPool, align_step2_spans, evaluate_step1_decisions  # noqa: E402
from dedup_checker import remove_unmerged_duplicates_from_docx  # noqa: E402
from lang_registry import find_lang_display_name, get_languages_file, load_languages  # noqa: E402
from master_sync import _accept_cell_revisions_inplace, sync_pending_with_latest_master  # noqa: E402
from security_utils import (  # noqa: E402
    create_ephemeral_workdir,
    get_secure_user_workdir,
    parse_api_keys,
)
from word_counter import (  # noqa: E402
    balance_trilingual_table_widths,
    ensure_three_column_tables_and_landscape,
    format_counted_docx_name,
    is_top_highlight_instruction_text,
    parse_doc_filename,
    resolve_doc_naming_label,
    update_top_highlight_instruction_lines,
)

# Ensure SCATTERED_REV_LANGUAGES_JSON env var points to external languages.json
get_languages_file()

W = docx_xml_utils.W
NS = docx_xml_utils.NS

_OUTPUT_COUNT_TAG_RE = re.compile(
    r"_\((?:(?:UP|RV|PR)\s+\d+\s+words(?:,\s*(?:UP|RV|PR)\s+\d+\s+words)?)\)_",
    re.IGNORECASE,
)
_WORD_COUNTS_STDOUT_RE = re.compile(r"Final Col 2 Word Counts:\s*UP\s*=\s*(\d+)\s*words,\s*RV\s*=\s*(\d+)\s*words")
_CLI_FLAG_SUPPORT_CACHE: Dict[Tuple[str, str], bool] = {}


def _script_supports_cli_flag(mod, flag: str, fallback_utils_attr: str = "") -> bool:
    key = (getattr(mod, "__name__", str(mod)), flag)
    if key in _CLI_FLAG_SUPPORT_CACHE:
        return _CLI_FLAG_SUPPORT_CACHE[key]

    supported = False
    mod_file = getattr(mod, "__file__", "")
    if mod_file:
        pf = Path(mod_file)
        if pf.is_file() and pf.suffix.lower() == ".py":
            try:
                supported = flag in pf.read_text(encoding="utf-8", errors="ignore")
                _CLI_FLAG_SUPPORT_CACHE[key] = supported
                return supported
            except OSError:
                pass
    if fallback_utils_attr:
        supported = callable(getattr(docx_xml_utils, fallback_utils_attr, None))
    _CLI_FLAG_SUPPORT_CACHE[key] = supported
    return supported


def ensure_library_index(
    library_dir: Union[str, Path],
    log_fn: Optional[Callable[[str], None]] = None,
    max_workers: int = 8,
    force_rebuild: bool = False,
) -> dict:
    """
    Build or refresh the persistent multi-threaded library index via `library_index.ensure_library_index`,
    verifying live `os.stat` `(mtime_ns, size)` on every run and forwarding `force_rebuild`.
    """
    _ensure_skill_scripts_on_path()
    try:
        lib_idx_mod = importlib.import_module("library_index")
        fn = getattr(lib_idx_mod, "ensure_library_index", None)
        if callable(fn):
            sig = inspect.signature(fn)
            if "force_rebuild" in sig.parameters:
                return fn(
                    library_dir,
                    log_fn=log_fn,
                    max_workers=max_workers,
                    force_rebuild=bool(force_rebuild),
                )
            if force_rebuild:
                resolved_lib = Path(library_dir).resolve()
                mem_cache = getattr(lib_idx_mod, "_MEM_INDEX_CACHE", None)
                if isinstance(mem_cache, dict):
                    mem_cache.pop(os.path.normcase(str(resolved_lib)), None)
                res_cache_fn = getattr(lib_idx_mod, "_resolve_cache_path", None)
                if callable(res_cache_fn):
                    try:
                        Path(res_cache_fn(resolved_lib)).unlink(missing_ok=True)
                    except OSError:
                        pass
            return fn(library_dir, log_fn=log_fn, max_workers=max_workers)
    except ModuleNotFoundError:
        pass

    # Standalone fallback (used only when testing in isolation without library_index.py)
    t0 = time.perf_counter()
    lib_p = Path(library_dir).resolve()
    if not lib_p.exists():
        raise FileNotFoundError(f"找不到语种总库文件夹：{lib_p}")

    cfg = docx_xml_utils.load_config()
    excl_kws = cfg.get("exclude_library_dir_keywords", ["标记结果", "初筛", "答案", "output", "temp"])
    docx_files: List[Path] = []
    for p in lib_p.rglob("*.docx"):
        if p.name.startswith("~$") or p.name.startswith("初筛_") or _OUTPUT_COUNT_TAG_RE.search(p.name):
            continue
        rel_str = str(p.relative_to(lib_p))
        if any(kw in rel_str for kw in excl_kws):
            continue
        try:
            _ = os.stat(p)
        except OSError:
            continue
        docx_files.append(p)

    def _count_rows(fp: Path) -> int:
        try:
            root, _, _ = docx_xml_utils.load_docx_xml(fp)
            cnt = 0
            for tr in root.findall(".//w:tbl/w:tr", docx_xml_utils.NS):
                tcs = tr.findall("w:tc", docx_xml_utils.NS)
                if len(tcs) >= 3 and docx_xml_utils.get_cell_plain_text(tcs[2]).strip():
                    cnt += 1
            return cnt
        except Exception:
            return 0

    total_rows = 0
    if docx_files:
        workers = max(1, min(int(max_workers or 8), len(docx_files)))
        with ThreadPoolExecutor(max_workers=workers) as ex:
            total_rows = sum(ex.map(_count_rows, docx_files))

    return {
        "total_files": len(docx_files),
        "total_rows": total_rows,
        "rebuilt_files": len(docx_files),
        "cache_path": "",
        "elapsed_sec": round(time.perf_counter() - t0, 3),
    }


class _LiveLogWriter(io.TextIOBase):
    """Captures stdout while streaming each completed line immediately to log_fn (safe even when log_fn=print)."""

    def __init__(self, log_fn: Optional[Callable[[str], None]] = None):
        super().__init__()
        self._log_fn = log_fn
        self._orig_stdout = sys.stdout
        self._in_emit = False
        self._full = []
        self._buf = ""

    def _emit_line(self, text: str) -> None:
        if not self._log_fn or self._in_emit:
            return
        self._in_emit = True
        cur_out = sys.stdout
        try:
            sys.stdout = self._orig_stdout
            self._log_fn(text)
        finally:
            sys.stdout = cur_out
            self._in_emit = False

    def write(self, s: str) -> int:
        if not s:
            return 0
        self._full.append(s)
        if self._log_fn and not self._in_emit:
            self._buf += s
            while "\n" in self._buf:
                line, self._buf = self._buf.split("\n", 1)
                line = line.rstrip("\r")
                if line.strip():
                    self._emit_line(f"  {line}")
        return len(s)

    def flush(self):
        if self._log_fn and not self._in_emit and self._buf.strip():
            rem = self._buf.rstrip()
            self._buf = ""
            self._emit_line(f"  {rem}")

    def getvalue(self) -> str:
        self.flush()
        return "".join(self._full)


def _run_module_main(mod_main, argv: list, log_fn: Optional[Callable[[str], None]] = None) -> str:
    writer = _LiveLogWriter(log_fn=log_fn)
    with patch.object(sys, "argv", argv), redirect_stdout(writer):
        mod_main()
    return writer.getvalue()


def _get_work_dir(stem: str) -> Path:
    return create_ephemeral_workdir(stem)


def clean_doc_label(stem: str) -> str:
    known_codes = {e["code"].upper() for e in load_languages() if e.get("code")}
    parsed = parse_doc_filename(stem, known_codes=known_codes, strip_draft_prefix=True)
    return docx_xml_utils.sanitize_filename_part(parsed["orig_name"], fallback="Document", max_len=140)


def _resolve_pipeline_doc_label(
    explicit_doc_name: str,
    naming_mode: str,
    scattered_keyword: str,
    fallback_stem: str,
    source_docx_path: Optional[Path] = None,
) -> str:
    exp = (explicit_doc_name or "").strip()
    if exp:
        return docx_xml_utils.sanitize_filename_part(exp, fallback="Document", max_len=140)
    fallback_lbl = clean_doc_label(fallback_stem)
    mode = (naming_mode or "keep").strip().lower()
    if mode in ("scattered", "full"):
        resolved = resolve_doc_naming_label(
            naming_mode=mode,
            fallback_orig_name=fallback_lbl,
            scattered_keyword=scattered_keyword,
            docx_path=source_docx_path,
        )
        return docx_xml_utils.sanitize_filename_part(resolved, fallback=fallback_lbl, max_len=140)
    return fallback_lbl


def _prepare_optional_pool(
    api_key: Union[str, List[str], ApiKeyPool],
    rotation_mode: str,
    enable_ai: bool,
) -> Tuple[Union[List[str], ApiKeyPool], str, bool, bool]:
    """Build an ApiKeyPool without raising ValueError when `api_key` is empty."""
    if isinstance(api_key, ApiKeyPool):
        has_keys = len(api_key) > 0
        eff_rot = api_key.rotation_mode
        return api_key, eff_rot, has_keys, bool(enable_ai and has_keys)

    keys = parse_api_keys(api_key)
    has_keys = len(keys) > 0
    eff_rot = rotation_mode if rotation_mode in ("round_robin", "failover") else "round_robin"
    if has_keys and enable_ai:
        return ApiKeyPool(keys, rotation_mode=eff_rot), eff_rot, True, True
    return keys, eff_rot, has_keys, False


def resolve_highlight_mode(
    highlight_mode: Optional[str] = None,
    highlight_colors: bool = True,
) -> Tuple[str, bool, bool]:
    """
    Normalize 3-way highlight mode into `(mode_id, use_highlights, highlight_col1)`:
      - `"all_three"` (default): Col 1 Chinese single underline (`u:single`) + Col 2 English yellow/cyan
        + Col 3 Foreign soft light-blue shading (`#D9EEF7`) / `lightGray`.
      - `"en_foreign"`: Col 2 English yellow/cyan + Col 3 Foreign `#D9EEF7` / `lightGray`; NO Col 1 mark.
      - `"revisions_only"`: NO highlights/shading/underline on any column, only Col 2 track changes (0 AI calls).
    """
    raw = (highlight_mode or "").strip().lower()
    if raw in ("revisions_only", "none", "off", "false", "0"):
        return "revisions_only", False, False
    if raw in ("en_foreign", "en_col3", "no_col1"):
        return "en_foreign", True, False
    if raw in ("all_three", "all", "three", "cn_en_foreign", "col1_underline", "default"):
        return "all_three", True, True
    if not highlight_colors:
        return "revisions_only", False, False
    return "all_three", True, True


def _strip_all_highlights_in_root(root: ET.Element) -> None:
    """Remove all `<w:highlight>` & UP `#D9EEF7` `<w:shd>` tags across `root` and strip top instruction lines."""
    for rPr in root.findall(".//w:rPr", NS):
        docx_xml_utils.remove_rpr_child(rPr, "highlight")
        shd_el = rPr.find("w:shd", NS)
        if shd_el is not None and (shd_el.attrib.get(f"{W}fill") or "").upper() == "D9EEF7":
            docx_xml_utils.remove_rpr_child(rPr, "shd")
    update_top_highlight_instruction_lines(root, has_up=False, has_rv=False)


def _clone_clean_col12_paragraphs(src_tc: ET.Element, target_tc: ET.Element) -> List[ET.Element]:
    """
    Clone clean paragraphs from matched library Col 1 (`s_tcs[0]`) or Col 2 (`s_tcs[1]`)
    without shading, highlights, or track changes for pre-diff raw trilingual export.
    """
    work_tc = copy.deepcopy(src_tc)
    _accept_cell_revisions_inplace(work_tc)
    tpl_pPr = docx_xml_utils.get_cell_template_ppr(target_tc, fallback_tc=work_tc)
    tpl_rPr = docx_xml_utils.get_first_run_rpr(target_tc, fallback_tc=work_tc)
    for non_base in ("color", "b", "bCs", "i", "iCs", "u", "highlight", "shd"):
        docx_xml_utils.remove_rpr_child(tpl_rPr, non_base)

    out_paras: List[ET.Element] = []
    for p in work_tc.findall("w:p", NS):
        new_p = ET.Element(f"{W}p")
        if tpl_pPr is not None:
            new_p.append(copy.deepcopy(tpl_pPr))
        for r in p.findall(".//w:r", NS):
            txt = "".join(t.text or "" for t in r.findall("w:t", NS))
            if not txt:
                continue
            r_rPr = copy.deepcopy(tpl_rPr)
            old_rPr = r.find("w:rPr", NS)
            if old_rPr is not None:
                for inline_tag in ("color", "b", "bCs", "i", "iCs", "u"):
                    ch_el = old_rPr.find(f"w:{inline_tag}", NS)
                    if ch_el is not None:
                        cp_el = copy.deepcopy(ch_el)
                        docx_xml_utils._strip_unsafe_attrs_recursive(cp_el)
                        docx_xml_utils.insert_rpr_child_ordered(r_rPr, cp_el)
            new_p.append(docx_xml_utils.make_run(txt, base_rPr=r_rPr, highlight=None, shd_fill=None, is_del=False))
        out_paras.append(new_p)

    if not out_paras:
        empty_p = ET.Element(f"{W}p")
        if tpl_pPr is not None:
            empty_p.append(copy.deepcopy(tpl_pPr))
        out_paras.append(empty_p)
    return out_paras


def _build_raw_trilingual_docx_from_manifest(
    manifest_path: Path,
    output_path: Path,
    balance_columns: bool = True,
) -> Path:
    """
    Fallback builder for pre-comparison extracted 3-language raw `.docx` (`对比前提取的三语原稿`):
      - Col 1 = matched library Col 1 (`s_tcs[0]` clean)
      - Col 2 = matched library old English (`s_tcs[1]` clean)
      - Col 3 = matched library Foreign (`s_tcs[2]` with preserved formatting & migrated Word comments)
    """
    with open(manifest_path, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    inp_path = Path(manifest["input_docx"]).resolve()
    root, _, ns_pairs = docx_xml_utils.load_docx_xml(inp_path)
    ensure_three_column_tables_and_landscape(root)
    orig_comments = docx_xml_utils.load_docx_comments_xml(inp_path)
    inp_had_comments = orig_comments is not None
    working_comments = orig_comments if orig_comments is not None else docx_xml_utils.new_comments_root()
    id_allocator = docx_xml_utils.CommentIdAllocator(doc_root=root, comments_root=working_comments)
    body = root.find("w:body", NS)

    src_cache: Dict[Path, tuple] = {}
    for rm in manifest.get("rows", []):
        body_el = body[rm["body_idx"]]
        tr = body_el.findall("w:tr", NS)[rm["row_idx"]]
        tcs = tr.findall("w:tc", NS)
        if len(tcs) < 3:
            continue

        src_file = Path(rm["source_file"])
        if src_file not in src_cache:
            s_root, _, _ = docx_xml_utils.load_docx_xml(src_file)
            s_cmap = docx_xml_utils.load_docx_comments_map(src_file)
            s_styles = docx_xml_utils.load_docx_styles_xml(src_file)
            src_cache[src_file] = (s_root.findall(".//w:tbl/w:tr", NS), s_cmap, s_styles)
        s_trs, s_cmap, s_styles = src_cache[src_file]
        s_tr = s_trs[rm["source_row_idx"]]
        s_tcs = s_tr.findall("w:tc", NS)

        step1_render_draft._populate_clean_library_cell(
            src_tc=s_tcs[0],
            target_tc=tcs[0],
            fallback_tc=tcs[0],
            src_comments_map=s_cmap,
            target_comments_root=working_comments,
            id_allocator=id_allocator,
            src_styles_root=s_styles,
            ensure_cjk_songti=True,
            preserve_shading_and_highlights=True,
        )
        step1_render_draft._populate_clean_library_cell(
            src_tc=s_tcs[1],
            target_tc=tcs[1],
            fallback_tc=tcs[1],
            src_comments_map=s_cmap,
            target_comments_root=working_comments,
            id_allocator=id_allocator,
            src_styles_root=s_styles,
            ensure_cjk_songti=False,
            preserve_shading_and_highlights=True,
        )
        step1_render_draft._populate_clean_library_cell(
            src_tc=s_tcs[2],
            target_tc=tcs[2],
            fallback_tc=tcs[1],
            src_comments_map=s_cmap,
            target_comments_root=working_comments,
            id_allocator=id_allocator,
            src_styles_root=s_styles,
            ensure_cjk_songti=False,
            preserve_shading_and_highlights=True,
        )

    for sr in manifest.get("skipped_rows", []):
        body_el = body[sr["body_idx"]]
        trs = body_el.findall("w:tr", NS)
        if 0 <= sr["row_idx"] < len(trs):
            tcs = trs[sr["row_idx"]].findall("w:tc", NS)
            if len(tcs) >= 3:
                _strip_all_highlights_in_root(trs[sr["row_idx"]])
                step1_render_draft.populate_skipped_col3_cell(tcs[2], fallback_tc=tcs[1])

    update_top_highlight_instruction_lines(root, has_up=False, has_rv=False)
    if balance_columns:
        balance_trilingual_table_widths(root)

    final_comments = (
        working_comments
        if (inp_had_comments or len(working_comments.findall(f"{W}comment")) > 0)
        else None
    )
    out_p = Path(output_path).resolve()
    docx_xml_utils.save_docx_xml(inp_path, out_p, root, ns_pairs, comments_root=final_comments)
    return out_p


def run_step1_pipeline(
    input_docx: str,
    library_dir: str,
    output_dir: str,
    api_key: Union[str, List[str], ApiKeyPool] = "",
    model: str = "gemini-3.5-flash-lite",
    api_mode: str = "auto",
    rotation_mode: str = "round_robin",
    lang: str = "",
    log_fn: Optional[Callable[[str], None]] = None,
    progress_fn: Optional[Callable[[float, str], None]] = None,
    enable_ai: bool = True,
    force_full_diff: bool = False,
    highlight_colors: bool = True,
    highlight_mode: Optional[str] = None,
    yellow_gap_words: int = 3,
    export_raw_trilingual: bool = False,
    raw_trilingual_dir: str = "",
    raw_trilingual_label: str = "",
    date_str: str = "",
    balance_columns: bool = True,
) -> dict:
    _, use_hl, _ = resolve_highlight_mode(highlight_mode, highlight_colors)
    highlight_colors = use_hl
    eff_enable_ai = bool(enable_ai and highlight_colors)
    pool_or_keys, eff_rot, has_keys, effective_ai = _prepare_optional_pool(
        api_key=api_key,
        rotation_mode=rotation_mode,
        enable_ai=eff_enable_ai,
    )
    inp_p = Path(input_docx).resolve()
    lib_p = Path(library_dir).resolve()
    out_d = Path(output_dir).resolve()
    out_d.mkdir(parents=True, exist_ok=True)

    if not inp_p.exists():
        raise FileNotFoundError(f"找不到待整理文档：{inp_p}")
    if not lib_p.exists():
        raise FileNotFoundError(f"找不到语种总库文件夹：{lib_p}")

    safe_inp_stem = docx_xml_utils.sanitize_filename_part(inp_p.stem, fallback="原稿")
    wdir = _get_work_dir(safe_inp_stem)

    # Standalone fallback: auto-expand 2-column tables if `xml-core` has not yet merged into `docx_xml_utils`
    effective_inp_p = inp_p
    if not callable(getattr(docx_xml_utils, "ensure_three_column_tables_and_landscape", None)):
        inp_root, _, inp_ns = docx_xml_utils.load_docx_xml(inp_p)
        exp_cnt = ensure_three_column_tables_and_landscape(inp_root)
        if exp_cnt > 0:
            pre_path = wdir / "_preprocessed_input.docx"
            inp_cmts = docx_xml_utils.load_docx_comments_xml(inp_p)
            docx_xml_utils.save_docx_xml(inp_p, pre_path, inp_root, inp_ns, comments_root=inp_cmts)
            effective_inp_p = pre_path
            if log_fn:
                log_fn(f"  ℹ️ 检测到双列表格，已自动扩展 {exp_cnt} 个表格为三列并设为横向 A4。")

    if progress_fn:
        progress_fn(5.0, "第一步 (1/3)：正在检索语种总库、提取浅蓝底色细粒度修订...")
    if log_fn:
        log_fn("▶ [第一步 1/3] 匹配语种总库、回退白底旧英文、提取浅蓝底色细粒度修订...")

    argv1 = [
        "step1_extract_and_diff.py",
        "--input", str(effective_inp_p),
        "--library", str(lib_p),
        "--workdir", str(wdir),
    ]
    clean_lang = (lang or "").strip().upper()
    if clean_lang and clean_lang != "AUTO":
        argv1.extend(["--lang", clean_lang])
    if force_full_diff and _script_supports_cli_flag(step1_extract_and_diff, "--force-full-diff", "format_up_rv_count_tag"):
        argv1.append("--force-full-diff")

    _run_module_main(step1_extract_and_diff.main, argv1, log_fn=log_fn)

    manifest_path = wdir / "diff_manifest.json"
    review_items_path = wdir / "step1_review_items.json"
    with open(manifest_path, "r", encoding="utf-8") as f:
        manifest = json.load(f)
    with open(review_items_path, "r", encoding="utf-8") as f:
        review_items = json.load(f)

    skipped_rows = manifest.get("skipped_rows") or []
    if skipped_rows and log_fn:
        log_fn(f"⚠️ 共跳过 {len(skipped_rows)} 行未在总库匹配到的段落")
        for sk in skipped_rows[:10]:
            art = sk.get("article_title") or "无标题表格"
            r_no = int(sk.get("row_idx", 0)) + 1
            preview = (sk.get("cn_preview") or "")[:36]
            log_fn(f"    - [{art}] 第 {r_no} 行：{preview}...")

    detected_lang = manifest.get("lang_code") or (clean_lang if clean_lang and clean_lang != "AUTO" else "XX")
    edited_cnt = sum(1 for x in review_items if x.get("has_edit"))

    step1_stats: dict = {}
    if not highlight_colors:
        if log_fn:
            log_fn("▶ [第一步 2/3] 已关闭高亮颜色标记，跳过黄/青颜色分级与 AI 调用（仅输出对比修订痕迹）...")
        decisions = {}
    else:
        if enable_ai and not has_keys and log_fn:
            log_fn("  ℹ️ 未填写 API Key，本次自动使用内置规则评估黄(UP)/青(RV)分级...")

        mode_desc = f"调用 {model}" if effective_ai else "使用本地内置规则"
        if progress_fn:
            progress_fn(22.0, f"第一步 (2/3)：{mode_desc}评估 {edited_cnt} 处改动句（黄 UP / 青 RV）...")
        if log_fn:
            log_fn(f"▶ [第一步 2/3] {mode_desc}评估 {edited_cnt} 处改动句（黄 UP / 青 RV 分级）...")

        def _sub_prog(frac: float, msg: str):
            if progress_fn:
                progress_fn(22.0 + frac * 63.0, f"第一步 (2/3)：{msg}")

        decisions = evaluate_step1_decisions(
            review_items=review_items,
            api_key=pool_or_keys,
            model=model,
            api_mode=api_mode,
            rotation_mode=eff_rot,
            log_fn=log_fn,
            progress_fn=_sub_prog,
            enable_ai=effective_ai,
            stats_out=step1_stats,
        )

    decisions_path = wdir / "step1_decisions.json"
    with open(decisions_path, "w", encoding="utf-8") as f:
        json.dump(decisions, f, ensure_ascii=False, indent=2)

    draft_path = docx_xml_utils.safe_join_output_file(out_d, f"初筛_{safe_inp_stem}.docx")
    if progress_fn:
        progress_fn(88.0, "第一步 (3/3)：正在渲染初筛 Word 文档（保留原稿字体/字号/排版）...")
    if log_fn:
        log_fn("▶ [第一步 3/3] 渲染初筛稿（保留原文档字体/字号/排版及全部修订痕迹）...")

    raw_tri_path_obj: Optional[Path] = None
    if export_raw_trilingual:
        tri_dir = Path(raw_trilingual_dir).resolve() if (raw_trilingual_dir or "").strip() else out_d
        tri_dir.mkdir(parents=True, exist_ok=True)
        safe_lang_tri = docx_xml_utils.sanitize_filename_part(detected_lang, fallback="XX", max_len=16)
        safe_lbl_tri = docx_xml_utils.sanitize_filename_part(
            (raw_trilingual_label or clean_doc_label(inp_p.stem)).strip(),
            fallback="Document",
            max_len=120,
        )
        eff_dt_tri = docx_xml_utils.sanitize_filename_part(
            (date_str or "").strip() or datetime.now().strftime("%Y%m%d"),
            fallback=datetime.now().strftime("%Y%m%d"),
            max_len=16,
        )
        raw_tri_name = f"{safe_lang_tri}_{safe_lbl_tri}_提取三语原稿_{eff_dt_tri}.docx"
        raw_tri_path_obj = docx_xml_utils.safe_join_output_file(tri_dir, raw_tri_name)

    supports_ygap = _script_supports_cli_flag(step1_render_draft, "--yellow-gap-words")
    supports_no_hl1 = _script_supports_cli_flag(step1_render_draft, "--no-highlight-colors")
    supports_raw_tri = _script_supports_cli_flag(step1_render_draft, "--raw-trilingual-output")

    argv1b = [
        "step1_render_draft.py",
        "--manifest", str(manifest_path),
        "--decisions", str(decisions_path),
        "--output", str(draft_path),
    ]
    clamped_gap = max(0, int(yellow_gap_words if yellow_gap_words is not None else 3))
    if supports_ygap:
        argv1b.extend(["--yellow-gap-words", str(clamped_gap)])
    if not highlight_colors and supports_no_hl1:
        argv1b.append("--no-highlight-colors")
    if raw_tri_path_obj is not None and supports_raw_tri:
        argv1b.extend(["--raw-trilingual-output", str(raw_tri_path_obj)])

    _run_module_main(step1_render_draft.main, argv1b, log_fn=log_fn)

    # Standalone fallback when step1_render_draft does not yet support --no-highlight-colors
    if not highlight_colors and not supports_no_hl1 and draft_path.is_file():
        d_root, _, d_ns = docx_xml_utils.load_docx_xml(draft_path)
        d_cmts = docx_xml_utils.load_docx_comments_xml(draft_path)
        _strip_all_highlights_in_root(d_root)
        docx_xml_utils.save_docx_xml(draft_path, draft_path, d_root, d_ns, comments_root=d_cmts)

    # Standalone fallback when step1_render_draft does not yet support --raw-trilingual-output
    if raw_tri_path_obj is not None and (not supports_raw_tri or not raw_tri_path_obj.is_file()):
        _build_raw_trilingual_docx_from_manifest(
            manifest_path=manifest_path,
            output_path=raw_tri_path_obj,
            balance_columns=bool(balance_columns),
        )
        if log_fn:
            log_fn(f"  📄 已导出对比前提取三语原稿：{raw_tri_path_obj.name}")

    meta_payload = {"lang_code": detected_lang, "doc_label": clean_doc_label(inp_p.stem)}
    for meta_dir in (wdir, get_secure_user_workdir(safe_inp_stem)):
        try:
            (meta_dir / "session_meta.json").write_text(
                json.dumps(meta_payload, ensure_ascii=False), encoding="utf-8"
            )
        except OSError:
            pass

    if progress_fn:
        progress_fn(100.0, "第一步已完成：初筛稿已生成！")

    used_fallback = bool(
        highlight_colors
        and (
            step1_stats.get("used_local_fallback", False)
            or (edited_cnt > 0 and not effective_ai)
        )
    )
    fb_reason = step1_stats.get("fallback_reason", "") if highlight_colors else ""
    if highlight_colors and not fb_reason and enable_ai and not has_keys:
        fb_reason = "未填写 API Key"

    return {
        "draft_path": str(draft_path),
        "lang_code": detected_lang,
        "edited_sentences": edited_cnt,
        "workdir": str(wdir),
        "skipped_rows": skipped_rows,
        "skipped_count": len(skipped_rows),
        "used_local_fallback": used_fallback,
        "fallback_reason": fb_reason,
        "raw_trilingual_path": str(raw_tri_path_obj) if raw_tri_path_obj is not None else "",
    }


def _postprocess_rendered_docx(
    saved_path: Path,
    lang_code: str,
    target_label: str,
    date_yyyymmdd: str,
    balance_columns: bool,
    step2_native_v17: bool,
    highlight_colors: bool = True,
    step2_supports_no_hl: bool = False,
) -> Tuple[Path, int, int]:
    """
    Ensure every rendered output file has:
      - Balanced 3-column table widths (when `balance_columns=True`)
      - Top 'Blue/Yellow highlighted parts...' instruction lines synchronized with UP/RV counts
        (or all highlights & header lines stripped when `highlight_colors=False`)
      - Zero word counts omitted in filename (`format_counted_docx_name`)
    """
    cfg = docx_xml_utils.load_config()
    col2_up = cfg.get("col2_up_highlight", "yellow")
    col2_rv = cfg.get("col2_rv_highlight", "cyan")

    root, _, ns_pairs = docx_xml_utils.load_docx_xml(saved_path)
    needs_xml_rewrite = not step2_native_v17

    if not highlight_colors:
        u_w, r_w = 0, 0
        if not step2_supports_no_hl or root.find(".//w:rPr/w:highlight", NS) is not None:
            _strip_all_highlights_in_root(root)
            needs_xml_rewrite = True
    else:
        u_w, r_w = docx_xml_utils.compute_col2_up_rv_words(root, up_color=col2_up, rv_color=col2_rv)

    desired_name = format_counted_docx_name(
        lang_code=lang_code,
        orig_name=target_label,
        up_words=u_w,
        rv_words=r_w,
        date_yyyymmdd=date_yyyymmdd,
    )
    desired_path = docx_xml_utils.safe_join_output_file(saved_path.parent, desired_name)

    if needs_xml_rewrite:
        comments_root = docx_xml_utils.load_docx_comments_xml(saved_path)
        if balance_columns:
            balance_trilingual_table_widths(root)
        update_top_highlight_instruction_lines(
            root,
            has_up=bool(highlight_colors and u_w > 0),
            has_rv=bool(highlight_colors and r_w > 0),
        )
        docx_xml_utils.save_docx_xml(saved_path, desired_path, root, ns_pairs, comments_root=comments_root)
        if os.path.normcase(str(desired_path.resolve())) != os.path.normcase(str(saved_path.resolve())):
            try:
                saved_path.unlink(missing_ok=True)
            except OSError:
                pass
        return desired_path, int(u_w), int(r_w)

    if os.path.normcase(str(desired_path.resolve())) != os.path.normcase(str(saved_path.resolve())):
        if desired_path.exists():
            desired_path.unlink()
        saved_path.rename(desired_path)
    return desired_path, int(u_w), int(r_w)


def run_step2_pipeline(
    draft_docx: str,
    output_dir: str,
    api_key: Union[str, List[str], ApiKeyPool] = "",
    model: str = "gemini-3.5-flash-lite",
    api_mode: str = "auto",
    rotation_mode: str = "round_robin",
    lang: str = "",
    doc_name: str = "",
    dual_outputs: bool = True,
    date_mmdd: str = "",
    log_fn: Optional[Callable[[str], None]] = None,
    progress_fn: Optional[Callable[[float, str], None]] = None,
    enable_ai: bool = True,
    naming_mode: str = "keep",
    scattered_keyword: str = "",
    balance_columns: bool = True,
    dedup_pkg_dir: str = "",
    date_str: str = "",
    highlight_colors: bool = True,
    highlight_mode: Optional[str] = None,
) -> dict:
    _, use_hl, highlight_col1 = resolve_highlight_mode(highlight_mode, highlight_colors)
    highlight_colors = use_hl
    eff_enable_ai = bool(enable_ai and highlight_colors)
    pool_or_keys, eff_rot, has_keys, effective_ai = _prepare_optional_pool(
        api_key=api_key,
        rotation_mode=rotation_mode,
        enable_ai=eff_enable_ai,
    )
    draft_p = Path(draft_docx).resolve()
    out_d = Path(output_dir).resolve()
    out_d.mkdir(parents=True, exist_ok=True)

    if not draft_p.exists():
        raise FileNotFoundError(f"找不到初筛稿文档：{draft_p}")

    base_stem = draft_p.stem[len("初筛_") :] if draft_p.stem.startswith("初筛_") else draft_p.stem
    wdir = _get_work_dir(base_stem)

    clean_lang = (lang or "").strip().upper()
    if not clean_lang or clean_lang == "AUTO":
        legacy_meta = get_secure_user_workdir(base_stem) / "session_meta.json"
        if legacy_meta.exists():
            try:
                clean_lang = json.loads(legacy_meta.read_text(encoding="utf-8")).get("lang_code", "XX")
            except Exception:
                clean_lang = "XX"
        else:
            clean_lang = "XX"

    base_label = _resolve_pipeline_doc_label(
        explicit_doc_name=doc_name,
        naming_mode=naming_mode,
        scattered_keyword=scattered_keyword,
        fallback_stem=draft_p.stem,
        source_docx_path=draft_p,
    )
    date_yyyymmdd = (date_str or date_mmdd or "").strip() or datetime.now().strftime("%Y%m%d")

    if progress_fn:
        progress_fn(5.0, "第二步 (1/3)：正在解析初筛稿高亮区间...")
    if log_fn:
        log_fn("▶ [第二步 1/3] 解析初筛稿高亮区间（提取第 2 列黄色与青色片段）...")

    argv2a = [
        "step2_parse_reviewed.py",
        "--draft", str(draft_p),
        "--workdir", str(wdir),
    ]
    _run_module_main(step2_parse_reviewed.main, argv2a, log_fn=log_fn)

    tasks_path = wdir / "step2_align_tasks.json"
    with open(tasks_path, "r", encoding="utf-8") as f:
        tasks_data = json.load(f)

    total_spans = sum(len(r.get("spans", [])) for r in tasks_data.get("row_tasks", []))
    total_rows = len(tasks_data.get("row_tasks", []))

    step2_stats: dict = {}
    if not highlight_colors:
        if log_fn:
            log_fn("▶ [第二步 2/3] 已关闭高亮颜色标记，跳过第 1/3 列高亮对齐与 AI 调用...")
        alignments = {}
        total_spans = 0
    else:
        if enable_ai and not has_keys and log_fn:
            log_fn("  ℹ️ 未填写 API Key，自动采用「本地位置估算模式」完成第 3 列外文标色...")

        mode_desc2 = f"调用 {model}" if effective_ai else "使用「本地位置估算模式」"
        if progress_fn:
            progress_fn(14.0, f"第二步 (2/3)：正在{mode_desc2}对齐 {total_rows} 行共 {total_spans} 处外文高亮...")
        if log_fn:
            log_fn(f"▶ [第二步 2/3] {mode_desc2}将 {total_spans} 处第 2 列高亮对齐至第 3 列外文...")

        def _sub_prog2(frac: float, msg: str):
            if progress_fn:
                progress_fn(14.0 + frac * 74.0, f"第二步 (2/3)：{msg}")

        lang_hint = find_lang_display_name(clean_lang, load_languages())
        alignments = align_step2_spans(
            tasks_data=tasks_data,
            find_fuzzy_fn=docx_xml_utils.find_fuzzy_substring_span,
            api_key=pool_or_keys,
            model=model,
            api_mode=api_mode,
            rotation_mode=eff_rot,
            log_fn=log_fn,
            progress_fn=_sub_prog2,
            target_lang_hint=lang_hint,
            enable_ai=effective_ai,
            stats_out=step2_stats,
        )

    alignments_path = wdir / "step2_alignments.json"
    with open(alignments_path, "w", encoding="utf-8") as f:
        json.dump(alignments, f, ensure_ascii=False, indent=2)

    supports_no_bal = _script_supports_cli_flag(
        step2_render_final, "--no-balance-columns", "balance_trilingual_table_widths"
    )
    supports_no_hl2 = _script_supports_cli_flag(step2_render_final, "--no-highlight-colors")
    supports_no_col1 = _script_supports_cli_flag(step2_render_final, "--no-col1-highlights")
    step2_native_v17 = supports_no_bal and callable(getattr(docx_xml_utils, "format_up_rv_count_tag", None))

    saved_files: List[str] = []
    up_words: Optional[int] = None
    rv_words: Optional[int] = None
    dedup_removed_count = 0
    dedup_report_path = ""
    clean_dedup_dir = (dedup_pkg_dir or "").strip()

    if dual_outputs:
        if progress_fn:
            progress_fn(89.0, "第二步 (3/3)：正在写入三份定稿（发译者版 + 完整留底版 + 无改点版）...")
        if log_fn:
            log_fn("▶ [第二步 3/3] 生成三份定稿（1. 发译者版[仅改点]  2. 完整留底版[全量]  3. 无改点版[仅无改点内容]）...")
        variants = [
            ("", "--remove-unchanged-tables", "发译者版(仅改点)"),
            ("_完整留底版", None, "完整留底版"),
            ("_无改点版", "--only-unchanged-tables", "无改点内容版"),
        ]
        for idx, (label_suffix, filter_flag, desc_lbl) in enumerate(variants):
            if progress_fn:
                progress_fn(89.0 + idx * 3.0, f"第二步 (3/3)：正在写出 {desc_lbl} ...")
            target_label = f"{base_label}{label_suffix}"
            argv2b = [
                "step2_render_final.py",
                "--tasks", str(tasks_path),
                "--alignments", str(alignments_path),
                "--outdir", str(out_d),
                "--lang", clean_lang,
                "--doc-name", target_label,
                "--date", date_yyyymmdd,
            ]
            if filter_flag:
                argv2b.append(filter_flag)
            if not balance_columns and supports_no_bal:
                argv2b.append("--no-balance-columns")
            if not highlight_colors and supports_no_hl2:
                argv2b.append("--no-highlight-colors")
            elif highlight_colors and not highlight_col1 and supports_no_col1:
                argv2b.append("--no-col1-highlights")
            out_txt = _run_module_main(step2_render_final.main, argv2b, log_fn=log_fn)
            m = re.search(r"Final Document Saved:\s*(.+)", out_txt)
            if m:
                raw_saved = Path(m.group(1).strip())
                final_p, u_w, r_w = _postprocess_rendered_docx(
                    saved_path=raw_saved,
                    lang_code=clean_lang,
                    target_label=target_label,
                    date_yyyymmdd=date_yyyymmdd,
                    balance_columns=bool(balance_columns),
                    step2_native_v17=step2_native_v17,
                    highlight_colors=bool(highlight_colors),
                    step2_supports_no_hl=supports_no_hl2,
                )
                if idx == 0 and clean_dedup_dir:
                    d_res = remove_unmerged_duplicates_from_docx(
                        docx_path=final_p,
                        unmerged_pkg_dir=clean_dedup_dir,
                        report_dir=out_d,
                        doc_tag=base_label,
                        log_fn=log_fn,
                        balance_columns=bool(balance_columns),
                        date_str=date_yyyymmdd,
                    )
                    dedup_removed_count = int(d_res.get("duplicate_count", 0))
                    dedup_report_path = str(d_res.get("report_path") or "")
                    u_w = int(d_res.get("up_words", u_w)) if highlight_colors else 0
                    r_w = int(d_res.get("rv_words", r_w)) if highlight_colors else 0
                    if dedup_removed_count > 0:
                        recalc_name = format_counted_docx_name(
                            lang_code=clean_lang,
                            orig_name=target_label,
                            up_words=u_w,
                            rv_words=r_w,
                            date_yyyymmdd=date_yyyymmdd,
                        )
                        recalc_p = docx_xml_utils.safe_join_output_file(out_d, recalc_name)
                        if os.path.normcase(str(recalc_p.resolve())) != os.path.normcase(str(final_p.resolve())):
                            if recalc_p.exists():
                                recalc_p.unlink()
                            final_p.rename(recalc_p)
                            final_p = recalc_p
                saved_files.append(str(final_p))
                if idx == 0 or (idx == 1 and up_words is None):
                    up_words, rv_words = u_w, r_w
    else:
        if progress_fn:
            progress_fn(92.0, "第二步 (3/3)：正在写出最终定稿...")
        if log_fn:
            log_fn("▶ [第二步 3/3] 生成最终定稿...")
        argv2b = [
            "step2_render_final.py",
            "--tasks", str(tasks_path),
            "--alignments", str(alignments_path),
            "--outdir", str(out_d),
            "--lang", clean_lang,
            "--doc-name", base_label,
            "--date", date_yyyymmdd,
        ]
        if not balance_columns and supports_no_bal:
            argv2b.append("--no-balance-columns")
        if not highlight_colors and supports_no_hl2:
            argv2b.append("--no-highlight-colors")
        elif highlight_colors and not highlight_col1 and supports_no_col1:
            argv2b.append("--no-col1-highlights")
        out_txt = _run_module_main(step2_render_final.main, argv2b, log_fn=log_fn)
        m = re.search(r"Final Document Saved:\s*(.+)", out_txt)
        if m:
            raw_saved = Path(m.group(1).strip())
            final_p, u_w, r_w = _postprocess_rendered_docx(
                saved_path=raw_saved,
                lang_code=clean_lang,
                target_label=base_label,
                date_yyyymmdd=date_yyyymmdd,
                balance_columns=bool(balance_columns),
                step2_native_v17=step2_native_v17,
                highlight_colors=bool(highlight_colors),
                step2_supports_no_hl=supports_no_hl2,
            )
            if clean_dedup_dir:
                d_res = remove_unmerged_duplicates_from_docx(
                    docx_path=final_p,
                    unmerged_pkg_dir=clean_dedup_dir,
                    report_dir=out_d,
                    doc_tag=base_label,
                    log_fn=log_fn,
                    balance_columns=bool(balance_columns),
                    date_str=date_yyyymmdd,
                )
                dedup_removed_count = int(d_res.get("duplicate_count", 0))
                dedup_report_path = str(d_res.get("report_path") or "")
                u_w = int(d_res.get("up_words", u_w)) if highlight_colors else 0
                r_w = int(d_res.get("rv_words", r_w)) if highlight_colors else 0
                if dedup_removed_count > 0:
                    recalc_name = format_counted_docx_name(
                        lang_code=clean_lang,
                        orig_name=base_label,
                        up_words=u_w,
                        rv_words=r_w,
                        date_yyyymmdd=date_yyyymmdd,
                    )
                    recalc_p = docx_xml_utils.safe_join_output_file(out_d, recalc_name)
                    if os.path.normcase(str(recalc_p.resolve())) != os.path.normcase(str(final_p.resolve())):
                        if recalc_p.exists():
                            recalc_p.unlink()
                        final_p.rename(recalc_p)
                        final_p = recalc_p
            saved_files.append(str(final_p))
            up_words, rv_words = u_w, r_w

    if (up_words is None or rv_words is None) and saved_files:
        if not highlight_colors:
            up_words, rv_words = 0, 0
        else:
            try:
                cfg = docx_xml_utils.load_config()
                r_doc, _, _ = docx_xml_utils.load_docx_xml(Path(saved_files[0]))
                up_words, rv_words = docx_xml_utils.compute_col2_up_rv_words(
                    r_doc,
                    up_color=cfg.get("col2_up_highlight", "yellow"),
                    rv_color=cfg.get("col2_rv_highlight", "cyan"),
                )
            except Exception:
                up_words, rv_words = 0, 0

    if progress_fn:
        progress_fn(100.0, f"第二步已完成：共生成 {len(saved_files)} 个定稿文件！")

    used_fallback = bool(
        highlight_colors
        and (
            step2_stats.get("used_local_fallback", False)
            or (total_spans > 0 and not effective_ai)
        )
    )
    fb_reason = step2_stats.get("fallback_reason", "") if highlight_colors else ""
    if highlight_colors and not fb_reason and enable_ai and not has_keys:
        fb_reason = "未填写 API Key"

    return {
        "saved_files": saved_files,
        "lang_code": clean_lang,
        "total_spans": total_spans,
        "up_words": int(up_words or 0),
        "rv_words": int(rv_words or 0),
        "used_local_fallback": used_fallback,
        "fallback_reason": fb_reason,
        "dedup_removed_count": dedup_removed_count,
        "dedup_report_path": dedup_report_path,
        "workdir": str(wdir),
    }


def _is_generated_sc1_artifact_name(filename: str) -> bool:
    """Skip previously generated output artifacts when scanning an input directory."""
    nm = (filename or "").strip()
    if not nm or nm.startswith("~$") or nm.startswith("."):
        return True
    skip_tokens = (
        "_提取三语原稿_",
        "_完整留底版",
        "_无改点版",
        "查重结果明细报告",
        "_初筛稿",
    )
    return any(tok in nm for tok in skip_tokens)


def resolve_sc1_input_docx_list(input_docx: Union[str, Path, Sequence[Union[str, Path]]]) -> List[Path]:
    """
    Resolve Scenario 1 (`零散改点整理`) input specification into an ordered, deduplicated list of `.docx` files.
    Supports:
      - A single `.docx` path
      - A folder path containing multiple `.docx` files (top-level first, recursive fallback)
      - Multiple `.docx` / folder paths joined by `;` or `\\n`
      - A Python list/tuple of paths
    """
    raw_items: List[str] = []
    if isinstance(input_docx, (list, tuple, set)):
        for x in input_docx:
            if str(x or "").strip():
                raw_items.append(str(x).strip())
    else:
        s = str(input_docx or "").strip()
        if s:
            for part in re.split(r"[;\r\n]+", s):
                clean_p = part.strip().strip('"').strip("'")
                if clean_p:
                    raw_items.append(clean_p)

    def _nat_key(p: Path):
        return [int(tok) if tok.isdigit() else tok.lower() for tok in re.split(r"(\d+)", p.name)]

    resolved: List[Path] = []
    seen_norm: set = set()

    def _add_file(fp: Path):
        rp = fp.resolve()
        nk = os.path.normcase(str(rp))
        if nk not in seen_norm:
            seen_norm.add(nk)
            resolved.append(rp)

    for item in raw_items:
        p = Path(item)
        if not p.exists():
            continue
        rp = p.resolve()
        if rp.is_file():
            if rp.suffix.lower() == ".docx" and not rp.name.startswith("~$"):
                _add_file(rp)
        elif rp.is_dir():
            top_docs = sorted(
                (f for f in rp.glob("*.docx") if f.is_file() and not _is_generated_sc1_artifact_name(f.name)),
                key=_nat_key,
            )
            if top_docs:
                for f in top_docs:
                    _add_file(f)
            else:
                sub_docs = sorted(
                    (f for f in rp.rglob("*.docx") if f.is_file() and not _is_generated_sc1_artifact_name(f.name)),
                    key=_nat_key,
                )
                for f in sub_docs:
                    _add_file(f)
    return resolved


def _run_single_oneclick_pipeline(
    inp_p: Path,
    lib_p: Path,
    out_d: Path,
    api_key: Union[str, List[str], ApiKeyPool] = "",
    model: str = "gemini-3.5-flash-lite",
    api_mode: str = "auto",
    rotation_mode: str = "round_robin",
    lang: str = "AUTO",
    doc_name: str = "",
    dual_outputs: bool = True,
    enable_ai: bool = True,
    naming_mode: str = "scattered",
    scattered_keyword: str = "",
    balance_columns: bool = True,
    dedup_pkg_dir: str = "",
    eff_date: str = "",
    log_fn: Optional[Callable[[str], None]] = None,
    progress_fn: Optional[Callable[[float, str], None]] = None,
    latest_master_path: str = "",
    highlight_colors: bool = True,
    highlight_mode: Optional[str] = None,
    yellow_gap_words: int = 3,
    export_raw_trilingual: bool = True,
    skip_index_warmup: bool = False,
) -> dict:
    """Execute one-click Scattered Revision pipeline for a single resolved `inp_p` `.docx` file."""
    out_d.mkdir(parents=True, exist_ok=True)

    eff_hl_mode, use_hl, _ = resolve_highlight_mode(highlight_mode, highlight_colors)
    highlight_colors = use_hl
    eff_enable_ai = bool(enable_ai and highlight_colors)
    pool_or_keys, eff_rot, has_keys, effective_ai = _prepare_optional_pool(
        api_key=api_key,
        rotation_mode=rotation_mode,
        enable_ai=eff_enable_ai,
    )

    eff_date = (eff_date or "").strip() or datetime.now().strftime("%Y%m%d")

    pre_saved_files: List[str] = []
    master_sync_res: Optional[dict] = None
    effective_input_p = inp_p
    clean_master_path = (latest_master_path or "").strip()
    if clean_master_path:
        if progress_fn:
            progress_fn(1.0, "正在核对最新中英底稿并同步更新待处理文档...")
        master_sync_res = sync_pending_with_latest_master(
            pending_docx=inp_p,
            latest_master_path=clean_master_path,
            output_dir=out_d,
            date_str=eff_date,
            log_fn=log_fn,
        )
        upd_path_str = str(master_sync_res["updated_docx_path"])
        pre_saved_files.append(upd_path_str)
        effective_input_p = Path(upd_path_str).resolve()

    if not skip_index_warmup:
        if progress_fn:
            progress_fn(3.0, "正在预热/加载语种总库检索索引...")
        idx_stats = ensure_library_index(str(lib_p), log_fn=log_fn)
        if log_fn:
            log_fn(
                f"📚 语种总库索引就绪：共 {idx_stats.get('total_files', 0)} 篇文档"
                f"（{idx_stats.get('total_rows', 0)} 段，耗时 {idx_stats.get('elapsed_sec', 0.0):.2f}s）"
            )

    resolved_doc_name = _resolve_pipeline_doc_label(
        explicit_doc_name=doc_name,
        naming_mode=naming_mode,
        scattered_keyword=scattered_keyword,
        fallback_stem=inp_p.stem,
        source_docx_path=effective_input_p,
    )

    tmp_draft_dir = tempfile.mkdtemp(prefix="scattered_rev_draft_")
    wdir1_str = ""
    wdir2_str = ""
    try:
        def _prog_s1(p: float, msg: str):
            if progress_fn:
                progress_fn(6.0 + (p / 100.0) * 44.0, msg)

        r1 = run_step1_pipeline(
            input_docx=str(effective_input_p),
            library_dir=str(lib_p),
            output_dir=tmp_draft_dir,
            api_key=pool_or_keys,
            model=model,
            api_mode=api_mode,
            rotation_mode=eff_rot,
            lang=lang,
            log_fn=log_fn,
            progress_fn=_prog_s1,
            enable_ai=effective_ai,
            highlight_colors=bool(highlight_colors),
            highlight_mode=eff_hl_mode,
            yellow_gap_words=yellow_gap_words,
            export_raw_trilingual=bool(export_raw_trilingual),
            raw_trilingual_dir=str(out_d),
            raw_trilingual_label=resolved_doc_name,
            date_str=eff_date,
            balance_columns=bool(balance_columns),
        )
        wdir1_str = str(r1.get("workdir") or "")

        det_lang = r1.get("lang_code") or "XX"
        skipped_rows = r1.get("skipped_rows") or []
        step1_failed_api = bool(effective_ai and r1.get("used_local_fallback") and r1.get("fallback_reason"))

        step2_enable_ai = effective_ai and bool(highlight_colors)
        if step1_failed_api and highlight_colors:
            if log_fn:
                log_fn(f"  ⚠️ {r1['fallback_reason']} -> 已自动切换「本地位置估算模式」完成第 3 列标色")
            step2_enable_ai = False

        def _prog_s2(p: float, msg: str):
            if progress_fn:
                progress_fn(50.0 + (p / 100.0) * 50.0, msg)

        step2_lang = det_lang if (not lang or lang.strip().upper() == "AUTO") else lang
        r2 = run_step2_pipeline(
            draft_docx=r1["draft_path"],
            output_dir=str(out_d),
            api_key=pool_or_keys,
            model=model,
            api_mode=api_mode,
            rotation_mode=eff_rot,
            lang=step2_lang,
            doc_name=resolved_doc_name,
            dual_outputs=dual_outputs,
            date_mmdd=eff_date,
            log_fn=log_fn,
            progress_fn=_prog_s2,
            enable_ai=step2_enable_ai,
            naming_mode=naming_mode,
            scattered_keyword=scattered_keyword,
            balance_columns=balance_columns,
            dedup_pkg_dir=dedup_pkg_dir,
            date_str=eff_date,
            highlight_colors=bool(highlight_colors),
            highlight_mode=eff_hl_mode,
        )
        wdir2_str = str(r2.get("workdir") or "")
    finally:
        shutil.rmtree(tmp_draft_dir, ignore_errors=True)
        for wd_s in (wdir1_str, wdir2_str):
            if wd_s:
                shutil.rmtree(wd_s, ignore_errors=True)

    all_saved: List[str] = []
    for fp_s in [*pre_saved_files, *r2["saved_files"], r1.get("raw_trilingual_path") or ""]:
        if fp_s and fp_s not in all_saved:
            all_saved.append(fp_s)

    used_fallback = bool(
        highlight_colors
        and (
            r1.get("used_local_fallback")
            or r2.get("used_local_fallback")
            or (enable_ai and not has_keys)
            or (not enable_ai)
        )
    )
    fallback_reason = (r2.get("fallback_reason") or r1.get("fallback_reason") or "") if highlight_colors else ""

    if progress_fn:
        progress_fn(100.0, f"一键整理已完成：共输出 {len(all_saved)} 份文档！")

    return {
        "saved_files": all_saved,
        "step2_files": list(r2.get("saved_files") or []),
        "lang_code": r2.get("lang_code") or det_lang,
        "skipped_count": len(skipped_rows),
        "skipped_rows": skipped_rows,
        "used_local_fallback": used_fallback,
        "fallback_reason": fallback_reason,
        "up_words": int(r2.get("up_words", 0)),
        "rv_words": int(r2.get("rv_words", 0)),
        "total_spans": int(r2.get("total_spans", 0)),
        "edited_sentences": int(r1.get("edited_sentences", 0)),
        "dedup_removed_count": int(r2.get("dedup_removed_count", 0)),
        "dedup_report_path": r2.get("dedup_report_path", ""),
        "raw_trilingual_path": r1.get("raw_trilingual_path", ""),
        "updated_master_docx_path": pre_saved_files[0] if pre_saved_files else "",
        "master_sync_result": master_sync_res,
    }


def _disambiguate_batch_dest(dst_dir: Path, filename: str, fallback_tag: str = "") -> Path:
    """Ensure `dst_dir / filename` never overwrites an existing file from an earlier document in the same batch."""
    dst_dir.mkdir(parents=True, exist_ok=True)
    cand = docx_xml_utils.safe_join_output_file(dst_dir, filename)
    if not cand.exists():
        return cand
    stem = Path(filename).stem
    ext = Path(filename).suffix or ".docx"
    # Insert fallback_tag or numeric counter before `_(RV` / `_(UP` / `_YYYYMMDD` if present
    m_tail = re.search(r"(_\((?:RV|UP)\b.*|_\d{8})$", stem)
    for k in range(2, 500):
        tag = f"_{fallback_tag}" if (k == 2 and fallback_tag and fallback_tag not in stem) else f"_{k}"
        if m_tail:
            new_stem = stem[: m_tail.start()] + tag + stem[m_tail.start() :]
        else:
            new_stem = f"{stem}{tag}"
        cand_k = docx_xml_utils.safe_join_output_file(dst_dir, f"{new_stem}{ext}")
        if not cand_k.exists():
            return cand_k
    return cand


def run_oneclick_pipeline(
    input_docx: Union[str, Path, Sequence[Union[str, Path]]],
    library_dir: str,
    output_dir: str,
    api_key: Union[str, List[str], ApiKeyPool] = "",
    model: str = "gemini-3.5-flash-lite",
    api_mode: str = "auto",
    rotation_mode: str = "round_robin",
    lang: str = "AUTO",
    doc_name: str = "",
    dual_outputs: bool = True,
    enable_ai: bool = True,
    naming_mode: str = "scattered",
    scattered_keyword: str = "",
    balance_columns: bool = True,
    dedup_pkg_dir: str = "",
    date_str: str = "",
    log_fn: Optional[Callable[[str], None]] = None,
    progress_fn: Optional[Callable[[float, str], None]] = None,
    date_mmdd: str = "",
    latest_master_path: str = "",
    highlight_colors: bool = True,
    highlight_mode: Optional[str] = None,
    yellow_gap_words: int = 3,
    export_raw_trilingual: bool = True,
) -> dict:
    """
    One-click end-to-end Scattered Revision pipeline supporting BOTH:
      - Single `.docx` input (flat output in `output_dir`, 100% backward-compatible)
      - Multi-file / Folder input (`len(input_files) > 1`): warms library index once, processes every
        `.docx` in order, preserves distinct per-document names without collisions, and neatly archives
        multi-doc variants into `给译者版/` & `完整留底版/` when `dual_outputs=True`.
    """
    input_files = resolve_sc1_input_docx_list(input_docx)
    lib_p = Path(library_dir).resolve()
    out_d = Path(output_dir).resolve()
    if not input_files:
        raise FileNotFoundError(f"未在待整理路径中找到有效的 .docx 文档：{input_docx}")
    if not lib_p.exists():
        raise FileNotFoundError(f"找不到语种总库文件夹：{lib_p}")
    out_d.mkdir(parents=True, exist_ok=True)
    eff_date = (date_str or date_mmdd or "").strip() or datetime.now().strftime("%Y%m%d")

    # Single-file fast path: 100% identical flat output behavior
    if len(input_files) == 1:
        res_single = _run_single_oneclick_pipeline(
            inp_p=input_files[0],
            lib_p=lib_p,
            out_d=out_d,
            api_key=api_key,
            model=model,
            api_mode=api_mode,
            rotation_mode=rotation_mode,
            lang=lang,
            doc_name=doc_name,
            dual_outputs=dual_outputs,
            enable_ai=enable_ai,
            naming_mode=naming_mode,
            scattered_keyword=scattered_keyword,
            balance_columns=balance_columns,
            dedup_pkg_dir=dedup_pkg_dir,
            eff_date=eff_date,
            log_fn=log_fn,
            progress_fn=progress_fn,
            latest_master_path=latest_master_path,
            highlight_colors=highlight_colors,
            highlight_mode=highlight_mode,
            yellow_gap_words=yellow_gap_words,
            export_raw_trilingual=export_raw_trilingual,
            skip_index_warmup=False,
        )
        res_single["processed_docs_count"] = 1
        res_single["failed_docs"] = []
        return res_single

    # Multi-document / Folder Batch Mode (`len(input_files) > 1`)
    total_docs = len(input_files)
    if log_fn:
        log_fn(f"📂 批量零散改点模式：共检测到 {total_docs} 份待整理 .docx 文档，开始预热总库索引并逐份整理...")
    if progress_fn:
        progress_fn(2.0, f"正在预热语种总库索引（共 {total_docs} 份待整理文档）...")

    idx_stats = ensure_library_index(str(lib_p), log_fn=log_fn)
    if log_fn:
        log_fn(
            f"📚 语种总库索引就绪：共 {idx_stats.get('total_files', 0)} 篇文档"
            f"（{idx_stats.get('total_rows', 0)} 段，耗时 {idx_stats.get('elapsed_sec', 0.0):.2f}s）"
        )

    eff_hl_mode, use_hl, _ = resolve_highlight_mode(highlight_mode, highlight_colors)
    shared_pool, eff_rot, _, _ = _prepare_optional_pool(
        api_key=api_key,
        rotation_mode=rotation_mode,
        enable_ai=bool(enable_ai and use_hl),
    )

    all_saved_batch: List[str] = []
    all_skipped_rows: List[dict] = []
    failed_docs: List[str] = []
    total_up = 0
    total_rv = 0
    total_spans = 0
    total_edited = 0
    total_dedup = 0
    any_local_fallback = False
    last_fb_reason = ""
    detected_lang_final = lang

    translator_dir = (out_d / "给译者版") if dual_outputs else out_d
    archive_dir = out_d / "完整留底版"
    unchanged_dir = out_d / "无改点版"
    raw_tri_dir = out_d / "提取三语原稿"
    master_upd_dir = out_d / "底稿同步更新"

    for doc_idx, cur_inp in enumerate(input_files):
        base_pct = 5.0 + (doc_idx / total_docs) * 92.0
        span_pct = 92.0 / total_docs
        if log_fn:
            log_fn(f"\n📄 === [{doc_idx + 1}/{total_docs}] 正在整理：{cur_inp.name} ===")

        def _sub_prog(pct_inner: float, msg_inner: str, _b=base_pct, _s=span_pct, _i=doc_idx + 1, _nm=cur_inp.name):
            if progress_fn:
                progress_fn(_b + (pct_inner / 100.0) * _s, f"[{_i}/{total_docs} {_nm}] {msg_inner}")

        # Determine safe per-file naming in batch mode so multiple files never collapse onto the same name
        eff_doc_name = ""
        eff_naming_mode = naming_mode
        eff_kw = (scattered_keyword or "").strip()
        if (doc_name or "").strip():
            eff_doc_name = f"{doc_name.strip()}_{doc_idx + 1}"
            eff_naming_mode = "keep"
        elif (naming_mode or "").strip().lower() == "scattered" and not eff_kw:
            # When no custom keyword is set in batch mode, preserve each file's own distinct title/stem!
            eff_naming_mode = "keep"

        stage_dir = Path(tempfile.mkdtemp(prefix=f"sc1_batch_{doc_idx + 1}_"))
        try:
            sub_res = _run_single_oneclick_pipeline(
                inp_p=cur_inp,
                lib_p=lib_p,
                out_d=stage_dir,
                api_key=shared_pool,
                model=model,
                api_mode=api_mode,
                rotation_mode=eff_rot,
                lang=lang,
                doc_name=eff_doc_name,
                dual_outputs=dual_outputs,
                enable_ai=enable_ai,
                naming_mode=eff_naming_mode,
                scattered_keyword=eff_kw,
                balance_columns=balance_columns,
                dedup_pkg_dir=dedup_pkg_dir,
                eff_date=eff_date,
                log_fn=log_fn,
                progress_fn=_sub_prog,
                latest_master_path=latest_master_path,
                highlight_colors=highlight_colors,
                highlight_mode=eff_hl_mode,
                yellow_gap_words=yellow_gap_words,
                export_raw_trilingual=export_raw_trilingual,
                skip_index_warmup=True,
            )

            orig_tag = clean_doc_label(cur_inp.stem)[:32]
            if sub_res.get("lang_code") and sub_res["lang_code"] != "XX":
                detected_lang_final = sub_res["lang_code"]
            total_up += int(sub_res.get("up_words", 0))
            total_rv += int(sub_res.get("rv_words", 0))
            total_spans += int(sub_res.get("total_spans", 0))
            total_edited += int(sub_res.get("edited_sentences", 0))
            total_dedup += int(sub_res.get("dedup_removed_count", 0))
            all_skipped_rows.extend(sub_res.get("skipped_rows") or [])
            if sub_res.get("used_local_fallback"):
                any_local_fallback = True
                last_fb_reason = sub_res.get("fallback_reason") or last_fb_reason

            # Move staged outputs into organized destination folders without collisions
            upd_master_s = sub_res.get("updated_master_docx_path") or ""
            if upd_master_s and Path(upd_master_s).is_file():
                src_m = Path(upd_master_s)
                dst_m = _disambiguate_batch_dest(master_upd_dir, src_m.name, orig_tag)
                shutil.move(str(src_m), str(dst_m))
                all_saved_batch.append(str(dst_m))

            step2_list = [Path(x) for x in (sub_res.get("step2_files") or []) if x and Path(x).is_file()]
            for s2_p in step2_list:
                nm = s2_p.name
                if "_完整留底版" in nm:
                    clean_nm = nm.replace("_完整留底版", "")
                    dst_p = _disambiguate_batch_dest(archive_dir, clean_nm, orig_tag)
                elif "_无改点版" in nm:
                    clean_nm = nm.replace("_无改点版", "")
                    dst_p = _disambiguate_batch_dest(unchanged_dir, clean_nm, orig_tag)
                else:
                    dst_p = _disambiguate_batch_dest(translator_dir, nm, orig_tag)
                shutil.move(str(s2_p), str(dst_p))
                all_saved_batch.append(str(dst_p))

            raw_tri_s = sub_res.get("raw_trilingual_path") or ""
            if raw_tri_s and Path(raw_tri_s).is_file():
                src_rt = Path(raw_tri_s)
                dst_rt = _disambiguate_batch_dest(raw_tri_dir, src_rt.name, orig_tag)
                shutil.move(str(src_rt), str(dst_rt))
                all_saved_batch.append(str(dst_rt))

            dedup_rep_s = sub_res.get("dedup_report_path") or ""
            if dedup_rep_s and Path(dedup_rep_s).is_file():
                src_dr = Path(dedup_rep_s)
                dst_dr = _disambiguate_batch_dest(out_d, src_dr.name, orig_tag)
                shutil.move(str(src_dr), str(dst_dr))
                all_saved_batch.append(str(dst_dr))

        except Exception as ex:
            failed_docs.append(f"{cur_inp.name}（{ex}）")
            if log_fn:
                log_fn(f"❌ [{doc_idx + 1}/{total_docs}] 处理失败：{cur_inp.name} -> {ex}")
        finally:
            shutil.rmtree(stage_dir, ignore_errors=True)

    if not all_saved_batch and failed_docs:
        raise RuntimeError(f"批量处理全部失败（共 {len(failed_docs)} 份）：\n" + "\n".join(failed_docs[:5]))

    if progress_fn:
        progress_fn(
            100.0,
            f"批量整理完成：成功 {total_docs - len(failed_docs)}/{total_docs} 份，共输出 {len(all_saved_batch)} 个文件！",
        )

    return {
        "saved_files": all_saved_batch,
        "lang_code": detected_lang_final,
        "processed_docs_count": total_docs - len(failed_docs),
        "total_input_docs": total_docs,
        "failed_docs": failed_docs,
        "skipped_count": len(all_skipped_rows),
        "skipped_rows": all_skipped_rows,
        "used_local_fallback": any_local_fallback,
        "fallback_reason": last_fb_reason,
        "up_words": total_up,
        "rv_words": total_rv,
        "total_spans": total_spans,
        "edited_sentences": total_edited,
        "dedup_removed_count": total_dedup,
        "duplicate_count": total_dedup,
        "dedup_report_path": "",
        "raw_trilingual_path": "",
    }
