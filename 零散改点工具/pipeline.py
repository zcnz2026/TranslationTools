from concurrent.futures import ThreadPoolExecutor
import importlib
import io
import json
import re
import shutil
import sys
import tempfile
import time
from contextlib import redirect_stdout
from datetime import datetime
from pathlib import Path
from typing import Callable, List, Optional, Tuple, Union
from unittest.mock import patch


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
from lang_registry import find_lang_display_name, get_languages_file, load_languages  # noqa: E402
from security_utils import get_secure_user_workdir, parse_api_keys  # noqa: E402
from word_counter import parse_doc_filename  # noqa: E402

# Ensure SCATTERED_REV_LANGUAGES_JSON env var points to external languages.json
get_languages_file()

_OUTPUT_COUNT_TAG_RE = re.compile(r"_\(UP\s+\d+\s+words,\s*RV\s+\d+\s+words\)_", re.IGNORECASE)
_WORD_COUNTS_STDOUT_RE = re.compile(r"Final Col 2 Word Counts:\s*UP\s*=\s*(\d+)\s*words,\s*RV\s*=\s*(\d+)\s*words")


def ensure_library_index(
    library_dir: Union[str, Path],
    log_fn: Optional[Callable[[str], None]] = None,
    max_workers: int = 8,
) -> dict:
    """
    Build or refresh the persistent multi-threaded library index via `library_index.ensure_library_index`,
    with a safe standalone fallback if `library_index` is not yet present on disk.
    """
    _ensure_skill_scripts_on_path()
    try:
        lib_idx_mod = importlib.import_module("library_index")
        fn = getattr(lib_idx_mod, "ensure_library_index", None)
        if callable(fn):
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
    """Captures stdout while streaming each completed line immediately to log_fn."""

    def __init__(self, log_fn: Optional[Callable[[str], None]] = None):
        super().__init__()
        self._log_fn = log_fn
        self._full = []
        self._buf = ""

    def write(self, s: str) -> int:
        if not s:
            return 0
        self._full.append(s)
        if self._log_fn:
            self._buf += s
            while "\n" in self._buf:
                line, self._buf = self._buf.split("\n", 1)
                line = line.rstrip("\r")
                if line.strip():
                    self._log_fn(f"  {line}")
        return len(s)

    def flush(self):
        if self._log_fn and self._buf.strip():
            self._log_fn(f"  {self._buf.rstrip()}")
            self._buf = ""

    def getvalue(self) -> str:
        self.flush()
        return "".join(self._full)


def _run_module_main(mod_main, argv: list, log_fn: Optional[Callable[[str], None]] = None) -> str:
    writer = _LiveLogWriter(log_fn=log_fn)
    with patch.object(sys, "argv", argv), redirect_stdout(writer):
        mod_main()
    return writer.getvalue()


def _get_work_dir(stem: str) -> Path:
    return get_secure_user_workdir(stem)


def clean_doc_label(stem: str) -> str:
    known_codes = {e["code"].upper() for e in load_languages() if e.get("code")}
    parsed = parse_doc_filename(stem, known_codes=known_codes, strip_draft_prefix=True)
    return docx_xml_utils.sanitize_filename_part(parsed["orig_name"], fallback="Document", max_len=140)


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
) -> dict:
    pool_or_keys, eff_rot, has_keys, effective_ai = _prepare_optional_pool(
        api_key=api_key,
        rotation_mode=rotation_mode,
        enable_ai=enable_ai,
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
    if progress_fn:
        progress_fn(5.0, "第一步 (1/3)：正在检索语种总库、提取浅蓝底色细粒度修订...")
    if log_fn:
        log_fn("▶ [第一步 1/3] 匹配语种总库、回退白底旧英文、提取浅蓝底色细粒度修订...")

    argv1 = [
        "step1_extract_and_diff.py",
        "--input", str(inp_p),
        "--library", str(lib_p),
        "--workdir", str(wdir),
    ]
    clean_lang = (lang or "").strip().upper()
    if clean_lang and clean_lang != "AUTO":
        argv1.extend(["--lang", clean_lang])

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

    step1_stats: dict = {}
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

    argv1b = [
        "step1_render_draft.py",
        "--manifest", str(manifest_path),
        "--decisions", str(decisions_path),
        "--output", str(draft_path),
    ]
    _run_module_main(step1_render_draft.main, argv1b, log_fn=log_fn)

    meta_path = wdir / "session_meta.json"
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump({"lang_code": detected_lang, "doc_label": clean_doc_label(inp_p.stem)}, f, ensure_ascii=False)

    if progress_fn:
        progress_fn(100.0, "第一步已完成：初筛稿已生成！")

    used_fallback = bool(
        step1_stats.get("used_local_fallback", False)
        or (edited_cnt > 0 and not effective_ai)
    )
    fb_reason = step1_stats.get("fallback_reason", "")
    if not fb_reason and enable_ai and not has_keys:
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
    }


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
) -> dict:
    pool_or_keys, eff_rot, has_keys, effective_ai = _prepare_optional_pool(
        api_key=api_key,
        rotation_mode=rotation_mode,
        enable_ai=enable_ai,
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
        meta_path = wdir / "session_meta.json"
        if meta_path.exists():
            with open(meta_path, "r", encoding="utf-8") as f:
                clean_lang = json.load(f).get("lang_code", "XX")
        else:
            clean_lang = "XX"

    base_label = (doc_name or "").strip() or clean_doc_label(draft_p.stem)
    date_yyyymmdd = (date_mmdd or "").strip() or datetime.now().strftime("%Y%m%d")

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
    step2_stats: dict = {}
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

    saved_files: List[str] = []
    up_words: Optional[int] = None
    rv_words: Optional[int] = None

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
                progress_fn(89.0 + idx * 3.5, f"第二步 (3/3)：正在写出 {desc_lbl} ...")
            argv2b = [
                "step2_render_final.py",
                "--tasks", str(tasks_path),
                "--alignments", str(alignments_path),
                "--outdir", str(out_d),
                "--lang", clean_lang,
                "--doc-name", f"{base_label}{label_suffix}",
                "--date", date_yyyymmdd,
            ]
            if filter_flag:
                argv2b.append(filter_flag)
            out_txt = _run_module_main(step2_render_final.main, argv2b, log_fn=log_fn)
            m = re.search(r"Final Document Saved:\s*(.+)", out_txt)
            if m:
                saved_files.append(m.group(1).strip())
            # Capture true UP / RV counts from 发译者版 (idx 0) or 完整留底版 (idx 1), never overwritten by 无改点版 (0, 0)
            if idx in (0, 1) and up_words is None:
                m_wc = _WORD_COUNTS_STDOUT_RE.search(out_txt)
                if m_wc:
                    up_words = int(m_wc.group(1))
                    rv_words = int(m_wc.group(2))
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
        out_txt = _run_module_main(step2_render_final.main, argv2b, log_fn=log_fn)
        m = re.search(r"Final Document Saved:\s*(.+)", out_txt)
        if m:
            saved_files.append(m.group(1).strip())
        m_wc = _WORD_COUNTS_STDOUT_RE.search(out_txt)
        if m_wc:
            up_words = int(m_wc.group(1))
            rv_words = int(m_wc.group(2))

    if (up_words is None or rv_words is None) and saved_files:
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
        step2_stats.get("used_local_fallback", False)
        or (total_spans > 0 and not effective_ai)
    )
    fb_reason = step2_stats.get("fallback_reason", "")
    if not fb_reason and enable_ai and not has_keys:
        fb_reason = "未填写 API Key"

    return {
        "saved_files": saved_files,
        "lang_code": clean_lang,
        "total_spans": total_spans,
        "up_words": int(up_words or 0),
        "rv_words": int(rv_words or 0),
        "used_local_fallback": used_fallback,
        "fallback_reason": fb_reason,
    }


def run_oneclick_pipeline(
    input_docx: str,
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
    log_fn: Optional[Callable[[str], None]] = None,
    progress_fn: Optional[Callable[[float, str], None]] = None,
    date_mmdd: str = "",
) -> dict:
    """
    One-click end-to-end Scattered Revision pipeline:
      1. Warms/loads the persistent multi-threaded library index;
      2. Runs Step 1 into an ephemeral temp directory (keeping `output_dir` 100% free of intermediate drafts);
      3. Runs Step 2 (with multi-threaded AI alignment or automatic local approximate Col 3 fallback);
      4. Outputs clean final `.docx` deliverables to `output_dir`.
    """
    inp_p = Path(input_docx).resolve()
    lib_p = Path(library_dir).resolve()
    out_d = Path(output_dir).resolve()
    if not inp_p.exists():
        raise FileNotFoundError(f"找不到待整理文档：{inp_p}")
    if not lib_p.exists():
        raise FileNotFoundError(f"找不到语种总库文件夹：{lib_p}")
    out_d.mkdir(parents=True, exist_ok=True)

    pool_or_keys, eff_rot, has_keys, effective_ai = _prepare_optional_pool(
        api_key=api_key,
        rotation_mode=rotation_mode,
        enable_ai=enable_ai,
    )

    if progress_fn:
        progress_fn(2.0, "正在预热/加载语种总库检索索引...")
    idx_stats = ensure_library_index(str(lib_p), log_fn=log_fn)
    if log_fn:
        log_fn(
            f"📚 语种总库索引就绪：共 {idx_stats.get('total_files', 0)} 篇文档"
            f"（{idx_stats.get('total_rows', 0)} 段，耗时 {idx_stats.get('elapsed_sec', 0.0):.2f}s）"
        )

    resolved_doc_name = (doc_name or "").strip() or clean_doc_label(inp_p.stem)
    tmp_draft_dir = tempfile.mkdtemp(prefix="scattered_rev_draft_")
    try:
        def _prog_s1(p: float, msg: str):
            if progress_fn:
                progress_fn(6.0 + (p / 100.0) * 44.0, msg)

        r1 = run_step1_pipeline(
            input_docx=str(inp_p),
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
        )

        det_lang = r1.get("lang_code") or "XX"
        skipped_rows = r1.get("skipped_rows") or []
        step1_failed_api = bool(effective_ai and r1.get("used_local_fallback") and r1.get("fallback_reason"))

        step2_enable_ai = effective_ai
        if step1_failed_api:
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
            date_mmdd=date_mmdd,
            log_fn=log_fn,
            progress_fn=_prog_s2,
            enable_ai=step2_enable_ai,
        )
    finally:
        shutil.rmtree(tmp_draft_dir, ignore_errors=True)

    used_fallback = bool(
        r1.get("used_local_fallback")
        or r2.get("used_local_fallback")
        or (enable_ai and not has_keys)
        or (not enable_ai)
    )
    fallback_reason = r2.get("fallback_reason") or r1.get("fallback_reason") or ""

    if progress_fn:
        progress_fn(100.0, f"一键整理已完成：共输出 {len(r2['saved_files'])} 份定稿！")

    return {
        "saved_files": r2["saved_files"],
        "lang_code": r2.get("lang_code") or det_lang,
        "skipped_count": len(skipped_rows),
        "skipped_rows": skipped_rows,
        "used_local_fallback": used_fallback,
        "fallback_reason": fallback_reason,
        "up_words": int(r2.get("up_words", 0)),
        "rv_words": int(r2.get("rv_words", 0)),
        "total_spans": int(r2.get("total_spans", 0)),
        "edited_sentences": int(r1.get("edited_sentences", 0)),
    }
