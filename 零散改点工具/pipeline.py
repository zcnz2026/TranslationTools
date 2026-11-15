import io
import json
import re
import sys
import tempfile
from contextlib import redirect_stdout
from datetime import datetime
from pathlib import Path
from typing import Callable, List, Optional, Union
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
from security_utils import get_secure_user_workdir  # noqa: E402

# Ensure SCATTERED_REV_LANGUAGES_JSON env var points to external languages.json
get_languages_file()


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
    s = stem.strip()
    if s.startswith("初筛_"):
        s = s[len("初筛_") :]
    s = re.sub(r"_[0-9]{4,8}$", "", s)
    return docx_xml_utils.sanitize_filename_part(s.strip() or stem.strip(), fallback="定稿")


def run_step1_pipeline(
    input_docx: str,
    library_dir: str,
    output_dir: str,
    api_key: Union[str, List[str], ApiKeyPool],
    model: str,
    api_mode: str = "auto",
    rotation_mode: str = "round_robin",
    lang: str = "",
    log_fn: Optional[Callable[[str], None]] = None,
    progress_fn: Optional[Callable[[float, str], None]] = None,
) -> dict:
    pool = api_key if isinstance(api_key, ApiKeyPool) else ApiKeyPool(api_key, rotation_mode=rotation_mode)
    inp_p = Path(input_docx).resolve()
    lib_p = Path(library_dir).resolve()
    out_d = Path(output_dir).resolve()
    out_d.mkdir(parents=True, exist_ok=True)

    if not inp_p.exists():
        raise FileNotFoundError(f"找不到待整理文档：{inp_p}")
    if not lib_p.exists():
        raise FileNotFoundError(f"找不到语种总库文件夹：{lib_p}")

    wdir = _get_work_dir(inp_p.stem)
    if progress_fn:
        progress_fn(5.0, "第一步 (1/3)：正在扫描语种总库、提取浅蓝底色细粒度修订...")
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

    detected_lang = manifest.get("lang_code") or (clean_lang if clean_lang != "AUTO" else "XX")
    edited_cnt = sum(1 for x in review_items if x.get("has_edit"))

    if progress_fn:
        progress_fn(22.0, f"第一步 (2/3)：调用 {model} 评估 {edited_cnt} 处改动句（黄 UP / 青 RV）...")
    if log_fn:
        log_fn(f"▶ [第一步 2/3] 调用 {model} 评估 {edited_cnt} 处改动句（黄 UP / 青 RV 分级）...")

    def _sub_prog(frac: float, msg: str):
        if progress_fn:
            progress_fn(22.0 + frac * 63.0, f"第一步 (2/3)：{msg}")

    decisions = evaluate_step1_decisions(
        review_items=review_items,
        api_key=pool,
        model=model,
        api_mode=api_mode,
        rotation_mode=pool.rotation_mode,
        log_fn=log_fn,
        progress_fn=_sub_prog,
    )
    decisions_path = wdir / "step1_decisions.json"
    with open(decisions_path, "w", encoding="utf-8") as f:
        json.dump(decisions, f, ensure_ascii=False, indent=2)

    safe_inp_stem = docx_xml_utils.sanitize_filename_part(inp_p.stem, fallback="原稿")
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

    return {
        "draft_path": str(draft_path),
        "lang_code": detected_lang,
        "edited_sentences": edited_cnt,
        "workdir": str(wdir),
    }


def run_step2_pipeline(
    draft_docx: str,
    output_dir: str,
    api_key: Union[str, List[str], ApiKeyPool],
    model: str,
    api_mode: str = "auto",
    rotation_mode: str = "round_robin",
    lang: str = "",
    doc_name: str = "",
    dual_outputs: bool = True,
    date_mmdd: str = "",
    log_fn: Optional[Callable[[str], None]] = None,
    progress_fn: Optional[Callable[[float, str], None]] = None,
) -> dict:
    pool = api_key if isinstance(api_key, ApiKeyPool) else ApiKeyPool(api_key, rotation_mode=rotation_mode)
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
    mmdd = (date_mmdd or "").strip() or datetime.now().strftime("%m%d")

    if progress_fn:
        progress_fn(5.0, "第二步 (1/3)：正在解析确认后的初筛稿高亮区间...")
    if log_fn:
        log_fn("▶ [第二步 1/3] 解析确认后的初筛稿高亮区间（提取第 2 列黄色与青色片段）...")

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
    if progress_fn:
        progress_fn(14.0, f"第二步 (2/3)：正在对齐 {total_rows} 行共 {total_spans} 处外文高亮...")
    if log_fn:
        log_fn(f"▶ [第二步 2/3] 调用 {model} 将 {total_spans} 处第 2 列高亮精准对齐至第 3 列外文...")

    def _sub_prog2(frac: float, msg: str):
        if progress_fn:
            progress_fn(14.0 + frac * 74.0, f"第二步 (2/3)：{msg}")

    lang_hint = find_lang_display_name(clean_lang, load_languages())
    alignments = align_step2_spans(
        tasks_data=tasks_data,
        find_fuzzy_fn=docx_xml_utils.find_fuzzy_substring_span,
        api_key=pool,
        model=model,
        api_mode=api_mode,
        rotation_mode=pool.rotation_mode,
        log_fn=log_fn,
        progress_fn=_sub_prog2,
        target_lang_hint=lang_hint,
    )
    alignments_path = wdir / "step2_alignments.json"
    with open(alignments_path, "w", encoding="utf-8") as f:
        json.dump(alignments, f, ensure_ascii=False, indent=2)

    saved_files = []
    if dual_outputs:
        if progress_fn:
            progress_fn(90.0, "第二步 (3/3)：正在写入双版本定稿（完整留底版 + 发译者版）...")
        if log_fn:
            log_fn("▶ [第二步 3/3] 生成双版本定稿（1. 完整留底版  2. 剔除无改点表的发译者版）...")
        for idx, (label_suffix, rm_unchanged) in enumerate([("_完整留底版", False), ("_发译者版", True)]):
            if progress_fn:
                progress_fn(90.0 + idx * 5.0, f"第二步 (3/3)：正在写出 {base_label}{label_suffix} ...")
            argv2b = [
                "step2_render_final.py",
                "--tasks", str(tasks_path),
                "--alignments", str(alignments_path),
                "--outdir", str(out_d),
                "--lang", clean_lang,
                "--doc-name", f"{base_label}{label_suffix}",
                "--date", mmdd,
            ]
            if rm_unchanged:
                argv2b.append("--remove-unchanged-tables")
            out_txt = _run_module_main(step2_render_final.main, argv2b, log_fn=log_fn)
            m = re.search(r"Final Document Saved:\s*(.+)", out_txt)
            if m:
                saved_files.append(m.group(1).strip())
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
            "--date", mmdd,
        ]
        out_txt = _run_module_main(step2_render_final.main, argv2b, log_fn=log_fn)
        m = re.search(r"Final Document Saved:\s*(.+)", out_txt)
        if m:
            saved_files.append(m.group(1).strip())

    if progress_fn:
        progress_fn(100.0, f"第二步已完成：共生成 {len(saved_files)} 个定稿文件！")

    return {
        "saved_files": saved_files,
        "lang_code": clean_lang,
        "total_spans": total_spans,
    }
