from datetime import datetime
import os
from pathlib import Path
import re
import shutil
import sys
from typing import Optional


def _ensure_skill_scripts():
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


_ensure_skill_scripts()

import docx_xml_utils  # noqa: E402
from lang_registry import load_languages  # noqa: E402

# Matches new & legacy word-count tags:
#   _(UP 120 words, RV 45 words)
#   _(UP  words, RV  words)
#   _(UP 120 words) / _(RV 45 words)
#   _(up452 rv609) / _(up0)
_WORD_COUNT_TAG_RE = re.compile(
    r"[\s_]*\(\s*(?:"
    r"(?:UP|RV)\s*\d*\s*(?:words?)?(?:\s*[,，/&\s]\s*(?:UP|RV)\s*\d*\s*(?:words?)?)*"
    r"|"
    r"(?:up|rv)\s*\d+(?:\s+(?:up|rv)\s*\d+)*"
    r")\s*\)",
    re.IGNORECASE,
)

_TRAILING_DATE_RE = re.compile(r"_((?:19|20)\d{6}|\d{4})$")


def parse_doc_filename(
    stem: str,
    known_codes: Optional[set[str]] = None,
    strip_draft_prefix: bool = True,
) -> dict:
    """
    Parse a .docx filename stem into structured components:
      - existing_lang: e.g. 'GR' if prefixed with 'GR_', else ''
      - orig_name: clean original document title (without lang prefix, word count tag, or trailing date)
      - existing_date: 8-digit YYYYMMDD if present at the end, else ''
      - had_word_count: True if stem already contained an UP/RV word count block
    """
    s = (stem or "").strip()
    if strip_draft_prefix and s.startswith("初筛_"):
        s = s[len("初筛_") :].strip()

    codes_upper = {c.upper() for c in (known_codes or set())}

    # 1. Extract trailing date before or after word count tag
    existing_date = ""
    m_date = _TRAILING_DATE_RE.search(s)
    if m_date:
        raw_dt = m_date.group(1)
        if len(raw_dt) == 8:
            existing_date = raw_dt
        s = s[: m_date.start()].strip(" _")

    # 2. Detect & strip all word-count parenthetical tags
    had_word_count = bool(_WORD_COUNT_TAG_RE.search(s))
    if had_word_count:
        s = _WORD_COUNT_TAG_RE.sub("", s).strip(" _")

    # Check again if a trailing date was sitting before the word count tag
    m_date2 = _TRAILING_DATE_RE.search(s)
    if m_date2:
        raw_dt2 = m_date2.group(1)
        if len(raw_dt2) == 8 and not existing_date:
            existing_date = raw_dt2
        s = s[: m_date2.start()].strip(" _")

    # 3. Extract leading language code (e.g. 'GR_God’s Words')
    existing_lang = ""
    m_lang = re.match(r"^([A-Za-z]{2,5})_(.+)$", s)
    if m_lang:
        cand_code = m_lang.group(1)
        rest_name = m_lang.group(2).strip()
        if rest_name and (cand_code.upper() in codes_upper or re.fullmatch(r"[A-Z]{2,3}", cand_code)):
            existing_lang = cand_code.upper()
            s = rest_name

    orig_name = s.strip() or (stem or "Document").strip()
    return {
        "existing_lang": existing_lang,
        "orig_name": orig_name,
        "existing_date": existing_date,
        "had_word_count": had_word_count,
    }


def detect_lang_from_text(text: str, lang_entries: list[dict]) -> str:
    t_low = (text or "").lower()
    best_code = ""
    best_len = 0
    for entry in lang_entries:
        code = entry.get("code", "")
        for kw in entry.get("keywords", []):
            k_low = kw.lower().strip()
            if k_low and k_low in t_low and len(k_low) > best_len:
                best_code = code
                best_len = len(k_low)
    return best_code


def format_counted_docx_name(
    lang_code: str,
    orig_name: str,
    up_words: int,
    rv_words: int,
    date_yyyymmdd: str,
) -> str:
    """
    Format standard output filename:
      {LANG}_{ORIG_NAME}_(UP {up_words} words, RV {rv_words} words)_{YYYYMMDD}.docx
    Example:
      GR_God’s Words_(UP 128 words, RV 64 words)_20261115.docx
    """
    safe_lang = docx_xml_utils.sanitize_filename_part((lang_code or "XX").strip().upper(), fallback="XX", max_len=16)
    safe_orig = docx_xml_utils.sanitize_filename_part((orig_name or "Document").strip(), fallback="Document", max_len=140)
    today_str = datetime.now().strftime("%Y%m%d")
    safe_date = docx_xml_utils.sanitize_filename_part((date_yyyymmdd or today_str).strip(), fallback=today_str, max_len=16)
    count_tag = f"(UP {int(up_words)} words, RV {int(rv_words)} words)"
    return f"{safe_lang}_{safe_orig}_{count_tag}_{safe_date}.docx"


def analyze_docx_word_count(
    docx_path: Path,
    default_lang_code: str = "AUTO",
    keep_existing_lang: bool = True,
    date_yyyymmdd: str = "",
    keep_existing_date: bool = False,
    strip_draft_prefix: bool = True,
    lang_entries: Optional[list[dict]] = None,
) -> dict:
    """
    Count Col 2 English words (UP=yellow, RV=cyan; punctuation excluded) in `docx_path`
    and compute its standardized target filename.
    """
    p = Path(docx_path).resolve()
    if not p.is_file():
        raise FileNotFoundError(f"找不到文件：{p}")

    entries = lang_entries if lang_entries is not None else load_languages()
    known_codes = {e["code"].upper() for e in entries if e.get("code")}

    cfg = docx_xml_utils.load_config()
    col2_up = cfg.get("col2_up_highlight", "yellow")
    col2_rv = cfg.get("col2_rv_highlight", "cyan")

    root, _, _ = docx_xml_utils.load_docx_xml(p)
    up_words, rv_words = docx_xml_utils.compute_col2_up_rv_words(root, up_color=col2_up, rv_color=col2_rv)

    parsed = parse_doc_filename(p.stem, known_codes=known_codes, strip_draft_prefix=strip_draft_prefix)

    # Resolve target language code
    sel_code = (default_lang_code or "AUTO").strip().upper()
    if keep_existing_lang and parsed["existing_lang"]:
        final_lang = parsed["existing_lang"]
    elif sel_code and sel_code != "AUTO":
        final_lang = sel_code
    else:
        det = detect_lang_from_text(f"{p.parent.name} {p.stem}", entries)
        final_lang = det or parsed["existing_lang"] or "XX"

    # Resolve target date (YYYYMMDD)
    today_str = datetime.now().strftime("%Y%m%d")
    if keep_existing_date and parsed["existing_date"]:
        final_date = parsed["existing_date"]
    else:
        final_date = (date_yyyymmdd or "").strip() or today_str

    new_filename = format_counted_docx_name(
        lang_code=final_lang,
        orig_name=parsed["orig_name"],
        up_words=up_words,
        rv_words=rv_words,
        date_yyyymmdd=final_date,
    )

    return {
        "path": str(p),
        "old_name": p.name,
        "new_name": new_filename,
        "lang_code": final_lang,
        "orig_name": parsed["orig_name"],
        "date_str": final_date,
        "up_words": up_words,
        "rv_words": rv_words,
        "total_words": up_words + rv_words,
        "had_word_count": parsed["had_word_count"],
    }


def apply_word_count_filename(
    docx_path: Path,
    new_filename: str,
    copy_to_dir: Optional[Path] = None,
) -> tuple[Path, str]:
    """
    Rename `docx_path` in place (or copy to `copy_to_dir`) with `new_filename`.
    Returns (final_path, action_description).
    """
    src = Path(docx_path).resolve()
    if not src.is_file():
        raise FileNotFoundError(f"原文档不存在：{src}")

    if copy_to_dir is not None:
        out_dir = Path(copy_to_dir).resolve()
        out_dir.mkdir(parents=True, exist_ok=True)
        dst = docx_xml_utils.safe_join_output_file(out_dir, new_filename)
        shutil.copy2(src, dst)
        return dst, "✅ 已另存副本"

    dst = docx_xml_utils.safe_join_output_file(src.parent, new_filename)
    if src.name == dst.name:
        return src, "✅ 文件名已是最新"

    # Check collision with a DIFFERENT file in the same directory
    if dst.exists() and os.path.normcase(str(src)) != os.path.normcase(str(dst)):
        raise FileExistsError(f"同目录下已存在同名文件，为防覆盖已跳过：{dst.name}")

    try:
        src.rename(dst)
    except PermissionError as pe:
        raise PermissionError(f"文件正在被 Word 或其他程序占用，请先关闭文档后重试：{src.name}") from pe

    return dst, "✅ 已更新文件名"
