from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
import hashlib
import importlib
import json
import os
from pathlib import Path
import re
import shutil
import sys
import tempfile
import threading
import time
from typing import Callable, Dict, List, Optional, Sequence, Tuple, Union
import xml.etree.ElementTree as ET
import zipfile


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
import step1_extract_and_diff  # noqa: E402
import step1_render_draft  # noqa: E402
import step2_parse_reviewed  # noqa: E402
import step2_render_final  # noqa: E402
from ai_client import ApiKeyPool, align_step2_spans, evaluate_step1_decisions  # noqa: E402
from dedup_checker import remove_unmerged_duplicates_from_docx  # noqa: E402
from lang_registry import find_lang_display_name, load_languages  # noqa: E402
from master_sync import _accept_cell_revisions_inplace  # noqa: E402
from security_utils import create_ephemeral_workdir  # noqa: E402
from pipeline import (  # noqa: E402
    _build_raw_trilingual_docx_from_manifest,
    _postprocess_rendered_docx,
    _prepare_optional_pool,
    _run_module_main,
    _script_supports_cli_flag,
    _strip_all_highlights_in_root,
    ensure_library_index,
    resolve_highlight_mode,
)
from word_counter import (  # noqa: E402
    balance_trilingual_table_widths,
    ensure_three_column_tables_and_landscape,
    extract_full_article_en_title,
    format_counted_docx_name,
    is_top_highlight_instruction_text,
    parse_doc_filename,
    update_top_highlight_instruction_lines,
)

W = docx_xml_utils.W
W_URI = docx_xml_utils.W_URI
R_URI = docx_xml_utils.R_URI
XML_SPACE = docx_xml_utils.XML_SPACE
PKG_REL_URI = docx_xml_utils.PKG_REL_URI
CONTENT_TYPES_URI = docx_xml_utils.CONTENT_TYPES_URI
NS = docx_xml_utils.NS
CN_NUM_CHARS = "一二三四五六七八九十百零"

MASTER_INDEX_SCHEMA_VERSION = 3

_BULLET_PREFIX_RE = re.compile(
    r"^\s*(?:[•●▪▸►*]+|\(\s*\d+\s*\)|（\s*\d+\s*）|\d+\s*[\.、\)）])\s*"
)
_LANG_STEM_PREFIX_RE = re.compile(r"^(?:中英[^\W\d_]{0,4}|[A-Z]{2,4})[\s_]+")
_CHAPTER_PREFIX_CAPTURE_RE = re.compile(r"^(\d+(?:-\d+)*(?:[\s_]+\d{2,4})?)[\s_]+(.*)$")
_SUB_EXCERPT_TITLE_RE = re.compile(r"^(?:选段|片段|excerpt|part)\s*\d+$", re.IGNORECASE)

_TRAILING_DATE_STEM_RE = re.compile(r"[\s_]+(?:19|20)\d{6}[\s_]*$")
_TRAILING_WORKFLOW_TAG_RE = re.compile(r"(?:[\s_]+(?:全篇TD|全篇RV|TD|RV)|(?<=[）》】])(?:全篇TD|全篇RV|TD|RV))\s*$", re.IGNORECASE)
_TRAILING_BRACKET_TAG_RE = re.compile(
    r"[\s_]*(?:"
    r"（\s*(?:审判|话|出书用|选段\s*\d+\s*-\s*\d+|\d+\s*-\s*\d+)\s*）"
    r"|\(\s*(?:审判|话|出书用|选段\s*\d+\s*-\s*\d+|\d+\s*-\s*\d+)\s*\)"
    r"|【[^】]*】"
    r"|\[(?:TEP|标)[^\]]*\]"
    r")\s*$",
    re.IGNORECASE,
)
_TRAILING_EXCERPT_RANGE_RE = re.compile(r"[\s_]*\d{1,3}\s*-\s*\d{1,3}\s*$")
_TRAILING_TITLE_PUNCT_RE = re.compile(r"[？?!！。．…；;：:]+$")
_NATURAL_SPLIT_RE = re.compile(r"(\d+)")

_PART_DIVIDER_RE = re.compile(r"^第[\s一二三四五六七八九十百\d]+部分")
_SUBPART_DIVIDER_RE = re.compile(
    r"^(?:附篇[\s\u3000]*[“\"]?神向全宇|基督在众教会行走时的说话[\s\u3000]*[一二三四五六七八九十\d]+\s*$)"
)

_WS_RE = re.compile(r"[\s\u3000\u200b\u200c\ufeff]+")
_LOOSE_KEEP_RE = re.compile(r"[\u4e00-\u9fff\u3400-\u4dbfA-Za-z0-9]+")

_MASTER_CACHE_LOCK = threading.Lock()
_MEM_MASTER_CACHE: Dict[str, Tuple[tuple, List[dict]]] = {}


# ---------------------------------------------------------------------------
# Title List Parsing & Prefix/Suffix Stripping (Item #7③)
# ---------------------------------------------------------------------------

def _clean_single_title_line(line: str) -> str:
    s = (line or "").strip().lstrip("\ufeff")
    if not s:
        return ""
    # Strip list bullets like '1. ', '2、', '(3) ', '• ' while keeping '1-01 ...' / '3-1-35 ...' chapter prefixes intact!
    s = _BULLET_PREFIX_RE.sub("", s).strip()
    if s.startswith("《") and s.endswith("》") and s.count("《") == 1 and s.count("》") == 1:
        s = s[1:-1].strip()
    return s


def _strip_title_suffixes_and_punct(title_str: str) -> str:
    """
    Iteratively strip trailing dates, workflow tags (`全篇TD`, `_RV`), parenthetical/bracket tags
    (`（审判）`, `(审判)`, `（话）`, `(话)`, `【标you】`, `【TEP...】`, `（选段10-18）`), excerpt ranges,
    and trailing punctuation (`？?!！。`).
    """
    s = (title_str or "").strip(" _")
    if not s:
        return ""
    for _ in range(8):
        prev = s
        s = _TRAILING_DATE_STEM_RE.sub("", s).strip(" _")
        s = _TRAILING_WORKFLOW_TAG_RE.sub("", s).strip(" _")
        s = _TRAILING_BRACKET_TAG_RE.sub("", s).strip(" _")
        # Strip bare trailing '1-9' / '10-18' only when Chinese characters precede it
        if re.search(r"[\u4e00-\u9fff]", s):
            s_no_rng = _TRAILING_EXCERPT_RANGE_RE.sub("", s).strip(" _")
            if s_no_rng:
                s = s_no_rng
        s = _TRAILING_TITLE_PUNCT_RE.sub("", s).strip()
        if s == prev:
            break
    if s.startswith("《") and s.endswith("》") and s.count("《") == 1 and s.count("》") == 1:
        s = s[1:-1].strip()
    return s


def split_title_prefix_and_core(raw_title: str, strip_lang_prefix: bool = False) -> Tuple[str, str]:
    """
    Parse a chapter title or stem into `(prefix_code, core_title)`:
      - `3-1-35 实行真理才是有实际` -> `("3-1-35", "实行真理才是有实际")`
      - `3-1-37 030 神的作工像人想象得那么简单吗？` -> `("3-1-37 030", "神的作工像人想象得那么简单吗")`
      - `3-1-39 032 七雷巨响——预言国度的福音将扩展全宇（审判）` -> `("3-1-39 032", "七雷巨响——预言国度的福音将扩展全宇")`
      - `3-1-42 034 信神要注重实际不是搞宗教仪式（审判）` -> `("3-1-42 034", "信神要注重实际不是搞宗教仪式")`
    """
    s = _clean_single_title_line(raw_title)
    if not s:
        return "", ""
    if strip_lang_prefix:
        s = _LANG_STEM_PREFIX_RE.sub("", s).strip(" _")

    prefix_code = ""
    m_pref = _CHAPTER_PREFIX_CAPTURE_RE.match(s)
    if m_pref:
        raw_pref = m_pref.group(1)
        rest = m_pref.group(2).strip(" _")
        if rest:
            prefix_code = re.sub(r"[\s_]+", " ", raw_pref.strip(" _"))
            s = rest

    core_title = _strip_title_suffixes_and_punct(s)
    return prefix_code, core_title or s


def _normalize_chapter_code_tuple(prefix_code: str) -> Optional[Tuple[int, ...]]:
    """Normalize primary chapter prefix (e.g. '3-1-39 032' -> (3, 1, 39), '1-001 001' -> (1, 1)) for structural matching."""
    if not prefix_code:
        return None
    primary = re.split(r"[\s_]+", prefix_code.strip())[0]
    if "-" not in primary:
        return None
    parts = primary.split("-")
    if not all(p.isdigit() for p in parts):
        return None
    return tuple(int(p) for p in parts)


def _natural_sort_key(s: str) -> tuple:
    return tuple(int(tok) if tok.isdigit() else tok.lower() for tok in _NATURAL_SPLIT_RE.split(s or ""))


def parse_titles_multiline(raw_text: str) -> List[str]:
    """
    Split multiline user input into clean, deduplicated non-empty article titles.
    Strips list bullet numbers ('1. ', '2、', '(1)') while keeping chapter prefixes ('1-01 ...', '3-1-35 ...') intact.
    """
    if not raw_text:
        return []
    titles: List[str] = []
    seen = set()
    for raw_line in str(raw_text).splitlines():
        cleaned = _clean_single_title_line(raw_line)
        if not cleaned:
            continue
        norm_key = _WS_RE.sub(" ", cleaned)
        if norm_key not in seen:
            seen.add(norm_key)
            titles.append(cleaned)
    return titles


def parse_titles_text_or_file(file_path: Union[str, Path]) -> List[str]:
    """
    Read a `.docx` or `.txt` title list (e.g. `需要对比的篇目标题.docx`) into a deduplicated list of title strings.
    """
    p = Path(file_path).resolve()
    if not p.is_file():
        raise FileNotFoundError(f"找不到篇目标题清单文件：{p}")

    if p.suffix.lower() == ".docx":
        root, _, _ = docx_xml_utils.load_docx_xml(p)
        body = root.find("w:body", NS)
        lines: List[str] = []
        if body is not None:
            for p_el in body.findall(".//w:p", NS):
                txt = "".join(t.text or "" for t in p_el.findall(".//w:t", NS)).strip()
                if txt and not is_top_highlight_instruction_text(txt):
                    lines.append(txt)
        return parse_titles_multiline("\n".join(lines))

    raw_bytes = p.read_bytes()
    decoded = None
    for enc in ("utf-8-sig", "utf-8", "gb18030", "utf-16"):
        try:
            decoded = raw_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if decoded is None:
        decoded = raw_bytes.decode("utf-8", errors="ignore")
    return parse_titles_multiline(decoded)


# ---------------------------------------------------------------------------
# Master CN-EN Directory Multi-Threaded Indexer (Single- & Multi-Article .docx)
# ---------------------------------------------------------------------------

def _norm_compact(s: str) -> str:
    s_clean = (s or "").replace("《", "").replace("》", "").replace("“", "").replace("”", "").replace('"', "")
    s_clean = re.sub(r"[—─―-]+", "—", s_clean)
    return _WS_RE.sub("", s_clean).lower()


def _norm_loose(s: str) -> str:
    return "".join(_LOOSE_KEEP_RE.findall(s or "")).lower()


def _strip_chapter_code_and_range(stem: str) -> str:
    _, core = split_title_prefix_and_core(stem, strip_lang_prefix=True)
    return core


def _resolve_master_cache_path(resolved_master_dir: Path) -> Path:
    norm_key = os.path.normcase(str(resolved_master_dir))
    sha1_hex = hashlib.sha1(norm_key.encode("utf-8")).hexdigest()
    fname = f"master_articles_{sha1_hex}.json"
    local_app = os.environ.get("LOCALAPPDATA", "").strip()
    base = (Path(local_app) / "ScatteredRevisionTool") if local_app else (Path(tempfile.gettempdir()) / "ScatteredRevisionTool")
    try:
        base.mkdir(parents=True, exist_ok=True)
        return base / fname
    except OSError:
        return Path(tempfile.gettempdir()) / fname


def _scan_master_docx_stats(master_path: Path) -> List[Tuple[str, str, int, int]]:
    """Scan master file/folder using live `os.stat` (`st_mtime_ns`, `st_size`) on every run."""
    if master_path.is_file():
        st = os.stat(master_path)
        return [(master_path.name, str(master_path.resolve()), st.st_mtime_ns, st.st_size)]

    found: List[Tuple[str, str, int, int]] = []

    def _walk(cur_dir: str, rel_prefix: str):
        try:
            with os.scandir(cur_dir) as it:
                entries = list(it)
        except OSError:
            return
        for entry in entries:
            ename = entry.name
            if entry.is_dir(follow_symlinks=False):
                low = ename.lower()
                if ename.startswith(".") or any(
                    k in low for k in ("output", "temp", "初筛", "标记结果", "给译者版", "完整留底版", "中英最新底稿包", "外文三语原稿包")
                ):
                    continue
                nxt = f"{rel_prefix}/{ename}" if rel_prefix else ename
                _walk(entry.path, nxt)
            elif entry.is_file(follow_symlinks=False):
                if not ename.lower().endswith(".docx") or ename.startswith("~$") or ename.startswith("."):
                    continue
                if (
                    ename.startswith("初筛_")
                    or ename.startswith("查重结果")
                    or ename.startswith("全篇对比_英文无改点")
                ):
                    continue
                rel_posix = f"{rel_prefix}/{ename}" if rel_prefix else ename
                try:
                    st = os.stat(entry.path)
                except OSError:
                    continue
                found.append((rel_posix, str(Path(entry.path).resolve()), st.st_mtime_ns, st.st_size))

    _walk(str(master_path), "")
    found.sort(key=lambda x: _natural_sort_key(x[0]))
    return found


def _extract_table_info(tbl: ET.Element) -> Tuple[int, str, str, str, bool]:
    """
    Inspect a `<w:tbl>` and return `(row_count, r0_cn, en_title, first_body_cn, has_blank_sep_row)`.
    Handles both standard Row 0 titles and 2-row `附篇一` + subtitle tables.
    """
    trs = tbl.findall("w:tr", NS)
    n_rows = len(trs)
    if n_rows == 0:
        return 0, "", "", "", False

    def _row_cells_text(tr: ET.Element) -> List[str]:
        return [
            re.sub(r"\s+", " ", docx_xml_utils.get_cell_plain_text(tc)).strip()
            for tc in tr.findall("w:tc", NS)
        ]

    c0 = _row_cells_text(trs[0])
    r0_cn = c0[0] if len(c0) >= 1 else ""
    r0_en = c0[1] if len(c0) >= 2 else ""

    has_blank_sep = False
    en_title = r0_en
    body_start_idx = 1

    if n_rows >= 3:
        c1 = _row_cells_text(trs[1])
        c2 = _row_cells_text(trs[2])
        if all(not x for x in c2) and len(c1) >= 2 and (r0_en.endswith(":") or r0_cn.startswith("附篇")):
            has_blank_sep = True
            r0_cn = f"{r0_cn} {c1[0]}".strip()
            en_title = f"{r0_en} {c1[1]}".strip() if r0_en else c1[1]
            body_start_idx = 3
        elif all(not x for x in c1):
            has_blank_sep = True
            body_start_idx = 2
    elif n_rows >= 2:
        c1 = _row_cells_text(trs[1])
        if all(not x for x in c1):
            has_blank_sep = True
            body_start_idx = 2

    first_body_cn = ""
    for r_i in range(body_start_idx, n_rows):
        cells = _row_cells_text(trs[r_i])
        if cells and cells[0]:
            first_body_cn = cells[0][:120]
            break
    if not first_body_cn and n_rows >= 1 and not has_blank_sep and len(r0_cn) > 30:
        first_body_cn = r0_cn[:120]

    return n_rows, r0_cn, en_title, first_body_cn, has_blank_sep


def _is_divider_cover_table(n_rows: int, p_titles: List[str], r0_cn: str, has_blank_sep: bool) -> Tuple[bool, str, str]:
    """
    Detect book/volume/part/sub-part cover divider tables in multi-article `.docx` files.
    Returns `(is_divider, kind, clean_section_label)` where `kind` is `'book'|'part'|'subpart'|''`.
    """
    if has_blank_sep or n_rows > 4:
        return False, "", ""
    joined_p = " ".join(p_titles).strip()
    cand_txt = joined_p or r0_cn
    compact_r0 = _WS_RE.sub("", r0_cn)
    compact_p = _WS_RE.sub("", joined_p)

    if compact_r0 == "话在肉身显现" or compact_p == "话在肉身显现":
        return True, "book", "话在肉身显现"
    if _PART_DIVIDER_RE.match(compact_r0) or _PART_DIVIDER_RE.match(compact_p):
        lbl = re.sub(r"[──—]+.*$", "", joined_p or r0_cn).strip()
        return True, "part", lbl
    if _SUBPART_DIVIDER_RE.match(cand_txt):
        return True, "subpart", re.sub(r"\s+", " ", cand_txt).strip()
    return False, "", ""


def _build_aliases_with_stripped_cores(raw_candidates: Sequence[str]) -> List[str]:
    aliases: List[str] = []
    for raw_c in raw_candidates:
        c_s = (raw_c or "").strip()
        if not c_s:
            continue
        if c_s not in aliases:
            aliases.append(c_s)
        _, core_c = split_title_prefix_and_core(c_s, strip_lang_prefix=True)
        if core_c and core_c not in aliases:
            aliases.append(core_c)
    return aliases


def _parse_single_master_docx(task: Tuple[str, str, int, int]) -> Tuple[str, dict]:
    rel_posix, abs_path_str, mtime_ns, size = task
    p = Path(abs_path_str)
    stem = p.stem.strip()
    stem_prefix, clean_stem_title = split_title_prefix_and_core(stem, strip_lang_prefix=True)
    parent_rel = str(Path(rel_posix).parent).replace("\\", "/")
    if parent_rel == ".":
        parent_rel = ""

    try:
        with zipfile.ZipFile(abs_path_str, "r") as zf:
            docx_xml_utils._validate_docx_zip(zf)
            xml_bytes = docx_xml_utils.read_docx_member_bounded(zf, "word/document.xml")
        root = docx_xml_utils.parse_safe_openxml_bytes(xml_bytes)
    except Exception as ex:
        return rel_posix, {
            "rel_path": rel_posix,
            "mtime_ns": mtime_ns,
            "size": size,
            "articles": [],
            "error": str(ex).splitlines()[0][:120],
        }

    body = root.find("w:body", NS)
    if body is None:
        return rel_posix, {"rel_path": rel_posix, "mtime_ns": mtime_ns, "size": size, "articles": []}

    raw_units = []
    pending_ps: List[Tuple[int, str]] = []
    for idx, ch in enumerate(body):
        loc = ch.tag.split("}")[-1]
        if loc == "p":
            txt = "".join(t.text or "" for t in ch.findall(".//w:t", NS)).strip()
            if txt and not is_top_highlight_instruction_text(txt):
                pending_ps.append((idx, txt))
        elif loc == "tbl":
            n_rows, r0_cn, en_title, first_body_cn, has_blank_sep = _extract_table_info(ch)
            start_idx = pending_ps[0][0] if pending_ps else idx
            end_idx = idx + 1
            raw_units.append({
                "start_idx": start_idx,
                "end_idx": end_idx,
                "tbl_idx": idx,
                "p_titles": [t for _, t in pending_ps],
                "n_rows": n_rows,
                "r0_cn": r0_cn,
                "en_title": en_title,
                "first_body_cn": first_body_cn,
                "has_blank_sep": has_blank_sep,
            })
            pending_ps = []

    if not raw_units:
        return rel_posix, {"rel_path": rel_posix, "mtime_ns": mtime_ns, "size": size, "articles": []}

    is_multi = False
    if len(raw_units) >= 2:
        sub_units_after_first = [
            u for u in raw_units[1:]
            if any(_SUB_EXCERPT_TITLE_RE.match(t.strip()) for t in (u["p_titles"] or [u["r0_cn"]]))
        ]
        if len(sub_units_after_first) == len(raw_units) - 1:
            is_multi = False
        else:
            is_multi = True

    articles: List[dict] = []
    if not is_multi:
        u0 = raw_units[0]
        cn_title = (u0["p_titles"][-1] if u0["p_titles"] else u0["r0_cn"]) or clean_stem_title or stem
        en_title = extract_full_article_en_title(root) or u0["en_title"]
        tot_rows = sum(u["n_rows"] for u in raw_units)
        first_body = next((u["first_body_cn"] for u in raw_units if u["first_body_cn"]), "")

        aliases = _build_aliases_with_stripped_cores((cn_title, stem, clean_stem_title, u0["r0_cn"]))

        sec_disp = parent_rel or "独立单篇"
        disp_en = f" ({en_title})" if en_title else ""
        display_label = f"{rel_posix}{disp_en} [{tot_rows}行]"

        articles.append({
            "cand_id": f"{rel_posix}::FULL",
            "display_label": display_label,
            "master_file": abs_path_str,
            "master_rel_path": rel_posix,
            "naming_stem": stem,
            "stem_prefix": stem_prefix,
            "is_multi_article": False,
            "body_range": None,
            "section_path": sec_disp,
            "cn_title": cn_title,
            "qualified_cn_title": cn_title,
            "en_title": en_title,
            "row_count": tot_rows,
            "first_body_cn": first_body,
            "aliases": aliases,
        })
    else:
        cur_part = ""
        cur_subpart = ""
        book_label = re.sub(r"\d{8}.*$", "", stem).strip(" _") or stem

        for u in raw_units:
            p_titles = u["p_titles"]
            if len(p_titles) >= 2:
                lead_p = p_titles[0]
                lead_compact = _WS_RE.sub("", lead_p)
                if _PART_DIVIDER_RE.match(lead_compact):
                    cur_part = re.sub(r"[──—]+.*$", "", lead_p).strip()
                    cur_subpart = ""
                    p_titles = p_titles[1:]
                elif _SUBPART_DIVIDER_RE.match(lead_p):
                    cur_subpart = re.sub(r"\s+", " ", lead_p).strip()
                    p_titles = p_titles[1:]

            is_div, div_kind, div_lbl = _is_divider_cover_table(
                u["n_rows"], p_titles, u["r0_cn"], u["has_blank_sep"]
            )
            if is_div:
                if div_kind == "part":
                    cur_part = div_lbl
                    cur_subpart = ""
                elif div_kind == "subpart":
                    cur_subpart = div_lbl
                continue

            if len(p_titles) >= 2 and p_titles[0].startswith("附篇"):
                cn_title = f"{p_titles[0]} {p_titles[-1]}".strip()
                leaf_sub = p_titles[-1].strip()
                extra_aliases = [cn_title, leaf_sub, p_titles[0].strip(), u["r0_cn"]]
            else:
                cn_title = (p_titles[-1] if p_titles else u["r0_cn"]).strip()
                extra_aliases = [cn_title, u["r0_cn"]]

            if not cn_title:
                continue

            sec_parts = [x for x in (cur_part, cur_subpart) if x]
            section_path = " › ".join(sec_parts)

            qual_prefix = ""
            if cur_subpart and "奥秘揭示" in cur_subpart:
                qual_prefix = "对神奥秘的揭示"
            elif cur_part:
                if "基督起初的发表" in cur_part:
                    qual_prefix = "基督起初的发表"
                elif "神向全宇的说话" in cur_part:
                    qual_prefix = "神向全宇的说话"
                elif "基督在众教会行走时的说话" in cur_part:
                    qual_prefix = "基督在众教会行走时的说话"
                else:
                    qual_prefix = cur_part

            is_generic_leaf = bool(
                re.match(r"^(?:附加[\s\u3000：:]*)?第[一二三四五六七八九十百零\d]+篇", _WS_RE.sub("", cn_title))
                or cn_title in ("内容简介", "写在前面的话", "后记", "前言")
            )
            qualified_cn_title = f"{qual_prefix}・{cn_title}" if (is_generic_leaf and qual_prefix) else cn_title

            aliases = _build_aliases_with_stripped_cores([*extra_aliases, qualified_cn_title])

            b_start, b_end = u["start_idx"], u["end_idx"]
            en_t = u["en_title"]
            disp_sec = f" › {section_path}" if section_path else ""
            disp_en = f" ({en_t})" if en_t else ""
            display_label = f"{book_label}{disp_sec} › {cn_title}{disp_en} [{u['n_rows']}行]"
            ch_pref, _ = split_title_prefix_and_core(cn_title)

            articles.append({
                "cand_id": f"{rel_posix}::{b_start}:{b_end}",
                "display_label": display_label,
                "master_file": abs_path_str,
                "master_rel_path": rel_posix,
                "naming_stem": qualified_cn_title,
                "stem_prefix": ch_pref,
                "is_multi_article": True,
                "body_range": [b_start, b_end],
                "section_path": section_path or book_label,
                "cn_title": cn_title,
                "qualified_cn_title": qualified_cn_title,
                "en_title": en_t,
                "row_count": u["n_rows"],
                "first_body_cn": u["first_body_cn"],
                "aliases": aliases,
            })

    return rel_posix, {
        "rel_path": rel_posix,
        "mtime_ns": mtime_ns,
        "size": size,
        "articles": articles,
    }


def _load_or_build_master_articles_with_stats(
    cn_en_master_dir: Union[str, Path],
    log_fn: Optional[Callable[[str], None]] = None,
    max_workers: int = 8,
    force_rebuild: bool = False,
) -> Tuple[List[dict], dict]:
    t0 = time.perf_counter()
    master_p = Path(cn_en_master_dir).resolve()
    if not master_p.exists():
        raise FileNotFoundError(f"找不到最新中英底稿目录或文件：{master_p}")

    norm_key = os.path.normcase(str(master_p))
    disk_stats = _scan_master_docx_stats(master_p)
    stat_sig = tuple((rp, mt, sz) for rp, _, mt, sz in disk_stats)

    with _MASTER_CACHE_LOCK:
        cache_path = _resolve_master_cache_path(master_p)
        if force_rebuild:
            _MEM_MASTER_CACHE.pop(norm_key, None)
        else:
            mem_hit = _MEM_MASTER_CACHE.get(norm_key)
            if mem_hit is not None and mem_hit[0] == stat_sig and cache_path.exists():
                arts_cached = mem_hit[1]
                return arts_cached, {
                    "total_files": len(disk_stats),
                    "total_articles": len(arts_cached),
                    "rebuilt_files": 0,
                    "cache_path": str(cache_path),
                    "elapsed_sec": round(time.perf_counter() - t0, 4),
                }

        cached_files: Dict[str, dict] = {}
        if not force_rebuild and cache_path.exists():
            try:
                raw = json.loads(cache_path.read_text(encoding="utf-8"))
                if (
                    isinstance(raw, dict)
                    and raw.get("version") == MASTER_INDEX_SCHEMA_VERSION
                    and isinstance(raw.get("files"), dict)
                ):
                    cached_files = raw["files"]
            except Exception:
                cached_files = {}

        next_files: Dict[str, dict] = {}
        to_parse = []
        for rp, abs_s, mt, sz in disk_stats:
            prev = None if force_rebuild else cached_files.get(rp)
            if (
                prev is not None
                and prev.get("mtime_ns") == mt
                and prev.get("size") == sz
                and isinstance(prev.get("articles"), list)
            ):
                next_files[rp] = prev
            else:
                to_parse.append((rp, abs_s, mt, sz))

        rebuilt_cnt = len(to_parse)
        deleted_cnt = sum(1 for old_k in cached_files if old_k not in {rp for rp, _, _, _ in disk_stats})

        if to_parse:
            workers = max(1, min(int(max_workers or 8), len(to_parse)))
            if log_fn and (force_rebuild or len(to_parse) >= 3):
                mode_lbl = "强制全量刷新中英底稿索引" if force_rebuild else "多线程扫描中英底稿目录"
                log_fn(f"📖 正在{mode_lbl}（解析 {len(to_parse)}/{len(disk_stats)} 个文档）...")
            if workers == 1:
                parsed = [_parse_single_master_docx(t) for t in to_parse]
            else:
                with ThreadPoolExecutor(max_workers=workers) as ex:
                    parsed = list(ex.map(_parse_single_master_docx, to_parse))
            for rp, rec in parsed:
                next_files[rp] = rec

        if force_rebuild or to_parse or deleted_cnt > 0 or not cache_path.exists():
            try:
                payload = {
                    "version": MASTER_INDEX_SCHEMA_VERSION,
                    "master_dir": str(master_p),
                    "updated_at": time.time(),
                    "files": next_files,
                }
                tmp_p = cache_path.with_name(f"{cache_path.name}.tmp_{os.getpid()}_{threading.get_ident()}")
                tmp_p.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
                os.replace(tmp_p, cache_path)
            except OSError:
                pass

        all_articles: List[dict] = []
        for rp in sorted(next_files.keys(), key=_natural_sort_key):
            all_articles.extend(next_files[rp].get("articles") or [])

        _MEM_MASTER_CACHE[norm_key] = (stat_sig, all_articles)
        elapsed = time.perf_counter() - t0
        if log_fn:
            log_fn(f"📖 中英底稿索引就绪：共识别 {len(all_articles)} 篇可提取章节（耗时 {elapsed:.2f}s）")
        return all_articles, {
            "total_files": len(disk_stats),
            "total_articles": len(all_articles),
            "rebuilt_files": rebuilt_cnt,
            "cache_path": str(cache_path),
            "elapsed_sec": round(elapsed, 4),
        }


def _load_or_build_master_articles(
    cn_en_master_dir: Union[str, Path],
    log_fn: Optional[Callable[[str], None]] = None,
    max_workers: int = 8,
    force_rebuild: bool = False,
) -> List[dict]:
    arts, _ = _load_or_build_master_articles_with_stats(
        cn_en_master_dir=cn_en_master_dir,
        log_fn=log_fn,
        max_workers=max_workers,
        force_rebuild=force_rebuild,
    )
    return arts


def refresh_master_index(
    master_path: Union[str, Path],
    log_fn: Optional[Callable[[str], None]] = None,
    force_rebuild: bool = False,
    max_workers: int = 8,
) -> dict:
    """
    Public helper for GUI `刷新底稿` button & live refresh checks:
    verifies live `os.stat` `(mtime_ns, size)` across `master_path` (or forces a 100% rebuild
    when `force_rebuild=True`) and returns summary stats dict.
    """
    _, stats = _load_or_build_master_articles_with_stats(
        cn_en_master_dir=master_path,
        log_fn=log_fn,
        max_workers=max_workers,
        force_rebuild=bool(force_rebuild),
    )
    return stats


# ---------------------------------------------------------------------------
# Article Resolution (`resolve_full_articles` & `resolve_direct_folder_articles`)
# ---------------------------------------------------------------------------

def _qualifier_synonyms(qual: str) -> List[str]:
    qn = _norm_compact(qual)
    syns = [qn]
    if "起初" in qn or "基督起初的发表" in qn:
        syns.extend(["起初", "基督起初的发表", "第一部分"])
    if "全宇" in qn or "神向全宇的说话" in qn:
        syns.extend(["全宇", "神向全宇的说话", "第二部分"])
    if "揭示" in qn or "奥秘" in qn:
        syns.extend(["揭示", "奥秘揭示", "对神奥秘的揭示", "神向全宇说话的奥秘揭示"])
    if "行走" in qn:
        syns.extend(["行走", "基督在众教会行走时的说话", "第三部分"])
    return list(dict.fromkeys(s for s in syns if s))


def _safe_substring_hit(needle_norm: str, target_norm: str) -> bool:
    if not needle_norm or not target_norm:
        return False
    pos = target_norm.find(needle_norm)
    while pos != -1:
        end_pos = pos + len(needle_norm)
        if needle_norm[-1] in CN_NUM_CHARS and end_pos < len(target_norm) and target_norm[end_pos] in CN_NUM_CHARS:
            pos = target_norm.find(needle_norm, pos + 1)
            continue
        if needle_norm.startswith("第") and pos > 0 and target_norm[pos - 1] in (CN_NUM_CHARS + "附加"):
            pos = target_norm.find(needle_norm, pos + 1)
            continue
        return True
    return False


def _match_single_query(query: str, all_articles: List[dict]) -> Tuple[str, List[dict]]:
    """Return `(query_prefix_code, matched_articles)` for a single query line."""
    clean_q = _clean_single_title_line(query)
    if not clean_q:
        return "", []

    parts = [p.strip().strip("《》") for p in re.split(r"[・·>›/｜|]+", clean_q) if p.strip()]
    if len(parts) == 1 and re.search(r"^(?:起初|全宇|揭示|基督起初的发表|神向全宇的说话|对神奥秘的揭示)[\s_]+(.+)$", clean_q):
        m_sp = re.match(r"^(起初|全宇|揭示|基督起初的发表|神向全宇的说话|对神奥秘的揭示)[\s_]+(.+)$", clean_q)
        if m_sp:
            parts = [m_sp.group(1).strip(), m_sp.group(2).strip()]

    leaf = parts[-1] if parts else clean_q
    qualifiers = parts[:-1] if len(parts) >= 2 else []
    q_prefix, leaf_core = split_title_prefix_and_core(leaf, strip_lang_prefix=False)
    if not q_prefix:
        q_prefix, _ = split_title_prefix_and_core(clean_q, strip_lang_prefix=False)

    pool = all_articles
    if qualifiers:
        for qual in qualifiers:
            syns = _qualifier_synonyms(qual)
            filtered = [
                art for art in pool
                if any(
                    syn in _norm_compact(
                        f"{art.get('master_rel_path', '')} {art.get('section_path', '')} {art.get('qualified_cn_title', '')}"
                    )
                    for syn in syns
                )
            ]
            if filtered:
                pool = filtered

    q_full_compact = _norm_compact(clean_q)
    q_full_loose = _norm_loose(clean_q)
    leaf_compact = _norm_compact(leaf)
    leaf_loose = _norm_loose(leaf)
    leaf_core_compact = _norm_compact(leaf_core)
    leaf_core_loose = _norm_loose(leaf_core)

    # Tier 1: Exact match on aliases / stripped core aliases / cn_title / en_title / stem
    exact_hits: List[dict] = []
    for art in pool:
        alias_compacts = set()
        alias_looses = set()
        for a in art.get("aliases", ()):
            alias_compacts.add(_norm_compact(a))
            alias_looses.add(_norm_loose(a))
            _, a_core = split_title_prefix_and_core(a, strip_lang_prefix=True)
            if a_core:
                alias_compacts.add(_norm_compact(a_core))
                alias_looses.add(_norm_loose(a_core))
        en_compact = _norm_compact(_strip_title_suffixes_and_punct(art.get("en_title", "")))
        en_raw_compact = _norm_compact(art.get("en_title", ""))

        if (
            q_full_compact in alias_compacts
            or leaf_compact in alias_compacts
            or (leaf_core_compact and leaf_core_compact in alias_compacts)
            or (q_full_loose and len(q_full_loose) >= 4 and q_full_loose in alias_looses)
            or (leaf_loose and len(leaf_loose) >= 4 and leaf_loose in alias_looses)
            or (leaf_core_loose and len(leaf_core_loose) >= 4 and leaf_core_loose in alias_looses)
            or (en_raw_compact and q_full_compact in (en_raw_compact, en_compact))
        ):
            exact_hits.append(art)

    if exact_hits:
        return q_prefix, exact_hits

    # Tier 2: Safe substring / containment match on aliases and stripped core titles
    sub_hits: List[dict] = []
    for art in pool:
        hay_list = []
        for a in art.get("aliases", ()):
            hay_list.append(_norm_compact(a))
            _, a_core = split_title_prefix_and_core(a, strip_lang_prefix=True)
            if a_core:
                hay_list.append(_norm_compact(a_core))
        if art.get("en_title"):
            hay_list.append(_norm_compact(art["en_title"]))

        if any(_safe_substring_hit(leaf_compact, h) for h in hay_list):
            sub_hits.append(art)
        elif leaf_core_compact and len(leaf_core_compact) >= 3 and any(
            _safe_substring_hit(leaf_core_compact, h) for h in hay_list
        ):
            sub_hits.append(art)

    return q_prefix, sub_hits


def resolve_full_articles(
    titles: Sequence[str],
    cn_en_master_dir: Union[str, Path],
    foreign_lib_dir: Union[str, Path] = "",
    log_fn: Optional[Callable[[str], None]] = None,
) -> dict:
    """
    Resolve user-supplied chapter titles against `cn_en_master_dir` (supporting both single-article
    `.docx` files and multi-article book compilation `.docx` files), preserving any numeric prefix
    codes (e.g. `3-1-39 032`) in `cand["query_prefix"]`.
    """
    clean_titles: List[str] = []
    seen_q = set()
    for raw_t in titles or []:
        for t in parse_titles_multiline(str(raw_t)):
            if t not in seen_q:
                seen_q.add(t)
                clean_titles.append(t)

    all_articles = _load_or_build_master_articles(cn_en_master_dir, log_fn=log_fn)

    if foreign_lib_dir and str(foreign_lib_dir).strip():
        flib_p = Path(foreign_lib_dir).resolve()
        if flib_p.is_dir():
            ensure_library_index(flib_p, log_fn=log_fn)

    resolved: List[dict] = []
    ambiguous: List[dict] = []
    unmatched: List[str] = []
    seen_cand_ids = set()

    for q in clean_titles:
        q_prefix, hits = _match_single_query(q, all_articles)
        if not hits:
            unmatched.append(q)
            if log_fn:
                log_fn(f"  ⚠️ 未在中英底稿目录找到匹配篇目：《{q}》")
        elif len(hits) == 1:
            cand = dict(hits[0])
            cand["query"] = q
            cand["query_prefix"] = q_prefix
            cand["direct_folder_mode"] = False
            if cand["cand_id"] not in seen_cand_ids:
                seen_cand_ids.add(cand["cand_id"])
                resolved.append(cand)
            if log_fn:
                log_fn(f"  ✅ 精准匹配：《{q}》 -> {cand['display_label']}")
        else:
            cand_copies = []
            for h in hits:
                cp = dict(h)
                cp["query"] = q
                cp["query_prefix"] = q_prefix
                cp["direct_folder_mode"] = False
                cand_copies.append(cp)
            ambiguous.append({
                "query": q,
                "candidates": cand_copies,
            })
            if log_fn:
                log_fn(f"  ❓ 存在 {len(cand_copies)} 篇同名/近似章节需确认：《{q}》")

    return {
        "resolved": resolved,
        "ambiguous": ambiguous,
        "unmatched": unmatched,
    }


def resolve_direct_folder_articles(
    cn_en_master_dir: Union[str, Path],
    foreign_lib_dir: Union[str, Path],
    log_fn: Optional[Callable[[str], None]] = None,
) -> dict:
    """
    Item #7① — Direct Folder-to-Folder Mode:
      1. Lists all `.docx` articles in `cn_en_master_dir` in natural sorted order (`1-01...`, `3-1-35...`);
      2. Matches each master `.docx` with its corresponding `.docx` in `foreign_lib_dir`
         (by stripped core title, chapter code, and first body Chinese paragraph);
      3. Preserves `"naming_stem": master_file.stem` so outputs are named directly after the
         `最新中英底稿` filename: `{LANG}_{master_stem}_{COUNT_TAG}_{DATE}.docx`.
    Returns `{"resolved": list[dict], "ambiguous": [], "unmatched": list[str]}`.
    """
    master_p = Path(cn_en_master_dir).resolve()
    flib_p = Path(foreign_lib_dir).resolve()
    if not master_p.exists():
        raise FileNotFoundError(f"找不到最新中英底稿目录：{master_p}")
    if not flib_p.exists() or not flib_p.is_dir():
        raise FileNotFoundError(f"找不到外文三语目录：{flib_p}")

    all_articles = _load_or_build_master_articles(master_p, log_fn=log_fn)
    all_articles = sorted(
        all_articles,
        key=lambda a: (_natural_sort_key(str(a.get("master_rel_path") or "")), (a.get("body_range") or [0, 0])[0]),
    )

    lib_idx_mod = importlib.import_module("library_index")
    lib_index = lib_idx_mod.get_library_index(flib_p, log_fn=log_fn)

    # Pre-compute normalized core titles & chapter code tuples for every foreign library file
    foreign_meta: Dict[str, dict] = {}
    for rel_p, f_info in lib_index.files_map.items():
        f_stem = Path(rel_p).stem
        f_pref, f_core = split_title_prefix_and_core(f_stem, strip_lang_prefix=True)
        core_norms = set()
        if f_core:
            core_norms.add(_norm_compact(f_core))
        for t_str in f_info.get("titles") or ():
            _, t_core = split_title_prefix_and_core(t_str, strip_lang_prefix=True)
            if t_core:
                core_norms.add(_norm_compact(t_core))
        foreign_meta[rel_p] = {
            "code_tuple": _normalize_chapter_code_tuple(f_pref),
            "core_norms": core_norms,
            "priority": lib_index.file_priority_key(rel_p),
        }

    resolved: List[dict] = []
    unmatched: List[str] = []

    for art in all_articles:
        m_file = Path(art["master_file"])
        master_stem = m_file.stem if not art.get("is_multi_article") else str(art.get("qualified_cn_title") or art.get("cn_title") or m_file.stem)
        m_pref, m_core = split_title_prefix_and_core(master_stem, strip_lang_prefix=True)
        _, cn_core = split_title_prefix_and_core(str(art.get("cn_title") or ""), strip_lang_prefix=True)
        m_code_tuple = _normalize_chapter_code_tuple(m_pref)

        m_cores = {_norm_compact(x) for x in (m_core, cn_core) if x and _norm_compact(x)}
        first_body = str(art.get("first_body_cn") or "").strip()

        # Score candidate foreign files
        cand_scores: Dict[str, float] = {}
        for rel_p, fmeta in foreign_meta.items():
            sc = 0.0
            f_cores = fmeta["core_norms"]
            if m_cores and (m_cores & f_cores):
                sc += 100.0
            elif m_cores:
                for mc in m_cores:
                    if len(mc) >= 3 and any(_safe_substring_hit(mc, fc) or _safe_substring_hit(fc, mc) for fc in f_cores if len(fc) >= 3):
                        sc += 65.0
                        break
            if m_code_tuple is not None and fmeta["code_tuple"] == m_code_tuple:
                sc += 45.0
            if sc > 0:
                cand_scores[rel_p] = sc

        # First body Chinese paragraph verification / fallback
        if first_body:
            fb_norm = lib_idx_mod.normalize_cn(first_body)
            exact_rows = lib_index.exact_cn_map.get(fb_norm) or []
            for r_hit in exact_rows:
                rp = r_hit["rel_path"]
                cand_scores[rp] = cand_scores.get(rp, 0.0) + 110.0
            if not exact_rows:
                tri_cands = lib_index.retrieve_trigram_candidates(first_body, top_k=15)
                for r_hit in tri_cands:
                    sim = step1_extract_and_diff.best_local_cn_similarity(first_body, r_hit["cn_norm"])
                    if sim >= 0.78:
                        rp = r_hit["rel_path"]
                        cand_scores[rp] = max(cand_scores.get(rp, 0.0), cand_scores.get(rp, 0.0) + 90.0 * sim)

        if not cand_scores:
            unmatched.append(master_stem)
            if log_fn:
                log_fn(f"  ⚠️ 未在外文目录找到对应篇目：《{master_stem}》")
            continue

        best_rel, best_sc = max(
            cand_scores.items(),
            key=lambda kv: (kv[1], foreign_meta[kv[0]]["priority"]),
        )
        if best_sc < 55.0:
            unmatched.append(master_stem)
            if log_fn:
                log_fn(f"  ⚠️ 未在外文目录找到高置信度对应篇目：《{master_stem}》")
            continue

        cp = dict(art)
        cp["naming_stem"] = master_stem
        cp["direct_folder_mode"] = True
        cp["query"] = master_stem
        cp["query_prefix"] = ""
        cp["matched_foreign_file"] = str(lib_index.resolve_file_path(best_rel))
        cp["matched_foreign_rel"] = best_rel
        resolved.append(cp)
        if log_fn:
            log_fn(f"  ✅ 顺序直连匹配：《{master_stem}》 -> {best_rel}")

    return {
        "resolved": resolved,
        "ambiguous": [],
        "unmatched": unmatched,
    }


# ---------------------------------------------------------------------------
# Full-Article Extraction, Zero-Change Report & Smart Subfolder Export
# ---------------------------------------------------------------------------

def _strip_master_shading_and_highlights(root: ET.Element) -> None:
    """
    Accept any unaccepted revisions and remove pre-existing `w:shd` and `w:highlight` from master
    table cells (`【标you】` annotations) so Step 1 compares 100% of sentences across the entire chapter.
    """
    for tbl in root.findall(".//w:tbl", NS):
        for tc in tbl.findall(".//w:tc", NS):
            _accept_cell_revisions_inplace(tc)
            tcPr = tc.find("w:tcPr", NS)
            if tcPr is not None:
                docx_xml_utils.remove_rpr_child(tcPr, "shd")
            for p in tc.findall("w:p", NS):
                pPr = p.find("w:pPr", NS)
                if pPr is not None:
                    docx_xml_utils.remove_rpr_child(pPr, "shd")
                    p_rPr = pPr.find("w:rPr", NS)
                    if p_rPr is not None:
                        docx_xml_utils.remove_rpr_child(p_rPr, "shd")
                        docx_xml_utils.remove_rpr_child(p_rPr, "highlight")
                for r in p.findall(".//w:r", NS):
                    rPr = r.find("w:rPr", NS)
                    if rPr is not None:
                        docx_xml_utils.remove_rpr_child(rPr, "shd")
                        docx_xml_utils.remove_rpr_child(rPr, "highlight")


def _export_master_article_backup_docx(article: dict, dst_docx_path: Path) -> Path:
    """
    Export processed article's authentic `中英最新底稿` `.docx`:
      - Single-article master `.docx`: direct file copy preserving 100% original bytes & formatting;
      - Multi-article compilation `.docx`: slices `body_range` while preserving master shading/colors.
    """
    master_p = Path(article["master_file"]).resolve()
    dst_docx_path.parent.mkdir(parents=True, exist_ok=True)
    body_range = article.get("body_range")
    if not article.get("is_multi_article") or not isinstance(body_range, (list, tuple)) or len(body_range) != 2:
        docx_xml_utils.slim_docx_file(master_p, dst_docx_path)
        return dst_docx_path

    root, _, ns_pairs = docx_xml_utils.load_docx_xml(master_p)
    body = root.find("w:body", NS)
    if body is not None:
        b_start, b_end = int(body_range[0]), int(body_range[1])
        children = list(body)
        sect_el = children[-1] if (children and children[-1].tag.split("}")[-1] == "sectPr") else None
        kept = children[b_start:b_end]
        for ch in children:
            if ch not in kept and ch is not sect_el:
                body.remove(ch)
    docx_xml_utils.save_docx_xml(master_p, dst_docx_path, root, ns_pairs, comments_root=None)
    return dst_docx_path


def _extract_standalone_article_docx(
    article: dict,
    dst_docx_path: Path,
) -> Tuple[str, Optional[Tuple[int, str]]]:
    """
    Extract `article` from `article["master_file"]` into a standalone 3-column Landscape `.docx` at `dst_docx_path`.
    Returns `(extracted_en_title, restored_heading_info_or_None)`.
    """
    master_p = Path(article["master_file"]).resolve()
    if not master_p.is_file():
        raise FileNotFoundError(f"找不到中英底稿文件：{master_p}")

    root, _, ns_pairs = docx_xml_utils.load_docx_xml(master_p)
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError(f"文档结构无效（缺少 w:body）：{master_p}")

    body_range = article.get("body_range")
    if article.get("is_multi_article") and isinstance(body_range, (list, tuple)) and len(body_range) == 2:
        b_start, b_end = int(body_range[0]), int(body_range[1])
        children = list(body)
        sect_el = children[-1] if (children and children[-1].tag.split("}")[-1] == "sectPr") else None
        kept = children[b_start:b_end]
        for ch in children:
            if ch not in kept and ch is not sect_el:
                body.remove(ch)

    _strip_master_shading_and_highlights(root)
    ensure_three_column_tables_and_landscape(root)

    en_title = (
        extract_full_article_en_title(root)
        or str(article.get("en_title") or "").strip()
        or str(article.get("cn_title") or "").strip()
        or "Article"
    )

    restore_info: Optional[Tuple[int, str]] = None
    qual_cn = str(article.get("qualified_cn_title") or "").strip()
    raw_cn = str(article.get("cn_title") or "").strip()
    if qual_cn and raw_cn and qual_cn != raw_cn:
        for idx_b, ch in enumerate(body):
            if ch.tag.split("}")[-1] == "p":
                ptxt = "".join(t.text or "" for t in ch.findall(".//w:t", NS)).strip()
                if ptxt == raw_cn:
                    t_nodes = ch.findall(".//w:t", NS)
                    if t_nodes:
                        restore_info = (idx_b, ptxt)
                        t_nodes[0].text = qual_cn
                        for extra_t in t_nodes[1:]:
                            extra_t.text = ""
                    break
            elif ch.tag.split("}")[-1] == "tbl":
                break

    docx_xml_utils.save_docx_xml(master_p, dst_docx_path, root, ns_pairs, comments_root=None)
    return en_title, restore_info


def _restore_heading_in_docx(docx_path: Path, restore_heading: Optional[Tuple[int, str]]) -> None:
    if restore_heading is None or not docx_path.is_file():
        return
    p_idx_restore, orig_ptxt = restore_heading
    d_root, _, d_ns = docx_xml_utils.load_docx_xml(docx_path)
    d_cmts = docx_xml_utils.load_docx_comments_xml(docx_path)
    d_body = d_root.find("w:body", NS)
    if d_body is not None and p_idx_restore < len(d_body):
        t_els = d_body[p_idx_restore].findall(".//w:t", NS)
        if t_els:
            t_els[0].text = orig_ptxt
            for et in t_els[1:]:
                et.text = ""
            docx_xml_utils.save_docx_xml(docx_path, docx_path, d_root, d_ns, comments_root=d_cmts)


def _repair_skipped_title_rows_in_manifest(
    manifest_path: Path,
    review_items_path: Path,
) -> int:
    """
    If Row 0 (e.g. short 2-char chapter title like '后记' with changed English title) was skipped by
    `step1_extract_and_diff`, recover it from Row 0 of the same table's matched `source_file`!
    """
    with open(manifest_path, "r", encoding="utf-8") as f:
        manifest = json.load(f)
    with open(review_items_path, "r", encoding="utf-8") as f:
        review_items = json.load(f)

    skipped = manifest.get("skipped_rows") or []
    rows_manifest = manifest.get("rows") or []
    if not skipped or not rows_manifest:
        return 0

    tbl_to_src: Dict[int, Path] = {}
    tbl_to_idx: Dict[int, int] = {}
    for rm in rows_manifest:
        b_idx = rm["body_idx"]
        if b_idx not in tbl_to_src and rm.get("source_file"):
            tbl_to_src[b_idx] = Path(rm["source_file"])
            tbl_to_idx[b_idx] = rm.get("tbl_idx", 0)

    if not tbl_to_src:
        return 0

    bilingual_2col = bool(manifest.get("bilingual_2col"))
    min_src_cols = 2 if bilingual_2col else 3
    tgt_src_idx = 1 if bilingual_2col else 2
    inp_root, _, _ = docx_xml_utils.load_docx_xml(Path(manifest["input_docx"]))
    inp_body = inp_root.find("w:body", NS)
    cfg = docx_xml_utils.load_config()
    rv_threshold = cfg.get("rv_word_change_ratio_threshold", 0.42)

    repaired = 0
    new_skipped = []
    src_xml_cache: Dict[Path, List[ET.Element]] = {}

    for sk in skipped:
        b_idx = sk["body_idx"]
        r_idx = sk["row_idx"]
        if r_idx in (0, 1) and b_idx in tbl_to_src:
            src_f = tbl_to_src[b_idx]
            if src_f.is_file():
                if src_f not in src_xml_cache:
                    s_root, _, _ = docx_xml_utils.load_docx_xml(src_f)
                    src_xml_cache[src_f] = s_root.findall(".//w:tbl/w:tr", NS)
                s_trs = src_xml_cache[src_f]
                inp_trs = inp_body[b_idx].findall("w:tr", NS)
                if r_idx < len(inp_trs) and r_idx < len(s_trs):
                    inp_tcs = inp_trs[r_idx].findall("w:tc", NS)
                    s_tcs = s_trs[r_idx].findall("w:tc", NS)
                    if len(inp_tcs) >= 2 and len(s_tcs) >= min_src_cols and docx_xml_utils.get_cell_plain_text(s_tcs[tgt_src_idx]).strip():
                        _, inp_cn_core = split_title_prefix_and_core(docx_xml_utils.get_cell_plain_text(inp_tcs[0]))
                        _, src_cn_core = split_title_prefix_and_core(docx_xml_utils.get_cell_plain_text(s_tcs[0]))
                        inp_cn = _WS_RE.sub("", inp_cn_core or docx_xml_utils.get_cell_plain_text(inp_tcs[0]))
                        src_cn = _WS_RE.sub("", src_cn_core or docx_xml_utils.get_cell_plain_text(s_tcs[0]))
                        same_appendix_r0 = (r_idx == 0 and inp_cn.startswith("附篇") and src_cn.startswith("附篇"))
                        if inp_cn and (inp_cn == src_cn or inp_cn in src_cn or src_cn in inp_cn or same_appendix_r0):
                            if bilingual_2col:
                                inp_txt0 = docx_xml_utils.get_cell_plain_text(inp_tcs[0])
                                src_txt0 = docx_xml_utils.get_cell_plain_text(s_tcs[0])
                                ops = docx_xml_utils.diff_cn_tokens(src_txt0, inp_txt0)
                                sent_op_lists = docx_xml_utils.segment_cn_diff_into_sentences(ops)
                                stats_fn = docx_xml_utils.cn_sentence_unit_stats
                            else:
                                inp_en = docx_xml_utils.get_cell_plain_text(inp_tcs[1])
                                src_en = docx_xml_utils.get_cell_plain_text(s_tcs[1])
                                ops = docx_xml_utils.diff_tokens(src_en, inp_en)
                                sent_op_lists = docx_xml_utils.segment_diff_into_sentences(ops)
                                stats_fn = docx_xml_utils.sentence_unit_stats
                            tbl_cnt = tbl_to_idx[b_idx]
                            sent_items = []
                            for s_i, s_ops in enumerate(sent_op_lists):
                                stats = stats_fn(s_ops)
                                sid = f"T{tbl_cnt}_R{r_idx}_G0_S{s_i}"
                                if not stats["has_edit"]:
                                    sug = "none"
                                elif stats["change_ratio"] >= rv_threshold:
                                    sug = "cyan"
                                else:
                                    sug = "yellow"
                                sent_items.append({
                                    "sent_id": sid,
                                    "suggested_action": sug,
                                    **stats,
                                    "ops": s_ops,
                                })
                                if stats["has_edit"]:
                                    review_items.insert(
                                        0,
                                        {
                                            "sent_id": sid,
                                            "article": sk.get("article_title", ""),
                                            "row_idx": r_idx,
                                            "has_edit": True,
                                            "suggested": sug,
                                            "change_ratio": stats["change_ratio"],
                                            "edit_hunks": stats["edit_hunks"],
                                            "new_words": stats["new_words"],
                                            "diff_preview": stats["diff_preview"],
                                        },
                                    )
                            rows_manifest.insert(
                                0,
                                {
                                    "tbl_idx": tbl_cnt,
                                    "body_idx": b_idx,
                                    "row_idx": r_idx,
                                    "article_title": sk.get("article_title", ""),
                                    "source_file": str(src_f),
                                    "source_row_idx": r_idx,
                                    "match_confidence": 1.0,
                                    "match_stage": "chapter_title_row0",
                                    "col3_preview": docx_xml_utils.get_cell_plain_text(s_tcs[tgt_src_idx])[:120],
                                    "regions": [
                                        {
                                            "region_idx": 0,
                                            "is_quoted": True,
                                            "fill_hex": None,
                                            "sentences": sent_items,
                                        }
                                    ],
                                },
                            )
                            repaired += 1
                            continue
        new_skipped.append(sk)

    if repaired > 0:
        rows_manifest.sort(key=lambda r: (r["body_idx"], r["row_idx"]))
        manifest["rows"] = rows_manifest
        manifest["skipped_rows"] = new_skipped
        with open(manifest_path, "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)
        with open(review_items_path, "w", encoding="utf-8") as f:
            json.dump(review_items, f, ensure_ascii=False, indent=2)

    return repaired


def _manifest_has_any_en_edits(manifest: dict, review_items: list) -> bool:
    if any(x.get("has_edit") for x in (review_items or [])):
        return True
    for rm in manifest.get("rows") or []:
        for reg in rm.get("regions") or []:
            if not reg.get("is_quoted"):
                continue
            for sent in reg.get("sentences") or []:
                if sent.get("has_edit"):
                    return True
                for k_op, t_op in sent.get("ops") or []:
                    if (
                        k_op in ("del", "ins")
                        and str(t_op).strip()
                        and not docx_xml_utils.is_pure_pronoun_tag_op(k_op, str(t_op))
                    ):
                        return True
    return False


def _docx_has_any_col2_revisions(docx_path: Path) -> bool:
    if not docx_path.is_file():
        return False
    root, _, _ = docx_xml_utils.load_docx_xml(docx_path)
    for tr in root.findall(".//w:tbl/w:tr", NS):
        if step2_render_final.is_skipped_row(tr):
            continue
        tcs = tr.findall("w:tc", NS)
        if len(tcs) >= 2:
            tc_diff = tcs[0] if len(tcs) == 2 else tcs[1]
            if docx_xml_utils.cell_has_substantive_revisions(tc_diff):
                return True
    return False


def _row_has_up_or_rv_changes(tr: ET.Element, up_color: str, rv_color: str) -> bool:
    if step2_render_final.is_skipped_row(tr):
        return False
    tcs = tr.findall("w:tc", NS)
    if len(tcs) < 2:
        return False
    tc_diff = tcs[0] if len(tcs) == 2 else tcs[1]
    if docx_xml_utils.cell_has_substantive_revisions(tc_diff):
        return True
    active_hls = {up_color, rv_color, "cyan", "blue", "darkCyan"}
    for hl in tc_diff.findall(".//w:rPr/w:highlight", NS):
        if hl.attrib.get(f"{W}val") in active_hls:
            return True
    return False


def _count_changed_body_rows_in_root(root: ET.Element, up_color: str, rv_color: str) -> int:
    cnt = 0
    for tbl in root.findall(".//w:tbl", NS):
        trs = tbl.findall("w:tr", NS)
        for r_i, tr in enumerate(trs):
            if r_i == 0 and len(trs) >= 2:
                # Skip counting unedited Row 0 header; count if Row 0 itself was edited
                if _row_has_up_or_rv_changes(tr, up_color, rv_color):
                    cnt += 1
                continue
            if _row_has_up_or_rv_changes(tr, up_color, rv_color):
                cnt += 1
    return cnt


def _filter_full_article_to_translator_version(
    root: ET.Element,
    up_color: str,
    rv_color: str,
) -> int:
    """
    Transform a rendered full-article `root` into the `仅改点发译者版`:
      - Inside single-table chapters: keeps Row 0 bold chapter title row (plus Row 1 subtitle if `附篇一..四`)
        AND ONLY rows that have UP/RV changes in Col 2!
      - Inside multi-table `选段` articles: keeps main chapter title table + only `选段` tables with changes.
    Returns count of kept changed rows.
    """
    body = root.find("w:body", NS)
    if body is None:
        return 0

    tbls = body.findall("w:tbl", NS)
    if not tbls:
        return 0

    def _is_blank_row(tr: ET.Element) -> bool:
        tcs = tr.findall("w:tc", NS)
        return bool(tcs) and all(not docx_xml_utils.get_cell_plain_text(c).strip() for c in tcs)

    def _title_row_count(trs: List[ET.Element]) -> int:
        if not trs:
            return 0
        if docx_xml_utils.is_two_row_chapter_header(trs):
            return 3 if (len(trs) >= 3 and _is_blank_row(trs[2])) else 2
        tcs0 = trs[0].findall("w:tc", NS)
        cn0 = docx_xml_utils.get_cell_plain_text(tcs0[0]).strip() if tcs0 else ""
        if len(trs) >= 2 and _is_blank_row(trs[1]):
            return 2
        if len(cn0) <= 60 and (
            (tcs0 and tcs0[0].find(".//w:b", NS) is not None)
            or (len(tcs0) >= 2 and tcs0[1].find(".//w:b", NS) is not None)
        ):
            return 1
        return 0

    kept_changed_rows = 0
    emptied_sub_tbls = set()

    for t_idx, tbl in enumerate(tbls):
        trs = tbl.findall("w:tr", NS)
        if not trs:
            continue
        hdr_cnt = _title_row_count(trs)
        if t_idx == 0 and hdr_cnt == 0 and len(tbls) == 1:
            hdr_cnt = 1

        tbl_changed = 0
        for r_i, tr in enumerate(list(trs)):
            if r_i < hdr_cnt:
                if _row_has_up_or_rv_changes(tr, up_color, rv_color):
                    kept_changed_rows += 1
                continue
            if _row_has_up_or_rv_changes(tr, up_color, rv_color):
                tbl_changed += 1
                kept_changed_rows += 1
            else:
                tbl.remove(tr)

        if len(tbls) > 1 and t_idx >= 1 and tbl_changed == 0:
            emptied_sub_tbls.add(id(tbl))

        if id(tbl) not in emptied_sub_tbls and not tbl.findall("w:tr", NS) and trs:
            tbl.append(trs[0])

    if emptied_sub_tbls:
        children = list(body)
        to_rem = set()
        for idx_c, ch in enumerate(children):
            if id(ch) in emptied_sub_tbls:
                to_rem.add(idx_c)
                b_i = idx_c - 1
                while b_i >= 0 and children[b_i].tag.split("}")[-1] == "p":
                    ptxt = "".join(t.text or "" for t in children[b_i].findall(".//w:t", NS)).strip()
                    if not ptxt or not is_top_highlight_instruction_text(ptxt):
                        to_rem.add(b_i)
                    if ptxt:
                        break
                    b_i -= 1
        for idx_c in sorted(to_rem, reverse=True):
            body.remove(children[idx_c])

    return kept_changed_rows


# ---------------------------------------------------------------------------
# Standalone OOXML Builder for Zero-Change Summary Report (Item #7②)
# ---------------------------------------------------------------------------

def _make_report_para(
    text: str,
    bold: bool = False,
    sz_half_pt: int = 22,
    color_hex: str = "1F2937",
    jc: str = "left",
    space_before: int = 60,
    space_after: int = 60,
) -> ET.Element:
    p = ET.Element(f"{W}p")
    pPr = ET.SubElement(p, f"{W}pPr")
    ET.SubElement(pPr, f"{W}spacing", {f"{W}before": str(space_before), f"{W}after": str(space_after)})
    if jc != "left":
        ET.SubElement(pPr, f"{W}jc", {f"{W}val": jc})
    r = ET.SubElement(p, f"{W}r")
    rPr = ET.SubElement(r, f"{W}rPr")
    ET.SubElement(
        rPr,
        f"{W}rFonts",
        {f"{W}ascii": "Georgia", f"{W}eastAsia": "微软雅黑", f"{W}hAnsi": "Georgia", f"{W}cs": "Georgia"},
    )
    if bold:
        ET.SubElement(rPr, f"{W}b")
        ET.SubElement(rPr, f"{W}bCs")
    if color_hex:
        ET.SubElement(rPr, f"{W}color", {f"{W}val": color_hex})
    ET.SubElement(rPr, f"{W}sz", {f"{W}val": str(sz_half_pt)})
    ET.SubElement(rPr, f"{W}szCs", {f"{W}val": str(sz_half_pt)})
    t = ET.SubElement(r, f"{W}t")
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        t.set(XML_SPACE, "preserve")
    t.text = text
    return p


def _make_report_cell(text: str, width_pct: str, bold: bool = False, shd_fill: Optional[str] = None) -> ET.Element:
    tc = ET.Element(f"{W}tc")
    tcPr = ET.SubElement(tc, f"{W}tcPr")
    ET.SubElement(tcPr, f"{W}tcW", {f"{W}w": width_pct, f"{W}type": "pct"})
    bdr = ET.SubElement(tcPr, f"{W}tcBorders")
    for side in ("top", "left", "bottom", "right"):
        ET.SubElement(bdr, f"{W}{side}", {f"{W}val": "single", f"{W}sz": "4", f"{W}space": "0", f"{W}color": "CBD5E1"})
    if shd_fill:
        ET.SubElement(tcPr, f"{W}shd", {f"{W}val": "clear", f"{W}color": "auto", f"{W}fill": shd_fill})
    tc.append(_make_report_para(text, bold=bold, sz_half_pt=20, space_before=40, space_after=40))
    return tc


def _write_unchanged_articles_report_docx(
    report_path: Path,
    unchanged_articles: List[dict],
    total_processed: int,
    lang_code: str,
    date_str: str,
) -> Path:
    """Write `全篇对比_英文无改点篇目汇总报告_{DATE}.docx` listing all 0-change articles cleanly."""
    with docx_xml_utils._ET_NS_LOCK:
        ET.register_namespace("w", W_URI)
        ET.register_namespace("r", R_URI)

    doc_root = ET.Element(f"{W}document")
    body = ET.SubElement(doc_root, f"{W}body")

    body.append(
        _make_report_para(
            "全篇改点对比 · 英文无改点篇目汇总报告",
            bold=True,
            sz_half_pt=32,
            color_hex="0F172A",
            jc="center",
            space_before=40,
            space_after=140,
        )
    )
    body.append(
        _make_report_para(
            f"整理日期：{date_str}    ｜    目标语种：{lang_code or 'XX'}    ｜    "
            f"共对比 {total_processed} 篇章节，其中英文无改点 {len(unchanged_articles)} 篇（已自动跳过导出给译者版与完整版）",
            bold=False,
            sz_half_pt=21,
            color_hex="334155",
            space_before=20,
            space_after=140,
        )
    )

    tbl = ET.SubElement(body, f"{W}tbl")
    tblPr = ET.SubElement(tbl, f"{W}tblPr")
    ET.SubElement(tblPr, f"{W}tblW", {f"{W}w": "5000", f"{W}type": "pct"})
    tblGrid = ET.SubElement(tbl, f"{W}tblGrid")
    col_specs = [
        ("序号", "400", "1165"),
        ("中文篇名 / 编号", "1500", "4368"),
        ("英文篇名 (Row 0)", "1700", "4950"),
        ("中英底稿来源", "900", "2621"),
        ("对比段落数", "500", "1456"),
    ]
    for _, _, dxa_w in col_specs:
        ET.SubElement(tblGrid, f"{W}gridCol", {f"{W}w": dxa_w})

    hdr_tr = ET.SubElement(tbl, f"{W}tr")
    for lbl, pct_w, _ in col_specs:
        hdr_tr.append(_make_report_cell(lbl, pct_w, bold=True, shd_fill="F1F5F9"))

    for idx_u, item in enumerate(unchanged_articles, 1):
        tr = ET.SubElement(tbl, f"{W}tr")
        cn_lbl = str(item.get("cn_title") or "")
        pref = str(item.get("query_prefix") or "").strip()
        if pref and not cn_lbl.startswith(pref):
            cn_lbl = f"{pref} {cn_lbl}"
        vals = [
            str(idx_u),
            cn_lbl,
            str(item.get("en_title") or "-"),
            str(item.get("master_rel_path") or "-"),
            f"{int(item.get('matched_rows', 0))} 段（无英文改点）",
        ]
        for val_s, (_, pct_w, _) in zip(vals, col_specs):
            tr.append(_make_report_cell(val_s, pct_w, bold=False))

    sect = ET.SubElement(body, f"{W}sectPr")
    ET.SubElement(sect, f"{W}pgSz", {f"{W}w": "16838", f"{W}h": "11906", f"{W}orient": "landscape"})
    ET.SubElement(
        sect,
        f"{W}pgMar",
        {
            f"{W}top": "1134",
            f"{W}right": "1134",
            f"{W}bottom": "1134",
            f"{W}left": "1134",
            f"{W}header": "720",
            f"{W}footer": "720",
            f"{W}gutter": "0",
        },
    )

    with docx_xml_utils._ET_NS_LOCK:
        ET.register_namespace("w", W_URI)
        ET.register_namespace("r", R_URI)
        doc_xml_str = ET.tostring(doc_root, encoding="unicode")

    doc_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + doc_xml_str).encode("utf-8")
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<w:styles xmlns:w="{W_URI}">'
        '<w:docDefaults>'
        '<w:rPrDefault><w:rPr>'
        '<w:rFonts w:ascii="Georgia" w:eastAsia="宋体" w:hAnsi="Georgia" w:cs="Georgia"/>'
        '<w:sz w:val="22"/><w:szCs w:val="22"/>'
        '</w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
        '</w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        '</w:styles>'
    ).encode("utf-8")
    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<Types xmlns="{CONTENT_TYPES_URI}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '</Types>'
    ).encode("utf-8")
    root_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<Relationships xmlns="{PKG_REL_URI}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>'
    ).encode("utf-8")
    doc_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<Relationships xmlns="{PKG_REL_URI}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>'
    ).encode("utf-8")

    report_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(report_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", root_rels_xml)
        zf.writestr("word/document.xml", doc_xml)
        zf.writestr("word/styles.xml", styles_xml)
        zf.writestr("word/_rels/document.xml.rels", doc_rels_xml)

    return report_path


def run_full_article_pipeline(
    selected_articles: Sequence[dict],
    foreign_lib_dir: Union[str, Path],
    output_dir: Union[str, Path],
    api_key: Union[str, List[str], ApiKeyPool] = "",
    model: str = "gemini-3.5-flash-lite",
    api_mode: str = "auto",
    rotation_mode: str = "round_robin",
    lang: str = "AUTO",
    enable_ai: bool = True,
    balance_columns: bool = True,
    dedup_pkg_dir: str = "",
    date_str: str = "",
    log_fn: Optional[Callable[[str], None]] = None,
    progress_fn: Optional[Callable[[float, str], None]] = None,
    use_title_list: bool = True,
    highlight_colors: bool = True,
    highlight_mode: Optional[str] = None,
    yellow_gap_words: int = 3,
    cn_en_master_dir: Union[str, Path] = "",
) -> dict:
    """
    Process full articles independently (`输出结果一篇一篇独立导出`) with v1.8 upgrades:
      - `use_title_list=True`: names outputs `{LANG}_{PREFIX_}{EN_TITLE}_{COUNT_TAG}_{DATE}.docx`;
      - `use_title_list=False` (direct folder mode): names outputs `{LANG}_{master_stem}_{COUNT_TAG}_{DATE}.docx`;
      - Skips `给译者版` & `完整版` for 0-change articles and writes `全篇对比_英文无改点篇目汇总报告_{DATE}.docx`;
      - Smart folder rule: `1` article -> flat in `output_dir`; `>= 2` articles -> 4 clean subfolders
        (`给译者版/`, `完整留底版/`, `中英最新底稿包/`, `外文三语原稿包/`).
    """
    flib_p = Path(foreign_lib_dir).resolve()
    out_d = Path(output_dir).resolve()
    if not flib_p.exists() or not flib_p.is_dir():
        raise FileNotFoundError(f"找不到外文三语总库目录：{flib_p}")
    out_d.mkdir(parents=True, exist_ok=True)

    articles_list = list(selected_articles or [])
    if not articles_list and not use_title_list and str(cn_en_master_dir or "").strip():
        direct_res = resolve_direct_folder_articles(cn_en_master_dir, flib_p, log_fn=log_fn)
        articles_list = list(direct_res.get("resolved") or [])

    if not articles_list:
        raise ValueError("请先至少选择或匹配 1 篇需要对比的章节。")

    _, use_hl, highlight_col1 = resolve_highlight_mode(highlight_mode, highlight_colors)
    highlight_colors = use_hl
    eff_enable_ai = bool(enable_ai and highlight_colors)
    pool_or_keys, eff_rot, has_keys, effective_ai = _prepare_optional_pool(
        api_key=api_key,
        rotation_mode=rotation_mode,
        enable_ai=eff_enable_ai,
    )

    final_date = docx_xml_utils.sanitize_filename_part(
        (date_str or "").strip() or datetime.now().strftime("%Y%m%d"),
        fallback=datetime.now().strftime("%Y%m%d"),
        max_len=16,
    )

    if progress_fn:
        progress_fn(2.0, "正在预热外文三语总库检索索引...")
    idx_stats = ensure_library_index(flib_p, log_fn=log_fn)
    if log_fn:
        log_fn(
            f"📚 外文三语总库就绪：共 {idx_stats.get('total_files', 0)} 篇文档"
            f"（{idx_stats.get('total_rows', 0)} 段，耗时 {idx_stats.get('elapsed_sec', 0.0):.2f}s）"
        )

    cfg = docx_xml_utils.load_config()
    col2_up = cfg.get("col2_up_highlight", "yellow")
    col2_rv = cfg.get("col2_rv_highlight", "cyan")
    clean_dedup_dir = (dedup_pkg_dir or "").strip()
    clamped_gap = max(0, int(yellow_gap_words if yellow_gap_words is not None else 3))

    supports_full_diff_flag = _script_supports_cli_flag(
        step1_extract_and_diff, "--force-full-diff", "format_up_rv_count_tag"
    )
    supports_ygap = _script_supports_cli_flag(step1_render_draft, "--yellow-gap-words")
    supports_no_hl1 = _script_supports_cli_flag(step1_render_draft, "--no-highlight-colors")
    supports_raw_tri = _script_supports_cli_flag(step1_render_draft, "--raw-trilingual-output")
    supports_no_bal = _script_supports_cli_flag(
        step2_render_final, "--no-balance-columns", "balance_trilingual_table_widths"
    )
    supports_no_hl2 = _script_supports_cli_flag(step2_render_final, "--no-highlight-colors")
    supports_no_col1 = _script_supports_cli_flag(step2_render_final, "--no-col1-highlights")
    supports_keep_row0 = _script_supports_cli_flag(step2_render_final, "--keep-first-title-row")
    step2_native_v17 = supports_no_bal and callable(getattr(docx_xml_utils, "format_up_rv_count_tag", None))

    total_arts = len(articles_list)
    use_subfolders = True
    trans_out_dir = out_d / "给译者版"
    full_out_dir = out_d / "完整留底版"
    master_pkg_dir = out_d / "中英最新底稿包"
    foreign_raw_pkg_dir = out_d / "外文三语原稿包"
    for sub_d in (trans_out_dir, full_out_dir, master_pkg_dir, foreign_raw_pkg_dir):
        sub_d.mkdir(parents=True, exist_ok=True)

    saved_files: List[str] = []
    article_results: List[dict] = []
    unchanged_articles: List[dict] = []
    duplicate_reports: List[str] = []
    final_lang_code = (lang or "AUTO").strip().upper()

    en_title_counts: Dict[str, int] = {}
    for a_item in articles_list:
        raw_et = str(a_item.get("en_title") or a_item.get("cn_title") or "").strip().lower()
        q_pf = str(a_item.get("query_prefix") or "").strip().lower()
        et_key = f"{q_pf}_{raw_et}" if q_pf else raw_et
        if et_key:
            en_title_counts[et_key] = en_title_counts.get(et_key, 0) + 1
    used_doc_labels: Dict[str, int] = {}

    for a_idx, art in enumerate(articles_list):
        base_pct = 5.0 + (a_idx / max(1, total_arts)) * 92.0
        span_pct = 92.0 / max(1, total_arts)
        cn_disp = art.get("cn_title") or art.get("query") or f"第{a_idx + 1}篇"

        if log_fn:
            log_fn(f"\n📖 === [{a_idx + 1}/{total_arts}] 开始全篇提取对比：《{cn_disp}》 ===")
        if progress_fn:
            progress_fn(base_pct, f"[{a_idx + 1}/{total_arts}] 正在提取《{cn_disp}》并比对外文总库...")

        tmp_root = create_ephemeral_workdir(f"full_art_{a_idx + 1}")
        try:
            extracted_docx = tmp_root / "extracted_article.docx"
            workdir = tmp_root / "work"
            workdir.mkdir(parents=True, exist_ok=True)

            en_title, restore_heading = _extract_standalone_article_docx(art, extracted_docx)

            is_direct_mode = (not use_title_list) or bool(art.get("direct_folder_mode"))
            query_prefix = str(art.get("query_prefix") or "").strip()
            if not query_prefix and use_title_list and art.get("query"):
                query_prefix, _ = split_title_prefix_and_core(str(art["query"]))

            if is_direct_mode:
                raw_stem = str(art.get("naming_stem") or Path(art["master_file"]).stem).strip()
                raw_label = parse_doc_filename(raw_stem, strip_draft_prefix=True)["orig_name"]
            else:
                clean_en_title = parse_doc_filename(en_title, strip_draft_prefix=True)["orig_name"] if en_title else en_title
                raw_label = f"{query_prefix}_{clean_en_title}" if query_prefix else clean_en_title
                et_lookup = f"{query_prefix.lower()}_{en_title.strip().lower()}" if query_prefix else en_title.strip().lower()
                if en_title_counts.get(et_lookup, 0) > 1:
                    sec_tag = str(art.get("section_path") or "").split("›")[-1].strip()
                    if sec_tag:
                        raw_label = f"{raw_label} ({sec_tag})"

            doc_label = docx_xml_utils.sanitize_filename_part(raw_label, fallback="Article", max_len=125)
            lbl_key = doc_label.lower()
            if lbl_key in used_doc_labels:
                used_doc_labels[lbl_key] += 1
                doc_label = docx_xml_utils.sanitize_filename_part(
                    f"{doc_label}_{used_doc_labels[lbl_key]}", fallback="Article", max_len=125
                )
            else:
                used_doc_labels[lbl_key] = 1

            # 1. Step 1a: Extract & full-chapter diff
            argv1 = [
                "step1_extract_and_diff.py",
                "--input", str(extracted_docx),
                "--library", str(flib_p),
                "--workdir", str(workdir),
            ]
            if final_lang_code and final_lang_code != "AUTO":
                argv1.extend(["--lang", final_lang_code])
            if supports_full_diff_flag:
                argv1.append("--force-full-diff")

            _run_module_main(step1_extract_and_diff.main, argv1, log_fn=log_fn)

            manifest_path = workdir / "diff_manifest.json"
            review_items_path = workdir / "step1_review_items.json"
            _repair_skipped_title_rows_in_manifest(manifest_path, review_items_path)

            with open(manifest_path, "r", encoding="utf-8") as f:
                manifest = json.load(f)
            with open(review_items_path, "r", encoding="utf-8") as f:
                review_items = json.load(f)

            det_lang = manifest.get("lang_code") or (
                final_lang_code if final_lang_code and final_lang_code != "AUTO" else "XX"
            )
            if not final_lang_code or final_lang_code == "AUTO":
                final_lang_code = det_lang

            safe_lang = docx_xml_utils.sanitize_filename_part(det_lang, fallback="XX", max_len=16)

            # Prepare source backup output paths (Item #3 & #4: 中英最新底稿 & 外文三语原稿)
            master_backup_name = f"{doc_label}_中英最新底稿_{final_date}.docx"
            master_backup_path = docx_xml_utils.safe_join_output_file(master_pkg_dir, master_backup_name)
            _export_master_article_backup_docx(art, master_backup_path)

            foreign_raw_name = f"{safe_lang}_{doc_label}_提取三语原稿_{final_date}.docx"
            foreign_raw_path = docx_xml_utils.safe_join_output_file(foreign_raw_pkg_dir, foreign_raw_name)

            # Check Item #7②: 0 English changes fast check
            has_en_changes = _manifest_has_any_en_edits(manifest, review_items)
            draft_path = tmp_root / "draft_article.docx"

            if has_en_changes:
                # 2. Step 1b: Evaluate UP/RV & render draft
                s1_stats: dict = {}
                if not highlight_colors:
                    decisions = {}
                else:
                    def _s1_prog(frac: float, msg: str):
                        if progress_fn:
                            progress_fn(base_pct + span_pct * (0.15 + frac * 0.35), f"[{a_idx + 1}/{total_arts}] {msg}")

                    decisions = evaluate_step1_decisions(
                        review_items=review_items,
                        api_key=pool_or_keys,
                        model=model,
                        api_mode=api_mode,
                        rotation_mode=eff_rot,
                        log_fn=log_fn,
                        progress_fn=_s1_prog,
                        enable_ai=effective_ai,
                        stats_out=s1_stats,
                    )
                    if effective_ai and s1_stats.get("used_local_fallback") and s1_stats.get("fallback_reason"):
                        effective_ai = False

                decisions_path = workdir / "step1_decisions.json"
                decisions_path.write_text(json.dumps(decisions, ensure_ascii=False, indent=2), encoding="utf-8")

                argv1b = [
                    "step1_render_draft.py",
                    "--manifest", str(manifest_path),
                    "--decisions", str(decisions_path),
                    "--output", str(draft_path),
                ]
                if supports_ygap:
                    argv1b.extend(["--yellow-gap-words", str(clamped_gap)])
                if not highlight_colors and supports_no_hl1:
                    argv1b.append("--no-highlight-colors")
                if supports_raw_tri:
                    argv1b.extend(["--raw-trilingual-output", str(foreign_raw_path)])

                _run_module_main(step1_render_draft.main, argv1b, log_fn=log_fn)

                if not highlight_colors and not supports_no_hl1 and draft_path.is_file():
                    d_root, _, d_ns = docx_xml_utils.load_docx_xml(draft_path)
                    d_cmts = docx_xml_utils.load_docx_comments_xml(draft_path)
                    _strip_all_highlights_in_root(d_root)
                    docx_xml_utils.save_docx_xml(draft_path, draft_path, d_root, d_ns, comments_root=d_cmts)

                _restore_heading_in_docx(draft_path, restore_heading)
                if not _docx_has_any_col2_revisions(draft_path):
                    has_en_changes = False

            if not supports_raw_tri or not foreign_raw_path.is_file():
                _build_raw_trilingual_docx_from_manifest(
                    manifest_path=manifest_path,
                    output_path=foreign_raw_path,
                    balance_columns=bool(balance_columns),
                )
            _restore_heading_in_docx(foreign_raw_path, restore_heading)

            saved_files.append(str(master_backup_path))
            saved_files.append(str(foreign_raw_path))

            # Item #7②: If 0 English changes -> skip `给译者版` and `完整版`, record in `unchanged_articles`!
            if not has_en_changes:
                if log_fn:
                    log_fn(f"ℹ️ 《{cn_disp}》英文对比无改点，已归入无改点汇总报告")
                unchanged_entry = {
                    "cn_title": str(cn_disp),
                    "en_title": str(en_title),
                    "doc_label": str(doc_label),
                    "query_prefix": query_prefix,
                    "master_rel_path": str(art.get("master_rel_path") or Path(art["master_file"]).name),
                    "matched_rows": len(manifest.get("rows") or []),
                    "skipped_rows": len(manifest.get("skipped_rows") or []),
                }
                unchanged_articles.append(unchanged_entry)
                article_results.append({
                    "cn_title": cn_disp,
                    "en_title": en_title,
                    "lang_code": det_lang,
                    "has_changes": False,
                    "full_path": "",
                    "translator_path": "",
                    "master_backup_path": str(master_backup_path),
                    "foreign_raw_path": str(foreign_raw_path),
                    "up_words": 0,
                    "rv_words": 0,
                    "full_up_words": 0,
                    "full_rv_words": 0,
                    "changed_rows": 0,
                    "skipped_count": len(manifest.get("skipped_rows") or []),
                    "dedup_removed_count": 0,
                    "dedup_report_path": "",
                })
                continue

            # 3. Step 2a + 2b: Parse spans & align Col 1 + Col 3
            argv2a = [
                "step2_parse_reviewed.py",
                "--draft", str(draft_path),
                "--workdir", str(workdir),
            ]
            _run_module_main(step2_parse_reviewed.main, argv2a, log_fn=log_fn)

            tasks_path = workdir / "step2_align_tasks.json"
            tasks_data = json.loads(tasks_path.read_text(encoding="utf-8"))

            if not highlight_colors:
                alignments = {}
            else:
                def _s2_prog(frac: float, msg: str):
                    if progress_fn:
                        progress_fn(base_pct + span_pct * (0.52 + frac * 0.35), f"[{a_idx + 1}/{total_arts}] {msg}")

                lang_hint = find_lang_display_name(det_lang, load_languages())
                s2_stats: dict = {}
                alignments = align_step2_spans(
                    tasks_data=tasks_data,
                    find_fuzzy_fn=docx_xml_utils.find_fuzzy_substring_span,
                    api_key=pool_or_keys,
                    model=model,
                    api_mode=api_mode,
                    rotation_mode=eff_rot,
                    log_fn=log_fn,
                    progress_fn=_s2_prog,
                    target_lang_hint=lang_hint,
                    enable_ai=effective_ai,
                    stats_out=s2_stats,
                )
                if effective_ai and s2_stats.get("used_local_fallback") and s2_stats.get("fallback_reason"):
                    effective_ai = False

            alignments_path = workdir / "step2_alignments.json"
            alignments_path.write_text(json.dumps(alignments, ensure_ascii=False, indent=2), encoding="utf-8")

            # 4A. Deliverable 1: 整篇完整留底版 `{LANG}_{DOC_LABEL}_{COUNT_TAG}_{DATE}.docx` (inside `完整留底版/`)
            full_label = doc_label
            argv_full = [
                "step2_render_final.py",
                "--tasks", str(tasks_path),
                "--alignments", str(alignments_path),
                "--outdir", str(full_out_dir),
                "--lang", det_lang,
                "--doc-name", full_label,
                "--date", final_date,
            ]
            if not balance_columns and supports_no_bal:
                argv_full.append("--no-balance-columns")
            if not highlight_colors and supports_no_hl2:
                argv_full.append("--no-highlight-colors")
            elif highlight_colors and not highlight_col1 and supports_no_col1:
                argv_full.append("--no-col1-highlights")
            out_full_txt = _run_module_main(step2_render_final.main, argv_full, log_fn=log_fn)
            m_full = re.search(r"Final Document Saved:\s*(.+)", out_full_txt)
            if not m_full:
                raise RuntimeError(f"未能成功写出《{cn_disp}》整篇完整版文档。")

            raw_full_path = Path(m_full.group(1).strip())
            full_final_path, full_up, full_rv = _postprocess_rendered_docx(
                saved_path=raw_full_path,
                lang_code=det_lang,
                target_label=full_label,
                date_yyyymmdd=final_date,
                balance_columns=bool(balance_columns),
                step2_native_v17=step2_native_v17,
                highlight_colors=bool(highlight_colors),
                step2_supports_no_hl=supports_no_hl2,
            )
            saved_files.append(str(full_final_path))

            # 4B. Deliverable 2: 仅改点给译者版 `{LANG}_{DOC_LABEL}_{COUNT_TAG}_{DATE}.docx`
            if supports_keep_row0:
                argv_trans = [
                    "step2_render_final.py",
                    "--tasks", str(tasks_path),
                    "--alignments", str(alignments_path),
                    "--outdir", str(trans_out_dir),
                    "--lang", det_lang,
                    "--doc-name", doc_label,
                    "--date", final_date,
                    "--remove-unchanged-tables",
                    "--keep-first-title-row",
                ]
                if not balance_columns and supports_no_bal:
                    argv_trans.append("--no-balance-columns")
                if not highlight_colors and supports_no_hl2:
                    argv_trans.append("--no-highlight-colors")
                elif highlight_colors and not highlight_col1 and supports_no_col1:
                    argv_trans.append("--no-col1-highlights")
                out_trans_txt = _run_module_main(step2_render_final.main, argv_trans, log_fn=log_fn)
                m_trans = re.search(r"Final Document Saved:\s*(.+)", out_trans_txt)
                if not m_trans:
                    raise RuntimeError(f"未能成功写出《{cn_disp}》给译者版文档。")
                raw_trans_path = Path(m_trans.group(1).strip())
                trans_path, trans_up, trans_rv = _postprocess_rendered_docx(
                    saved_path=raw_trans_path,
                    lang_code=det_lang,
                    target_label=doc_label,
                    date_yyyymmdd=final_date,
                    balance_columns=bool(balance_columns),
                    step2_native_v17=step2_native_v17,
                    highlight_colors=bool(highlight_colors),
                    step2_supports_no_hl=supports_no_hl2,
                )
                chk_root, _, _ = docx_xml_utils.load_docx_xml(trans_path)
                kept_chg_rows = _count_changed_body_rows_in_root(chk_root, col2_up, col2_rv)
            else:
                trans_root, _, trans_ns = docx_xml_utils.load_docx_xml(full_final_path)
                trans_cmts = docx_xml_utils.load_docx_comments_xml(full_final_path)
                kept_chg_rows = _filter_full_article_to_translator_version(trans_root, col2_up, col2_rv)
                if balance_columns:
                    balance_trilingual_table_widths(trans_root)

                if not highlight_colors:
                    _strip_all_highlights_in_root(trans_root)
                    trans_up, trans_rv = 0, 0
                else:
                    trans_up, trans_rv = docx_xml_utils.compute_col2_up_rv_words(
                        trans_root, up_color=col2_up, rv_color=col2_rv
                    )
                    update_top_highlight_instruction_lines(
                        trans_root, has_up=(trans_up > 0), has_rv=(trans_rv > 0)
                    )

                trans_name = format_counted_docx_name(
                    lang_code=det_lang,
                    orig_name=doc_label,
                    up_words=trans_up,
                    rv_words=trans_rv,
                    date_yyyymmdd=final_date,
                )
                trans_path = docx_xml_utils.safe_join_output_file(trans_out_dir, trans_name)
                docx_xml_utils.save_docx_xml(
                    full_final_path, trans_path, trans_root, trans_ns, comments_root=trans_cmts
                )

            dedup_removed = 0
            dedup_rep = ""
            if clean_dedup_dir:
                d_res = remove_unmerged_duplicates_from_docx(
                    docx_path=trans_path,
                    unmerged_pkg_dir=clean_dedup_dir,
                    report_dir=out_d,
                    doc_tag=doc_label,
                    log_fn=log_fn,
                    balance_columns=bool(balance_columns),
                    date_str=final_date,
                )
                dedup_removed = int(d_res.get("duplicate_count", 0))
                dedup_rep = str(d_res.get("report_path") or "")
                trans_up = int(d_res.get("up_words", trans_up)) if highlight_colors else 0
                trans_rv = int(d_res.get("rv_words", trans_rv)) if highlight_colors else 0
                if dedup_rep:
                    duplicate_reports.append(dedup_rep)
                if dedup_removed > 0:
                    recalc_name = format_counted_docx_name(
                        lang_code=det_lang,
                        orig_name=doc_label,
                        up_words=trans_up,
                        rv_words=trans_rv,
                        date_yyyymmdd=final_date,
                    )
                    recalc_p = docx_xml_utils.safe_join_output_file(trans_out_dir, recalc_name)
                    if os.path.normcase(str(recalc_p.resolve())) != os.path.normcase(str(trans_path.resolve())):
                        if recalc_p.exists():
                            recalc_p.unlink()
                        trans_path.rename(recalc_p)
                        trans_path = recalc_p

            saved_files.append(str(trans_path))

            if log_fn:
                log_fn(f"  ✅ 《{cn_disp}》导出完成：")
                log_fn(f"     - 给译者版：{trans_path.name}（保留改点行 {kept_chg_rows} 段）")
                log_fn(f"     - 完整留底版：{full_final_path.name}")

            article_results.append({
                "cn_title": cn_disp,
                "en_title": en_title,
                "lang_code": det_lang,
                "has_changes": True,
                "full_path": str(full_final_path),
                "translator_path": str(trans_path),
                "master_backup_path": str(master_backup_path),
                "foreign_raw_path": str(foreign_raw_path),
                "up_words": int(trans_up),
                "rv_words": int(trans_rv),
                "full_up_words": int(full_up),
                "full_rv_words": int(full_rv),
                "changed_rows": kept_chg_rows,
                "skipped_count": len(manifest.get("skipped_rows") or []),
                "dedup_removed_count": dedup_removed,
                "dedup_report_path": dedup_rep,
            })
        finally:
            shutil.rmtree(tmp_root, ignore_errors=True)

    unchanged_report_path = ""
    if unchanged_articles:
        rep_name = f"全篇对比_英文无改点篇目汇总报告_{final_date}.docx"
        rep_p = docx_xml_utils.safe_join_output_file(out_d, rep_name)
        _write_unchanged_articles_report_docx(
            report_path=rep_p,
            unchanged_articles=unchanged_articles,
            total_processed=total_arts,
            lang_code=final_lang_code if final_lang_code != "AUTO" else "XX",
            date_str=final_date,
        )
        unchanged_report_path = str(rep_p)
        saved_files.append(unchanged_report_path)
        if log_fn:
            log_fn(f"\n📋 已生成英文无改点篇目汇总报告（共 {len(unchanged_articles)} 篇）：{rep_p.name}")

    if progress_fn:
        progress_fn(
            100.0,
            f"全篇对比整理完成：共处理 {total_arts} 篇章节"
            f"（有改点 {total_arts - len(unchanged_articles)} 篇，无改点 {len(unchanged_articles)} 篇）！",
        )

    return {
        "saved_files": saved_files,
        "article_results": article_results,
        "unchanged_articles": unchanged_articles,
        "unchanged_count": len(unchanged_articles),
        "unchanged_report_path": unchanged_report_path,
        "zero_change_report_path": unchanged_report_path,
        "lang_code": final_lang_code if final_lang_code != "AUTO" else "XX",
        "duplicate_reports": duplicate_reports,
        "used_subfolders": use_subfolders,
    }
