from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
import hashlib
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
from pipeline import (  # noqa: E402
    _postprocess_rendered_docx,
    _prepare_optional_pool,
    _run_module_main,
    _script_supports_cli_flag,
    ensure_library_index,
)
from word_counter import (  # noqa: E402
    balance_trilingual_table_widths,
    ensure_three_column_tables_and_landscape,
    extract_full_article_en_title,
    format_counted_docx_name,
    is_top_highlight_instruction_text,
    update_top_highlight_instruction_lines,
)

W = docx_xml_utils.W
NS = docx_xml_utils.NS
CN_NUM_CHARS = "一二三四五六七八九十百零"

MASTER_INDEX_SCHEMA_VERSION = 2

_BULLET_PREFIX_RE = re.compile(
    r"^\s*(?:[•●▪▸►*]+|\(\s*\d+\s*\)|（\s*\d+\s*）|\d+\s*[\.、\)）])\s*"
)
_CHAPTER_CODE_PREFIX_RE = re.compile(r"^\d+(?:-\d+)+[\s_]+")
_EXCERPT_SUFFIX_RE = re.compile(r"[\s_]*(?:（选段\s*\d+\s*-\s*\d+）|\(选段\s*\d+\s*-\s*\d+\)|\d+\s*-\s*\d+)\s*$")
_SUB_EXCERPT_TITLE_RE = re.compile(r"^(?:选段|片段|excerpt|part)\s*\d+$", re.IGNORECASE)

_PART_DIVIDER_RE = re.compile(r"^第[\s一二三四五六七八九十百\d]+部分")
_SUBPART_DIVIDER_RE = re.compile(
    r"^(?:附篇[\s\u3000]*[“\"]?神向全宇|基督在众教会行走时的说话[\s\u3000]*[一二三四五六七八九十\d]+\s*$)"
)

_WS_RE = re.compile(r"[\s\u3000\u200b\u200c\ufeff]+")
_LOOSE_KEEP_RE = re.compile(r"[\u4e00-\u9fff\u3400-\u4dbfA-Za-z0-9]+")

_MASTER_CACHE_LOCK = threading.Lock()
_MEM_MASTER_CACHE: Dict[str, Tuple[tuple, List[dict]]] = {}


# ---------------------------------------------------------------------------
# Title List Parsing (`parse_titles_multiline` & `parse_titles_text_or_file`)
# ---------------------------------------------------------------------------

def _clean_single_title_line(line: str) -> str:
    s = (line or "").strip().lstrip("\ufeff")
    if not s:
        return ""
    # Strip list bullets like '1. ', '2、', '(3) ', '• ' while keeping '1-01 ...' chapter prefixes intact!
    s = _BULLET_PREFIX_RE.sub("", s).strip()
    if s.startswith("《") and s.endswith("》") and s.count("《") == 1 and s.count("》") == 1:
        s = s[1:-1].strip()
    return s


def parse_titles_multiline(raw_text: str) -> List[str]:
    """
    Split multiline user input into clean, deduplicated non-empty article titles.
    Strips list bullet numbers ('1. ', '2、', '(1)') while keeping chapter prefixes ('1-01 ...') intact.
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
    return _WS_RE.sub("", s_clean).lower()


def _norm_loose(s: str) -> str:
    return "".join(_LOOSE_KEEP_RE.findall(s or "")).lower()


def _strip_chapter_code_and_range(stem: str) -> str:
    s = (stem or "").strip()
    s = _CHAPTER_CODE_PREFIX_RE.sub("", s).strip(" _")
    s = _EXCERPT_SUFFIX_RE.sub("", s).strip(" _")
    return s


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
    if master_path.is_file():
        st = os.stat(master_path)
        return [(master_path.name, str(master_path), st.st_mtime_ns, st.st_size)]

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
                if ename.startswith(".") or any(k in ename.lower() for k in ("output", "temp", "初筛", "标记结果")):
                    continue
                nxt = f"{rel_prefix}/{ename}" if rel_prefix else ename
                _walk(entry.path, nxt)
            elif entry.is_file(follow_symlinks=False):
                if not ename.lower().endswith(".docx") or ename.startswith("~$") or ename.startswith("."):
                    continue
                if ename.startswith("初筛_") or ename.startswith("查重结果"):
                    continue
                rel_posix = f"{rel_prefix}/{ename}" if rel_prefix else ename
                try:
                    st = os.stat(entry.path)
                except OSError:
                    continue
                found.append((rel_posix, str(Path(entry.path).resolve()), st.st_mtime_ns, st.st_size))

    _walk(str(master_path), "")
    found.sort(key=lambda x: x[0])
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
    Detect book/volume/part/sub-part cover divider tables in multi-article `.docx` files
    (e.g. `话在肉身显现` Volume 1, `第一部分`, `第二部分`, `附篇 神向全宇说话的奥秘揭示`, `第三部分`, `基督在众教会行走时的说话 一`).
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
        lbl = re.sub(r"──.*$", "", joined_p or r0_cn).strip()
        return True, "part", lbl
    if _SUBPART_DIVIDER_RE.match(cand_txt):
        return True, "subpart", re.sub(r"\s+", " ", cand_txt).strip()
    return False, "", ""


def _parse_single_master_docx(task: Tuple[str, str, int, int]) -> Tuple[str, dict]:
    rel_posix, abs_path_str, mtime_ns, size = task
    p = Path(abs_path_str)
    stem = p.stem.strip()
    clean_stem_title = _strip_chapter_code_and_range(stem)
    parent_rel = str(Path(rel_posix).parent).replace("\\", "/")
    if parent_rel == ".":
        parent_rel = ""

    try:
        with zipfile.ZipFile(abs_path_str, "r") as zf:
            docx_xml_utils._validate_docx_zip(zf)
            xml_bytes = zf.read("word/document.xml")
        docx_xml_utils._assert_safe_openxml_bytes(xml_bytes)
        root = ET.fromstring(xml_bytes)
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

    # Scan all <w:tbl> blocks with their immediately preceding non-empty <w:p> headings
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

    # Determine whether this `.docx` is a single-article file (including multi-table `选段1..9` files!)
    # or a multi-article compilation/book summary `.docx`.
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

        aliases = []
        for a in (cn_title, stem, clean_stem_title, u0["r0_cn"]):
            a_s = (a or "").strip()
            if a_s and a_s not in aliases:
                aliases.append(a_s)

        sec_disp = parent_rel or "独立单篇"
        disp_en = f" ({en_title})" if en_title else ""
        display_label = f"{rel_posix}{disp_en} [{tot_rows}行]"

        articles.append({
            "cand_id": f"{rel_posix}::FULL",
            "display_label": display_label,
            "master_file": abs_path_str,
            "master_rel_path": rel_posix,
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
            # Check if an earlier paragraph in `p_titles` was a standalone section heading before a chapter heading
            if len(p_titles) >= 2:
                lead_p = p_titles[0]
                lead_compact = _WS_RE.sub("", lead_p)
                if _PART_DIVIDER_RE.match(lead_compact):
                    cur_part = re.sub(r"──.*$", "", lead_p).strip()
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

            # Build qualified title for generic chapter names like '第一篇' / '内容简介' / '后记'
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

            aliases = []
            for a in [*extra_aliases, qualified_cn_title]:
                a_s = (a or "").strip()
                if a_s and a_s not in aliases:
                    aliases.append(a_s)

            b_start, b_end = u["start_idx"], u["end_idx"]
            en_t = u["en_title"]
            disp_sec = f" › {section_path}" if section_path else ""
            disp_en = f" ({en_t})" if en_t else ""
            display_label = f"{book_label}{disp_sec} › {cn_title}{disp_en} [{u['n_rows']}行]"

            articles.append({
                "cand_id": f"{rel_posix}::{b_start}:{b_end}",
                "display_label": display_label,
                "master_file": abs_path_str,
                "master_rel_path": rel_posix,
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


def _load_or_build_master_articles(
    cn_en_master_dir: Union[str, Path],
    log_fn: Optional[Callable[[str], None]] = None,
    max_workers: int = 8,
) -> List[dict]:
    t0 = time.perf_counter()
    master_p = Path(cn_en_master_dir).resolve()
    if not master_p.exists():
        raise FileNotFoundError(f"找不到最新中英底稿目录：{master_p}")

    norm_key = os.path.normcase(str(master_p))
    disk_stats = _scan_master_docx_stats(master_p)
    stat_sig = tuple((rp, mt, sz) for rp, _, mt, sz in disk_stats)

    with _MASTER_CACHE_LOCK:
        mem_hit = _MEM_MASTER_CACHE.get(norm_key)
        if mem_hit is not None and mem_hit[0] == stat_sig:
            return mem_hit[1]

        cache_path = _resolve_master_cache_path(master_p)
        cached_files: Dict[str, dict] = {}
        if cache_path.exists():
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
            prev = cached_files.get(rp)
            if (
                prev is not None
                and prev.get("mtime_ns") == mt
                and prev.get("size") == sz
                and isinstance(prev.get("articles"), list)
            ):
                next_files[rp] = prev
            else:
                to_parse.append((rp, abs_s, mt, sz))

        if to_parse:
            workers = max(1, min(int(max_workers or 8), len(to_parse)))
            if log_fn and len(to_parse) >= 3:
                log_fn(f"📖 正在多线程扫描中英底稿目录（解析 {len(to_parse)}/{len(disk_stats)} 个文档）...")
            if workers == 1:
                parsed = [_parse_single_master_docx(t) for t in to_parse]
            else:
                with ThreadPoolExecutor(max_workers=workers) as ex:
                    parsed = list(ex.map(_parse_single_master_docx, to_parse))
            for rp, rec in parsed:
                next_files[rp] = rec

            try:
                payload = {
                    "version": MASTER_INDEX_SCHEMA_VERSION,
                    "master_dir": str(master_p),
                    "updated_at": time.time(),
                    "files": next_files,
                }
                tmp_p = cache_path.with_name(f"{cache_path.name}.tmp_{os.getpid()}")
                tmp_p.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
                os.replace(tmp_p, cache_path)
            except OSError:
                pass

        all_articles: List[dict] = []
        for rp in sorted(next_files.keys()):
            all_articles.extend(next_files[rp].get("articles") or [])

        _MEM_MASTER_CACHE[norm_key] = (stat_sig, all_articles)
        elapsed = time.perf_counter() - t0
        if log_fn:
            log_fn(f"📖 中英底稿索引就绪：共识别 {len(all_articles)} 篇可提取章节（耗时 {elapsed:.2f}s）")
        return all_articles


# ---------------------------------------------------------------------------
# Article Resolution & Disambiguation (`resolve_full_articles`)
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
        # Prevent '第一篇' from matching '第十一篇'/'第二十一篇', or '十' from matching '十一'
        if needle_norm[-1] in CN_NUM_CHARS and end_pos < len(target_norm) and target_norm[end_pos] in CN_NUM_CHARS:
            pos = target_norm.find(needle_norm, pos + 1)
            continue
        if needle_norm.startswith("第") and pos > 0 and target_norm[pos - 1] in (CN_NUM_CHARS + "附加"):
            pos = target_norm.find(needle_norm, pos + 1)
            continue
        return True
    return False


def _match_single_query(query: str, all_articles: List[dict]) -> List[dict]:
    clean_q = _clean_single_title_line(query)
    if not clean_q:
        return []

    # Split optional section/volume qualifiers separated by '・', '·', '>', '›', '/'
    parts = [p.strip().strip("《》") for p in re.split(r"[・·>›/｜|]+", clean_q) if p.strip()]
    if len(parts) == 1 and re.search(r"^(?:起初|全宇|揭示|基督起初的发表|神向全宇的说话|对神奥秘的揭示)[\s_]+(.+)$", clean_q):
        m_sp = re.match(r"^(起初|全宇|揭示|基督起初的发表|神向全宇的说话|对神奥秘的揭示)[\s_]+(.+)$", clean_q)
        if m_sp:
            parts = [m_sp.group(1).strip(), m_sp.group(2).strip()]

    leaf = parts[-1] if parts else clean_q
    qualifiers = parts[:-1] if len(parts) >= 2 else []

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
    leaf_stripped_compact = _norm_compact(_strip_chapter_code_and_range(leaf))

    # Tier 1: Exact match on aliases / cn_title / en_title / stem
    exact_hits: List[dict] = []
    for art in pool:
        alias_compacts = {_norm_compact(a) for a in art.get("aliases", ())}
        alias_looses = {_norm_loose(a) for a in art.get("aliases", ())}
        en_compact = _norm_compact(art.get("en_title", ""))
        if (
            q_full_compact in alias_compacts
            or leaf_compact in alias_compacts
            or (leaf_stripped_compact and leaf_stripped_compact in alias_compacts)
            or (q_full_loose and len(q_full_loose) >= 4 and q_full_loose in alias_looses)
            or (leaf_loose and len(leaf_loose) >= 4 and leaf_loose in alias_looses)
            or (en_compact and q_full_compact == en_compact)
        ):
            exact_hits.append(art)

    if exact_hits:
        return exact_hits

    # Tier 2: Safe substring / containment match on aliases and stems
    sub_hits: List[dict] = []
    for art in pool:
        hay_list = [_norm_compact(a) for a in art.get("aliases", ())]
        if art.get("en_title"):
            hay_list.append(_norm_compact(art["en_title"]))
        if any(_safe_substring_hit(leaf_compact, h) for h in hay_list):
            sub_hits.append(art)
        elif leaf_stripped_compact and len(leaf_stripped_compact) >= 4 and any(
            _safe_substring_hit(leaf_stripped_compact, h) for h in hay_list
        ):
            sub_hits.append(art)

    return sub_hits


def resolve_full_articles(
    titles: Sequence[str],
    cn_en_master_dir: Union[str, Path],
    foreign_lib_dir: Union[str, Path] = "",
    log_fn: Optional[Callable[[str], None]] = None,
) -> dict:
    """
    Resolve user-supplied chapter titles against `cn_en_master_dir` (supporting both single-article
    `.docx` files and multi-article book compilation `.docx` files).
    Returns:
      {
        "resolved": list[dict],
        "ambiguous": [{"query": str, "candidates": list[dict]}, ...],
        "unmatched": list[str],
      }
    """
    clean_titles: List[str] = []
    seen_q = set()
    for raw_t in titles or []:
        for t in parse_titles_multiline(str(raw_t)):
            if t not in seen_q:
                seen_q.add(t)
                clean_titles.append(t)

    all_articles = _load_or_build_master_articles(cn_en_master_dir, log_fn=log_fn)

    # Pre-warm foreign library index if provided and valid
    if foreign_lib_dir and str(foreign_lib_dir).strip():
        flib_p = Path(foreign_lib_dir).resolve()
        if flib_p.is_dir():
            ensure_library_index(flib_p, log_fn=log_fn)

    resolved: List[dict] = []
    ambiguous: List[dict] = []
    unmatched: List[str] = []
    seen_cand_ids = set()

    for q in clean_titles:
        hits = _match_single_query(q, all_articles)
        if not hits:
            unmatched.append(q)
            if log_fn:
                log_fn(f"  ⚠️ 未在中英底稿目录找到匹配篇目：《{q}》")
        elif len(hits) == 1:
            cand = dict(hits[0])
            cand["query"] = q
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


# ---------------------------------------------------------------------------
# Full-Article Extraction, Full-Chapter Comparison & Dual-Deliverable Export
# ---------------------------------------------------------------------------

def _strip_master_shading_and_highlights(root: ET.Element) -> None:
    """
    Remove pre-existing `w:shd` and `w:highlight` from master table cells (`【标you】` annotations)
    so Step 1 compares 100% of sentences across the entire chapter and leaves clean white backgrounds.
    """
    for tbl in root.findall(".//w:tbl", NS):
        for tc in tbl.findall(".//w:tc", NS):
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

    # Temporarily qualify bare generic <w:p> title (e.g. '第一篇' -> '基督起初的发表・第一篇')
    # so Step 1 Stage 1 scopes directly to the right foreign library chapter!
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


def _repair_skipped_title_rows_in_manifest(
    manifest_path: Path,
    review_items_path: Path,
) -> int:
    """
    If Row 0 (e.g. short 2-char chapter title like '后记' with changed English title) was skipped by
    `step1_extract_and_diff`, recover it from Row 0 of the same table's matched `source_file`!
    Also ensure unchanged Row 0 chapter titles are never falsely flagged as UP/RV.
    """
    with open(manifest_path, "r", encoding="utf-8") as f:
        manifest = json.load(f)
    with open(review_items_path, "r", encoding="utf-8") as f:
        review_items = json.load(f)

    skipped = manifest.get("skipped_rows") or []
    rows_manifest = manifest.get("rows") or []
    if not skipped or not rows_manifest:
        return 0

    # Map body_idx -> dominant matched source_file from body rows of that table
    tbl_to_src: Dict[int, Path] = {}
    tbl_to_idx: Dict[int, int] = {}
    for rm in rows_manifest:
        b_idx = rm["body_idx"]
        if b_idx not in tbl_to_src and rm.get("source_file"):
            tbl_to_src[b_idx] = Path(rm["source_file"])
            tbl_to_idx[b_idx] = rm.get("tbl_idx", 0)

    if not tbl_to_src:
        return 0

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
                    if len(inp_tcs) >= 2 and len(s_tcs) >= 3 and docx_xml_utils.get_cell_plain_text(s_tcs[2]).strip():
                        inp_cn = _WS_RE.sub("", docx_xml_utils.get_cell_plain_text(inp_tcs[0]))
                        src_cn = _WS_RE.sub("", docx_xml_utils.get_cell_plain_text(s_tcs[0]))
                        if inp_cn and (inp_cn == src_cn or inp_cn in src_cn or src_cn in inp_cn):
                            inp_en = docx_xml_utils.get_cell_plain_text(inp_tcs[1])
                            src_en = docx_xml_utils.get_cell_plain_text(s_tcs[1])
                            ops = docx_xml_utils.diff_tokens(src_en, inp_en)
                            sent_op_lists = docx_xml_utils.segment_diff_into_sentences(ops)
                            tbl_cnt = tbl_to_idx[b_idx]
                            sent_items = []
                            for s_i, s_ops in enumerate(sent_op_lists):
                                stats = docx_xml_utils.sentence_unit_stats(s_ops)
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
                                    "col3_preview": docx_xml_utils.get_cell_plain_text(s_tcs[2])[:120],
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


def _row_has_up_or_rv_changes(tr: ET.Element, up_color: str, rv_color: str) -> bool:
    if step2_render_final.is_skipped_row(tr):
        return False
    tcs = tr.findall("w:tc", NS)
    if len(tcs) < 2:
        return False
    tc2 = tcs[1]
    if tc2.find(".//w:ins", NS) is not None or tc2.find(".//w:del", NS) is not None:
        return True
    active_hls = {up_color, rv_color, "cyan", "blue", "darkCyan"}
    for hl in tc2.findall(".//w:rPr/w:highlight", NS):
        if hl.attrib.get(f"{W}val") in active_hls:
            return True
    return False


def _filter_full_article_to_translator_version(
    root: ET.Element,
    up_color: str,
    rv_color: str,
) -> int:
    """
    Transform a rendered full-article `root` into the `仅改点发译者版`:
      - Inside single-table chapters: keeps Row 0 bold chapter title row (plus Row 1 subtitle if `附篇一..四`)
        AND ONLY rows that have UP/RV changes in Col 2!
      - Inside multi-table `选段` articles: keeps main chapter title table + only `选段` tables with changes
        (pruning unchanged rows inside each kept `选段` table).
    Returns count of kept changed body rows.
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
        tcs0 = trs[0].findall("w:tc", NS)
        cn0 = docx_xml_utils.get_cell_plain_text(tcs0[0]).strip() if tcs0 else ""
        en0 = docx_xml_utils.get_cell_plain_text(tcs0[1]).strip() if len(tcs0) >= 2 else ""
        if len(trs) >= 3 and _is_blank_row(trs[2]) and (en0.endswith(":") or cn0.startswith("附篇")):
            return 2
        if len(trs) >= 2 and _is_blank_row(trs[1]):
            return 1
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
            # Always preserve Row 0 title row if single-table article
            hdr_cnt = 1

        tbl_changed = 0
        for r_i, tr in enumerate(list(trs)):
            if r_i < hdr_cnt:
                continue
            if _row_has_up_or_rv_changes(tr, up_color, rv_color):
                tbl_changed += 1
                kept_changed_rows += 1
            else:
                tbl.remove(tr)

        # In multi-table `选段` articles (len(tbls) > 1), drop unchanged `选段` sub-tables (t_idx >= 1)
        if len(tbls) > 1 and t_idx >= 1 and tbl_changed == 0:
            emptied_sub_tbls.add(id(tbl))

        # ECMA-376: ensure every remaining <w:tbl> has at least 1 <w:tr>
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
) -> dict:
    """
    Process each selected full article independently (`输出结果一篇一篇独立导出`):
      1. Extracts standalone 3-col temp `.docx` from master file;
      2. Runs full-chapter sentence-level comparison against `foreign_lib_dir`;
      3. Extracts clean English title from Row 0 Col 2 as `doc_label`;
      4. Exports TWO deliverables per article into `output_dir`:
         - 整篇完整版: `{LANG}_{EN_TITLE}_完整版_{COUNT_TAG}_{DATE}.docx`
         - 仅改点发译者版: `{LANG}_{EN_TITLE}_{COUNT_TAG}_{DATE}.docx` (Row 0 bold title + changed rows only,
           with optional `dedup_pkg_dir` duplicate removal & report).
    """
    articles_list = list(selected_articles or [])
    if not articles_list:
        raise ValueError("请先至少选择或匹配 1 篇需要对比的章节。")

    flib_p = Path(foreign_lib_dir).resolve()
    out_d = Path(output_dir).resolve()
    if not flib_p.exists() or not flib_p.is_dir():
        raise FileNotFoundError(f"找不到外文三语总库目录：{flib_p}")
    out_d.mkdir(parents=True, exist_ok=True)

    pool_or_keys, eff_rot, has_keys, effective_ai = _prepare_optional_pool(
        api_key=api_key,
        rotation_mode=rotation_mode,
        enable_ai=enable_ai,
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

    supports_full_diff_flag = _script_supports_cli_flag(
        step1_extract_and_diff, "--force-full-diff", "format_up_rv_count_tag"
    )
    supports_no_bal = _script_supports_cli_flag(
        step2_render_final, "--no-balance-columns", "balance_trilingual_table_widths"
    )
    step2_native_v17 = supports_no_bal and callable(getattr(docx_xml_utils, "format_up_rv_count_tag", None))

    saved_files: List[str] = []
    article_results: List[dict] = []
    duplicate_reports: List[str] = []
    final_lang_code = (lang or "AUTO").strip().upper()
    total_arts = len(articles_list)

    # Pre-count English titles across this batch so multiple 'Chapter 1' / 'Introduction' selections
    # in the same batch never overwrite each other's output files!
    en_title_counts: Dict[str, int] = {}
    for a_item in articles_list:
        raw_et = str(a_item.get("en_title") or a_item.get("cn_title") or "").strip().lower()
        if raw_et:
            en_title_counts[raw_et] = en_title_counts.get(raw_et, 0) + 1
    used_doc_labels: Dict[str, int] = {}

    for a_idx, art in enumerate(articles_list):
        base_pct = 5.0 + (a_idx / max(1, total_arts)) * 93.0
        span_pct = 93.0 / max(1, total_arts)
        cn_disp = art.get("cn_title") or art.get("query") or f"第{a_idx + 1}篇"

        if log_fn:
            log_fn(f"\n📖 === [{a_idx + 1}/{total_arts}] 开始全篇提取对比：《{cn_disp}》 ===")
        if progress_fn:
            progress_fn(base_pct, f"[{a_idx + 1}/{total_arts}] 正在提取《{cn_disp}》并比对外文总库...")

        tmp_root = Path(tempfile.mkdtemp(prefix=f"full_art_{a_idx + 1}_")).resolve()
        try:
            extracted_docx = tmp_root / "extracted_article.docx"
            workdir = tmp_root / "work"
            workdir.mkdir(parents=True, exist_ok=True)

            en_title, restore_heading = _extract_standalone_article_docx(art, extracted_docx)
            raw_label = en_title
            if en_title_counts.get(en_title.strip().lower(), 0) > 1:
                sec_tag = str(art.get("section_path") or "").split("›")[-1].strip()
                if sec_tag:
                    raw_label = f"{en_title} ({sec_tag})"
            doc_label = docx_xml_utils.sanitize_filename_part(raw_label, fallback="Article", max_len=120)
            lbl_key = doc_label.lower()
            if lbl_key in used_doc_labels:
                used_doc_labels[lbl_key] += 1
                doc_label = docx_xml_utils.sanitize_filename_part(
                    f"{doc_label}_{used_doc_labels[lbl_key]}", fallback="Article", max_len=120
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

            # 2. Step 1b: Evaluate UP/RV & render draft
            def _s1_prog(frac: float, msg: str):
                if progress_fn:
                    progress_fn(base_pct + span_pct * (0.15 + frac * 0.35), f"[{a_idx + 1}/{total_arts}] {msg}")

            s1_stats: dict = {}
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

            draft_path = tmp_root / "draft_article.docx"
            argv1b = [
                "step1_render_draft.py",
                "--manifest", str(manifest_path),
                "--decisions", str(decisions_path),
                "--output", str(draft_path),
            ]
            _run_module_main(step1_render_draft.main, argv1b, log_fn=log_fn)

            # Restore clean original <w:p> title if we temporarily prepended section qualifier for scoping
            if restore_heading is not None:
                p_idx_restore, orig_ptxt = restore_heading
                d_root, _, d_ns = docx_xml_utils.load_docx_xml(draft_path)
                d_cmts = docx_xml_utils.load_docx_comments_xml(draft_path)
                d_body = d_root.find("w:body", NS)
                if d_body is not None and p_idx_restore < len(d_body):
                    t_els = d_body[p_idx_restore].findall(".//w:t", NS)
                    if t_els:
                        t_els[0].text = orig_ptxt
                        for et in t_els[1:]:
                            et.text = ""
                        docx_xml_utils.save_docx_xml(draft_path, draft_path, d_root, d_ns, comments_root=d_cmts)

            # 3. Step 2a + 2b: Parse spans & align Col 3
            argv2a = [
                "step2_parse_reviewed.py",
                "--draft", str(draft_path),
                "--workdir", str(workdir),
            ]
            _run_module_main(step2_parse_reviewed.main, argv2a, log_fn=log_fn)

            tasks_path = workdir / "step2_align_tasks.json"
            tasks_data = json.loads(tasks_path.read_text(encoding="utf-8"))

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

            # 4A. Deliverable 1: 整篇完整版 `{LANG}_{EN_TITLE}_完整版_{COUNT_TAG}_{DATE}.docx`
            full_label = f"{doc_label}_完整版"
            argv_full = [
                "step2_render_final.py",
                "--tasks", str(tasks_path),
                "--alignments", str(alignments_path),
                "--outdir", str(out_d),
                "--lang", det_lang,
                "--doc-name", full_label,
                "--date", final_date,
            ]
            if not balance_columns and supports_no_bal:
                argv_full.append("--no-balance-columns")
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
            )
            saved_files.append(str(full_final_path))

            # 4B. Deliverable 2: 仅改点发译者版 `{LANG}_{EN_TITLE}_{COUNT_TAG}_{DATE}.docx`
            # Build directly from `full_final_path` so Col 3 highlights & comments are already applied,
            # keeping Row 0 bold title row + ONLY rows with UP/RV changes!
            trans_root, _, trans_ns = docx_xml_utils.load_docx_xml(full_final_path)
            trans_cmts = docx_xml_utils.load_docx_comments_xml(full_final_path)
            kept_chg_rows = _filter_full_article_to_translator_version(trans_root, col2_up, col2_rv)
            if balance_columns:
                balance_trilingual_table_widths(trans_root)

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
            trans_path = docx_xml_utils.safe_join_output_file(out_d, trans_name)
            docx_xml_utils.save_docx_xml(full_final_path, trans_path, trans_root, trans_ns, comments_root=trans_cmts)

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
                trans_up = int(d_res.get("up_words", trans_up))
                trans_rv = int(d_res.get("rv_words", trans_rv))
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
                    recalc_p = docx_xml_utils.safe_join_output_file(out_d, recalc_name)
                    if os.path.normcase(str(recalc_p.resolve())) != os.path.normcase(str(trans_path.resolve())):
                        if recalc_p.exists():
                            recalc_p.unlink()
                        trans_path.rename(recalc_p)
                        trans_path = recalc_p

            saved_files.append(str(trans_path))

            if log_fn:
                log_fn(f"  ✅ 《{cn_disp}》导出完成：")
                log_fn(f"     - 整篇完整版：{full_final_path.name}")
                log_fn(f"     - 仅改点发译者版：{trans_path.name}（保留改点行 {kept_chg_rows} 段）")

            article_results.append({
                "cn_title": cn_disp,
                "en_title": en_title,
                "lang_code": det_lang,
                "full_path": str(full_final_path),
                "translator_path": str(trans_path),
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

    if progress_fn:
        progress_fn(100.0, f"全篇对比整理完成：共处理 {total_arts} 篇章节，导出 {len(saved_files)} 份文档！")

    return {
        "saved_files": saved_files,
        "article_results": article_results,
        "lang_code": final_lang_code if final_lang_code != "AUTO" else "XX",
        "duplicate_reports": duplicate_reports,
    }
