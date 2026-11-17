from concurrent.futures import ThreadPoolExecutor
import copy
from datetime import datetime
import difflib
import os
from pathlib import Path
import re
import sys
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
from lang_registry import load_languages  # noqa: E402
from word_counter import (  # noqa: E402
    balance_trilingual_table_widths,
    detect_lang_from_text,
    format_counted_docx_name,
    format_up_rv_count_tag,
    is_top_highlight_instruction_text,
    parse_doc_filename,
    update_top_highlight_instruction_lines,
)

W_URI = docx_xml_utils.W_URI
R_URI = docx_xml_utils.R_URI
XML_URI = docx_xml_utils.XML_URI
PKG_REL_URI = docx_xml_utils.PKG_REL_URI
CONTENT_TYPES_URI = docx_xml_utils.CONTENT_TYPES_URI
W = docx_xml_utils.W
NS = docx_xml_utils.NS
XML_SPACE = docx_xml_utils.XML_SPACE

_WS_RE = re.compile(r"[\s\u3000\u200b\u200c\ufeff]+")
_LOOSE_KEEP_RE = re.compile(r"[\u4e00-\u9fff\u3400-\u4dbfA-Za-z0-9]+")
_CN_SENT_SPLIT_RE = re.compile(r"(?<=[。？！!?；;\n])(?:[”’\"'」』）\)]*)")
_EN_WORD_RE = re.compile(r"[A-Za-z0-9]+(?:[’'][A-Za-z0-9]+)*")
_DEDUP_REPORT_PREFIX = "查重结果"

_EN_STOPWORDS = frozenset({
    "the", "and", "that", "this", "with", "from", "have", "they", "their",
    "your", "what", "when", "will", "into", "only", "even", "more", "than",
    "some", "such", "were", "been", "which", "there", "would", "could",
    "should", "about", "for", "are", "was", "not", "but", "you", "all",
    "can", "had", "her", "his", "one", "our", "out", "who", "has", "him",
})

_PKG_CACHE_LOCK = threading.Lock()
_PKG_INDEX_CACHE: Dict[str, Tuple[tuple, "UnmergedPackageIndex"]] = {}


# ---------------------------------------------------------------------------
# Text Normalization & Colored Segment Extraction
# ---------------------------------------------------------------------------

def normalize_cn(s: str) -> str:
    return _WS_RE.sub("", s or "")


def normalize_cn_loose(s: str) -> str:
    return "".join(_LOOSE_KEEP_RE.findall(s or "")).lower()


def split_cn_sentences(text: str) -> List[str]:
    if not text:
        return []
    parts = _CN_SENT_SPLIT_RE.split(text.strip())
    return [ns for p in parts if (ns := normalize_cn(p))]


def best_local_cn_similarity(needle: str, haystack: str) -> float:
    """Fast local substring similarity of `needle` inside `haystack` on loose CJK/alnum text."""
    q = normalize_cn_loose(needle)
    h = normalize_cn_loose(haystack)
    if not q or not h:
        return 0.0
    if q in h:
        return 1.0
    if h in q:
        return len(h) / len(q)

    lq = len(q)
    lh = len(h)
    if lh <= int(lq * 1.35) + 12:
        sm = difflib.SequenceMatcher(None, q, h, autojunk=False)
        matched = sum(blk.size for blk in sm.get_matching_blocks())
        return matched / lq

    win_len = min(lh, int(lq * 1.25) + 8)
    step = max(1, lq // 4)
    q_grams = {q[i : i + 3] for i in range(max(0, lq - 2))}
    best_starts = [0]
    if q_grams:
        scored_wins = []
        pos = 0
        while pos < lh:
            sub = h[pos : pos + win_len]
            hits = sum(1 for g in q_grams if g in sub)
            scored_wins.append((hits, pos))
            if pos + win_len >= lh:
                break
            pos += step
        scored_wins.sort(key=lambda x: x[0], reverse=True)
        if scored_wins and scored_wins[0][0] == 0:
            return 0.0
        best_starts = [p for _, p in scored_wins[:2]]

    best_sim = 0.0
    for st in best_starts:
        sub = h[st : st + win_len]
        sm = difflib.SequenceMatcher(None, q, sub, autojunk=False)
        matched = sum(blk.size for blk in sm.get_matching_blocks())
        sim = matched / lq
        if sim > best_sim:
            best_sim = sim
            if best_sim >= 0.98:
                break
    return best_sim


def _has_visual_color_on_props(prop_el: Optional[ET.Element]) -> bool:
    if prop_el is None:
        return False
    shd = prop_el.find("w:shd", NS)
    if shd is not None:
        fill = (shd.attrib.get(f"{W}fill") or "").strip().lower()
        val = (shd.attrib.get(f"{W}val") or "").strip().lower()
        col = (shd.attrib.get(f"{W}color") or "").strip().lower()
        if fill not in ("", "auto", "ffffff", "none"):
            return True
        if val not in ("", "clear", "nil") and col not in ("", "auto", "ffffff", "none"):
            return True

    hl = prop_el.find("w:highlight", NS)
    if hl is not None:
        hval = (hl.attrib.get(f"{W}val") or "").strip().lower()
        if hval not in ("", "none", "white"):
            return True

    col_el = prop_el.find("w:color", NS)
    if col_el is not None:
        cval = (col_el.attrib.get(f"{W}val") or "").strip().lower()
        theme = (col_el.attrib.get(f"{W}themeColor") or "").strip().lower()
        if cval not in ("", "auto", "000000", "black", "ffffff"):
            return True
        if theme and theme not in ("text1", "dark1", "background1", "light1"):
            return True
    return False


def extract_colored_segments(tc: Optional[ET.Element]) -> List[str]:
    """Extract visually colored runs/sentences (`w:shd`, `w:highlight`, or non-black `w:color`) from cell `tc`."""
    if tc is None:
        return []
    tc_colored = _has_visual_color_on_props(tc.find("w:tcPr", NS))
    char_list: List[str] = []
    flag_list: List[bool] = []

    for p_i, p in enumerate(tc.findall("w:p", NS)):
        if p_i > 0:
            char_list.append("\n")
            flag_list.append(False)
        pPr = p.find("w:pPr", NS)
        p_colored = _has_visual_color_on_props(pPr) or (
            pPr is not None and _has_visual_color_on_props(pPr.find("w:rPr", NS))
        )
        for r in p.findall(".//w:r", NS):
            txt = "".join(t.text or "" for t in r.findall("w:t", NS))
            if not txt:
                continue
            rPr = r.find("w:rPr", NS)
            is_col = p_colored or _has_visual_color_on_props(rPr)
            for ch in txt:
                char_list.append(ch)
                flag_list.append(is_col)

    full_txt = "".join(char_list)
    if not full_txt.strip():
        return []

    segments: List[str] = []
    if any(flag_list):
        n = len(char_list)
        i = 0
        while i < n:
            if flag_list[i]:
                j = i + 1
                while j < n:
                    if flag_list[j]:
                        j += 1
                    elif (
                        j + 1 < n
                        and flag_list[j + 1]
                        and char_list[j] not in "\n。？！.!?"
                        and char_list[j].strip() in ("", "，", ",", "、", "“", "”", '"', "'", "’")
                    ):
                        j += 1
                    else:
                        break
                seg = "".join(char_list[i:j]).strip()
                if seg:
                    segments.append(seg)
                    seg_loose = normalize_cn_loose(seg)
                    if 2 <= len(seg_loose) < 12:
                        s_bound = i
                        while s_bound > 0 and char_list[s_bound - 1] not in "。？！!?\n；;":
                            s_bound -= 1
                        e_bound = j
                        while e_bound < n and char_list[e_bound] not in "。？！!?\n；;":
                            e_bound += 1
                        if e_bound < n and char_list[e_bound] in "。？！!?；;":
                            e_bound += 1
                        enclosing = "".join(char_list[s_bound:e_bound]).strip()
                        if enclosing and enclosing != seg:
                            segments.append(enclosing)
                i = j
            else:
                i += 1
    elif tc_colored:
        segments.append(full_txt.strip())

    dedup: List[str] = []
    seen = set()
    for s in sorted(segments, key=lambda x: len(normalize_cn_loose(x)), reverse=True):
        key = normalize_cn(s)
        if key and key not in seen:
            seen.add(key)
            dedup.append(s)
    return dedup


def _en_lexical_overlap(en_a: str, en_b: str) -> float:
    words_a = {
        m.group(0).lower()
        for m in _EN_WORD_RE.finditer(en_a or "")
        if len(m.group(0)) >= 4 and m.group(0).lower() not in _EN_STOPWORDS
    }
    words_b = {
        m.group(0).lower()
        for m in _EN_WORD_RE.finditer(en_b or "")
        if len(m.group(0)) >= 4 and m.group(0).lower() not in _EN_STOPWORDS
    }
    if not words_a or not words_b:
        return 0.5
    return len(words_a & words_b) / max(1, min(len(words_a), len(words_b)))


# ---------------------------------------------------------------------------
# Row Classification & Unmerged Package Multi-Threaded 3-Gram Index
# ---------------------------------------------------------------------------

def _is_blank_tr(tr: ET.Element) -> bool:
    tcs = tr.findall("w:tc", NS)
    return bool(tcs) and all(not docx_xml_utils.get_cell_plain_text(tc).strip() for tc in tcs)


def _row_has_col2_revisions_or_highlights(tr: ET.Element) -> bool:
    tcs = tr.findall("w:tc", NS)
    if len(tcs) < 2:
        return False
    tc2 = tcs[1]
    if tc2.find(".//w:ins", NS) is not None or tc2.find(".//w:del", NS) is not None:
        return True
    for hl in tc2.findall(".//w:rPr/w:highlight", NS):
        hval = (hl.attrib.get(f"{W}val") or "").strip().lower()
        if hval in ("yellow", "cyan", "blue", "darkcyan"):
            return True
    return False


def _header_row_count_in_table(trs: List[ET.Element], preceding_title: str) -> int:
    """
    Detect leading unedited chapter-title header row(s) in a full-article style table
    (e.g. Row 0 = title, Row 1 = blank; or Row 0 = '附篇一', Row 1 = subtitle, Row 2 = blank)
    so chapter titles are never falsely deleted as 'duplicate paragraphs'.
    """
    if len(trs) < 2:
        return 0
    tr0 = trs[0]
    if _row_has_col2_revisions_or_highlights(tr0):
        return 0
    tcs0 = tr0.findall("w:tc", NS)
    if not tcs0:
        return 0
    cn0 = normalize_cn(docx_xml_utils.get_cell_plain_text(tcs0[0]))
    if not cn0 or len(cn0) > 60:
        return 0

    # 2-row Appendix header: Row 0 ('附篇一'), Row 1 (subtitle), Row 2 (blank)
    if len(trs) >= 3 and _is_blank_tr(trs[2]) and not _row_has_col2_revisions_or_highlights(trs[1]):
        tcs1 = trs[1].findall("w:tc", NS)
        cn1 = normalize_cn(docx_xml_utils.get_cell_plain_text(tcs1[0])) if tcs1 else ""
        if cn1 and len(cn1) <= 60:
            return 2

    if _is_blank_tr(trs[1]):
        return 1
    if preceding_title and cn0 == normalize_cn(preceding_title):
        return 1
    return 0


def _scan_pkg_docx_files(pkg_dir: Path) -> List[Tuple[str, str, int, int]]:
    """Return sorted list of `(rel_posix, abs_path_str, mtime_ns, size)` for valid `.docx` files in `pkg_dir`."""
    results: List[Tuple[str, str, int, int]] = []

    def _walk(cur_dir: str, rel_prefix: str):
        try:
            with os.scandir(cur_dir) as it:
                entries = list(it)
        except OSError:
            return
        for entry in entries:
            ename = entry.name
            if entry.is_dir(follow_symlinks=False):
                if ename.startswith("."):
                    continue
                nxt = f"{rel_prefix}/{ename}" if rel_prefix else ename
                _walk(entry.path, nxt)
            elif entry.is_file(follow_symlinks=False):
                if not ename.lower().endswith(".docx"):
                    continue
                if ename.startswith("~$") or ename.startswith(".") or ename.startswith("初筛_"):
                    continue
                if ename.startswith(_DEDUP_REPORT_PREFIX):
                    continue
                rel_posix = f"{rel_prefix}/{ename}" if rel_prefix else ename
                try:
                    st = os.stat(entry.path)
                except OSError:
                    continue
                results.append((rel_posix, str(Path(entry.path).resolve()), st.st_mtime_ns, st.st_size))

    _walk(str(pkg_dir), "")
    results.sort(key=lambda x: x[0])
    return results


def _parse_single_pkg_docx(task: Tuple[str, str, int, int]) -> List[dict]:
    rel_posix, abs_path_str, _mtime_ns, _size = task
    p = Path(abs_path_str)
    try:
        root, _, _ = docx_xml_utils.load_docx_xml(p)
    except Exception:
        return []

    body = root.find("w:body", NS)
    if body is None:
        return []

    fallback_title = p.stem
    cur_title = ""
    tbl_no = 0
    rows_out: List[dict] = []

    for ch in body:
        loc = ch.tag.split("}")[-1]
        if loc == "p":
            ptxt = "".join(t.text or "" for t in ch.findall(".//w:t", NS)).strip()
            if ptxt and not is_top_highlight_instruction_text(ptxt) and len(ptxt) <= 200:
                cur_title = ptxt
        elif loc == "tbl":
            tbl_no += 1
            trs = ch.findall("w:tr", NS)
            hdr_cnt = _header_row_count_in_table(trs, cur_title)
            eff_title = cur_title
            if not eff_title and hdr_cnt >= 1:
                tcs0 = trs[0].findall("w:tc", NS)
                if tcs0:
                    eff_title = docx_xml_utils.get_cell_plain_text(tcs0[0]).strip()
            if not eff_title:
                eff_title = fallback_title

            for r_idx, tr in enumerate(trs):
                if r_idx < hdr_cnt:
                    continue
                tcs = tr.findall("w:tc", NS)
                if not tcs:
                    continue
                if len(tcs) >= 3 and docx_xml_utils.get_cell_plain_text(tcs[2]).strip() == docx_xml_utils.SKIPPED_ROW_MARKER:
                    continue
                cn_raw = docx_xml_utils.get_cell_plain_text(tcs[0]).strip()
                cn_norm = normalize_cn(cn_raw)
                cn_loose = normalize_cn_loose(cn_raw)
                if len(cn_loose) < 6:
                    continue

                en_raw = docx_xml_utils.get_cell_plain_text(tcs[1]).strip() if len(tcs) >= 2 else ""
                col3_raw = docx_xml_utils.get_cell_plain_text(tcs[2]).strip() if len(tcs) >= 3 else ""

                sents = split_cn_sentences(cn_raw)
                open_sents = sents[:3] if sents else [cn_norm[:80]]
                close_sents = sents[-3:] if sents else [cn_norm[-80:]]
                open_loose = normalize_cn_loose("".join(open_sents))
                close_loose = normalize_cn_loose("".join(close_sents))

                colored_cn = [
                    seg for seg in extract_colored_segments(tcs[0])
                    if len(normalize_cn_loose(seg)) >= 10
                ]

                rows_out.append({
                    "abs_path": abs_path_str,
                    "abs_normcase": os.path.normcase(abs_path_str),
                    "file_name": p.name,
                    "rel_path": rel_posix,
                    "article_title": eff_title,
                    "table_no": tbl_no,
                    "row_no": r_idx + 1,
                    "cn_raw": cn_raw,
                    "en_raw": en_raw,
                    "col3_raw": col3_raw,
                    "cn_norm": cn_norm,
                    "cn_loose": cn_loose,
                    "open_sents": open_sents,
                    "close_sents": close_sents,
                    "open_loose": open_loose,
                    "close_loose": close_loose,
                    "colored_cn_segs": colored_cn,
                })

    return rows_out


class UnmergedPackageIndex:
    """Fast 3-gram inverted index over all `.docx` paragraphs in `未入库改点包`."""

    def __init__(self, pkg_dir: Path, file_count: int, rows: List[dict], elapsed_sec: float):
        self.pkg_dir = pkg_dir
        self.file_count = file_count
        self.rows = rows
        self.elapsed_sec = elapsed_sec

        exact_map: Dict[str, List[int]] = {}
        postings: Dict[str, List[int]] = {}
        colored_entries: List[Tuple[str, int]] = []

        for idx, r in enumerate(rows):
            cl = r["cn_loose"]
            exact_map.setdefault(cl, []).append(idx)
            n = len(cl)
            if n >= 3:
                seen_g = set()
                for p in range(n - 2):
                    g = cl[p : p + 3]
                    if g not in seen_g:
                        seen_g.add(g)
                        postings.setdefault(g, []).append(idx)
            for cseg in r["colored_cn_segs"]:
                c_loose = normalize_cn_loose(cseg)
                if len(c_loose) >= 10:
                    colored_entries.append((c_loose, idx))

        self.exact_loose_map = exact_map
        self.trigram_postings = postings
        self.colored_entries = colored_entries

    def retrieve_candidates(
        self,
        query_text: str,
        top_k: int = 30,
        exclude_normcase: str = "",
    ) -> List[dict]:
        q_loose = normalize_cn_loose(query_text)
        if len(q_loose) < 3:
            return []
        q_grams = {q_loose[i : i + 3] for i in range(len(q_loose) - 2)}
        if not q_grams:
            return []

        counts: Dict[int, int] = {}
        postings = self.trigram_postings
        for g in q_grams:
            plist = postings.get(g)
            if plist:
                for r_i in plist:
                    if exclude_normcase and self.rows[r_i]["abs_normcase"] == exclude_normcase:
                        continue
                    counts[r_i] = counts.get(r_i, 0) + 1
        if not counts:
            return []
        ranked = sorted(counts.items(), key=lambda kv: kv[1], reverse=True)[:top_k]
        return [self.rows[r_i] for r_i, _ in ranked]


def get_unmerged_package_index(
    unmerged_pkg_dir: Union[str, Path],
    log_fn: Optional[Callable[[str], None]] = None,
    max_workers: int = 8,
) -> UnmergedPackageIndex:
    t0 = time.perf_counter()
    pkg_p = Path(unmerged_pkg_dir).resolve()
    if not pkg_p.exists() or not pkg_p.is_dir():
        raise FileNotFoundError(f"找不到未入库改点包目录：{pkg_p}")

    norm_key = os.path.normcase(str(pkg_p))
    file_stats = _scan_pkg_docx_files(pkg_p)
    stat_sig = tuple((rp, mt, sz) for rp, _, mt, sz in file_stats)

    with _PKG_CACHE_LOCK:
        cached = _PKG_INDEX_CACHE.get(norm_key)
        if cached is not None and cached[0] == stat_sig:
            return cached[1]

    all_rows: List[dict] = []
    if file_stats:
        workers = max(1, min(int(max_workers or 8), len(file_stats)))
        if workers == 1:
            parsed_lists = [_parse_single_pkg_docx(t) for t in file_stats]
        else:
            with ThreadPoolExecutor(max_workers=workers) as ex:
                parsed_lists = list(ex.map(_parse_single_pkg_docx, file_stats))
        for rlist in parsed_lists:
            all_rows.extend(rlist)

    elapsed = round(time.perf_counter() - t0, 4)
    idx_obj = UnmergedPackageIndex(pkg_p, len(file_stats), all_rows, elapsed)
    with _PKG_CACHE_LOCK:
        _PKG_INDEX_CACHE[norm_key] = (stat_sig, idx_obj)

    if log_fn:
        log_fn(
            f"📦 未入库改点包索引就绪：共扫描 {len(file_stats)} 篇文档、{len(all_rows)} 个中文段落（耗时 {elapsed:.3f}s）"
        )
    return idx_obj


# ---------------------------------------------------------------------------
# 4-Stage Col 1 Chinese Duplicate Matcher
# ---------------------------------------------------------------------------

def match_row_4stage(
    tc1: ET.Element,
    tc2: Optional[ET.Element],
    inp_cn: str,
    inp_en: str,
    pkg_index: UnmergedPackageIndex,
    exclude_normcase: str = "",
) -> Optional[Tuple[dict, float, str]]:
    """
    4-Stage Col 1 Chinese duplicate detection against `pkg_index`:
      Stage 1: Full paragraph exact / high similarity (>= 0.88)
      Stage 2: Opening 1-3 sentences (`段首前1-3句`)
      Stage 3: Closing 1-3 sentences (`段尾后1-3句`)
      Stage 4: Colored sentence/run segments (`w:shd` / `w:highlight` / `w:color`)
    Returns `(matched_pkg_row, similarity, stage_label)` or `None`.
    """
    inp_loose = normalize_cn_loose(inp_cn)
    inp_norm = normalize_cn(inp_cn)
    if len(inp_loose) < 6 or not pkg_index.rows:
        return None

    # ------------------------------------------------------------------
    # Stage 1: Full paragraph exact or high similarity (>= 0.88)
    # ------------------------------------------------------------------
    exact_ids = pkg_index.exact_loose_map.get(inp_loose, ())
    for r_i in exact_ids:
        cand = pkg_index.rows[r_i]
        if exclude_normcase and cand["abs_normcase"] == exclude_normcase:
            continue
        return cand, 1.0, "第1级：整段精确匹配"

    full_cands = pkg_index.retrieve_candidates(inp_cn, top_k=30, exclude_normcase=exclude_normcase)
    best_s1 = None
    best_s1_sim = 0.0
    for cand in full_cands:
        cl = cand["cn_loose"]
        len_ratio = min(len(inp_loose), len(cl)) / max(1, max(len(inp_loose), len(cl)))
        if len_ratio < 0.78:
            continue
        sim = difflib.SequenceMatcher(None, inp_loose, cl, autojunk=False).ratio()
        if sim >= 0.88 and sim > best_s1_sim:
            best_s1_sim = sim
            best_s1 = cand
    if best_s1 is not None:
        return best_s1, round(best_s1_sim, 4), "第1级：整段高相似匹配"

    # ------------------------------------------------------------------
    # Stage 2: Opening 1-3 sentences (`段首前1-3句`)
    # ------------------------------------------------------------------
    cn_sents = split_cn_sentences(inp_cn)
    open_queries: List[str] = []
    for k_cnt in (3, 2, 1):
        if len(cn_sents) >= k_cnt:
            q_s = "".join(cn_sents[:k_cnt])
            if len(normalize_cn_loose(q_s)) >= 14 and q_s not in open_queries:
                open_queries.append(q_s)
    if not open_queries and len(inp_norm) >= 36:
        open_queries.append(inp_norm[:45])

    best_s2 = None
    best_s2_sim = 0.0
    for oq in open_queries:
        oq_loose = normalize_cn_loose(oq)
        oq_len = len(oq_loose)
        thresh = 0.90 if oq_len >= 20 else 0.94
        cands = pkg_index.retrieve_candidates(oq, top_k=25, exclude_normcase=exclude_normcase)
        for cand in cands:
            head_len = max(len(cand["open_loose"]), int(oq_len * 1.35) + 18)
            head_slice = cand["cn_loose"][:head_len]
            sim = 1.0 if oq_loose in head_slice else best_local_cn_similarity(oq_loose, head_slice)
            if sim >= thresh and sim > best_s2_sim:
                best_s2_sim = sim
                best_s2 = cand
        if best_s2 is not None and best_s2_sim >= 0.96:
            break

    # Also check if a package row's opening 1-3 sentences match input's opening (when package row is shorter/longer)
    if best_s2 is None and full_cands:
        for cand in full_cands[:15]:
            co_loose = cand["open_loose"]
            if len(co_loose) >= 16:
                inp_head = inp_loose[: max(len(co_loose) + 20, int(len(co_loose) * 1.35))]
                sim = 1.0 if co_loose in inp_head else best_local_cn_similarity(co_loose, inp_head)
                if sim >= 0.92 and sim > best_s2_sim:
                    best_s2_sim = sim
                    best_s2 = cand

    if best_s2 is not None:
        return best_s2, round(best_s2_sim, 4), "第2级：段首前1-3句匹配"

    # ------------------------------------------------------------------
    # Stage 3: Closing 1-3 sentences (`段尾后1-3句`)
    # ------------------------------------------------------------------
    close_queries: List[str] = []
    for k_cnt in (3, 2, 1):
        if len(cn_sents) >= k_cnt:
            q_c = "".join(cn_sents[-k_cnt:])
            if len(normalize_cn_loose(q_c)) >= 14 and q_c not in close_queries:
                close_queries.append(q_c)
    if not close_queries and len(inp_norm) >= 36:
        close_queries.append(inp_norm[-45:])

    best_s3 = None
    best_s3_sim = 0.0
    for cq in close_queries:
        cq_loose = normalize_cn_loose(cq)
        cq_len = len(cq_loose)
        thresh = 0.90 if cq_len >= 20 else 0.94
        cands = pkg_index.retrieve_candidates(cq, top_k=25, exclude_normcase=exclude_normcase)
        for cand in cands:
            tail_len = max(len(cand["close_loose"]), int(cq_len * 1.35) + 18)
            tail_slice = cand["cn_loose"][-tail_len:]
            sim = 1.0 if cq_loose in tail_slice else best_local_cn_similarity(cq_loose, tail_slice)
            if sim >= thresh and sim > best_s3_sim:
                best_s3_sim = sim
                best_s3 = cand
        if best_s3 is not None and best_s3_sim >= 0.96:
            break

    if best_s3 is None and full_cands:
        for cand in full_cands[:15]:
            cc_loose = cand["close_loose"]
            if len(cc_loose) >= 16:
                inp_tail = inp_loose[-max(len(cc_loose) + 20, int(len(cc_loose) * 1.35)) :]
                sim = 1.0 if cc_loose in inp_tail else best_local_cn_similarity(cc_loose, inp_tail)
                if sim >= 0.92 and sim > best_s3_sim:
                    best_s3_sim = sim
                    best_s3 = cand

    if best_s3 is not None:
        return best_s3, round(best_s3_sim, 4), "第3级：段尾后1-3句匹配"

    # ------------------------------------------------------------------
    # Stage 4: Colored sentence segments (`w:shd` / `w:highlight` / `w:color`)
    # ------------------------------------------------------------------
    best_s4 = None
    best_s4_sim = 0.0

    inp_colored_cn = extract_colored_segments(tc1)
    for cseg in inp_colored_cn:
        c_loose = normalize_cn_loose(cseg)
        clen = len(c_loose)
        if clen < 10:
            continue
        cands = pkg_index.retrieve_candidates(cseg, top_k=25, exclude_normcase=exclude_normcase)
        for cand in cands:
            sim = 1.0 if c_loose in cand["cn_loose"] else (
                best_local_cn_similarity(c_loose, cand["cn_loose"]) if clen >= 16 else 0.0
            )
            if sim >= 0.90 and sim > best_s4_sim:
                if clen >= 16 or _en_lexical_overlap(inp_en, cand["en_raw"]) >= 0.20:
                    best_s4_sim = sim
                    best_s4 = cand

    # Reverse check: package row has colored Col 1 Chinese segment contained in input Col 1
    if best_s4 is None and full_cands:
        for cand in full_cands[:20]:
            for pkg_cseg in cand["colored_cn_segs"]:
                pc_loose = normalize_cn_loose(pkg_cseg)
                if len(pc_loose) < 10:
                    continue
                sim = 1.0 if pc_loose in inp_loose else (
                    best_local_cn_similarity(pc_loose, inp_loose) if len(pc_loose) >= 16 else 0.0
                )
                if sim >= 0.90 and sim > best_s4_sim:
                    if len(pc_loose) >= 16 or _en_lexical_overlap(inp_en, cand["en_raw"]) >= 0.20:
                        best_s4_sim = sim
                        best_s4 = cand

    # Secondary check: Col 2 highlighted English sentence mapped to middle Chinese sentence
    if best_s4 is None and tc2 is not None and len(cn_sents) >= 3:
        col2_segs = extract_colored_segments(tc2)
        if col2_segs and inp_en:
            en_low = inp_en.lower()
            for eseg in col2_segs[:3]:
                pos = en_low.find(eseg.lower()[:30])
                if pos == -1:
                    continue
                rel_ratio = pos / max(1, len(inp_en))
                est_idx = min(len(cn_sents) - 1, max(0, int(round(rel_ratio * (len(cn_sents) - 1)))))
                for s_idx in {est_idx, max(0, est_idx - 1), min(len(cn_sents) - 1, est_idx + 1)}:
                    msent = cn_sents[s_idx]
                    m_loose = normalize_cn_loose(msent)
                    if len(m_loose) < 16:
                        continue
                    cands = pkg_index.retrieve_candidates(msent, top_k=15, exclude_normcase=exclude_normcase)
                    for cand in cands:
                        sim = 1.0 if m_loose in cand["cn_loose"] else best_local_cn_similarity(m_loose, cand["cn_loose"])
                        if sim >= 0.93 and _en_lexical_overlap(inp_en, cand["en_raw"]) >= 0.25 and sim > best_s4_sim:
                            best_s4_sim = sim
                            best_s4 = cand

    if best_s4 is not None:
        return best_s4, round(best_s4_sim, 4), "第4级：带颜色句子片段匹配"

    return None


# ---------------------------------------------------------------------------
# Standalone Landscape A4 OOXML Report Builder
# ---------------------------------------------------------------------------

def _sanitize_snapshot_tc(src_tc: Optional[ET.Element], rev_id_counter: List[int]) -> ET.Element:
    """
    Clone a table cell for the report's 3-column snapshot table, preserving inline formatting,
    highlights (`yellow` / `cyan` / `lightGray`), and `<w:ins>` / `<w:del>` track changes while
    stripping external references (comments, bookmarks, hyperlinks, style IDs, paraId).
    """
    tc_out = ET.Element(f"{W}tc")
    tcPr = ET.SubElement(tc_out, f"{W}tcPr")
    bdr = ET.SubElement(tcPr, f"{W}tcBorders")
    for side in ("top", "left", "bottom", "right"):
        ET.SubElement(
            bdr,
            f"{W}{side}",
            {f"{W}val": "single", f"{W}sz": "4", f"{W}space": "0", f"{W}color": "BFBFBF"},
        )

    if src_tc is None:
        ET.SubElement(tc_out, f"{W}p")
        return tc_out

    allowed_rpr = {
        "rFonts", "b", "bCs", "i", "iCs", "u", "strike", "color",
        "sz", "szCs", "highlight", "shd", "rtl", "cs", "lang",
    }

    def _clean_rpr(rpr_in: Optional[ET.Element]) -> ET.Element:
        rpr_out = ET.Element(f"{W}rPr")
        ET.SubElement(
            rpr_out,
            f"{W}rFonts",
            {f"{W}ascii": "Georgia", f"{W}eastAsia": "宋体", f"{W}hAnsi": "Georgia", f"{W}cs": "Georgia"},
        )
        ET.SubElement(rpr_out, f"{W}sz", {f"{W}val": "21"})
        ET.SubElement(rpr_out, f"{W}szCs", {f"{W}val": "21"})
        if rpr_in is not None:
            for ch in rpr_in:
                loc = ch.tag.split("}")[-1]
                if loc in allowed_rpr:
                    cp_ch = ET.Element(f"{W}{loc}", {
                        k: v for k, v in ch.attrib.items()
                        if k.startswith(f"{{{W_URI}}}") or k.startswith(f"{{{XML_URI}}}") or not k.startswith("{")
                    })
                    docx_xml_utils.insert_rpr_child_ordered(rpr_out, cp_ch)
        return rpr_out

    def _clean_run(r_in: ET.Element, is_del: bool = False) -> Optional[ET.Element]:
        t_tag = "delText" if is_del else "t"
        txt_parts = []
        for sub in r_in:
            sloc = sub.tag.split("}")[-1]
            if sloc == t_tag and sub.text:
                txt_parts.append(sub.text)
            elif not is_del and sloc == "t" and sub.text:
                txt_parts.append(sub.text)
            elif sloc == "tab":
                txt_parts.append("\t")
            elif sloc in ("br", "cr"):
                txt_parts.append("\n")
        txt = "".join(txt_parts)
        if not txt:
            return None
        r_out = ET.Element(f"{W}r")
        r_out.append(_clean_rpr(r_in.find("w:rPr", NS)))
        t_el = ET.Element(f"{W}{'delText' if is_del else 't'}")
        if txt.startswith(" ") or txt.endswith(" ") or "  " in txt:
            t_el.set(XML_SPACE, "preserve")
        t_el.text = txt
        r_out.append(t_el)
        return r_out

    def _walk_inline(container: ET.Element, dst_p: ET.Element) -> None:
        for ch in container:
            loc = ch.tag.split("}")[-1]
            if loc == "r":
                cr = _clean_run(ch, is_del=False)
                if cr is not None:
                    dst_p.append(cr)
            elif loc in ("ins", "del"):
                is_d = loc == "del"
                sub_runs = []
                for r_sub in ch.findall(".//w:r", NS):
                    cr = _clean_run(r_sub, is_del=is_d)
                    if cr is not None:
                        sub_runs.append(cr)
                if sub_runs:
                    rev_id_counter[0] += 1
                    tr_el = ET.Element(
                        f"{W}{loc}",
                        {
                            f"{W}id": str(rev_id_counter[0]),
                            f"{W}author": ch.attrib.get(f"{W}author", "整理组"),
                            f"{W}date": ch.attrib.get(f"{W}date", "2026-01-01T00:00:00Z"),
                        },
                    )
                    for sr in sub_runs:
                        tr_el.append(sr)
                    dst_p.append(tr_el)
            elif loc in ("hyperlink", "smartTag", "customXml", "fldSimple", "sdtContent"):
                _walk_inline(ch, dst_p)
            elif loc == "sdt":
                sc = ch.find("w:sdtContent", NS)
                if sc is not None:
                    _walk_inline(sc, dst_p)

    for p_in in src_tc.findall("w:p", NS):
        p_out = ET.Element(f"{W}p")
        pPr_out = ET.SubElement(p_out, f"{W}pPr")
        ET.SubElement(pPr_out, f"{W}spacing", {f"{W}before": "40", f"{W}after": "40", f"{W}line": "280", f"{W}lineRule": "auto"})
        old_bidi = p_in.find("./w:pPr/w:bidi", NS)
        if old_bidi is not None:
            ET.SubElement(pPr_out, f"{W}bidi")
        _walk_inline(p_in, p_out)
        tc_out.append(p_out)

    if not tc_out.findall("w:p", NS):
        ET.SubElement(tc_out, f"{W}p")
    return tc_out


def _make_report_p(
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


def _write_dedup_report_docx(
    report_path: Path,
    pkg_index: UnmergedPackageIndex,
    checked_file_count: int,
    total_checked_rows: int,
    matches: List[dict],
    date_str: str,
) -> Path:
    """Create a standalone, 100% ECMA-376 compliant Landscape A4 Word report at `report_path`."""
    ET.register_namespace("w", W_URI)
    ET.register_namespace("r", R_URI)

    doc_root = ET.Element(f"{W}document")
    body = ET.SubElement(doc_root, f"{W}body")

    body.append(
        _make_report_p(
            "未入库改点包 · 第一列中文四级查重结果明细报告",
            bold=True,
            sz_half_pt=32,
            color_hex="0F172A",
            jc="center",
            space_before=40,
            space_after=140,
        )
    )
    body.append(
        _make_report_p(
            f"检查日期：{date_str}    ｜    未入库改点包目录：{pkg_index.pkg_dir}（共 {pkg_index.file_count} 篇文档，{len(pkg_index.rows)} 个段落）",
            bold=False,
            sz_half_pt=20,
            color_hex="475569",
            space_before=20,
            space_after=40,
        )
    )
    summary_color = "B91C1C" if matches else "15803D"
    body.append(
        _make_report_p(
            f"查重汇总：共检查待处理文档 {checked_file_count} 篇（{total_checked_rows} 个有效段落），发现重复段落 {len(matches)} 处。",
            bold=True,
            sz_half_pt=23,
            color_hex=summary_color,
            space_before=40,
            space_after=160,
        )
    )

    if not matches:
        body.append(
            _make_report_p(
                "✅ 未发现任何与「未入库改点包」重复的段落。",
                bold=True,
                sz_half_pt=24,
                color_hex="15803D",
                space_before=120,
                space_after=120,
            )
        )
    else:
        rev_counter = [100]
        col_pcts = ("1100", "1950", "1950")
        col_dxas = ("3203", "5678", "5679")
        hdr_labels = ("第 1 列：中文原文", "第 2 列：英文改点", "第 3 列：外文译文")

        for idx_m, m in enumerate(matches, 1):
            sim_pct = f"{float(m.get('similarity', 1.0)) * 100:.1f}%"
            body.append(
                _make_report_p(
                    f"【重复项 #{idx_m}】 命中阶段：{m['stage']}（相似度：{sim_pct}）",
                    bold=True,
                    sz_half_pt=23,
                    color_hex="991B1B",
                    space_before=140,
                    space_after=40,
                )
            )
            body.append(
                _make_report_p(
                    f"  (a) 待处理文档：{m['input_file']}  ｜  篇目标题：{m['input_title']}  ｜  位置：第 {m['input_table_no']} 表 · 第 {m['input_row_no']} 行",
                    bold=False,
                    sz_half_pt=20,
                    color_hex="1E293B",
                    space_before=20,
                    space_after=20,
                )
            )
            body.append(
                _make_report_p(
                    f"  (b) 重复包来源：{m['source_rel_path']}  ｜  篇目标题：{m['source_title']}  ｜  位置：第 {m['source_table_no']} 表 · 第 {m['source_row_no']} 行",
                    bold=False,
                    sz_half_pt=20,
                    color_hex="0369A1",
                    space_before=20,
                    space_after=60,
                )
            )

            tbl = ET.Element(f"{W}tbl")
            tblPr = ET.SubElement(tbl, f"{W}tblPr")
            ET.SubElement(tblPr, f"{W}tblW", {f"{W}w": "5000", f"{W}type": "pct"})
            tblBdr = ET.SubElement(tblPr, f"{W}tblBorders")
            for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
                ET.SubElement(
                    tblBdr,
                    f"{W}{side}",
                    {f"{W}val": "single", f"{W}sz": "4", f"{W}space": "0", f"{W}color": "CBD5E1"},
                )
            tblGrid = ET.SubElement(tbl, f"{W}tblGrid")
            for dxa_w in col_dxas:
                ET.SubElement(tblGrid, f"{W}gridCol", {f"{W}w": dxa_w})

            # Header row
            tr_hdr = ET.SubElement(tbl, f"{W}tr")
            for c_i, h_lbl in enumerate(hdr_labels):
                tc_h = ET.SubElement(tr_hdr, f"{W}tc")
                tcPr_h = ET.SubElement(tc_h, f"{W}tcPr")
                ET.SubElement(tcPr_h, f"{W}tcW", {f"{W}w": col_pcts[c_i], f"{W}type": "pct"})
                ET.SubElement(tcPr_h, f"{W}shd", {f"{W}val": "clear", f"{W}color": "auto", f"{W}fill": "F1F5F9"})
                tc_h.append(_make_report_p(h_lbl, bold=True, sz_half_pt=20, color_hex="334155", space_before=40, space_after=40))

            # Snapshot content row
            tr_snap = ET.SubElement(tbl, f"{W}tr")
            snap_tcs: List[Optional[ET.Element]] = m.get("_snapshot_tcs") or [None, None, None]
            for c_i in range(3):
                src_tc = snap_tcs[c_i] if c_i < len(snap_tcs) else None
                clean_tc = _sanitize_snapshot_tc(src_tc, rev_counter)
                tcPr_s = clean_tc.find("w:tcPr", NS)
                tcW_s = ET.Element(f"{W}tcW", {f"{W}w": col_pcts[c_i], f"{W}type": "pct"})
                tcPr_s.insert(0, tcW_s)
                tr_snap.append(clean_tc)

            body.append(tbl)

    sectPr = ET.SubElement(body, f"{W}sectPr")
    ET.SubElement(sectPr, f"{W}pgSz", {f"{W}w": "16838", f"{W}h": "11906", f"{W}orient": "landscape"})
    ET.SubElement(
        sectPr,
        f"{W}pgMar",
        {
            f"{W}top": "1080",
            f"{W}right": "1080",
            f"{W}bottom": "1080",
            f"{W}left": "1080",
            f"{W}header": "720",
            f"{W}footer": "720",
            f"{W}gutter": "0",
        },
    )

    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        + ET.tostring(doc_root, encoding="unicode")
    ).encode("utf-8")

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


# ---------------------------------------------------------------------------
# Core File Inspection, Row/Table Pruning, and Public API
# ---------------------------------------------------------------------------

def _remove_emptied_tables_and_headings(body: ET.Element, emptied_tbl_elements: set) -> None:
    """
    Remove `<w:tbl>` elements that lost all their content rows during deduplication,
    along with their preceding heading `<w:p>`(s) and trailing blank separator `<w:p>`(s),
    without ever deleting top 'Blue/Yellow highlighted parts...' instruction paragraphs.
    """
    if not emptied_tbl_elements:
        return
    children = list(body)
    n = len(children)
    to_remove = set()

    def _p_text(el: ET.Element) -> str:
        return "".join(t.text or "" for t in el.findall(".//w:t", NS)).strip()

    def _is_blank_p(el: ET.Element) -> bool:
        return el.tag.split("}")[-1] == "p" and _p_text(el) == ""

    def _is_heading_p(el: ET.Element) -> bool:
        if el.tag.split("}")[-1] != "p":
            return False
        txt = _p_text(el)
        return bool(txt) and not is_top_highlight_instruction_text(txt)

    for idx, ch in enumerate(children):
        if id(ch) not in emptied_tbl_elements:
            continue
        to_remove.add(idx)
        b_idx = idx - 1
        if b_idx >= 1 and _is_blank_p(children[b_idx]) and _is_heading_p(children[b_idx - 1]):
            to_remove.add(b_idx)
            b_idx -= 1
        while b_idx >= 0 and _is_heading_p(children[b_idx]):
            to_remove.add(b_idx)
            b_idx -= 1
        f_idx = idx + 1
        while f_idx < n and _is_blank_p(children[f_idx]):
            to_remove.add(f_idx)
            f_idx += 1

    for idx in sorted(to_remove, reverse=True):
        body.remove(children[idx])

    non_sect = [c for c in list(body) if c.tag.split("}")[-1] != "sectPr"]
    if not non_sect:
        body.insert(0, ET.Element(f"{W}p"))


def _check_single_input_docx(
    inp_path: Path,
    pkg_index: UnmergedPackageIndex,
    prune_duplicates_in_tree: bool,
    balance_columns: bool,
) -> Tuple[ET.Element, list, Optional[ET.Element], int, List[dict]]:
    """
    Check `inp_path` against `pkg_index`, optionally pruning duplicate rows in `root`.
    Returns `(root, ns_pairs, comments_root, checked_rows_count, file_matches)`.
    """
    root, _, ns_pairs = docx_xml_utils.load_docx_xml(inp_path)
    comments_root = docx_xml_utils.load_docx_comments_xml(inp_path)
    body = root.find("w:body", NS)
    if body is None:
        return root, ns_pairs, comments_root, 0, []

    inp_normcase = os.path.normcase(str(inp_path.resolve()))
    cur_title = ""
    tbl_no = 0
    checked_rows = 0
    file_matches: List[dict] = []
    emptied_tbl_ids = set()

    for ch in list(body):
        loc = ch.tag.split("}")[-1]
        if loc == "p":
            ptxt = "".join(t.text or "" for t in ch.findall(".//w:t", NS)).strip()
            if ptxt and not is_top_highlight_instruction_text(ptxt) and len(ptxt) <= 200:
                cur_title = ptxt
        elif loc == "tbl":
            tbl_no += 1
            trs = ch.findall("w:tr", NS)
            hdr_cnt = _header_row_count_in_table(trs, cur_title)
            eff_title = cur_title
            if not eff_title and hdr_cnt >= 1:
                tcs0 = trs[0].findall("w:tc", NS)
                if tcs0:
                    eff_title = docx_xml_utils.get_cell_plain_text(tcs0[0]).strip()
            if not eff_title:
                eff_title = inp_path.stem

            dup_trs: List[ET.Element] = []
            had_content_rows = False

            for r_idx, tr in enumerate(trs):
                if r_idx < hdr_cnt:
                    continue
                tcs = tr.findall("w:tc", NS)
                if not tcs:
                    continue
                if len(tcs) >= 3 and docx_xml_utils.get_cell_plain_text(tcs[2]).strip() == docx_xml_utils.SKIPPED_ROW_MARKER:
                    continue
                cn_raw = docx_xml_utils.get_cell_plain_text(tcs[0]).strip()
                cn_loose = normalize_cn_loose(cn_raw)
                if len(cn_loose) < 6:
                    continue

                had_content_rows = True
                checked_rows += 1
                en_raw = docx_xml_utils.get_cell_plain_text(tcs[1]).strip() if len(tcs) >= 2 else ""
                col3_raw = docx_xml_utils.get_cell_plain_text(tcs[2]).strip() if len(tcs) >= 3 else ""

                hit = match_row_4stage(
                    tc1=tcs[0],
                    tc2=tcs[1] if len(tcs) >= 2 else None,
                    inp_cn=cn_raw,
                    inp_en=en_raw,
                    pkg_index=pkg_index,
                    exclude_normcase=inp_normcase,
                )
                if hit is not None:
                    src_row, sim, stage_lbl = hit
                    snap_tcs = [
                        copy.deepcopy(tcs[0]),
                        copy.deepcopy(tcs[1]) if len(tcs) >= 2 else None,
                        copy.deepcopy(tcs[2]) if len(tcs) >= 3 else None,
                    ]
                    file_matches.append({
                        "input_file": inp_path.name,
                        "input_path": str(inp_path),
                        "input_title": eff_title,
                        "input_table_no": tbl_no,
                        "input_row_no": r_idx + 1,
                        "source_file": src_row["file_name"],
                        "source_rel_path": src_row["rel_path"],
                        "source_path": src_row["abs_path"],
                        "source_title": src_row["article_title"],
                        "source_table_no": src_row["table_no"],
                        "source_row_no": src_row["row_no"],
                        "stage": stage_lbl,
                        "similarity": sim,
                        "cn_text": cn_raw,
                        "en_text": en_raw,
                        "col3_text": col3_raw,
                        "_snapshot_tcs": snap_tcs,
                    })
                    dup_trs.append(tr)

            if prune_duplicates_in_tree and dup_trs:
                for d_tr in dup_trs:
                    ch.remove(d_tr)
                rem_trs = ch.findall("w:tr", NS)
                rem_content = [
                    rt for idx_rt, rt in enumerate(rem_trs)
                    if idx_rt >= min(hdr_cnt, len(rem_trs))
                    and not _is_blank_tr(rt)
                    and len(normalize_cn_loose(docx_xml_utils.get_cell_plain_text(rt.findall("w:tc", NS)[0]))) >= 6
                ]
                if had_content_rows and not rem_content:
                    emptied_tbl_ids.add(id(ch))

    if prune_duplicates_in_tree and emptied_tbl_ids:
        _remove_emptied_tables_and_headings(body, emptied_tbl_ids)

    if prune_duplicates_in_tree and balance_columns:
        balance_trilingual_table_widths(root)

    return root, ns_pairs, comments_root, checked_rows, file_matches


def check_and_export_duplicates(
    input_files: Sequence[Union[str, Path]],
    unmerged_pkg_dir: Union[str, Path],
    output_dir: Union[str, Path],
    export_deduped_files: bool = True,
    balance_columns: bool = True,
    date_str: str = "",
    log_fn: Optional[Callable[[str], None]] = None,
    progress_fn: Optional[Callable[[float, str], None]] = None,
) -> dict:
    """
    Check all `.docx` files in `input_files` against `unmerged_pkg_dir` using 4-stage Col 1 Chinese matching,
    export `查重结果明细报告_{YYYYMMDD}.docx` (Landscape A4), and optionally export deduplicated `.docx` files
    with recalculated `(UP X words, RV Y words)` (0 omitted) and updated top highlight instruction lines.
    """
    pkg_dir = Path(unmerged_pkg_dir).resolve()
    out_d = Path(output_dir).resolve()
    if not pkg_dir.exists() or not pkg_dir.is_dir():
        raise FileNotFoundError(f"找不到未入库改点包目录：{pkg_dir}")
    out_d.mkdir(parents=True, exist_ok=True)

    # Expand any directories passed in `input_files` and filter valid .docx files
    resolved_inputs: List[Path] = []
    seen_inp = set()
    for item in input_files or []:
        p = Path(item).resolve()
        if p.is_dir():
            cands = sorted(p.rglob("*.docx"))
        elif p.is_file():
            cands = [p]
        else:
            continue
        for cp in cands:
            if cp.name.startswith("~$") or cp.name.startswith(".") or cp.name.startswith(_DEDUP_REPORT_PREFIX):
                continue
            nk = os.path.normcase(str(cp))
            if nk not in seen_inp:
                seen_inp.add(nk)
                resolved_inputs.append(cp)

    if not resolved_inputs:
        raise ValueError("未提供任何有效的待查重 .docx 文档。")

    final_date = docx_xml_utils.sanitize_filename_part(
        (date_str or "").strip() or datetime.now().strftime("%Y%m%d"),
        fallback=datetime.now().strftime("%Y%m%d"),
        max_len=16,
    )

    if progress_fn:
        progress_fn(5.0, "正在扫描并构建「未入库改点包」3-gram 检索索引...")
    pkg_index = get_unmerged_package_index(pkg_dir, log_fn=log_fn)

    cfg = docx_xml_utils.load_config()
    col2_up = cfg.get("col2_up_highlight", "yellow")
    col2_rv = cfg.get("col2_rv_highlight", "cyan")
    lang_entries = load_languages()
    known_codes = {e["code"].upper() for e in lang_entries if e.get("code")}

    total_checked_rows = 0
    all_matches: List[dict] = []
    deduped_files: List[str] = []
    n_files = len(resolved_inputs)

    for idx_f, inp_p in enumerate(resolved_inputs):
        if progress_fn:
            pct = 12.0 + (idx_f / max(1, n_files)) * 75.0
            progress_fn(pct, f"正在查重 ({idx_f + 1}/{n_files})：{inp_p.name} ...")
        if log_fn:
            log_fn(f"🔍 [{idx_f + 1}/{n_files}] 正在查验：{inp_p.name}")

        root, ns_pairs, comments_root, f_rows, f_matches = _check_single_input_docx(
            inp_path=inp_p,
            pkg_index=pkg_index,
            prune_duplicates_in_tree=bool(export_deduped_files),
            balance_columns=bool(balance_columns),
        )
        total_checked_rows += f_rows
        all_matches.extend(f_matches)

        if log_fn:
            if f_matches:
                log_fn(f"  🚨 发现 {len(f_matches)} 处重复段落（共检查 {f_rows} 段）：")
                for fm in f_matches:
                    log_fn(
                        f"     - 第{fm['input_table_no']}表第{fm['input_row_no']}行 [{fm['input_title']}] "
                        f"-> 重复于《{fm['source_file']}》第{fm['source_table_no']}表第{fm['source_row_no']}行 "
                        f"({fm['stage']}, {fm['similarity'] * 100:.1f}%)"
                    )
            else:
                log_fn(f"  ✅ 无重复段落（共检查 {f_rows} 段）")

        if export_deduped_files:
            up_w, rv_w = docx_xml_utils.compute_col2_up_rv_words(root, up_color=col2_up, rv_color=col2_rv)
            update_top_highlight_instruction_lines(root, has_up=(up_w > 0), has_rv=(rv_w > 0))

            parsed = parse_doc_filename(inp_p.stem, known_codes=known_codes)
            lang_code = parsed["existing_lang"] or detect_lang_from_text(f"{inp_p.parent.name} {inp_p.stem}", lang_entries) or "XX"
            out_name = format_counted_docx_name(
                lang_code=lang_code,
                orig_name=parsed["orig_name"],
                up_words=up_w,
                rv_words=rv_w,
                date_yyyymmdd=final_date,
            )
            dst_path = docx_xml_utils.safe_join_output_file(out_d, out_name)
            # Never overwrite the user's original input file if output_dir == inp_p.parent
            if os.path.normcase(str(dst_path.resolve())) == os.path.normcase(str(inp_p.resolve())):
                safe_dedup_orig = f"{parsed['orig_name']}_去重定稿"
                out_name = format_counted_docx_name(
                    lang_code=lang_code,
                    orig_name=safe_dedup_orig,
                    up_words=up_w,
                    rv_words=rv_w,
                    date_yyyymmdd=final_date,
                )
                dst_path = docx_xml_utils.safe_join_output_file(out_d, out_name)

            docx_xml_utils.save_docx_xml(inp_p, dst_path, root, ns_pairs, comments_root=comments_root)
            deduped_files.append(str(dst_path))
            if log_fn:
                tag_disp = format_up_rv_count_tag(up_w, rv_w) or "0 words"
                log_fn(f"  💾 已导出去重定稿：{dst_path.name} [{tag_disp}]")

    if progress_fn:
        progress_fn(92.0, "正在生成《查重结果明细报告》...")

    report_name = f"查重结果明细报告_{final_date}.docx"
    report_path = docx_xml_utils.safe_join_output_file(out_d, report_name)
    _write_dedup_report_docx(
        report_path=report_path,
        pkg_index=pkg_index,
        checked_file_count=n_files,
        total_checked_rows=total_checked_rows,
        matches=all_matches,
        date_str=final_date,
    )

    # Strip internal XML element references from returned public dicts
    public_matches = [
        {k: v for k, v in m.items() if not k.startswith("_")}
        for m in all_matches
    ]

    if log_fn:
        log_fn(f"📋 查重明细报告已保存：{report_path}")
    if progress_fn:
        progress_fn(100.0, f"查重完成：共发现 {len(public_matches)} 处重复段落！")

    return {
        "report_path": str(report_path),
        "deduped_files": deduped_files,
        "total_checked_rows": total_checked_rows,
        "duplicate_count": len(public_matches),
        "matches": public_matches,
    }


def remove_unmerged_duplicates_from_docx(
    docx_path: Union[str, Path],
    unmerged_pkg_dir: Union[str, Path],
    report_dir: Union[str, Path],
    doc_tag: str = "",
    log_fn: Optional[Callable[[str], None]] = None,
    balance_columns: bool = True,
    date_str: str = "",
) -> dict:
    """
    In-place helper used by `pipeline.py` & `full_article_processor.py` during One-Click runs:
    strips rows duplicating `unmerged_pkg_dir` from `docx_path` (`发译者版`), updates top
    'Blue/Yellow highlighted parts...' lines, recalculates UP/RV words, and writes
    `查重结果明细_{doc_tag}_{date}.docx` in `report_dir` if `duplicate_count > 0`.
    """
    target_p = Path(docx_path).resolve()
    pkg_dir = Path(unmerged_pkg_dir).resolve()
    rep_dir = Path(report_dir).resolve()

    cfg = docx_xml_utils.load_config()
    col2_up = cfg.get("col2_up_highlight", "yellow")
    col2_rv = cfg.get("col2_rv_highlight", "cyan")

    if not target_p.is_file():
        raise FileNotFoundError(f"找不到待去重文档：{target_p}")
    if not pkg_dir.exists() or not pkg_dir.is_dir():
        raise FileNotFoundError(f"找不到未入库改点包目录：{pkg_dir}")

    pkg_index = get_unmerged_package_index(pkg_dir, log_fn=log_fn)
    root, ns_pairs, comments_root, checked_rows, matches = _check_single_input_docx(
        inp_path=target_p,
        pkg_index=pkg_index,
        prune_duplicates_in_tree=True,
        balance_columns=bool(balance_columns),
    )

    up_w, rv_w = docx_xml_utils.compute_col2_up_rv_words(root, up_color=col2_up, rv_color=col2_rv)
    dup_cnt = len(matches)
    report_path_str = ""

    if dup_cnt > 0:
        update_top_highlight_instruction_lines(root, has_up=(up_w > 0), has_rv=(rv_w > 0))
        docx_xml_utils.save_docx_xml(target_p, target_p, root, ns_pairs, comments_root=comments_root)

        final_date = docx_xml_utils.sanitize_filename_part(
            (date_str or "").strip() or datetime.now().strftime("%Y%m%d"),
            fallback=datetime.now().strftime("%Y%m%d"),
            max_len=16,
        )
        clean_tag = docx_xml_utils.sanitize_filename_part((doc_tag or "").strip(), fallback="", max_len=80)
        rep_fname = (
            f"查重结果明细_{clean_tag}_{final_date}.docx"
            if clean_tag
            else f"查重结果明细_{final_date}.docx"
        )
        rep_dir.mkdir(parents=True, exist_ok=True)
        rep_path = docx_xml_utils.safe_join_output_file(rep_dir, rep_fname)
        _write_dedup_report_docx(
            report_path=rep_path,
            pkg_index=pkg_index,
            checked_file_count=1,
            total_checked_rows=checked_rows,
            matches=matches,
            date_str=final_date,
        )
        report_path_str = str(rep_path)
        if log_fn:
            log_fn(f"  🧹 已从发译者版自动剔除 {dup_cnt} 处未入库重复段落，查重明细已导出：{rep_path.name}")
    else:
        if log_fn:
            log_fn(f"  ✅ 未入库改点包查重通过：0 处重复（共核验 {checked_rows} 段）")

    return {
        "duplicate_count": dup_cnt,
        "report_path": report_path_str,
        "up_words": int(up_w),
        "rv_words": int(rv_w),
    }
