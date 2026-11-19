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
from library_index import (  # noqa: E402
    normalize_cn,
    normalize_cn_loose,
    split_cn_sentences,
)
from step1_extract_and_diff import (  # noqa: E402
    best_local_cn_similarity,
    compute_en_overlap_score,
    extract_colored_segments,
)

W = docx_xml_utils.W
W_URI = docx_xml_utils.W_URI
XML_SPACE = docx_xml_utils.XML_SPACE
NS = docx_xml_utils.NS

_WS_RE = re.compile(r"[\s\u3000\u200b\u200c\ufeff]+")
_SYNC_EN_TOK_RE = re.compile(r"[A-Za-z0-9]+(?:[’'][A-Za-z0-9]+)*|\s+|.", re.DOTALL)
_SENT_END_BEFORE_INS_RE = re.compile(r"[.?!。？！][”’\"'\)\]]*\s+$")
_CLOSING_OR_CLAUSE_PUNCT = frozenset(".,;:?!，。；：？！、”’\"')）】]")
_CHAPTER_PREFIX_STRIP_RE = re.compile(r"^(?:中英[^\W\d_]{1,4}[\s_]+)?(?:\d+(?:-\d+)*(?:[\s_]+\d{2,4})?[\s_]+)?")
_TRAILING_TAG_STRIP_RE = re.compile(
    r"(?:[\s_]*(?:（[^）]{1,16}）|\([^)]{1,16}\)|【[^】]{1,24}】|\[[^\]]{1,24}\]|(?:19|20)\d{6}))+[\s_]*$"
)
_TRAILING_PUNCT_STRIP_RE = re.compile(r"[？?!！。．…；;：:]+$")

_MASTER_SYNC_LOCK = threading.Lock()
_MASTER_SYNC_CACHE: Dict[str, Tuple[tuple, "_MasterRowIndex"]] = {}


def _local_name(tag: str) -> str:
    return tag.split("}")[-1] if "}" in tag else tag


def _norm_title_key(title: str) -> str:
    s = (title or "").strip().strip("《》“”\"'")
    if not s:
        return ""
    s = _CHAPTER_PREFIX_STRIP_RE.sub("", s).strip(" _")
    s = _TRAILING_TAG_STRIP_RE.sub("", s).strip(" _")
    s = _TRAILING_PUNCT_STRIP_RE.sub("", s).strip()
    s = re.sub(r"[—─―-]+", "—", s)
    return _WS_RE.sub("", s).lower()


def _clean_cell_text_accepted(tc: ET.Element) -> str:
    """Read plain text of a cell as if all track changes are accepted (ignoring <w:del>/<w:moveFrom>)."""
    p_texts: List[str] = []

    def _walk(node: ET.Element, buf: List[str]):
        for ch in node:
            loc = _local_name(ch.tag)
            if loc in ("del", "moveFrom", "rPr", "pPr", "tcPr"):
                continue
            if loc == "t":
                if ch.text:
                    buf.append(ch.text)
            elif loc == "tab":
                buf.append("\t")
            elif loc == "noBreakHyphen":
                buf.append("-")
            else:
                _walk(ch, buf)

    for p in tc.findall("w:p", NS):
        p_buf: List[str] = []
        _walk(p, p_buf)
        p_texts.append("".join(p_buf))
    return "\n".join(p_texts)


def _accept_cell_revisions_inplace(tc: ET.Element) -> bool:
    """
    Accept all track changes inside `tc` in-place:
      - Unwraps `<w:ins>` and `<w:moveTo>` (keeping runs and their `<w:rPr>` intact);
      - Removes `<w:del>` and `<w:moveFrom>` (rescuing any nested comment range markers);
      - Removes `<w:rPrChange>`, `<w:pPrChange>`, `<w:tcPrChange>`, and `<w:moveFromRange*>`/`<w:moveToRange*>`.
    Returns True if any revision markup was modified/removed.
    """
    changed = False
    tcPr = tc.find("w:tcPr", NS)
    if tcPr is not None:
        for ch in list(tcPr):
            if _local_name(ch.tag) in ("tcPrChange", "cellIns", "cellDel", "cellMerge"):
                tcPr.remove(ch)
                changed = True

    def _rescue_comment_markers(node: ET.Element) -> List[ET.Element]:
        rescued: List[ET.Element] = []
        for elem in node.iter():
            eloc = _local_name(elem.tag)
            if eloc in ("commentRangeStart", "commentRangeEnd"):
                rescued.append(ET.Element(elem.tag, dict(elem.attrib)))
        return rescued

    def _process_container(parent: ET.Element) -> None:
        nonlocal changed
        new_children: List[ET.Element] = []
        mutated_here = False
        for ch in list(parent):
            loc = _local_name(ch.tag)
            if loc in ("ins", "moveTo"):
                mutated_here = True
                changed = True
                _process_container(ch)
                for sub in list(ch):
                    new_children.append(sub)
            elif loc in ("del", "moveFrom"):
                mutated_here = True
                changed = True
                for c_marker in _rescue_comment_markers(ch):
                    new_children.append(c_marker)
            elif loc in (
                "moveFromRangeStart",
                "moveFromRangeEnd",
                "moveToRangeStart",
                "moveToRangeEnd",
                "pPrChange",
                "rPrChange",
            ):
                mutated_here = True
                changed = True
            else:
                if loc in ("p", "pPr", "r", "rPr", "sdt", "sdtContent", "hyperlink", "smartTag", "customXml"):
                    _process_container(ch)
                new_children.append(ch)

        if mutated_here:
            for ch in list(parent):
                parent.remove(ch)
            for ch in new_children:
                parent.append(ch)

    _process_container(tc)
    return changed


def _scan_master_files(master_path: Path, exclude_paths: Sequence[Path] = ()) -> List[Tuple[str, str, int, int]]:
    excl_norms = {os.path.normcase(str(p.resolve())) for p in (exclude_paths or ()) if p is not None}
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
                if ename.startswith(".") or any(k in low for k in ("output", "temp", "初筛", "标记结果")):
                    continue
                nxt = f"{rel_prefix}/{ename}" if rel_prefix else ename
                _walk(entry.path, nxt)
            elif entry.is_file(follow_symlinks=False):
                if not ename.lower().endswith(".docx") or ename.startswith("~$") or ename.startswith("."):
                    continue
                if ename.startswith("初筛_") or ename.startswith("查重结果"):
                    continue
                abs_resolved = str(Path(entry.path).resolve())
                if os.path.normcase(abs_resolved) in excl_norms:
                    continue
                rel_posix = f"{rel_prefix}/{ename}" if rel_prefix else ename
                try:
                    st = os.stat(entry.path)
                except OSError:
                    continue
                found.append((rel_posix, abs_resolved, st.st_mtime_ns, st.st_size))

    _walk(str(master_path), "")
    found.sort(key=lambda x: x[0])
    return found


def _parse_single_master_file_rows(task: Tuple[str, str, int, int]) -> List[dict]:
    rel_posix, abs_path_str, _, _ = task
    stem = Path(abs_path_str).stem.strip()
    try:
        with zipfile.ZipFile(abs_path_str, "r") as zf:
            docx_xml_utils._validate_docx_zip(zf)
            xml_bytes = docx_xml_utils.read_docx_member_bounded(zf, "word/document.xml")
        root = docx_xml_utils.parse_safe_openxml_bytes(xml_bytes)
    except Exception:
        return []

    body = root.find("w:body", NS)
    if body is None:
        return []

    rows_out: List[dict] = []
    pending_headings: List[str] = []
    cur_section = ""

    for ch in body:
        loc = _local_name(ch.tag)
        if loc == "p":
            txt = "".join(t.text or "" for t in ch.findall(".//w:t", NS)).strip()
            if txt and len(txt) <= 240:
                pending_headings.append(txt)
                if re.match(r"^(?:第[\s一二三四五六七八九十百\d]+部分|附篇|基督在众教会)", txt):
                    cur_section = txt
        elif loc == "tbl":
            trs = ch.findall("w:tr", NS)
            tbl_titles = []
            if stem:
                tbl_titles.append(stem)
            if cur_section:
                tbl_titles.append(cur_section)
            for hp in pending_headings[-3:]:
                if hp not in tbl_titles:
                    tbl_titles.append(hp)
            if trs:
                tcs0 = trs[0].findall("w:tc", NS)
                if tcs0:
                    r0_cn = _clean_cell_text_accepted(tcs0[0]).strip()
                    if r0_cn and len(r0_cn) <= 80 and r0_cn not in tbl_titles:
                        tbl_titles.append(r0_cn)

            title_norms = set()
            for t_cand in tbl_titles:
                nk = _norm_title_key(t_cand)
                if nk:
                    title_norms.add(nk)
                for sub_part in re.split(r"[・·>›/]+", t_cand):
                    snk = _norm_title_key(sub_part)
                    if snk:
                        title_norms.add(snk)

            for r_idx, tr in enumerate(trs):
                tcs = tr.findall("w:tc", NS)
                if len(tcs) < 2:
                    continue
                cn_txt = _clean_cell_text_accepted(tcs[0])
                cn_norm = normalize_cn(cn_txt)
                if not cn_norm:
                    continue
                en_txt = _clean_cell_text_accepted(tcs[1])
                cn_loose = normalize_cn_loose(cn_txt)
                sents = split_cn_sentences(cn_txt)
                open_str = "".join(sents[:3]) if sents else cn_norm[:80]
                close_str = "".join(sents[-3:]) if sents else cn_norm[-80:]
                rows_out.append({
                    "rel_path": rel_posix,
                    "row_idx": r_idx,
                    "title_norms": frozenset(title_norms),
                    "cn_text": cn_txt,
                    "en_text": en_txt,
                    "cn_norm": cn_norm,
                    "cn_loose": cn_loose,
                    "open_loose": normalize_cn_loose(open_str),
                    "close_loose": normalize_cn_loose(close_str),
                })
            pending_headings = []

    return rows_out


class _MasterRowIndex:
    def __init__(self, rows: List[dict]):
        self.rows = rows
        self.exact_cn_map: Dict[str, List[dict]] = {}
        self.trigram_postings: Dict[str, List[int]] = {}
        for idx, r in enumerate(rows):
            self.exact_cn_map.setdefault(r["cn_norm"], []).append(r)
            s = r["cn_loose"]
            n = len(s)
            if n >= 3:
                seen_local = set()
                positions = list(range(0, n - 2, 2))
                if (n - 3) not in positions:
                    positions.append(n - 3)
                for p in positions:
                    gram = s[p : p + 3]
                    if gram not in seen_local:
                        seen_local.add(gram)
                        self.trigram_postings.setdefault(gram, []).append(idx)

    def retrieve_trigram_candidates(self, query_cn: str, top_k: int = 25) -> List[dict]:
        q_loose = normalize_cn_loose(query_cn)
        if len(q_loose) < 3:
            return []
        q_grams = {q_loose[i : i + 3] for i in range(len(q_loose) - 2)}
        counts: Dict[int, int] = {}
        for g in q_grams:
            plist = self.trigram_postings.get(g)
            if plist:
                for r_i in plist:
                    counts[r_i] = counts.get(r_i, 0) + 1
        if not counts:
            return []
        ranked = sorted(counts.items(), key=lambda kv: kv[1], reverse=True)[:top_k]
        return [self.rows[r_i] for r_i, _ in ranked]


def _load_master_row_index(
    latest_master_path: Path,
    exclude_paths: Sequence[Path] = (),
    log_fn: Optional[Callable[[str], None]] = None,
) -> _MasterRowIndex:
    t0 = time.perf_counter()
    resolved = latest_master_path.resolve()
    if not resolved.exists():
        raise FileNotFoundError(f"找不到最新中英底稿路径：{resolved}")

    norm_key = os.path.normcase(str(resolved))
    disk_stats = _scan_master_files(resolved, exclude_paths=exclude_paths)
    stat_sig = tuple((rp, mt, sz) for rp, _, mt, sz in disk_stats)

    with _MASTER_SYNC_LOCK:
        hit = _MASTER_SYNC_CACHE.get(norm_key)
        if hit is not None and hit[0] == stat_sig:
            return hit[1]

        all_rows: List[dict] = []
        if disk_stats:
            workers = max(1, min(8, len(disk_stats)))
            if workers == 1:
                parsed_lists = [_parse_single_master_file_rows(disk_stats[0])]
            else:
                with ThreadPoolExecutor(max_workers=workers) as ex:
                    parsed_lists = list(ex.map(_parse_single_master_file_rows, disk_stats))
            for rlist in parsed_lists:
                all_rows.extend(rlist)

        idx_obj = _MasterRowIndex(all_rows)
        _MASTER_SYNC_CACHE[norm_key] = (stat_sig, idx_obj)
        if log_fn:
            log_fn(
                f"📖 最新中英底稿已加载：共 {len(disk_stats)} 个文件、{len(all_rows)} 个段落"
                f"（耗时 {time.perf_counter() - t0:.2f}s）"
            )
        return idx_obj


def _match_pending_row_to_master(
    cur_title: str,
    p_cn: str,
    p_en: str,
    tc0: ET.Element,
    tc1: ET.Element,
    master_idx: _MasterRowIndex,
) -> Optional[dict]:
    """
    Match a pending row against `master_idx` via preceding title + 5-stage Chinese similarity
    (full / opening / colored / ending + English lexical overlap verification).
    """
    p_cn_norm = normalize_cn(p_cn)
    p_cn_loose = normalize_cn_loose(p_cn)
    if not p_cn_norm:
        return None

    title_keys = set()
    if cur_title:
        tk_full = _norm_title_key(cur_title)
        if tk_full:
            title_keys.add(tk_full)
        for part in re.split(r"[・·>›/]+", cur_title):
            pk = _norm_title_key(part)
            if pk:
                title_keys.add(pk)

    def _title_hit(cand: dict) -> bool:
        if not title_keys:
            return False
        c_norms = cand["title_norms"]
        for qk in title_keys:
            if qk in c_norms or any(qk in cn_k or cn_k in qk for cn_k in c_norms if len(cn_k) >= 2):
                return True
        return False

    # Fast exact Chinese match (Stages 1 & 2 exact)
    exact_cands = master_idx.exact_cn_map.get(p_cn_norm)
    if exact_cands:
        if len(exact_cands) == 1:
            return exact_cands[0]
        scored_exact = []
        for cand in exact_cands:
            t_bonus = 0.25 if _title_hit(cand) else 0.0
            en_ov = compute_en_overlap_score(p_en, cand["en_text"])
            scored_exact.append((t_bonus + en_ov, cand))
        scored_exact.sort(key=lambda x: x[0], reverse=True)
        return scored_exact[0][1]

    # Stage 1: Title-scoped fuzzy search
    if title_keys:
        scoped = [r for r in master_idx.rows if _title_hit(r)]
        if scoped:
            best_sc = 0.0
            best_cand = None
            for cand in scoped:
                if abs(len(cand["cn_loose"]) - len(p_cn_loose)) / max(1, len(p_cn_loose)) <= 0.45:
                    sim = difflib.SequenceMatcher(None, p_cn_loose, cand["cn_loose"], autojunk=False).ratio()
                else:
                    sim = best_local_cn_similarity(p_cn_loose[:80], cand["cn_loose"]) * 0.92
                en_ov = compute_en_overlap_score(p_en, cand["en_text"])
                combined = sim * 0.85 + en_ov * 0.15
                if sim >= 0.72 and combined > best_sc:
                    best_sc = combined
                    best_cand = cand
            if best_cand is not None:
                return best_cand

    # Stage 2: Full paragraph + Opening 1-3 sentences 3-gram search
    cn_sents = split_cn_sentences(p_cn)
    open_q = "".join(cn_sents[:3]) if cn_sents else p_cn_norm[:80]
    pool_s2 = master_idx.retrieve_trigram_candidates(p_cn, top_k=25)
    for c_open in master_idx.retrieve_trigram_candidates(open_q, top_k=15):
        if c_open not in pool_s2:
            pool_s2.append(c_open)

    best_s2 = None
    best_s2_score = 0.0
    open_loose = normalize_cn_loose(open_q)
    for cand in pool_s2:
        full_sim = difflib.SequenceMatcher(None, p_cn_loose, cand["cn_loose"], autojunk=False).ratio()
        open_sim = (
            best_local_cn_similarity(open_loose, cand["cn_loose"][: max(len(open_loose) + 30, int(len(open_loose) * 1.4))])
            if len(open_loose) >= 12
            else full_sim
        )
        en_ov = compute_en_overlap_score(p_en, cand["en_text"])
        eff_sim = max(full_sim, open_sim * 0.94 if en_ov >= 0.30 else 0.0)
        if eff_sim >= 0.74:
            comb = eff_sim * 0.82 + en_ov * 0.18 + (0.05 if _title_hit(cand) else 0.0)
            if comb > best_s2_score:
                best_s2_score = comb
                best_s2 = cand
    if best_s2 is not None:
        return best_s2

    # Stage 3: Colored Chinese segments (`有颜色的句子部分`)
    col_segs = extract_colored_segments(tc0)
    col_en_segs = extract_colored_segments(tc1)
    best_s3 = None
    best_s3_score = 0.0
    for cseg in col_segs:
        c_loose = normalize_cn_loose(cseg)
        if len(c_loose) < 10:
            continue
        for cand in master_idx.retrieve_trigram_candidates(cseg, top_k=20):
            sim = 1.0 if c_loose in cand["cn_loose"] else best_local_cn_similarity(c_loose, cand["cn_loose"])
            en_ov = compute_en_overlap_score(p_en, cand["en_text"], colored_en_segs=col_en_segs)
            if sim >= 0.88 and (len(c_loose) >= 18 or en_ov >= 0.25):
                comb = sim * 0.80 + en_ov * 0.20 + (0.05 if _title_hit(cand) else 0.0)
                if comb > best_s3_score:
                    best_s3_score = comb
                    best_s3 = cand
    if best_s3 is not None:
        return best_s3

    # Stage 4: Paragraph ending (`段落结尾`, closing 1-3 sentences)
    close_q = "".join(cn_sents[-3:]) if cn_sents else p_cn_norm[-80:]
    close_loose = normalize_cn_loose(close_q)
    if len(close_loose) >= 12:
        best_s4 = None
        best_s4_score = 0.0
        for cand in master_idx.retrieve_trigram_candidates(close_q, top_k=20):
            tail_len = max(len(close_loose) + 30, int(len(close_loose) * 1.4))
            tail_slice = cand["cn_loose"][-tail_len:]
            sim = 1.0 if close_loose in tail_slice else best_local_cn_similarity(close_loose, tail_slice)
            en_ov = compute_en_overlap_score(p_en, cand["en_text"], colored_en_segs=col_en_segs)
            if sim >= 0.88 and en_ov >= 0.25:
                comb = sim * 0.80 + en_ov * 0.20 + (0.05 if _title_hit(cand) else 0.0)
                if comb > best_s4_score:
                    best_s4_score = comb
                    best_s4 = cand
        if best_s4 is not None:
            return best_s4

    return None


# ---------------------------------------------------------------------------
# Surgical Format-Preserving Cell Text Updater
# ---------------------------------------------------------------------------

def _sanitize_run_rpr_exact(rPr: Optional[ET.Element]) -> Optional[ET.Element]:
    if rPr is None:
        return None
    cp = copy.deepcopy(rPr)
    docx_xml_utils.remove_rpr_child(cp, "rPrChange")
    docx_xml_utils._strip_unsafe_attrs_recursive(cp)
    return cp if len(cp) > 0 else None


def _rpr_key(rPr: Optional[ET.Element]) -> bytes:
    if rPr is None or len(rPr) == 0:
        return b""
    return ET.tostring(rPr, encoding="utf-8")


def _get_shd_fill(rPr: Optional[ET.Element]) -> Optional[str]:
    if rPr is None:
        return None
    shd = rPr.find("w:shd", NS)
    if shd is None:
        return None
    fill = (shd.attrib.get(f"{W}fill") or "").strip().upper()
    if fill in ("", "AUTO", "FFFFFF", "NONE"):
        return None
    return fill


def _get_special_font_color(rPr: Optional[ET.Element]) -> Optional[str]:
    """Return non-black/non-auto font color hex (e.g. '0000FF', 'FF0000') if present, else None."""
    if rPr is None:
        return None
    col = rPr.find("w:color", NS)
    if col is None:
        return None
    val = (col.attrib.get(f"{W}val") or "").strip().upper()
    if val in ("", "AUTO", "000000", "BLACK"):
        return None
    return val


def _extract_cell_char_styles(tc: ET.Element) -> Tuple[str, List[Tuple[str, Optional[ET.Element], int]], List[Optional[ET.Element]]]:
    """
    Extract `(full_text, char_records, orig_pPrs)` from `tc` (after revisions have been accepted).
    Each item in `char_records` is `(ch, rPr_copy_or_None, p_idx)`.
    """
    orig_pPrs: List[Optional[ET.Element]] = []
    char_records: List[Tuple[str, Optional[ET.Element], int]] = []

    def _walk_inline(node: ET.Element, p_idx: int, inherited_rpr: Optional[ET.Element]):
        for ch in node:
            loc = _local_name(ch.tag)
            if loc in ("pPr", "del", "moveFrom", "commentRangeStart", "commentRangeEnd", "commentReference"):
                continue
            if loc == "r":
                r_rpr = _sanitize_run_rpr_exact(ch.find("w:rPr", NS))
                eff_rpr = r_rpr if r_rpr is not None else inherited_rpr
                for sub in ch:
                    sloc = _local_name(sub.tag)
                    if sloc == "t":
                        if sub.text:
                            for c_ch in sub.text:
                                char_records.append((c_ch, eff_rpr, p_idx))
                    elif sloc == "tab":
                        char_records.append(("\t", eff_rpr, p_idx))
                    elif sloc == "noBreakHyphen":
                        char_records.append(("-", eff_rpr, p_idx))
            elif loc in ("sdt", "sdtContent", "hyperlink", "smartTag", "customXml", "ins", "moveTo", "fldSimple"):
                _walk_inline(ch, p_idx, inherited_rpr)

    for p_idx, p in enumerate(tc.findall("w:p", NS)):
        pPr = p.find("w:pPr", NS)
        if pPr is not None:
            cp_ppr = copy.deepcopy(pPr)
            docx_xml_utils.remove_ppr_child(cp_ppr, "pPrChange")
            docx_xml_utils._strip_unsafe_attrs_recursive(cp_ppr)
            orig_pPrs.append(cp_ppr)
        else:
            orig_pPrs.append(None)

        if p_idx > 0:
            char_records.append(("\n", None, p_idx - 1))
        _walk_inline(p, p_idx, None)

    full_text = "".join(ch for ch, _, _ in char_records)
    return full_text, char_records, orig_pPrs


def _tokenize_with_spans(text: str, is_chinese: bool) -> Tuple[List[str], List[Tuple[int, int]]]:
    if not text:
        return [], []
    if is_chinese:
        toks = list(text)
        spans = [(i, i + 1) for i in range(len(toks))]
        return toks, spans
    toks = []
    spans = []
    for m in _SYNC_EN_TOK_RE.finditer(text):
        toks.append(m.group(0))
        spans.append((m.start(), m.end()))
    return toks, spans


def _is_inside_open_curly_quotes(prefix_text: str) -> bool:
    return prefix_text.count("“") > prefix_text.count("”")


def _choose_insert_rpr_and_pidx(
    pending_text: str,
    pending_chars: List[Tuple[str, Optional[ET.Element], int]],
    c_pos: int,
    ins_text: str,
    fallback_rpr: Optional[ET.Element],
) -> Tuple[Optional[ET.Element], int]:
    """
    Choose `(rPr_copy, p_idx)` for an inserted span at `c_pos` in `pending_chars`:
      - Keeps `w:shd` background shading continuous across enclosing shaded spans (ZERO white holes);
      - Respects sentence boundaries when inserting right between a white sentence and a shaded sentence;
      - Strips edge-only special quote `<w:color>` (e.g. `0000FF`) when inserting outside quotes,
        while ALWAYS keeping `w:shd`, font family, and font size intact.
    """
    n = len(pending_chars)
    if n == 0:
        return (copy.deepcopy(fallback_rpr) if fallback_rpr is not None else None), 0

    left_idx = c_pos - 1 if c_pos > 0 else None
    while left_idx is not None and left_idx >= 0 and pending_chars[left_idx][0] == "\n":
        left_idx = left_idx - 1 if left_idx > 0 else None

    right_idx = c_pos if c_pos < n else None
    while right_idx is not None and right_idx < n and pending_chars[right_idx][0] == "\n":
        right_idx = right_idx + 1 if right_idx + 1 < n else None

    if left_idx is None and right_idx is None:
        return (copy.deepcopy(fallback_rpr) if fallback_rpr is not None else None), 0
    if left_idx is None:
        chosen_idx = right_idx
    elif right_idx is None:
        chosen_idx = left_idx
    else:
        left_rpr = pending_chars[left_idx][1]
        right_rpr = pending_chars[right_idx][1]
        left_shd = _get_shd_fill(left_rpr)
        right_shd = _get_shd_fill(right_rpr)
        if left_shd == right_shd:
            chosen_idx = left_idx
        else:
            ins_stripped = ins_text.lstrip()
            starts_with_punct = bool(ins_stripped) and ins_stripped[0] in _CLOSING_OR_CLAUSE_PUNCT
            after_sent_end_ws = bool(_SENT_END_BEFORE_INS_RE.search(pending_text[:c_pos]))
            if after_sent_end_ws and not starts_with_punct:
                chosen_idx = right_idx
            else:
                chosen_idx = left_idx

    donor_rpr = copy.deepcopy(pending_chars[chosen_idx][1]) if pending_chars[chosen_idx][1] is not None else None
    donor_pidx = pending_chars[chosen_idx][2]

    # Guard against bleeding inline quote <w:color> (e.g. 0000FF) onto insertions outside the quote
    if donor_rpr is not None:
        spec_col = _get_special_font_color(donor_rpr)
        if spec_col is not None:
            left_col = _get_special_font_color(pending_chars[left_idx][1]) if left_idx is not None else None
            right_col = _get_special_font_color(pending_chars[right_idx][1]) if right_idx is not None else None
            strictly_inside_colored = (left_col == spec_col and right_col == spec_col)
            inside_quotes = _is_inside_open_curly_quotes(pending_text[:c_pos])
            if not strictly_inside_colored and not inside_quotes:
                other_idx = right_idx if chosen_idx == left_idx else left_idx
                other_rpr = pending_chars[other_idx][1] if other_idx is not None else None
                other_col_el = other_rpr.find("w:color", NS) if other_rpr is not None else None
                if other_col_el is not None and _get_special_font_color(other_rpr) is None:
                    docx_xml_utils.insert_rpr_child_ordered(donor_rpr, copy.deepcopy(other_col_el))
                else:
                    docx_xml_utils.remove_rpr_child(donor_rpr, "color")

    return donor_rpr, donor_pidx


def _make_exact_text_run(text: str, rPr_template: Optional[ET.Element]) -> ET.Element:
    r_el = ET.Element(f"{W}r")
    if rPr_template is not None and len(rPr_template) > 0:
        r_el.append(copy.deepcopy(rPr_template))
    if "\t" not in text:
        t_el = ET.Element(f"{W}t")
        if text.startswith(" ") or text.endswith(" ") or "  " in text:
            t_el.set(XML_SPACE, "preserve")
        t_el.text = text
        r_el.append(t_el)
    else:
        parts = text.split("\t")
        for idx_p, part in enumerate(parts):
            if idx_p > 0:
                r_el.append(ET.Element(f"{W}tab"))
            if part:
                t_el = ET.Element(f"{W}t")
                if part.startswith(" ") or part.endswith(" ") or "  " in part:
                    t_el.set(XML_SPACE, "preserve")
                t_el.text = part
                r_el.append(t_el)
    return r_el


def _surgical_update_cell_text(
    tc: ET.Element,
    latest_text: str,
    is_chinese: bool,
    working_comments_root: Optional[ET.Element],
    id_allocator: docx_xml_utils.MarkupIdAllocator,
) -> bool:
    """
    Update `tc` text to `latest_text` while 100% preserving `tc`'s `<w:pPr>`, per-character `<w:rPr>`
    (`DEEAF6` / `F2CEED` / `None` shading, `0000FF` / `FF0000` colors, fonts, sizes, bold), and Word comments.
    Returns True if text was changed.
    """
    pending_text, pending_chars, orig_pPrs = _extract_cell_char_styles(tc)
    if pending_text == latest_text:
        return False

    comment_anchors = (
        docx_xml_utils.extract_cell_comment_anchors(tc, working_comments_root)
        if working_comments_root is not None
        else []
    )

    fallback_rpr: Optional[ET.Element] = None
    for _, rpr_cand, _ in pending_chars:
        if rpr_cand is not None:
            fallback_rpr = copy.deepcopy(rpr_cand)
            break
    if fallback_rpr is None:
        ppr_rpr = tc.find("./w:p/w:pPr/w:rPr", NS)
        if ppr_rpr is not None:
            fallback_rpr = _sanitize_run_rpr_exact(ppr_rpr)

    p_toks, p_spans = _tokenize_with_spans(pending_text, is_chinese=is_chinese)
    m_toks, _ = _tokenize_with_spans(latest_text, is_chinese=is_chinese)

    sm = difflib.SequenceMatcher(None, p_toks, m_toks, autojunk=False)
    out_chars: List[Tuple[str, Optional[ET.Element], int, bool]] = []
    # 4th tuple field: is_from_edit (True for replace/insert, False for exact equal)

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            if i1 < i2:
                c_s = p_spans[i1][0]
                c_e = p_spans[i2 - 1][1]
                for ch, rpr_el, pidx in pending_chars[c_s:c_e]:
                    out_chars.append((ch, rpr_el, pidx, False))
        elif tag == "delete":
            continue
        elif tag == "replace":
            c_s = p_spans[i1][0]
            c_e = p_spans[i2 - 1][1]
            src_slice = pending_chars[c_s:c_e]
            non_nl_src = [(ch, rpr_el, pidx) for ch, rpr_el, pidx in src_slice if ch != "\n"]
            if not non_nl_src:
                donor_rpr, donor_pidx = _choose_insert_rpr_and_pidx(
                    pending_text, pending_chars, c_s, "".join(m_toks[j1:j2]), fallback_rpr
                )
                for m_tok in m_toks[j1:j2]:
                    for ch in m_tok:
                        out_chars.append((ch, donor_rpr, donor_pidx, True))
            else:
                first_key = _rpr_key(non_nl_src[0][1])
                uniform = all(_rpr_key(rpr_el) == first_key for _, rpr_el, _ in non_nl_src[1:])
                if uniform:
                    u_rpr, u_pidx = non_nl_src[0][1], non_nl_src[0][2]
                    for m_tok in m_toks[j1:j2]:
                        for ch in m_tok:
                            out_chars.append((ch, u_rpr, u_pidx, True))
                else:
                    new_toks = m_toks[j1:j2]
                    tot_new_chars = max(1, sum(len(t) for t in new_toks))
                    cur_off = 0
                    src_len = len(non_nl_src)
                    for m_tok in new_toks:
                        mid = (cur_off + len(m_tok) * 0.5) / tot_new_chars
                        d_idx = min(src_len - 1, max(0, int(mid * src_len)))
                        _, d_rpr, d_pidx = non_nl_src[d_idx]
                        for ch in m_tok:
                            out_chars.append((ch, d_rpr, d_pidx, True))
                        cur_off += len(m_tok)
        elif tag == "insert":
            c_pos = p_spans[i1][0] if i1 < len(p_spans) else len(pending_chars)
            ins_str = "".join(m_toks[j1:j2])
            donor_rpr, donor_pidx = _choose_insert_rpr_and_pidx(
                pending_text, pending_chars, c_pos, ins_str, fallback_rpr
            )
            for ch in ins_str:
                out_chars.append((ch, donor_rpr, donor_pidx, True))

    # Zero white-hole continuity pass: ensure no edit-induced whitespace/punctuation inside a shaded block
    # accidentally drops `w:shd` when both neighbors share the exact same non-None `w:shd` fill!
    n_out = len(out_chars)
    idx_o = 0
    while idx_o < n_out:
        ch_o, rpr_o, pidx_o, from_edit = out_chars[idx_o]
        if from_edit and ch_o != "\n" and _get_shd_fill(rpr_o) is None:
            j_o = idx_o + 1
            while (
                j_o < n_out
                and out_chars[j_o][3]
                and out_chars[j_o][0] != "\n"
                and _get_shd_fill(out_chars[j_o][1]) is None
            ):
                j_o += 1
            if idx_o > 0 and j_o < n_out:
                l_shd = _get_shd_fill(out_chars[idx_o - 1][1])
                r_shd = _get_shd_fill(out_chars[j_o][1])
                if l_shd is not None and l_shd == r_shd:
                    donor_shd_el = out_chars[idx_o - 1][1].find("w:shd", NS)
                    if donor_shd_el is not None:
                        for k_o in range(idx_o, j_o):
                            patched_rpr = (
                                copy.deepcopy(out_chars[k_o][1])
                                if out_chars[k_o][1] is not None
                                else (copy.deepcopy(out_chars[idx_o - 1][1]) or ET.Element(f"{W}rPr"))
                            )
                            docx_xml_utils.insert_rpr_child_ordered(patched_rpr, copy.deepcopy(donor_shd_el))
                            out_chars[k_o] = (out_chars[k_o][0], patched_rpr, out_chars[k_o][2], True)
            idx_o = j_o
        else:
            idx_o += 1

    # Partition `out_chars` into paragraphs on '\n' and merge adjacent runs with identical serialized <w:rPr>
    para_groups: List[List[Tuple[str, Optional[ET.Element]]]] = [[]]
    for ch_o, rpr_o, _, _ in out_chars:
        if ch_o == "\n":
            para_groups.append([])
        else:
            para_groups[-1].append((ch_o, rpr_o))

    tcPr = tc.find("w:tcPr", NS)
    for ch in list(tc):
        if ch is not tcPr:
            tc.remove(ch)

    for p_i, p_chars in enumerate(para_groups):
        p_el = ET.Element(f"{W}p")
        if orig_pPrs:
            donor_pPr = orig_pPrs[p_i] if p_i < len(orig_pPrs) else orig_pPrs[-1]
            if donor_pPr is not None:
                p_el.append(copy.deepcopy(donor_pPr))

        merged_runs: List[Tuple[List[str], Optional[ET.Element], bytes]] = []
        for ch_c, rpr_c in p_chars:
            rk = _rpr_key(rpr_c)
            if merged_runs and merged_runs[-1][2] == rk:
                merged_runs[-1][0].append(ch_c)
            else:
                merged_runs.append(([ch_c], rpr_c, rk))

        for ch_list, rpr_c, _ in merged_runs:
            run_txt = "".join(ch_list)
            if run_txt:
                p_el.append(_make_exact_text_run(run_txt, rpr_c))

        tc.append(p_el)

    if comment_anchors and working_comments_root is not None:
        docx_xml_utils.reanchor_comments_to_cell(tc, comment_anchors, working_comments_root, id_allocator)

    return True


def sync_pending_with_latest_master(
    pending_docx: Union[str, Path],
    latest_master_path: Union[str, Path],
    output_dir: Optional[Union[str, Path]] = None,
    date_str: str = "",
    log_fn: Optional[Callable[[str], None]] = None,
) -> dict:
    """
    Synchronize `pending_docx` Col 1 & Col 2 text with `latest_master_path` (either a single `.docx`
    file or a directory of single/multi-article `.docx` files), while 100% preserving `pending_docx`'s
    original fonts, sizes, inline font colors (`0000FF`, `FF0000`), background shadings (`DEEAF6`,
    `F2CEED`, `None` white), `<w:pPr>`, and Word comments.

    Saves `{pending_stem}{YYYYMMDD}更新.docx` into `output_dir` (defaults to `pending_docx.parent`)
    and returns:
      {"updated_docx_path": str(out_path), "matched_rows": int, "updated_rows": int, "unmatched_rows": int}
    """
    pend_p = Path(pending_docx).resolve()
    mast_p = Path(latest_master_path).resolve()
    if not pend_p.is_file():
        raise FileNotFoundError(f"找不到待处理文档：{pend_p}")
    if not mast_p.exists():
        raise FileNotFoundError(f"找不到最新中英底稿文件或目录：{mast_p}")

    out_d = Path(output_dir).resolve() if output_dir is not None and str(output_dir).strip() else pend_p.parent.resolve()
    out_d.mkdir(parents=True, exist_ok=True)

    eff_date = docx_xml_utils.sanitize_filename_part(
        (date_str or "").strip() or datetime.now().strftime("%Y%m%d"),
        fallback=datetime.now().strftime("%Y%m%d"),
        max_len=16,
    )
    safe_stem = docx_xml_utils.sanitize_filename_part(pend_p.stem, fallback="待处理稿子", max_len=140)
    out_name = f"{safe_stem}{eff_date}更新.docx"
    out_path = docx_xml_utils.safe_join_output_file(out_d, out_name)

    if log_fn:
        log_fn(f"🔄 正在核对最新中英底稿：{pend_p.name} <-> {mast_p.name} ...")

    master_idx = _load_master_row_index(mast_p, exclude_paths=(pend_p, out_path), log_fn=log_fn)

    root, _, ns_pairs = docx_xml_utils.load_docx_xml(pend_p)
    orig_comments = docx_xml_utils.load_docx_comments_xml(pend_p)
    had_comments = orig_comments is not None
    working_comments = orig_comments if orig_comments is not None else docx_xml_utils.new_comments_root()
    id_allocator = docx_xml_utils.MarkupIdAllocator(doc_root=root, comments_root=working_comments)

    body = root.find("w:body", NS)
    if body is None:
        raise ValueError(f"待处理文档缺少 w:body：{pend_p}")

    matched_rows = 0
    updated_rows = 0
    unmatched_rows = 0
    cur_title = ""

    for ch in body:
        loc = _local_name(ch.tag)
        if loc == "p":
            ptxt = "".join(t.text or "" for t in ch.findall(".//w:t", NS)).strip()
            if ptxt:
                cur_title = ptxt
        elif loc == "tbl":
            for r_i, tr in enumerate(ch.findall("w:tr", NS)):
                tcs = tr.findall("w:tc", NS)
                if len(tcs) < 2:
                    continue

                rev_c0 = _accept_cell_revisions_inplace(tcs[0])
                rev_c1 = _accept_cell_revisions_inplace(tcs[1])

                p_cn = docx_xml_utils.get_cell_plain_text(tcs[0])
                if not normalize_cn(p_cn):
                    continue
                p_en = docx_xml_utils.get_cell_plain_text(tcs[1])

                m_row = _match_pending_row_to_master(cur_title, p_cn, p_en, tcs[0], tcs[1], master_idx)
                if m_row is None:
                    unmatched_rows += 1
                    if log_fn:
                        log_fn(f"  ⚠️ 未在最新底稿匹配到段落：[{cur_title or '无标题'}] 第 {r_i + 1} 行（保留原内容）")
                    continue

                matched_rows += 1
                upd_c0 = _surgical_update_cell_text(
                    tcs[0], m_row["cn_text"], is_chinese=True, working_comments_root=working_comments, id_allocator=id_allocator
                )
                upd_c1 = _surgical_update_cell_text(
                    tcs[1], m_row["en_text"], is_chinese=False, working_comments_root=working_comments, id_allocator=id_allocator
                )
                if rev_c0 or rev_c1 or upd_c0 or upd_c1:
                    updated_rows += 1

    final_comments = (
        working_comments
        if (had_comments or len(working_comments.findall(f"{W}comment")) > 0)
        else None
    )
    docx_xml_utils.save_docx_xml(pend_p, out_path, root, ns_pairs, comments_root=final_comments)

    if log_fn:
        log_fn(
            f"✅ 最新底稿核对更新完成：匹配 {matched_rows} 段，更新内容 {updated_rows} 段，"
            f"未匹配 {unmatched_rows} 段 -> 已保存：{out_path.name}"
        )

    return {
        "updated_docx_path": str(out_path),
        "matched_rows": matched_rows,
        "updated_rows": updated_rows,
        "unmatched_rows": unmatched_rows,
    }
