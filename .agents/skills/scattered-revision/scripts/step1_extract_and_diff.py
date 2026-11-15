import argparse
import difflib
import json
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from docx_xml_utils import (
    NS,
    W,
    diff_tokens,
    extract_char_shading,
    get_cell_plain_text,
    load_config,
    load_docx_xml,
    segment_diff_into_sentences,
    sentence_unit_stats,
    split_into_bg_regions,
)
from library_index import (
    LibraryIndex,
    get_library_index,
    is_excluded_library_file,
    normalize_cn,
    normalize_cn_loose,
    split_cn_sentences,
)

CN_NUM_CHARS = "一二三四五六七八九十百零"

_EN_WORD_RE = re.compile(r"[A-Za-z0-9]+(?:[’'][A-Za-z0-9]+)*")
_EN_STOPWORDS = frozenset({
    "the", "and", "that", "this", "with", "from", "have", "they", "their",
    "your", "what", "when", "will", "into", "only", "even", "more", "than",
    "some", "such", "were", "been", "which", "there", "would", "could",
    "should", "about", "for", "are", "was", "not", "but", "you", "all",
    "can", "had", "her", "his", "one", "our", "out", "who", "has", "him",
    "how", "man", "men", "god", "people", "things", "these", "those", "them",
    "then", "also", "must", "does", "did", "upon", "unto", "very", "just",
})


def detect_lang_code(library_dir: Path, matched_files: list, cfg: dict) -> str:
    lang_map = cfg.get("language_code_map", {})
    search_str = str(library_dir) + " " + " ".join(p.name for p in matched_files)
    for kw, code in lang_map.items():
        if kw in search_str:
            return code
    return "XX"


def index_library_files(library_dir: Path, exclude_keywords: list) -> list:
    files = []
    for p in library_dir.rglob("*.docx"):
        rel_p = p.relative_to(library_dir)
        if is_excluded_library_file(rel_p, exclude_keywords):
            continue
        files.append(p)
    return files


def _leaf_matches_stem(leaf: str, stem: str) -> bool:
    norm_leaf = normalize_cn(leaf)
    norm_stem = normalize_cn(stem)
    if not norm_leaf or not norm_stem:
        return False
    pos = norm_stem.find(norm_leaf)
    while pos != -1:
        end_pos = pos + len(norm_leaf)
        # Prevent '十' from matching prefix of '十一', '二十' from '二十一', etc.
        if norm_leaf and norm_leaf[-1] in CN_NUM_CHARS and end_pos < len(norm_stem) and norm_stem[end_pos] in CN_NUM_CHARS:
            pos = norm_stem.find(norm_leaf, pos + 1)
            continue
        return True
    return False


def match_title_to_index_rels(title: str, lib_index: LibraryIndex) -> List[str]:
    """
    Match `title` against indexed library files using both file stems AND header paragraphs inside each .docx.
    Returns ordered list of candidate `rel_path` strings (empty list if unmatched).
    """
    clean_title = (title or "").strip().strip("《》")
    if not clean_title:
        return []

    parts = [p.strip().strip("《》") for p in re.split(r"[・·]", clean_title) if p.strip()]
    leaf = parts[-1] if parts else clean_title

    sub_leaf_m = re.match(r"^附篇[一二三四五六七八九十百零]+[\s\u3000]+(.+)$", leaf)
    leaf_variants = [leaf]
    if sub_leaf_m:
        leaf_variants.append(sub_leaf_m.group(1).strip())
    if clean_title not in leaf_variants:
        leaf_variants.append(clean_title)

    all_rels = list(lib_index.files_map.keys())
    pool = all_rels
    if len(parts) >= 2:
        for qual in parts[:-1]:
            q_norm = normalize_cn(qual)
            synonyms = [q_norm]
            if "基督起初的发表" in q_norm:
                synonyms.append("起初")
            elif "神向全宇的说话" in q_norm:
                synonyms.append("全宇")
            elif "对神奥秘的揭示" in q_norm:
                synonyms.append("揭示")
            filtered = [
                rel_p for rel_p in pool
                if any(syn in normalize_cn(rel_p) for syn in synonyms)
            ]
            if filtered:
                pool = filtered

    def _search_pool(cand_pool: List[str]) -> List[str]:
        matched = []
        for lv in leaf_variants:
            for rel_p in cand_pool:
                f_info = lib_index.files_map[rel_p]
                titles = f_info.get("titles") or [Path(rel_p).stem]
                if any(_leaf_matches_stem(lv, t_cand) for t_cand in titles):
                    if rel_p not in matched:
                        matched.append(rel_p)
            if matched:
                break
        return matched

    hits = _search_pool(pool)
    if not hits and pool is not all_rels:
        hits = _search_pool(all_rels)

    hits.sort(key=lambda rp: len(Path(rp).stem))
    return hits


def match_library_file(title: str, lib_files: list, first_row_cn: str = "", src_xml_cache: dict = None):
    """Legacy-compatible file matcher."""
    if not title or not title.strip():
        return None
    clean_title = title.strip().strip("《》")
    parts = [p.strip().strip("《》") for p in re.split(r"[・·]", clean_title) if p.strip()]
    leaf = parts[-1] if parts else clean_title

    sub_leaf_m = re.match(r"^附篇[一二三四五六七八九十百零]+[\s\u3000]+(.+)$", leaf)
    leaf_variants = [leaf]
    if sub_leaf_m:
        leaf_variants.append(sub_leaf_m.group(1).strip())

    pool = lib_files
    if len(parts) >= 2:
        for qual in parts[:-1]:
            q_norm = normalize_cn(qual)
            synonyms = [q_norm]
            if "基督起初的发表" in q_norm:
                synonyms.append("起初")
            elif "神向全宇的说话" in q_norm:
                synonyms.append("全宇")
            elif "对神奥秘的揭示" in q_norm:
                synonyms.append("揭示")
            filtered = [p for p in pool if any(syn in normalize_cn(str(p)) for syn in synonyms)]
            if filtered:
                pool = filtered

    hits = []
    for lv in leaf_variants:
        hits = [p for p in pool if _leaf_matches_stem(lv, p.stem)]
        if hits:
            break
    if not hits and pool is not lib_files:
        for lv in leaf_variants:
            hits = [p for p in lib_files if _leaf_matches_stem(lv, p.stem)]
            if hits:
                break

    if len(hits) == 1:
        return hits[0]
    if len(hits) > 1:
        hits.sort(key=lambda p: len(p.stem))
        if first_row_cn and src_xml_cache is not None:
            best_f = None
            best_sc = -1.0
            for cand in hits:
                if cand not in src_xml_cache:
                    s_root, _, _ = load_docx_xml(cand)
                    src_xml_cache[cand] = s_root
                _, s_tcs, sc = find_matching_source_row(first_row_cn, src_xml_cache[cand])
                if s_tcs is not None and sc > best_sc:
                    best_sc = sc
                    best_f = cand
                if best_sc == 1.0:
                    break
            if best_f is not None:
                return best_f
        return hits[0]
    return None


def find_matching_source_row(inp_cn: str, src_root):
    """Legacy XML row matcher retained for backward compatibility."""
    inp_norm = normalize_cn(inp_cn)
    best_score = 0.0
    best_row = None
    best_idx = -1

    all_rows = src_root.findall(".//w:tbl/w:tr", NS)
    for idx, tr in enumerate(all_rows):
        tcs = tr.findall("w:tc", NS)
        if len(tcs) < 3:
            continue
        src_cn = get_cell_plain_text(tcs[0])
        src_norm = normalize_cn(src_cn)
        if not src_norm:
            continue
        if src_norm == inp_norm:
            return idx, tcs, 1.0
        if inp_norm in src_norm or src_norm in inp_norm:
            score = min(len(inp_norm), len(src_norm)) / max(len(inp_norm), len(src_norm))
            if score > best_score:
                best_score = score
                best_row = tcs
                best_idx = idx
        else:
            if abs(len(src_norm) - len(inp_norm)) / max(1, len(inp_norm)) < 0.35:
                score = difflib.SequenceMatcher(None, inp_norm, src_norm).ratio()
                if score > best_score:
                    best_score = score
                    best_row = tcs
                    best_idx = idx

    if best_score >= 0.75:
        return best_idx, best_row, round(best_score, 4)
    return -1, None, 0.0


def _has_visual_color_on_props(prop_el: Optional[ET.Element]) -> bool:
    """Return True if <w:rPr>, <w:pPr>, or <w:tcPr> has non-white shading, highlight, or non-black font color."""
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


def extract_colored_segments(tc: ET.Element) -> List[str]:
    """
    Extract visually colored text segments (`有颜色的句子部分`) from a table cell `tc`:
    runs, paragraphs, or cells with non-white `w:shd`, `w:highlight`, or non-black/non-auto `w:color`.
    """
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
            # Ignore deleted track-change runs if any exist in input
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
                # Bridge tiny 1-2 char uncolored punctuation/spaces between colored runs in the same sentence
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
                    # If the colored run is a short phrase inside a Chinese clause, also include its enclosing sentence
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

    # Deduplicate preserving order, preferring longer informative segments first
    dedup: List[str] = []
    seen = set()
    for s in sorted(segments, key=lambda x: len(normalize_cn_loose(x)), reverse=True):
        key = normalize_cn(s)
        if key and key not in seen:
            seen.add(key)
            dedup.append(s)
    return dedup


def compute_en_overlap_score(inp_en: str, cand_en: str, colored_en_segs: Optional[List[str]] = None) -> float:
    """
    Stage 5 verification signal: compute English lexical + 3-gram phrase overlap between input Col 2 and library Col 2.
    Returns score in [0.0, 1.0] (returns 0.5 neutral if either side has no English words).
    """
    inp_words = [m.group(0).lower() for m in _EN_WORD_RE.finditer(inp_en or "") if len(m.group(0)) >= 2]
    lib_words = [m.group(0).lower() for m in _EN_WORD_RE.finditer(cand_en or "") if len(m.group(0)) >= 2]
    if not inp_words or not lib_words:
        return 0.5

    inp_set = set(inp_words)
    lib_set = set(lib_words)
    word_containment = len(inp_set & lib_set) / max(1, min(len(inp_set), len(lib_set)))

    inp_cw = {w for w in inp_set if len(w) >= 4 and w not in _EN_STOPWORDS}
    lib_cw = {w for w in lib_set if len(w) >= 4 and w not in _EN_STOPWORDS}
    cw_containment = (
        len(inp_cw & lib_cw) / max(1, min(len(inp_cw), len(lib_cw)))
        if (inp_cw and lib_cw)
        else word_containment
    )

    inp_3g = {tuple(inp_words[i : i + 3]) for i in range(max(0, len(inp_words) - 2))}
    lib_3g = {tuple(lib_words[i : i + 3]) for i in range(max(0, len(lib_words) - 2))}
    trigram_hits = len(inp_3g & lib_3g)
    trigram_score = min(1.0, trigram_hits / max(1.0, min(4.0, len(inp_3g) * 0.35))) if inp_3g and lib_3g else 0.0

    col_bonus = 0.0
    if colored_en_segs:
        for cseg in colored_en_segs[:3]:
            cw = {
                m.group(0).lower()
                for m in _EN_WORD_RE.finditer(cseg)
                if len(m.group(0)) >= 4 and m.group(0).lower() not in _EN_STOPWORDS
            }
            if len(cw) >= 2:
                ov = len(cw & lib_cw) / len(cw)
                if ov > col_bonus:
                    col_bonus = ov

    return min(1.0, max(word_containment * 0.75, cw_containment, trigram_score * 0.9, col_bonus * 0.85))


def best_local_cn_similarity(needle: str, haystack: str) -> float:
    """
    Compute best local substring similarity of `needle` inside `haystack` (on loose CJK/alnum text),
    returning [0.0, 1.0]. Fast O(N) substring / 3-gram window locating followed by bounded SequenceMatcher.
    """
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

    # Locate top candidate windows in `h` via 3-gram density
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


def _iter_scope_rows(lib_index: LibraryIndex, allowed_rels: Optional[Set[str]]) -> List[dict]:
    if allowed_rels is None:
        return lib_index.rows
    out = []
    for rp in allowed_rels:
        out.extend(lib_index.rows_by_rel.get(rp, ()))
    return out


def _search_stages_2_to_5(
    inp_cn: str,
    inp_en: str,
    colored_cn_segs: List[str],
    colored_en_segs: List[str],
    lib_index: LibraryIndex,
    allowed_rels: Optional[Set[str]] = None,
) -> Tuple[Optional[dict], float, str]:
    """
    Execute Stages 2 -> 3 -> 4 + Stage 5 Candidate Verification within `allowed_rels` (or full library when None).
    Returns `(matched_row_dict_or_None, confidence, stage_label)`.
    """
    inp_norm = normalize_cn(inp_cn)
    inp_loose = normalize_cn_loose(inp_cn)
    if not inp_norm or not inp_loose:
        return None, 0.0, ""

    scope_rows = _iter_scope_rows(lib_index, allowed_rels)
    if not scope_rows:
        return None, 0.0, ""

    def _en_sc(cand_row: dict) -> float:
        return compute_en_overlap_score(inp_en, cand_row.get("en_plain", ""), colored_en_segs)

    # ------------------------------------------------------------------
    # STAGE 2A: Exact or Full-Paragraph Col 1 Chinese Match + Stage 5
    # ------------------------------------------------------------------
    exact_hits = (
        [r for r in lib_index.exact_cn_map.get(inp_norm, ()) if (allowed_rels is None or r["rel_path"] in allowed_rels)]
    )
    if exact_hits:
        if len(exact_hits) == 1 and (len(inp_loose) >= 6 or _en_sc(exact_hits[0]) >= 0.20):
            return exact_hits[0], 1.0, "stage2_full_exact"
        scored_ex = sorted(exact_hits, key=lambda r: _en_sc(r), reverse=True)
        best_r = scored_ex[0]
        if len(inp_loose) >= 6 or _en_sc(best_r) >= 0.20:
            return best_r, 1.0, "stage2_full_exact"

    full_cands = lib_index.retrieve_trigram_candidates(inp_cn, top_k=25, allowed_rels=allowed_rels)
    best_full = None
    best_full_rank = 0.0
    best_full_sim = 0.0
    for r in full_cands:
        r_norm = r["cn_norm"]
        if inp_norm in r_norm or r_norm in inp_norm:
            c_sim = min(len(inp_norm), len(r_norm)) / max(1, max(len(inp_norm), len(r_norm)))
        elif abs(len(r_norm) - len(inp_norm)) / max(1, len(inp_norm)) <= 0.42:
            c_sim = difflib.SequenceMatcher(None, inp_norm, r_norm, autojunk=False).ratio()
        else:
            continue
        if c_sim < 0.74:
            continue
        esc = _en_sc(r)
        if c_sim >= 0.86 or (c_sim >= 0.74 and esc >= 0.12):
            rank = 0.70 * c_sim + 0.30 * esc
            if rank > best_full_rank:
                best_full_rank = rank
                best_full_sim = c_sim
                best_full = r
    if best_full is not None:
        return best_full, round(best_full_sim, 4), "stage2_full_fuzzy"

    # ------------------------------------------------------------------
    # STAGE 2B: Opening 1-3 Sentences (`前几句话` Prefix / Fuzzy Anchor) + Stage 5
    # ------------------------------------------------------------------
    cn_sents = split_cn_sentences(inp_cn)
    open_queries: List[str] = []
    for k_cnt in (3, 2, 1):
        if len(cn_sents) >= k_cnt:
            q_s = "".join(cn_sents[:k_cnt])
            if len(normalize_cn_loose(q_s)) >= 10 and q_s not in open_queries:
                open_queries.append(q_s)
    for prefix_len in (60, 32, 18):
        if len(inp_norm) >= prefix_len:
            q_p = inp_norm[:prefix_len]
            if len(normalize_cn_loose(q_p)) >= 10 and q_p not in open_queries:
                open_queries.append(q_p)

    best_open = None
    best_open_rank = 0.0
    best_open_sim = 0.0
    for oq in open_queries:
        oq_loose = normalize_cn_loose(oq)
        oq_len = len(oq_loose)
        # Direct prefix/substring hits + trigram candidates
        cands_map: Dict[Tuple[str, int], dict] = {}
        for r in scope_rows:
            if oq in r["cn_norm"][: max(120, len(oq) * 2)]:
                cands_map[(r["rel_path"], r["row_idx"])] = r
        for r in lib_index.retrieve_trigram_candidates(oq, top_k=20, allowed_rels=allowed_rels):
            cands_map[(r["rel_path"], r["row_idx"])] = r

        for r in cands_map.values():
            head_target = "".join(r.get("opening_sentences") or [])
            head_slice = r["cn_norm"][: max(len(head_target), int(len(oq) * 1.6) + 24)]
            osim = best_local_cn_similarity(oq, head_slice)
            if osim < 0.82:
                continue
            esc = _en_sc(r)
            if esc >= 0.12 or (oq_len >= 22 and osim >= 0.92):
                rank = 0.62 * osim + 0.38 * esc + min(0.05, oq_len / 1000.0)
                if rank > best_open_rank:
                    best_open_rank = rank
                    best_open_sim = osim
                    best_open = r
        if best_open is not None and best_open_sim >= 0.92 and best_open_rank >= 0.72:
            break
    if best_open is not None:
        return best_open, round(best_open_sim, 4), "stage2_opening"

    # ------------------------------------------------------------------
    # STAGE 3: Colored Sentence / Run Parts (`有颜色的句子部分`) + Stage 5
    # ------------------------------------------------------------------
    best_col = None
    best_col_rank = 0.0
    best_col_sim = 0.0

    for cseg in colored_cn_segs:
        c_norm = normalize_cn(cseg)
        c_loose = normalize_cn_loose(cseg)
        clen = len(c_loose)
        if clen < 6:
            continue
        cands_map = {}
        for r in scope_rows:
            if c_norm in r["cn_norm"]:
                cands_map[(r["rel_path"], r["row_idx"])] = r
        for r in lib_index.retrieve_trigram_candidates(cseg, top_k=25, allowed_rels=allowed_rels):
            cands_map[(r["rel_path"], r["row_idx"])] = r

        for r in cands_map.values():
            csim = 1.0 if c_norm in r["cn_norm"] else best_local_cn_similarity(cseg, r["cn_norm"])
            if csim < 0.78:
                continue
            esc = _en_sc(r)
            accepted = (
                (clen >= 8 and (esc >= 0.10 or (clen >= 16 and csim >= 0.90)))
                or (clen >= 6 and csim >= 0.95 and esc >= 0.25)
            )
            if accepted:
                rank = 0.60 * csim + 0.40 * esc + min(0.06, clen / 800.0)
                if rank > best_col_rank:
                    best_col_rank = rank
                    best_col_sim = csim
                    best_col = r
        if best_col is not None and best_col_sim >= 0.95 and best_col_rank >= 0.72:
            break

    # Secondary Stage 3: If Col 2 had colored English segments while Col 1 had uncolored middle sentences
    if best_col is None and len(cn_sents) >= 3:
        mid_sents = [s for s in cn_sents[1:-1] if len(normalize_cn_loose(s)) >= 12]
        for mseg in sorted(mid_sents, key=len, reverse=True)[:3]:
            mlen = len(normalize_cn_loose(mseg))
            cands_map = {}
            for r in scope_rows:
                if mseg in r["cn_norm"]:
                    cands_map[(r["rel_path"], r["row_idx"])] = r
            for r in lib_index.retrieve_trigram_candidates(mseg, top_k=20, allowed_rels=allowed_rels):
                cands_map[(r["rel_path"], r["row_idx"])] = r
            for r in cands_map.values():
                msim = 1.0 if mseg in r["cn_norm"] else best_local_cn_similarity(mseg, r["cn_norm"])
                if msim < 0.85:
                    continue
                esc = _en_sc(r)
                if esc >= 0.16 or (mlen >= 22 and msim >= 0.95 and esc >= 0.10):
                    rank = 0.58 * msim + 0.42 * esc
                    if rank > best_col_rank:
                        best_col_rank = rank
                        best_col_sim = msim
                        best_col = r

    if best_col is not None:
        return best_col, round(best_col_sim, 4), "stage3_colored"

    # ------------------------------------------------------------------
    # STAGE 4: Paragraph Ending (`段落结尾`: Last 1-3 Sentences / Trailing 25-80 Chars) + Stage 5
    # ------------------------------------------------------------------
    close_queries: List[str] = []
    for k_cnt in (3, 2, 1):
        if len(cn_sents) >= k_cnt:
            q_c = "".join(cn_sents[-k_cnt:])
            if len(normalize_cn_loose(q_c)) >= 10 and q_c not in close_queries:
                close_queries.append(q_c)
    for tail_len in (80, 45, 25):
        if len(inp_norm) >= tail_len:
            q_t = inp_norm[-tail_len:]
            if len(normalize_cn_loose(q_t)) >= 10 and q_t not in close_queries:
                close_queries.append(q_t)

    best_close = None
    best_close_rank = 0.0
    best_close_sim = 0.0
    for cq in close_queries:
        cq_loose = normalize_cn_loose(cq)
        cq_len = len(cq_loose)
        cands_map = {}
        for r in scope_rows:
            if r["cn_norm"].endswith(cq) or cq in r["cn_norm"][-max(140, len(cq) * 2) :]:
                cands_map[(r["rel_path"], r["row_idx"])] = r
        for r in lib_index.retrieve_trigram_candidates(cq, top_k=25, allowed_rels=allowed_rels):
            cands_map[(r["rel_path"], r["row_idx"])] = r

        for r in cands_map.values():
            tail_target = "".join(r.get("closing_sentences") or [])
            tail_slice = r["cn_norm"][-max(len(tail_target) + 20, int(len(cq) * 1.6) + 30) :]
            tsim = 1.0 if cq in tail_slice else best_local_cn_similarity(cq, tail_slice)
            if tsim < 0.80:
                continue
            esc = _en_sc(r)
            if esc >= 0.10 or (cq_len >= 20 and tsim >= 0.90):
                rank = 0.62 * tsim + 0.38 * esc + min(0.05, cq_len / 1000.0)
                if rank > best_close_rank:
                    best_close_rank = rank
                    best_close_sim = tsim
                    best_close = r
        if best_close is not None and best_close_sim >= 0.92 and best_close_rank >= 0.72:
            break

    if best_close is not None:
        return best_close, round(best_close_sim, 4), "stage4_ending"

    return None, 0.0, ""


def search_library_row_5stage(
    cur_title: str,
    tc1: ET.Element,
    tc2: ET.Element,
    inp_cn: str,
    inp_en: str,
    lib_index: LibraryIndex,
) -> Tuple[Optional[dict], float, str]:
    """
    5-Stage Cascading Search for a single input row:
      Stage 1: Search inside `cur_title`-matched library file(s) first; automatically fall back to full library!
      Stage 2: Full Col 1 Chinese paragraph + opening 1-3 sentences (`前几句话`).
      Stage 3: Colored sentence/run parts (`有颜色的句子部分`) from Col 1 (and Col 2).
      Stage 4: Paragraph ending (`段落结尾`, last 1-3 Chinese sentences / trailing ~25-80 chars).
      Stage 5: Candidate verification combining Chinese segment similarity with Col 2 English token overlap.
    Never raises on missing/non-standard title or unmatched row.
    """
    colored_cn_segs = extract_colored_segments(tc1)
    colored_en_segs = extract_colored_segments(tc2)

    title_rels = match_title_to_index_rels(cur_title, lib_index) if cur_title else []
    if title_rels:
        matched_row, conf, stage = _search_stages_2_to_5(
            inp_cn=inp_cn,
            inp_en=inp_en,
            colored_cn_segs=colored_cn_segs,
            colored_en_segs=colored_en_segs,
            lib_index=lib_index,
            allowed_rels=set(title_rels),
        )
        if matched_row is not None:
            return matched_row, conf, f"title_scoped:{stage}"

    # Fallback: full-library index search across all files
    matched_row, conf, stage = _search_stages_2_to_5(
        inp_cn=inp_cn,
        inp_en=inp_en,
        colored_cn_segs=colored_cn_segs,
        colored_en_segs=colored_en_segs,
        lib_index=lib_index,
        allowed_rels=None,
    )
    if matched_row is not None:
        return matched_row, conf, f"full_library:{stage}"

    return None, 0.0, ""


def main():
    parser = argparse.ArgumentParser(description="Step 1a: Extract library rows and compute sentence diffs.")
    parser.add_argument("--input", required=True, help="Path to 英文底稿.docx")
    parser.add_argument("--library", required=True, help="Path to target language total library directory")
    parser.add_argument("--workdir", required=True, help="Directory to store intermediate JSON manifest")
    parser.add_argument("--lang", default="", help="Optional explicit language code (e.g. FR, PT)")
    args = parser.parse_args()

    cfg = load_config()
    inp_path = Path(args.input).resolve()
    lib_dir = Path(args.library).resolve()
    workdir = Path(args.workdir).resolve()
    workdir.mkdir(parents=True, exist_ok=True)

    lib_index = get_library_index(lib_dir, log_fn=print)
    inp_root, _, _ = load_docx_xml(inp_path)
    body = inp_root.find("w:body", NS)

    # Check if ANY shading exists across all tables in input Col 2
    any_doc_shading = False
    for tbl in body.findall("w:tbl", NS):
        for tr in tbl.findall("w:tr", NS):
            tcs = tr.findall("w:tc", NS)
            if len(tcs) >= 2:
                _, fills = extract_char_shading(tcs[1])
                if any(f is not None for f in fills):
                    any_doc_shading = True
                    break

    cur_title = ""
    tbl_counter = 0
    rows_manifest = []
    skipped_rows = []
    matched_lib_paths: List[Path] = []

    rv_threshold = cfg.get("rv_word_change_ratio_threshold", 0.42)

    for body_idx, ch in enumerate(body):
        local = ch.tag.split("}")[-1]
        if local == "p":
            txt = "".join(t.text or "" for t in ch.findall(".//w:t", NS)).strip()
            if txt:
                cur_title = txt
        elif local == "tbl":
            for r_idx, tr in enumerate(ch.findall("w:tr", NS)):
                tcs = tr.findall("w:tc", NS)
                if len(tcs) < 3:
                    continue
                inp_cn = get_cell_plain_text(tcs[0])
                inp_en, inp_fills = extract_char_shading(tcs[1])
                if not normalize_cn(inp_cn) and not inp_en.strip():
                    continue
                if not any_doc_shading:
                    inp_fills = ["NONE_FULL"] * len(inp_en)

                matched_row, match_conf, match_stage = search_library_row_5stage(
                    cur_title=cur_title,
                    tc1=tcs[0],
                    tc2=tcs[1],
                    inp_cn=inp_cn,
                    inp_en=inp_en,
                    lib_index=lib_index,
                )

                if matched_row is None:
                    cn_prev = normalize_cn(inp_cn)[:80]
                    title_disp = cur_title if cur_title else "<无标题>"
                    print(
                        f"⚠️ 未在总库匹配到原文，已跳过该行 [表格 {tbl_counter + 1} 第 {r_idx + 1} 行 | 标题: {title_disp}]: {cn_prev}"
                    )
                    skipped_rows.append({
                        "body_idx": body_idx,
                        "row_idx": r_idx,
                        "article_title": cur_title,
                        "cn_preview": cn_prev,
                    })
                    continue

                src_file = lib_index.resolve_file_path(matched_row["rel_path"])
                if src_file not in matched_lib_paths:
                    matched_lib_paths.append(src_file)

                s_row_idx = int(matched_row["row_idx"])
                src_en = matched_row.get("en_plain", "")
                col3_preview = matched_row.get("col3_preview", "")

                regions = split_into_bg_regions(src_en, inp_en, inp_fills)
                region_entries = []
                for reg_idx, (fill_hex, src_seg, inp_seg) in enumerate(regions):
                    if fill_hex is None:
                        region_entries.append({
                            "region_idx": reg_idx,
                            "is_quoted": False,
                            "fill_hex": None,
                            "src_text": src_seg,
                        })
                    else:
                        actual_fill = None if fill_hex == "NONE_FULL" else fill_hex
                        ops = diff_tokens(src_seg, inp_seg)
                        sent_op_lists = segment_diff_into_sentences(ops)
                        sent_items = []
                        for s_i, s_ops in enumerate(sent_op_lists):
                            stats = sentence_unit_stats(s_ops)
                            sent_id = f"T{tbl_counter}_R{r_idx}_G{reg_idx}_S{s_i}"
                            if not stats["has_edit"]:
                                suggested = "none"
                            elif stats["change_ratio"] >= rv_threshold or (stats["change_ratio"] >= 0.32 and stats["edit_hunks"] >= 4):
                                suggested = "cyan"
                            else:
                                suggested = "yellow"
                            sent_items.append({
                                "sent_id": sent_id,
                                "suggested_action": suggested,
                                **stats,
                                "ops": s_ops,
                            })
                        region_entries.append({
                            "region_idx": reg_idx,
                            "is_quoted": True,
                            "fill_hex": actual_fill,
                            "sentences": sent_items,
                        })

                rows_manifest.append({
                    "tbl_idx": tbl_counter,
                    "body_idx": body_idx,
                    "row_idx": r_idx,
                    "article_title": cur_title,
                    "source_file": str(src_file),
                    "source_row_idx": s_row_idx,
                    "match_confidence": match_conf,
                    "match_stage": match_stage,
                    "col3_preview": col3_preview,
                    "regions": region_entries,
                })
            tbl_counter += 1

    all_idx_files = [lib_index.resolve_file_path(rp) for rp in list(lib_index.files_map.keys())[:20]]
    lang_code = args.lang.strip().upper() or detect_lang_code(lib_dir, matched_lib_paths or all_idx_files, cfg)
    manifest = {
        "input_docx": str(inp_path),
        "library_dir": str(lib_dir),
        "lang_code": lang_code,
        "any_doc_shading": any_doc_shading,
        "rows": rows_manifest,
        "skipped_rows": skipped_rows,
    }
    manifest_path = workdir / "diff_manifest.json"
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    # Also write a lightweight summary for Agent inspection (including sentence sequence context)
    summary_items = []
    for rm in rows_manifest:
        for reg in rm["regions"]:
            if not reg["is_quoted"]:
                continue
            for s in reg["sentences"]:
                if s["has_edit"]:
                    summary_items.append({
                        "sent_id": s["sent_id"],
                        "article": rm["article_title"],
                        "row_idx": rm["row_idx"],
                        "has_edit": True,
                        "suggested": s["suggested_action"],
                        "change_ratio": s["change_ratio"],
                        "edit_hunks": s["edit_hunks"],
                        "new_words": s["new_words"],
                        "diff_preview": s["diff_preview"],
                    })
                else:
                    summary_items.append({
                        "sent_id": s["sent_id"],
                        "article": rm["article_title"],
                        "row_idx": rm["row_idx"],
                        "has_edit": False,
                        "suggested": "none",
                        "new_words": s["new_words"],
                        "text_preview": s["new_text"][:80],
                    })
    summary_path = workdir / "step1_review_items.json"
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary_items, f, ensure_ascii=False, indent=2)

    print(f"Manifest written to: {manifest_path}")
    print(
        f"Review items ({len(summary_items)} sentences across {len(rows_manifest)} matched rows, "
        f"{len(skipped_rows)} skipped rows) written to: {summary_path}"
    )
    print(f"Detected language code: {lang_code}")


if __name__ == "__main__":
    main()
