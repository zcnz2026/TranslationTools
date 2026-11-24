"""
Scenario 3A: 改点入库（发稿 + 回稿 + 总库 + 可选通读定稿表）

Strict safety & fidelity guarantees:
  1. ZERO mutation to the 3-language library directory (`library_dir` is strictly read-only).
  2. Outputs clean `改前/` (verbatim copies), `改后/` (exact original filenames & dates, surgical
     sentence-level write-back preserving styles/blue scripture/red text/existing shadings),
     `修订稿/` (track changes shown on `改后`, with single / batch-N / per-article options),
     and Step 1A/1B verification reports + annotated diff `.docx`.
  3. Preserves Col 1 Chinese shading color from `回稿` and overlays ONLY onto the corresponding
     span in `改后/` without clearing any pre-existing shadings elsewhere in the library cell.
"""

from __future__ import annotations

import copy
from datetime import datetime
import difflib
import os
from pathlib import Path
import re
import shutil
import sys
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
import library_index  # noqa: E402
import step1_extract_and_diff  # noqa: E402
from ai_client import ApiKeyPool, call_gemini_json  # noqa: E402
from pipeline import ensure_library_index  # noqa: E402
from sheets_reader import (  # noqa: E402
    ColumnMappingConfig,
    GoogleOAuthManager,
    ParsedSheetData,
    ProofreadRow,
    clean_cell_spaces,
    is_google_sheet_url,
    load_proofread_table,
    parse_title_link_table,
)

W = docx_xml_utils.W
W_URI = docx_xml_utils.W_URI
R_URI = docx_xml_utils.R_URI
NS = docx_xml_utils.NS
XML_SPACE = docx_xml_utils.XML_SPACE

RV_HIGHLIGHT_VALS = {"cyan", "blue", "darkCyan"}
STRIP_COL23_HIGHLIGHTS = {"yellow", "cyan", "blue", "darkCyan", "lightGray", "darkGray"}
STRIP_COL23_SHD_FILLS = {"D9EEF7", "DAE9F7", "D9E2F3", "DEEAF6", "BDD7EE", "FFFF00"}
_NO_CHANGE_COMMENT_RE = re.compile(
    r"(不需要改|无需修改|不用改|不用修改|保持原样|原文已准确|无须修改|no\s*change|no\s*need\s*to\s*change|keep\s*original|already\s*accurate)",
    re.IGNORECASE,
)
_SENT_SPLIT_RE = re.compile(r"(?<=[.!?。？！；;])\s+")
_PUNCT_SPACE_RE = re.compile(r"[ \t]+([,;:!?，。；：！？]|\.(?!\.))")
_DOUBLE_SPACE_RE = re.compile(r"[ \t]{2,}")
_TOKEN_RE = re.compile(r"\S+|\s+")


# ---------------------------------------------------------------------------
# Low-level XML helpers (reject revisions, unwrap SDT, comments, surgical diff)
# ---------------------------------------------------------------------------

def unwrap_sdt_in_element(el: ET.Element) -> None:
    """Recursively unwrap `<w:sdt>` content controls, keeping their `<w:sdtContent>` children."""
    idx = 0
    while idx < len(el):
        ch = el[idx]
        loc = ch.tag.split("}")[-1]
        if loc == "sdt":
            sdt_content = ch.find("w:sdtContent", NS)
            el.remove(ch)
            if sdt_content is not None:
                children = list(sdt_content)
                for offset, sub in enumerate(children):
                    el.insert(idx + offset, sub)
                continue
        else:
            unwrap_sdt_in_element(ch)
            idx += 1


def reject_all_revisions_in_cell(tc: ET.Element) -> None:
    """
    Reject all Word track changes in `tc`:
      - Remove `<w:ins>` elements completely.
      - Unwrap `<w:del>` elements and convert `<w:delText>` back to `<w:t>`.
    """
    unwrap_sdt_in_element(tc)
    for p in tc.findall(".//w:p", NS):
        new_children: List[ET.Element] = []
        for ch in list(p):
            loc = ch.tag.split("}")[-1]
            if loc == "ins":
                continue
            elif loc == "del":
                for sub in list(ch):
                    sub_loc = sub.tag.split("}")[-1]
                    if sub_loc == "r":
                        for dt in sub.findall("w:delText", NS):
                            dt.tag = f"{W}t"
                            if dt.text and (dt.text[0].isspace() or dt.text[-1].isspace()):
                                dt.set(XML_SPACE, "preserve")
                        new_children.append(sub)
                    elif sub_loc in ("commentRangeStart", "commentRangeEnd"):
                        new_children.append(sub)
            else:
                new_children.append(ch)
        for ch in list(p):
            p.remove(ch)
        for ch in new_children:
            p.append(ch)


def accept_all_revisions_in_cell_safe(tc: ET.Element) -> None:
    """Unwrap `<w:sdt>` and accept all track changes in `tc`."""
    unwrap_sdt_in_element(tc)
    docx_xml_utils.accept_all_revisions_in_cell(tc)


def strip_comments_in_cell(tc: ET.Element) -> None:
    """Remove all `<w:commentRangeStart>`, `<w:commentRangeEnd>`, and `<w:commentReference>` runs from `tc`."""
    for p in tc.findall(".//w:p", NS):
        for ch in list(p):
            loc = ch.tag.split("}")[-1]
            if loc in ("commentRangeStart", "commentRangeEnd"):
                p.remove(ch)
            elif loc == "r":
                for cref in list(ch.findall("w:commentReference", NS)):
                    ch.remove(cref)
                if not ch.findall("w:t", NS) and not ch.findall("w:delText", NS) and not ch.findall("w:br", NS) and not ch.findall("w:tab", NS):
                    if all(sub.tag.split("}")[-1] == "rPr" for sub in ch):
                        p.remove(ch)


def get_cell_comment_texts(tc: ET.Element, comments_map: Dict[str, ET.Element]) -> List[str]:
    """Return plain text strings of all Word comments anchored in `tc`."""
    if not comments_map:
        return []
    ids: List[str] = []
    for tag in ("commentRangeStart", "commentReference"):
        for el in tc.findall(f".//w:{tag}", NS):
            cid = el.attrib.get(f"{W}id", "")
            if cid and cid not in ids:
                ids.append(cid)
    out: List[str] = []
    for cid in ids:
        c_el = comments_map.get(cid)
        if c_el is not None:
            txt = "".join(t.text or "" for t in c_el.findall(".//w:t", NS)).strip()
            if txt:
                out.append(txt)
    return out


def extract_non_cyan_rejected_text(tc: ET.Element) -> Tuple[str, bool]:
    """
    Return `(rejected_text_excluding_cyan_runs, has_cyan)` from a sent-draft Col 2 cell.
    Cyan (`RV`) regions in `发稿` are allowed to differ from the library, so Step 1A compares
    only the non-cyan portion when checking whether the library changed since dispatch.
    """
    work = copy.deepcopy(tc)
    has_cyan = False
    for hl in work.findall(".//w:rPr/w:highlight", NS):
        if hl.attrib.get(f"{W}val") in RV_HIGHLIGHT_VALS:
            has_cyan = True
            break
    reject_all_revisions_in_cell(work)
    if not has_cyan:
        return docx_xml_utils.get_cell_plain_text(work), False

    para_parts: List[str] = []
    for p in work.findall("w:p", NS):
        buf: List[str] = []
        for r in p.findall(".//w:r", NS):
            hl = r.find("./w:rPr/w:highlight", NS)
            if hl is not None and hl.attrib.get(f"{W}val") in RV_HIGHLIGHT_VALS:
                continue
            buf.append("".join(t.text or "" for t in r.findall("w:t", NS)))
        para_parts.append("".join(buf))
    return "\n".join(para_parts), True


def cell_has_sent_edits(tc2: ET.Element) -> bool:
    """Return True if sent-draft Col 2 cell contains track changes or UP/RV highlights."""
    if tc2.find(".//w:ins", NS) is not None or tc2.find(".//w:del", NS) is not None:
        return True
    for hl in tc2.findall(".//w:rPr/w:highlight", NS):
        if hl.attrib.get(f"{W}val") in ({"yellow"} | RV_HIGHLIGHT_VALS):
            return True
    return False


def cell_has_cyan_rv(tc2: ET.Element) -> bool:
    for hl in tc2.findall(".//w:rPr/w:highlight", NS):
        if hl.attrib.get(f"{W}val") in RV_HIGHLIGHT_VALS:
            return True
    return False


def _norm_compare_text(s: str) -> str:
    s = (s or "").replace("\u00A0", " ").replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"[ \t]+", " ", s)
    lines = [ln.strip() for ln in s.split("\n") if ln.strip()]
    return "\n".join(lines)


def _norm_compare_flat_ws(s: str) -> str:
    """
    Normalize text for consistency comparison by treating paragraph breaks (`\\n`) and multiple spaces
    identically to single spaces (and stripping space before punctuation), so isolated RV paragraphs split
    by blank lines in `发稿` Col 3 are recognized as 100% consistent with single-paragraph library text.
    """
    s = (s or "").replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    s = _PUNCT_SPACE_RE.sub(r"\1", s)
    return s


def _align_sent_text_breaks_to_library(sent_text: str, lib_text: str) -> str:
    """
    If `sent_text` has identical wording to `lib_text` modulo artificial paragraph breaks (`\\n\\n`),
    return `lib_text` so Step 1B / Step 3 diffs never produce fake newline revisions or boundary collisions.
    """
    if not sent_text or not lib_text:
        return sent_text
    if _norm_compare_flat_ws(sent_text) == _norm_compare_flat_ws(lib_text):
        return lib_text
    return sent_text


def _extract_rejected_cyan_segments(tc2: ET.Element) -> Tuple[str, bool, List[Tuple[str, bool]]]:
    """
    Reject all Word revisions in `tc2` first, then extract contiguous `(text, is_cyan_rv)` segments
    (bridging pure whitespace between two adjacent cyan runs so each cyan sentence forms a single segment).
    """
    work = copy.deepcopy(tc2)
    reject_all_revisions_in_cell(work)
    full_rej = docx_xml_utils.get_cell_plain_text(work)

    raw_chunks: List[Tuple[str, bool]] = []
    paras = work.findall("w:p", NS)
    for p_i, p in enumerate(paras):
        for r in p.findall(".//w:r", NS):
            txt = "".join(t.text or "" for t in r.findall("w:t", NS))
            if not txt:
                continue
            hl = r.find("./w:rPr/w:highlight", NS)
            is_rv = bool(hl is not None and hl.attrib.get(f"{W}val") in RV_HIGHLIGHT_VALS)
            if raw_chunks and raw_chunks[-1][1] == is_rv:
                raw_chunks[-1] = (raw_chunks[-1][0] + txt, is_rv)
            else:
                raw_chunks.append((txt, is_rv))
        if p_i < len(paras) - 1:
            if raw_chunks and not raw_chunks[-1][1]:
                raw_chunks[-1] = (raw_chunks[-1][0] + "\n", False)
            else:
                raw_chunks.append(("\n", False))

    if not any(ic for _, ic in raw_chunks):
        return full_rej, False, [(full_rej, False)] if full_rej else []

    # Bridge whitespace-only chunks sandwiched between two cyan chunks
    bridged = list(raw_chunks)
    for idx in range(1, len(bridged) - 1):
        txt_c, ic_c = bridged[idx]
        if not ic_c and not txt_c.strip() and bridged[idx - 1][1] and bridged[idx + 1][1]:
            bridged[idx] = (txt_c, True)

    merged: List[Tuple[str, bool]] = []
    for txt_c, ic_c in bridged:
        if merged and merged[-1][1] == ic_c:
            merged[-1] = (merged[-1][0] + txt_c, ic_c)
        else:
            merged.append((txt_c, ic_c))

    return full_rej, True, merged


def analyze_sent_col2_vs_library(sent_tc2: ET.Element, lib_c2_text: str) -> dict:
    """
    Step 1A + Step 1B English Analyzer:
      1. Rejects all `<w:ins>` / `<w:del>` revisions on `sent_tc2`.
      2. Compares rejected `sent_tc2` against `lib_c2_text` (`三语总库` current Col 2):
         - Cyan (`RV`) spans in `发稿` naturally differ from `lib_c2_text` (generating revisions on cyan);
         - Automatically ACCEPTS those Cyan (`RV`) revisions;
         - Checks whether ANY remaining non-cyan revisions exist (`has_remaining_diff`).
      3. Simultaneously reconstructs:
         - `step1a_target`: text with Cyan (`RV`) revisions accepted, so `render_diff_track_changes_into_cell(s_c2_rej, step1a_target)`
           renders ONLY the remaining non-cyan library mismatches!
         - `true_old_baseline`: pre-dispatch English text where non-cyan parts come from `sent_tc2` (rejected)
           and Cyan (`RV`) parts come from `lib_c2_text` (the library's original English before RV rewrite),
           so Step 1B (`true_old_baseline -> ret_c2_clean`) shows full revisions on Cyan (`RV`) sentences AND
           Step 3 `merge_keep_library_text` never reverts Cyan (`RV`) back to old library English!
    """
    s_rej, has_cyan, segs = _extract_rejected_cyan_segments(sent_tc2)
    norm_s_flat = _norm_compare_flat_ws(s_rej)
    norm_l_flat = _norm_compare_flat_ws(lib_c2_text)

    if not lib_c2_text.strip():
        return {
            "sent_rej": s_rej,
            "has_cyan": has_cyan,
            "is_excerpt": False,
            "has_remaining_diff": False,
            "step1a_target": s_rej,
            "true_old_baseline": s_rej,
        }

    if not has_cyan:
        is_exc = bool(norm_s_flat and norm_s_flat != norm_l_flat and norm_s_flat in norm_l_flat)
        eq_flat = (norm_s_flat == norm_l_flat)
        has_diff = (not eq_flat) and (not is_exc)
        return {
            "sent_rej": s_rej,
            "has_cyan": False,
            "is_excerpt": is_exc,
            "has_remaining_diff": has_diff,
            "step1a_target": lib_c2_text if has_diff else s_rej,
            "true_old_baseline": _align_sent_text_breaks_to_library(s_rej, lib_c2_text) if eq_flat else s_rej,
        }

    # All-cyan cell (100% of sent_tc2 is highlighted cyan, zero non-cyan anchors)
    nc_indices = [i for i, (tx, ic) in enumerate(segs) if not ic and tx.strip()]
    if not nc_indices:
        is_exc = len(lib_c2_text.strip()) > len(s_rej.strip()) * 2.5 + 120
        return {
            "sent_rej": s_rej,
            "has_cyan": True,
            "is_excerpt": is_exc,
            "has_remaining_diff": False,  # Accepting cyan leaves 0 remaining revisions
            "step1a_target": s_rej,
            "true_old_baseline": s_rej if is_exc else lib_c2_text,
        }

    # Anchor each non-empty non-cyan segment in `lib_c2_text` in monotonic order
    anchors: Dict[int, Tuple[int, int]] = {}
    search_cur = 0
    all_nc_exact = True
    for idx in nc_indices:
        seg_txt = segs[idx][0]
        core = seg_txt.strip()
        loc = docx_xml_utils.find_fuzzy_substring_span(lib_c2_text, core, search_from=search_cur)
        if loc is not None:
            anchors[idx] = loc
            search_cur = loc[1]
        else:
            all_nc_exact = False

    if all_nc_exact and len(anchors) == len(nc_indices):
        first_nc = nc_indices[0]
        last_nc = nc_indices[-1]
        first_starts_at_zero = (first_nc == 0)
        last_ends_at_end = (last_nc == len(segs) - 1)
        leading_unquoted = lib_c2_text[: anchors[first_nc][0]].strip() if first_starts_at_zero else ""
        trailing_unquoted = lib_c2_text[anchors[last_nc][1] :].strip() if last_ends_at_end else ""
        is_exc = bool(leading_unquoted or trailing_unquoted)

        # Build `true_old_baseline` by replacing each cyan span with the exact library gap between surrounding non-cyan anchors!
        baseline_parts: List[str] = []
        for i, (tx, ic) in enumerate(segs):
            if not ic:
                baseline_parts.append(tx)
            else:
                # Find nearest preceding & following anchored non-cyan segments
                prev_nc = max((k for k in anchors if k < i), default=None)
                next_nc = min((k for k in anchors if k > i), default=None)
                gap_s = anchors[prev_nc][1] if prev_nc is not None else 0
                gap_e = anchors[next_nc][0] if next_nc is not None else len(lib_c2_text)
                lib_gap = lib_c2_text[gap_s:gap_e]
                # Preserve leading/trailing spaces consistently with `tx`
                lead_ws = tx[: len(tx) - len(tx.lstrip())]
                trail_ws = tx[len(tx.rstrip()) :]
                baseline_parts.append(lead_ws + lib_gap.strip() + trail_ws if lib_gap.strip() else tx)

        return {
            "sent_rej": s_rej,
            "has_cyan": True,
            "is_excerpt": is_exc,
            "has_remaining_diff": False,
            "step1a_target": s_rej,
            "true_old_baseline": "".join(baseline_parts),
        }

    # At least one non-cyan segment differed from `lib_c2_text` -> real library modification in non-cyan region!
    # Reconstruct `step1a_target` (accepting cyan segments while keeping library's modified non-cyan segments)
    # and `true_old_baseline` via token-level alignment of non-cyan tokens.
    s_tok_pairs: List[Tuple[str, bool]] = []
    for tx, ic in segs:
        for tk_s in _TOKEN_RE.findall(tx):
            s_tok_pairs.append((tk_s, ic))
    s_toks = [t for t, _ in s_tok_pairs]
    l_toks = _TOKEN_RE.findall(lib_c2_text)
    sm = difflib.SequenceMatcher(None, s_toks, l_toks, autojunk=False)

    step1a_out: List[str] = []
    baseline_out: List[str] = []
    has_rem = False

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        s_piece = "".join(s_toks[i1:i2])
        l_piece = "".join(l_toks[j1:j2])
        if tag == "equal":
            step1a_out.append(s_piece)
            baseline_out.append(s_piece)
            continue
        if _norm_compare_flat_ws(s_piece) == _norm_compare_flat_ws(l_piece):
            step1a_out.append(s_piece)
            baseline_out.append(l_piece)
            continue
        touches_cyan = any(ic for _, ic in s_tok_pairs[i1:i2] if _.strip())
        if not touches_cyan and not s_piece.strip():
            left_cy = next((ic for tx, ic in reversed(s_tok_pairs[:i1]) if tx.strip()), False)
            right_cy = next((ic for tx, ic in s_tok_pairs[i2:] if tx.strip()), False)
            touches_cyan = bool(left_cy or right_cy)
        if touches_cyan:
            # Accept cyan revision for Step 1A; record library old English for Step 1B/3 baseline
            step1a_out.append(s_piece)
            baseline_out.append(l_piece)
        else:
            has_rem = True
            step1a_out.append(l_piece)
            baseline_out.append(s_piece)

    return {
        "sent_rej": s_rej,
        "has_cyan": True,
        "is_excerpt": False,
        "has_remaining_diff": has_rem,
        "step1a_target": "".join(step1a_out),
        "true_old_baseline": "".join(baseline_out),
    }


def _norm_cn_key(s: str) -> str:
    return re.sub(r"[\s\u3000\u00A0]+", "", (s or "").strip())


def extract_col1_shading_spans(tc1: ET.Element) -> Tuple[Optional[str], List[Tuple[str, str]]]:
    """
    Extract non-marker Col 1 Chinese shading colors from `回稿`:
      - Ignores scattered-revision extraction filter light-blue markers (`STRIP_COL23_SHD_FILLS` such as `DAE9F7`),
        so library original Col 1 shadings/highlights remain untouched unless custom non-marker shading exists.
      - Returns `(cell_level_fill_hex_or_None, [(shaded_text_snippet, fill_hex), ...])`.
    """
    ignore_fills = {"AUTO", "FFFFFF", "NONE"} | STRIP_COL23_SHD_FILLS
    cell_fill: Optional[str] = None
    tc_pr = tc1.find("w:tcPr", NS)
    if tc_pr is not None:
        shd = tc_pr.find("w:shd", NS)
        if shd is not None:
            fv = (shd.attrib.get(f"{W}fill") or "").strip().upper()
            if fv and fv not in ignore_fills:
                cell_fill = fv

    run_spans: List[Tuple[str, str]] = []
    for p in tc1.findall("w:p", NS):
        p_fill: Optional[str] = None
        p_pr = p.find("w:pPr", NS)
        if p_pr is not None:
            pshd = p_pr.find("w:shd", NS)
            if pshd is not None:
                fv = (pshd.attrib.get(f"{W}fill") or "").strip().upper()
                if fv and fv not in ignore_fills:
                    p_fill = fv
        if p_fill:
            ptxt = "".join(t.text or "" for t in p.findall(".//w:t", NS)).strip()
            if ptxt:
                run_spans.append((ptxt, p_fill))
                continue

        cur_txt = ""
        cur_fill = ""
        for r in p.findall(".//w:r", NS):
            txt = "".join(t.text or "" for t in r.findall("w:t", NS))
            if not txt:
                continue
            rshd = r.find("./w:rPr/w:shd", NS)
            fv = (rshd.attrib.get(f"{W}fill") or "").strip().upper() if rshd is not None else ""
            if fv in ignore_fills:
                fv = ""
            if fv == cur_fill:
                cur_txt += txt
            else:
                if cur_fill and cur_txt.strip():
                    run_spans.append((cur_txt.strip(), cur_fill))
                cur_fill = fv
                cur_txt = txt
        if cur_fill and cur_txt.strip():
            run_spans.append((cur_txt.strip(), cur_fill))

    return cell_fill, run_spans


def apply_col1_shading_overlay(
    target_tc1: ET.Element,
    cell_fill: Optional[str],
    run_spans: Sequence[Tuple[str, str]],
) -> int:
    """
    Overlay Col 1 Chinese shading colors onto `target_tc1` WITHOUT clearing any pre-existing
    shadings/highlights elsewhere in `target_tc1`.
    Returns count of overlaid spans.
    """
    applied = 0
    if cell_fill:
        tc_pr = target_tc1.find("w:tcPr", NS)
        if tc_pr is None:
            tc_pr = ET.Element(f"{W}tcPr")
            target_tc1.insert(0, tc_pr)
        for old_shd in list(tc_pr.findall("w:shd", NS)):
            tc_pr.remove(old_shd)
        ET.SubElement(tc_pr, f"{W}shd", {f"{W}val": "clear", f"{W}color": "auto", f"{W}fill": cell_fill})
        applied += 1

    if not run_spans:
        return applied

    full_txt = docx_xml_utils.get_cell_plain_text(target_tc1)
    if not full_txt.strip():
        return applied

    char_spans: List[Tuple[int, int, str]] = []
    cursor = 0
    for snippet, fill_hex in run_spans:
        loc = docx_xml_utils.find_fuzzy_substring_span(full_txt, snippet, search_from=cursor)
        if loc is None and cursor > 0:
            loc = docx_xml_utils.find_fuzzy_substring_span(full_txt, snippet, search_from=0)
        if loc is not None:
            s_i, e_i = loc
            char_spans.append((s_i, e_i, f"shd:{fill_hex}"))
            cursor = e_i
            applied += 1

    if char_spans:
        # Apply overlay WITHOUT stripping existing shadings/highlights outside `char_spans`
        _overlay_run_shading_preserve_existing(target_tc1, char_spans)
    return applied


def _overlay_run_shading_preserve_existing(
    tc: ET.Element,
    char_spans: Sequence[Tuple[int, int, str]],
) -> None:
    """Apply `shd:HEX` to exact `[start, end)` character intervals in `tc` while preserving all existing formatting & shadings outside those intervals."""
    if not char_spans:
        return

    full_len = len(docx_xml_utils.get_cell_plain_text(tc))
    overlay_arr: List[Optional[str]] = [None] * full_len
    for s_i, e_i, spec in char_spans:
        fill = spec[4:] if spec.startswith("shd:") else spec
        for idx in range(max(0, s_i), min(full_len, e_i)):
            overlay_arr[idx] = fill

    pos = 0
    paras = tc.findall("w:p", NS)
    for p_idx, p in enumerate(paras):
        new_children: List[ET.Element] = []
        for ch in list(p):
            loc = ch.tag.split("}")[-1]
            if loc != "r":
                new_children.append(ch)
                continue
            r_txt = "".join(t.text or "" for t in ch.findall("w:t", NS))
            if not r_txt:
                new_children.append(ch)
                continue
            r_len = len(r_txt)
            seg_fills = [overlay_arr[pos + i] if (pos + i) < full_len else None for i in range(r_len)]
            pos += r_len
            if all(f is None for f in seg_fills):
                new_children.append(ch)
                continue
            base_rPr = ch.find("w:rPr", NS)
            # Group consecutive chars with identical overlay fill
            i = 0
            while i < r_len:
                j = i + 1
                while j < r_len and seg_fills[j] == seg_fills[i]:
                    j += 1
                piece = r_txt[i:j]
                ov_fill = seg_fills[i]
                if ov_fill is None:
                    r_copy = copy.deepcopy(ch)
                    for old_t in list(r_copy.findall("w:t", NS)):
                        r_copy.remove(old_t)
                    t_el = ET.SubElement(r_copy, f"{W}t")
                    t_el.text = piece
                    if piece[0].isspace() or piece[-1].isspace():
                        t_el.set(XML_SPACE, "preserve")
                    new_children.append(r_copy)
                else:
                    r_copy = copy.deepcopy(ch)
                    for old_t in list(r_copy.findall("w:t", NS)):
                        r_copy.remove(old_t)
                    rPr = r_copy.find("w:rPr", NS)
                    if rPr is None:
                        rPr = ET.Element(f"{W}rPr")
                        r_copy.insert(0, rPr)
                    docx_xml_utils.insert_rpr_child_ordered(
                        rPr,
                        ET.Element(f"{W}shd", {f"{W}val": "clear", f"{W}color": "auto", f"{W}fill": ov_fill}),
                    )
                    t_el = ET.SubElement(r_copy, f"{W}t")
                    t_el.text = piece
                    if piece[0].isspace() or piece[-1].isspace():
                        t_el.set(XML_SPACE, "preserve")
                    new_children.append(r_copy)
                i = j
        for ch in list(p):
            p.remove(ch)
        for ch in new_children:
            p.append(ch)
        if p_idx < len(paras) - 1:
            pos += 1  # newline between paragraphs


# ---------------------------------------------------------------------------
# Surgical Sentence/Token-Level Write-Back (Preserves Library Inline Styles!)
# ---------------------------------------------------------------------------

def _flatten_paragraph_chars(p: ET.Element) -> List[Tuple[str, Optional[ET.Element]]]:
    """Return list of `(char, rPr_element_or_None)` for all `<w:r>/<w:t>` text in paragraph `p`."""
    out: List[Tuple[str, Optional[ET.Element]]] = []
    for r in p.findall(".//w:r", NS):
        rPr = r.find("w:rPr", NS)
        txt = "".join(t.text or "" for t in r.findall("w:t", NS))
        for ch in txt:
            out.append((ch, rPr))
    return out


def _build_runs_from_char_styles(
    items: Sequence[Tuple[str, Optional[ET.Element]]],
    fallback_rPr: Optional[ET.Element],
) -> List[ET.Element]:
    """Coalesce adjacent `(char, rPr)` pairs sharing identical XML rPr into `<w:r>` elements."""
    if not items:
        return []
    runs: List[ET.Element] = []
    cur_txt: List[str] = []
    cur_key: Optional[str] = None
    cur_rPr: Optional[ET.Element] = None

    def _flush():
        if not cur_txt:
            return
        s = "".join(cur_txt)
        r_el = ET.Element(f"{W}r")
        use_rPr = cur_rPr if cur_rPr is not None else fallback_rPr
        if use_rPr is not None:
            r_el.append(copy.deepcopy(use_rPr))
        t_el = ET.SubElement(r_el, f"{W}t")
        t_el.text = s
        if s[0].isspace() or s[-1].isspace():
            t_el.set(XML_SPACE, "preserve")
        runs.append(r_el)

    for ch, rPr in items:
        key = ET.tostring(rPr, encoding="unicode") if rPr is not None else "__NONE__"
        if cur_txt and key == cur_key:
            cur_txt.append(ch)
        else:
            _flush()
            cur_txt = [ch]
            cur_key = key
            cur_rPr = rPr
    _flush()
    return runs


def surgical_update_paragraph_text(
    p: ET.Element,
    new_text: str,
    fallback_rPr: Optional[ET.Element] = None,
) -> bool:
    """
    Surgically update paragraph `p` from its current text to `new_text` using token-level diff,
    preserving every untouched word's exact `<w:rPr>` (font, size, blue scripture, red text,
    bold, italic, existing shading/highlights). Inserted/replaced tokens inherit the nearest
    neighboring run's `<w:rPr>`!
    Returns True if text actually changed.
    """
    char_styles = _flatten_paragraph_chars(p)
    old_text = "".join(c for c, _ in char_styles)
    if old_text == new_text:
        return False

    old_tokens = _TOKEN_RE.findall(old_text)
    new_tokens = _TOKEN_RE.findall(new_text)

    # Map each old token to its slice in char_styles
    old_tok_slices: List[List[Tuple[str, Optional[ET.Element]]]] = []
    pos = 0
    for tok in old_tokens:
        old_tok_slices.append(char_styles[pos : pos + len(tok)])
        pos += len(tok)

    sm = difflib.SequenceMatcher(None, old_tokens, new_tokens, autojunk=False)
    out_chars: List[Tuple[str, Optional[ET.Element]]] = []
    default_rPr = char_styles[0][1] if char_styles else fallback_rPr

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for oi in range(i1, i2):
                out_chars.extend(old_tok_slices[oi])
        elif tag in ("replace", "insert"):
            # Pick inherited style from replaced span first, else nearest preceding/following token
            inh_rPr = default_rPr
            if tag == "replace" and i1 < i2 and old_tok_slices[i1]:
                inh_rPr = old_tok_slices[i1][0][1]
            elif i1 > 0 and old_tok_slices[i1 - 1]:
                inh_rPr = old_tok_slices[i1 - 1][-1][1]
            elif i2 < len(old_tok_slices) and old_tok_slices[i2]:
                inh_rPr = old_tok_slices[i2][0][1]
            for nj in range(j1, j2):
                for ch in new_tokens[nj]:
                    out_chars.append((ch, inh_rPr))
        elif tag == "delete":
            continue

    pPr = p.find("w:pPr", NS)
    for ch in list(p):
        if ch is not pPr:
            p.remove(ch)
    for r_el in _build_runs_from_char_styles(out_chars, fallback_rPr=default_rPr):
        p.append(r_el)
    return True


def surgical_update_cell_text(tc: ET.Element, new_text: str) -> bool:
    """
    Update cell `tc` to `new_text` surgically paragraph-by-paragraph without clobbering
    untouched paragraphs or inline formatting (blue scripture, red text, existing shadings).
    """
    accept_all_revisions_in_cell_safe(tc)
    paras = tc.findall("w:p", NS)
    old_lines = ["".join(t.text or "" for t in p.findall(".//w:t", NS)) for p in paras]
    new_lines = (new_text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")

    if old_lines == new_lines:
        return False

    tpl_pPr = copy.deepcopy(paras[0].find("w:pPr", NS)) if (paras and paras[0].find("w:pPr", NS) is not None) else None
    tpl_rPr = docx_xml_utils.get_first_run_rpr(tc)

    changed = False
    if len(old_lines) == len(new_lines):
        for p, nl in zip(paras, new_lines):
            if surgical_update_paragraph_text(p, nl, fallback_rPr=tpl_rPr):
                changed = True
        return changed

    # Paragraph count differs: align paragraphs via SequenceMatcher so untouched paragraphs stay 100% verbatim!
    sm = difflib.SequenceMatcher(None, old_lines, new_lines, autojunk=False)
    new_para_els: List[ET.Element] = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for oi in range(i1, i2):
                new_para_els.append(paras[oi])
        elif tag == "replace":
            common = min(i2 - i1, j2 - j1)
            for k in range(common):
                p_cur = paras[i1 + k]
                surgical_update_paragraph_text(p_cur, new_lines[j1 + k], fallback_rPr=tpl_rPr)
                new_para_els.append(p_cur)
                changed = True
            for k in range(common, j2 - j1):
                p_new = ET.Element(f"{W}p")
                if tpl_pPr is not None:
                    p_new.append(copy.deepcopy(tpl_pPr))
                r_el = docx_xml_utils.make_run(new_lines[j1 + k], base_rPr=tpl_rPr)
                p_new.append(r_el)
                new_para_els.append(p_new)
                changed = True
        elif tag == "insert":
            for nj in range(j1, j2):
                p_new = ET.Element(f"{W}p")
                if tpl_pPr is not None:
                    p_new.append(copy.deepcopy(tpl_pPr))
                r_el = docx_xml_utils.make_run(new_lines[nj], base_rPr=tpl_rPr)
                p_new.append(r_el)
                new_para_els.append(p_new)
                changed = True
        elif tag == "delete":
            changed = True

    if not new_para_els:
        p_empty = ET.Element(f"{W}p")
        if tpl_pPr is not None:
            p_empty.append(copy.deepcopy(tpl_pPr))
        new_para_els.append(p_empty)

    tcPr = tc.find("w:tcPr", NS)
    for ch in list(tc):
        if ch is not tcPr:
            tc.remove(ch)
    for p_el in new_para_els:
        tc.append(p_el)
    return changed


def merge_keep_library_text(sent_rejected: str, lib_current: str, returned_clean: str) -> str:
    """
    3-way merge for conflict policy ① (`keep_library`):
      - `sent_rejected` is the base.
      - `lib_current` has changes made in the library since dispatch -> KEEP these!
      - `returned_clean` has the translator's edits -> APPLY where they don't collide with `lib_current`.
    """
    base_t = _TOKEN_RE.findall(sent_rejected or "")
    lib_t = _TOKEN_RE.findall(lib_current or "")
    ret_t = _TOKEN_RE.findall(returned_clean or "")

    # Identify base token indices modified in `lib_current`
    lib_ops = difflib.SequenceMatcher(None, base_t, lib_t, autojunk=False).get_opcodes()
    ret_ops = difflib.SequenceMatcher(None, base_t, ret_t, autojunk=False).get_opcodes()

    lib_edits: Dict[int, Tuple[int, List[str]]] = {}
    for tag, i1, i2, j1, j2 in lib_ops:
        if tag != "equal":
            lib_edits[i1] = (i2, lib_t[j1:j2])

    ret_edits: Dict[int, Tuple[int, List[str]]] = {}
    for tag, i1, i2, j1, j2 in ret_ops:
        if tag != "equal":
            ret_edits[i1] = (i2, ret_t[j1:j2])

    out: List[str] = []
    idx = 0
    n = len(base_t)
    while idx <= n:
        if idx in lib_edits:
            end_i, rep = lib_edits[idx]
            out.extend(rep)
            idx = max(end_i, idx + 1) if end_i > idx else idx
            if end_i == idx:
                if idx < n:
                    out.append(base_t[idx])
                idx += 1
            continue
        if idx in ret_edits:
            end_i, rep = ret_edits[idx]
            # Check if any index in [idx, end_i) collides with a library edit
            collides = any(k in lib_edits for k in range(idx, max(idx + 1, end_i)))
            if not collides:
                out.extend(rep)
                idx = max(end_i, idx + 1) if end_i > idx else idx
                if end_i == idx:
                    if idx < n:
                        out.append(base_t[idx])
                    idx += 1
                continue
        if idx < n:
            out.append(base_t[idx])
        idx += 1
    return "".join(out)


def render_diff_track_changes_into_cell(
    tc: ET.Element,
    old_text: str,
    new_text: str,
    author: str = "Review",
    date_iso: str = "",
    rev_id_counter: Optional[List[int]] = None,
) -> bool:
    """
    Replace `tc` paragraphs with native Word `<w:del>` / `<w:ins>` track changes showing
    the diff from `old_text` -> `new_text` while keeping `tc`'s existing paragraph & run styles.
    Returns True if any diff exists.
    """
    if not date_iso:
        date_iso = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
    if rev_id_counter is None:
        rev_id_counter = [100]

    old_norm = (old_text or "").replace("\r\n", "\n").replace("\r", "\n")
    new_norm = (new_text or "").replace("\r\n", "\n").replace("\r", "\n")
    if old_norm == new_norm:
        return False

    tpl_pPr = docx_xml_utils.get_cell_template_ppr(tc)
    tpl_rPr = docx_xml_utils.get_first_run_rpr(tc)

    old_lines = old_norm.split("\n")
    new_lines = new_norm.split("\n")
    max_lines = max(len(old_lines), len(new_lines))

    new_paras: List[ET.Element] = []
    for li in range(max_lines):
        ol = old_lines[li] if li < len(old_lines) else ""
        nl = new_lines[li] if li < len(new_lines) else ""
        p_el = ET.Element(f"{W}p")
        if tpl_pPr is not None:
            p_el.append(copy.deepcopy(tpl_pPr))
        if ol == nl:
            if nl:
                p_el.append(docx_xml_utils.make_run(nl, base_rPr=tpl_rPr))
        else:
            o_toks = _TOKEN_RE.findall(ol)
            n_toks = _TOKEN_RE.findall(nl)
            sm = difflib.SequenceMatcher(None, o_toks, n_toks, autojunk=False)
            for tag, i1, i2, j1, j2 in sm.get_opcodes():
                if tag == "equal":
                    s = "".join(o_toks[i1:i2])
                    if s:
                        p_el.append(docx_xml_utils.make_run(s, base_rPr=tpl_rPr))
                elif tag == "delete":
                    s = "".join(o_toks[i1:i2])
                    if s:
                        rev_id_counter[0] += 1
                        p_el.append(
                            docx_xml_utils.make_track_el(
                                "del",
                                rev_id_counter[0],
                                author,
                                date_iso,
                                docx_xml_utils.make_run(s, base_rPr=tpl_rPr, is_del=True),
                            )
                        )
                elif tag == "insert":
                    s = "".join(n_toks[j1:j2])
                    if s:
                        rev_id_counter[0] += 1
                        p_el.append(
                            docx_xml_utils.make_track_el(
                                "ins",
                                rev_id_counter[0],
                                author,
                                date_iso,
                                docx_xml_utils.make_run(s, base_rPr=tpl_rPr),
                            )
                        )
                elif tag == "replace":
                    s_del = "".join(o_toks[i1:i2])
                    s_ins = "".join(n_toks[j1:j2])
                    if s_del:
                        rev_id_counter[0] += 1
                        p_el.append(
                            docx_xml_utils.make_track_el(
                                "del",
                                rev_id_counter[0],
                                author,
                                date_iso,
                                docx_xml_utils.make_run(s_del, base_rPr=tpl_rPr, is_del=True),
                            )
                        )
                    if s_ins:
                        rev_id_counter[0] += 1
                        p_el.append(
                            docx_xml_utils.make_track_el(
                                "ins",
                                rev_id_counter[0],
                                author,
                                date_iso,
                                docx_xml_utils.make_run(s_ins, base_rPr=tpl_rPr),
                            )
                        )
        new_paras.append(p_el)

    tcPr = tc.find("w:tcPr", NS)
    for ch in list(tc):
        if ch is not tcPr:
            tc.remove(ch)
    for p_el in new_paras:
        tc.append(p_el)
    return True


def add_comment_to_cell(
    tc: ET.Element,
    comment_text: str,
    comments_root: ET.Element,
    id_allocator: docx_xml_utils.CommentIdAllocator,
    author: str = "验收检查",
) -> None:
    """Attach a native Word comment (`<w:comment>`) spanning the first paragraph of `tc`."""
    p = tc.find("w:p", NS)
    if p is None:
        p = ET.SubElement(tc, f"{W}p")
    cid = str(id_allocator.next_id())
    date_iso = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")

    c_el = ET.SubElement(
        comments_root,
        f"{W}comment",
        {f"{W}id": cid, f"{W}author": author, f"{W}date": date_iso, f"{W}initials": "QA"},
    )
    for ln in (comment_text or "").splitlines():
        cp = ET.SubElement(c_el, f"{W}p")
        cr = ET.SubElement(cp, f"{W}r")
        ct = ET.SubElement(cr, f"{W}t")
        ct.text = ln
        if ln and (ln[0].isspace() or ln[-1].isspace()):
            ct.set(XML_SPACE, "preserve")

    pPr = p.find("w:pPr", NS)
    ins_pos = 1 if (pPr is not None and len(p) > 0 and p[0] is pPr) else 0
    p.insert(ins_pos, ET.Element(f"{W}commentRangeStart", {f"{W}id": cid}))
    p.append(ET.Element(f"{W}commentRangeEnd", {f"{W}id": cid}))
    ref_r = ET.Element(f"{W}r")
    ET.SubElement(ref_r, f"{W}commentReference", {f"{W}id": cid})
    p.append(ref_r)


def save_docx_without_track_revisions_flag(
    src_path: Path,
    dst_path: Path,
    root: ET.Element,
    ns_pairs: Sequence[Tuple[str, str]],
    comments_root: Optional[ET.Element] = None,
) -> None:
    """Save `.docx`, slim embedded fonts/thumbnails, and disable `<w:trackRevisions/>` inside `word/settings.xml`."""
    docx_xml_utils.save_docx_xml(
        src_path,
        dst_path,
        root,
        ns_pairs,
        comments_root=comments_root,
        disable_track_revisions=True,
    )


# ---------------------------------------------------------------------------
# Auto-Pairing of Sent (`发稿`) and Returned (`回稿`) Files
# ---------------------------------------------------------------------------

def _collect_docx_files(path_str: Union[str, Path]) -> List[Path]:
    p = Path(path_str).resolve()
    if p.is_file():
        return [p] if p.suffix.lower() == ".docx" else []
    if p.is_dir():
        return sorted(
            f
            for f in p.rglob("*.docx")
            if not f.name.startswith("~$") and not f.name.startswith(".")
        )
    return []


def _extract_doc_cn_signature(docx_path: Path) -> set:
    try:
        root, _, _ = docx_xml_utils.load_docx_xml(docx_path)
        sig = set()
        for tr in root.findall(".//w:tbl/w:tr", NS):
            tcs = tr.findall("w:tc", NS)
            if len(tcs) >= 3:
                k = _norm_cn_key(docx_xml_utils.get_cell_plain_text(tcs[0]))
                if len(k) >= 6:
                    sig.add(k[:40])
        return sig
    except Exception:
        return set()


def pair_sent_and_returned_files(
    sent_path: Union[str, Path],
    returned_path: Union[str, Path],
) -> Tuple[List[Tuple[Path, Path]], List[Path], List[Path]]:
    """
    Auto-pair single or batch folders of `发稿` and `回稿` using filename similarity + Col 1 Chinese overlap.
    Returns `(paired_list, unpaired_sent, unpaired_returned)`.
    """
    sent_files = _collect_docx_files(sent_path)
    ret_files = _collect_docx_files(returned_path)
    if not sent_files:
        raise FileNotFoundError(f"未在发稿路径中找到 .docx 文档：{sent_path}")
    if not ret_files:
        raise FileNotFoundError(f"未在回稿路径中找到 .docx 文档：{returned_path}")

    if len(sent_files) == 1 and len(ret_files) == 1:
        return [(sent_files[0], ret_files[0])], [], []

    sent_sigs = {f: _extract_doc_cn_signature(f) for f in sent_files}
    ret_sigs = {f: _extract_doc_cn_signature(f) for f in ret_files}

    def _clean_name_for_pair(name: str) -> str:
        s = re.sub(r"^\d+\s+", "", name)
        s = re.sub(r"(?:[-_\s]*(?:净稿|修订稿|回稿|定稿|副本|\(\d+\)|\（\d+\）))+$", "", s)
        return re.sub(r"[\s_\-]+", "", s).lower()

    scored: List[Tuple[float, Path, Path]] = []
    for sf in sent_files:
        sn = _clean_name_for_pair(sf.stem)
        ss = sent_sigs[sf]
        for rf in ret_files:
            rn = _clean_name_for_pair(rf.stem)
            rs = ret_sigs[rf]
            name_sim = difflib.SequenceMatcher(None, sn, rn).ratio()
            overlap = (len(ss & rs) / max(1, min(len(ss), len(rs)))) if (ss and rs) else 0.0
            total = overlap * 0.70 + name_sim * 0.30
            scored.append((total, sf, rf))

    scored.sort(key=lambda x: x[0], reverse=True)
    used_s = set()
    used_r = set()
    pairs: List[Tuple[Path, Path]] = []
    for score, sf, rf in scored:
        if sf in used_s or rf in used_r:
            continue
        if score >= 0.35:
            pairs.append((sf, rf))
            used_s.add(sf)
            used_r.add(rf)

    unpaired_s = [f for f in sent_files if f not in used_s]
    unpaired_r = [f for f in ret_files if f not in used_r]
    return pairs, unpaired_s, unpaired_r


# ---------------------------------------------------------------------------
# Step 1B: Anomaly Heuristics on Returned Col 3
# ---------------------------------------------------------------------------

def _detect_col3_anomalies(
    sent_tc2: ET.Element,
    sent_col2_rej: str,
    ret_col2_clean: str,
    sent_col3_clean: str,
    ret_col3_clean: str,
    ret_col3_comments: Sequence[str],
) -> List[str]:
    """
    Detect structural/severe errors in translator's Col 3 revision (NOT grammar/wording nitpicking):
      1. Missed edit (Col 2 changed, Col 3 untouched, and no '不需要改' translator comment).
      2. Excessive sentence deletion / truncation on non-cyan small edits.
      3. Duplicate consecutive sentence accidentally pasted.
      4. Wrong-location edit (Col 2 edit in one half of paragraph, Col 3 edit in opposite half).
      5. Cyan RV boundary / severe sentence-count mismatch.
    """
    issues: List[str] = []
    has_c2_edit = _norm_compare_text(sent_col2_rej) != _norm_compare_text(ret_col2_clean) or cell_has_sent_edits(sent_tc2)
    c3_changed = _norm_compare_text(sent_col3_clean) != _norm_compare_text(ret_col3_clean)
    is_cyan = cell_has_cyan_rv(sent_tc2)

    if not has_c2_edit:
        return issues

    # 1. Missed edit check (unless translator left a "不需要改" comment)
    if not c3_changed:
        has_valid_excuse = any(_NO_CHANGE_COMMENT_RE.search(c) for c in ret_col3_comments)
        if not has_valid_excuse:
            issues.append("【疑似漏改】发稿第 2 列英文有改点，但回稿第 3 列外文未作任何修改（且无「不需要改」类批注说明）。")
        return issues

    # 2. Check sentence deletion / truncation
    old_sents = [s.strip() for s in _SENT_SPLIT_RE.split(sent_col3_clean.strip()) if s.strip()]
    new_sents = [s.strip() for s in _SENT_SPLIT_RE.split(ret_col3_clean.strip()) if s.strip()]
    en_old_len = max(1, len(sent_col2_rej.strip()))
    en_new_len = max(1, len(ret_col2_clean.strip()))
    en_ratio = en_new_len / en_old_len
    f_old_len = max(1, len(sent_col3_clean.strip()))
    f_new_len = max(1, len(ret_col3_clean.strip()))
    f_ratio = f_new_len / f_old_len

    if len(old_sents) >= 3 and len(new_sents) <= len(old_sents) - 2 and en_ratio > 0.85:
        issues.append(
            f"【疑似多删句子】外文原句数 {len(old_sents)} 句，改后仅剩 {len(new_sents)} 句（英文篇幅基本未缩减），请核对是否误删整句。"
        )
    elif not is_cyan and f_old_len >= 80 and f_ratio < 0.65 and en_ratio > 0.85:
        issues.append(
            f"【疑似大段误删】非青色局部改点，但外文篇幅缩减超过 35%（{f_old_len} -> {f_new_len} 字符），请核对是否误删内容。"
        )

    # 3. Duplicate consecutive sentence check
    for idx in range(len(new_sents) - 1):
        s_a = re.sub(r"\s+", " ", new_sents[idx]).lower()
        s_b = re.sub(r"\s+", " ", new_sents[idx + 1]).lower()
        if len(s_a) >= 25 and s_a == s_b:
            issues.append("【疑似重复句】回稿第 3 列出现连续完全相同的长句，请核对是否重复粘贴。")
            break

    # 4. Wrong-location edit check (only on longer paragraphs with localized single edit)
    if not is_cyan and len(sent_col2_rej) >= 120 and len(sent_col3_clean) >= 120:
        en_ops = [op for op in difflib.SequenceMatcher(None, sent_col2_rej, ret_col2_clean, autojunk=False).get_opcodes() if op[0] != "equal"]
        f_ops = [op for op in difflib.SequenceMatcher(None, sent_col3_clean, ret_col3_clean, autojunk=False).get_opcodes() if op[0] != "equal"]
        if en_ops and f_ops:
            en_center = ((en_ops[0][1] + en_ops[-1][2]) * 0.5) / en_old_len
            f_center = ((f_ops[0][1] + f_ops[-1][2]) * 0.5) / f_old_len
            en_span = (en_ops[-1][2] - en_ops[0][1]) / en_old_len
            f_span = (f_ops[-1][2] - f_ops[0][1]) / f_old_len
            if en_span < 0.25 and f_span < 0.25 and abs(en_center - f_center) > 0.48:
                issues.append(
                    f"【疑似改错位置】英文改点位于段落约 {int(en_center * 100)}% 处，但外文改动出现在约 {int(f_center * 100)}% 处，请核对是否误改到了其他类似短语。"
                )

    # 5. Cyan RV boundary check
    if is_cyan:
        if f_new_len < f_old_len * 0.45 and en_ratio > 0.75:
            issues.append("【青色重翻衔接异常】青色重翻片段改后外文篇幅明显过短，请重点核对前后文衔接句是否遗漏。")
        elif abs(len(old_sents) - len(new_sents)) >= 3 and abs(len(_SENT_SPLIT_RE.split(sent_col2_rej)) - len(_SENT_SPLIT_RE.split(ret_col2_clean))) <= 1:
            issues.append("【青色重翻句数差异】青色重翻外文句数变化较大，请核对前后文衔接处是否与中英一致。")

    return issues


# ---------------------------------------------------------------------------
# Step 2: Clean Formatting & Apply Optional Proofread Table Updates
# ---------------------------------------------------------------------------

def clean_cell_text_artifacts(text: str) -> str:
    """Strip double spaces, spaces before punctuation, and excess blank lines."""
    lines = (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    out: List[str] = []
    for ln in lines:
        s = _DOUBLE_SPACE_RE.sub(" ", ln.replace("\u00A0", " "))
        s = _PUNCT_SPACE_RE.sub(r"\1", s)
        s = s.strip()
        if s:
            out.append(s)
    return "\n".join(out)


def clean_col23_highlights_and_artifacts(tc: ET.Element) -> None:
    """
    Clean Col 2 or Col 3 cell in `回稿`:
      - Accept all track changes & unwrap `<w:sdt>`
      - Strip comments
      - Strip yellow/cyan/blue/gray `<w:highlight>` & UP light-blue `<w:shd>`
      - Normalize double spaces, space-before-punct, and empty trailing paragraphs.
    """
    accept_all_revisions_in_cell_safe(tc)
    strip_comments_in_cell(tc)
    for rPr in tc.findall(".//w:rPr", NS):
        hl = rPr.find("w:highlight", NS)
        if hl is not None and hl.attrib.get(f"{W}val") in STRIP_COL23_HIGHLIGHTS:
            rPr.remove(hl)
        shd = rPr.find("w:shd", NS)
        if shd is not None and (shd.attrib.get(f"{W}fill") or "").upper() in STRIP_COL23_SHD_FILLS:
            rPr.remove(shd)
    tcPr = tc.find("w:tcPr", NS)
    if tcPr is not None:
        shd = tcPr.find("w:shd", NS)
        if shd is not None and (shd.attrib.get(f"{W}fill") or "").upper() in STRIP_COL23_SHD_FILLS:
            tcPr.remove(shd)

    # Remove empty redundant paragraphs (keep at least 1 paragraph per cell for valid OOXML)
    paras = tc.findall("w:p", NS)
    non_empty = [p for p in paras if "".join(t.text or "" for t in p.findall(".//w:t", NS)).strip()]
    if non_empty and len(non_empty) < len(paras):
        for p in paras:
            if p not in non_empty:
                tc.remove(p)

    # Clean double spaces & space-before-punctuation inside text runs
    for p in tc.findall("w:p", NS):
        cur_txt = "".join(t.text or "" for t in p.findall(".//w:t", NS))
        cleaned = clean_cell_text_artifacts(cur_txt)
        if cleaned != cur_txt:
            surgical_update_paragraph_text(p, cleaned)


PROOFREAD_5COL_HEADERS = [
    "来源表格与行号",
    "状态 / 异常说明",
    "中文",
    "English",
    "改前列（目标语）",
    "定稿列（Final Version-全）",
    "改动列（Final Version-改动行）",
]
PROOFREAD_5COL_WIDTHS = [1550, 2020, 2000, 2200, 2260, 2270, 2270]

_CLOSING_WRAPPERS = '”"’\'」』）)》】»“„›'
_SENT_END_CHARS = set(".?!。？！…။।។")
_CLAUSE_END_CHARS = set(",;:，；：、—–-၊")
_FIRST_ALPHA_RE = re.compile(r'^[“"‘\'(\[（【«„\s]*([A-Za-zÀ-ÖØ-öø-ÿА-Яа-яЁё])')


def extract_proofread_sources_list(raw_sources: Optional[Sequence[str]]) -> List[str]:
    """
    Extract all Google Sheets URLs (each preserving its `#gid=...`) and local `.xlsx/.csv` paths
    from multi-line text, 2-column `篇目标题\\t链接` paste, or comma/semicolon-separated inputs.
    """
    if not raw_sources:
        return []
    blob = "\n".join(str(x or "") for x in raw_sources if str(x or "").strip())
    if not blob.strip():
        return []
    out: List[str] = []
    seen = set()

    # First try structured 2-column / multi-URL parser
    for _, urls in parse_title_link_table(blob):
        for u in urls:
            u_clean = u.strip().strip('"').strip("'")
            if u_clean and u_clean not in seen:
                if is_google_sheet_url(u_clean) or Path(u_clean).suffix.lower() in (".xlsx", ".csv"):
                    seen.add(u_clean)
                    out.append(u_clean)

    # Also scan raw lines directly in case of standalone local paths or URLs
    for raw_ln in blob.splitlines():
        ln = raw_ln.strip()
        if not ln:
            continue
        urls = re.findall(r"https?://[^\s,，;；\t]+", ln)
        if urls:
            for u in urls:
                if u not in seen:
                    seen.add(u)
                    out.append(u)
        else:
            for seg in raw_ln.split("\t"):
                cand = seg.strip().strip('"').strip("'")
                if cand and cand not in seen and Path(cand).suffix.lower() in (".xlsx", ".csv"):
                    seen.add(cand)
                    out.append(cand)
    return out


def _classify_clause_or_sentence_end(text: str) -> Tuple[str, str]:
    """
    Classify the ending punctuation of `text` (ignoring trailing closing quotes/brackets):
      - `("SENT", ch)`   : full sentence end (`. ? ! 。 ？ ！ … ။ । ។`)
      - `("CLAUSE", ch)` : half-sentence / clause pause (`, ; : ， ； ： 、 — – - ၊`)
      - `("NONE", "")`   : no ending punctuation
    """
    s = (text or "").strip().rstrip(_CLOSING_WRAPPERS).rstrip()
    if not s:
        return "NONE", ""
    last_ch = s[-1]
    if last_ch in _SENT_END_CHARS:
        return "SENT", last_ch
    if last_ch in _CLAUSE_END_CHARS:
        return "CLAUSE", last_ch
    return "NONE", ""


def _count_sentence_stops(text: str) -> int:
    s = (text or "").strip()
    if not s:
        return 0
    cnt = sum(1 for ch in re.sub(r"\.{2,}", "…", s) if ch in _SENT_END_CHARS)
    return max(1, cnt) if cnt > 0 else 0


def pick_proofread_new_text(pr: ProofreadRow) -> str:
    """
    Pick the effective post-proofread target text from `ProofreadRow`:
      - Prefer `改动列` (`pr.changes`) when non-empty and different from `pr.target`;
      - Otherwise use `定稿列` (`pr.final`) when non-empty and different from `pr.target`.
    """
    old_s = clean_cell_spaces(pr.target)
    chg_s = clean_cell_spaces(pr.changes)
    fin_s = clean_cell_spaces(pr.final)
    if chg_s and chg_s != old_s:
        return chg_s
    if fin_s and fin_s != old_s:
        return fin_s
    return ""


def make_proofread_5col_row(source_pos: str, status_reason: str, pr: ProofreadRow) -> List[str]:
    """Format a 7-cell report row containing location, status/reason, and all 5 original proofread columns."""
    return [
        source_pos,
        status_reason,
        clean_cell_spaces(pr.cn),
        clean_cell_spaces(pr.en),
        clean_cell_spaces(pr.target),
        clean_cell_spaces(pr.final),
        clean_cell_spaces(pr.changes),
    ]


def _batch_check_proofread_semantic_with_ai(
    items: Sequence[Tuple[int, ProofreadRow, str, str]],
    api_key: Union[str, List[str], ApiKeyPool],
    model: str = "gemini-3.5-flash-lite",
    api_mode: str = "auto",
) -> Dict[int, Tuple[bool, str]]:
    """
    Batch-verify changed proofread rows against Chinese (`r.cn`) & English (`r.en`) via Gemini:
    checks whether meaning matches Chinese, or whether a full sentence became half a sentence / vice versa.
    Returns `{item_idx: (is_consistent, reason)}`.
    """
    results: Dict[int, Tuple[bool, str]] = {}
    if not items or not api_key:
        return results

    batch_size = 10
    for b_start in range(0, len(items), batch_size):
        chunk = items[b_start : b_start + batch_size]
        lines_prompt = [
            "你是多语种翻译定稿质检专家。请逐条核对以下通读修改行的【改后外文】是否与【中文原文】（及英文参考）在意思和句子完整性上吻合：",
            "重点排查：1) 原来是一整句被误删成半句（漏译）；2) 原来是半句被误扩成一整句或串入了上下行内容；3) 改后意思与中文原文不符或相反。",
            "若只是正常的近义词润色、语法微调且意思完整吻合，请判为 `consistent: true`。",
            '请严格只返回 JSON：{"items": [{"id": 整数ID, "consistent": true或false, "reason": "若不一致请简明指出原因，一致则留空"}]}',
            "",
        ]
        for item_id, pr, old_t, new_t in chunk:
            lines_prompt.append(
                f"[ID={item_id}]\n"
                f"中文：{clean_cell_spaces(pr.cn)}\n"
                f"English：{clean_cell_spaces(pr.en)}\n"
                f"改前外文：{old_t}\n"
                f"改后外文：{new_t}\n"
            )
        try:
            resp = call_gemini_json(
                api_key=api_key,
                model=model or "gemini-3.5-flash-lite",
                system_instruction="你是严谨的多语种翻译通读质检助手，只输出合法 JSON。",
                user_prompt="\n".join(lines_prompt),
                api_mode=api_mode or "auto",
            )
            if isinstance(resp, dict) and isinstance(resp.get("items"), list):
                for obj in resp["items"]:
                    if isinstance(obj, dict) and "id" in obj:
                        iid = int(obj["id"])
                        ok = bool(obj.get("consistent", True))
                        rsn = str(obj.get("reason") or "").strip()
                        results[iid] = (ok, rsn)
        except Exception:
            continue
    return results


def evaluate_proofread_sheet_rows(
    parsed: ParsedSheetData,
    chapter_label: str = "",
    enable_ai: bool = False,
    api_key: Union[str, List[str], ApiKeyPool] = "",
    model: str = "gemini-3.5-flash-lite",
    api_mode: str = "auto",
) -> Tuple[List[dict], List[List[str]]]:
    """
    Unified strict validator for both Scenario 3A and Scenario 3B proofread sheets:
      1. `改动列 == 改前列` (when `改动列` is non-empty) -> DO NOT replace, report all 5 columns.
      2. `改动列 != 定稿列` (when both columns exist/are populated) -> DO NOT replace, report all 5 columns.
      3. All valid changed rows -> compare `改前` vs `改后` for:
         - 半句变整句 / 整句变半句 / 末尾标点丢失 / 句号后下一行首字母小写 / 句数突变 / 篇幅暴涨暴跌
         - 与中文原文意思是否吻合（AI 批量核对或本地大幅改动预警）
         Any abnormal changed row is reported with all 5 columns.
    Returns `(replaceable_items, blocked_5col_rows)` where each item in `replaceable_items` is:
      `{"row": ProofreadRow, "source_pos": str, "old_text": str, "new_text": str, "warnings": List[str]}`.
    """
    src_short = parsed.sheet_name or Path(parsed.source).name
    if chapter_label:
        src_prefix = f"{chapter_label} [{src_short}]"
    else:
        src_prefix = src_short

    has_changes_col = ("changes" in parsed.column_indices) and any(
        bool(clean_cell_spaces(r.changes)) for r in parsed.rows
    )
    has_final_col = ("final" in parsed.column_indices) and any(
        bool(clean_cell_spaces(r.final)) for r in parsed.rows
    )

    replaceable: List[dict] = []
    blocked_5col_rows: List[List[str]] = []
    ai_candidates: List[Tuple[int, ProofreadRow, str, str]] = []

    for idx, r in enumerate(parsed.rows):
        old_t = clean_cell_spaces(r.target)
        fin_t = clean_cell_spaces(r.final)
        chg_t = clean_cell_spaces(r.changes)
        pos_label = f"{src_prefix} 第{r.row_number}行"

        if not old_t and not fin_t and not chg_t:
            continue

        # Rule 3b: `改动列` non-empty and identical to `改前列` -> report 5 columns, do NOT replace
        if chg_t and chg_t == old_t:
            blocked_5col_rows.append(
                make_proofread_5col_row(pos_label, "未替换：改动列与改前列内容相同（无实际改动）", r)
            )
            continue

        # Rule 3a: `改动列` and `定稿列` differ -> CANNOT replace, report 5 columns
        if chg_t and fin_t and chg_t != fin_t:
            blocked_5col_rows.append(
                make_proofread_5col_row(pos_label, "禁止替换：改动列与定稿列内容不一致", r)
            )
            continue

        if has_changes_col and has_final_col:
            if fin_t and fin_t != old_t and not chg_t:
                blocked_5col_rows.append(
                    make_proofread_5col_row(pos_label, "禁止替换：定稿列有改动但改动列为空（改动列与定稿列不一致）", r)
                )
                continue
            if chg_t and chg_t != old_t and not fin_t:
                blocked_5col_rows.append(
                    make_proofread_5col_row(pos_label, "禁止替换：改动列有改动但定稿列为空（改动列与定稿列不一致）", r)
                )
                continue

        # Determine effective new text for genuinely modified rows
        if chg_t and chg_t != old_t:
            eff_new = chg_t
        elif (not has_changes_col) and fin_t and fin_t != old_t:
            eff_new = fin_t
        else:
            continue

        if not old_t:
            blocked_5col_rows.append(
                make_proofread_5col_row(pos_label, "禁止替换：改前列为空，无法定位原句", r)
            )
            continue

        # Rule 3c: Compare `改前` vs `改后` on every changed row (half-sentence <-> full-sentence, length, punctuation)
        warnings: List[str] = []
        old_kind, old_p = _classify_clause_or_sentence_end(old_t)
        new_kind, new_p = _classify_clause_or_sentence_end(eff_new)

        if old_kind in ("CLAUSE", "NONE") and new_kind == "SENT":
            msg = f"【半句变整句】改前末尾为半句（{f'标点「{old_p}」' if old_p else '无句末标点'}），改后变为整句标点「{new_p}」，请核对是否与中文吻合"
            if old_kind == "CLAUSE" and idx + 1 < len(parsed.rows):
                nxt = parsed.rows[idx + 1]
                nxt_txt = clean_cell_spaces(nxt.changes) or clean_cell_spaces(nxt.final) or clean_cell_spaces(nxt.target)
                m_al = _FIRST_ALPHA_RE.match(nxt_txt)
                if m_al and m_al.group(1).islower():
                    msg += f"（且下一行第{nxt.row_number}行首字母「{m_al.group(1)}」仍为小写）"
            warnings.append(msg)
        elif old_kind == "SENT" and new_kind in ("CLAUSE", "NONE"):
            warnings.append(
                f"【整句变半句】改前为完整句（末尾「{old_p}」），改后变为{f'半句标点「{new_p}」' if new_p else '无末尾标点'}，请核对是否漏译或与中文不符"
            )
        elif old_kind == "CLAUSE" and new_kind == "NONE":
            warnings.append(f"【末尾标点丢失】改前末尾有标点「{old_p}」，改后末尾标点缺失")

        old_sc = _count_sentence_stops(old_t)
        new_sc = _count_sentence_stops(eff_new)
        if old_sc >= 1 and new_sc >= 1 and old_sc != new_sc and not any("半句" in w for w in warnings):
            warnings.append(f"【句数变化】改前约 {old_sc} 句变为改后 {new_sc} 句，请核对是否与中文一致")

        len_ratio = len(eff_new) / max(1, len(old_t))
        old_words = old_t.split()
        new_words = eff_new.split()
        word_sim = (
            difflib.SequenceMatcher(None, [w.lower() for w in old_words], [w.lower() for w in new_words], autojunk=False).ratio()
            if old_words
            else 1.0
        )

        if len(old_t) >= 16 and len_ratio < 0.68:
            warnings.append(
                f"【疑似一句变半句/大幅删减】长度由 {len(old_t)} 缩减至 {len(eff_new)} 字符（-{int((1 - len_ratio) * 100)}%），请对照中文核对有无漏译"
            )
        elif len(old_t) >= 16 and len_ratio > 1.45:
            warnings.append(
                f"【疑似半句扩写/大幅增加】长度由 {len(old_t)} 增加至 {len(eff_new)} 字符（+{int((len_ratio - 1) * 100)}%），请对照中文核对是否多译或串行"
            )

        rec_idx = len(replaceable)
        replaceable.append({
            "row": r,
            "source_pos": pos_label,
            "old_text": old_t,
            "new_text": eff_new,
            "warnings": warnings,
            "word_sim": word_sim,
        })
        if enable_ai and api_key and clean_cell_spaces(r.cn):
            ai_candidates.append((rec_idx, r, old_t, eff_new))
        elif len(old_words) >= 5 and word_sim < 0.55:
            warnings.append(f"【大幅改动需核对中文】词相似度较低（{word_sim:.2f}），请人工核对改后意思是否与中文一致")

    if ai_candidates:
        ai_res = _batch_check_proofread_semantic_with_ai(ai_candidates, api_key=api_key, model=model, api_mode=api_mode)
        for rec_idx, pr, _, _ in ai_candidates:
            if rec_idx in ai_res:
                is_ok, rsn = ai_res[rec_idx]
                if not is_ok:
                    replaceable[rec_idx]["warnings"].append(
                        f"【与中文意思不吻合（AI核对）】{rsn or '改后译文与中文原文存在语义或完整性差异，请人工复核'}"
                    )
            else:
                ws = replaceable[rec_idx]["word_sim"]
                if len(replaceable[rec_idx]["old_text"].split()) >= 5 and ws < 0.55:
                    replaceable[rec_idx]["warnings"].append(
                        f"【大幅改动需核对中文】词相似度较低（{ws:.2f}），请人工核对改后意思是否与中文一致"
                    )

    return replaceable, blocked_5col_rows


# ---------------------------------------------------------------------------
# Report & Aggregated Revision Document Builders
# ---------------------------------------------------------------------------

def _write_simple_report_docx(
    out_path: Path,
    title: str,
    summary_lines: Sequence[str],
    headers: Sequence[str],
    rows_data: Sequence[Sequence[str]],
    col_widths: Optional[Sequence[int]] = None,
    extra_sections: Optional[Sequence[Tuple[str, Sequence[str], Sequence[Sequence[str]], Optional[Sequence[int]]]]] = None,
) -> Path:
    doc = ET.Element(f"{W}document")
    body = ET.SubElement(doc, f"{W}body")

    def _add_p(text: str, bold: bool = False, sz: str = "22", color: str = ""):
        p = ET.SubElement(body, f"{W}p")
        r = ET.SubElement(p, f"{W}r")
        rPr = ET.SubElement(r, f"{W}rPr")
        ET.SubElement(rPr, f"{W}rFonts", {f"{W}ascii": "Georgia", f"{W}eastAsia": "宋体", f"{W}hAnsi": "Georgia"})
        if bold:
            ET.SubElement(rPr, f"{W}b")
        if color:
            ET.SubElement(rPr, f"{W}color", {f"{W}val": color})
        ET.SubElement(rPr, f"{W}sz", {f"{W}val": sz})
        ET.SubElement(rPr, f"{W}szCs", {f"{W}val": sz})
        t = ET.SubElement(r, f"{W}t")
        t.text = text
        if text and (text[0].isspace() or text[-1].isspace()):
            t.set(XML_SPACE, "preserve")

    def _add_table(
        sec_headers: Sequence[str],
        sec_rows: Sequence[Sequence[str]],
        sec_widths: Optional[Sequence[int]] = None,
    ):
        if not sec_headers or not sec_rows:
            return
        tbl = ET.SubElement(body, f"{W}tbl")
        tblPr = ET.SubElement(tbl, f"{W}tblPr")
        ET.SubElement(tblPr, f"{W}tblW", {f"{W}w": "14570", f"{W}type": "dxa"})
        ET.SubElement(tblPr, f"{W}tblLayout", {f"{W}type": "fixed"})
        borders = ET.SubElement(tblPr, f"{W}tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            ET.SubElement(borders, f"{W}{side}", {f"{W}val": "single", f"{W}sz": "4", f"{W}space": "0", f"{W}color": "BFBFBF"})

        if sec_widths and len(sec_widths) == len(sec_headers):
            widths = [int(w) for w in sec_widths]
        else:
            default_w = max(1000, 14570 // len(sec_headers))
            widths = [default_w] * len(sec_headers)

        grid = ET.SubElement(tbl, f"{W}tblGrid")
        for w_val in widths:
            ET.SubElement(grid, f"{W}gridCol", {f"{W}w": str(w_val)})

        font_sz = "17" if len(sec_headers) >= 6 else "19"

        def _add_row(vals: Sequence[str], is_hdr: bool = False):
            tr = ET.SubElement(tbl, f"{W}tr")
            for c_idx, v in enumerate(vals):
                w_cur = widths[c_idx] if c_idx < len(widths) else widths[-1]
                tc = ET.SubElement(tr, f"{W}tc")
                tcPr = ET.SubElement(tc, f"{W}tcPr")
                ET.SubElement(tcPr, f"{W}tcW", {f"{W}w": str(w_cur), f"{W}type": "dxa"})
                if is_hdr:
                    ET.SubElement(tcPr, f"{W}shd", {f"{W}val": "clear", f"{W}color": "auto", f"{W}fill": "D9E2F3"})
                lines_cell = str(v or "").splitlines() or [""]
                for ln_txt in lines_cell:
                    p = ET.SubElement(tc, f"{W}p")
                    r = ET.SubElement(p, f"{W}r")
                    rPr = ET.SubElement(r, f"{W}rPr")
                    ET.SubElement(rPr, f"{W}rFonts", {f"{W}ascii": "Georgia", f"{W}eastAsia": "宋体", f"{W}hAnsi": "Georgia"})
                    if is_hdr:
                        ET.SubElement(rPr, f"{W}b")
                    ET.SubElement(rPr, f"{W}sz", {f"{W}val": font_sz})
                    ET.SubElement(rPr, f"{W}szCs", {f"{W}val": font_sz})
                    t = ET.SubElement(r, f"{W}t")
                    t.text = ln_txt
                    if ln_txt and (ln_txt[0].isspace() or ln_txt[-1].isspace()):
                        t.set(XML_SPACE, "preserve")

        _add_row(sec_headers, is_hdr=True)
        for row_vals in sec_rows:
            _add_row(row_vals, is_hdr=False)

    _add_p(title, bold=True, sz="28", color="1F4E79")
    for line in summary_lines:
        _add_p(line, bold=False, sz="20")

    _add_table(headers, rows_data, col_widths)

    if extra_sections:
        for sec_title, sec_hdrs, sec_rows, sec_ws in extra_sections:
            if not sec_rows:
                continue
            _add_p("", bold=False, sz="16")
            _add_p(sec_title, bold=True, sz="24", color="1F4E79")
            _add_table(sec_hdrs, sec_rows, sec_ws)

    sect = ET.SubElement(body, f"{W}sectPr")
    ET.SubElement(sect, f"{W}pgSz", {f"{W}w": "16838", f"{W}h": "11906", f"{W}orient": "landscape"})
    ET.SubElement(
        sect,
        f"{W}pgMar",
        {f"{W}top": "1134", f"{W}right": "1134", f"{W}bottom": "1247", f"{W}left": "1134", f"{W}header": "720", f"{W}footer": "720", f"{W}gutter": "0"},
    )

    with docx_xml_utils._ET_NS_LOCK:
        ET.register_namespace("w", W_URI)
        ET.register_namespace("r", R_URI)
        doc_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + ET.tostring(doc, encoding="unicode")).encode("utf-8")

    styles_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<w:styles xmlns:w="{W_URI}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'
    ).encode("utf-8")
    ct_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '</Types>'
    ).encode("utf-8")
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>'
    ).encode("utf-8")
    drels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>'
    ).encode("utf-8")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", ct_xml)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("word/document.xml", doc_xml)
        zf.writestr("word/styles.xml", styles_xml)
        zf.writestr("word/_rels/document.xml.rels", drels_xml)
    return out_path


def _merge_revision_docs(
    rev_doc_paths: Sequence[Tuple[str, Path]],
    out_path: Path,
) -> Path:
    """
    Merge multiple revision `.docx` documents into one combined `.docx` preserving all `<w:ins>`/`<w:del>`
    track changes, stripping unmerged chapter comment refs, and re-sequencing `w:id` globally so Word
    never reports duplicate markup IDs or missing comment targets.
    """
    if not rev_doc_paths:
        return out_path
    first_title, first_p = rev_doc_paths[0]
    base_root, _, ns_pairs = docx_xml_utils.load_docx_xml(first_p)
    base_body = base_root.find("w:body", NS)
    sect_pr = base_body.find("w:sectPr", NS)
    if sect_pr is not None:
        base_body.remove(sect_pr)

    def _make_banner_p(text: str) -> ET.Element:
        p = ET.Element(f"{W}p")
        r = ET.SubElement(p, f"{W}r")
        rPr = ET.SubElement(r, f"{W}rPr")
        ET.SubElement(rPr, f"{W}b")
        ET.SubElement(rPr, f"{W}color", {f"{W}val": "1F4E79"})
        ET.SubElement(rPr, f"{W}sz", {f"{W}val": "26"})
        ET.SubElement(rPr, f"{W}szCs", {f"{W}val": "26"})
        t = ET.SubElement(r, f"{W}t")
        t.text = f"═══ {text} ═══"
        return p

    base_body.insert(0, _make_banner_p(first_title))

    for title, doc_p in rev_doc_paths[1:]:
        sub_root, _, _ = docx_xml_utils.load_docx_xml(doc_p)
        sub_body = sub_root.find("w:body", NS)
        if sub_body is None:
            continue
        pb_p = ET.Element(f"{W}p")
        pb_r = ET.SubElement(pb_p, f"{W}r")
        ET.SubElement(pb_r, f"{W}br", {f"{W}type": "page"})
        base_body.append(pb_p)
        base_body.append(_make_banner_p(title))
        for ch in list(sub_body):
            if ch.tag.split("}")[-1] != "sectPr":
                base_body.append(copy.deepcopy(ch))

    # Strip library comment anchors on the combined preview copy & renumber all revision w:id uniquely
    for tc in base_body.findall(".//w:tc", NS):
        strip_comments_in_cell(tc)
    next_rev_id = 1
    for el in base_body.iter():
        if el.tag.split("}")[-1] in ("ins", "del"):
            el.set(f"{W}id", str(next_rev_id))
            next_rev_id += 1

    if sect_pr is not None:
        base_body.append(sect_pr)
    docx_xml_utils.save_docx_xml(first_p, out_path, base_root, ns_pairs, comments_root=None)
    return out_path


def _extract_rows_with_titles(doc_root: ET.Element) -> List[Tuple[str, ET.Element]]:
    """
    Walk `w:body` in document order, tracking preceding `<w:p>` chapter titles before each `<w:tbl>`.
    Returns `[(chapter_title_hint, tr_element), ...]` for all rows with >= 3 cells.
    """
    body = doc_root.find("w:body", NS)
    if body is None:
        return [("", tr) for tr in doc_root.findall(".//w:tbl/w:tr", NS) if len(tr.findall("w:tc", NS)) >= 3]

    out: List[Tuple[str, ET.Element]] = []
    cur_title = ""
    for ch in body:
        loc = ch.tag.split("}")[-1]
        if loc == "p":
            ptxt = "".join(t.text or "" for t in ch.findall(".//w:t", NS)).strip()
            if ptxt:
                # Strip common workflow suffixes such as "通读", "改点" while keeping actual chapter title
                cleaned_t = re.sub(r"(?:\s*[-—_]*\s*(?:通读|改点|初筛|定稿))+$", "", ptxt).strip()
                cur_title = cleaned_t or ptxt
        elif loc == "tbl":
            for tr in ch.findall("w:tr", NS):
                if len(tr.findall("w:tc", NS)) >= 3:
                    out.append((cur_title, tr))
    return out


# ---------------------------------------------------------------------------
# Main Scenario 3A Orchestrator: `run_library_ingestion_pipeline`
# ---------------------------------------------------------------------------

def run_library_ingestion_pipeline(
    sent_path: Union[str, Path],
    returned_path: Union[str, Path],
    library_dir: Union[str, Path],
    output_dir: Union[str, Path],
    conflict_policy: str = "keep_library",  # "keep_library" | "skip_paragraph"
    revision_merge_mode: str = "single",    # "single" | "batch" | "per_article"
    revision_batch_size: int = 10,
    check_only: bool = False,
    proofread_sources: Optional[Sequence[str]] = None,
    col_config: Optional[ColumnMappingConfig] = None,
    oauth: Optional[GoogleOAuthManager] = None,
    api_key: Union[str, List[str], ApiKeyPool] = "",
    model: str = "gemini-3.5-flash-lite",
    api_mode: str = "auto",
    date_str: str = "",
    log_fn: Optional[Callable[[str], None]] = None,
    progress_fn: Optional[Callable[[float, str], None]] = None,
) -> dict:
    """
    Execute Scenario 3A (`改点入库`) end-to-end:
      - Step 1A: Reject revisions on `发稿`, compare with `library_dir` (normalizing artificial Col 3 `\\n` breaks),
                 automatically ACCEPT normal Cyan (`RV`) English diffs, and output ONLY remaining non-cyan/foreign diffs.
      - Step 1B: Diff `true_old_baseline` (library old English on Cyan + rejected `发稿` on non-Cyan) vs `回稿`
                 so both UP and Cyan RV show track changes, and detect structural anomalies with Word comments.
      - Step 2: Clean formatting artifacts on `回稿` + validate & strictly exact-match replace multi-link proofread tables
                (reporting full 5 columns whenever `改动列!=定稿列`, `改动列==改前列`, not found in `回稿`, or clause/meaning anomaly).
      - Step 3: Extract matched library chapters to `改前/`, perform surgical write-back to `改后/`
                (preserving original filenames, dates, inline styles, & overlaying Col 1 Chinese shading),
                and generate `修订稿/` (shown on `改后`).
    """
    lib_p = Path(library_dir).resolve()
    out_d = Path(output_dir).resolve()
    if not lib_p.is_dir():
        raise FileNotFoundError(f"找不到语种总库目录：{lib_p}")
    out_d.mkdir(parents=True, exist_ok=True)
    eff_date = docx_xml_utils.sanitize_filename_part(
        (date_str or "").strip() or datetime.now().strftime("%Y%m%d"),
        fallback=datetime.now().strftime("%Y%m%d"),
        max_len=16,
    )

    if progress_fn:
        progress_fn(3.0, "正在配对发稿与回稿并加载语种总库索引...")
    pairs, unpaired_s, unpaired_r = pair_sent_and_returned_files(sent_path, returned_path)
    if log_fn:
        log_fn(f"📂 自动配对完成：成功配对 {len(pairs)} 组发稿/回稿")
        for sf, rf in pairs:
            log_fn(f"   • 发稿：{sf.name}  ↔  回稿：{rf.name}")
        if unpaired_s:
            log_fn(f"⚠️ 未配对发稿 ({len(unpaired_s)})：{', '.join(f.name for f in unpaired_s)}")
        if unpaired_r:
            log_fn(f"⚠️ 未配对回稿 ({len(unpaired_r)})：{', '.join(f.name for f in unpaired_r)}")

    if not pairs:
        raise ValueError("未能配对任何发稿与回稿文档，请检查文件路径或命名。")

    idx_stats = ensure_library_index(str(lib_p), log_fn=log_fn)
    lib_index = library_index.get_library_index(lib_p)
    if log_fn:
        log_fn(f"📚 总库索引就绪：共 {idx_stats.get('total_files', 0)} 篇文档（{idx_stats.get('total_rows', 0)} 段）")

    # Optional: load & strictly validate multiple Google Sheets links / local `.xlsx/.csv` proofread tables
    clean_pr_sources = extract_proofread_sources_list(proofread_sources)
    pending_pr_items: List[dict] = []
    proofread_5col_rows: List[List[str]] = []

    for s_idx, s_clean in enumerate(clean_pr_sources):
        try:
            parsed = load_proofread_table(s_clean, col_config=col_config, oauth=oauth)
            rep_items, blocked_5col = evaluate_proofread_sheet_rows(
                parsed=parsed,
                chapter_label="",
                enable_ai=bool(api_key),
                api_key=api_key,
                model=model,
                api_mode=api_mode,
            )
            for it in rep_items:
                it["matched"] = False
                pending_pr_items.append(it)
            proofread_5col_rows.extend(blocked_5col)
            if log_fn:
                log_fn(
                    f"📊 已读取通读表 [{s_idx + 1}/{len(clean_pr_sources)}]（{parsed.sheet_name}）："
                    f"共 {len(parsed.rows)} 行，有效待替换 {len(rep_items)} 条，列校验拦截 {len(blocked_5col)} 条"
                )
        except Exception as ex:
            err_pr = ProofreadRow(row_number=0, cn="", en="", target="", changes="", final="", source_label=s_clean)
            proofread_5col_rows.append(make_proofread_5col_row(s_clean[:60], f"读取通读表失败：{ex}", err_pr))
            if log_fn:
                log_fn(f"⚠️ 读取通读表失败：{s_clean} -> {ex}")

    lib_doc_cache: Dict[Path, Tuple[ET.Element, List[ET.Element]]] = {}

    def _get_lib_rows(lib_file: Path) -> List[ET.Element]:
        if lib_file not in lib_doc_cache:
            lr, _, _ = docx_xml_utils.load_docx_xml(lib_file)
            lib_doc_cache[lib_file] = (lr, lr.findall(".//w:tbl/w:tr", NS))
        return lib_doc_cache[lib_file][1]

    # Collect per-library-file updates: {lib_file: [update_spec, ...]}
    lib_updates: Dict[Path, List[dict]] = {}
    step1a_mismatches: List[dict] = []
    step1b_anomalies: List[dict] = []
    saved_reports: List[str] = []
    proofread_applied_total = 0

    for p_idx, (sf, rf) in enumerate(pairs):
        pct_base = 8.0 + (p_idx / max(1, len(pairs))) * 45.0
        if progress_fn:
            progress_fn(pct_base, f"正在验收检查 [{p_idx + 1}/{len(pairs)}]：{rf.name}")
        if log_fn:
            log_fn(f"\n🔍 === [{p_idx + 1}/{len(pairs)}] 第一步验收检查：{rf.name} ===")

        s_root, _, s_ns = docx_xml_utils.load_docx_xml(sf)
        r_root, _, r_ns = docx_xml_utils.load_docx_xml(rf)
        r_comments_map = docx_xml_utils.load_docx_comments_map(rf)

        s_titled_rows = _extract_rows_with_titles(s_root)
        r_rows = [tr for tr in r_root.findall(".//w:tbl/w:tr", NS) if len(tr.findall("w:tc", NS)) >= 3]

        # Build lookup of returned rows by Col 1 Chinese key (with sequential fallback)
        r_by_cn: Dict[str, List[ET.Element]] = {}
        for r_tr in r_rows:
            r_tcs = r_tr.findall("w:tc", NS)
            r_c1 = copy.deepcopy(r_tcs[0])
            accept_all_revisions_in_cell_safe(r_c1)
            k = _norm_cn_key(docx_xml_utils.get_cell_plain_text(r_c1))
            if k:
                r_by_cn.setdefault(k, []).append(r_tr)

        # Prepare Step 1B annotated diff document (cloned from `发稿`, with Col 1 revisions rejected cleanly)
        qa_root = copy.deepcopy(s_root)
        for _tc_qa in qa_root.findall(".//w:tbl/w:tr/w:tc", NS):
            reject_all_revisions_in_cell(_tc_qa)
            strip_comments_in_cell(_tc_qa)
        qa_rows = [tr for tr in qa_root.findall(".//w:tbl/w:tr", NS) if len(tr.findall("w:tc", NS)) >= 3]
        qa_comments_root = docx_xml_utils.new_comments_root()
        qa_id_alloc = docx_xml_utils.CommentIdAllocator(doc_root=qa_root, comments_root=qa_comments_root)
        rev_counter = [100]

        # Prepare Step 1A mismatch diff document (FIRST reject ALL `发稿` revisions across all cells!)
        mism_root = copy.deepcopy(s_root)
        for _tc_m in mism_root.findall(".//w:tbl/w:tr/w:tc", NS):
            reject_all_revisions_in_cell(_tc_m)
            strip_comments_in_cell(_tc_m)
        mism_rows = [tr for tr in mism_root.findall(".//w:tbl/w:tr", NS) if len(tr.findall("w:tc", NS)) >= 3]
        mism_row_indices_kept: set = set()
        mism_has_any = False

        # Pre-pass: match all rows of `(sf, rf)` so exact proofread items can match by CN first, then exact Col 3 fallback
        pair_row_contexts: List[dict] = []
        for r_i, (row_title_hint, s_tr) in enumerate(s_titled_rows):
            s_tcs = s_tr.findall("w:tc", NS)
            s_c1_copy = copy.deepcopy(s_tcs[0])
            reject_all_revisions_in_cell(s_c1_copy)
            cn_text = docx_xml_utils.get_cell_plain_text(s_c1_copy).strip()
            cn_key = _norm_cn_key(cn_text)
            if not cn_key:
                continue

            r_tr = None
            if cn_key in r_by_cn and r_by_cn[cn_key]:
                r_tr = r_by_cn[cn_key].pop(0)
            elif r_i < len(r_rows):
                r_tr = r_rows[r_i]
            if r_tr is None:
                continue
            r_tcs = r_tr.findall("w:tc", NS)

            r_c1_clean = copy.deepcopy(r_tcs[0])
            accept_all_revisions_in_cell_safe(r_c1_clean)
            c1_cell_fill, c1_run_spans = extract_col1_shading_spans(r_c1_clean)

            r_c2_clean_tc = copy.deepcopy(r_tcs[1])
            clean_col23_highlights_and_artifacts(r_c2_clean_tc)
            r_c2_clean = docx_xml_utils.get_cell_plain_text(r_c2_clean_tc)

            r_c3_comments = get_cell_comment_texts(r_tcs[2], r_comments_map)
            r_c3_clean_tc = copy.deepcopy(r_tcs[2])
            clean_col23_highlights_and_artifacts(r_c3_clean_tc)
            r_c3_clean = docx_xml_utils.get_cell_plain_text(r_c3_clean_tc).replace("\u00A0", " ")

            pair_row_contexts.append({
                "r_i": r_i,
                "row_title_hint": row_title_hint,
                "s_tcs": s_tcs,
                "s_c1_copy": s_c1_copy,
                "cn_text": cn_text,
                "cn_key": cn_key,
                "c1_cell_fill": c1_cell_fill,
                "c1_run_spans": c1_run_spans,
                "r_c2_clean": r_c2_clean,
                "r_c3_clean": r_c3_clean,
                "r_c3_comments": r_c3_comments,
            })

        # Apply strict exact-match proofread replacements (`改前列` must match `回稿` 100% verbatim!)
        if pending_pr_items and pair_row_contexts:
            # Pass 1: CN key overlap + 100% exact target match in `r_c3_clean`
            for item in pending_pr_items:
                if item["matched"]:
                    continue
                old_exact = item["old_text"]
                new_exact = item["new_text"]
                pr_cn_k = _norm_cn_key(item["row"].cn)
                if not pr_cn_k:
                    continue
                for ctx in pair_row_contexts:
                    if (pr_cn_k in ctx["cn_key"] or ctx["cn_key"] in pr_cn_k) and (old_exact in ctx["r_c3_clean"]):
                        ctx["r_c3_clean"] = ctx["r_c3_clean"].replace(old_exact, new_exact, 1)
                        item["matched"] = True
                        proofread_applied_total += 1
                        break

            # Pass 2: Fallback 100% exact target match across `pair_row_contexts` when CN in sheet was blank/partial
            for item in pending_pr_items:
                if item["matched"]:
                    continue
                old_exact = item["old_text"]
                new_exact = item["new_text"]
                for ctx in pair_row_contexts:
                    if old_exact in ctx["r_c3_clean"]:
                        ctx["r_c3_clean"] = ctx["r_c3_clean"].replace(old_exact, new_exact, 1)
                        item["matched"] = True
                        proofread_applied_total += 1
                        break

        for ctx in pair_row_contexts:
            r_i = ctx["r_i"]
            row_title_hint = ctx["row_title_hint"]
            s_tcs = ctx["s_tcs"]
            s_c1_copy = ctx["s_c1_copy"]
            cn_text = ctx["cn_text"]
            c1_cell_fill = ctx["c1_cell_fill"]
            c1_run_spans = ctx["c1_run_spans"]
            r_c2_clean = ctx["r_c2_clean"]
            r_c3_clean = ctx["r_c3_clean"]
            r_c3_comments = ctx["r_c3_comments"]

            # Extract rejected `发稿` Col 2 & Col 3
            s_c2_rej_full_tc = copy.deepcopy(s_tcs[1])
            reject_all_revisions_in_cell(s_c2_rej_full_tc)
            s_c2_rej_full = docx_xml_utils.get_cell_plain_text(s_c2_rej_full_tc)

            s_c3_rej_tc = copy.deepcopy(s_tcs[2])
            reject_all_revisions_in_cell(s_c3_rej_tc)
            s_c3_rej = docx_xml_utils.get_cell_plain_text(s_c3_rej_tc)

            # Locate in 3-language library via 5-stage library index search (with title disambiguation + fallback)
            matched_row, _, _ = step1_extract_and_diff.search_library_row_5stage(
                cur_title=row_title_hint,
                tc1=s_c1_copy,
                tc2=s_c2_rej_full_tc,
                inp_cn=cn_text,
                inp_en=s_c2_rej_full,
                lib_index=lib_index,
            )
            if matched_row is None and row_title_hint:
                matched_row, _, _ = step1_extract_and_diff.search_library_row_5stage(
                    cur_title="",
                    tc1=s_c1_copy,
                    tc2=s_c2_rej_full_tc,
                    inp_cn=cn_text,
                    inp_en=s_c2_rej_full,
                    lib_index=lib_index,
                )

            lib_file_p: Optional[Path] = None
            lib_row_idx: int = -1
            lib_c2_text = ""
            lib_c3_text = ""
            is_mismatch = False
            mismatch_reason = ""
            c2_is_excerpt = False
            c3_is_excerpt = False
            true_old_c2 = s_c2_rej_full
            s_c3_aligned = s_c3_rej

            if matched_row is not None:
                lib_file_p = lib_index.resolve_file_path(matched_row["rel_path"]).resolve()
                lib_row_idx = int(matched_row["row_idx"])
                lib_trs = _get_lib_rows(lib_file_p)
                if 0 <= lib_row_idx < len(lib_trs):
                    ltcs = lib_trs[lib_row_idx].findall("w:tc", NS)
                    if len(ltcs) >= 3:
                        lt2 = copy.deepcopy(ltcs[1])
                        lt3 = copy.deepcopy(ltcs[2])
                        accept_all_revisions_in_cell_safe(lt2)
                        accept_all_revisions_in_cell_safe(lt3)
                        lib_c2_text = docx_xml_utils.get_cell_plain_text(lt2)
                        lib_c3_text = docx_xml_utils.get_cell_plain_text(lt3)

                        # Step 1A Col 2: reject `发稿` revisions -> diff vs `总库` -> accept Cyan (`RV`) revisions -> check remaining!
                        c2_info = analyze_sent_col2_vs_library(s_tcs[1], lib_c2_text)
                        en_diff = bool(c2_info["has_remaining_diff"])
                        c2_is_excerpt = bool(c2_info["is_excerpt"])
                        step1a_c2_target = str(c2_info["step1a_target"])
                        true_old_c2 = str(c2_info["true_old_baseline"])

                        # Step 1A Col 3: treat isolated RV blank-line splits (`\n\n`) identically to single spaces!
                        norm_s3_flat = _norm_compare_flat_ws(s_c3_rej)
                        norm_l3_flat = _norm_compare_flat_ws(lib_c3_text)
                        c3_is_excerpt = bool(norm_s3_flat and norm_s3_flat != norm_l3_flat and norm_s3_flat in norm_l3_flat)
                        f_diff = (norm_s3_flat != norm_l3_flat) and not c3_is_excerpt
                        s_c3_aligned = _align_sent_text_breaks_to_library(s_c3_rej, lib_c3_text)

                        # If library Col 3 is single-paragraph and `发稿` only had artificial RV `\n\n` breaks,
                        # collapse any leftover `\n` in `r_c3_clean` back to space so `改后/` matches library paragraphing
                        if "\n" not in lib_c3_text.strip() and "\n" in s_c3_rej.strip() and not f_diff and "\n" in r_c3_clean.strip():
                            r_c3_clean = re.sub(r"\s*\n+\s*", " ", r_c3_clean).strip()

                        if en_diff or f_diff:
                            is_mismatch = True
                            mism_has_any = True
                            mism_row_indices_kept.add(r_i)
                            reasons = []
                            if en_diff:
                                reasons.append("英文(接受青色修订后剩余部分)与总库不一致")
                            if f_diff:
                                reasons.append("外文与总库不一致")
                            mismatch_reason = "；".join(reasons)
                            step1a_mismatches.append({
                                "sent_file": sf.name,
                                "chapter": lib_file_p.name,
                                "row_idx": lib_row_idx + 1,
                                "cn_preview": cn_text[:50],
                                "reason": mismatch_reason,
                            })
                            if r_i < len(mism_rows):
                                m_tcs = mism_rows[r_i].findall("w:tc", NS)
                                if en_diff:
                                    render_diff_track_changes_into_cell(
                                        m_tcs[1], s_c2_rej_full, step1a_c2_target, author="总库现状", rev_id_counter=rev_counter
                                    )
                                if f_diff:
                                    render_diff_track_changes_into_cell(
                                        m_tcs[2], s_c3_aligned, lib_c3_text, author="总库现状", rev_id_counter=rev_counter
                                    )

            # Step 1B: Diff `true_old_c2` (including library old English for Cyan RV!) vs `回稿` & check anomalies
            if r_i < len(qa_rows):
                qa_tcs = qa_rows[r_i].findall("w:tc", NS)
                render_diff_track_changes_into_cell(
                    qa_tcs[1], true_old_c2, r_c2_clean, author="英外验收", rev_id_counter=rev_counter
                )
                render_diff_track_changes_into_cell(
                    qa_tcs[2], s_c3_aligned, r_c3_clean, author="英外验收", rev_id_counter=rev_counter
                )
                anoms = _detect_col3_anomalies(
                    sent_tc2=s_tcs[1],
                    sent_col2_rej=true_old_c2,
                    ret_col2_clean=r_c2_clean,
                    sent_col3_clean=s_c3_aligned,
                    ret_col3_clean=r_c3_clean,
                    ret_col3_comments=r_c3_comments,
                )
                if anoms:
                    comment_msg = "\n".join(anoms)
                    add_comment_to_cell(qa_tcs[2], comment_msg, qa_comments_root, qa_id_alloc)
                    step1b_anomalies.append({
                        "file": rf.name,
                        "chapter": lib_file_p.name if lib_file_p else "未定位篇目",
                        "row_idx": r_i + 1,
                        "cn_preview": cn_text[:45],
                        "issues": " | ".join(anoms),
                    })

            # Queue Step 3 write-back update if matched in library
            if lib_file_p is not None and lib_row_idx >= 0:
                if is_mismatch and conflict_policy == "skip_paragraph":
                    continue
                lib_updates.setdefault(lib_file_p, []).append({
                    "row_idx": lib_row_idx,
                    "sent_c2_base": true_old_c2,
                    "sent_c3_base": s_c3_aligned,
                    "ret_c2_clean": r_c2_clean,
                    "ret_c3_clean": r_c3_clean,
                    "c2_is_excerpt": c2_is_excerpt,
                    "c3_is_excerpt": c3_is_excerpt,
                    "c1_cell_fill": c1_cell_fill,
                    "c1_run_spans": c1_run_spans,
                    "had_conflict": is_mismatch,
                    "conflict_policy": conflict_policy,
                })

        # Save Step 1B annotated diff doc
        safe_stem = docx_xml_utils.sanitize_filename_part(rf.stem, fallback="回稿", max_len=100)
        qa_out = docx_xml_utils.safe_join_output_file(out_d, f"步1B_英外改动痕迹验收稿_{safe_stem}_{eff_date}.docx")
        final_qa_cmts = qa_comments_root if len(qa_comments_root.findall(f"{W}comment")) > 0 else None
        docx_xml_utils.save_docx_xml(sf, qa_out, qa_root, s_ns, comments_root=final_qa_cmts)
        saved_reports.append(str(qa_out))

        # Save Step 1A mismatch diff doc ONLY if remaining mismatches exist (removing clean rows so only mismatched rows remain!)
        if mism_has_any:
            row_idx_map = {id(tr): idx_r for idx_r, tr in enumerate(mism_rows)}
            for tbl_m in list(mism_root.findall(".//w:tbl", NS)):
                for tr_m in list(tbl_m.findall("w:tr", NS)):
                    idx_m = row_idx_map.get(id(tr_m))
                    if idx_m is not None and idx_m not in mism_row_indices_kept:
                        tbl_m.remove(tr_m)
                if len(tbl_m.findall("w:tr", NS)) == 0:
                    m_body = mism_root.find("w:body", NS)
                    if m_body is not None and tbl_m in list(m_body):
                        m_body.remove(tbl_m)
            mism_out = docx_xml_utils.safe_join_output_file(out_d, f"步1A_发稿与总库不一致核对稿_{safe_stem}_{eff_date}.docx")
            docx_xml_utils.save_docx_xml(sf, mism_out, mism_root, s_ns)
            saved_reports.append(str(mism_out))

    # Finalize Proofread 5-Column Audit Ledger:
    #   - Matched items with clause/meaning warnings -> report all 5 columns for review
    #   - Unmatched items (`找不到完全和回稿一致的`) -> DO NOT replace, report all 5 columns!
    pr_unmatched_cnt = 0
    pr_warn_cnt = 0
    for item in pending_pr_items:
        r_obj = item["row"]
        warns = item["warnings"]
        if not item["matched"]:
            pr_unmatched_cnt += 1
            reason_txt = "未替换：回稿中未找到与改前列完全一致的内容"
            if warns:
                reason_txt += "；另发现：" + "；".join(warns)
            proofread_5col_rows.append(make_proofread_5col_row(item["source_pos"], reason_txt, r_obj))
        elif warns:
            pr_warn_cnt += 1
            proofread_5col_rows.append(
                make_proofread_5col_row(item["source_pos"], "已替换（请人工复核异常）：" + "；".join(warns), r_obj)
            )

    if clean_pr_sources and log_fn:
        log_fn(
            f"📊 通读定稿表核销汇总：成功精确替换 {proofread_applied_total} 条，"
            f"未找到完全一致/列不一致未替换 {len(proofread_5col_rows) - pr_warn_cnt} 条，改后异常提醒 {pr_warn_cnt} 条。"
        )

    # Write combined Step 1A/1B + Proofread 5-Column summary report
    rep_rows = []
    for m in step1a_mismatches:
        rep_rows.append(["Step 1A 发稿vs总库", m["chapter"], str(m["row_idx"]), m["cn_preview"], m["reason"]])
    for a in step1b_anomalies:
        rep_rows.append(["Step 1B 回稿验收", a["chapter"], str(a["row_idx"]), a["cn_preview"], a["issues"]])

    extra_secs = []
    if proofread_5col_rows:
        extra_secs.append((
            f"二、通读定稿表核查与未替换/异常明细（共 {len(proofread_5col_rows)} 条，完整列出 5 列）",
            PROOFREAD_5COL_HEADERS,
            proofread_5col_rows,
            PROOFREAD_5COL_WIDTHS,
        ))

    summary_rep_path = docx_xml_utils.safe_join_output_file(out_d, f"步1_改点入库验收报告_{eff_date}.docx")
    _write_simple_report_docx(
        out_path=summary_rep_path,
        title=f"改点入库验收检查报告（{eff_date}）",
        summary_lines=[
            f"配对文档数：{len(pairs)} 组 | 定位涉及总库篇目：{len(lib_updates)} 篇",
            f"Step 1A 发稿与总库不一致段落（已拒绝发稿修订并接受青色正常差异）：{len(step1a_mismatches)} 处（策略：{'保留总库改动' if conflict_policy == 'keep_library' else '跳过不一致段落'}）",
            f"Step 1B 回稿结构/漏改/衔接异常：{len(step1b_anomalies)} 处（已在验收痕迹稿对应位置添加 Word 批注）",
            f"通读定稿表：共读取 {len(clean_pr_sources)} 个表格，成功精确替换 {proofread_applied_total} 条，需人工核查/未替换汇报 {len(proofread_5col_rows)} 条",
        ],
        headers=["检查阶段", "涉及篇目", "段落行", "中文预览", "检查结果 / 异常说明"],
        rows_data=rep_rows or [["全部通过", "-", "-", "-", "未发现总库冲突或回稿结构异常，可放心入库。"]],
        col_widths=[2100, 3300, 1100, 3270, 4800],
        extra_sections=extra_secs,
    )
    saved_reports.append(str(summary_rep_path))

    if log_fn:
        log_fn(
            f"📋 验收报告已生成：总库不一致 {len(step1a_mismatches)} 处，回稿异常 {len(step1b_anomalies)} 处，"
            f"通读表汇报项 {len(proofread_5col_rows)} 条。"
        )

    if check_only:
        if progress_fn:
            progress_fn(100.0, "仅验收检查完成（未执行入库回写）。")
        return {
            "saved_reports": saved_reports,
            "before_dir": "",
            "after_dir": "",
            "revision_files": [],
            "mismatches_count": len(step1a_mismatches),
            "anomalies_count": len(step1b_anomalies),
            "updated_chapters": 0,
            "updated_rows": 0,
            "proofread_applied": proofread_applied_total,
            "proofread_report_count": len(proofread_5col_rows),
        }

    # ------------------------------------------------------------------
    # Step 3: Zero-mutation extract to `改前/`, surgical write-back to `改后/`,
    #         and generate `修订稿/` (shown on `改后`)
    # ------------------------------------------------------------------
    before_dir = out_d / "改前"
    after_dir = out_d / "改后"
    rev_dir = out_d / "修订稿"
    before_dir.mkdir(parents=True, exist_ok=True)
    after_dir.mkdir(parents=True, exist_ok=True)
    rev_dir.mkdir(parents=True, exist_ok=True)

    updated_chapters = 0
    updated_rows_total = 0
    per_chapter_rev_docs: List[Tuple[str, Path]] = []
    lib_files_sorted = sorted(lib_updates.keys(), key=lambda p: p.name.lower())
    global_rev_ctr = [2000]

    def _resolve_updated_cell_text(
        cur_cell_text: str,
        sent_base: str,
        ret_clean: str,
        is_excerpt: bool,
        had_conflict: bool,
        pol: str,
    ) -> str:
        if is_excerpt:
            # Substitute only the quoted excerpt inside `cur_cell_text` (preserving unquoted stanzas/paragraphs)
            s_trim = (sent_base or "").strip()
            r_trim = (ret_clean or "").strip()
            if s_trim and s_trim in cur_cell_text:
                return cur_cell_text.replace(s_trim, r_trim, 1)
            loc = docx_xml_utils.find_fuzzy_substring_span(cur_cell_text, s_trim) if s_trim else None
            if loc is not None:
                return cur_cell_text[: loc[0]] + r_trim + cur_cell_text[loc[1] :]
        if had_conflict and pol == "keep_library":
            return merge_keep_library_text(sent_base, cur_cell_text, ret_clean)
        return ret_clean

    for c_idx, lib_file in enumerate(lib_files_sorted):
        if progress_fn:
            progress_fn(
                55.0 + (c_idx / max(1, len(lib_files_sorted))) * 38.0,
                f"正在回写篇目 [{c_idx + 1}/{len(lib_files_sorted)}]：{lib_file.name}",
            )
        # 1. Losslessly slimmed copy to `改前/` (NEVER touch `lib_file`!)
        before_copy = docx_xml_utils.safe_join_output_file(before_dir, lib_file.name)
        docx_xml_utils.slim_docx_file(lib_file, before_copy)

        # 2. Load from `before_copy` to produce `改后/` (exact original filename & date!)
        after_copy = docx_xml_utils.safe_join_output_file(after_dir, lib_file.name)
        root_after, _, ns_after = docx_xml_utils.load_docx_xml(before_copy)
        cmts_after = docx_xml_utils.load_docx_comments_xml(before_copy)
        trs_after = root_after.findall(".//w:tbl/w:tr", NS)

        # Also prepare a revision copy where track changes are shown on `改后`
        root_rev = copy.deepcopy(root_after)
        trs_rev = root_rev.findall(".//w:tbl/w:tr", NS)
        orig_row_texts: Dict[int, Tuple[str, str]] = {}
        touched_row_indices: List[int] = []
        chapter_changed = False

        for upd in lib_updates[lib_file]:
            r_idx = int(upd["row_idx"])
            if not (0 <= r_idx < len(trs_after)):
                continue
            tcs_a = trs_after[r_idx].findall("w:tc", NS)
            tcs_r = trs_rev[r_idx].findall("w:tc", NS)
            if len(tcs_a) < 3:
                continue

            if r_idx not in orig_row_texts:
                c2_init = copy.deepcopy(tcs_a[1])
                c3_init = copy.deepcopy(tcs_a[2])
                accept_all_revisions_in_cell_safe(c2_init)
                accept_all_revisions_in_cell_safe(c3_init)
                orig_row_texts[r_idx] = (
                    docx_xml_utils.get_cell_plain_text(c2_init),
                    docx_xml_utils.get_cell_plain_text(c3_init),
                )

            cur_c2 = docx_xml_utils.get_cell_plain_text(tcs_a[1])
            cur_c3 = docx_xml_utils.get_cell_plain_text(tcs_a[2])
            new_c2 = _resolve_updated_cell_text(
                cur_c2,
                upd["sent_c2_base"],
                upd["ret_c2_clean"],
                bool(upd.get("c2_is_excerpt")),
                bool(upd.get("had_conflict")),
                str(upd.get("conflict_policy") or "keep_library"),
            )
            new_c3 = _resolve_updated_cell_text(
                cur_c3,
                upd["sent_c3_base"],
                upd["ret_c3_clean"],
                bool(upd.get("c3_is_excerpt")),
                bool(upd.get("had_conflict")),
                str(upd.get("conflict_policy") or "keep_library"),
            )

            ch2 = surgical_update_cell_text(tcs_a[1], new_c2)
            ch3 = surgical_update_cell_text(tcs_a[2], new_c3)
            shd_cnt = apply_col1_shading_overlay(tcs_a[0], upd["c1_cell_fill"], upd["c1_run_spans"])

            if ch2 or ch3 or shd_cnt > 0:
                chapter_changed = True
                if r_idx not in touched_row_indices:
                    touched_row_indices.append(r_idx)
                    updated_rows_total += 1
                if shd_cnt > 0:
                    apply_col1_shading_overlay(tcs_r[0], upd["c1_cell_fill"], upd["c1_run_spans"])

        for r_idx in touched_row_indices:
            tcs_a = trs_after[r_idx].findall("w:tc", NS)
            tcs_r = trs_rev[r_idx].findall("w:tc", NS)
            orig_c2, orig_c3 = orig_row_texts[r_idx]
            fin_c2 = docx_xml_utils.get_cell_plain_text(tcs_a[1])
            fin_c3 = docx_xml_utils.get_cell_plain_text(tcs_a[2])
            if orig_c2 != fin_c2:
                render_diff_track_changes_into_cell(tcs_r[1], orig_c2, fin_c2, author="改点入库", rev_id_counter=global_rev_ctr)
            if orig_c3 != fin_c3:
                render_diff_track_changes_into_cell(tcs_r[2], orig_c3, fin_c3, author="改点入库", rev_id_counter=global_rev_ctr)

        save_docx_without_track_revisions_flag(before_copy, after_copy, root_after, ns_after, comments_root=cmts_after)
        if chapter_changed:
            updated_chapters += 1
            rev_single_p = docx_xml_utils.safe_join_output_file(rev_dir, f"修订对比_{lib_file.name}")
            docx_xml_utils.save_docx_xml(before_copy, rev_single_p, root_rev, ns_after, comments_root=cmts_after)
            per_chapter_rev_docs.append((lib_file.stem, rev_single_p))

    # Aggregate revision docs according to `revision_merge_mode`
    revision_files: List[str] = []
    if per_chapter_rev_docs:
        mode = (revision_merge_mode or "single").strip().lower()
        if mode == "single":
            merged_p = docx_xml_utils.safe_join_output_file(rev_dir, f"改前后修订对比汇总_{len(per_chapter_rev_docs)}篇_{eff_date}.docx")
            _merge_revision_docs(per_chapter_rev_docs, merged_p)
            revision_files.append(str(merged_p))
            for _, sp in per_chapter_rev_docs:
                sp.unlink(missing_ok=True)
        elif mode == "batch":
            bsz = max(1, int(revision_batch_size or 10))
            for b_start in range(0, len(per_chapter_rev_docs), bsz):
                chunk = per_chapter_rev_docs[b_start : b_start + bsz]
                b_num = (b_start // bsz) + 1
                bp = docx_xml_utils.safe_join_output_file(
                    rev_dir,
                    f"改前后修订对比汇总_第{b_num}批({len(chunk)}篇)_{eff_date}.docx",
                )
                _merge_revision_docs(chunk, bp)
                revision_files.append(str(bp))
            for _, sp in per_chapter_rev_docs:
                sp.unlink(missing_ok=True)
        else:
            revision_files.extend(str(sp) for _, sp in per_chapter_rev_docs)

    if progress_fn:
        progress_fn(100.0, f"改点入库完成：已更新 {updated_chapters} 篇（{updated_rows_total} 段），原总库零改动！")
    if log_fn:
        log_fn(f"\n✅ 改点入库完成！已提取 {len(lib_files_sorted)} 篇至「改前/」，回写 {updated_chapters} 篇至「改后/」，生成修订稿 {len(revision_files)} 份。")

    return {
        "saved_reports": saved_reports,
        "before_dir": str(before_dir),
        "after_dir": str(after_dir),
        "revision_dir": str(rev_dir),
        "revision_files": revision_files,
        "mismatches_count": len(step1a_mismatches),
        "anomalies_count": len(step1b_anomalies),
        "updated_chapters": updated_chapters,
        "updated_rows": updated_rows_total,
        "proofread_applied": proofread_applied_total,
        "proofread_report_count": len(proofread_5col_rows),
    }
