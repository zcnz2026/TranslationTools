import argparse
import difflib
import json
import re
from datetime import datetime
from pathlib import Path
import xml.etree.ElementTree as ET

from docx_xml_utils import (
    NS,
    SKIPPED_ROW_MARKER,
    W,
    _is_top_instruction_paragraph_text,
    accept_all_revisions_in_cell,
    apply_highlights_to_col3_cell,
    balance_trilingual_table_widths,
    compute_col2_up_rv_words,
    count_msword_en_words,
    find_fuzzy_substring_span,
    format_up_rv_count_tag,
    get_cell_plain_text,
    insert_rpr_child_ordered,
    is_two_row_chapter_header,
    load_config,
    load_docx_xml,
    make_run,
    safe_join_output_file,
    sanitize_filename_part,
    save_docx_xml,
    strip_all_highlights_in_cell,
    update_top_highlight_instruction_lines,
)

RV_HIGHLIGHT_VALS = {"cyan", "blue", "darkCyan"}
_CN_SENT_END_RE = re.compile(r"[。？！!?；;\n]+(?:[”’\"'」』）\)]*)")
_CN_CLAUSE_END_RE = re.compile(r"[，,、：:；;。？！!?\n]+(?:[”’\"'」』）\)]*)")


def is_skipped_row(tr: ET.Element) -> bool:
    tcs = tr.findall("w:tc", NS)
    if len(tcs) < 3:
        return False
    return get_cell_plain_text(tcs[2]).strip() == SKIPPED_ROW_MARKER


def remove_skipped_rows_in_body(body: ET.Element) -> int:
    """Remove all skipped rows (`【未在总库匹配到原文，已跳过】` in Col 3) from tables in `body`."""
    removed_rows = 0
    for tbl in body.findall("w:tbl", NS):
        for tr in list(tbl.findall("w:tr", NS)):
            if is_skipped_row(tr):
                tbl.remove(tr)
                removed_rows += 1
    return removed_rows


def normalize_rv_highlights_in_col2(tc2: ET.Element, target_rv_color: str = "cyan"):
    """
    Normalize any RV highlight aliases (e.g. blue / darkCyan set manually in Word) to target_rv_color
    WITHOUT accepting or stripping <w:del> / <w:ins> track changes inside cyan regions.
    """
    for r in tc2.findall(".//w:r", NS):
        rPr = r.find("w:rPr", NS)
        if rPr is None:
            continue
        hl = rPr.find("w:highlight", NS)
        hval = hl.attrib.get(f"{W}val") if hl is not None else None
        if hval in RV_HIGHLIGHT_VALS and hval != target_rv_color:
            insert_rpr_child_ordered(rPr, ET.Element(f"{W}highlight", {f"{W}val": target_rv_color}))


def accept_rv_revisions_in_col2(tc2: ET.Element, target_rv_color: str = "cyan"):
    """
    Accept all track changes (remove w:del, unwrap w:ins) where highlight is RV (cyan/blue),
    leaving yellow (UP) track changes intact.
    """
    for p in tc2.findall("w:p", NS):
        new_children = []
        for ch in list(p):
            local = ch.tag.split("}")[-1]
            if local == "del":
                kept_runs = []
                for r in list(ch):
                    if r.tag.split("}")[-1] != "r":
                        continue
                    hl = r.find("./w:rPr/w:highlight", NS)
                    hval = hl.attrib.get(f"{W}val") if hl is not None else None
                    if hval not in RV_HIGHLIGHT_VALS:
                        kept_runs.append(r)
                if kept_runs:
                    for r in list(ch):
                        ch.remove(r)
                    for r in kept_runs:
                        ch.append(r)
                    new_children.append(ch)
            elif local == "ins":
                pending_non_rv = []
                for r in list(ch):
                    rloc = r.tag.split("}")[-1]
                    if rloc in ("commentRangeStart", "commentRangeEnd"):
                        if pending_non_rv:
                            ins_copy = ET.Element(ch.tag, dict(ch.attrib))
                            for pr in pending_non_rv:
                                ins_copy.append(pr)
                            new_children.append(ins_copy)
                            pending_non_rv = []
                        new_children.append(r)
                        continue
                    if rloc != "r":
                        continue
                    hl = r.find("./w:rPr/w:highlight", NS)
                    hval = hl.attrib.get(f"{W}val") if hl is not None else None
                    if hval in RV_HIGHLIGHT_VALS:
                        if pending_non_rv:
                            ins_copy = ET.Element(ch.tag, dict(ch.attrib))
                            for pr in pending_non_rv:
                                ins_copy.append(pr)
                            new_children.append(ins_copy)
                            pending_non_rv = []
                        rPr = r.find("w:rPr", NS)
                        if rPr is not None:
                            insert_rpr_child_ordered(rPr, ET.Element(f"{W}highlight", {f"{W}val": target_rv_color}))
                        new_children.append(r)
                    else:
                        pending_non_rv.append(r)
                if pending_non_rv:
                    ins_copy = ET.Element(ch.tag, dict(ch.attrib))
                    for pr in pending_non_rv:
                        ins_copy.append(pr)
                    new_children.append(ins_copy)
            else:
                if local == "r":
                    hl = ch.find("./w:rPr/w:highlight", NS)
                    hval = hl.attrib.get(f"{W}val") if hl is not None else None
                    if hval in RV_HIGHLIGHT_VALS and hval != target_rv_color:
                        insert_rpr_child_ordered(ch.find("w:rPr", NS), ET.Element(f"{W}highlight", {f"{W}val": target_rv_color}))
                new_children.append(ch)

        for ch in list(p):
            p.remove(ch)
        for ch in new_children:
            p.append(ch)


def row_has_col2_changes(tr: ET.Element, active_hl_vals: set = None) -> bool:
    """Return True if table row `tr` is non-skipped and has Col 2 track changes (`<w:ins>`/`<w:del>`) or active highlights."""
    if is_skipped_row(tr):
        return False
    tcs = tr.findall("w:tc", NS)
    if len(tcs) < 2:
        return False
    tc2 = tcs[1]
    if tc2.find(".//w:del", NS) is not None or tc2.find(".//w:ins", NS) is not None:
        return True
    if active_hl_vals:
        for hl in tc2.findall(".//w:rPr/w:highlight", NS):
            if hl.attrib.get(f"{W}val") in active_hl_vals:
                return True
    return False


def _is_blank_row(tr: ET.Element, active_hl_vals: set = None) -> bool:
    if row_has_col2_changes(tr, active_hl_vals):
        return False
    for tc in tr.findall("w:tc", NS):
        if get_cell_plain_text(tc).strip():
            return False
    return True


def table_has_col2_changes(tbl_el: ET.Element, active_hl_vals: set = None) -> bool:
    return any(row_has_col2_changes(tr, active_hl_vals) for tr in tbl_el.findall("w:tr", NS))


def filter_tables_in_body(
    body: ET.Element,
    up_color: str,
    rv_color: str,
    keep_changed: bool,
    keep_first_title_row: bool = False,
    no_highlight_colors: bool = False,
) -> tuple[int, int]:
    """
    Row-level & table-level filtering in `body`:
      - `keep_changed=True` (`给译者版` `--remove-unchanged-tables`):
          Inside EVERY `<w:tbl>`, delete every `<w:tr>` where `not row_has_col2_changes(tr, active_hl_vals)`
          (keeping Row 0 + optional blank Row 1 ONLY if `keep_first_title_row=True` AND at least one row in that table has changes)!
          If 0 changed rows exist in `<w:tbl>`, remove the entire `<w:tbl>` + preceding title `<w:p>`(s) + blank `<w:p>`(s).
      - `keep_changed=False` (`无改点版` `--only-unchanged-tables`):
          Inside EVERY `<w:tbl>`, delete changed `<w:tr>`s so only unchanged rows remain (and remove emptied tables + titles).
    Returns `(kept_tables, removed_tables)`.
    """
    active_hl_vals = set() if no_highlight_colors else ({up_color, rv_color} | RV_HIGHLIGHT_VALS)
    children = list(body)
    n = len(children)
    to_remove = set()
    removed_tables = 0
    kept_tables = 0

    def _is_blank_p(el: ET.Element) -> bool:
        if el.tag.split("}")[-1] != "p":
            return False
        return "".join(t.text or "" for t in el.findall(".//w:t", NS)).strip() == ""

    def _is_title_p(el: ET.Element) -> bool:
        if el.tag.split("}")[-1] != "p":
            return False
        ptxt = "".join(t.text or "" for t in el.findall(".//w:t", NS)).strip()
        return bool(ptxt) and not _is_top_instruction_paragraph_text(ptxt)

    for idx, ch in enumerate(children):
        if ch.tag.split("}")[-1] != "tbl":
            continue
        rem_trs = ch.findall("w:tr", NS)
        if not rem_trs:
            should_remove = True
        else:
            header_idxs = set()
            if keep_first_title_row and len(rem_trs) >= 1:
                header_idxs.add(0)
                if is_two_row_chapter_header(rem_trs):
                    header_idxs.add(1)
                    if len(rem_trs) >= 3 and _is_blank_row(rem_trs[2], active_hl_vals):
                        header_idxs.add(2)
                elif len(rem_trs) >= 2 and _is_blank_row(rem_trs[1], active_hl_vals):
                    header_idxs.add(1)

            if keep_changed:
                changed_idxs = {i for i, tr in enumerate(rem_trs) if row_has_col2_changes(tr, active_hl_vals)}
                if not changed_idxs:
                    should_remove = True
                else:
                    keep_idxs = changed_idxs | header_idxs
                    for i, tr in enumerate(rem_trs):
                        if i not in keep_idxs:
                            ch.remove(tr)
                    should_remove = len(ch.findall("w:tr", NS)) == 0
            else:
                for tr in rem_trs:
                    if row_has_col2_changes(tr, active_hl_vals):
                        ch.remove(tr)
                surviving = ch.findall("w:tr", NS)
                if keep_first_title_row:
                    surviving_set = set(surviving)
                    body_unchanged = [
                        rem_trs[i]
                        for i in range(len(rem_trs))
                        if i not in header_idxs and rem_trs[i] in surviving_set and not _is_blank_row(rem_trs[i], active_hl_vals)
                    ]
                    should_remove = len(body_unchanged) == 0
                else:
                    has_nonblank = any(not _is_blank_row(tr, active_hl_vals) for tr in surviving)
                    should_remove = (not surviving) or (not has_nonblank)

        if should_remove:
            removed_tables += 1
            to_remove.add(idx)
            b_idx = idx - 1
            # Remove optional blank paragraph(s) between title <w:p> and <w:tbl>
            tmp_blanks = []
            while b_idx >= 0 and _is_blank_p(children[b_idx]):
                tmp_blanks.append(b_idx)
                b_idx -= 1
            if b_idx >= 0 and _is_title_p(children[b_idx]):
                for bi in tmp_blanks:
                    to_remove.add(bi)
                while b_idx >= 0 and _is_title_p(children[b_idx]):
                    to_remove.add(b_idx)
                    b_idx -= 1
            f_idx = idx + 1
            while f_idx < n and _is_blank_p(children[f_idx]):
                to_remove.add(f_idx)
                f_idx += 1
        else:
            kept_tables += 1

    if to_remove:
        for idx in sorted(to_remove, reverse=True):
            body.remove(children[idx])

    rem_children = list(body)
    if rem_children and rem_children[-1].tag.split("}")[-1] == "sectPr":
        while len(rem_children) >= 3 and _is_blank_p(rem_children[-2]) and _is_blank_p(rem_children[-3]):
            body.remove(rem_children[-2])
            rem_children = list(body)

    # ECMA-376 safety: ensure at least one <w:p> exists before <w:sectPr> if all tables were removed
    non_sect = [c for c in list(body) if c.tag.split("}")[-1] != "sectPr"]
    if not non_sect:
        body.insert(0, ET.Element(f"{W}p"))

    return kept_tables, removed_tables


def _full_nonws_char_spans(text: str, hl_color: str) -> list:
    if not text or not text.strip():
        return []
    lead = len(text) - len(text.lstrip())
    trail = len(text) - len(text.rstrip())
    end = max(lead, len(text) - trail)
    return [(lead, end, hl_color)] if end > lead else []


def _split_ranges_by_regex(text: str, sep_re: re.Pattern) -> list[tuple[int, int]]:
    ranges = []
    cur = 0
    for m in sep_re.finditer(text):
        end_p = m.end()
        if text[cur:end_p].strip():
            ranges.append((cur, end_p))
            cur = end_p
    if cur < len(text) and text[cur:].strip():
        ranges.append((cur, len(text)))
    return ranges or [(0, len(text))]


def _fallback_col1_span_range(
    col1_text: str,
    old_en_full: str,
    sp: dict,
    search_from: int = 0,
    failed_target: str = "",
) -> tuple[int, int] | None:
    """
    Resilient local Col 1 Chinese span locator used when `col1_substring` was omitted (e.g. legacy string map)
    or had minor quote/character divergence from `col1_text`.
    """
    n_cn = len(col1_text)
    if n_cn == 0 or not col1_text.strip():
        return None

    # 1. If a non-empty target missed exact/regex fuzzy match, try bounded SequenceMatcher alignment first
    clean_t = (failed_target or "").strip()
    if len(clean_t) >= 2:
        sub_h = col1_text[search_from:] if search_from < n_cn else col1_text
        base_off = search_from if search_from < n_cn else 0
        sm = difflib.SequenceMatcher(None, clean_t, sub_h, autojunk=False)
        blocks = [b for b in sm.get_matching_blocks() if b.size > 0]
        matched = sum(b.size for b in blocks)
        if blocks and matched >= max(2, int(len(clean_t) * 0.65)):
            s_i = base_off + blocks[0].b
            e_i = base_off + blocks[-1].b + blocks[-1].size
            if e_i > s_i:
                return (s_i, e_i)

    # 2. Proportional sentence / clause mapping from Col 2 English offsets onto Col 1 Chinese
    en_total = max(1, len(old_en_full))
    c_start = int(sp.get("old_en_char_start", 0))
    c_end = int(sp.get("old_en_char_end", c_start))
    r_start = max(0.0, min(1.0, c_start / en_total))
    r_end = max(r_start, min(1.0, c_end / en_total))
    est_center = max(search_from, min(n_cn - 1, round(((r_start + r_end) * 0.5) * n_cn)))
    est_start = max(search_from, min(n_cn - 1, round(r_start * n_cn)))
    est_end = max(est_start + 1, min(n_cn, round(r_end * n_cn)))

    if sp.get("type") == "rv":
        sents = _split_ranges_by_regex(col1_text, _CN_SENT_END_RE)
        chosen = [
            (s_a, s_b)
            for (s_a, s_b) in sents
            if s_b > search_from and (
                (s_a < est_end and s_b > est_start)
                or (s_a <= est_center < s_b)
            )
        ]
        if not chosen:
            avail = [(s_a, s_b) for (s_a, s_b) in sents if s_b > search_from] or sents
            best_s = min(avail, key=lambda pair: abs(((pair[0] + pair[1]) * 0.5) - est_center))
            chosen = [best_s]
        s_out = max(search_from, chosen[0][0])
        e_out = chosen[-1][1]
        while s_out < e_out and col1_text[s_out].isspace():
            s_out += 1
        return (s_out, e_out) if e_out > s_out else None

    clauses = _split_ranges_by_regex(col1_text, _CN_CLAUSE_END_RE)
    avail_cl = [(c_a, c_b) for (c_a, c_b) in clauses if c_b > search_from] or clauses
    best_cl = min(avail_cl, key=lambda pair: abs(((pair[0] + pair[1]) * 0.5) - est_center))
    cl_a, cl_b = max(search_from, best_cl[0]), best_cl[1]
    en_words = max(1, count_msword_en_words(sp.get("new_en") or sp.get("old_en_in_col3") or "word"))
    if en_words <= 4 and (cl_b - cl_a) > 12:
        win_len = min(cl_b - cl_a, max(4, en_words * 2 + 2))
        rel = (est_center - cl_a) / max(1, cl_b - cl_a)
        s_out = cl_a + max(0, min((cl_b - cl_a) - win_len, round(rel * ((cl_b - cl_a) - win_len))))
        e_out = min(cl_b, s_out + win_len)
    else:
        s_out, e_out = cl_a, cl_b
    while s_out < e_out and col1_text[s_out].isspace():
        s_out += 1
    return (s_out, e_out) if e_out > s_out else None


def compute_final_col2_word_counts(root: ET.Element, up_color: str, rv_color: str):
    return compute_col2_up_rv_words(root, up_color=up_color, rv_color=rv_color)


def main():
    parser = argparse.ArgumentParser(description="Step 2b: Highlight Col 1 & Col 3 (UP+RV), count Col 2 English words & output final docx.")
    parser.add_argument("--tasks", required=True, help="Path to step2_align_tasks.json")
    parser.add_argument("--alignments", required=True, help="Path to step2_alignments.json")
    parser.add_argument("--outdir", required=True, help="Output directory for final .docx")
    parser.add_argument("--lang", required=True, help="Language code, e.g. GR, FR, PT")
    parser.add_argument("--doc-name", required=True, help="Original document name for filename")
    parser.add_argument("--date", default="", help="Optional YYYYMMDD string, defaults to current local YYYYMMDD")
    parser.add_argument("--remove-unchanged-tables", action="store_true", help="Keep only changed rows/tables (translator version)")
    parser.add_argument("--only-unchanged-tables", action="store_true", help="Keep only unchanged rows/tables (unchanged-only version)")
    parser.add_argument("--keep-first-title-row", action="store_true", help="Full-article mode: preserve Row 0 chapter title row in changed tables")
    parser.add_argument("--no-highlight-colors", action="store_true", help="Skip Col 1/Col 3 highlights and top instruction lines")
    parser.add_argument("--no-col1-highlights", action="store_true", help="Skip Col 1 Chinese yellow/cyan highlights (highlight Col 2 & Col 3 only)")
    parser.add_argument("--no-balance-columns", action="store_true", help="Disable automatic adaptive 3-column width balancing")
    parser.add_argument("--no-header-lines", action="store_true", help="Disable automatic top Blue/Yellow highlight instruction lines")
    args = parser.parse_args()

    cfg = load_config()
    col2_up = cfg.get("col2_up_highlight", "yellow")
    col2_rv = cfg.get("col2_rv_highlight", "cyan")
    col3_up = cfg.get("col3_up_highlight", "shd:D9EEF7")
    col3_rv = cfg.get("col3_rv_highlight", "lightGray")
    use_hl = not bool(args.no_highlight_colors)
    highlight_col1 = use_hl and not bool(args.no_col1_highlights)

    with open(args.tasks, "r", encoding="utf-8") as f:
        tasks_data = json.load(f)
    with open(args.alignments, "r", encoding="utf-8") as f:
        align_map = json.load(f)

    draft_path = Path(tasks_data["draft_docx"])
    root, _, ns_pairs = load_docx_xml(draft_path)
    body = root.find("w:body", NS)

    # 1. Accept any unaccepted Col 1 Chinese track changes; normalize/accept Col 2 RV highlights (or strip highlights when --no-highlight-colors)
    accept_rv = cfg.get("accept_col2_rv_revisions", False)
    for tbl in body.findall("w:tbl", NS):
        for tr in tbl.findall("w:tr", NS):
            tcs = tr.findall("w:tc", NS)
            if tcs:
                accept_all_revisions_in_cell(tcs[0])
                strip_all_highlights_in_cell(tcs[0])
            if not use_hl:
                for tc_c in tcs:
                    strip_all_highlights_in_cell(tc_c)
                continue
            if is_skipped_row(tr):
                continue
            if len(tcs) >= 2:
                if accept_rv:
                    accept_rv_revisions_in_col2(tcs[1], target_rv_color=col2_rv)
                else:
                    normalize_rv_highlights_in_col2(tcs[1], target_rv_color=col2_rv)

    # 2. Apply Col 1 Chinese single underline (`u:single`) & Col 3 Foreign (`shd:D9EEF7` for UP, `lightGray` for RV)
    unmatched = []
    matched_count = 0
    if use_hl:
        for rt in tasks_data.get("row_tasks", []):
            tbl_el = body[rt["body_idx"]]
            tr = tbl_el.findall("w:tr", NS)[rt["row_idx"]]
            if is_skipped_row(tr):
                continue
            tcs = tr.findall("w:tc", NS)
            tc1, tc3 = tcs[0], tcs[2]
            col1_text = get_cell_plain_text(tc1)
            col3_text = get_cell_plain_text(tc3)

            if rt.get("is_full_cell_rv"):
                if highlight_col1:
                    c1_full = _full_nonws_char_spans(col1_text, "u:single")
                    apply_highlights_to_col3_cell(tc1, c1_full)
                c3_full = _full_nonws_char_spans(col3_text, col3_rv)
                apply_highlights_to_col3_cell(tc3, c3_full)
                matched_count += len(rt.get("spans", []))
                continue

            col1_char_spans = []
            col3_char_spans = []
            col1_cursor = 0
            col3_cursor = 0
            old_en_full = rt.get("old_en_full", "")

            for sp in rt["spans"]:
                sid = sp["span_id"]
                entry = align_map.get(sid, "")
                if isinstance(entry, dict):
                    col1_target = entry.get("col1_substring") or ""
                    col3_target = entry.get("col3_substring") or ""
                else:
                    col1_target = ""
                    col3_target = str(entry or "")

                # Synchronous Col 1 Chinese single underline (never yellow/cyan highlight) when highlight_col1 is enabled
                if highlight_col1 and col1_text.strip():
                    hl_c1 = "u:single"
                    loc1 = (
                        find_fuzzy_substring_span(col1_text, col1_target, search_from=col1_cursor)
                        if col1_target.strip()
                        else None
                    )
                    if loc1 is None:
                        loc1 = _fallback_col1_span_range(
                            col1_text=col1_text,
                            old_en_full=old_en_full,
                            sp=sp,
                            search_from=col1_cursor,
                            failed_target=col1_target,
                        )
                    if loc1 is not None:
                        s1, e1 = loc1
                        col1_char_spans.append((s1, e1, hl_c1))
                        col1_cursor = e1

                # Synchronous Col 3 Foreign highlight (soft light-blue shading for 'up', lightGray for 'rv')
                if not col3_target or not col3_target.strip():
                    unmatched.append((sid, "EMPTY_ALIGNMENT", sp["old_en_in_col3"][:60]))
                    continue

                loc3 = find_fuzzy_substring_span(col3_text, col3_target, search_from=col3_cursor)
                if loc3 is None:
                    unmatched.append((sid, col3_target, sp["old_en_in_col3"][:60]))
                    continue

                s3, e3 = loc3
                hl_c3 = col3_rv if sp["type"] == "rv" else col3_up
                col3_char_spans.append((s3, e3, hl_c3))
                col3_cursor = e3
                matched_count += 1

            if highlight_col1:
                apply_highlights_to_col3_cell(tc1, col1_char_spans)
            apply_highlights_to_col3_cell(tc3, col3_char_spans)

    if unmatched:
        print(f"WARNING: {len(unmatched)} span(s) failed to match in Col 3 and were skipped:")
        for sid, tstr, oen in unmatched:
            print(f"  - {sid}: target={tstr!r} (from old_en={oen!r})")

    # 2.5 Skipped-row removal + row-level & table-level filtering for translator / unchanged-only outputs
    if args.only_unchanged_tables:
        rem_skip = remove_skipped_rows_in_body(body)
        kept_cnt, removed_cnt = filter_tables_in_body(
            body,
            col2_up,
            col2_rv,
            keep_changed=False,
            keep_first_title_row=bool(args.keep_first_title_row),
            no_highlight_colors=not use_hl,
        )
        print(f"Kept {kept_cnt} unchanged tables (removed {removed_cnt} changed/empty tables, {rem_skip} skipped rows).")
    elif args.remove_unchanged_tables:
        rem_skip = remove_skipped_rows_in_body(body)
        kept_cnt, removed_cnt = filter_tables_in_body(
            body,
            col2_up,
            col2_rv,
            keep_changed=True,
            keep_first_title_row=bool(args.keep_first_title_row),
            no_highlight_colors=not use_hl,
        )
        print(f"Kept {kept_cnt} changed tables (removed {removed_cnt} unchanged/empty tables, {rem_skip} skipped rows).")

    # 3. Compute exact UP & RV English word counts from Col 2 (punctuation & skipped rows excluded)
    up_words, rv_words = (compute_final_col2_word_counts(root, col2_up, col2_rv) if use_hl else (0, 0))

    # 3.5 Adaptive trilingual 3-column width balancing & top highlight instruction lines
    if not args.no_balance_columns:
        balance_trilingual_table_widths(root)

    if not use_hl:
        update_top_highlight_instruction_lines(root, has_up=False, has_rv=False)
        for p_top in body.findall("w:p", NS):
            for rpr_top in p_top.findall(f".//{W}rPr"):
                if rpr_top.find(f"{W}highlight") is not None:
                    for hl_el in list(rpr_top.findall(f"{W}highlight")):
                        rpr_top.remove(hl_el)
    elif not args.no_header_lines:
        update_top_highlight_instruction_lines(root, has_up=(up_words > 0), has_rv=(rv_words > 0))
    else:
        header_paras = cfg.get("header_paragraphs", [])
        if header_paras:
            for idx, hp_text in enumerate(header_paras):
                p_el = ET.Element(f"{W}p")
                p_el.append(make_run(hp_text))
                body.insert(idx, p_el)

    count_tag = format_up_rv_count_tag(up_words, rv_words) if use_hl else ""

    date_str = sanitize_filename_part(
        args.date.strip() or datetime.now().strftime("%Y%m%d"),
        fallback=datetime.now().strftime("%Y%m%d"),
        max_len=16,
    )
    lang_str = sanitize_filename_part(args.lang.strip().upper(), fallback="XX", max_len=16)
    doc_label = sanitize_filename_part(args.doc_name.strip(), fallback="Document", max_len=140)
    if count_tag:
        out_filename = f"{lang_str}_{doc_label}_{count_tag}_{date_str}.docx"
    else:
        out_filename = f"{lang_str}_{doc_label}_{date_str}.docx"

    outdir = Path(args.outdir).resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    out_path = safe_join_output_file(outdir, out_filename)

    save_docx_xml(draft_path, out_path, root, ns_pairs)
    total_attempted = matched_count + len(unmatched)
    print(f"Successfully aligned {matched_count}/{total_attempted} Col 3 spans ({len(unmatched)} unmatched).")
    print(f"Final Col 2 Word Counts: UP = {up_words} words, RV = {rv_words} words")
    print(f"Final Document Saved: {out_path}")


if __name__ == "__main__":
    main()
