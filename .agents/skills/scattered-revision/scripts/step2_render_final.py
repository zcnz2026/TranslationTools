import argparse
import json
from datetime import datetime
from pathlib import Path
import xml.etree.ElementTree as ET

from docx_xml_utils import (
    NS,
    W,
    apply_highlights_to_col3_cell,
    compute_col2_up_rv_words,
    find_fuzzy_substring_span,
    get_cell_plain_text,
    insert_rpr_child_ordered,
    load_config,
    load_docx_xml,
    make_run,
    safe_join_output_file,
    sanitize_filename_part,
    save_docx_xml,
)

RV_HIGHLIGHT_VALS = {"cyan", "blue", "darkCyan"}


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
                    if r.tag.split("}")[-1] != "r":
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


def table_has_col2_changes(tbl_el: ET.Element, active_hl_vals: set) -> bool:
    for tr in tbl_el.findall("w:tr", NS):
        tcs = tr.findall("w:tc", NS)
        if len(tcs) < 2:
            continue
        tc2 = tcs[1]
        if tc2.find(".//w:del", NS) is not None or tc2.find(".//w:ins", NS) is not None:
            return True
        for hl in tc2.findall(".//w:rPr/w:highlight", NS):
            if hl.attrib.get(f"{W}val") in active_hl_vals:
                return True
    return False


def filter_tables_in_body(body: ET.Element, up_color: str, rv_color: str, keep_changed: bool) -> tuple[int, int]:
    """
    Filter tables in `body` along with each removed table's preceding title paragraph(s) and trailing blank separator(s).
    - keep_changed=True  -> keep ONLY tables with Col 2 changes/highlights (translator version)
    - keep_changed=False -> keep ONLY tables with ZERO Col 2 changes/highlights (unchanged-only version)
    Returns (kept_tables, removed_tables).
    """
    active_hl_vals = {up_color, rv_color} | RV_HIGHLIGHT_VALS
    children = list(body)
    n = len(children)
    to_remove = set()
    removed_tables = 0
    kept_tables = 0

    def _is_blank_p(el: ET.Element) -> bool:
        if el.tag.split("}")[-1] != "p":
            return False
        return "".join(t.text or "" for t in el.findall(".//w:t", NS)).strip() == ""

    def _is_nonempty_p(el: ET.Element) -> bool:
        if el.tag.split("}")[-1] != "p":
            return False
        return "".join(t.text or "" for t in el.findall(".//w:t", NS)).strip() != ""

    for idx, ch in enumerate(children):
        if ch.tag.split("}")[-1] != "tbl":
            continue
        has_change = table_has_col2_changes(ch, active_hl_vals)
        should_remove = (not has_change) if keep_changed else has_change
        if should_remove:
            removed_tables += 1
            to_remove.add(idx)
            b_idx = idx - 1
            while b_idx >= 0 and _is_nonempty_p(children[b_idx]):
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


def compute_final_col2_word_counts(root: ET.Element, up_color: str, rv_color: str):
    return compute_col2_up_rv_words(root, up_color=up_color, rv_color=rv_color)


def main():
    parser = argparse.ArgumentParser(description="Step 2b: Highlight Col 3 (UP+RV), count Col 2 English words & output final docx.")
    parser.add_argument("--tasks", required=True, help="Path to step2_align_tasks.json")
    parser.add_argument("--alignments", required=True, help="Path to step2_alignments.json (span_id -> target Col 3 text)")
    parser.add_argument("--outdir", required=True, help="Output directory for final .docx")
    parser.add_argument("--lang", required=True, help="Language code, e.g. GR, FR, PT")
    parser.add_argument("--doc-name", required=True, help="Original document name for filename")
    parser.add_argument("--date", default="", help="Optional YYYYMMDD string, defaults to current local YYYYMMDD")
    parser.add_argument("--remove-unchanged-tables", action="store_true", help="Keep only tables that have Col 2 highlights/revisions")
    parser.add_argument("--only-unchanged-tables", action="store_true", help="Keep only tables that have ZERO Col 2 highlights/revisions")
    args = parser.parse_args()

    cfg = load_config()
    col2_up = cfg.get("col2_up_highlight", "yellow")
    col2_rv = cfg.get("col2_rv_highlight", "cyan")
    col3_up = cfg.get("col3_up_highlight", "yellow")
    col3_rv = cfg.get("col3_rv_highlight", "lightGray")

    with open(args.tasks, "r", encoding="utf-8") as f:
        tasks_data = json.load(f)
    with open(args.alignments, "r", encoding="utf-8") as f:
        align_map = json.load(f)

    draft_path = Path(tasks_data["draft_docx"])
    root, _, ns_pairs = load_docx_xml(draft_path)
    body = root.find("w:body", NS)

    # 1. Handle RV highlights/revisions in ALL Col 2 cells (default: retain Track Changes inside cyan!)
    accept_rv = cfg.get("accept_col2_rv_revisions", False)
    for tbl in body.findall("w:tbl", NS):
        for tr in tbl.findall("w:tr", NS):
            tcs = tr.findall("w:tc", NS)
            if len(tcs) >= 2:
                if accept_rv:
                    accept_rv_revisions_in_col2(tcs[1], target_rv_color=col2_rv)
                else:
                    normalize_rv_highlights_in_col2(tcs[1], target_rv_color=col2_rv)

    # 2. Apply Col 3 highlights (yellow for UP, lightGray for RV, NO extra line breaks)
    unmatched = []
    matched_count = 0
    for rt in tasks_data["row_tasks"]:
        tbl_el = body[rt["body_idx"]]
        tr = tbl_el.findall("w:tr", NS)[rt["row_idx"]]
        tc3 = tr.findall("w:tc", NS)[2]
        col3_text = get_cell_plain_text(tc3)

        char_spans = []
        cursor = 0
        for sp in rt["spans"]:
            sid = sp["span_id"]
            target_str = align_map.get(sid, "")
            if isinstance(target_str, dict):
                target_str = target_str.get("col3_substring", "")
            if not target_str or not target_str.strip():
                unmatched.append((sid, "EMPTY_ALIGNMENT", sp["old_en_in_col3"][:60]))
                continue

            loc = find_fuzzy_substring_span(col3_text, target_str, search_from=cursor)
            if loc is None:
                unmatched.append((sid, target_str, sp["old_en_in_col3"][:60]))
                continue

            s_idx, e_idx = loc
            hl_color = col3_rv if sp["type"] == "rv" else col3_up
            char_spans.append((s_idx, e_idx, hl_color))
            cursor = e_idx
            matched_count += 1

        apply_highlights_to_col3_cell(tc3, char_spans)

    if unmatched:
        print("WARNING: Some spans failed to match in Col 3:")
        for sid, tstr, oen in unmatched:
            print(f"  - {sid}: target={tstr!r} (from old_en={oen!r})")
        raise SystemExit(2)

    # 2.5 Optional table filtering
    if args.only_unchanged_tables:
        kept_cnt, removed_cnt = filter_tables_in_body(body, col2_up, col2_rv, keep_changed=False)
        print(f"Kept {kept_cnt} unchanged tables (removed {removed_cnt} changed tables).")
    elif args.remove_unchanged_tables:
        kept_cnt, removed_cnt = filter_tables_in_body(body, col2_up, col2_rv, keep_changed=True)
        print(f"Kept {kept_cnt} changed tables (removed {removed_cnt} unchanged tables).")

    # 3. Optional top header paragraphs (if configured in config.json; empty by default)
    header_paras = cfg.get("header_paragraphs", [])
    if header_paras:
        for idx, hp_text in enumerate(header_paras):
            p_el = ET.Element(f"{W}p")
            p_el.append(make_run(hp_text))
            body.insert(idx, p_el)

    # 4. Compute exact UP & RV English word counts from Col 2 (punctuation excluded)
    up_words, rv_words = compute_final_col2_word_counts(root, col2_up, col2_rv)
    count_tag = f"(UP {up_words} words, RV {rv_words} words)"

    date_str = sanitize_filename_part(
        args.date.strip() or datetime.now().strftime("%Y%m%d"),
        fallback=datetime.now().strftime("%Y%m%d"),
        max_len=16,
    )
    lang_str = sanitize_filename_part(args.lang.strip().upper(), fallback="XX", max_len=16)
    doc_label = sanitize_filename_part(args.doc_name.strip(), fallback="Document", max_len=140)
    out_filename = f"{lang_str}_{doc_label}_{count_tag}_{date_str}.docx"

    outdir = Path(args.outdir).resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    out_path = safe_join_output_file(outdir, out_filename)

    save_docx_xml(draft_path, out_path, root, ns_pairs)
    print(f"Successfully aligned {matched_count}/{matched_count} Col 3 spans (0 unmatched).")
    print(f"Final Col 2 Word Counts: UP = {up_words} words, RV = {rv_words} words")
    print(f"Final Document Saved: {out_path}")


if __name__ == "__main__":
    main()
