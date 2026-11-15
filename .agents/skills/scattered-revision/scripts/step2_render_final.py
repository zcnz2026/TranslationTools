import argparse
import json
from datetime import datetime
from pathlib import Path
import xml.etree.ElementTree as ET

from docx_xml_utils import (
    NS,
    W,
    apply_highlights_to_col3_cell,
    count_msword_en_words,
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
                # Split runs: RV runs get unwrapped into p, non-RV runs stay inside w:ins
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


def remove_unchanged_tables_from_body(body: ET.Element, up_color: str, rv_color: str) -> int:
    """
    Remove tables whose Col 2 has zero UP/RV highlights and zero track changes,
    along with each removed table's immediately preceding title paragraph(s) and trailing blank separator paragraphs.
    Returns number of removed tables.
    """
    active_hl_vals = {up_color, rv_color} | RV_HIGHLIGHT_VALS
    children = list(body)
    n = len(children)
    to_remove = set()
    removed_tables = 0

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
        has_change = False
        for tr in ch.findall("w:tr", NS):
            tcs = tr.findall("w:tc", NS)
            if len(tcs) < 2:
                continue
            tc2 = tcs[1]
            if tc2.find(".//w:del", NS) is not None or tc2.find(".//w:ins", NS) is not None:
                has_change = True
                break
            for hl in tc2.findall(".//w:rPr/w:highlight", NS):
                if hl.attrib.get(f"{W}val") in active_hl_vals:
                    has_change = True
                    break
            if has_change:
                break

        if not has_change:
            removed_tables += 1
            to_remove.add(idx)
            # Preceding non-empty title paragraph(s)
            b_idx = idx - 1
            while b_idx >= 0 and _is_nonempty_p(children[b_idx]):
                to_remove.add(b_idx)
                b_idx -= 1
            # Trailing blank separator paragraph(s)
            f_idx = idx + 1
            while f_idx < n and _is_blank_p(children[f_idx]):
                to_remove.add(f_idx)
                f_idx += 1

    if to_remove:
        for idx in sorted(to_remove, reverse=True):
            body.remove(children[idx])
        # Keep at most 1 trailing blank paragraph immediately before sectPr
        rem_children = list(body)
        if rem_children and rem_children[-1].tag.split("}")[-1] == "sectPr":
            while len(rem_children) >= 3 and _is_blank_p(rem_children[-2]) and _is_blank_p(rem_children[-3]):
                body.remove(rem_children[-2])
                rem_children = list(body)

    return removed_tables


def compute_final_col2_word_counts(root: ET.Element, up_color: str, rv_color: str):
    up_chunks = []
    rv_chunks = []
    for tbl in root.findall(".//w:tbl", NS):
        for tr in tbl.findall("w:tr", NS):
            tcs = tr.findall("w:tc", NS)
            if len(tcs) < 2:
                continue
            tc2 = tcs[1]
            for p in tc2.findall("w:p", NS):
                u_buf = []
                r_buf = []
                for elem in p.iter():
                    if elem.tag.split("}")[-1] == "r":
                        hl = elem.find("./w:rPr/w:highlight", NS)
                        hval = hl.attrib.get(f"{W}val") if hl is not None else None
                        wt = "".join(t.text or "" for t in elem.findall("w:t", NS))
                        if not wt:
                            continue
                        if hval == up_color:
                            u_buf.append(wt)
                        elif u_buf:
                            u_buf.append(" ")
                        if hval == rv_color:
                            r_buf.append(wt)
                        elif r_buf:
                            r_buf.append(" ")
                up_chunks.append("".join(u_buf))
                rv_chunks.append("".join(r_buf))
    return count_msword_en_words(" ".join(up_chunks)), count_msword_en_words(" ".join(rv_chunks))


def main():
    parser = argparse.ArgumentParser(description="Step 2b: Accept RV revisions, highlight Col 3 (UP+RV), count words & output final docx.")
    parser.add_argument("--tasks", required=True, help="Path to step2_align_tasks.json")
    parser.add_argument("--alignments", required=True, help="Path to step2_alignments.json (span_id -> target Col 3 text)")
    parser.add_argument("--outdir", required=True, help="Output directory for final .docx")
    parser.add_argument("--lang", required=True, help="Language code, e.g. FR, PT")
    parser.add_argument("--doc-name", required=True, help="Document type/title label for filename")
    parser.add_argument("--date", default="", help="Optional MMDD string, defaults to current local MMDD")
    parser.add_argument("--remove-unchanged-tables", action="store_true", help="Remove tables that have 0 Col 2 highlights/revisions")
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

    # 2.5 Optional: remove completely unchanged tables (and their titles) for translator-facing version
    if args.remove_unchanged_tables:
        removed_cnt = remove_unchanged_tables_from_body(body, col2_up, col2_rv)
        print(f"Removed {removed_cnt} unchanged tables (with 0 Col 2 revisions/highlights).")

    # 3. Optional top header paragraphs (if configured in config.json; empty by default)
    header_paras = cfg.get("header_paragraphs", [])
    if header_paras:
        for idx, hp_text in enumerate(header_paras):
            p_el = ET.Element(f"{W}p")
            p_el.append(make_run(hp_text))
            body.insert(idx, p_el)

    # 4. Compute exact UP & RV word counts from Col 2 only
    up_words, rv_words = compute_final_col2_word_counts(root, col2_up, col2_rv)
    parts = []
    if up_words > 0:
        parts.append(f"up{up_words}")
    if rv_words > 0:
        parts.append(f"rv{rv_words}")
    count_tag = f"({' '.join(parts)})" if parts else "(up0)"

    date_str = sanitize_filename_part(args.date.strip() or datetime.now().strftime("%m%d"), fallback="0101", max_len=16)
    lang_str = sanitize_filename_part(args.lang.strip().upper(), fallback="XX", max_len=16)
    doc_label = sanitize_filename_part(args.doc_name.strip(), fallback="定稿", max_len=120)
    out_filename = f"{lang_str}_{doc_label}_{count_tag}_{date_str}.docx"

    outdir = Path(args.outdir).resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    out_path = safe_join_output_file(outdir, out_filename)

    save_docx_xml(draft_path, out_path, root, ns_pairs)
    print(f"Successfully aligned {matched_count}/{matched_count} Col 3 spans (0 unmatched).")
    print(f"Final Col 2 Word Counts: UP = {up_words}, RV = {rv_words}")
    print(f"Final Document Saved: {out_path}")


if __name__ == "__main__":
    main()
