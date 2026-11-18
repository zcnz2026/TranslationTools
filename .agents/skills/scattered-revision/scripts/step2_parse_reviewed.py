import argparse
import json
from pathlib import Path

from docx_xml_utils import (
    NS,
    SKIPPED_ROW_MARKER,
    W,
    accept_all_revisions_in_cell,
    get_cell_plain_text,
    load_docx_xml,
)

RV_HIGHLIGHT_VALS = {"cyan", "blue", "darkCyan"}
UP_HIGHLIGHT_VALS = {"yellow"}


def extract_col2_spans(tc2):
    """
    Parse Col 2 XML (even after manual edits in MS Word) into ordered segments:
    [{'hl_group': 'up'|'rv'|None, 'old_en': str, 'new_en': str}, ...]
    """
    raw_items = []
    paragraphs = tc2.findall("w:p", NS)
    for p_i, p in enumerate(paragraphs):
        if p_i > 0:
            raw_items.append((None, "\n", "\n"))
        for ch in p:
            local = ch.tag.split("}")[-1]
            if local == "r":
                hl_el = ch.find("./w:rPr/w:highlight", NS)
                hval = hl_el.attrib.get(f"{W}val") if hl_el is not None else None
                txt = "".join(t.text or "" for t in ch.findall("w:t", NS))
                if txt:
                    grp = "rv" if hval in RV_HIGHLIGHT_VALS else ("up" if hval in UP_HIGHLIGHT_VALS else None)
                    raw_items.append((grp, txt, txt))
            elif local == "ins":
                for r in ch.findall(".//w:r", NS):
                    hl_el = r.find("./w:rPr/w:highlight", NS)
                    hval = hl_el.attrib.get(f"{W}val") if hl_el is not None else None
                    txt = "".join(t.text or "" for t in r.findall("w:t", NS))
                    if txt:
                        grp = "rv" if hval in RV_HIGHLIGHT_VALS else ("up" if hval in UP_HIGHLIGHT_VALS else None)
                        raw_items.append((grp, "", txt))
            elif local == "del":
                for r in ch.findall(".//w:r", NS):
                    hl_el = r.find("./w:rPr/w:highlight", NS)
                    hval = hl_el.attrib.get(f"{W}val") if hl_el is not None else None
                    txt = "".join(t.text or "" for t in r.findall("w:delText", NS))
                    if txt:
                        grp = "rv" if hval in RV_HIGHLIGHT_VALS else ("up" if hval in UP_HIGHLIGHT_VALS else None)
                        raw_items.append((grp, txt, ""))

    merged = []
    for grp, old_t, new_t in raw_items:
        if merged and merged[-1]["hl_group"] == grp:
            merged[-1]["old_en"] += old_t
            merged[-1]["new_en"] += new_t
        else:
            merged.append({"hl_group": grp, "old_en": old_t, "new_en": new_t})
    return merged


def _is_full_cell_rv_segments(segments: list) -> bool:
    """Return True when every visible/non-whitespace segment in Col 2 is highlighted `rv` (whole-cell RV)."""
    has_rv = False
    for seg in segments:
        has_visible_or_old = bool(seg["new_en"].strip() or seg["old_en"].strip())
        if not has_visible_or_old:
            continue
        if seg["hl_group"] == "rv":
            has_rv = True
        else:
            return False
    return has_rv


def main():
    parser = argparse.ArgumentParser(description="Step 2a: Parse reviewed Step 1 .docx and output Col 1 & Col 3 alignment tasks.")
    parser.add_argument("--draft", required=True, help="Path to reviewed Step 1 draft .docx")
    parser.add_argument("--workdir", required=True, help="Working directory for JSON files")
    args = parser.parse_args()

    draft_path = Path(args.draft).resolve()
    workdir = Path(args.workdir).resolve()
    workdir.mkdir(parents=True, exist_ok=True)

    root, _, _ = load_docx_xml(draft_path)
    body = root.find("w:body", NS)

    cur_title = ""
    global_row_idx = 0
    row_tasks = []

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
                col3_text = get_cell_plain_text(tcs[2])
                if not col3_text.strip() or col3_text.strip() == SKIPPED_ROW_MARKER:
                    global_row_idx += 1
                    continue

                accept_all_revisions_in_cell(tcs[0])
                col1_text = get_cell_plain_text(tcs[0])

                segments = extract_col2_spans(tcs[1])
                is_full_cell_rv = _is_full_cell_rv_segments(segments)
                old_en_full = "".join(seg["old_en"] for seg in segments)
                spans_to_align = []
                char_cursor = 0

                for seg in segments:
                    raw_old = seg["old_en"]
                    raw_start = char_cursor
                    raw_end = char_cursor + len(raw_old)
                    char_cursor = raw_end

                    if seg["hl_group"] in ("up", "rv"):
                        stripped_old = raw_old.strip()
                        if stripped_old:
                            lead_ws = len(raw_old) - len(raw_old.lstrip())
                            trail_ws = len(raw_old) - len(raw_old.rstrip())
                            c_start = raw_start + lead_ws
                            c_end = raw_end - trail_ws
                        else:
                            c_start = raw_start
                            c_end = raw_end

                        left_ctx = old_en_full[max(0, c_start - 45) : c_start]
                        right_ctx = old_en_full[c_end : c_end + 45]
                        spans_to_align.append({
                            "span_id": f"R{global_row_idx}_S{len(spans_to_align)}",
                            "type": seg["hl_group"],  # 'up' -> Col1 yellow / Col3 yellow; 'rv' -> Col1 cyan / Col3 lightGray
                            "old_en_in_col3": stripped_old,
                            "old_en_char_start": c_start,
                            "old_en_char_end": c_end,
                            "new_en": seg["new_en"].strip(),
                            "left_old_context": left_ctx,
                            "right_old_context": right_ctx,
                        })

                if spans_to_align:
                    row_tasks.append({
                        "global_row_idx": global_row_idx,
                        "body_idx": body_idx,
                        "row_idx": r_idx,
                        "article_title": cur_title,
                        "col1_text": col1_text,
                        "col3_text": col3_text,
                        "old_en_full": old_en_full,
                        "is_full_cell_rv": is_full_cell_rv,
                        "spans": spans_to_align,
                    })
                global_row_idx += 1

    out_json = workdir / "step2_align_tasks.json"
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump({
            "draft_docx": str(draft_path),
            "row_tasks": row_tasks,
        }, f, ensure_ascii=False, indent=2)

    total_spans = sum(len(rt["spans"]) for rt in row_tasks)
    print(f"Alignment tasks written to: {out_json}")
    print(f"Total rows with highlights: {len(row_tasks)}, total spans to align to Col 1 & Col 3: {total_spans}")


if __name__ == "__main__":
    main()
