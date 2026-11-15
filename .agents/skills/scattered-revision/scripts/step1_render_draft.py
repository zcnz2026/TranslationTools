import argparse
import copy
import json
from datetime import datetime, timezone
from pathlib import Path
import xml.etree.ElementTree as ET

from docx_xml_utils import (
    NS,
    W,
    apply_full_sentence_highlight,
    compute_precision_yellow,
    count_msword_en_words,
    get_cell_template_ppr,
    get_first_run_rpr,
    insert_rpr_child_ordered,
    load_config,
    load_docx_xml,
    make_run,
    make_track_el,
    remove_rpr_child,
    save_docx_xml,
)


INLINE_SEMANTIC_RPR_TAGS = {
    "b", "bCs", "i", "iCs", "u", "strike", "dstrike", "vertAlign", "caps", "smallCaps", "rtl", "cs"
}


def merge_styled_ops(styled_ops: list) -> list:
    """Merge consecutive (kind, text, hl, shd) tuples that share identical kind, hl, and shd."""
    merged = []
    for kind, txt, hl, shd in styled_ops:
        if not txt:
            continue
        if merged and merged[-1][0] == kind and merged[-1][2] == hl and merged[-1][3] == shd:
            merged[-1] = (kind, merged[-1][1] + txt, hl, shd)
        else:
            merged.append((kind, txt, hl, shd))
    return merged


def clean_col3_paragraphs(src_tc3: ET.Element, target_tc3: ET.Element, target_tc2: ET.Element) -> list:
    """
    Clone text/inline-emphasis from library Col 3 while strictly preserving the input document's
    original paragraph layout (pPr) and font/size formatting (rPr) from target_tc3 / target_tc2,
    as well as RTL directionality (bidi/rtl) and complex-script fonts for non-Latin scripts.
    """
    tpl_pPr = get_cell_template_ppr(target_tc3, fallback_tc=target_tc2)
    tpl_rPr = get_first_run_rpr(target_tc3, fallback_tc=target_tc2)

    paras = []
    for p in src_tc3.findall("w:p", NS):
        cp = copy.deepcopy(p)
        old_pPr = cp.find("w:pPr", NS)
        old_bidi = old_pPr.find("w:bidi", NS) if old_pPr is not None else None
        if old_pPr is not None:
            cp.remove(old_pPr)
        if tpl_pPr is not None:
            new_pPr = copy.deepcopy(tpl_pPr)
            if old_bidi is not None and new_pPr.find("w:bidi", NS) is None:
                after_bidi = {"adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "jc", "textAlignment", "outlineLvl", "rPr", "sectPr", "pPrChange"}
                ins_pos = len(new_pPr)
                for idx_c, ch_p in enumerate(new_pPr):
                    if ch_p.tag.split("}")[-1] in after_bidi:
                        ins_pos = idx_c
                        break
                new_pPr.insert(ins_pos, copy.deepcopy(old_bidi))
            cp.insert(0, new_pPr)
        elif old_bidi is not None:
            new_pPr = ET.Element(f"{W}pPr")
            new_pPr.append(copy.deepcopy(old_bidi))
            cp.insert(0, new_pPr)

        for r in cp.findall(".//w:r", NS):
            old_rPr = r.find("w:rPr", NS)
            new_rPr = copy.deepcopy(tpl_rPr)
            if old_rPr is not None:
                old_rf = old_rPr.find("w:rFonts", NS)
                new_rf = new_rPr.find("w:rFonts", NS)
                if old_rf is not None and new_rf is not None:
                    cs_font = old_rf.attrib.get(f"{W}cs")
                    if cs_font and f"{W}cs" not in new_rf.attrib:
                        new_rf.set(f"{W}cs", cs_font)
                for ch in old_rPr:
                    local = ch.tag.split("}")[-1]
                    if local in INLINE_SEMANTIC_RPR_TAGS:
                        insert_rpr_child_ordered(new_rPr, copy.deepcopy(ch))
                r.remove(old_rPr)
            if len(new_rPr) > 0:
                r.insert(0, new_rPr)
        paras.append(cp)
    return paras


def main():
    parser = argparse.ArgumentParser(description="Step 1b: Render Step 1 Draft .docx with yellow/cyan & retained track changes.")
    parser.add_argument("--manifest", required=True, help="Path to diff_manifest.json")
    parser.add_argument("--decisions", default="", help="Optional JSON file mapping sent_id -> 'yellow'|'cyan'|'none'")
    parser.add_argument("--output", required=True, help="Output Step 1 draft .docx path")
    args = parser.parse_args()

    cfg = load_config()
    author = cfg.get("default_author", "整理组")
    up_color = cfg.get("col2_up_highlight", "yellow")
    rv_color = cfg.get("col2_rv_highlight", "cyan")
    date_iso = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    with open(args.manifest, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    decisions = {}
    if args.decisions and Path(args.decisions).exists():
        with open(args.decisions, "r", encoding="utf-8") as f:
            decisions = json.load(f)

    inp_path = Path(manifest["input_docx"])
    root, _, ns_pairs = load_docx_xml(inp_path)
    body = root.find("w:body", NS)

    src_cache = {}
    rev_id = 1
    up_texts = []
    rv_texts = []

    for rm in manifest["rows"]:
        body_el = body[rm["body_idx"]]
        tr = body_el.findall("w:tr", NS)[rm["row_idx"]]
        tcs = tr.findall("w:tc", NS)
        tc2, tc3 = tcs[1], tcs[2]

        src_file = Path(rm["source_file"])
        if src_file not in src_cache:
            s_root, _, _ = load_docx_xml(src_file)
            src_cache[src_file] = s_root.findall(".//w:tbl/w:tr", NS)
        s_tr = src_cache[src_file][rm["source_row_idx"]]
        s_tcs = s_tr.findall("w:tc", NS)

        # 1. Populate Col 3 from library Col 3 while preserving target_tc3/tc2 original font & size
        col3_paras = clean_col3_paragraphs(s_tcs[2], tc3, tc2)
        tc3_Pr = tc3.find("w:tcPr", NS)
        for ch in list(tc3):
            if ch is not tc3_Pr:
                tc3.remove(ch)
        for cp in col3_paras:
            tc3.append(cp)

        # 2. Rebuild Col 2 with white=old, shaded=diff + yellow/cyan (retaining track changes in Step 1!)
        base_rPr = get_first_run_rpr(tc2, fallback_tc=tcs[0])
        orig_pPrs = [
            copy.deepcopy(p_item.find("w:pPr", NS)) if p_item.find("w:pPr", NS) is not None else None
            for p_item in tc2.findall("w:p", NS)
        ]

        styled_stream = []  # (kind, text, hl, shd)
        for reg in rm["regions"]:
            if not reg["is_quoted"]:
                styled_stream.append(("eq", reg["src_text"], None, None))
            else:
                shd_hex = reg["fill_hex"]
                sents = reg["sentences"]
                actions = [decisions.get(s["sent_id"], s["suggested_action"]) for s in sents]
                idx_s = 0
                while idx_s < len(sents):
                    if actions[idx_s] == "cyan":
                        c_ops = []
                        while idx_s < len(sents) and actions[idx_s] == "cyan":
                            c_ops.extend(sents[idx_s]["ops"])
                            idx_s += 1
                        toks = apply_full_sentence_highlight(c_ops, rv_color)
                        cur_rv_buf = []
                        for k, t, hl in toks:
                            s_fill = None if k == "del" else shd_hex
                            styled_stream.append((k, t, hl, s_fill))
                            if hl == rv_color and k in ("eq", "ins"):
                                cur_rv_buf.append(t)
                        rv_texts.append("".join(cur_rv_buf))
                    else:
                        nc_ops = []
                        while idx_s < len(sents) and actions[idx_s] != "cyan":
                            s_obj = sents[idx_s]
                            act = actions[idx_s]
                            if not s_obj["has_edit"] or act == "none":
                                for k, t in s_obj["ops"]:
                                    if k != "del":
                                        nc_ops.append(("eq", t))
                            else:
                                nc_ops.extend(s_obj["ops"])
                            idx_s += 1
                        toks = compute_precision_yellow(nc_ops)
                        cur_up_buf = []
                        for k, t, hl in toks:
                            hl_mapped = up_color if hl == "yellow" else None
                            s_fill = None if k == "del" else shd_hex
                            styled_stream.append((k, t, hl_mapped, s_fill))
                            if hl_mapped == up_color and k in ("eq", "ins"):
                                cur_up_buf.append(t)
                            elif cur_up_buf:
                                cur_up_buf.append(" ")
                        if cur_up_buf:
                            up_texts.append("".join(cur_up_buf))

        # Bridge pure-whitespace gaps between two identically highlighted tokens
        for idx_tok in range(1, len(styled_stream) - 1):
            k_c, t_c, hl_c, shd_c = styled_stream[idx_tok]
            if k_c == "eq" and hl_c is None and t_c.isspace():
                hl_prev = styled_stream[idx_tok - 1][2]
                hl_next = styled_stream[idx_tok + 1][2]
                if hl_prev is not None and hl_prev == hl_next:
                    styled_stream[idx_tok] = (k_c, t_c, hl_prev, shd_c)

        # Split styled_stream on '\n' if multi-paragraph inside cell
        para_streams = [[]]
        for kind, txt, hl, shd in styled_stream:
            if "\n" not in txt or kind in ("del", "ins"):
                para_streams[-1].append((kind, txt, hl, shd))
            else:
                parts = txt.split("\n")
                for p_i, part in enumerate(parts):
                    if p_i > 0:
                        para_streams.append([])
                    if part:
                        para_streams[-1].append((kind, part, hl, shd))

        tc2_Pr = tc2.find("w:tcPr", NS)
        for ch in list(tc2):
            if ch is not tc2_Pr:
                tc2.remove(ch)

        for p_idx, p_stream in enumerate(para_streams):
            p_el = ET.Element(f"{W}p")
            chosen_pPr = (
                orig_pPrs[p_idx]
                if p_idx < len(orig_pPrs)
                else (orig_pPrs[-1] if orig_pPrs else None)
            )
            if chosen_pPr is not None:
                p_el.append(copy.deepcopy(chosen_pPr))
            for kind, txt, hl, shd in merge_styled_ops(p_stream):
                if kind == "eq":
                    p_el.append(make_run(txt, base_rPr=base_rPr, highlight=hl, shd_fill=shd, is_del=False))
                elif kind == "del":
                    r_el = make_run(txt, base_rPr=base_rPr, highlight=hl, shd_fill=None, is_del=True)
                    p_el.append(make_track_el("del", rev_id, author, date_iso, r_el))
                    rev_id += 1
                elif kind == "ins":
                    r_el = make_run(txt, base_rPr=base_rPr, highlight=hl, shd_fill=shd, is_del=False)
                    p_el.append(make_track_el("ins", rev_id, author, date_iso, r_el))
                    rev_id += 1
            tc2.append(p_el)

    out_path = Path(args.output).resolve()
    save_docx_xml(inp_path, out_path, root, ns_pairs)

    up_words = count_msword_en_words(" ".join(up_texts))
    rv_words = count_msword_en_words(" ".join(rv_texts))
    print(f"Step 1 Draft saved to: {out_path}")
    print(f"Provisional counts (pending human review): UP = {up_words} words, RV = {rv_words} words")


if __name__ == "__main__":
    main()
