import argparse
import copy
import json
from datetime import datetime, timezone
from pathlib import Path
import xml.etree.ElementTree as ET

from docx_xml_utils import (
    INLINE_SEMANTIC_RPR_TAGS,
    NS,
    SKIPPED_ROW_MARKER,
    W,
    CommentIdAllocator,
    apply_full_sentence_highlight,
    build_rpr,
    clean_and_migrate_col3_paragraphs,
    compute_precision_yellow,
    count_msword_en_words,
    extract_cell_comment_anchors,
    get_cell_template_ppr,
    get_first_run_rpr,
    insert_rpr_child_ordered,
    load_config,
    load_docx_comments_map,
    load_docx_comments_xml,
    load_docx_styles_xml,
    load_docx_xml,
    make_run,
    make_track_el,
    new_comments_root,
    reanchor_comments_to_cell,
    remove_rpr_child,
    save_docx_xml,
)


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


def clean_col3_paragraphs(
    src_tc3: ET.Element,
    target_tc3: ET.Element,
    target_tc2: ET.Element,
    src_comments_map: dict = None,
    target_comments_root: ET.Element = None,
    id_allocator: CommentIdAllocator = None,
    src_styles_root: ET.Element = None,
) -> list:
    """
    Clone text, inline emphasis, and Word comments from library Col 3 while strictly preserving
    `src_tc3`'s own paragraph & run formatting (`src_styles_root`) and stripping invalid OOXML artifacts.
    """
    return clean_and_migrate_col3_paragraphs(
        src_tc3=src_tc3,
        target_tc3=target_tc3,
        target_tc2=target_tc2,
        src_comments_map=src_comments_map,
        target_comments_root=target_comments_root,
        id_allocator=id_allocator,
        src_styles_root=src_styles_root,
    )


def populate_skipped_col3_cell(tc3: ET.Element, fallback_tc: ET.Element = None) -> None:
    """Write a single red (`C00000`) paragraph `【未在总库匹配到原文，已跳过】` into Col 3 while keeping Col 1 & Col 2 intact."""
    tpl_pPr = get_cell_template_ppr(tc3, fallback_tc=fallback_tc)
    tpl_rPr = get_first_run_rpr(tc3, fallback_tc=fallback_tc)

    tc3_Pr = tc3.find("w:tcPr", NS)
    for ch in list(tc3):
        if ch is not tc3_Pr:
            tc3.remove(ch)

    p_el = ET.Element(f"{W}p")
    if tpl_pPr is not None:
        p_el.append(copy.deepcopy(tpl_pPr))

    r_el = ET.Element(f"{W}r")
    rPr = build_rpr(tpl_rPr, highlight=None, shd_fill=None)
    insert_rpr_child_ordered(rPr, ET.Element(f"{W}color", {f"{W}val": "C00000"}))
    r_el.append(rPr)
    t_el = ET.Element(f"{W}t")
    t_el.text = SKIPPED_ROW_MARKER
    r_el.append(t_el)
    p_el.append(r_el)
    tc3.append(p_el)


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
    orig_comments_root = load_docx_comments_xml(inp_path)
    inp_had_comments = orig_comments_root is not None
    working_comments_root = orig_comments_root if orig_comments_root is not None else new_comments_root()
    id_allocator = CommentIdAllocator(doc_root=root, comments_root=working_comments_root)

    body = root.find("w:body", NS)

    src_cache = {}  # Path -> (list_of_trs, comments_map_dict, styles_root)
    up_texts = []
    rv_texts = []

    for rm in manifest.get("rows", []):
        body_el = body[rm["body_idx"]]
        tr = body_el.findall("w:tr", NS)[rm["row_idx"]]
        tcs = tr.findall("w:tc", NS)
        tc2, tc3 = tcs[1], tcs[2]

        src_file = Path(rm["source_file"])
        if src_file not in src_cache:
            s_root, _, _ = load_docx_xml(src_file)
            s_cmap = load_docx_comments_map(src_file)
            s_styles = load_docx_styles_xml(src_file)
            src_cache[src_file] = (s_root.findall(".//w:tbl/w:tr", NS), s_cmap, s_styles)
        s_trs, s_cmap, s_styles = src_cache[src_file]
        s_tr = s_trs[rm["source_row_idx"]]
        s_tcs = s_tr.findall("w:tc", NS)

        # 1. Populate Col 3 from library Col 3 while preserving src_tc3's own formatting AND migrating Word comments
        col3_paras = clean_col3_paragraphs(
            src_tc3=s_tcs[2],
            target_tc3=tc3,
            target_tc2=tc2,
            src_comments_map=s_cmap,
            target_comments_root=working_comments_root,
            id_allocator=id_allocator,
            src_styles_root=s_styles,
        )
        tc3_Pr = tc3.find("w:tcPr", NS)
        for ch in list(tc3):
            if ch is not tc3_Pr:
                tc3.remove(ch)
        for cp in col3_paras:
            tc3.append(cp)

        # 2. Extract any input Col 2 comments BEFORE clearing tc2 so they can be re-anchored onto rebuilt tc2!
        tc2_comment_anchors = extract_cell_comment_anchors(tc2, working_comments_root)

        # 3. Rebuild Col 2 with white=old, shaded=diff + yellow/cyan (retaining track changes in Step 1!)
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
                    p_el.append(make_track_el("del", id_allocator.next_id(), author, date_iso, r_el))
                elif kind == "ins":
                    r_el = make_run(txt, base_rPr=base_rPr, highlight=hl, shd_fill=shd, is_del=False)
                    p_el.append(make_track_el("ins", id_allocator.next_id(), author, date_iso, r_el))
            tc2.append(p_el)

        if tc2_comment_anchors:
            reanchor_comments_to_cell(tc2, tc2_comment_anchors, working_comments_root, id_allocator)

    # 3. Populate skipped rows: keep Col 1 & Col 2 intact, write red `【未在总库匹配到原文，已跳过】` into Col 3
    skipped_list = manifest.get("skipped_rows", [])
    for sr in skipped_list:
        body_el = body[sr["body_idx"]]
        trs = body_el.findall("w:tr", NS)
        if 0 <= sr["row_idx"] < len(trs):
            tcs = trs[sr["row_idx"]].findall("w:tc", NS)
            if len(tcs) >= 3:
                populate_skipped_col3_cell(tcs[2], fallback_tc=tcs[1])

    final_comments_root = (
        working_comments_root
        if (inp_had_comments or len(working_comments_root.findall(f"{W}comment")) > 0)
        else None
    )

    out_path = Path(args.output).resolve()
    save_docx_xml(inp_path, out_path, root, ns_pairs, comments_root=final_comments_root)

    up_words = count_msword_en_words(" ".join(up_texts))
    rv_words = count_msword_en_words(" ".join(rv_texts))
    print(f"Step 1 Draft saved to: {out_path}")
    print(f"Provisional counts (pending human review): UP = {up_words} words, RV = {rv_words} words")
    if skipped_list:
        print(f"Skipped unmatched rows marked in Col 3: {len(skipped_list)}")


if __name__ == "__main__":
    main()
