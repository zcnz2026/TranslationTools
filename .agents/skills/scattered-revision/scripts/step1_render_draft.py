import argparse
import copy
import json
from datetime import datetime, timezone
from pathlib import Path
import xml.etree.ElementTree as ET

from docx_xml_utils import (
    DEFAULT_INLINE_STYLE,
    INLINE_SEMANTIC_RPR_TAGS,
    NS,
    SKIPPED_ROW_MARKER,
    W,
    CommentIdAllocator,
    accept_all_revisions_in_cell,
    apply_full_sentence_highlight,
    balance_trilingual_table_widths,
    build_rpr,
    clean_and_migrate_col3_paragraphs,
    clean_col2_ppr,
    compute_precision_yellow,
    count_msword_en_words,
    extract_cell_comment_anchors,
    get_cell_template_ppr,
    get_clean_col2_base_rpr,
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
    should_promote_region_to_cyan,
    split_text_by_style_spans,
    strip_all_highlights_in_cell,
    update_top_highlight_instruction_lines,
)


def merge_styled_ops(styled_ops: list) -> list:
    """
    Merge consecutive styled tuples that share identical `(kind, hl, shd, inline_style)`.
    Supports both legacy 4-tuples `(kind, txt, hl, shd)` and 5-tuples `(kind, txt, hl, shd, inline_style)`.
    """
    merged = []
    for item in styled_ops:
        if len(item) == 4:
            kind, txt, hl, shd = item
            if not txt:
                continue
            if merged and len(merged[-1]) == 4 and merged[-1][0] == kind and merged[-1][2] == hl and merged[-1][3] == shd:
                merged[-1] = (kind, merged[-1][1] + txt, hl, shd)
            else:
                merged.append((kind, txt, hl, shd))
        else:
            kind, txt, hl, shd, inl_st = item
            if not txt:
                continue
            if (
                merged
                and len(merged[-1]) == 5
                and merged[-1][0] == kind
                and merged[-1][2] == hl
                and merged[-1][3] == shd
                and merged[-1][4] == inl_st
            ):
                merged[-1] = (kind, merged[-1][1] + txt, hl, shd, inl_st)
            else:
                merged.append((kind, txt, hl, shd, inl_st))
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


def _populate_clean_library_cell(
    src_tc: ET.Element,
    target_tc: ET.Element,
    fallback_tc: ET.Element,
    src_comments_map: dict,
    target_comments_root: ET.Element,
    id_allocator: CommentIdAllocator,
    src_styles_root: ET.Element,
    ensure_cjk_songti: bool = False,
) -> None:
    """Replace `target_tc` content with clean library `src_tc` paragraphs (preserving inline styles & comments, zero shading/highlights)."""
    paras = clean_and_migrate_col3_paragraphs(
        src_tc3=src_tc,
        target_tc3=target_tc,
        target_tc2=fallback_tc,
        src_comments_map=src_comments_map,
        target_comments_root=target_comments_root,
        id_allocator=id_allocator,
        src_styles_root=src_styles_root,
    )
    if ensure_cjk_songti:
        for p_el in paras:
            for rf in p_el.findall(f".//{W}rFonts"):
                if not rf.attrib.get(f"{W}eastAsia"):
                    rf.set(f"{W}eastAsia", "宋体")

    tc_Pr = target_tc.find("w:tcPr", NS)
    if tc_Pr is not None:
        for shd_el in list(tc_Pr.findall("w:shd", NS)):
            tc_Pr.remove(shd_el)
    for ch in list(target_tc):
        if ch is not tc_Pr:
            target_tc.remove(ch)
    for cp in paras:
        target_tc.append(cp)


def export_raw_trilingual_docx(
    inp_path: Path,
    raw_out_path: Path,
    manifest: dict,
    src_cache: dict,
) -> None:
    """
    Export the pre-comparison extracted 3-language raw `.docx` (`对比前提取的三语原稿`):
      - Col 1 = matched library Col 1 (`s_tcs[0]` clean)
      - Col 2 = matched library old English (`s_tcs[1]` clean)
      - Col 3 = matched library Foreign (`s_tcs[2]` with preserved formatting & comments)
      - Adaptive balanced 3-column widths & Landscape A4.
    """
    raw_root, _, raw_ns_pairs = load_docx_xml(inp_path)
    raw_comments_root = new_comments_root()
    raw_id_alloc = CommentIdAllocator(doc_root=raw_root, comments_root=raw_comments_root)
    raw_body = raw_root.find("w:body", NS)

    update_top_highlight_instruction_lines(raw_root, has_up=False, has_rv=False)

    for rm in manifest.get("rows", []):
        body_el = raw_body[rm["body_idx"]]
        tr = body_el.findall("w:tr", NS)[rm["row_idx"]]
        r_tcs = tr.findall("w:tc", NS)
        if len(r_tcs) < 3:
            continue

        src_file = Path(rm["source_file"])
        if src_file not in src_cache:
            s_root, _, _ = load_docx_xml(src_file)
            s_cmap = load_docx_comments_map(src_file)
            s_styles = load_docx_styles_xml(src_file)
            src_cache[src_file] = (s_root.findall(".//w:tbl/w:tr", NS), s_cmap, s_styles)
        s_trs, s_cmap, s_styles = src_cache[src_file]
        s_tcs = s_trs[rm["source_row_idx"]].findall("w:tc", NS)

        _populate_clean_library_cell(
            src_tc=s_tcs[0],
            target_tc=r_tcs[0],
            fallback_tc=r_tcs[0],
            src_comments_map=s_cmap,
            target_comments_root=raw_comments_root,
            id_allocator=raw_id_alloc,
            src_styles_root=s_styles,
            ensure_cjk_songti=True,
        )
        _populate_clean_library_cell(
            src_tc=s_tcs[1],
            target_tc=r_tcs[1],
            fallback_tc=r_tcs[1],
            src_comments_map=s_cmap,
            target_comments_root=raw_comments_root,
            id_allocator=raw_id_alloc,
            src_styles_root=s_styles,
            ensure_cjk_songti=False,
        )
        _populate_clean_library_cell(
            src_tc=s_tcs[2],
            target_tc=r_tcs[2],
            fallback_tc=r_tcs[1],
            src_comments_map=s_cmap,
            target_comments_root=raw_comments_root,
            id_allocator=raw_id_alloc,
            src_styles_root=s_styles,
            ensure_cjk_songti=False,
        )

    for sr in manifest.get("skipped_rows", []):
        body_el = raw_body[sr["body_idx"]]
        trs = body_el.findall("w:tr", NS)
        if 0 <= sr["row_idx"] < len(trs):
            r_tcs = trs[sr["row_idx"]].findall("w:tc", NS)
            if len(r_tcs) >= 3:
                accept_all_revisions_in_cell(r_tcs[0])
                strip_all_highlights_in_cell(r_tcs[0])
                strip_all_highlights_in_cell(r_tcs[1])
                populate_skipped_col3_cell(r_tcs[2], fallback_tc=r_tcs[1])

    balance_trilingual_table_widths(raw_root)

    final_raw_comments = (
        raw_comments_root
        if len(raw_comments_root.findall(f"{W}comment")) > 0
        else new_comments_root()
    )
    save_docx_xml(inp_path, raw_out_path, raw_root, raw_ns_pairs, comments_root=final_raw_comments)
    print(f"Raw Trilingual Document saved to: {raw_out_path}")


def _bridge_yellow_gaps_in_styled_stream(styled_stream: list, up_color: str, max_gap_words: int) -> None:
    """
    Ensure any two `up_color` (yellow) spans in the same cell's `styled_stream` separated by
    `0 <= g_words <= max_gap_words` unhighlighted tokens (with no `\\n` and no `rv` highlight in between)
    are bridged into one continuous `up_color` highlight!
    """
    if not styled_stream:
        return

    # 1. Bridge pure horizontal whitespace between two identically highlighted tokens (yellow or cyan)
    for idx_tok in range(1, len(styled_stream) - 1):
        k_c, t_c, hl_c, shd_c, inl_c = styled_stream[idx_tok]
        if k_c == "eq" and hl_c is None and t_c.isspace() and "\n" not in t_c:
            hl_prev = styled_stream[idx_tok - 1][2]
            hl_next = styled_stream[idx_tok + 1][2]
            if hl_prev is not None and hl_prev == hl_next:
                styled_stream[idx_tok] = (k_c, t_c, hl_prev, shd_c, inl_c)

    if max_gap_words <= 0:
        return

    # 2. Bridge unhighlighted internal gaps between two `up_color` spans when `g_words <= max_gap_words`
    n = len(styled_stream)
    i = 0
    while i < n:
        if styled_stream[i][2] == up_color:
            j = i + 1
            while j < n and styled_stream[j][2] == up_color:
                j += 1
            # Scan unhighlighted gap [j .. m - 1]
            m = j
            has_other_hl_or_nl = False
            gap_vis_parts = []
            while m < n and styled_stream[m][2] is None:
                k_m, t_m, _, _, _ = styled_stream[m]
                if "\n" in t_m:
                    has_other_hl_or_nl = True
                    break
                if k_m in ("eq", "ins"):
                    gap_vis_parts.append(t_m)
                m += 1
            if not has_other_hl_or_nl and j < m < n and styled_stream[m][2] == up_color:
                g_words = count_msword_en_words("".join(gap_vis_parts))
                if g_words <= max_gap_words:
                    for p in range(j, m):
                        kp, tp, _, sp_fill, inlp = styled_stream[p]
                        styled_stream[p] = (kp, tp, up_color, sp_fill, inlp)
                    i = m
                    continue
            i = j
        else:
            i += 1


def main():
    parser = argparse.ArgumentParser(description="Step 1b: Render Step 1 Draft .docx with yellow/cyan & retained track changes.")
    parser.add_argument("--manifest", required=True, help="Path to diff_manifest.json")
    parser.add_argument("--decisions", default="", help="Optional JSON file mapping sent_id -> 'yellow'|'cyan'|'none'")
    parser.add_argument("--output", required=True, help="Output Step 1 draft .docx path")
    parser.add_argument(
        "--yellow-gap-words",
        type=int,
        default=3,
        help="Bridge two yellow spans in the same cell when separated by <= N unhighlighted English words (default: 3)",
    )
    parser.add_argument(
        "--no-highlight-colors",
        action="store_true",
        help="Render Col 2 track changes, shading, and inline colors without any <w:highlight> tags",
    )
    parser.add_argument(
        "--raw-trilingual-output",
        default="",
        help="Optional output path to export the pre-comparison extracted 3-language raw .docx",
    )
    args = parser.parse_args()

    cfg = load_config()
    author = cfg.get("default_author", "整理组")
    up_color = cfg.get("col2_up_highlight", "yellow")
    rv_color = cfg.get("col2_rv_highlight", "cyan")
    use_hl = not bool(args.no_highlight_colors)
    max_gap_words = max(0, int(args.yellow_gap_words))
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

    # Preprocess Col 1 Chinese: accept 100% of track changes in Col 1 across all table rows!
    for tr_all in body.findall(".//w:tbl/w:tr", NS):
        tcs_all = tr_all.findall("w:tc", NS)
        if tcs_all:
            accept_all_revisions_in_cell(tcs_all[0])
            if not use_hl:
                for tc_c in tcs_all:
                    strip_all_highlights_in_cell(tc_c)

    if not use_hl:
        update_top_highlight_instruction_lines(root, has_up=False, has_rv=False)
        for p_top in body.findall("w:p", NS):
            for rpr_top in p_top.findall(f".//{W}rPr"):
                remove_rpr_child(rpr_top, "highlight")

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

        # 3. Rebuild Col 2 with clean black unbolded base_rPr + exact per-char inline styles + yellow/cyan & track changes
        base_rPr = get_clean_col2_base_rpr(tc2, fallback_tc=tcs[0])
        orig_pPrs = [
            clean_col2_ppr(p_item.find("w:pPr", NS))
            for p_item in tc2.findall("w:p", NS)
        ]

        styled_stream = []  # (kind, text, hl, shd, inline_style)
        for reg in rm["regions"]:
            if not reg["is_quoted"]:
                for sub_txt, inl_st in split_text_by_style_spans(reg["src_text"], reg.get("style_spans", [])):
                    styled_stream.append(("eq", sub_txt, None, None, inl_st))
            else:
                shd_hex = reg["fill_hex"]
                sents = reg["sentences"]
                actions = [decisions.get(s["sent_id"], s["suggested_action"]) for s in sents]

                # Item 8.4: Promote heavily fragmented whole paragraph/region to clean 'cyan' (RV)
                if use_hl and should_promote_region_to_cyan(sents, actions, max_gap_words=max_gap_words):
                    actions = ["cyan"] * len(sents)

                idx_s = 0
                while idx_s < len(sents):
                    if use_hl and actions[idx_s] == "cyan":
                        c_ops = []
                        c_op_styles = []
                        while idx_s < len(sents) and actions[idx_s] == "cyan":
                            s_cur = sents[idx_s]
                            s_ops = s_cur["ops"]
                            s_sts = s_cur.get("op_styles") or [[] for _ in s_ops]
                            c_ops.extend(s_ops)
                            c_op_styles.extend(s_sts)
                            idx_s += 1
                        toks = apply_full_sentence_highlight(c_ops, rv_color)
                        for (k, t, hl), op_sp in zip(toks, c_op_styles):
                            s_fill = None if k == "del" else shd_hex
                            for sub_t, inl_st in split_text_by_style_spans(t, op_sp):
                                styled_stream.append((k, sub_t, hl, s_fill, inl_st))
                    else:
                        nc_ops = []
                        nc_op_styles = []
                        while idx_s < len(sents) and not (use_hl and actions[idx_s] == "cyan"):
                            s_obj = sents[idx_s]
                            act = actions[idx_s]
                            s_ops = s_obj["ops"]
                            s_sts = s_obj.get("op_styles") or [[] for _ in s_ops]
                            if not s_obj["has_edit"] or (use_hl and act == "none"):
                                for (k, t), op_sp in zip(s_ops, s_sts):
                                    if k != "del":
                                        nc_ops.append(("eq", t))
                                        nc_op_styles.append(op_sp)
                            else:
                                nc_ops.extend(s_ops)
                                nc_op_styles.extend(s_sts)
                            idx_s += 1

                        if use_hl:
                            toks = compute_precision_yellow(nc_ops, max_gap_words=max_gap_words)
                        else:
                            toks = [(k, t, None) for k, t in nc_ops]

                        for (k, t, hl), op_sp in zip(toks, nc_op_styles):
                            hl_mapped = up_color if (use_hl and hl == "yellow") else None
                            s_fill = None if k == "del" else shd_hex
                            for sub_t, inl_st in split_text_by_style_spans(t, op_sp):
                                styled_stream.append((k, sub_t, hl_mapped, s_fill, inl_st))

        if use_hl:
            _bridge_yellow_gaps_in_styled_stream(styled_stream, up_color=up_color, max_gap_words=max_gap_words)
            u_buf = []
            r_buf = []
            for k_c, t_c, hl_c, _, _ in styled_stream:
                if k_c in ("eq", "ins"):
                    if hl_c == up_color:
                        u_buf.append(t_c)
                    elif u_buf:
                        u_buf.append(" ")
                    if hl_c == rv_color:
                        r_buf.append(t_c)
                    elif r_buf:
                        r_buf.append(" ")
            if u_buf:
                up_texts.append("".join(u_buf))
            if r_buf:
                rv_texts.append("".join(r_buf))

        # Split styled_stream on '\n' if multi-paragraph inside cell
        para_streams = [[]]
        for kind, txt, hl, shd, inl_st in styled_stream:
            if "\n" not in txt or kind in ("del", "ins"):
                para_streams[-1].append((kind, txt, hl, shd, inl_st))
            else:
                parts = txt.split("\n")
                for p_i, part in enumerate(parts):
                    if p_i > 0:
                        para_streams.append([])
                    if part:
                        para_streams[-1].append((kind, part, hl, shd, inl_st))

        tc2_Pr = tc2.find("w:tcPr", NS)
        if tc2_Pr is not None:
            for shd_tc in list(tc2_Pr.findall("w:shd", NS)):
                tc2_Pr.remove(shd_tc)
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
            for kind, txt, hl, shd, inl_st in merge_styled_ops(p_stream):
                eff_inl = inl_st if inl_st != DEFAULT_INLINE_STYLE else None
                if kind == "eq":
                    p_el.append(make_run(txt, base_rPr=base_rPr, highlight=hl, shd_fill=shd, is_del=False, inline_style=eff_inl))
                elif kind == "del":
                    r_el = make_run(txt, base_rPr=base_rPr, highlight=hl, shd_fill=None, is_del=True, inline_style=eff_inl)
                    p_el.append(make_track_el("del", id_allocator.next_id(), author, date_iso, r_el))
                elif kind == "ins":
                    r_el = make_run(txt, base_rPr=base_rPr, highlight=hl, shd_fill=shd, is_del=False, inline_style=eff_inl)
                    p_el.append(make_track_el("ins", id_allocator.next_id(), author, date_iso, r_el))
            tc2.append(p_el)

        if tc2_comment_anchors:
            reanchor_comments_to_cell(tc2, tc2_comment_anchors, working_comments_root, id_allocator)

    # 4. Populate skipped rows: keep Col 1 & Col 2 intact, write red `【未在总库匹配到原文，已跳过】` into Col 3
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

    if args.raw_trilingual_output and args.raw_trilingual_output.strip():
        raw_out_path = Path(args.raw_trilingual_output.strip()).resolve()
        export_raw_trilingual_docx(inp_path, raw_out_path, manifest, src_cache)

    up_words = count_msword_en_words(" ".join(up_texts))
    rv_words = count_msword_en_words(" ".join(rv_texts))
    print(f"Step 1 Draft saved to: {out_path}")
    print(f"Provisional counts (pending human review): UP = {up_words} words, RV = {rv_words} words")
    if skipped_list:
        print(f"Skipped unmatched rows marked in Col 3: {len(skipped_list)}")


if __name__ == "__main__":
    main()
