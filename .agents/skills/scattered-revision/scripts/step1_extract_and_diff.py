import argparse
import difflib
import json
import re
from pathlib import Path

from docx_xml_utils import (
    NS,
    diff_tokens,
    extract_char_shading,
    get_cell_plain_text,
    load_config,
    load_docx_xml,
    segment_diff_into_sentences,
    sentence_unit_stats,
    split_into_bg_regions,
)


def normalize_cn(s: str) -> str:
    return re.sub(r"[\s\u3000]+", "", s)


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
        if p.name.startswith("~$"):
            continue
        rel_str = str(p.relative_to(library_dir))
        if any(kw in rel_str for kw in exclude_keywords):
            continue
        files.append(p)
    return files


CN_NUM_CHARS = "一二三四五六七八九十百零"


def _leaf_matches_stem(leaf: str, stem: str) -> bool:
    norm_leaf = normalize_cn(leaf)
    norm_stem = normalize_cn(stem)
    pos = norm_stem.find(norm_leaf)
    while pos != -1:
        end_pos = pos + len(norm_leaf)
        # Prevent '十' from matching prefix of '十一', '二十' from '二十一', etc.
        if norm_leaf and norm_leaf[-1] in CN_NUM_CHARS and end_pos < len(norm_stem) and norm_stem[end_pos] in CN_NUM_CHARS:
            pos = norm_stem.find(norm_leaf, pos + 1)
            continue
        return True
    return False


def match_library_file(title: str, lib_files: list, first_row_cn: str = "", src_xml_cache: dict = None):
    clean_title = title.strip().strip("《》")
    parts = [p.strip().strip("《》") for p in re.split(r"[・·]", clean_title) if p.strip()]
    leaf = parts[-1] if parts else clean_title

    # Also strip leading '附篇X　' if present (e.g. '附篇一　神的显现带来了新的时代' -> '神的显现带来了新的时代')
    sub_leaf_m = re.match(r"^附篇[一二三四五六七八九十百零]+[\s\u3000]+(.+)$", leaf)
    leaf_variants = [leaf]
    if sub_leaf_m:
        leaf_variants.append(sub_leaf_m.group(1).strip())

    # Narrow candidates by volume / section qualifiers when hierarchical '・' is present
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
            # Quick length filter before SequenceMatcher
            if abs(len(src_norm) - len(inp_norm)) / max(1, len(inp_norm)) < 0.35:
                score = difflib.SequenceMatcher(None, inp_norm, src_norm).ratio()
                if score > best_score:
                    best_score = score
                    best_row = tcs
                    best_idx = idx

    if best_score >= 0.75:
        return best_idx, best_row, round(best_score, 4)
    return -1, None, 0.0


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

    lib_files = index_library_files(lib_dir, cfg.get("exclude_library_dir_keywords", []))
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
    matched_lib_paths = []
    src_xml_cache = {}

    rv_threshold = cfg.get("rv_word_change_ratio_threshold", 0.42)

    for body_idx, ch in enumerate(body):
        local = ch.tag.split("}")[-1]
        if local == "p":
            txt = "".join(t.text or "" for t in ch.findall(".//w:t", NS)).strip()
            if txt:
                cur_title = txt
        elif local == "tbl":
            trs_in_tbl = ch.findall("w:tr", NS)
            first_cn = get_cell_plain_text(trs_in_tbl[0].findall("w:tc", NS)[0]) if trs_in_tbl else ""
            src_file = match_library_file(cur_title, lib_files, first_row_cn=first_cn, src_xml_cache=src_xml_cache) if cur_title else None
            if src_file is None:
                raise RuntimeError(f"Cannot find matching library .docx for title: {cur_title!r} in {lib_dir}")
            matched_lib_paths.append(src_file)
            if src_file not in src_xml_cache:
                s_root, _, _ = load_docx_xml(src_file)
                src_xml_cache[src_file] = s_root
            s_root = src_xml_cache[src_file]

            for r_idx, tr in enumerate(ch.findall("w:tr", NS)):
                tcs = tr.findall("w:tc", NS)
                inp_cn = get_cell_plain_text(tcs[0])
                inp_en, inp_fills = extract_char_shading(tcs[1])
                if not any_doc_shading:
                    # Document has no shading differentiation -> treat full text as active
                    inp_fills = ["NONE_FULL"] * len(inp_en)

                s_row_idx, s_tcs, match_conf = find_matching_source_row(inp_cn, s_root)
                if s_tcs is None:
                    raise RuntimeError(f"Cannot match Chinese paragraph in [{cur_title}] Row {r_idx}: {inp_cn[:50]!r}")

                src_en = get_cell_plain_text(s_tcs[1])
                src_fr = get_cell_plain_text(s_tcs[2])

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
                    "col3_preview": src_fr[:120],
                    "regions": region_entries,
                })
            tbl_counter += 1

    lang_code = args.lang.strip().upper() or detect_lang_code(lib_dir, matched_lib_paths, cfg)
    manifest = {
        "input_docx": str(inp_path),
        "library_dir": str(lib_dir),
        "lang_code": lang_code,
        "any_doc_shading": any_doc_shading,
        "rows": rows_manifest,
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
    print(f"Review items ({len(summary_items)} edited sentences across {len(rows_manifest)} rows) written to: {summary_path}")
    print(f"Detected language code: {lang_code}")


if __name__ == "__main__":
    main()
