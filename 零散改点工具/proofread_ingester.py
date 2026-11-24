"""
Scenario 3B: 通读定稿入库（独立模式 — Google Sheets / Excel -> 三语总库）

4-Step Pipeline:
  Step 1: Match chapter titles in the 3-language library (read-only!) and copy verbatim to `改前/`.
  Step 2: Open each link (strictly by `gid`), strip excess spaces, and validate edits:
            - minor word edits -> OK
            - large length gap or heavy lexical divergence -> verify semantic consistency with Chinese (AI if enabled, else flag for review)
            - trailing comma `,` changed to period `.` -> check next sentence starts with uppercase letter
            - original trailing punctuation lost -> alert
  Step 3: Copy `改前/` -> `改后/` (exact original filenames) and apply exact-match replacements ONLY:
            - if pre-edit (`目标语言`) differs from `改前` doc -> DO NOT replace, report mismatch
            - if replaced > 1 occurrences -> report exact replacement count
  Step 4: Merge all Part sheets in order and compare merged `定稿列` vs merged `改后` Col 3 AND
          merged `英文` vs merged `改后` Col 2; if any difference exists, output a track-changes revision `.docx`!
"""

from __future__ import annotations

from datetime import datetime
import difflib
import json
from pathlib import Path
import re
import shutil
import sys
from typing import Callable, Dict, List, Optional, Sequence, Tuple, Union
import xml.etree.ElementTree as ET


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
from library_ingester import (  # noqa: E402
    PROOFREAD_5COL_HEADERS,
    PROOFREAD_5COL_WIDTHS,
    _write_simple_report_docx,
    accept_all_revisions_in_cell_safe,
    evaluate_proofread_sheet_rows,
    make_proofread_5col_row,
    pick_proofread_new_text,
    render_diff_track_changes_into_cell,
    save_docx_without_track_revisions_flag,
    surgical_update_cell_text,
)
from pipeline import ensure_library_index  # noqa: E402
from sheets_reader import (  # noqa: E402
    ColumnMappingConfig,
    GoogleOAuthManager,
    ProofreadRow,
    clean_cell_spaces,
    load_proofread_table,
    parse_title_link_table,
)

W = docx_xml_utils.W
NS = docx_xml_utils.NS

_TRAILING_PUNCT_RE = re.compile(r'([.!?;:,。！？；：，…]+)["”’\'」』）\)]*$')
_COMMA_END_RE = re.compile(r'[,，]["”’\'」』）\)]*$')
_PERIOD_END_RE = re.compile(r'[.。]["”’\'」』）\)]*$')
_FIRST_LETTER_RE = re.compile(r'^[“"‘\'(\[（【\s]*([A-Za-zÀ-ÖØ-öø-ÿ])')


def _norm_ws_text(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").replace("\u00A0", " ")).strip()


def _match_library_file_by_title(
    title: str,
    lib_index: step1_extract_and_diff.LibraryIndex,
) -> Optional[Path]:
    """Match a chapter title to its exact `.docx` file inside the 3-language library."""
    raw_t = (title or "").strip()
    stripped_t = re.sub(r"(?i)(?:\s*[-—_]*\s*(?:part\s*\d+|p\d+|通读|定稿|副本|\(\d+\)|\（\d+\）))+$", "", raw_t).strip()
    for cand_title in dict.fromkeys([raw_t, stripped_t]):
        if not cand_title:
            continue
        rels = step1_extract_and_diff.match_title_to_index_rels(cand_title, lib_index)
        if rels:
            p = lib_index.resolve_file_path(rels[0])
            if p.is_file():
                return p.resolve()

    # Fallback: fuzzy match against all indexed file stems & row-0 Chinese titles
    clean_t = re.sub(r"[\s\u3000《》“”\"'_\-]+", "", stripped_t or raw_t).lower()
    if not clean_t:
        return None
    best_p: Optional[Path] = None
    best_score = 0.0
    rel_keys = list(getattr(lib_index, "files_map", {}).keys()) or list(lib_index.keys())
    rows_by_rel = getattr(lib_index, "rows_by_rel", {}) or {}
    for rel_p in rel_keys:
        stem = Path(rel_p).stem
        candidates_str = [stem]
        r_list = rows_by_rel.get(rel_p) or []
        if r_list:
            r0_cn = str(r_list[0].get("cn") or "").strip()
            if r0_cn and len(r0_cn) <= 80:
                candidates_str.append(r0_cn)
        for cs in candidates_str:
            clean_cs = re.sub(r"[\s\u3000《》“”\"'_\-]+", "", cs).lower()
            if not clean_cs:
                continue
            if len(clean_t) >= 4 and (clean_t in clean_cs or clean_cs in clean_t):
                cand = lib_index.resolve_file_path(rel_p)
                if cand.is_file():
                    return cand.resolve()
            ratio = difflib.SequenceMatcher(None, clean_t, clean_cs, autojunk=False).ratio()
            if ratio > best_score:
                best_score = ratio
                best_p = lib_index.resolve_file_path(rel_p)
    if best_p is not None and best_score >= 0.62 and best_p.is_file():
        return best_p.resolve()
    return None


def _check_semantic_consistency_with_ai(
    cn_text: str,
    old_target: str,
    new_target: str,
    api_key: Union[str, List[str], ApiKeyPool],
    model: str = "gemini-3.5-flash-lite",
    api_mode: str = "auto",
) -> Tuple[Optional[bool], str]:
    """
    Optionally ask Gemini whether `new_target` is semantically faithful to `cn_text`.
    Returns `(is_consistent_or_None, brief_explanation)`.
    """
    if not api_key:
        return None, "未配置 API Key，请人工核对改后译文与中文原文意思是否一致"
    prompt = (
        "你是多语种翻译质量核对助手。请判断【改后译文】的整体意思是否与【中文原文】一致（只检查有无重大漏译、错译、跑偏，不挑剔润色风格）。\n"
        f"【中文原文】：{cn_text}\n"
        f"【改前译文】：{old_target}\n"
        f"【改后译文】：{new_target}\n"
        '请仅返回 JSON：{"consistent": true或false, "reason": "一句话简明说明"}'
    )
    try:
        obj = call_gemini_json(
            api_key=api_key,
            model=model,
            system_instruction="你是多语种翻译质量核对助手，只返回严格 JSON。",
            user_prompt=prompt,
            api_mode=api_mode,
        )
        if isinstance(obj, dict):
            return bool(obj.get("consistent", True)), str(obj.get("reason") or "")
    except Exception as ex:
        return None, f"AI 核对跳过（{str(ex)[:60]}），请人工复核"
    return None, "请人工核对与中文原文意思是否一致"


def _validate_proofread_rows(
    chapter_title: str,
    rows: Sequence[ProofreadRow],
    enable_ai: bool = False,
    api_key: Union[str, List[str], ApiKeyPool] = "",
    model: str = "gemini-3.5-flash-lite",
    api_mode: str = "auto",
) -> Tuple[List[dict], List[Tuple[ProofreadRow, str]]]:
    """
    Step 2: Inspect all rows from one sheet/part:
      - Returns `(alerts_list, active_edits_list)` where `active_edits_list` is `[(row, new_text), ...]`.
    """
    alerts: List[dict] = []
    active_edits: List[Tuple[ProofreadRow, str]] = []

    for idx, r in enumerate(rows):
        old_t = clean_cell_spaces(r.target)
        new_t = pick_proofread_new_text(r)
        if not old_t or not new_t or old_t == new_t:
            continue

        active_edits.append((r, new_t))

        # 1. Trailing punctuation lost check
        m_old_p = _TRAILING_PUNCT_RE.search(old_t)
        m_new_p = _TRAILING_PUNCT_RE.search(new_t)
        if m_old_p and not m_new_p:
            alerts.append({
                "chapter": chapter_title,
                "sheet": r.sheet_name,
                "row": r.row_number,
                "type": "末尾标点丢失",
                "detail": f"改前末尾有标点「{m_old_p.group(1)}」，改后末尾标点缺失。",
                "old": old_t[:80],
                "new": new_t[:80],
            })

        # 2. Trailing comma -> period: verify next row's first letter is capitalized
        if _COMMA_END_RE.search(old_t) and _PERIOD_END_RE.search(new_t):
            if idx + 1 < len(rows):
                nxt = rows[idx + 1]
                nxt_txt = clean_cell_spaces(nxt.changes) or clean_cell_spaces(nxt.final) or clean_cell_spaces(nxt.target)
                m_fl = _FIRST_LETTER_RE.match(nxt_txt)
                if m_fl and m_fl.group(1).islower():
                    alerts.append({
                        "chapter": chapter_title,
                        "sheet": r.sheet_name,
                        "row": r.row_number,
                        "type": "句号后首字母未大写",
                        "detail": f"本行末尾由逗号改为句号，但下一行（第 {nxt.row_number} 行）首字母「{m_fl.group(1)}」仍为小写。",
                        "old": old_t[-40:],
                        "new": f"{new_t[-40:]} || 下一行：{nxt_txt[:40]}",
                    })

        # 3. Length gap or heavy lexical divergence -> verify semantic consistency with Chinese
        old_words = old_t.split()
        new_words = new_t.split()
        len_ratio = len(new_t) / max(1, len(old_t))
        word_sim = difflib.SequenceMatcher(None, [w.lower() for w in old_words], [w.lower() for w in new_words], autojunk=False).ratio()

        large_len_gap = (len(old_t) >= 25) and (len_ratio < 0.60 or len_ratio > 1.55)
        heavy_word_diff = (len(old_words) >= 5) and (word_sim < 0.55)

        if large_len_gap or heavy_word_diff:
            tag = "篇幅长度差距大" if large_len_gap else "大量字词变动"
            if enable_ai and api_key and r.cn:
                ok, reason = _check_semantic_consistency_with_ai(
                    cn_text=r.cn,
                    old_target=old_t,
                    new_target=new_t,
                    api_key=api_key,
                    model=model,
                    api_mode=api_mode,
                )
                if ok is False:
                    alerts.append({
                        "chapter": chapter_title,
                        "sheet": r.sheet_name,
                        "row": r.row_number,
                        "type": f"{tag}（AI核对不一致）",
                        "detail": reason or "改后译文与中文原文语义存在差异，请人工复核。",
                        "old": old_t[:80],
                        "new": new_t[:80],
                    })
                elif ok is None:
                    alerts.append({
                        "chapter": chapter_title,
                        "sheet": r.sheet_name,
                        "row": r.row_number,
                        "type": f"{tag}（需核对中文）",
                        "detail": reason,
                        "old": old_t[:80],
                        "new": new_t[:80],
                    })
            else:
                alerts.append({
                    "chapter": chapter_title,
                    "sheet": r.sheet_name,
                    "row": r.row_number,
                    "type": f"{tag}（需核对中文）",
                    "detail": f"长度比 {len_ratio:.2f}、词相似度 {word_sim:.2f}，请核对改后意思是否与中文一致。",
                    "old": old_t[:80],
                    "new": new_t[:80],
                })

    return alerts, active_edits


def _write_full_comparison_revision_docx(
    out_path: Path,
    chapter_title: str,
    en_table_merged: str,
    en_doc_merged: str,
    f_table_merged: str,
    f_doc_merged: str,
) -> Path:
    """Write a clean 2-section revision `.docx` showing exact diffs between merged sheet vs `改后` doc."""
    # Create a base report skeleton first, then populate diff cells with native track changes
    _write_simple_report_docx(
        out_path=out_path,
        title=f"全篇合并核对修订稿 — {chapter_title}",
        summary_lines=[
            "说明：以下将通读表格全部 Part 合并文本（原值）与「改后」文档对应列全量文本（新值）做逐字比对，差异以原生修订显示。",
        ],
        headers=["核对列", "合并比对修订痕迹（删除线=通读表合并内容，下划线=改后文档内容）"],
        rows_data=[
            ["第 2 列（英文全篇合并）", "一致" if _norm_ws_text(en_table_merged) == _norm_ws_text(en_doc_merged) else "存在差异"],
            ["第 3 列（目标语全篇合并）", "一致" if _norm_ws_text(f_table_merged) == _norm_ws_text(f_doc_merged) else "存在差异"],
        ],
    )
    root, _, ns_pairs = docx_xml_utils.load_docx_xml(out_path)
    trs = root.findall(".//w:tbl/w:tr", NS)
    rev_ctr = [300]
    if len(trs) >= 2 and _norm_ws_text(en_table_merged) != _norm_ws_text(en_doc_merged):
        tcs = trs[1].findall("w:tc", NS)
        if len(tcs) >= 2:
            render_diff_track_changes_into_cell(
                tcs[1], _norm_ws_text(en_table_merged), _norm_ws_text(en_doc_merged), author="全篇核对", rev_id_counter=rev_ctr
            )
    if len(trs) >= 3 and _norm_ws_text(f_table_merged) != _norm_ws_text(f_doc_merged):
        tcs = trs[2].findall("w:tc", NS)
        if len(tcs) >= 2:
            render_diff_track_changes_into_cell(
                tcs[1], _norm_ws_text(f_table_merged), _norm_ws_text(f_doc_merged), author="全篇核对", rev_id_counter=rev_ctr
            )
    docx_xml_utils.save_docx_xml(out_path, out_path, root, ns_pairs)
    return out_path


def run_proofread_ingestion_pipeline(
    title_links_text: str,
    library_dir: Union[str, Path],
    output_dir: Union[str, Path],
    col_config: Optional[ColumnMappingConfig] = None,
    oauth: Optional[GoogleOAuthManager] = None,
    enable_ai: bool = False,
    api_key: Union[str, List[str], ApiKeyPool] = "",
    model: str = "gemini-3.5-flash-lite",
    api_mode: str = "auto",
    date_str: str = "",
    log_fn: Optional[Callable[[str], None]] = None,
    progress_fn: Optional[Callable[[float, str], None]] = None,
) -> dict:
    """
    Execute Scenario 3B (`通读定稿入库`) 4-step workflow without mutating `library_dir`.
    """
    lib_p = Path(library_dir).resolve()
    out_d = Path(output_dir).resolve()
    if not lib_p.is_dir():
        raise FileNotFoundError(f"找不到三语总库目录：{lib_p}")
    out_d.mkdir(parents=True, exist_ok=True)

    chapter_groups = parse_title_link_table(title_links_text)
    if not chapter_groups:
        raise ValueError("请先在输入区粘贴「篇目标题 + 通读表格链接（或本地 .xlsx 路径）」。")

    eff_date = docx_xml_utils.sanitize_filename_part(
        (date_str or "").strip() or datetime.now().strftime("%Y%m%d"),
        fallback=datetime.now().strftime("%Y%m%d"),
        max_len=16,
    )

    if progress_fn:
        progress_fn(3.0, "正在加载三语总库索引...")
    ensure_library_index(str(lib_p), log_fn=log_fn)
    lib_index = library_index.get_library_index(lib_p)

    before_dir = out_d / "改前"
    after_dir = out_d / "改后"
    verify_dir = out_d / "全篇核对修订稿"
    before_dir.mkdir(parents=True, exist_ok=True)
    after_dir.mkdir(parents=True, exist_ok=True)

    all_qa_alerts: List[dict] = []
    all_replace_logs: List[dict] = []
    five_col_report_rows: List[List[str]] = []
    verify_rev_files: List[str] = []
    unmatched_titles: List[str] = []
    processed_chapters = 0

    def _infer_library_files_from_rows(rows: Sequence[ProofreadRow]) -> List[Path]:
        """When a local sheet is dragged in without an explicit chapter title, infer matched library chapters from Chinese rows."""
        found: List[Path] = []
        seen = set()
        dummy_tc = ET.Element(f"{W}tc")
        for pr in rows:
            cn_s = (pr.cn or "").strip()
            if not cn_s or len(cn_s) < 6:
                # Could be a short section/chapter title row (e.g. '独一无二的神自己 二') when en & target are empty
                if cn_s and not (pr.en or "").strip() and not (pr.target or "").strip():
                    by_t = _match_library_file_by_title(cn_s, lib_index)
                    if by_t is not None and by_t not in seen:
                        seen.add(by_t)
                        found.append(by_t)
                continue
            if not (pr.en or "").strip() and not (pr.target or "").strip():
                by_t = _match_library_file_by_title(cn_s, lib_index)
                if by_t is not None and by_t not in seen:
                    seen.add(by_t)
                    found.append(by_t)
                continue
            # Only probe rows that actually carry edits (or first few rows if none edited yet)
            if pick_proofread_new_text(pr) or len(found) == 0:
                m_row, _, _ = step1_extract_and_diff.search_library_row_5stage(
                    cur_title="",
                    tc1=dummy_tc,
                    tc2=dummy_tc,
                    inp_cn=cn_s,
                    inp_en=(pr.en or "").strip(),
                    lib_index=lib_index,
                )
                if m_row is not None:
                    lp = lib_index.resolve_file_path(m_row["rel_path"]).resolve()
                    if lp.is_file() and lp not in seen:
                        seen.add(lp)
                        found.append(lp)
        return found

    for g_idx, (ch_title, sources) in enumerate(chapter_groups):
        pct = 8.0 + (g_idx / max(1, len(chapter_groups))) * 85.0
        if progress_fn:
            progress_fn(pct, f"正在处理 [{g_idx + 1}/{len(chapter_groups)}]：{ch_title}（共 {len(sources)} 个表格）")
        if log_fn:
            log_fn(f"\n📑 === [{g_idx + 1}/{len(chapter_groups)}] 通读定稿入库：《{ch_title}》（{len(sources)} 个链接/表格） ===")

        # Read linked sheets first (strictly by gid) & validate edits
        chapter_edits: List[Tuple[ProofreadRow, str]] = []
        all_group_rows: List[ProofreadRow] = []
        merged_sheet_en_parts: List[str] = []
        merged_sheet_final_parts: List[str] = []

        for s_idx, src in enumerate(sources):
            try:
                parsed = load_proofread_table(src, col_config=col_config, oauth=oauth)
                if log_fn:
                    log_fn(f"  [读取表格] 第 {s_idx + 1}/{len(sources)} 个表格（{parsed.sheet_name}，共 {len(parsed.rows)} 行）")
            except Exception as ex:
                all_qa_alerts.append({
                    "chapter": ch_title,
                    "sheet": src[:60],
                    "row": 0,
                    "type": "表格读取失败",
                    "detail": str(ex),
                    "old": "",
                    "new": "",
                })
                if log_fn:
                    log_fn(f"  ❌ 读取表格失败：{src} -> {ex}")
                continue

            all_group_rows.extend(parsed.rows)
            rep_items, blocked_5col = evaluate_proofread_sheet_rows(
                parsed=parsed,
                chapter_label=ch_title,
                enable_ai=enable_ai,
                api_key=api_key,
                model=model,
                api_mode=api_mode,
            )
            five_col_report_rows.extend(blocked_5col)
            for b_row in blocked_5col:
                all_qa_alerts.append({
                    "chapter": ch_title,
                    "sheet": parsed.sheet_name,
                    "row": 0,
                    "type": "列一致性校验未通过",
                    "detail": b_row[1],
                })

            accepted_by_row_id: Dict[int, str] = {}
            for item in rep_items:
                pr_obj = item["row"]
                new_t = item["new_text"]
                warns = item["warnings"]
                chapter_edits.append((pr_obj, new_t))
                accepted_by_row_id[id(pr_obj)] = new_t
                if warns:
                    warn_joined = "；".join(warns)
                    five_col_report_rows.append(
                        make_proofread_5col_row(item["source_pos"], "改后异常提醒：" + warn_joined, pr_obj)
                    )
                    all_qa_alerts.append({
                        "chapter": ch_title,
                        "sheet": parsed.sheet_name,
                        "row": pr_obj.row_number,
                        "type": "改前改后对比提醒",
                        "detail": warn_joined,
                    })

            for pr in parsed.rows:
                en_clean = clean_cell_spaces(pr.en)
                if en_clean:
                    merged_sheet_en_parts.append(en_clean)
                # Only use accepted new text if row passed `改动列 == 定稿列` consistency check; otherwise keep `target`
                fin_clean = accepted_by_row_id.get(id(pr)) or clean_cell_spaces(pr.target)
                if fin_clean:
                    merged_sheet_final_parts.append(fin_clean)

        # Match library chapter by title (fallback to inferring chapters from sheet Chinese rows ONLY when
        # the title is a default dragged file stem or placeholder, so explicit unknown titles are reported!)
        explicit_lib_file = _match_library_file_by_title(ch_title, lib_index)
        first_src_stem = Path(sources[0]).stem if (sources and not sources[0].startswith(("http://", "https://"))) else ""
        allow_row_inference = (ch_title in ("", "未命名篇目")) or (first_src_stem and ch_title.strip() == first_src_stem.strip())
        target_lib_files: List[Path] = (
            [explicit_lib_file]
            if explicit_lib_file is not None
            else (_infer_library_files_from_rows(all_group_rows) if allow_row_inference else [])
        )
        is_single_chapter_sheet = explicit_lib_file is not None or len(target_lib_files) == 1

        if not target_lib_files:
            unmatched_titles.append(ch_title)
            if log_fn:
                log_fn(f"⚠️ 未在总库找到篇目《{ch_title}》，已跳过。")
            continue

        remaining_unmatched_edits: Dict[int, Tuple[ProofreadRow, str]] = {
            id(pr): (pr, new_text) for pr, new_text in chapter_edits
        }

        for lib_file in target_lib_files:
            before_copy = docx_xml_utils.safe_join_output_file(before_dir, lib_file.name)
            after_copy = docx_xml_utils.safe_join_output_file(after_dir, lib_file.name)
            docx_xml_utils.slim_docx_file(lib_file, before_copy)
            if log_fn:
                log_fn(f"  [步骤 1] 已提取改前文档：{lib_file.name}")

            # Step 3: Exact-match replacement into `改后/` (do NOT replace if pre-edit differs; report if >1 hits)
            root_after, _, ns_after = docx_xml_utils.load_docx_xml(before_copy)
            cmts_after = docx_xml_utils.load_docx_comments_xml(before_copy)
            trs_after = [tr for tr in root_after.findall(".//w:tbl/w:tr", NS) if len(tr.findall("w:tc", NS)) >= 3]

            replaced_ok = 0
            multi_hit_cnt = 0

            for pr, new_text in chapter_edits:
                old_exact = clean_cell_spaces(pr.target)
                hits: List[Tuple[int, ET.Element, str]] = []
                for r_i, tr in enumerate(trs_after):
                    tc3 = tr.findall("w:tc", NS)[2]
                    accept_all_revisions_in_cell_safe(tc3)
                    cell_txt = docx_xml_utils.get_cell_plain_text(tc3).replace("\u00A0", " ")
                    if old_exact in cell_txt:
                        hits.append((r_i, tc3, cell_txt))

                if not hits:
                    continue

                remaining_unmatched_edits.pop(id(pr), None)
                total_occ = sum(cell_txt.count(old_exact) for _, _, cell_txt in hits)
                for _, tc3, cell_txt in hits:
                    updated_txt = cell_txt.replace(old_exact, new_text)
                    surgical_update_cell_text(tc3, updated_txt)
                replaced_ok += 1

                if total_occ > 1:
                    multi_hit_cnt += 1
                    all_replace_logs.append({
                        "chapter": lib_file.stem if not is_single_chapter_sheet else ch_title,
                        "row": pr.row_number,
                        "status": f"多处替换提示（共替换 {total_occ} 处）",
                        "hits": total_occ,
                        "detail": f"改前句在文档中匹配并替换了 {total_occ} 处，请核对是否均需替换：{old_exact[:60]}",
                    })

            save_docx_without_track_revisions_flag(before_copy, after_copy, root_after, ns_after, comments_root=cmts_after)
            processed_chapters += 1
            if log_fn:
                log_fn(
                    f"  [步骤 3] {lib_file.name} 精确入库完成：成功替换 {replaced_ok} 条，多处替换提示 {multi_hit_cnt} 条。"
                )

            # Step 4: Full merged comparison (for single-chapter multi-Part proofread tables)
            if is_single_chapter_sheet:
                doc_en_parts: List[str] = []
                doc_f_parts: List[str] = []
                for tr in trs_after:
                    tcs = tr.findall("w:tc", NS)
                    c2_t = clean_cell_spaces(docx_xml_utils.get_cell_plain_text(tcs[1]))
                    c3_t = clean_cell_spaces(docx_xml_utils.get_cell_plain_text(tcs[2]))
                    if c2_t:
                        doc_en_parts.append(c2_t)
                    if c3_t:
                        doc_f_parts.append(c3_t)

                sheet_en_all = _norm_ws_text(" ".join(merged_sheet_en_parts))
                doc_en_all = _norm_ws_text(" ".join(doc_en_parts))
                sheet_f_all = _norm_ws_text(" ".join(merged_sheet_final_parts))
                doc_f_all = _norm_ws_text(" ".join(doc_f_parts))

                en_differs = bool(sheet_en_all and sheet_en_all != doc_en_all)
                f_differs = bool(sheet_f_all and sheet_f_all != doc_f_all)
                if en_differs or f_differs:
                    verify_dir.mkdir(parents=True, exist_ok=True)
                    safe_ch = docx_xml_utils.sanitize_filename_part(ch_title, fallback="篇目", max_len=90)
                    rev_p = docx_xml_utils.safe_join_output_file(verify_dir, f"全篇合并核对修订稿_{safe_ch}_{eff_date}.docx")
                    _write_full_comparison_revision_docx(
                        out_path=rev_p,
                        chapter_title=ch_title,
                        en_table_merged=sheet_en_all,
                        en_doc_merged=doc_en_all,
                        f_table_merged=sheet_f_all,
                        f_doc_merged=doc_f_all,
                    )
                    verify_rev_files.append(str(rev_p))
                    if log_fn:
                        diff_cols = "、".join([x for x, flg in (("英文", en_differs), ("目标语", f_differs)) if flg])
                        log_fn(f"  ⚠️ [步骤 4] 全篇合并核对发现差异（{diff_cols}），已输出修订稿：{rev_p.name}")
                else:
                    if log_fn:
                        log_fn("  ✅ [步骤 4] 全篇合并核对完全一致（英文与目标语 0 差异）。")

        if remaining_unmatched_edits:
            if log_fn:
                log_fn(f"  ⚠️ [步骤 3] 改前不一致未替换：{len(remaining_unmatched_edits)} 条（详见检查报告 5 列明细）。")
            for pr, _ in remaining_unmatched_edits.values():
                old_exact = clean_cell_spaces(pr.target)
                pos_lbl = f"{ch_title} [{pr.sheet_name or Path(pr.source_label).name}] 第{pr.row_number}行"
                five_col_report_rows.append(
                    make_proofread_5col_row(pos_lbl, "未替换：改前内容与文档不完全一致", pr)
                )
                all_replace_logs.append({
                    "chapter": ch_title,
                    "row": pr.row_number,
                    "status": "改前不一致（未替换）",
                    "hits": 0,
                    "detail": f"表格改前句在总库改前文档中未找到完全一致文本：{old_exact[:70]}",
                })

    # Write combined QA + Replacement Report (including full 5-column detail table!)
    rep_rows = []
    for rl in all_replace_logs:
        rep_rows.append([rl["chapter"], f"第{rl['row']}行", rl["status"], rl["detail"]])
    for ut in unmatched_titles:
        rep_rows.append([ut, "-", "总库未匹配篇目", "未在总库目录中定位到该篇目文件"])
    if not rep_rows and all_qa_alerts:
        for al in all_qa_alerts:
            rep_rows.append([al["chapter"], f"{al['sheet']} 第{al['row']}行" if al.get("row") else al["sheet"], al["type"], al["detail"]])

    extra_secs = []
    if five_col_report_rows:
        extra_secs.append((
            f"二、未替换 / 列不一致 / 改后异常 5 列完整明细（共 {len(five_col_report_rows)} 条）",
            PROOFREAD_5COL_HEADERS,
            five_col_report_rows,
            PROOFREAD_5COL_WIDTHS,
        ))

    report_path = docx_xml_utils.safe_join_output_file(out_d, f"通读定稿入库检查报告_{eff_date}.docx")
    _write_simple_report_docx(
        out_path=report_path,
        title=f"通读定稿入库检查与替换报告（{eff_date}）",
        summary_lines=[
            f"输入篇目数：{len(chapter_groups)} 篇 | 成功处理：{processed_chapters} 篇 | 未匹配篇目：{len(unmatched_titles)} 篇",
            f"步骤 2 列一致性校验及改前改后对比提醒：{len(all_qa_alerts)} 条",
            f"步骤 3 改前不一致未替换 / 多处替换提醒：{len(all_replace_logs)} 条",
            f"步骤 4 全篇合并核对差异修订稿：{len(verify_rev_files)} 份",
        ],
        headers=["篇目", "位置", "检查项 / 状态", "详细说明"],
        rows_data=rep_rows or [["全部通过", "-", "正常", "所有通读改点检查、精确替换及全篇合并核对均无异常。"]],
        col_widths=[3200, 2200, 2600, 6570],
        extra_sections=extra_secs,
    )

    if progress_fn:
        progress_fn(100.0, f"通读定稿入库完成：共处理 {processed_chapters} 篇！")
    if log_fn:
        log_fn(f"\n✅ 通读定稿入库全部完成！报告已保存至：{report_path.name}")

    return {
        "before_dir": str(before_dir),
        "after_dir": str(after_dir),
        "report_path": str(report_path),
        "verify_revision_files": verify_rev_files,
        "processed_chapters": processed_chapters,
        "unmatched_titles": unmatched_titles,
        "qa_alerts_count": len(all_qa_alerts),
        "replace_alerts_count": len(all_replace_logs),
    }
