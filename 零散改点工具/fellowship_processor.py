"""Scenario 4 — Brother Fellowship Manuscript Processor (弟兄交通稿件整理).

Capabilities:
1. Safe XML Pre-cleaning (Zero Data Loss):
   - Unwraps all <w:sdt> content controls first so inner text/runs are never lost.
   - Turns OFF Word Track Changes mode (<w:trackRevisions/> removed in word/settings.xml -> '修订: 关闭').
   - Accepts all remaining <w:ins>/<w:del>/*Change revisions.
   - Removes all review comments (commentRangeStart/End, commentReference, word/comments*.xml) and stray highlights.
2. Header & Title Normalization:
   - Strips 'Notice from the Church...' / translator remarks above the main table.
   - Extracts the top title line, normalizes part numbers ('Part 1' / '-1' -> '-01') for single-part outputs,
     and strips the part suffix for merged outputs.
3. Smart Quote Normalization (Col 2 & Col 3):
   - Converts straight quotes (" and ') to curly quotes (“ ” and ‘ ’), handling mid-word apostrophes (’) cleanly,
     with per-language skip support (e.g. French « »).
4. Two-Tier Col 3 Blue (#0000FF) & Bold (<w:b/>) Restoration:
   - Tier 1 (Deterministic, 0 API calls): Main title row, date/speaker row, Question X headings, whole-blue closing
     rows, and 'Answer:' / '解答：' prefixes.
   - Tier 2 (AI Semantic Span Alignment via ApiKeyPool): Intra-paragraph scripture & God's Word quotes; excludes
     outer quotation marks & citation brackets from bold while preserving bold on internal sentence punctuation.
5. Typography, Spacing & Content-Proportional Column Widths:
   - Main Title Row (Row 0): preserves 14pt (四号, w:sz=28).
   - Body Text: 12pt (小四, w:sz=24).
   - Post-Quote Citations ('” (Filipenses 1:21).', '«...» (Առակներ 10։21)։', '„...“ (Лука 9:50).', etc.):
     balanced-bracket matcher sets the bracket + inner text to 10.5pt (五号, w:sz=21) while keeping trailing
     external punctuation (., 。, ։) at 12pt (小四, w:sz=24), and excluding narrative book titles & '(Yes.)'/'(pl)'.
   - Table paragraphs: exact 20pt line spacing (w:line=400, w:lineRule=exact), before=0, after=0.
   - 3-column widths dynamically balanced by visual character length so all 3 columns have comparable height.
6. Three Output Modes:
   - 'single_only': Clean & format each part individually (with -01..-0N suffix).
   - 'merged_only': Merge all parts in natural order into 1 unified document (without -01 suffix).
   - 'both': Output both individual formatted parts and the merged document.
"""

import copy
from datetime import datetime
import json
import os
from pathlib import Path
import re
import tempfile
from typing import Callable, Dict, List, Optional, Set, Tuple
import xml.etree.ElementTree as ET
import zipfile

from ai_client import ApiKeyPool, call_gemini_json
from lang_registry import detect_lang_code_from_hint, get_tri_prefix, load_languages


MAX_DOCX_ENTRIES = 5000
MAX_DOCX_UNCOMPRESSED_BYTES = 300 * 1024 * 1024


def validate_docx_archive(docx_path: Path) -> None:
    """Validate .docx zip archive against zip-bomb / path-traversal / missing document.xml."""
    p = Path(docx_path)
    if not p.is_file():
        raise FileNotFoundError(f"未找到文档：{p}")
    with zipfile.ZipFile(p, "r") as zf:
        infos = zf.infolist()
        if len(infos) > MAX_DOCX_ENTRIES:
            raise ValueError(f"文档内部条目数异常 ({len(infos)})：{p.name}")
        total_uncompressed = 0
        has_doc_xml = False
        for info in infos:
            fn = info.filename.replace("\\", "/")
            if fn.startswith("/") or ".." in fn.split("/"):
                raise ValueError(f"文档包含不安全的归档路径：{info.filename}")
            total_uncompressed += max(0, info.file_size)
            if total_uncompressed > MAX_DOCX_UNCOMPRESSED_BYTES:
                raise ValueError(f"文档解压体积超过安全上限 (300MB)：{p.name}")
            if fn == "word/document.xml":
                has_doc_xml = True
        if not has_doc_xml:
            raise ValueError(f"不是有效的 Word (.docx) 文档（缺少 word/document.xml）：{p.name}")


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_NS = "http://www.w3.org/XML/1998/namespace"

W = f"{{{W_NS}}}"
XML_SPACE = f"{{{XML_NS}}}space"

COMMON_NAMESPACES = {
    "w": W_NS,
    "r": R_NS,
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
}
for _pfx, _uri in COMMON_NAMESPACES.items():
    ET.register_namespace(_pfx, _uri)


OUTPUT_MODES = [
    ("仅输出各单篇（默认，不合并为一个总文档）", "single_only"),
    ("仅输出汇总合并总文档（按顺序合并为 1 份，去掉序号）", "merged_only"),
    ("各单篇 + 汇总合并总文档同时输出", "both"),
]

QUOTE_MODES = [
    ("自动规范化为弯引号 “” ‘’（法语保留 « »）", "auto"),
    ("强制统一替换为弯引号 “” ‘’", "curly"),
    ("保持原文档引号不替换", "skip"),
]

CLOSING_QUOTE_CHARS: Set[str] = {
    "”",  # U+201D right double quotation mark
    '"',  # U+0022 straight double quote
    "’",  # U+2019 right single quotation mark
    "'",  # U+0027 straight single quote
    "»",  # U+00BB right-pointing double angle quotation mark
    "“",  # U+201C left double quotation mark (used as CLOSING quote in German/Bulgarian/Czech „...“)
    "‟",  # U+201F double high-reversed-9 quotation mark
    "›",  # U+203A single right-pointing angle quotation mark
    "」",  # U+300D CJK corner bracket
    "』",  # U+300F CJK white corner bracket
}

ALL_OUTER_QUOTE_CHARS: Set[str] = CLOSING_QUOTE_CHARS | {
    "„",  # U+201E double low-9 quotation mark
    "«",  # U+00AB left-pointing double angle quotation mark
    "‘",  # U+2018 left single quotation mark
    "‚",  # U+201A single low-9 quotation mark
    "‹",  # U+2039 single left-pointing angle quotation mark
    "「",  # U+300C
    "『",  # U+300E
}

BRACKET_PAIRS: Dict[str, str] = {
    "(": ")",
    "（": "）",
    "[": "]",
    "［": "］",
    "〔": "〕",
    "【": "】",
}

BOILERPLATE_HEADER_PATTERNS = [
    re.compile(r"notice\s+from\s+the\s+church", re.I),
    re.compile(r"three\s+types\s+of\s+[“\"']?you", re.I),
    re.compile(r"you\(pl\)", re.I),
    re.compile(r"you\(s\)", re.I),
    re.compile(r"differences\s+between\s+english\s+and\s+chinese", re.I),
    re.compile(r"proper\s+grammar\s+in\s+your\s+language", re.I),
    re.compile(r"no\s+need\s+to\s+review\s+the\s+part\s+in\s+blue\s+font", re.I),
    re.compile(r"^(?:译者备注|审校说明|交接说明|翻译说明)[:：]"),
]


def _local_tag(el: ET.Element) -> str:
    t = el.tag
    return t.split("}", 1)[1] if "}" in t else t


def _get_p_text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.findall(f".//{W}t"))


def _get_tc_text(tc: ET.Element) -> str:
    return "\n".join(_get_p_text(p) for p in tc.findall(f"{W}p"))


def sanitize_filename_part(s: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', " ", (s or "").strip())
    cleaned = re.sub(r"\s+", " ", cleaned).strip().rstrip(".")
    return cleaned or "弟兄交通"


# ---------------- Part Number & Title Parsing ----------------

_TITLE_PART_WORD_RE = re.compile(r"[\s\-—–_]*(?:[Pp]art|[Pp]t\.?)\s*0*(\d{1,3})\s*$")
_TITLE_DASH_NUM_RE = re.compile(r"(?<!\d\d\d\d)[\-—–_]\s*0*(\d{1,3})\s*$")
_FILE_PART_WORD_RE = re.compile(r"(?:[Pp]art|[Pp]t\.?)\s*0*(\d{1,3})(?!\d)")
_FILE_DASH_NUM_RE = re.compile(r"[\-—–_]0*(\d{1,2})(?=[\-—–_\s(.]|$)")


def extract_part_num_from_filename(filename: str) -> Optional[int]:
    stem = Path(filename).stem
    m = _FILE_PART_WORD_RE.search(stem)
    if m:
        return int(m.group(1))
    # Remove 8-digit dates like 20261115 before searching -01
    stem_no_date = re.sub(r"20\d{6}", "", stem)
    # Remove Chinese dates like 2026年11月2日 so we match -01 after it
    m2 = re.search(r"日[\s\-—–_]*0*(\d{1,2})(?=[\-—–_\s]|$)", stem_no_date)
    if m2:
        return int(m2.group(1))
    matches = list(_FILE_DASH_NUM_RE.finditer(stem_no_date))
    if matches:
        # Take the last 2-digit segment marker if not at the very beginning (02_How to...)
        for cand in reversed(matches):
            if cand.start() > 3:
                return int(cand.group(1))
    return None


def parse_title_and_part(raw_title: str, fallback_filename: str = "") -> Tuple[str, Optional[int]]:
    """Parse top title line into (base_title_without_part, part_number_int_or_None).

    Examples:
      '《怎样实行真理（二）》-上面的交通 2026年11月2日-01' -> ('《怎样实行真理（二）》-上面的交通 2026年11月2日', 1)
      '怎样分辨各种人（二）Part 1' -> ('怎样分辨各种人（二）', 1)
    """
    t = re.sub(r"\s+", " ", (raw_title or "").strip())
    # Strip trailing '-三语稿-20261115' if re-processing an already named title
    t = re.sub(r"[\-—–_]三语稿(?:[\-—–_]20\d{6})?\s*$", "", t).strip()

    m1 = _TITLE_PART_WORD_RE.search(t)
    if m1:
        base = t[: m1.start()].strip().rstrip("-—–_")
        return base, int(m1.group(1))

    # Match trailing '-01' (making sure we don't strip '-2' in '2026-11-02' unless preceded by '日' or non-date)
    m2 = re.search(r"(.*?)(?:[\s]*[\-—–_]\s*0*(\d{1,2}))\s*$", t)
    if m2:
        prefix = m2.group(1).strip()
        # Guard against ISO date YYYY-MM-DD at end
        if not re.search(r"20\d{2}[\-—–_]\d{1,2}$", prefix):
            return prefix, int(m2.group(2))

    fn_part = extract_part_num_from_filename(fallback_filename) if fallback_filename else None
    return t, fn_part


def natural_sort_docx_paths(paths: List[Path]) -> List[Path]:
    def _key(p: Path):
        pnum = extract_part_num_from_filename(p.name)
        nat_tokens = [
            int(tok) if tok.isdigit() else tok.lower()
            for tok in re.split(r"(\d+)", p.stem)
        ]
        return (0 if pnum is not None else 1, pnum if pnum is not None else 999999, nat_tokens)

    return sorted(paths, key=_key)


# ---------------- Step 1: XML Pre-Cleaning (SDTs, Track Changes, Comments, Highlights) ----------------

def unwrap_all_sdts(root: ET.Element) -> int:
    """Safely unwrap all <w:sdt> content controls innermost-first without losing a single run or paragraph."""
    unwrapped = 0
    while True:
        parent_map = {child: parent for parent in root.iter() for child in parent}
        sdts = root.findall(f".//{W}sdt")
        if not sdts:
            break
        # Pick an innermost <w:sdt> (one that contains no nested <w:sdt>)
        target_sdt = None
        for s in sdts:
            if not s.findall(f".//{W}sdt"):
                target_sdt = s
                break
        if target_sdt is None:
            target_sdt = sdts[-1]

        parent = parent_map.get(target_sdt)
        if parent is None:
            break
        siblings = list(parent)
        idx = siblings.index(target_sdt)
        sdt_content = target_sdt.find(f"{W}sdtContent")

        parent_tag = _local_tag(parent)
        children_to_hoist: List[ET.Element] = []
        if sdt_content is not None:
            raw_children = list(sdt_content)
            # Edge case: block container (body/tc) holding bare <w:r> inside sdtContent
            if parent_tag in ("body", "tc") and raw_children and all(_local_tag(c) == "r" for c in raw_children):
                wrap_p = ET.Element(f"{W}p")
                for rc in raw_children:
                    wrap_p.append(rc)
                children_to_hoist = [wrap_p]
            # Edge case: inline container (<w:p>) holding <w:p> inside sdtContent
            elif parent_tag == "p" and any(_local_tag(c) == "p" for c in raw_children):
                for c in raw_children:
                    if _local_tag(c) == "p":
                        for sub in list(c):
                            if _local_tag(sub) != "pPr":
                                children_to_hoist.append(sub)
                    else:
                        children_to_hoist.append(c)
            else:
                children_to_hoist = raw_children

        parent.remove(target_sdt)
        for offset, ch in enumerate(children_to_hoist):
            parent.insert(idx + offset, ch)
        unwrapped += 1
    return unwrapped


def accept_revisions_and_strip_comments(root: ET.Element) -> Tuple[int, int]:
    """Accept all <w:ins>/<w:del>/*Change revisions and strip all comment markers & highlights."""
    rev_count = 0
    cmt_count = 0

    DROP_TAGS = {
        "del",
        "moveFrom",
        "moveFromRangeStart",
        "moveFromRangeEnd",
        "moveToRangeStart",
        "moveToRangeEnd",
        "rPrChange",
        "pPrChange",
        "tblPrChange",
        "tblGridChange",
        "trPrChange",
        "tcPrChange",
        "sectPrChange",
    }
    UNWRAP_TAGS = {"ins", "moveTo"}
    COMMENT_TAGS = {"commentRangeStart", "commentRangeEnd", "commentReference"}

    changed = True
    while changed:
        changed = False
        for parent in list(root.iter()):
            for child in list(parent):
                lt = _local_tag(child)
                if lt in DROP_TAGS:
                    parent.remove(child)
                    rev_count += 1
                    changed = True
                elif lt in UNWRAP_TAGS:
                    idx = list(parent).index(child)
                    hoist = list(child)
                    parent.remove(child)
                    for off, sub in enumerate(hoist):
                        parent.insert(idx + off, sub)
                    rev_count += 1
                    changed = True
                elif lt in COMMENT_TAGS:
                    parent.remove(child)
                    cmt_count += 1
                    changed = True

    # Remove empty comment-reference wrapper runs and clear stray highlights / paste shadings
    for p in root.findall(f".//{W}p"):
        ppr = p.find(f"{W}pPr")
        if ppr is not None:
            pr_rpr = ppr.find(f"{W}rPr")
            if pr_rpr is not None:
                for hl in list(pr_rpr.findall(f"{W}highlight")):
                    pr_rpr.remove(hl)
            p_shd = ppr.find(f"{W}shd")
            if p_shd is not None:
                ppr.remove(p_shd)

        for r in list(p.findall(f"{W}r")):
            rpr = r.find(f"{W}rPr")
            if rpr is not None:
                for hl in list(rpr.findall(f"{W}highlight")):
                    rpr.remove(hl)
                r_shd = rpr.find(f"{W}shd")
                if r_shd is not None:
                    rpr.remove(r_shd)
                rstyle = rpr.find(f"{W}rStyle")
                if rstyle is not None and "CommentReference" in rstyle.attrib.get(f"{W}val", ""):
                    has_content = any(
                        _local_tag(c) not in ("rPr",) for c in r
                    )
                    if not has_content:
                        p.remove(r)
                        continue

    # Clean table cell background shadings (e.g. 'fcfcfc' gray web-paste background)
    for tcPr in root.findall(f".//{W}tcPr"):
        for shd in list(tcPr.findall(f"{W}shd")):
            tcPr.remove(shd)

    return rev_count, cmt_count


def disable_track_revisions_in_settings_xml(settings_bytes: bytes) -> bytes:
    """Remove <w:trackRevisions/> from word/settings.xml so Word opens with '修订: 关闭'."""
    try:
        s_root = ET.fromstring(settings_bytes)
        removed = False
        for tr_el in list(s_root.findall(f".//{W}trackRevisions")):
            # Find direct parent and remove
            for parent in s_root.iter():
                if tr_el in list(parent):
                    parent.remove(tr_el)
                    removed = True
                    break
        if removed:
            return ET.tostring(s_root, encoding="utf-8", xml_declaration=True)
        return settings_bytes
    except Exception:
        return settings_bytes


def strip_comment_rels_bytes(rels_bytes: bytes) -> bytes:
    try:
        ET.register_namespace("", REL_PKG_NS)
        r_root = ET.fromstring(rels_bytes)
        changed = False
        for rel in list(r_root):
            target = rel.attrib.get("Target", "").lower()
            rtype = rel.attrib.get("Type", "").lower()
            if "comment" in target or "comment" in rtype:
                r_root.remove(rel)
                changed = True
        if changed:
            return ET.tostring(r_root, encoding="utf-8", xml_declaration=True)
        return rels_bytes
    except Exception:
        return rels_bytes


def strip_comment_content_types_bytes(ct_bytes: bytes) -> bytes:
    try:
        ET.register_namespace("", CT_NS)
        c_root = ET.fromstring(ct_bytes)
        changed = False
        for ov in list(c_root):
            pn = ov.attrib.get("PartName", "").lower()
            if "/word/comments" in pn:
                c_root.remove(ov)
                changed = True
        if changed:
            return ET.tostring(c_root, encoding="utf-8", xml_declaration=True)
        return ct_bytes
    except Exception:
        return ct_bytes


# ---------------- Step 2: Top Boilerplate Cleanup & Title Line Normalization ----------------

def _is_boilerplate_header_paragraph(text: str) -> bool:
    s = text.strip()
    if not s:
        return True
    for pat in BOILERPLATE_HEADER_PATTERNS:
        if pat.search(s):
            return True
    return False


def clean_body_header_and_extract_title(
    root: ET.Element, fallback_filename: str = ""
) -> Tuple[str, Optional[int], ET.Element]:
    """Clean all boilerplate above first <w:tbl>, returning (base_title, part_num, kept_title_p)."""
    body = root.find(f"{W}body")
    if body is None:
        raise ValueError("Invalid Word document: missing <w:body>")

    children = list(body)
    first_tbl_idx = None
    for i, el in enumerate(children):
        if _local_tag(el) == "tbl":
            first_tbl_idx = i
            break

    if first_tbl_idx is None:
        raise ValueError(f"文档中未找到三列表格：{fallback_filename}")

    pre_ps = [el for el in children[:first_tbl_idx] if _local_tag(el) == "p"]
    candidate_ps = [p for p in pre_ps if not _is_boilerplate_header_paragraph(_get_p_text(p))]

    kept_p: Optional[ET.Element] = None
    raw_title = ""
    if candidate_ps:
        kept_p = candidate_ps[-1]
        raw_title = _get_p_text(kept_p).strip()

    # Fallback: construct title from Table Row 0 / Row 1 Col 1 if no header title paragraph existed
    tbl = children[first_tbl_idx]
    rows = tbl.findall(f"{W}tr")
    if not raw_title and rows:
        tcs0 = rows[0].findall(f"{W}tc")
        r0_c1 = _get_tc_text(tcs0[0]).strip() if tcs0 else ""
        r1_c1 = ""
        if len(rows) > 1:
            tcs1 = rows[1].findall(f"{W}tc")
            cand_r1 = _get_tc_text(tcs1[0]).strip() if tcs1 else ""
            if len(cand_r1) <= 50 and ("交通" in cand_r1 or "202" in cand_r1):
                r1_c1 = re.sub(r"\s+", " ", cand_r1)
        if r0_c1 and r1_c1:
            t_part = r0_c1 if r0_c1.startswith("《") else f"《{r0_c1}》"
            raw_title = f"{t_part}-{r1_c1}"
        else:
            raw_title = r0_c1 or Path(fallback_filename).stem

    base_title, part_num = parse_title_and_part(raw_title, fallback_filename)

    # Remove all pre-table elements and insert single normalized title paragraph
    for el in children[:first_tbl_idx]:
        body.remove(el)

    if kept_p is None:
        kept_p = ET.Element(f"{W}p")

    set_clean_top_title_paragraph(
        kept_p,
        f"{base_title}-{part_num:02d}" if part_num is not None else base_title,
    )
    body.insert(0, kept_p)

    # Also remove trailing empty paragraphs after table except keeping at most 1 clean empty <w:p> before sectPr
    post_els = list(body)
    last_tbl_idx = max(i for i, el in enumerate(post_els) if _local_tag(el) == "tbl")
    trailing_ps = [el for el in post_els[last_tbl_idx + 1 :] if _local_tag(el) == "p"]
    if len(trailing_ps) > 1:
        for extra_p in trailing_ps[1:]:
            if not _get_p_text(extra_p).strip():
                body.remove(extra_p)

    return base_title, part_num, kept_p


def set_clean_top_title_paragraph(p: ET.Element, title_text: str):
    """Format the single top title paragraph above the table cleanly (Blue #0000FF, Bold, 14pt)."""
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        ppr = ET.Element(f"{W}pPr")

    # Preserve existing rFonts if any
    existing_rfonts = None
    first_r = p.find(f".//{W}r/{W}rPr/{W}rFonts")
    if first_r is not None:
        existing_rfonts = copy.deepcopy(first_r)

    for ch in list(p):
        p.remove(ch)

    sp = ppr.find(f"{W}spacing")
    if sp is None:
        sp = ET.SubElement(ppr, f"{W}spacing")
    sp.attrib[f"{W}before"] = "0"
    sp.attrib[f"{W}after"] = "60"
    sp.attrib[f"{W}line"] = "400"
    sp.attrib[f"{W}lineRule"] = "exact"

    p.append(ppr)
    r = ET.SubElement(p, f"{W}r")
    rpr = ET.SubElement(r, f"{W}rPr")
    if existing_rfonts is not None:
        rpr.append(existing_rfonts)
    else:
        rf = ET.SubElement(rpr, f"{W}rFonts")
        rf.attrib[f"{W}ascii"] = "Georgia"
        rf.attrib[f"{W}hAnsi"] = "Georgia"
        rf.attrib[f"{W}eastAsia"] = "SimSun"
    ET.SubElement(rpr, f"{W}b")
    ET.SubElement(rpr, f"{W}bCs")
    c_el = ET.SubElement(rpr, f"{W}color")
    c_el.attrib[f"{W}val"] = "0000FF"
    sz = ET.SubElement(rpr, f"{W}sz")
    sz.attrib[f"{W}val"] = "28"
    sz_cs = ET.SubElement(rpr, f"{W}szCs")
    sz_cs.attrib[f"{W}val"] = "28"

    t_el = ET.SubElement(r, f"{W}t")
    if title_text.startswith(" ") or title_text.endswith(" "):
        t_el.attrib[XML_SPACE] = "preserve"
    t_el.text = title_text


# ---------------- Step 3: Smart Curly Quotes Normalization (Col 2 & Col 3) ----------------

_OPEN_CONTEXT_CHARS = set(" \t\n\r\u00a0\u3000([｛{〔【«„—–-/:：;；")


def smart_curly_quotes_in_paragraph(p: ET.Element):
    """Convert straight quotes (" and ') across all <w:t> nodes of <w:p> into curly quotes (“ ” and ‘ ’)."""
    t_nodes = p.findall(f".//{W}t")
    if not t_nodes:
        return
    full = "".join(t.text or "" for t in t_nodes)
    if '"' not in full and "'" not in full:
        return

    out_chars = list(full)
    n = len(out_chars)
    in_double = False
    in_single = False

    for i, ch in enumerate(out_chars):
        prev_c = out_chars[i - 1] if i > 0 else ""
        next_c = out_chars[i + 1] if i + 1 < n else ""

        if ch == '"':
            if i == 0 or prev_c in _OPEN_CONTEXT_CHARS or (not in_double and prev_c in "‘‚"):
                out_chars[i] = "“"
                in_double = True
            else:
                out_chars[i] = "”"
                in_double = False
        elif ch == "'":
            # Mid-word apostrophe / contraction / Malagasy/French elision (e.g. today's, ankamaroan'ny, l'homme)
            if prev_c and prev_c.isalnum() and next_c and next_c.isalnum():
                out_chars[i] = "’"
            # Possessive plural at end of word (e.g. brothers' fellowship) when not inside single quote
            elif prev_c and prev_c.isalnum() and not in_single:
                out_chars[i] = "’"
            elif i == 0 or prev_c in _OPEN_CONTEXT_CHARS or prev_c == "“":
                out_chars[i] = "‘"
                in_single = True
            else:
                out_chars[i] = "’"
                in_single = False

    replaced = "".join(out_chars)
    pos = 0
    for t in t_nodes:
        orig_len = len(t.text or "")
        if orig_len > 0:
            seg = replaced[pos : pos + orig_len]
            t.text = seg
            if seg.startswith(" ") or seg.endswith(" "):
                t.attrib[XML_SPACE] = "preserve"
            pos += orig_len


# ---------------- Step 4: XML Run Splitting & Character-Level Span Formatting ----------------

def _unwrap_inline_wrappers_in_p(p: ET.Element):
    """Unwrap <w:hyperlink> / <w:smartTag> if needed, keeping <w:r> direct children of <w:p> for clean slicing."""
    # Only unwrap smartTag; keep hyperlinks unless needed, or hoist runs cleanly
    for parent in list(p.iter()):
        for ch in list(parent):
            if _local_tag(ch) in ("smartTag", "fldSimple"):
                idx = list(parent).index(ch)
                subs = list(ch)
                parent.remove(ch)
                for off, s in enumerate(subs):
                    parent.insert(idx + off, s)


def _normalize_single_t_per_run(p: ET.Element):
    """Ensure each text-bearing <w:r> in <w:p> has a single consolidated <w:t> when possible."""
    _unwrap_inline_wrappers_in_p(p)
    for r in list(p.findall(f"{W}r")):
        non_rpr = [c for c in r if _local_tag(c) != "rPr"]
        if len(non_rpr) > 1 and all(_local_tag(c) == "t" for c in non_rpr):
            combined = "".join(c.text or "" for c in non_rpr)
            first_t = non_rpr[0]
            for extra_t in non_rpr[1:]:
                r.remove(extra_t)
            first_t.text = combined
            if combined.startswith(" ") or combined.endswith(" "):
                first_t.attrib[XML_SPACE] = "preserve"


def split_paragraph_runs_at_offsets(p: ET.Element, cut_offsets: Set[int]) -> List[Tuple[int, int, ET.Element]]:
    """Split <w:r> children of <w:p> at exact character offsets, returning [(start, end, run_el), ...]."""
    _normalize_single_t_per_run(p)
    valid_cuts = sorted(c for c in cut_offsets if c > 0)

    pos = 0
    children_snapshot = list(p)
    for child in children_snapshot:
        if _local_tag(child) != "r":
            continue
        non_rpr = [c for c in child if _local_tag(c) != "rPr"]
        if len(non_rpr) != 1 or _local_tag(non_rpr[0]) != "t":
            # Could contain <w:tab/> or <w:br/> alongside or without <w:t>
            t_len = sum(len(c.text or "") for c in non_rpr if _local_tag(c) == "t")
            pos += t_len
            continue

        txt = non_rpr[0].text or ""
        r_start = pos
        r_end = pos + len(txt)
        pos = r_end
        if not txt:
            continue

        inner_cuts = [c for c in valid_cuts if r_start < c < r_end]
        if not inner_cuts:
            continue

        boundaries = [r_start] + inner_cuts + [r_end]
        idx_in_p = list(p).index(child)
        rpr = child.find(f"{W}rPr")
        p.remove(child)

        for seg_i in range(len(boundaries) - 1):
            s0 = boundaries[seg_i] - r_start
            s1 = boundaries[seg_i + 1] - r_start
            piece = txt[s0:s1]
            new_r = ET.Element(f"{W}r")
            if rpr is not None:
                new_r.append(copy.deepcopy(rpr))
            new_t = ET.SubElement(new_r, f"{W}t")
            if piece.startswith(" ") or piece.endswith(" "):
                new_t.attrib[XML_SPACE] = "preserve"
            new_t.text = piece
            p.insert(idx_in_p + seg_i, new_r)

    # Collect final (start, end, r) list
    result: List[Tuple[int, int, ET.Element]] = []
    cur = 0
    for child in p:
        if _local_tag(child) != "r":
            continue
        ts = child.findall(f"{W}t")
        seg_txt = "".join(t.text or "" for t in ts)
        if seg_txt:
            result.append((cur, cur + len(seg_txt), child))
            cur += len(seg_txt)
    return result


def _ensure_rpr(r: ET.Element) -> ET.Element:
    rpr = r.find(f"{W}rPr")
    if rpr is None:
        rpr = ET.Element(f"{W}rPr")
        r.insert(0, rpr)
    return rpr


def set_run_font_size(r: ET.Element, half_pt_str: str):
    rpr = _ensure_rpr(r)
    for tag_name in ("sz", "szCs"):
        el = rpr.find(f"{W}{tag_name}")
        if el is None:
            el = ET.SubElement(rpr, f"{W}{tag_name}")
        el.attrib[f"{W}val"] = half_pt_str


def set_run_bold(r: ET.Element, bold: bool):
    rpr = _ensure_rpr(r)
    for tag_name in ("b", "bCs"):
        el = rpr.find(f"{W}{tag_name}")
        if bold:
            if el is None:
                el = ET.SubElement(rpr, f"{W}{tag_name}")
            else:
                el.attrib.pop(f"{W}val", None)
        else:
            if el is not None:
                rpr.remove(el)


def set_run_blue(r: ET.Element, blue: bool):
    rpr = _ensure_rpr(r)
    c_el = rpr.find(f"{W}color")
    if blue:
        if c_el is None:
            c_el = ET.SubElement(rpr, f"{W}color")
        c_el.attrib[f"{W}val"] = "0000FF"
    else:
        if c_el is not None and is_blue_hex(c_el.attrib.get(f"{W}val", "")):
            rpr.remove(c_el)


def is_blue_hex(val: str) -> bool:
    v = (val or "").strip().lstrip("#").lower()
    if v == "blue":
        return True
    if len(v) == 6 and all(c in "0123456789abcdef" for c in v):
        r_c = int(v[0:2], 16)
        g_c = int(v[2:4], 16)
        b_c = int(v[4:6], 16)
        return b_c >= 140 and b_c > r_c * 1.4 and b_c >= g_c
    return False


def is_run_blue(r: ET.Element) -> bool:
    rpr = r.find(f"{W}rPr")
    if rpr is None:
        return False
    c_el = rpr.find(f"{W}color")
    if c_el is None:
        return False
    return is_blue_hex(c_el.attrib.get(f"{W}val", ""))


def is_run_bold(r: ET.Element) -> bool:
    rpr = r.find(f"{W}rPr")
    if rpr is None:
        return False
    b_el = rpr.find(f"{W}b")
    if b_el is None:
        return False
    return b_el.attrib.get(f"{W}val", "true").lower() not in ("0", "false", "off")


# ---------------- Step 5: Post-Quote Scripture / God's Word Citation 10.5pt (五号) Detector ----------------

_CITATION_KEYWORDS_RE = re.compile(
    r"(?:vol\.?|bd\.?|tome|tomo|boky|卷|话|話|word|wort|palabras?|palavra|parole|слово|λόγος|[:：։])",
    re.I,
)
_SHORT_RESPONSE_EXCLUSION_RE = re.compile(
    r"^(?:pl|s|yes|no|right|amen|true|false|correct|ok|okay|eny|tsia|sí|si|oui|non|ja|nein|да|нет|так|ні"
    r"|是|不是|对|不对|没有|有|能|不能|行|不行|好|阿们|阿們|实际|不实际)[.。!！?？\s]*$",
    re.I,
)


def _is_valid_citation_inner(inner_text: str) -> bool:
    s = inner_text.strip()
    if len(s) < 2:
        return False
    if _SHORT_RESPONSE_EXCLUSION_RE.match(s):
        return False
    # Almost all Bible & God's Word citations contain at least one digit (Unicode-aware: ASCII, Myanmar ၁၂, etc.)
    if any(ch.isdigit() for ch in s):
        return True
    # Also accept volume/book citations without Arabic digits (e.g. Vol. I, Bd. I, 《话...》, or nested quotes)
    if len(s) >= 4 and _CITATION_KEYWORDS_RE.search(s):
        return True
    return False


def find_post_quote_citation_spans(full_text: str) -> List[Tuple[int, int]]:
    """Return [(start, end), ...] character spans of post-quote citations to set to 10.5pt (五号, w:sz=21).

    Strictly keeps trailing external periods (., 。, ։) OUTSIDE the returned span so they stay 12pt (小四).
    """
    spans: List[Tuple[int, int]] = []
    n = len(full_text)
    i = 0

    while i < n:
        ch = full_text[i]

        # Case A: Tortoise-shell brackets 〔...〕 used specifically for scripture citations
        if ch == "〔":
            close_idx = full_text.find("〕", i + 1)
            if close_idx != -1:
                inner = full_text[i + 1 : close_idx]
                if len(inner.strip()) >= 2:
                    spans.append((i, close_idx + 1))
                    i = close_idx + 1
                    continue

        # Case B: Any closing quote followed by optional spaces (or '。') and a citation bracket / 《话・...》
        if ch in CLOSING_QUOTE_CHARS:
            j = i + 1
            while j < n and full_text[j] in " \t\u00a0\u3000":
                j += 1

            # Allow Chinese pattern: ”（太7:23） OR ”。《话・卷一...》
            if j < n and full_text[j] in BRACKET_PAIRS:
                open_br = full_text[j]
                close_br = BRACKET_PAIRS[open_br]
                depth = 1
                k = j + 1
                while k < n and depth > 0:
                    if full_text[k] == open_br:
                        depth += 1
                    elif full_text[k] == close_br:
                        depth -= 1
                    k += 1
                if depth == 0:
                    inner = full_text[j + 1 : k - 1]
                    if _is_valid_citation_inner(inner):
                        spans.append((j, k))
                        i = k
                        continue

            # Chinese post-quote 《话・...》 citation right after ” or ”。
            k_zh = j
            if k_zh < n and full_text[k_zh] in "。.":
                k_zh += 1
                while k_zh < n and full_text[k_zh] in " \t\u00a0\u3000":
                    k_zh += 1
            if k_zh < n and full_text[k_zh] == "《":
                end_book = full_text.find("》", k_zh + 1)
                if end_book != -1:
                    book_inner = full_text[k_zh + 1 : end_book]
                    if any(tag in book_inner for tag in ("话", "話", "卷", "・", "·")):
                        spans.append((k_zh, end_book + 1))
                        i = end_book + 1
                        continue

        i += 1

    return spans


def apply_font_sizes_to_paragraph(p: ET.Element, is_main_title_row: bool = False) -> int:
    """Apply 14pt (sz=28) on main title row, or 12pt (sz=24) body + 10.5pt (sz=21) on post-quote citations."""
    if is_main_title_row:
        for r in p.findall(f"{W}r"):
            set_run_font_size(r, "28")
        return 0

    full_txt = _get_p_text(p)
    if not full_txt:
        return 0

    cit_spans = find_post_quote_citation_spans(full_txt)
    if not cit_spans:
        for r in p.findall(f"{W}r"):
            set_run_font_size(r, "24")
        return 0

    cuts: Set[int] = set()
    for s0, s1 in cit_spans:
        cuts.add(s0)
        cuts.add(s1)

    sliced_runs = split_paragraph_runs_at_offsets(p, cuts)
    for r_s, r_e, r_el in sliced_runs:
        in_cit = any(r_s >= c0 and r_e <= c1 for c0, c1 in cit_spans)
        set_run_font_size(r_el, "21" if in_cit else "24")
    return len(cit_spans)


# ---------------- Step 6: Col 3 Blue & Bold Analysis & Two-Tier Restoration ----------------

def extract_col_style_profile(p: ET.Element) -> Dict:
    """Analyze a Col 1 or Col 2 paragraph's blue & bold segments."""
    segments: List[Tuple[str, bool, bool]] = []
    for r in p.findall(f"{W}r"):
        txt = "".join(t.text or "" for t in r.findall(f"{W}t"))
        if not txt:
            continue
        bl = is_run_blue(r)
        bd = is_run_bold(r)
        if segments and segments[-1][1] == bl and segments[-1][2] == bd:
            prev_t, _, _ = segments[-1]
            segments[-1] = (prev_t + txt, bl, bd)
        else:
            segments.append((txt, bl, bd))

    non_ws_total = 0
    blue_non_ws = 0
    bold_non_ws = 0
    for txt, bl, bd in segments:
        n_chars = len(txt.strip())
        if not n_chars:
            continue
        non_ws_total += n_chars
        if bl:
            blue_non_ws += n_chars
        if bd:
            bold_non_ws += n_chars

    if non_ws_total == 0:
        return {"mode": "plain", "segments": segments}

    all_blue = blue_non_ws == non_ws_total
    all_bold = bold_non_ws == non_ws_total
    any_blue = blue_non_ws > 0
    any_bold = bold_non_ws > 0

    if not any_blue and not any_bold:
        return {"mode": "plain", "segments": segments}
    if all_blue and all_bold:
        return {"mode": "whole_blue_bold", "segments": segments}
    if all_blue and not any_bold:
        return {"mode": "whole_blue_only", "segments": segments}
    if all_bold and not any_blue:
        return {"mode": "whole_bold_only", "segments": segments}

    # Check prefix-only bold/blue (e.g., 'Answer: ...' or '解答：...')
    if len(segments) >= 2:
        first_txt, f_bl, f_bd = segments[0]
        rest_plain = all((not bl and not bd) or not tx.strip() for tx, bl, bd in segments[1:])
        if rest_plain and (f_bd or f_bl) and len(first_txt.strip()) <= 28 and re.search(r"[:：]\s*$", first_txt):
            return {
                "mode": "colon_prefix",
                "prefix_blue": f_bl,
                "prefix_bold": f_bd,
                "segments": segments,
            }

    return {"mode": "partial_ai", "segments": segments}


def format_annotated_source(segments: List[Tuple[str, bool, bool]]) -> str:
    parts = []
    for txt, bl, bd in segments:
        if bl and bd:
            parts.append(f"<BB>{txt}</BB>")
        elif bl:
            parts.append(f"<BL>{txt}</BL>")
        elif bd:
            parts.append(f"<BD>{txt}</BD>")
        else:
            parts.append(txt)
    return "".join(parts)


def apply_whole_paragraph_style(p_fl: ET.Element, blue: bool, bold: bool):
    for r in p_fl.findall(f"{W}r"):
        txt = "".join(t.text or "" for t in r.findall(f"{W}t"))
        if not txt:
            continue
        set_run_blue(r, blue)
        set_run_bold(r, bold if txt.strip() else False)


def apply_colon_prefix_style(p_fl: ET.Element, blue: bool, bold: bool) -> bool:
    full = _get_p_text(p_fl)
    m = re.match(r"^(\s*[^:：\n]{1,35}[:：])", full)
    if not m:
        return False
    end_idx = m.end(1)
    sliced = split_paragraph_runs_at_offsets(p_fl, {end_idx})
    for s0, s1, r_el in sliced:
        if s1 <= end_idx:
            set_run_blue(r_el, blue)
            set_run_bold(r_el, bold)
        else:
            set_run_blue(r_el, False)
            set_run_bold(r_el, False)
    return True


def _trim_outer_quotes_from_bold_span(full_text: str, s0: int, s1: int) -> Tuple[int, int]:
    """Ensure outer quotation marks and surrounding spaces are NOT bolded, while keeping inner punctuation bold."""
    while s0 < s1 and (full_text[s0].isspace() or full_text[s0] in ALL_OUTER_QUOTE_CHARS):
        s0 += 1
    while s1 > s0 and (full_text[s1 - 1].isspace() or full_text[s1 - 1] in ALL_OUTER_QUOTE_CHARS):
        s1 -= 1
    return s0, s1


def _locate_substring_fuzzy(haystack: str, needle: str, start_cursor: int = 0) -> Optional[Tuple[int, int]]:
    if not needle:
        return None
    idx = haystack.find(needle, start_cursor)
    if idx != -1:
        return idx, idx + len(needle)
    idx_any = haystack.find(needle)
    if idx_any != -1:
        return idx_any, idx_any + len(needle)

    # Quote/whitespace normalized fallback
    def _norm_char(c: str) -> str:
        if c in ('"', "“", "”", "„", "«", "»"):
            return '"'
        if c in ("'", "‘", "’", "‚"):
            return "'"
        if c in ("\u00a0", "\u3000", "\t"):
            return " "
        return c

    h_norm = "".join(_norm_char(c) for c in haystack)
    n_norm = "".join(_norm_char(c) for c in needle)
    if len(h_norm) == len(haystack) and len(n_norm) == len(needle):
        idx2 = h_norm.find(n_norm, start_cursor)
        if idx2 != -1:
            return idx2, idx2 + len(n_norm)
        idx3 = h_norm.find(n_norm)
        if idx3 != -1:
            return idx3, idx3 + len(n_norm)
    return None


def apply_ai_spans_to_col3_paragraph(p_fl: ET.Element, span_items: List[Dict]) -> int:
    """Apply [{'text': ..., 'blue': bool, 'bold': bool}, ...] onto p_fl with outer-quote unbold protection."""
    full_txt = _get_p_text(p_fl)
    if not full_txt or not span_items:
        return 0

    n = len(full_txt)
    char_blue = [False] * n
    char_bold = [False] * n

    cursor = 0
    matched_cnt = 0
    for item in span_items:
        sub = str(item.get("text", ""))
        if not sub.strip():
            continue
        bl = bool(item.get("blue", False))
        bd = bool(item.get("bold", False))
        if not bl and not bd:
            continue

        loc = _locate_substring_fuzzy(full_txt, sub, cursor)
        if loc is None:
            loc = _locate_substring_fuzzy(full_txt, sub.strip(), cursor)
        if loc is None:
            continue

        s0, s1 = loc
        cursor = s1
        matched_cnt += 1

        if bl:
            for idx in range(s0, s1):
                char_blue[idx] = True

        if bd:
            b0, b1 = _trim_outer_quotes_from_bold_span(full_txt, s0, s1)
            for idx in range(b0, b1):
                char_bold[idx] = True

    if matched_cnt == 0:
        return 0

    # Also auto-extend blue to immediately adjacent outer quotation marks & post-quote citation if model omitted them
    for i in range(1, n):
        if char_blue[i] and not char_blue[i - 1] and full_txt[i - 1] in ALL_OUTER_QUOTE_CHARS:
            char_blue[i - 1] = True
    for i in range(n - 2, -1, -1):
        if char_blue[i] and not char_blue[i + 1] and full_txt[i + 1] in ALL_OUTER_QUOTE_CHARS:
            char_blue[i + 1] = True

    # Ensure post-quote citations immediately following a blue closing quote also get blue (not bold!)
    for c0, c1 in find_post_quote_citation_spans(full_txt):
        # Check if quote right before c0 was blue
        probe = c0 - 1
        while probe >= 0 and full_txt[probe].isspace():
            probe -= 1
        if probe >= 0 and char_blue[probe]:
            for idx in range(probe + 1, c1):
                char_blue[idx] = True
                char_bold[idx] = False
            if c1 < n and full_txt[c1] in ".。։":
                char_blue[c1] = True
                char_bold[c1] = False

    cuts: Set[int] = set()
    for i in range(1, n):
        if char_blue[i] != char_blue[i - 1] or char_bold[i] != char_bold[i - 1]:
            cuts.add(i)

    sliced = split_paragraph_runs_at_offsets(p_fl, cuts)
    for r_s, r_e, r_el in sliced:
        set_run_blue(r_el, char_blue[r_s])
        set_run_bold(r_el, char_bold[r_s])

    return matched_cnt


_FELLOWSHIP_ALIGN_SYSTEM_PROMPT = """You are a precision multilingual typesetting alignment engine for trilingual Christian fellowship manuscripts (Column 1: Chinese, Column 2: English with formatting tags, Column 3: Target Foreign Language).

In Column 2 (English), styled spans are marked with:
- <BB>...</BB> : BOTH Blue font (#0000FF) AND Bold (typically direct words of God / Lord Jesus / scripture verses)
- <BL>...</BL> : Blue font ONLY, NOT Bold (typically speaker intro like 'The Lord Jesus said: “', closing quote + citation like '” (Matthew 7:23).', or blue-only quotes)
- <BD>...</BD> : Bold ONLY, NOT Blue

Your task:
For each item, locate the EXACT corresponding verbatim substrings inside `col3_text` (Target Foreign Language) in their exact order of appearance, and return JSON:
{
  "results": [
    {
      "id": <item_id>,
      "spans": [
        {"text": "<exact verbatim substring copied from col3_text>", "blue": true, "bold": false},
        {"text": "<exact verbatim substring copied from col3_text>", "blue": true, "bold": true}
      ]
    }
  ]
}

STRICT RULES:
1. Every `text` MUST be copied character-for-character from `col3_text`. Do NOT translate or alter spelling/punctuation.
2. Outer quotation marks (“ ” « » „ “ ‘ ’) and citation parentheses like `(Matio 7:23)` MUST have `"bold": false` (and `"blue": true` if blue in English).
3. Punctuation INSIDE a bolded quote sentence (commas, semicolons, colons, exclamation/question marks that are inside `<BB>...</BB>`) MUST stay included inside the `"bold": true` span!
4. If the same phrase is quoted multiple times in one paragraph, output each occurrence as a separate span in exact sequential order.
"""


_FELLOWSHIP_ALIGN_SCHEMA = {
    "type": "object",
    "properties": {
        "results": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "id": {"type": "integer"},
                    "spans": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "text": {"type": "string"},
                                "blue": {"type": "boolean"},
                                "bold": {"type": "boolean"},
                            },
                            "required": ["text", "blue", "bold"],
                        },
                    },
                },
                "required": ["id", "spans"],
            },
        }
    },
    "required": ["results"],
}


def _normalize_existing_col3_styled_paragraph(p_fl: ET.Element) -> None:
    """Normalize existing Col 3 blue shades to #0000FF and unbold outer quotes / post-quote citations locally."""
    full_txt = _get_p_text(p_fl)
    if not full_txt:
        return
    n = len(full_txt)
    char_blue = [False] * n
    char_bold = [False] * n
    pos = 0
    for r in p_fl.findall(f"{W}r"):
        seg = "".join(t.text or "" for t in r.findall(f"{W}t"))
        if not seg:
            continue
        bl = is_run_blue(r)
        bd = is_run_bold(r)
        for i in range(pos, min(n, pos + len(seg))):
            char_blue[i] = bl
            char_bold[i] = bd
        pos += len(seg)

    # Outer quotes are never bold
    for i, ch in enumerate(full_txt):
        if ch in ALL_OUTER_QUOTE_CHARS:
            char_bold[i] = False

    # Post-quote citations are never bold
    for c0, c1 in find_post_quote_citation_spans(full_txt):
        for idx in range(c0, c1):
            char_bold[idx] = False

    cuts: Set[int] = set()
    for i in range(1, n):
        if char_blue[i] != char_blue[i - 1] or char_bold[i] != char_bold[i - 1]:
            cuts.add(i)

    sliced = split_paragraph_runs_at_offsets(p_fl, cuts)
    for r_s, _r_e, r_el in sliced:
        set_run_blue(r_el, char_blue[r_s])
        set_run_bold(r_el, char_bold[r_s])


def sync_table_col3_blue_and_bold(
    tbl: ET.Element,
    api_pool: Optional[ApiKeyPool],
    model: str,
    api_mode: str,
    enable_ai_col3: bool = True,
    log_fn: Optional[Callable[[str], None]] = None,
) -> Dict:
    """Synchronize Col 3 blue (#0000FF) and bold (<w:b/>) from Col 2/Col 1 across all rows of `tbl`."""
    rows = tbl.findall(f"{W}tr")
    tier1_synced = 0
    skipped_existing = 0
    ai_tasks: List[Dict] = []

    for r_idx, tr in enumerate(rows):
        tcs = tr.findall(f"{W}tc")
        if len(tcs) < 3:
            continue
        ps_zh = tcs[0].findall(f"{W}p")
        ps_en = tcs[1].findall(f"{W}p")
        ps_fl = tcs[2].findall(f"{W}p")

        # Row 0 main title check
        if r_idx == 0 and _is_main_title_table_row(tr):
            for p_fl in ps_fl:
                apply_whole_paragraph_style(p_fl, blue=True, bold=True)
            tier1_synced += 1
            continue

        pair_count = min(len(ps_en), len(ps_fl))
        for p_i in range(pair_count):
            p_en = ps_en[p_i]
            p_fl = ps_fl[p_i]
            p_zh = ps_zh[p_i] if p_i < len(ps_zh) else None

            prof = extract_col_style_profile(p_en)
            mode = prof["mode"]
            if mode == "plain":
                continue
            elif mode == "whole_blue_bold":
                apply_whole_paragraph_style(p_fl, blue=True, bold=True)
                tier1_synced += 1
            elif mode == "whole_blue_only":
                apply_whole_paragraph_style(p_fl, blue=True, bold=False)
                tier1_synced += 1
            elif mode == "whole_bold_only":
                apply_whole_paragraph_style(p_fl, blue=False, bold=True)
                tier1_synced += 1
            elif mode == "colon_prefix":
                if apply_colon_prefix_style(p_fl, blue=prof["prefix_blue"], bold=prof["prefix_bold"]):
                    tier1_synced += 1
            elif mode == "partial_ai":
                col3_txt = _get_p_text(p_fl)
                if not col3_txt.strip():
                    continue
                # Optimization 1: If Col 3 ALREADY has translator-preserved blue/bold matching Col 2's needs, normalize locally & skip AI
                prof_fl = extract_col_style_profile(p_fl)
                en_has_blue = any(bl for _, bl, _ in prof["segments"])
                en_has_bold = any(bd for _, _, bd in prof["segments"])
                fl_has_blue = any(bl for _, bl, _ in prof_fl["segments"])
                fl_has_bold = any(bd for _, _, bd in prof_fl["segments"])
                if (not en_has_blue or fl_has_blue) and (not en_has_bold or fl_has_bold):
                    _normalize_existing_col3_styled_paragraph(p_fl)
                    skipped_existing += 1
                    continue

                ai_tasks.append(
                    {
                        "row_idx": r_idx,
                        "p_fl": p_fl,
                        "annotated_en": format_annotated_source(prof["segments"]),
                        "zh_text": _get_p_text(p_zh) if p_zh is not None else "",
                        "col3_text": col3_txt,
                    }
                )

    ai_synced = 0
    failed_rows: List[int] = []
    if ai_tasks:
        row_nums = sorted({t["row_idx"] + 1 for t in ai_tasks})
        if not enable_ai_col3:
            if log_fn:
                log_fn(
                    f"   ℹ️ 纯本地极速模式：第 {row_nums} 行含段内局部引文已跳过 AI（整格标题/问答已全部本地同步）。"
                )
        elif api_pool is None or len(api_pool) == 0:
            if log_fn:
                log_fn(
                    f"⚠️ 提示：第 {row_nums} 行含段内局部经文/神话引文，当前未填写 API Key，已跳过跨语种段内引文定位（整格标题/问答已全部本地同步）。"
                )
        else:
            BATCH_SIZE = 3
            for b_start in range(0, len(ai_tasks), BATCH_SIZE):
                batch = ai_tasks[b_start : b_start + BATCH_SIZE]
                payload_items = []
                for local_id, task in enumerate(batch):
                    payload_items.append(
                        {
                            "id": local_id,
                            "col1_chinese_ref": task["zh_text"][:1500],
                            "col2_annotated_english": task["annotated_en"],
                            "col3_text": task["col3_text"],
                        }
                    )
                user_prompt = (
                    "Align the <BB>, <BL>, and <BD> spans from `col2_annotated_english` onto `col3_text`:\n"
                    + json.dumps(payload_items, ensure_ascii=False, indent=2)
                )
                try:
                    resp = call_gemini_json(
                        api_key=api_pool,
                        model=model,
                        system_instruction=_FELLOWSHIP_ALIGN_SYSTEM_PROMPT,
                        user_prompt=user_prompt,
                        json_schema=_FELLOWSHIP_ALIGN_SCHEMA,
                        api_mode=api_mode,
                        log_fn=log_fn,
                    )
                    res_list = resp.get("results", []) if isinstance(resp, dict) else []
                    by_id = {int(x.get("id", -1)): x.get("spans", []) for x in res_list if isinstance(x, dict)}
                    for local_id, task in enumerate(batch):
                        spans = by_id.get(local_id, [])
                        applied_n = apply_ai_spans_to_col3_paragraph(task["p_fl"], spans)
                        if applied_n > 0:
                            ai_synced += 1
                            if log_fn:
                                log_fn(f"   ✓ 第 {task['row_idx'] + 1} 行第 3 列已精准还原 {applied_n} 处蓝字/加粗引文")
                except Exception as ai_ex:
                    batch_rows = sorted({t["row_idx"] + 1 for t in batch})
                    failed_rows.extend(batch_rows)
                    if log_fn:
                        log_fn(
                            f"   ⚠️ 第 {batch_rows} 行第 3 列段内引文 AI 定位暂受限（{ai_ex}），已保留该段原文并继续导出文档，请人工核对该行蓝粗！"
                        )

    return {
        "tier1_synced": tier1_synced,
        "skipped_existing": skipped_existing,
        "ai_paragraphs_total": len(ai_tasks),
        "ai_synced": ai_synced,
        "failed_rows": sorted(set(failed_rows)),
    }


# ---------------- Step 7: Spacing, Typography & Content-Proportional 3-Column Widths ----------------

def _is_main_title_table_row(tr: ET.Element) -> bool:
    """Detect if Row 0 of a table is the centered Main Title Row (保留四号 14pt)."""
    tcs = tr.findall(f"{W}tc")
    if len(tcs) < 2:
        return False
    c1_txt = _get_tc_text(tcs[0]).strip()
    c2_txt = _get_tc_text(tcs[1]).strip()
    if not c1_txt or len(c1_txt) > 80 or len(c2_txt) > 160:
        return False
    # Exclude speaker/date row ('上面的交通 2026年...')
    if "上面的交通" in c1_txt or "Fellowship From" in c2_txt:
        return False
    p0 = tcs[0].find(f"{W}p")
    if p0 is not None:
        jc = p0.find(f"{W}pPr/{W}jc")
        if jc is not None and jc.attrib.get(f"{W}val") == "center":
            return True
    # Also check if input Row 0 had sz=28
    for r in tcs[0].findall(f".//{W}r"):
        sz = r.find(f"{W}rPr/{W}sz")
        if sz is not None and sz.attrib.get(f"{W}val") == "28":
            return True
    return False


def apply_curly_quotes_to_table(
    tbl: ET.Element,
    apply_curly_quotes_col2: bool = True,
    apply_curly_quotes_col3: bool = True,
) -> None:
    """Normalize straight quotes (" and ') to curly quotes (“ ” and ‘ ’) on Col 2 & Col 3 prior to span alignment."""
    for tr in tbl.findall(f"{W}tr"):
        tcs = tr.findall(f"{W}tc")
        for c_idx, tc in enumerate(tcs):
            if c_idx == 1 and apply_curly_quotes_col2:
                for p in tc.findall(f"{W}p"):
                    smart_curly_quotes_in_paragraph(p)
            elif c_idx == 2 and apply_curly_quotes_col3:
                for p in tc.findall(f"{W}p"):
                    smart_curly_quotes_in_paragraph(p)


def enforce_table_spacing_and_typography(
    tbl: ET.Element,
    apply_curly_quotes_col2: bool = False,
    apply_curly_quotes_col3: bool = False,
) -> int:
    """Apply optional quote rules, 14pt title / 12pt body / 10.5pt citations, and exact 20pt line spacing."""
    if apply_curly_quotes_col2 or apply_curly_quotes_col3:
        apply_curly_quotes_to_table(tbl, apply_curly_quotes_col2, apply_curly_quotes_col3)
    rows = tbl.findall(f"{W}tr")
    citation_total = 0

    for r_idx, tr in enumerate(rows):
        is_title_row = r_idx == 0 and _is_main_title_table_row(tr)
        tcs = tr.findall(f"{W}tc")
        for tc in tcs:
            for p in tc.findall(f"{W}p"):
                # Exact 20pt line spacing (400 twips), before=0, after=0
                ppr = p.find(f"{W}pPr")
                if ppr is None:
                    ppr = ET.Element(f"{W}pPr")
                    p.insert(0, ppr)
                sp = ppr.find(f"{W}spacing")
                if sp is None:
                    sp = ET.SubElement(ppr, f"{W}spacing")
                for rm_attr in ("beforeLines", "afterLines", "beforeAutospacing", "afterAutospacing"):
                    sp.attrib.pop(f"{W}{rm_attr}", None)
                sp.attrib[f"{W}before"] = "0"
                sp.attrib[f"{W}after"] = "0"
                sp.attrib[f"{W}line"] = "400"
                sp.attrib[f"{W}lineRule"] = "exact"

                # Font sizes: 14pt on Row 0 main title, 12pt body + 10.5pt post-quote citations
                citation_total += apply_font_sizes_to_paragraph(p, is_main_title_row=is_title_row)

    return citation_total


def _estimate_visual_dxa_length(text: str) -> float:
    """Estimate horizontal unwrapped line length (in dxa at 12pt) to equalize vertical column heights."""
    total = 0.0
    for ch in text:
        if ch in "\r\n":
            continue
        cp = ord(ch)
        if ch == " ":
            total += 68.0
        elif (
            0x4E00 <= cp <= 0x9FFF
            or 0x3400 <= cp <= 0x4DBF
            or 0x3000 <= cp <= 0x303F
            or 0xFF00 <= cp <= 0xFFEF
            or 0xAC00 <= cp <= 0xD7AF
        ):
            # CJK / Korean / Fullwidth punctuation: full em width (12pt = 240 dxa)
            total += 240.0
        elif 0x1000 <= cp <= 0x109F or 0x0E00 <= cp <= 0x0FFF or 0x1780 <= cp <= 0x17FF:
            # Myanmar / Thai / Khmer / Lao scripts
            total += 148.0
        elif ch.isupper():
            total += 152.0
        else:
            total += 120.0
    return total


def balance_three_column_widths(root: ET.Element, tbl: ET.Element) -> Tuple[int, int, int]:
    """Dynamically compute (w1, w2, w3) proportional to visual content length so column heights match closely."""
    # Determine total table width in dxa
    total_w = 0
    tblPr = tbl.find(f"{W}tblPr")
    if tblPr is None:
        tblPr = ET.Element(f"{W}tblPr")
        tbl.insert(0, tblPr)

    tblW = tblPr.find(f"{W}tblW")
    if tblW is not None and tblW.attrib.get(f"{W}type") == "dxa":
        try:
            total_w = int(float(tblW.attrib.get(f"{W}w", "0")))
        except ValueError:
            total_w = 0

    if total_w < 6000:
        # Derive from page setup <w:sectPr>/<w:pgSz> - <w:pgMar>
        sect = root.find(f".//{W}sectPr")
        if sect is not None:
            pgSz = sect.find(f"{W}pgSz")
            pgMar = sect.find(f"{W}pgMar")
            if pgSz is not None:
                try:
                    pw = int(float(pgSz.attrib.get(f"{W}w", "16838")))
                    ml = int(float(pgMar.attrib.get(f"{W}left", "1134"))) if pgMar is not None else 1134
                    mr = int(float(pgMar.attrib.get(f"{W}right", "1134"))) if pgMar is not None else 1134
                    if pw - ml - mr >= 6000:
                        total_w = pw - ml - mr
                except ValueError:
                    pass
    if total_w < 6000:
        total_w = 14560

    vis_sums = [0.0, 0.0, 0.0]
    for tr in tbl.findall(f"{W}tr"):
        tcs = tr.findall(f"{W}tc")
        if len(tcs) < 3:
            continue
        for c_i in range(3):
            vis_sums[c_i] += _estimate_visual_dxa_length(_get_tc_text(tcs[c_i]))

    sum_all = sum(vis_sums)
    if sum_all <= 1.0:
        ratios = [1.0 / 3.0, 1.0 / 3.0, 1.0 / 3.0]
    else:
        raw_ratios = [v / sum_all for v in vis_sums]
        # Clamp each column between 21% and 48% so no column ever gets squeezed or overly wide, then renormalize
        clamped = [max(0.21, min(0.48, r)) for r in raw_ratios]
        c_sum = sum(clamped)
        ratios = [r / c_sum for r in clamped]

    CELL_PAD_DXA = 216  # ~108 dxa left + 108 dxa right cell margin
    usable_total = max(3000, total_w - 3 * CELL_PAD_DXA)
    w1 = int(round(usable_total * ratios[0])) + CELL_PAD_DXA
    w2 = int(round(usable_total * ratios[1])) + CELL_PAD_DXA
    w3 = total_w - w1 - w2

    # Write fixed table width & layout
    if tblW is None:
        tblW = ET.SubElement(tblPr, f"{W}tblW")
    tblW.attrib[f"{W}w"] = str(total_w)
    tblW.attrib[f"{W}type"] = "dxa"

    tblLayout = tblPr.find(f"{W}tblLayout")
    if tblLayout is None:
        tblLayout = ET.SubElement(tblPr, f"{W}tblLayout")
    tblLayout.attrib[f"{W}type"] = "fixed"

    # Update <w:tblGrid>
    tblGrid = tbl.find(f"{W}tblGrid")
    if tblGrid is None:
        tblGrid = ET.Element(f"{W}tblGrid")
        tbl.insert(1, tblGrid)
    for ch in list(tblGrid):
        tblGrid.remove(ch)
    for col_w in (w1, w2, w3):
        gc = ET.SubElement(tblGrid, f"{W}gridCol")
        gc.attrib[f"{W}w"] = str(col_w)

    # Update every row's <w:tcW>
    widths_tuple = (w1, w2, w3)
    for tr in tbl.findall(f"{W}tr"):
        tcs = tr.findall(f"{W}tc")
        if len(tcs) == 3:
            for c_i, tc in enumerate(tcs):
                tcPr = tc.find(f"{W}tcPr")
                if tcPr is None:
                    tcPr = ET.Element(f"{W}tcPr")
                    tc.insert(0, tcPr)
                tcW = tcPr.find(f"{W}tcW")
                if tcW is None:
                    tcW = ET.SubElement(tcPr, f"{W}tcW")
                tcW.attrib[f"{W}w"] = str(widths_tuple[c_i])
                tcW.attrib[f"{W}type"] = "dxa"

    return widths_tuple


# ---------------- Step 8: Package I/O, Preview Inspection & Full Pipeline ----------------

def write_cleaned_docx(
    src_docx: Path,
    dst_docx: Path,
    doc_root: ET.Element,
):
    dst_docx.parent.mkdir(parents=True, exist_ok=True)
    doc_bytes = ET.tostring(doc_root, encoding="utf-8", xml_declaration=True)

    fd, tmp_name = tempfile.mkstemp(prefix=".fel_", suffix=".docx", dir=str(dst_docx.parent))
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        with zipfile.ZipFile(src_docx, "r") as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                fn = item.filename
                fn_low = fn.lower()
                # Drop all comment XML parts completely
                if fn_low.startswith("word/comments"):
                    continue

                raw = zin.read(fn)
                if fn == "word/document.xml":
                    zout.writestr(item, doc_bytes)
                elif fn == "word/settings.xml":
                    zout.writestr(item, disable_track_revisions_in_settings_xml(raw))
                elif fn == "word/_rels/document.xml.rels":
                    zout.writestr(item, strip_comment_rels_bytes(raw))
                elif fn == "[Content_Types].xml":
                    zout.writestr(item, strip_comment_content_types_bytes(raw))
                else:
                    zout.writestr(item, raw)
        os.replace(str(tmp_path), str(dst_docx))
    except Exception:
        try:
            if tmp_path.exists():
                tmp_path.unlink()
        except Exception:
            pass
        raise


def inspect_fellowship_docx_quick(docx_path: Path) -> Dict:
    """Fast local inspection for GUI preview table (title, part_num, row_count)."""
    validate_docx_archive(docx_path)
    with zipfile.ZipFile(docx_path, "r") as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find(f"{W}body")
    tbl = body.find(f"{W}tbl") if body is not None else None
    row_cnt = len(tbl.findall(f"{W}tr")) if tbl is not None else 0

    raw_title = ""
    if body is not None:
        for el in body:
            if _local_tag(el) == "tbl":
                break
            if _local_tag(el) == "p":
                tx = _get_p_text(el).strip()
                if tx and not _is_boilerplate_header_paragraph(tx):
                    raw_title = tx
    if not raw_title and tbl is not None and row_cnt > 0:
        tcs0 = tbl.findall(f"{W}tr")[0].findall(f"{W}tc")
        if tcs0:
            raw_title = _get_tc_text(tcs0[0]).strip()

    base_title, part_num = parse_title_and_part(raw_title, docx_path.name)
    return {
        "path": str(docx_path),
        "filename": docx_path.name,
        "base_title": base_title or docx_path.stem,
        "part_num": part_num,
        "row_count": row_cnt,
    }


def build_fellowship_output_name(
    tri_prefix: str,
    base_title: str,
    date_yyyymmdd: str,
    part_num: Optional[int] = None,
) -> str:
    """Format output filename per spec:
    Single part: 中英缅_《怎样实行真理（二）》-上面的交通 2026年11月2日-01-三语稿-20261115.docx
    Merged:      中英缅_《怎样实行真理（二）》-上面的交通 2026年11月2日-三语稿-20261115.docx
    """
    pfx = sanitize_filename_part(tri_prefix).rstrip("_")
    clean_title = sanitize_filename_part(base_title)
    dt = re.sub(r"[^0-9]", "", date_yyyymmdd or "") or datetime.now().strftime("%Y%m%d")
    if part_num is not None:
        return f"{pfx}_{clean_title}-{part_num:02d}-三语稿-{dt}.docx"
    return f"{pfx}_{clean_title}-三语稿-{dt}.docx"


def _is_duplicate_header_row_in_subsequent_part(cand_tr: ET.Element, part1_header_rows: List[ET.Element]) -> bool:
    """Check if `cand_tr` in Part 2..N is a repeated main title / date header row already in Part 1."""
    tcs_c = cand_tr.findall(f"{W}tc")
    if not tcs_c:
        return False
    c1_cand = re.sub(r"\s+", "", _get_tc_text(tcs_c[0]))
    if not c1_cand or len(c1_cand) > 60:
        return False
    for h_tr in part1_header_rows:
        tcs_h = h_tr.findall(f"{W}tc")
        if tcs_h:
            c1_h = re.sub(r"\s+", "", _get_tc_text(tcs_h[0]))
            if c1_cand == c1_h:
                return True
    return False


def run_fellowship_pipeline(
    input_files: List[str],
    output_dir: str,
    output_mode: str = "single_only",  # "single_only", "merged_only", "both"
    lang_code: str = "AUTO",
    tri_prefix_override: str = "",
    quote_mode: str = "auto",  # "auto", "curly", "skip"
    date_yyyymmdd: str = "",
    enable_ai_col3: bool = True,
    api_key=None,
    model: str = "gemini-3.5-flash-lite",
    api_mode: str = "auto",
    rotation_mode: str = "round_robin",
    log_fn: Optional[Callable[[str], None]] = None,
    progress_fn: Optional[Callable[[float, str], None]] = None,
) -> Dict:
    """Execute Scenario 4 Fellowship Manuscript cleaning, styling, and optional multi-part merging."""

    def log(msg: str):
        if log_fn:
            log_fn(msg)

    def prog(pct: float, stage: str):
        if progress_fn:
            progress_fn(pct, stage)

    raw_paths = [Path(x).resolve() for x in input_files if str(x).strip()]
    sorted_paths = natural_sort_docx_paths(raw_paths)
    if not sorted_paths:
        raise ValueError("未提供任何待整理的 .docx 文档。")

    out_dir_p = Path(output_dir).resolve()
    out_dir_p.mkdir(parents=True, exist_ok=True)

    langs = load_languages()
    resolved_code = (lang_code or "AUTO").strip().upper()
    if resolved_code == "AUTO":
        hint_blob = " ".join([sorted_paths[0].parent.name] + [p.name for p in sorted_paths[:3]])
        det = detect_lang_code_from_hint(hint_blob, langs)
        if det:
            resolved_code = det

    tri_prefix = (tri_prefix_override or "").strip().rstrip("_")
    if not tri_prefix:
        tri_prefix = get_tri_prefix(resolved_code, langs)

    dt_str = re.sub(r"[^0-9]", "", date_yyyymmdd or "") or datetime.now().strftime("%Y%m%d")

    pool = None
    if enable_ai_col3:
        if isinstance(api_key, ApiKeyPool):
            pool = api_key
        elif api_key:
            pool = ApiKeyPool(api_key, rotation_mode=rotation_mode)

    # Check part sequence continuity
    part_nums = [extract_part_num_from_filename(p.name) for p in sorted_paths]
    known_pnums = [n for n in part_nums if n is not None]
    if len(known_pnums) >= 2:
        expected = list(range(min(known_pnums), max(known_pnums) + 1))
        missing = sorted(set(expected) - set(known_pnums))
        if missing:
            log(f"⚠️ 提醒：检测到的分篇序号存在跳号，缺失 Part {missing}，请留意是否漏拖文件！")

    apply_q_col2 = quote_mode != "skip"
    if quote_mode == "skip":
        apply_q_col3 = False
    elif quote_mode == "curly":
        apply_q_col3 = True
    else:
        # "auto": skip curly replacement on French if guillemets « » are used
        apply_q_col3 = resolved_code != "FR"

    processed_docs: List[Dict] = []
    ai_failed_items: List[str] = []
    total_files = len(sorted_paths)

    for idx, src_p in enumerate(sorted_paths):
        base_pct = 5.0 + (idx / max(1, total_files)) * 75.0
        prog(base_pct, f"正在清洗并还原格式 [{idx + 1}/{total_files}]：{src_p.name}")
        log(f"▶ [{idx + 1}/{total_files}] 处理文档：{src_p.name}")

        validate_docx_archive(src_p)
        with zipfile.ZipFile(src_p, "r") as zf:
            doc_root = ET.fromstring(zf.read("word/document.xml"))

        # 1. Highest priority: unwrap <w:sdt> content controls
        sdt_n = unwrap_all_sdts(doc_root)
        # 2. Accept track changes & strip comments + highlights
        rev_n, cmt_n = accept_revisions_and_strip_comments(doc_root)
        # 3. Clean boilerplate header & normalize title
        base_title, part_num, title_p = clean_body_header_and_extract_title(doc_root, src_p.name)

        # If multiple files were provided without explicit part numbers, assign sequential 1..N
        if part_num is None and total_files > 1:
            part_num = idx + 1
            set_clean_top_title_paragraph(title_p, f"{base_title}-{part_num:02d}")

        body = doc_root.find(f"{W}body")
        tbl = body.find(f"{W}tbl")

        # Auto-detect French guillemets in auto quote mode
        use_q_col3 = apply_q_col3
        if quote_mode == "auto" and use_q_col3:
            sample_c3 = "".join(_get_tc_text(tr.findall(f"{W}tc")[2])[:300] for tr in tbl.findall(f"{W}tr")[:3] if len(tr.findall(f"{W}tc")) >= 3)
            if "«" in sample_c3 and "»" in sample_c3:
                use_q_col3 = False

        # 4. Curly quotes on Col 2 & Col 3 BEFORE span alignment so AI & citation detector see clean quotes
        apply_curly_quotes_to_table(
            tbl=tbl,
            apply_curly_quotes_col2=apply_q_col2,
            apply_curly_quotes_col3=use_q_col3,
        )

        # 5. Sync Col 3 blue (#0000FF) & bold (<w:b/>) BEFORE citation font-size slicing
        sync_stats = sync_table_col3_blue_and_bold(
            tbl=tbl,
            api_pool=pool,
            model=model,
            api_mode=api_mode,
            enable_ai_col3=enable_ai_col3,
            log_fn=log,
        )
        if sync_stats.get("failed_rows"):
            ai_failed_items.append(f"{src_p.name}（第 {sync_stats['failed_rows']} 行）")

        # 6. Exact 20pt line spacing + 14pt title / 12pt body / 10.5pt post-quote citations
        cit_cnt = enforce_table_spacing_and_typography(tbl=tbl)

        # 7. Balance 3-column widths for single-part output
        w_tuple = balance_three_column_widths(doc_root, tbl)

        skip_note = (
            f", 已有蓝粗保留 {sync_stats['skipped_existing']} 段"
            if sync_stats.get("skipped_existing")
            else ""
        )
        log(
            f"   ✓ 拆内容控件 {sdt_n} 处 | 清批注/修订 {cmt_n + rev_n} 处 | "
            f"第3列蓝粗同步(整格/前缀 {sync_stats['tier1_synced']} 处{skip_note}, AI引文 {sync_stats['ai_synced']} 段) | "
            f"五号字出处 {cit_cnt} 处 | 列宽 {w_tuple}"
        )

        processed_docs.append(
            {
                "src_path": src_p,
                "root": doc_root,
                "base_title": base_title,
                "part_num": part_num,
                "title_p": title_p,
                "tbl": tbl,
            }
        )

    saved_files: List[str] = []

    # Save individual formatted parts if requested ("single_only" or "both")
    if output_mode in ("single_only", "both"):
        prog(84.0, "正在导出单篇整理定稿（修订模式：关闭）...")
        for item in processed_docs:
            out_name = build_fellowship_output_name(
                tri_prefix=tri_prefix,
                base_title=item["base_title"],
                date_yyyymmdd=dt_str,
                part_num=item["part_num"],
            )
            dst_p = out_dir_p / out_name
            # Ensure top title has -XX for single part if part_num exists
            if item["part_num"] is not None:
                set_clean_top_title_paragraph(item["title_p"], f"{item['base_title']}-{item['part_num']:02d}")
            write_cleaned_docx(item["src_path"], dst_p, item["root"])
            saved_files.append(str(dst_p))
            log(f"✅ 已输出单篇：{dst_p.name}")

    # Build & save merged unified document if requested ("merged_only" or "both")
    if output_mode in ("merged_only", "both"):
        prog(92.0, "正在按顺序合并总文档并全局均衡三列宽度...")
        first = processed_docs[0]
        merged_root = copy.deepcopy(first["root"])
        merged_body = merged_root.find(f"{W}body")
        merged_title_p = merged_body.find(f"{W}p")
        merged_tbl = merged_body.find(f"{W}tbl")

        # Strip -01 from top title paragraph in merged document
        if merged_title_p is not None:
            set_clean_top_title_paragraph(merged_title_p, first["base_title"])

        p1_rows = merged_tbl.findall(f"{W}tr")
        p1_headers = p1_rows[:2] if len(p1_rows) >= 2 else p1_rows[:1]

        for extra in processed_docs[1:]:
            extra_rows = extra["tbl"].findall(f"{W}tr")
            for er_i, er in enumerate(extra_rows):
                if er_i < 2 and _is_duplicate_header_row_in_subsequent_part(er, p1_headers):
                    continue
                merged_tbl.append(copy.deepcopy(er))

        # Globally balance 3-column widths across all merged rows
        merged_widths = balance_three_column_widths(merged_root, merged_tbl)
        total_merged_rows = len(merged_tbl.findall(f"{W}tr"))

        merged_name = build_fellowship_output_name(
            tri_prefix=tri_prefix,
            base_title=first["base_title"],
            date_yyyymmdd=dt_str,
            part_num=None,
        )
        merged_dst = out_dir_p / merged_name
        write_cleaned_docx(first["src_path"], merged_dst, merged_root)
        saved_files.append(str(merged_dst))
        log(f"🎉 已输出汇总合并总文档（共 {total_merged_rows} 行 | 列宽 {merged_widths}）：{merged_dst.name}")

    prog(100.0, f"整理完成！共输出 {len(saved_files)} 份文档")
    return {
        "saved_files": saved_files,
        "tri_prefix": tri_prefix,
        "lang_code": resolved_code,
        "ai_failed_items": ai_failed_items,
    }
