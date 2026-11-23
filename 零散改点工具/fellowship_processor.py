"""Scenario 4 — Multi-Part & Brother Fellowship Manuscript Processor (分part稿整理/合并).

Capabilities:
1. Safe XML Pre-cleaning & OOXML Integrity (Zero Data Loss, Zero Word Warnings):
   - Unwraps all <w:sdt> content controls innermost-first so inner text/runs are never lost.
   - Turns OFF Word Track Changes mode (<w:trackRevisions/> removed in word/settings.xml -> '修订: 关闭').
   - Accepts all remaining <w:ins>/<w:del>/*Change revisions.
   - Removes all review comments & people.xml (commentRangeStart/End, commentReference, word/comments*.xml,
     word/people.xml + rels + content-types), cleans internal/orphaned bookmarks & w14:paraId, and enforces
     globally unique w:id attributes.
2. Single-Doc & Multi-Doc Multi-Part Table Merging + Header Normalization:
   - Automatically detects and merges internal Part 1 / Part 2 / Part 3 tables inside a single .docx into 1 unified
     table (stripping 'part1' suffix, deleting intermediate 'Part 2'/'Part 3' headings & blank separators, and
     aligning all vertical borders).
   - Strips 'Notice from the Church...' / translator remarks above the main table and cleans noise ('片段', 'Part N',
     'Cleaned File') from base titles.
3. Smart Quote Normalization + Cell Trailing Whitespace & Empty Paragraph Cleanup:
   - Strips trailing spaces from paragraph ends and removes trailing empty <w:p> elements at the bottom of non-empty
     cells while preserving intentional whole-row blank separator rows (['', '', '']).
   - Converts straight quotes (" and ') to curly quotes (“ ” and ‘ ’), with French « » preservation.
4. Decoupled Col 3 Blue (#0000FF), Red (#FF0000 when sync_red=True) & Bold (<w:b/>) Synchronization:
   - Strictly mirrors Col 2's exact (color, bold) per span — never bundling blue + bold blindly!
   - Deterministic leading prefix synchronization ('Question 10:' -> 'Swali la 10:', 'Answer:', '1.', 'I.', etc.).
   - Deterministic quote-pair & inter-quote narrative bridge synchronization (keeping narrative bridges like
     'Ο Ησαΐας απάντησε:' unstyled even if input Col 3 was erroneously all-bold).
   - Red font synchronization (sync_red=True): preserves existing Col 3 red runs if present; otherwise aligns Col 2's
     red spans via AI or offline proportional clause/word fallback (local_align_col3_spans).
5. Standard Multilingual Fonts, Thai Typography (18/16/14pt), Landscape A4 & Citation Sizing:
   - Forces Landscape A4 (<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>).
   - Applies standard fonts per language from 各语系字体.xlsx (Col 1 宋体, Col 2 Georgia, Col 3 get_lang_typography).
   - Sizes: default 28 / 24 / 21 half-pts (14pt / 12pt / 10.5pt); Thai (TH) Angsana New 36 / 32 / 28 (18 / 16 / 14pt).
   - Detects multi-script Bible citations (with or without preceding quote) AND cross-column God's Word citations
     ('摘自' / '(The Word, ...)' / sz=21), applying citation_sz + unbold + Col 2's exact color.
   - Strips staggered gridSpan/gridBefore/vMerge artifacts and balances 3-column widths cleanly.
"""

import copy
from datetime import datetime
import json
import os
from pathlib import Path
import re
import sys
import tempfile
from typing import Callable, Dict, List, Optional, Set, Tuple
import unicodedata
import xml.etree.ElementTree as ET
import zipfile


def _ensure_skill_scripts_fp():
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


_ensure_skill_scripts_fp()
import docx_xml_utils  # noqa: E402

from ai_client import (
    ApiKeyPool,
    call_gemini_json,
    format_friendly_api_error,
    local_align_col3_spans,
)
from lang_registry import (
    detect_lang_code_from_hint,
    get_lang_typography,
    get_tri_prefix,
    load_languages,
)


MAX_DOCX_ENTRIES = 5000
MAX_DOCX_MEMBER_BYTES = 100 * 1024 * 1024
MAX_DOCX_UNCOMPRESSED_BYTES = 300 * 1024 * 1024
_DTD_RE = re.compile(rb"<!\s*(?:DOCTYPE|ENTITY)\b", re.IGNORECASE)


def _read_zip_member_bounded(zf: zipfile.ZipFile, member: str | zipfile.ZipInfo, max_bytes: int = MAX_DOCX_MEMBER_BYTES) -> bytes:
    """Read a ZIP member enforcing an actual decompressed stream byte limit (defends against forged ZIP headers)."""
    with zf.open(member, "r") as fp:
        data = fp.read(max_bytes + 1)
    if len(data) > max_bytes:
        mname = member.filename if isinstance(member, zipfile.ZipInfo) else str(member)
        raise ValueError(f"文档内部成员解压流体积超过安全上限 ({max_bytes // (1024 * 1024)}MB)：{mname}")
    return data


def _parse_safe_xml_bytes(xml_bytes: bytes, max_bytes: int = MAX_DOCX_MEMBER_BYTES) -> ET.Element:
    """Reject DTD / ENTITY declarations and oversized XML payloads before ElementTree parsing."""
    if len(xml_bytes) > max_bytes:
        raise ValueError("XML 内容体积超过安全上限。")
    if _DTD_RE.search(xml_bytes):
        raise ValueError("检测到不安全的 XML DOCTYPE/ENTITY 声明，已拒绝解析。")
    return ET.fromstring(xml_bytes)


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
            raw_fn = info.filename
            if "\x00" in raw_fn or ":" in raw_fn or raw_fn.startswith(("/", "\\")):
                raise ValueError(f"文档包含不安全的归档路径：{raw_fn}")
            fn = raw_fn.replace("\\", "/")
            if ".." in fn.split("/"):
                raise ValueError(f"文档包含路径穿越风险条目：{raw_fn}")
            if info.file_size < 0 or info.file_size > MAX_DOCX_MEMBER_BYTES:
                raise ValueError(f"文档单条目体积超过安全上限 (100MB)：{raw_fn}")
            total_uncompressed += info.file_size
            if total_uncompressed > MAX_DOCX_UNCOMPRESSED_BYTES:
                raise ValueError(f"文档解压体积超过安全上限 (300MB)：{p.name}")
            if fn == "word/document.xml":
                has_doc_xml = True
        if not has_doc_xml:
            raise ValueError(f"不是有效的 Word (.docx) 文档（缺少 word/document.xml）：{p.name}")


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
REL_PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_NS = "http://www.w3.org/XML/1998/namespace"

MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
W = f"{{{W_NS}}}"
R_ATTR = f"{{{R_NS}}}"
WP = f"{{{WP_NS}}}"
W14 = f"{{{W14_NS}}}"
MC = f"{{{MC_NS}}}"
XML_SPACE = f"{{{XML_NS}}}space"

COMMON_NAMESPACES = {
    "w": W_NS,
    "r": R_NS,
    "wp": WP_NS,
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "mc": MC_NS,
    "w14": W14_NS,
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
}
for _pfx, _uri in COMMON_NAMESPACES.items():
    ET.register_namespace(_pfx, _uri)

# Strict ECMA-376 CT_PPr child element ordering so Word never flags out-of-order paragraph properties
_PPR_TAG_ORDER: Dict[str, int] = {
    name: idx
    for idx, name in enumerate(
        [
            "pStyle",
            "keepNext",
            "keepLines",
            "pageBreakBefore",
            "framePr",
            "widowControl",
            "numPr",
            "suppressLineNumbers",
            "pBdr",
            "shd",
            "tabs",
            "suppressAutoHyphens",
            "kinsoku",
            "wordWrap",
            "overflowPunct",
            "topLinePunct",
            "autoSpaceDE",
            "autoSpaceDN",
            "bidi",
            "adjustRightInd",
            "snapToGrid",
            "spacing",
            "ind",
            "contextualSpacing",
            "mirrorIndents",
            "suppressOverlap",
            "jc",
            "textDirection",
            "textAlignment",
            "textboxTightWrap",
            "outlineLvl",
            "divId",
            "cnfStyle",
            "rPr",
            "sectPr",
            "pPrChange",
        ]
    )
}

# Strict ECMA-376 CT_RPr child element ordering so Word never flags out-of-order run properties
_RPR_TAG_ORDER: Dict[str, int] = {
    name: idx
    for idx, name in enumerate(
        [
            "rStyle",
            "rFonts",
            "b",
            "bCs",
            "i",
            "iCs",
            "caps",
            "smallCaps",
            "strike",
            "dstrike",
            "outline",
            "shadow",
            "emboss",
            "imprint",
            "noProof",
            "snapToGrid",
            "vanish",
            "webHidden",
            "color",
            "spacing",
            "w",
            "kern",
            "position",
            "sz",
            "szCs",
            "highlight",
            "u",
            "effect",
            "bdr",
            "shd",
            "fitText",
            "vertAlign",
            "rtl",
            "cs",
            "em",
            "lang",
            "eastAsianLayout",
            "specVanish",
            "oMath",
        ]
    )
}


OUTPUT_MODES = [
    ("仅输出各单篇（默认，单文档含多part自动合并）", "single_only"),
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
    "“",  # U+201C left double quotation mark (used as CLOSING quote in German/Bulgarian/Czech/Polish „...“)
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
CLOSING_TO_OPENING_BRACKET: Dict[str, str] = {v: k for k, v in BRACKET_PAIRS.items()}

BOILERPLATE_HEADER_PATTERNS = [
    re.compile(r"notice\s+from\s+the\s+church", re.I),
    re.compile(r"three\s+types\s+of\s+[“\"']?you", re.I),
    re.compile(r"you\(pl\)", re.I),
    re.compile(r"you\(s\)", re.I),
    re.compile(r"differences\s+between\s+english\s+and\s+chinese", re.I),
    re.compile(r"proper\s+grammar\s+in\s+your\s+language", re.I),
    re.compile(r"no\s+need\s+to\s+review\s+the\s+part\s+in\s+blue\s+font", re.I),
    re.compile(r"words\s+in\s+blue\s+font\s+are\s+already\s+translated", re.I),
    re.compile(r"high[\-\s]*frequency\s+vocabular", re.I),
    re.compile(r"^(?:译者备注|审校说明|交接说明|翻译说明|注意|备注)[:：]"),
    re.compile(r"^note\s*:", re.I),
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
    return cleaned or "整理稿"


# ---------------- Part Number, Noise Cleaning & Script Detection ----------------

_TITLE_PART_WORD_RE = re.compile(r"[\s\-—–_]*(?:[Pp]art|[Pp]t\.?)\s*0*(\d{1,3})\s*$")
_FILE_PART_WORD_RE = re.compile(r"(?:[Pp]art|[Pp]t\.?)\s*0*(\d{1,3})(?!\d)")
_FILE_DASH_NUM_RE = re.compile(r"[\-—–_]0*(\d{1,2})(?=[\-—–_\s(.]|$)")


def clean_title_noise(raw: str, is_filename_stem: bool = False) -> str:
    """Clean 'Cleaned File', word-count tags, trailing date/片段 noise, and Part markers from a title."""
    s = re.sub(r"\s+", " ", (raw or "").strip())
    if not s:
        return ""
    # Strip trailing '-三语稿-20261115'
    s = re.sub(r"[\s\-—–_]*三语稿(?:[\s\-—–_]*20\d{6})?\s*$", "", s).strip()
    # Strip '- Cleaned File' / '_Cleaned file'
    s = re.sub(r"[\s\-—–_]*[Cc]leaned\s+[Ff]ile\s*$", "", s).strip()
    # Strip word count parentheticals e.g. '_(RV 4038 words, Bible)' or '(3304 words)'
    s = re.sub(r"[\s_]*\((?:UP|RV|\d+\s*words)[^)]*\)", "", s, flags=re.I).strip()
    # Strip trailing '_20260714片段' / '20261102' / '片段'
    s = re.sub(r"[\s\-—–_]*20\d{6}\s*片段\s*$", "", s).strip()
    s = re.sub(r"[\s_]+20\d{6}\s*$", "", s).strip()
    s = re.sub(r"[\s\-—–_]*片段\s*$", "", s).strip()
    if is_filename_stem:
        # Strip leading index + lang code from stems like '71 HI_38_...' or '中英斯_22_...'
        s = re.sub(r"^中英[^\s_\-]{1,4}[_\-\s]+", "", s).strip()
        s = re.sub(r"^\d+\s+(?=[A-Z]{2,3}[_\-])", "", s).strip()
        s = re.sub(r"^[A-Z]{2,3}[_\-]+", "", s).strip()
        s = re.sub(r"^\d{1,3}[_\-]+(?=[\u4e00-\u9fff《])", "", s).strip()
    return s.strip(" _-—–")


def extract_part_num_from_filename(filename: str) -> Optional[int]:
    stem = Path(filename).stem
    m = _FILE_PART_WORD_RE.search(stem)
    if m:
        return int(m.group(1))
    stem_no_noise = re.sub(r"[\s\-—–_]*[Cc]leaned\s+[Ff]ile", "", stem)
    stem_no_noise = re.sub(r"\((?:UP|RV|\d+\s*words)[^)]*\)", "", stem_no_noise, flags=re.I)
    stem_no_date = re.sub(r"20\d{6}(?:片段)?", "", stem_no_noise)
    m2 = re.search(r"日[\s\-—–_]*0*(\d{1,2})(?=[\-—–_\s]|$)", stem_no_date)
    if m2:
        return int(m2.group(1))
    matches = list(_FILE_DASH_NUM_RE.finditer(stem_no_date))
    if matches:
        for cand in reversed(matches):
            if cand.start() > 3:
                return int(cand.group(1))
    return None


def parse_title_and_part(raw_title: str, fallback_filename: str = "") -> Tuple[str, Optional[int]]:
    """Parse top title line into (base_title_without_part, part_number_int_or_None)."""
    t = clean_title_noise(raw_title, is_filename_stem=False)

    m1 = _TITLE_PART_WORD_RE.search(t)
    if m1:
        base = clean_title_noise(t[: m1.start()], is_filename_stem=False)
        return base, int(m1.group(1))

    m2 = re.search(r"(.*?)(?:[\s]*[\-—–_]\s*0*(\d{1,2}))\s*$", t)
    if m2:
        prefix = m2.group(1).strip()
        if not re.search(r"20\d{2}[\-—–_]\d{1,2}$", prefix):
            return clean_title_noise(prefix, is_filename_stem=False), int(m2.group(2))

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


def detect_lang_code_from_col3_script(sample_text: str) -> str:
    """Fallback Unicode script detector when filename/folder hint does not specify language."""
    if not sample_text:
        return ""
    counts: Dict[str, int] = {}
    for ch in sample_text:
        cp = ord(ch)
        if 0x0E00 <= cp <= 0x0E7F:
            counts["TH"] = counts.get("TH", 0) + 1
        elif 0x0E80 <= cp <= 0x0EFF:
            counts["LO"] = counts.get("LO", 0) + 1
        elif 0x1000 <= cp <= 0x109F:
            counts["MY"] = counts.get("MY", 0) + 1
        elif 0x1780 <= cp <= 0x17FF:
            counts["KM"] = counts.get("KM", 0) + 1
        elif 0x0370 <= cp <= 0x03FF or 0x1F00 <= cp <= 0x1FFF:
            counts["GR"] = counts.get("GR", 0) + 1
        elif 0xAC00 <= cp <= 0xD7AF:
            counts["KO"] = counts.get("KO", 0) + 1
        elif 0x0590 <= cp <= 0x05FF:
            counts["HE"] = counts.get("HE", 0) + 1
        elif 0x0530 <= cp <= 0x058F:
            counts["HY"] = counts.get("HY", 0) + 1
        elif 0x10A0 <= cp <= 0x10FF:
            counts["KA"] = counts.get("KA", 0) + 1
        elif 0x0900 <= cp <= 0x097F:
            counts["HI"] = counts.get("HI", 0) + 1
        elif 0x3040 <= cp <= 0x30FF:
            counts["JA"] = counts.get("JA", 0) + 1
        elif 0x0980 <= cp <= 0x09FF:
            counts["BN"] = counts.get("BN", 0) + 1
        elif 0x0B80 <= cp <= 0x0BFF:
            counts["TA"] = counts.get("TA", 0) + 1
        elif 0x0A00 <= cp <= 0x0A7F:
            counts["PA"] = counts.get("PA", 0) + 1
        elif 0x0C00 <= cp <= 0x0C7F:
            counts["TE"] = counts.get("TE", 0) + 1
        elif 0x1200 <= cp <= 0x137F:
            counts["AM"] = counts.get("AM", 0) + 1
    if not counts:
        return ""
    best_lang, best_cnt = max(counts.items(), key=lambda kv: kv[1])
    return best_lang if best_cnt >= 3 else ""


# ---------------- Step 1: XML Pre-Cleaning & OOXML Integrity Sanitizer ----------------

def unwrap_all_sdts(root: ET.Element) -> int:
    """Safely unwrap all <w:sdt> content controls innermost-first without losing a single run or paragraph."""
    unwrapped = 0
    while True:
        parent_map = {child: parent for parent in root.iter() for child in parent}
        sdts = root.findall(f".//{W}sdt")
        if not sdts:
            break
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
            if parent_tag in ("body", "tc") and raw_children and all(_local_tag(c) == "r" for c in raw_children):
                wrap_p = ET.Element(f"{W}p")
                for rc in raw_children:
                    wrap_p.append(rc)
                children_to_hoist = [wrap_p]
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
        "proofErr",
        "permStart",
        "permEnd",
    }
    UNWRAP_TAGS = {"ins", "moveTo", "customXml"}
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
                    has_content = any(_local_tag(c) not in ("rPr",) for c in r)
                    if not has_content:
                        p.remove(r)
                        continue

    for tcPr in root.findall(f".//{W}tcPr"):
        for shd in list(tcPr.findall(f"{W}shd")):
            tcPr.remove(shd)

    return rev_count, cmt_count


def _sort_ppr_children(ppr: ET.Element) -> None:
    """Sort `<w:pPr>` children strictly according to ECMA-376 CT_PPr schema sequence."""
    children = list(ppr)
    if len(children) <= 1:
        return
    ordered = sorted(
        enumerate(children),
        key=lambda pair: (_PPR_TAG_ORDER.get(_local_tag(pair[1]), 900), pair[0]),
    )
    if [c for _, c in ordered] != children:
        for ch in children:
            ppr.remove(ch)
        for _, ch in ordered:
            ppr.append(ch)


def _upsert_ppr_child(ppr: ET.Element, local_name: str) -> ET.Element:
    existing = ppr.find(f"{W}{local_name}")
    if existing is not None:
        return existing
    el = ET.Element(f"{W}{local_name}")
    target_rank = _PPR_TAG_ORDER.get(local_name, 900)
    inserted = False
    for idx, ch in enumerate(list(ppr)):
        if _PPR_TAG_ORDER.get(_local_tag(ch), 900) > target_rank:
            ppr.insert(idx, el)
            inserted = True
            break
    if not inserted:
        ppr.append(el)
    return el


def sanitize_fellowship_ooxml(root: ET.Element, valid_rel_ids: Optional[Set[str]] = None) -> None:
    """Guarantee 100% valid Word OOXML: clean bookmarks, broken rels, w14/w15/w16 attrs, mc:Ignorable, and CT_PPr/CT_RPr order."""
    # 1. Strip Word 2010/2012/2016 extension attributes (w14:paraId, w14:textId, w16cid:*, etc.)
    ext_ns_prefixes = (
        "{http://schemas.microsoft.com/office/word/2010/",
        "{http://schemas.microsoft.com/office/word/2012/",
        "{http://schemas.microsoft.com/office/word/2015/",
        "{http://schemas.microsoft.com/office/word/2016/",
        "{http://schemas.microsoft.com/office/word/2018/",
        "{http://schemas.microsoft.com/office/word/2020/",
        "{http://schemas.microsoft.com/office/word/2023/",
        "{http://schemas.microsoft.com/office/drawing/2010/",
    )
    for el in root.iter():
        for attr_k in list(el.attrib.keys()):
            if attr_k.startswith(ext_ns_prefixes):
                el.attrib.pop(attr_k, None)

    # 2. Unwrap <w:hyperlink> whose r:id is missing in target document.xml.rels
    if valid_rel_ids is not None:
        r_id_attr = f"{R_ATTR}id"
        for parent in list(root.iter()):
            for child in list(parent):
                if _local_tag(child) == "hyperlink":
                    rid = child.attrib.get(r_id_attr, "")
                    if rid and rid not in valid_rel_ids:
                        idx = list(parent).index(child)
                        hoist = list(child)
                        parent.remove(child)
                        for off, sub in enumerate(hoist):
                            parent.insert(idx + off, sub)

    # 3. Clean hidden bookmarks (_GoBack, _Hlk*, _Toc*, _*) & orphaned half-bookmarks, then renumber uniquely
    bm_starts: Dict[str, List[Tuple[ET.Element, ET.Element]]] = {}
    bm_ends: Dict[str, List[Tuple[ET.Element, ET.Element]]] = {}
    for parent in root.iter():
        for child in list(parent):
            lt = _local_tag(child)
            if lt == "bookmarkStart":
                bid = child.attrib.get(f"{W}id", "")
                bname = child.attrib.get(f"{W}name", "")
                if not bid or bname.startswith("_"):
                    parent.remove(child)
                else:
                    bm_starts.setdefault(bid, []).append((parent, child))
            elif lt == "bookmarkEnd":
                bid = child.attrib.get(f"{W}id", "")
                if not bid:
                    parent.remove(child)
                else:
                    bm_ends.setdefault(bid, []).append((parent, child))

    all_bids = set(bm_starts.keys()) | set(bm_ends.keys())
    next_wid = 1
    for bid in sorted(all_bids):
        s_list = bm_starts.get(bid, [])
        e_list = bm_ends.get(bid, [])
        if len(s_list) == 1 and len(e_list) == 1:
            new_id_str = str(next_wid)
            next_wid += 1
            s_list[0][1].attrib[f"{W}id"] = new_id_str
            e_list[0][1].attrib[f"{W}id"] = new_id_str
        else:
            for p_el, c_el in s_list + e_list:
                if c_el in list(p_el):
                    p_el.remove(c_el)

    # 4. Ensure globally unique wp:docPr id on any inline drawings
    next_docpr_id = 1
    for doc_pr in root.findall(f".//{WP}docPr"):
        doc_pr.attrib["id"] = str(next_docpr_id)
        next_docpr_id += 1

    # 5. Enforce strict ECMA-376 CT_PPr and CT_RPr child element ordering
    for ppr in root.findall(f".//{W}pPr"):
        _sort_ppr_children(ppr)
    for rpr in root.findall(f".//{W}rPr"):
        r_children = list(rpr)
        if len(r_children) > 1:
            r_ord = sorted(
                enumerate(r_children),
                key=lambda pair: (_RPR_TAG_ORDER.get(_local_tag(pair[1]), 999), pair[0]),
            )
            if [c for _, c in r_ord] != r_children:
                for ch in r_children:
                    rpr.remove(ch)
                for _, ch in r_ord:
                    rpr.append(ch)

    # 6. CRITICAL: Clean `mc:Ignorable` so NO undeclared namespace prefix remains when ElementTree serializes!
    #    (Word throws '发现无法读取的内容' if `mc:Ignorable` lists any prefix without a matching `xmlns:<prefix>`!)
    mc_ig_attr = f"{MC}Ignorable"
    if mc_ig_attr in root.attrib:
        used_uris: Set[str] = set()
        for el in root.iter():
            if el.tag.startswith("{"):
                used_uris.add(el.tag[1 : el.tag.index("}")])
            for ak in el.attrib:
                if ak != mc_ig_attr and ak.startswith("{"):
                    used_uris.add(ak[1 : ak.index("}")])
        uri_to_pfx = {uri: pfx for pfx, uri in COMMON_NAMESPACES.items()}
        declared_pfxs = {uri_to_pfx[u] for u in used_uris if u in uri_to_pfx}
        kept_ig = [p for p in root.attrib[mc_ig_attr].split() if p in declared_pfxs]
        if kept_ig:
            root.attrib[mc_ig_attr] = " ".join(kept_ig)
        else:
            root.attrib.pop(mc_ig_attr, None)


def disable_track_revisions_in_settings_xml(settings_bytes: bytes) -> bytes:
    """Remove <w:trackRevisions/> from word/settings.xml so Word opens with '修订: 关闭'."""
    try:
        s_root = _parse_safe_xml_bytes(settings_bytes)
        removed = False
        for tr_el in list(s_root.findall(f".//{W}trackRevisions")):
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


def strip_comment_and_people_rels_bytes(rels_bytes: bytes) -> Tuple[bytes, Set[str]]:
    """Strip comments* and people* relationships from document.xml.rels and return (bytes, remaining_rel_ids)."""
    valid_ids: Set[str] = set()
    try:
        ET.register_namespace("", REL_PKG_NS)
        r_root = _parse_safe_xml_bytes(rels_bytes)
        changed = False
        for rel in list(r_root):
            target = rel.attrib.get("Target", "").lower()
            rtype = rel.attrib.get("Type", "").lower()
            if "comment" in target or "comment" in rtype or "people" in target or "/people" in rtype:
                r_root.remove(rel)
                changed = True
            else:
                rid = rel.attrib.get("Id", "")
                if rid:
                    valid_ids.add(rid)
        if changed:
            return ET.tostring(r_root, encoding="utf-8", xml_declaration=True), valid_ids
        return rels_bytes, valid_ids
    except Exception:
        return rels_bytes, valid_ids


def strip_comment_and_people_content_types_bytes(ct_bytes: bytes) -> bytes:
    try:
        ET.register_namespace("", CT_NS)
        c_root = _parse_safe_xml_bytes(ct_bytes)
        changed = False
        for ov in list(c_root):
            pn = ov.attrib.get("PartName", "").lower()
            if "/word/comments" in pn or pn == "/word/people.xml":
                c_root.remove(ov)
                changed = True
        if changed:
            return ET.tostring(c_root, encoding="utf-8", xml_declaration=True)
        return ct_bytes
    except Exception:
        return ct_bytes


# ---------------- Step 2: Single-Doc Multi-Part Merge, Header Cleanup & Landscape Page ----------------

def _is_boilerplate_header_paragraph(text: str) -> bool:
    s = text.strip()
    if not s:
        return True
    for pat in BOILERPLATE_HEADER_PATTERNS:
        if pat.search(s):
            return True
    return False


def _is_all_empty_row(tr: ET.Element) -> bool:
    tcs = tr.findall(f"{W}tc")
    if not tcs:
        return True
    return all(not _get_tc_text(tc).strip() for tc in tcs)


def _is_section_heading_row(tr: ET.Element) -> bool:
    """Detect short section/question heading rows (e.g. '4．人对教会的观念看法', '什么是教会')."""
    tcs = tr.findall(f"{W}tc")
    if not tcs:
        return False
    c1 = _get_tc_text(tcs[0]).strip()
    if not c1 or len(c1) > 85:
        return False
    for r in tcs[0].findall(f".//{W}r"):
        if is_run_bold(r) or is_run_blue(r):
            return True
    if re.match(r"^(?:[0-9一二三四五六七八九十]+[．.、]|第[一二三四五六七八九十0-9]+条|[（(][一二三四五六七八九十0-9]+[)）])", c1):
        return True
    return False


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


def strip_staggered_grid_props_from_table(tbl: ET.Element) -> int:
    """Strip <w:gridSpan>, <w:gridBefore>, <w:gridAfter>, <w:wBefore>, <w:wAfter>, <w:vMerge>, <w:hMerge>
    and row-level <w:tblPrEx> on 3-cell rows so merged/normalized tables never exhibit staircase borders.
    """
    BAD_TCPR_TAGS = {"gridSpan", "gridBefore", "gridAfter", "wBefore", "wAfter", "vMerge", "hMerge"}
    BAD_TRPR_TAGS = {"gridBefore", "gridAfter", "wBefore", "wAfter"}
    removed = 0

    for tr in tbl.findall(f"{W}tr"):
        for tbl_pr_ex in list(tr.findall(f"{W}tblPrEx")):
            tr.remove(tbl_pr_ex)
            removed += 1
        trPr = tr.find(f"{W}trPr")
        if trPr is not None:
            for ch in list(trPr):
                if _local_tag(ch) in BAD_TRPR_TAGS:
                    trPr.remove(ch)
                    removed += 1

        tcs = tr.findall(f"{W}tc")
        if len(tcs) == 3:
            for tc in tcs:
                tcPr = tc.find(f"{W}tcPr")
                if tcPr is not None:
                    for ch in list(tcPr):
                        if _local_tag(ch) in BAD_TCPR_TAGS:
                            tcPr.remove(ch)
                            removed += 1
    return removed


def _append_part_rows_into_table(
    target_tbl: ET.Element,
    extra_tbl: ET.Element,
    p1_headers: List[ET.Element],
    deep_copy: bool = False,
) -> int:
    """Append rows from `extra_tbl` into `target_tbl`, deduplicating repeated headers and double-blank boundaries."""
    extra_rows = list(extra_tbl.findall(f"{W}tr"))
    appended = 0
    seen_content_in_extra = False

    for er_i, er in enumerate(extra_rows):
        if not seen_content_in_extra and er_i < 2 and _is_duplicate_header_row_in_subsequent_part(er, p1_headers):
            continue

        if not seen_content_in_extra and _is_all_empty_row(er):
            cur_target_rows = target_tbl.findall(f"{W}tr")
            last_is_empty = bool(cur_target_rows) and _is_all_empty_row(cur_target_rows[-1])
            if last_is_empty:
                continue
            # Keep leading empty row ONLY if the next row in extra_tbl is a legitimate section heading
            next_tr = extra_rows[er_i + 1] if er_i + 1 < len(extra_rows) else None
            if next_tr is None or not _is_section_heading_row(next_tr):
                continue

        if not _is_all_empty_row(er):
            seen_content_in_extra = True

        row_to_add = copy.deepcopy(er) if deep_copy else er
        target_tbl.append(row_to_add)
        appended += 1

    return appended


def merge_internal_part_tables_in_body(
    root: ET.Element,
    output_mode: str = "single_only",
) -> Tuple[bool, int]:
    """Detect and merge multiple Part tables inside a single .docx into the first <w:tbl>.

    Returns `(was_internally_merged, merged_extra_table_count)`.
    """
    body = root.find(f"{W}body")
    if body is None:
        return False, 0

    children = list(body)
    tbl_indices = [i for i, el in enumerate(children) if _local_tag(el) == "tbl"]
    if len(tbl_indices) < 2:
        return False, 0

    # Inspect headings before tbl[0] and between tbl[k] and tbl[k+1]
    has_part_heading = False
    inter_texts: List[str] = []
    for pos in range(tbl_indices[0]):
        if _local_tag(children[pos]) == "p":
            tx = _get_p_text(children[pos]).strip()
            if tx and _TITLE_PART_WORD_RE.search(tx):
                has_part_heading = True

    for t_pos_i in range(len(tbl_indices) - 1):
        seg_els = children[tbl_indices[t_pos_i] + 1 : tbl_indices[t_pos_i + 1]]
        for el in seg_els:
            if _local_tag(el) == "p":
                tx = _get_p_text(el).strip()
                if tx:
                    inter_texts.append(tx)
                    if _TITLE_PART_WORD_RE.search(tx) or re.search(r"(?:[Pp]art|[Pp]t\.?)\s*\d+", tx):
                        has_part_heading = True

    should_merge = has_part_heading or output_mode in ("merged_only", "both")
    if not should_merge:
        # Also merge if all tables are 3-column tables separated only by blank paragraphs
        if not inter_texts:
            should_merge = True

    if not should_merge:
        return False, 0

    first_tbl = children[tbl_indices[0]]
    p1_rows = first_tbl.findall(f"{W}tr")
    p1_headers = p1_rows[:2] if len(p1_rows) >= 2 else p1_rows[:1]

    extra_tbl_cnt = 0
    # Remove intermediate paragraphs & merge subsequent tables into first_tbl
    for k in range(1, len(tbl_indices)):
        prev_idx = tbl_indices[k - 1]
        cur_idx = tbl_indices[k]
        for mid_el in children[prev_idx + 1 : cur_idx]:
            if mid_el in list(body):
                body.remove(mid_el)

        extra_tbl = children[cur_idx]
        _append_part_rows_into_table(first_tbl, extra_tbl, p1_headers, deep_copy=False)
        if extra_tbl in list(body):
            body.remove(extra_tbl)
        extra_tbl_cnt += 1

    strip_staggered_grid_props_from_table(first_tbl)
    return True, extra_tbl_cnt


def enforce_landscape_page_setup(root: ET.Element) -> None:
    """Force document page setup to Landscape A4 (w=16838, h=11906, orient=landscape)."""
    body = root.find(f"{W}body")
    if body is None:
        return

    # Remove any intermediate paragraph-level section breaks so the whole document is 1 unified section
    for ppr in body.findall(f".//{W}pPr"):
        for inner_sect in list(ppr.findall(f"{W}sectPr")):
            ppr.remove(inner_sect)

    sect = body.find(f"{W}sectPr")
    if sect is None:
        sect = ET.SubElement(body, f"{W}sectPr")

    pgSz = sect.find(f"{W}pgSz")
    if pgSz is None:
        pgSz = ET.Element(f"{W}pgSz")
        pgMar = sect.find(f"{W}pgMar")
        if pgMar is not None:
            sect.insert(list(sect).index(pgMar), pgSz)
        else:
            sect.insert(0, pgSz)

    pgSz.attrib[f"{W}w"] = "16838"
    pgSz.attrib[f"{W}h"] = "11906"
    pgSz.attrib[f"{W}orient"] = "landscape"

    pgMar = sect.find(f"{W}pgMar")
    if pgMar is None:
        pgMar = ET.Element(f"{W}pgMar")
        pgMar.attrib[f"{W}top"] = "1134"
        pgMar.attrib[f"{W}right"] = "1134"
        pgMar.attrib[f"{W}bottom"] = "1134"
        pgMar.attrib[f"{W}left"] = "1134"
        pgMar.attrib[f"{W}header"] = "851"
        pgMar.attrib[f"{W}footer"] = "992"
        pgMar.attrib[f"{W}gutter"] = "0"
        sect.insert(list(sect).index(pgSz) + 1, pgMar)


def clean_body_header_and_extract_title(
    root: ET.Element,
    fallback_filename: str = "",
    force_strip_part: bool = False,
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
    candidate_ps: List[ET.Element] = []
    for p in pre_ps:
        tx = _get_p_text(p).strip()
        if not _is_boilerplate_header_paragraph(tx):
            # Exclude numbered vocabulary items under a 'Note:' header (e.g. '1. achieve a tacit understanding')
            if re.match(r"^\d+\.\s+[a-z]", tx):
                continue
            candidate_ps.append(p)

    kept_p: Optional[ET.Element] = None
    raw_title = ""
    if candidate_ps:
        kept_p = candidate_ps[-1]
        raw_title = _get_p_text(kept_p).strip()

    tbl = children[first_tbl_idx]
    rows = tbl.findall(f"{W}tr")
    if not raw_title and rows:
        next_r = rows[1] if len(rows) > 1 else None
        if _is_main_title_table_row(rows[0], next_r):
            tcs0 = rows[0].findall(f"{W}tc")
            r0_c1 = _get_tc_text(tcs0[0]).strip() if tcs0 else ""
            r1_c1 = ""
            if len(rows) > 1:
                tcs1 = rows[1].findall(f"{W}tc")
                cand_r1 = _get_tc_text(tcs1[0]).strip() if tcs1 else ""
                if 0 < len(cand_r1) <= 50 and ("交通" in cand_r1 or "202" in cand_r1):
                    r1_c1 = re.sub(r"\s+", " ", cand_r1)
            if r0_c1 and r1_c1:
                t_part = r0_c1 if r0_c1.startswith("《") else f"《{r0_c1}》"
                raw_title = f"{t_part}-{r1_c1}"
            else:
                raw_title = r0_c1
        if not raw_title:
            raw_title = clean_title_noise(Path(fallback_filename).stem, is_filename_stem=True)

    base_title, part_num = parse_title_and_part(raw_title, fallback_filename)
    if force_strip_part:
        part_num = None

    for el in children[:first_tbl_idx]:
        body.remove(el)

    if kept_p is None:
        kept_p = ET.Element(f"{W}p")

    set_clean_top_title_paragraph(
        kept_p,
        f"{base_title}-{part_num:02d}" if part_num is not None else base_title,
    )
    body.insert(0, kept_p)

    post_els = list(body)
    last_tbl_idx = max(i for i, el in enumerate(post_els) if _local_tag(el) == "tbl")
    trailing_ps = [el for el in post_els[last_tbl_idx + 1 :] if _local_tag(el) == "p"]
    if len(trailing_ps) > 1:
        for extra_p in trailing_ps[1:]:
            if not _get_p_text(extra_p).strip():
                body.remove(extra_p)

    return base_title, part_num, kept_p


def set_clean_top_title_paragraph(p: ET.Element, title_text: str):
    """Format the single top title paragraph above the table cleanly (Blue #0000FF, Bold, 14pt, Outline Level 2 '二层级')."""
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        ppr = ET.Element(f"{W}pPr")

    for ch in list(p):
        p.remove(ch)

    sp = _upsert_ppr_child(ppr, "spacing")
    for rm_attr in ("beforeLines", "afterLines", "beforeAutospacing", "afterAutospacing"):
        sp.attrib.pop(f"{W}{rm_attr}", None)
    sp.attrib[f"{W}before"] = "0"
    sp.attrib[f"{W}after"] = "60"
    sp.attrib[f"{W}line"] = "400"
    sp.attrib[f"{W}lineRule"] = "exact"

    # Set Outline Level 2 (w:val="1" in 0-indexed OOXML -> Word 导航窗格显示为 2 级标题)
    ol = _upsert_ppr_child(ppr, "outlineLvl")
    ol.attrib[f"{W}val"] = "1"
    _sort_ppr_children(ppr)

    p.append(ppr)
    r = ET.SubElement(p, f"{W}r")
    rpr = ET.SubElement(r, f"{W}rPr")
    rf = ET.SubElement(rpr, f"{W}rFonts")
    rf.attrib[f"{W}ascii"] = "Georgia"
    rf.attrib[f"{W}hAnsi"] = "Georgia"
    rf.attrib[f"{W}eastAsia"] = "宋体"
    rf.attrib[f"{W}cs"] = "Georgia"
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


# ---------------- Step 3: Cell Trailing Spaces / Empty Paragraphs & Smart Curly Quotes ----------------

def _p_has_drawing_or_object(p: ET.Element) -> bool:
    for el in p.iter():
        if _local_tag(el) in ("drawing", "pict", "object"):
            return True
    return False


def strip_trailing_spaces_in_paragraph(p: ET.Element) -> int:
    """Strip trailing spaces (' ', '\\t', '\\u3000', '\\u00a0') and trailing soft <w:br/> from <w:p>."""
    removed_chars = 0
    # Strip trailing soft <w:br/> inside last run(s) when no text follows
    while True:
        runs = p.findall(f"{W}r")
        if not runs:
            break
        last_r = runs[-1]
        non_rpr = [c for c in last_r if _local_tag(c) != "rPr"]
        if not non_rpr:
            p.remove(last_r)
            continue
        last_child = non_rpr[-1]
        lt = _local_tag(last_child)
        if lt == "br" and last_child.attrib.get(f"{W}type", "") not in ("page", "column"):
            last_r.remove(last_child)
            removed_chars += 1
            if not [c for c in last_r if _local_tag(c) != "rPr"]:
                p.remove(last_r)
            continue
        if lt == "t":
            orig = last_child.text or ""
            trimmed = orig.rstrip(" \t\u3000\u00a0")
            if len(trimmed) < len(orig):
                removed_chars += len(orig) - len(trimmed)
                last_child.text = trimmed
                if trimmed.startswith(" ") or trimmed.endswith(" "):
                    last_child.attrib[XML_SPACE] = "preserve"
                else:
                    last_child.attrib.pop(XML_SPACE, None)
            if trimmed == "":
                last_r.remove(last_child)
                if not [c for c in last_r if _local_tag(c) != "rPr"]:
                    p.remove(last_r)
                continue
        break
    return removed_chars


def clean_table_trailing_spaces_and_empty_paragraphs(tbl: ET.Element) -> Tuple[int, int]:
    """In every non-empty table cell: strip trailing spaces on each paragraph, and remove trailing empty <w:p>s
    at the bottom of the cell, while keeping intentional whole-row blank separator rows (['', '', '']).
    """
    stripped_p_cnt = 0
    removed_empty_p_cnt = 0

    for tr in tbl.findall(f"{W}tr"):
        tcs = tr.findall(f"{W}tc")
        if _is_all_empty_row(tr):
            # Intentional blank separator row: keep row, ensure exactly 1 clean empty <w:p> per cell
            for tc in tcs:
                ps = tc.findall(f"{W}p")
                for p in ps:
                    strip_trailing_spaces_in_paragraph(p)
                while len(ps) > 1:
                    tc.remove(ps.pop())
                    removed_empty_p_cnt += 1
            continue

        for tc in tcs:
            ps = tc.findall(f"{W}p")
            for p in ps:
                if strip_trailing_spaces_in_paragraph(p) > 0:
                    stripped_p_cnt += 1
            # Remove trailing empty <w:p> elements at bottom of cell (always keeping >= 1 <w:p> for valid OOXML)
            while len(ps) > 1 and not _get_p_text(ps[-1]).strip() and not _p_has_drawing_or_object(ps[-1]):
                tc.remove(ps.pop())
                removed_empty_p_cnt += 1

    return stripped_p_cnt, removed_empty_p_cnt


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
            if prev_c and prev_c.isalnum() and next_c and next_c.isalnum():
                out_chars[i] = "’"
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


# ---------------- Step 4: XML Run Splitting & Strict Ordered <w:rPr> Setters ----------------

def _unwrap_inline_wrappers_in_p(p: ET.Element):
    for parent in list(p.iter()):
        for ch in list(parent):
            if _local_tag(ch) in ("smartTag", "fldSimple"):
                idx = list(parent).index(ch)
                subs = list(ch)
                parent.remove(ch)
                for off, s in enumerate(subs):
                    parent.insert(idx + off, s)


def _normalize_single_t_per_run(p: ET.Element):
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


def _upsert_rpr_child(rpr: ET.Element, local_name: str) -> ET.Element:
    existing = rpr.find(f"{W}{local_name}")
    if existing is not None:
        return existing
    new_el = ET.Element(f"{W}{local_name}")
    target_rank = _RPR_TAG_ORDER.get(local_name, 999)
    insert_pos = len(rpr)
    for idx, ch in enumerate(rpr):
        ch_rank = _RPR_TAG_ORDER.get(_local_tag(ch), 999)
        if ch_rank > target_rank:
            insert_pos = idx
            break
    rpr.insert(insert_pos, new_el)
    return new_el


def set_rpr_font_family(rpr: ET.Element, font_name: str, col_idx: int) -> None:
    """Apply standard font family onto `<w:rPr>/<w:rFonts>`, stripping theme font overrides."""
    if not font_name:
        return
    rf = _upsert_rpr_child(rpr, "rFonts")
    for theme_attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        rf.attrib.pop(f"{W}{theme_attr}", None)
    if col_idx == 0:
        rf.attrib[f"{W}ascii"] = font_name
        rf.attrib[f"{W}hAnsi"] = font_name
        rf.attrib[f"{W}eastAsia"] = font_name
        rf.attrib[f"{W}cs"] = font_name
    elif col_idx == 1:
        rf.attrib[f"{W}ascii"] = font_name
        rf.attrib[f"{W}hAnsi"] = font_name
        rf.attrib[f"{W}eastAsia"] = "宋体"
        rf.attrib[f"{W}cs"] = font_name
    else:
        rf.attrib[f"{W}ascii"] = font_name
        rf.attrib[f"{W}hAnsi"] = font_name
        rf.attrib[f"{W}eastAsia"] = font_name
        rf.attrib[f"{W}cs"] = font_name


def set_run_font_family(r: ET.Element, font_name: str, col_idx: int) -> None:
    if not font_name:
        return
    set_rpr_font_family(_ensure_rpr(r), font_name, col_idx)


def set_rpr_font_size(rpr: ET.Element, half_pt_str: str) -> None:
    for tag_name in ("sz", "szCs"):
        el = _upsert_rpr_child(rpr, tag_name)
        el.attrib[f"{W}val"] = half_pt_str


def set_run_font_size(r: ET.Element, half_pt_str: str):
    set_rpr_font_size(_ensure_rpr(r), half_pt_str)


def set_run_bold(r: ET.Element, bold: bool):
    rpr = _ensure_rpr(r)
    for tag_name in ("b", "bCs"):
        el = rpr.find(f"{W}{tag_name}")
        if bold:
            el = _upsert_rpr_child(rpr, tag_name)
            el.attrib.pop(f"{W}val", None)
        else:
            if el is not None:
                rpr.remove(el)


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


def is_red_hex(val: str) -> bool:
    v = (val or "").strip().lstrip("#").lower()
    if v in ("red", "darkred"):
        return True
    if len(v) == 6 and all(c in "0123456789abcdef" for c in v):
        r_c = int(v[0:2], 16)
        g_c = int(v[2:4], 16)
        b_c = int(v[4:6], 16)
        return r_c >= 150 and r_c > g_c * 1.8 and r_c > b_c * 1.8
    return False


def get_run_color_kind(r: ET.Element) -> str:
    """Return 'blue', 'red', or 'none' for a <w:r>."""
    rpr = r.find(f"{W}rPr")
    if rpr is None:
        return "none"
    c_el = rpr.find(f"{W}color")
    if c_el is None:
        return "none"
    raw = c_el.attrib.get(f"{W}val", "")
    if is_blue_hex(raw):
        return "blue"
    if is_red_hex(raw):
        return "red"
    return "none"


def set_run_color_kind(r: ET.Element, color_kind: str) -> None:
    """Set run color strictly to 'blue' (#0000FF), 'red' (#FF0000), or 'none' (default black/auto)."""
    rpr = _ensure_rpr(r)
    c_el = rpr.find(f"{W}color")
    if color_kind == "blue":
        c_el = _upsert_rpr_child(rpr, "color")
        c_el.attrib.clear()
        c_el.attrib[f"{W}val"] = "0000FF"
    elif color_kind == "red":
        c_el = _upsert_rpr_child(rpr, "color")
        existing_hex = c_el.attrib.get(f"{W}val", "")
        c_el.attrib.clear()
        c_el.attrib[f"{W}val"] = existing_hex.upper() if is_red_hex(existing_hex) else "FF0000"
    else:
        if c_el is not None:
            rpr.remove(c_el)


def set_run_blue(r: ET.Element, blue: bool):
    rpr = _ensure_rpr(r)
    c_el = rpr.find(f"{W}color")
    if blue:
        c_el = _upsert_rpr_child(rpr, "color")
        c_el.attrib.clear()
        c_el.attrib[f"{W}val"] = "0000FF"
    else:
        if c_el is not None and is_blue_hex(c_el.attrib.get(f"{W}val", "")):
            rpr.remove(c_el)


def is_run_blue(r: ET.Element) -> bool:
    return get_run_color_kind(r) == "blue"


def is_run_red(r: ET.Element) -> bool:
    return get_run_color_kind(r) == "red"


def is_run_bold(r: ET.Element) -> bool:
    rpr = r.find(f"{W}rPr")
    if rpr is None:
        return False
    b_el = rpr.find(f"{W}b")
    if b_el is None:
        return False
    return b_el.attrib.get(f"{W}val", "true").lower() not in ("0", "false", "off")


def get_run_sz_val(r: ET.Element) -> Optional[str]:
    rpr = r.find(f"{W}rPr")
    if rpr is None:
        return None
    sz_el = rpr.find(f"{W}sz")
    if sz_el is None:
        return None
    return sz_el.attrib.get(f"{W}val")


# ---------------- Step 5: Multi-Script Bible & God's Word ('神话的摘自') Citation Detector ----------------

_MULTI_DIGIT_CLASS = (
    r"0-9"
    r"\u0966-\u096f"  # Devanagari ०-९
    r"\u09e6-\u09ef"  # Bengali ০-৯
    r"\u1040-\u1049"  # Myanmar ၀-၉
    r"\u17e0-\u17e9"  # Khmer ០-៩
    r"\u0e50-\u0e59"  # Thai ๐-๙
    r"\u0ed0-\u0ed9"  # Lao ໐-໙
    r"\u0660-\u0669"  # Arabic-Indic ٠-٩
    r"\u06f0-\u06f9"  # Extended Arabic-Indic ۰-۹
)
_ANY_MULTI_DIGIT_RE = re.compile(f"[{_MULTI_DIGIT_CLASS}]")

# Standalone Bible chapter:verse inside (...) / （...） / [...]:
# e.g. '(βλ. Ησαΐας 6:8)', '(Matthew 7:23)', '(Առակներ 10։21)', '(यूहन्ना ३:१६)', '(太7:23)'
_STANDALONE_COLON_VERSE_INNER_RE = re.compile(
    rf"^[^()（）\[\]]*?[^\W\d_][^()（）\[\]]*[:：։]\s*[{_MULTI_DIGIT_CLASS}][{_MULTI_DIGIT_CLASS}~–—\-\.,;\s]{{0,18}}$",
    re.UNICODE,
)
# European comma chapter,verse style e.g. '(Mt 7, 23)', '(1 Kor 13, 4-7)' — requires chapter digits BEFORE comma
_STANDALONE_COMMA_VERSE_INNER_RE = re.compile(
    rf"^[^()（）\[\]]*?[^\W\d_][^()（）\[\]]*?\s*[{_MULTI_DIGIT_CLASS}]{{1,3}}\s*,\s*([{_MULTI_DIGIT_CLASS}]{{1,3}}(?:[~–—\-.\s]+[{_MULTI_DIGIT_CLASS}]{{1,3}})?)\s*$",
    re.UNICODE,
)

# Explicit source-citation prefixes inside (...) / （...） / [...] even without preceding quote
_EXPLICIT_SOURCE_INNER_RE = re.compile(
    r"^(?:摘自|选自|引自|参看|参(?=[《\sA-Za-z\u4e00-\u9fff])|《(?:话|話)"
    r"|(?:cf\.|see|from|βλ\.|πρβλ\.|см\.|zob\.|por\.|vgl\.)\s+\S+"
    r"|[«\"“„]?\s*(?:The\s+Word|La\s+Palabra|A\s+Palavra|La\s+Parole|Das\s+Wort|La\s+Parola|Ο\s+Λόγος|Слово|Słowo|Ny\s+Teny|Neno|พระวจนะ|वचन)\b.*(?:Vol|Tome|Tomo|Bd|Boky|Juz|เล่ม|खंड|τόμ|том|t\.|\d))",
    re.I,
)

_CITATION_KEYWORDS_RE = re.compile(
    r"(?:摘自|选自|引自|vol\.?|bd\.?|tome|tomo|boky|juz\.?|เล่ม|खंड|τόμ\.?|том|ks\.|卷|篇|章|话|話"
    r"|word|wort|palabras?|palavra|parole|слово|słowo|λόγος|พระวจนะ|वचन|[:：։])",
    re.I,
)

_SHORT_RESPONSE_EXCLUSION_RE = re.compile(
    r"^(?:pl|s|yes|no|right|amen|true|false|correct|ok|okay|it\s+is|they\s+are|we\s+do|we\s+cannot|we\s+can"
    r"|eny|tsia|ndiyo|hapana|la|sí|si|oui|non|ja|nein|да|нет|так|ні|jest|nie|zgadza\s+się|हाँ|नहीं|ใช่|ไม่ใช่"
    r"|是|不是|对|不对|没有|有|能|不能|行|不行|好|阿们|阿們|实际|不实际|可以|不可以|算|不算|重要|可怕)[.。!！?？\s]*$",
    re.I,
)


def _is_valid_citation_inner(inner_text: str) -> bool:
    s = inner_text.strip()
    if len(s) < 2:
        return False
    if _SHORT_RESPONSE_EXCLUSION_RE.match(s):
        return False
    if s.lower() in ("pl", "s", "1", "2", "3", "i", "ii", "iii", "iv"):
        return False
    if _ANY_MULTI_DIGIT_RE.search(s):
        return True
    if len(s) >= 4 and _CITATION_KEYWORDS_RE.search(s):
        return True
    return False


def _is_standalone_bracket_citation_inner(inner_text: str) -> bool:
    s = inner_text.strip()
    if len(s) < 3 or len(s) > 220:
        return False
    if _SHORT_RESPONSE_EXCLUSION_RE.match(s):
        return False
    if _STANDALONE_COLON_VERSE_INNER_RE.match(s):
        return True
    m_comma = _STANDALONE_COMMA_VERSE_INNER_RE.match(s)
    if m_comma:
        verse_part = m_comma.group(1).strip()
        if not re.match(r"^(?:19|20)\d{2}$", verse_part):
            return True
    if _EXPLICIT_SOURCE_INNER_RE.search(s):
        return True
    return False


def _find_balanced_bracket_end(text: str, open_pos: int) -> int:
    """Given `text[open_pos]` in `BRACKET_PAIRS`, return index right AFTER matching closing bracket, or -1."""
    open_br = text[open_pos]
    close_br = BRACKET_PAIRS.get(open_br)
    if not close_br:
        return -1
    depth = 1
    k = open_pos + 1
    n = len(text)
    while k < n and depth > 0:
        if text[k] == open_br:
            depth += 1
        elif text[k] == close_br:
            depth -= 1
        k += 1
    return k if depth == 0 else -1


def find_post_quote_citation_spans(full_text: str) -> List[Tuple[int, int]]:
    """Return sorted, non-overlapping [(start, end), ...] character spans of scripture & God's Word citations.

    Supports:
    - Multi-script digits [0-9०-९০-৯၀-၉០-៩๐-๙...] and colons/commas inside (...)/（...）/[...], even if
      preceded by '.' or without any preceding quotation mark (e.g. '(βλ. Ησαΐας 6:8).').
    - Explicit source citations like '（摘自《...》）', '(The Word, Vol. 1...)'.
    - Post-quote citations right after closing quotes (allowing optional '.'/'。' before bracket or '《话・...》').
    - Tortoise-shell brackets '〔...〕' and trailing '——摘自《...》'.
    Strictly keeps trailing external sentence periods (., 。, ։) OUTSIDE the returned span.
    """
    raw_spans: List[Tuple[int, int]] = []
    n = len(full_text)
    i = 0

    while i < n:
        ch = full_text[i]

        # Case A: Tortoise-shell brackets 〔...〕
        if ch == "〔":
            close_idx = full_text.find("〕", i + 1)
            if close_idx != -1 and len(full_text[i + 1 : close_idx].strip()) >= 2:
                raw_spans.append((i, close_idx + 1))
                i = close_idx + 1
                continue

        # Case B: Post-quote bracket or 《话・...》 citation (allowing optional spaces & optional '.'/'。' after quote)
        if ch in CLOSING_QUOTE_CHARS:
            j = i + 1
            while j < n and full_text[j] in " \t\u00a0\u3000":
                j += 1
            j_after_dot = j
            if j_after_dot < n and full_text[j_after_dot] in ".。":
                j_after_dot += 1
                while j_after_dot < n and full_text[j_after_dot] in " \t\u00a0\u3000":
                    j_after_dot += 1

            for cand_j in (j, j_after_dot):
                if cand_j < n and full_text[cand_j] in BRACKET_PAIRS:
                    k = _find_balanced_bracket_end(full_text, cand_j)
                    if k != -1:
                        inner = full_text[cand_j + 1 : k - 1]
                        if _is_valid_citation_inner(inner):
                            raw_spans.append((cand_j, k))
                            i = k
                            break
            else:
                if j_after_dot < n and full_text[j_after_dot] == "《":
                    end_book = full_text.find("》", j_after_dot + 1)
                    if end_book != -1:
                        book_inner = full_text[j_after_dot + 1 : end_book]
                        if any(tag in book_inner for tag in ("话", "話", "卷", "・", "·")):
                            raw_spans.append((j_after_dot, end_book + 1))
                            i = end_book + 1
                            continue
                i += 1
                continue
            continue

        # Case C: Standalone bracketed Bible / God's Word citation anywhere (even without preceding quote!)
        if ch in BRACKET_PAIRS:
            k = _find_balanced_bracket_end(full_text, i)
            if k != -1:
                inner = full_text[i + 1 : k - 1]
                if _is_standalone_bracket_citation_inner(inner):
                    raw_spans.append((i, k))
                    i = k
                    continue

        i += 1

    # Case D: Trailing '——摘自《...》' or '——《话·...》' at end of paragraph
    m_dash_zh = re.search(r"(?:——|—|--)\s*(?:摘自)?《[^》]{2,120}》\s*$", full_text)
    if m_dash_zh:
        raw_spans.append((m_dash_zh.start(), m_dash_zh.end()))

    if not raw_spans:
        return []

    raw_spans.sort(key=lambda x: (x[0], -(x[1] - x[0])))
    merged: List[Tuple[int, int]] = []
    for s0, s1 in raw_spans:
        if merged and s0 < merged[-1][1]:
            if s1 > merged[-1][1]:
                merged[-1] = (merged[-1][0], s1)
        else:
            merged.append((s0, s1))
    return merged


def extract_paragraph_source_citation_hints(
    p_zh: Optional[ET.Element],
    p_en: Optional[ET.Element],
) -> Dict:
    """Inspect Col 1 & Col 2 BEFORE font-size normalization to detect God's Word / source citations
    (including runs originally marked w:sz='21', '摘自', '《话...》', '(The Word, ...)', '(from ...)').

    Returns:
      {
        "whole_para_citation": bool,
        "whole_para_blue": bool,
        "trailing_citation": bool,
        "trailing_blue": bool,
        "post_quote_ratios": List[Tuple[float, bool]],  # (rel_pos_in_en_or_zh, is_blue)
      }
    """
    info = {
        "whole_para_citation": False,
        "whole_para_blue": False,
        "trailing_citation": False,
        "trailing_blue": False,
        "post_quote_ratios": [],
    }

    zh_txt = _get_p_text(p_zh) if p_zh is not None else ""
    en_txt = _get_p_text(p_en) if p_en is not None else ""

    def _collect_sz21_and_color(p_el: Optional[ET.Element], full_len: int):
        if p_el is None or full_len == 0:
            return [False] * full_len, [False] * full_len
        is_21 = [False] * full_len
        is_bl = [False] * full_len
        pos = 0
        for r in p_el.findall(f"{W}r"):
            seg = "".join(t.text or "" for t in r.findall(f"{W}t"))
            if not seg:
                continue
            sz_v = get_run_sz_val(r)
            bl_v = is_run_blue(r)
            end_p = min(full_len, pos + len(seg))
            for k in range(pos, end_p):
                if sz_v == "21":
                    is_21[k] = True
                if bl_v:
                    is_bl[k] = True
            pos = end_p
        return is_21, is_bl

    zh_21, zh_bl = _collect_sz21_and_color(p_zh, len(zh_txt))
    en_21, en_bl = _collect_sz21_and_color(p_en, len(en_txt))

    # 1. Check whole-paragraph standalone source line (e.g. '——摘自《话在肉身显现·作工与进入》' or all sz=21 source)
    zh_strip = zh_txt.strip()
    en_strip = en_txt.strip()
    if zh_strip and len(zh_strip) <= 160:
        zh_non_ws = [k for k, c in enumerate(zh_txt) if not c.isspace()]
        all_zh_21 = bool(zh_non_ws) and all(zh_21[k] for k in zh_non_ws)
        if re.match(r"^(?:——|—|--)?\s*(?:[（(\[])?\s*摘自《.+》[)）\]]?\s*$", zh_strip) or (
            all_zh_21 and ("摘自" in zh_strip or "《话" in zh_strip or "《話" in zh_strip)
        ):
            en_non_ws = [k for k, c in enumerate(en_txt) if not c.isspace()]
            whole_bl = (any(en_bl[k] for k in en_non_ws) if en_non_ws else any(zh_bl[k] for k in zh_non_ws))
            info["whole_para_citation"] = True
            info["whole_para_blue"] = whole_bl
            return info

    # 2. Check trailing citation at the end of Col 1 or Col 2
    trailing_hit = False
    trailing_blue_val = False

    # Check Col 2 trailing bracket or sz=21 tail
    if en_strip:
        en_spans = find_post_quote_citation_spans(en_txt)
        # Also check any trailing balanced (...) / [...] at end of en_txt whose runs have sz=21 or 'from '/'The Word'
        en_rstripped = en_txt.rstrip(" \t\u00a0\u3000.。")
        if en_rstripped and en_rstripped[-1] in CLOSING_TO_OPENING_BRACKET:
            close_ch = en_rstripped[-1]
            open_ch = CLOSING_TO_OPENING_BRACKET[close_ch]
            depth = 1
            p_idx = len(en_rstripped) - 2
            while p_idx >= 0 and depth > 0:
                if en_rstripped[p_idx] == close_ch:
                    depth += 1
                elif en_rstripped[p_idx] == open_ch:
                    depth -= 1
                p_idx -= 1
            if depth == 0:
                b_s = p_idx + 1
                b_e = len(en_rstripped)
                inner_en = en_txt[b_s + 1 : b_e - 1].strip()
                has_21_in_br = any(en_21[k] for k in range(b_s, b_e) if not en_txt[k].isspace())
                looks_source = bool(
                    re.search(r"\b(?:The\s+Word|Vol\.\s*\d|Chapter\s+\d|Appears\s+in\s+the\s+Flesh)\b|^from\s+", inner_en, re.I)
                )
                if (has_21_in_br or looks_source or (b_s, b_e) in en_spans) and not _SHORT_RESPONSE_EXCLUSION_RE.match(inner_en) and len(inner_en) >= 3:
                    trailing_hit = True
                    trailing_blue_val = any(en_bl[k] for k in range(b_s, b_e) if not en_txt[k].isspace())

        for s0, s1 in en_spans:
            sp_bl = any(en_bl[k] for k in range(s0, s1) if not en_txt[k].isspace())
            if s1 >= len(en_txt.rstrip(" \t\u00a0\u3000.。")):
                trailing_hit = True
                trailing_blue_val = sp_bl
            else:
                info["post_quote_ratios"].append(((s0 + s1) / (2.0 * max(1, len(en_txt))), sp_bl))

    # Check Col 1 trailing '（摘自...）' / '《话・...》' / sz=21 tail
    if zh_strip:
        m_zh_tail = re.search(
            r"(?:[（(\[]\s*(?:摘自|选自|引自)?《?[^()（）\[\]]{2,150}》?\s*[)）\]]|(?:——|—|--)?\s*(?:摘自)?《(?:话|話)[^》]{2,150}》)[.。]?\s*$",
            zh_txt,
        )
        if m_zh_tail:
            s0, s1 = m_zh_tail.start(), m_zh_tail.end()
            seg_zh = zh_txt[s0:s1].strip(" .。()（）[]")
            has_zh_21 = any(zh_21[k] for k in range(s0, s1) if not zh_txt[k].isspace())
            if ("摘自" in seg_zh or "《话" in zh_txt[s0:s1] or "《話" in zh_txt[s0:s1] or has_zh_21) and not _SHORT_RESPONSE_EXCLUSION_RE.match(seg_zh):
                if not trailing_hit:
                    trailing_hit = True
                    # Prefer Col 2's color at paragraph tail if Col 2 exists, else Col 1's color
                    if en_txt.strip():
                        tail_start_en = max(0, len(en_txt) - 40)
                        trailing_blue_val = any(en_bl[k] for k in range(tail_start_en, len(en_txt)) if not en_txt[k].isspace())
                    else:
                        trailing_blue_val = any(zh_bl[k] for k in range(s0, s1) if not zh_txt[k].isspace())

        if not info["post_quote_ratios"]:
            for s0, s1 in find_post_quote_citation_spans(zh_txt):
                if s1 < len(zh_txt.rstrip(" \t\u00a0\u3000.。")):
                    sp_bl = any(zh_bl[k] for k in range(s0, s1) if not zh_txt[k].isspace())
                    info["post_quote_ratios"].append(((s0 + s1) / (2.0 * max(1, len(zh_txt))), sp_bl))

    info["trailing_citation"] = trailing_hit
    info["trailing_blue"] = trailing_blue_val
    return info


def find_col3_citation_spans_with_hints(
    fl_txt: str,
    hints: Optional[Dict] = None,
) -> List[Tuple[int, int, Optional[bool]]]:
    """Return [(start, end, override_blue_or_None), ...] citation spans in Col 3 `fl_txt`."""
    if not fl_txt:
        return []

    if hints and hints.get("whole_para_citation"):
        return [(0, len(fl_txt), bool(hints.get("whole_para_blue", False)))]

    base_spans = find_post_quote_citation_spans(fl_txt)
    out_map: Dict[Tuple[int, int], Optional[bool]] = {(s0, s1): None for s0, s1 in base_spans}

    if hints:
        # 1. Trailing source citation from Col 1 / Col 2 -> match trailing balanced (...) / （...） / [...] in Col 3
        if hints.get("trailing_citation"):
            t_blue = bool(hints.get("trailing_blue", False))
            rstripped = fl_txt.rstrip(" \t\u00a0\u3000.。։।॥")
            matched_tail: Optional[Tuple[int, int]] = None
            if rstripped and rstripped[-1] in CLOSING_TO_OPENING_BRACKET:
                close_ch = rstripped[-1]
                open_ch = CLOSING_TO_OPENING_BRACKET[close_ch]
                depth = 1
                p_idx = len(rstripped) - 2
                while p_idx >= 0 and depth > 0:
                    if rstripped[p_idx] == close_ch:
                        depth += 1
                    elif rstripped[p_idx] == open_ch:
                        depth -= 1
                    p_idx -= 1
                if depth == 0:
                    b_s = p_idx + 1
                    b_e = len(rstripped)
                    inner = fl_txt[b_s + 1 : b_e - 1].strip()
                    if len(inner) >= 3 and not _SHORT_RESPONSE_EXCLUSION_RE.match(inner):
                        matched_tail = (b_s, b_e)
            if matched_tail is not None:
                out_map[matched_tail] = t_blue
            else:
                # Check if base_spans already had a span near the very end and attach t_blue
                for s0, s1 in list(out_map.keys()):
                    if s1 >= len(rstripped) - 2:
                        out_map[(s0, s1)] = t_blue

        # 2. Mid-paragraph post-quote source citations from Col 1 / Col 2 (even when Col 3 book title has no digits)
        pq_ratios: List[Tuple[float, bool]] = hints.get("post_quote_ratios") or []
        if pq_ratios:
            n = len(fl_txt)
            cand_pq_brackets: List[Tuple[int, int]] = []
            idx = 0
            while idx < n:
                if fl_txt[idx] in CLOSING_QUOTE_CHARS:
                    j = idx + 1
                    while j < n and fl_txt[j] in " \t\u00a0\u3000.。":
                        j += 1
                    if j < n and fl_txt[j] in BRACKET_PAIRS:
                        k = _find_balanced_bracket_end(fl_txt, j)
                        if k != -1:
                            inner = fl_txt[j + 1 : k - 1].strip()
                            if len(inner) >= 5 and not _SHORT_RESPONSE_EXCLUSION_RE.match(inner):
                                cand_pq_brackets.append((j, k))
                                idx = k
                                continue
                idx += 1

            for ratio, pq_blue in pq_ratios:
                best_br = None
                best_dist = 0.38
                for b_s, b_e in cand_pq_brackets:
                    b_ratio = (b_s + b_e) / (2.0 * max(1, n))
                    dist = abs(b_ratio - ratio)
                    if dist < best_dist:
                        best_dist = dist
                        best_br = (b_s, b_e)
                if best_br is not None:
                    out_map[best_br] = pq_blue

    # Merge overlapping intervals
    items = sorted(out_map.items(), key=lambda kv: (kv[0][0], -(kv[0][1] - kv[0][0])))
    merged: List[Tuple[int, int, Optional[bool]]] = []
    for (s0, s1), bl_override in items:
        if merged and s0 < merged[-1][1]:
            prev_s0, prev_s1, prev_bl = merged[-1]
            merged[-1] = (prev_s0, max(prev_s1, s1), bl_override if bl_override is not None else prev_bl)
        else:
            merged.append((s0, s1, bl_override))
    return merged


# ---------------- Step 6: Decoupled Col 3 Blue, Red & Bold Synchronization ----------------

def _extract_char_styles(p: ET.Element, include_red: bool = True) -> Tuple[str, List[str], List[bool]]:
    """Return `(full_text, char_color_list, char_bold_list)` where color is 'blue' | 'red' | 'none'."""
    chunks: List[str] = []
    colors: List[str] = []
    bolds: List[bool] = []
    for r in p.findall(f"{W}r"):
        seg = "".join(t.text or "" for t in r.findall(f"{W}t"))
        if not seg:
            continue
        ck = get_run_color_kind(r)
        if ck == "red" and not include_red:
            ck = "none"
        bd = is_run_bold(r)
        chunks.append(seg)
        colors.extend([ck] * len(seg))
        bolds.extend([bd] * len(seg))
    return "".join(chunks), colors, bolds


def _compress_style_segments(text: str, colors: List[str], bolds: List[bool], start: int = 0) -> List[Tuple[str, str, bool]]:
    """Compress per-char arrays from `start` into contiguous `(substr, color_kind, is_bold)` segments."""
    segs: List[Tuple[str, str, bool]] = []
    n = len(text)
    if start >= n:
        return segs
    cur_s = start
    cur_c = colors[start]
    cur_b = bolds[start]
    for i in range(start + 1, n):
        if colors[i] != cur_c or bolds[i] != cur_b:
            segs.append((text[cur_s:i], cur_c, cur_b))
            cur_s = i
            cur_c = colors[i]
            cur_b = bolds[i]
    segs.append((text[cur_s:n], cur_c, cur_b))
    return segs


def find_top_level_quotes(text: str, start: int = 0) -> List[Tuple[int, int, int, int]]:
    """Find top-level quotation spans in `text[start:]` -> `[(open_idx, inner_start, inner_end, close_end), ...]`.
    Supports “...”, «...», „...“ / „...”, 「...」, and straight "...".
    """
    QUOTE_OPEN_TO_CLOSERS = {
        "“": ("”",),
        "«": ("»",),
        "„": ("“", "”"),
        "「": ("」",),
        '"': ('"',),
    }
    spans: List[Tuple[int, int, int, int]] = []
    n = len(text)
    i = max(0, start)
    while i < n:
        ch = text[i]
        if ch in QUOTE_OPEN_TO_CLOSERS:
            closers = QUOTE_OPEN_TO_CLOSERS[ch]
            if ch == '"':
                close_pos = text.find('"', i + 1)
                if close_pos != -1:
                    spans.append((i, i + 1, close_pos, close_pos + 1))
                    i = close_pos + 1
                    continue
            else:
                depth = 1
                k = i + 1
                while k < n and depth > 0:
                    if text[k] == ch and ch not in closers:
                        depth += 1
                    elif text[k] in closers:
                        depth -= 1
                        if depth == 0:
                            break
                    k += 1
                if k < n and depth == 0:
                    spans.append((i, i + 1, k, k + 1))
                    i = k + 1
                    continue
                if ch == "“":
                    # Fallback when translator opened with `“` and closed with straight `"`
                    fb_close = text.find('"', i + 1)
                    next_open = text.find("“", i + 1)
                    if fb_close != -1 and (next_open == -1 or fb_close < next_open):
                        spans.append((i, i + 1, fb_close, fb_close + 1))
                        i = fb_close + 1
                        continue
        i += 1
    return spans


def _detect_leading_styled_prefix(
    en_txt: str,
    en_colors: List[str],
    en_bolds: List[bool],
    fl_txt: str,
) -> Optional[Tuple[int, int, str, bool]]:
    """Detect leading styled prefix in Col 2 (e.g. 'Question 10:', 'Answer:', '1.', 'I.') and match in Col 3.

    Returns `(en_pfx_end, fl_pfx_end, pfx_color, pfx_bold)` or None.
    """
    n_en = len(en_txt)
    if n_en == 0 or not fl_txt.strip():
        return None

    i0 = 0
    while i0 < n_en and en_txt[i0].isspace():
        i0 += 1
    if i0 >= n_en:
        return None

    c0, b0 = en_colors[i0], en_bolds[i0]
    if c0 == "none" and not b0:
        return None

    # Walk contiguous non-plain prefix in en_txt (allowing spaces inside 'Question 10:')
    i_end = i0
    while i_end < n_en:
        if en_txt[i_end].isspace():
            # Peek next non-space char
            pk = i_end
            while pk < n_en and en_txt[pk].isspace():
                pk += 1
            if pk < n_en and (en_colors[pk], en_bolds[pk]) == (c0, b0) and (pk - i0) <= 45:
                i_end = pk
                continue
            break
        if (en_colors[i_end], en_bolds[i_end]) == (c0, b0):
            i_end += 1
        elif en_txt[i_end] in ":：" and i_end > i0:
            # Include trailing colon even if colon had blue-only or plain style right after bold 'Question 10'
            i_end += 1
            break
        else:
            break

    if i_end >= n_en:
        return None  # Whole paragraph styled -> handled by whole-paragraph branch

    # Verify plain text immediately follows the prefix in Col 2
    probe = i_end
    while probe < n_en and en_txt[probe].isspace():
        probe += 1
    if probe >= n_en or (en_colors[probe], en_bolds[probe]) != ("none", False):
        return None

    pfx_str = en_txt[i0:i_end].strip()
    if not pfx_str or len(pfx_str) > 48:
        return None

    # Pattern A: Colon prefix ('Question 10:', 'Q10:', 'Answer:', 'Item 1:', etc.)
    if pfx_str.endswith((":", "：")) or re.match(r"^(?:Question|Q|Answer|A|Problem|Item)\s*\d+\s*[.:：]?$", pfx_str, re.I):
        m_fl_colon = re.match(r"^(\s*[^:：\n]{1,60}[:：])", fl_txt)
        if m_fl_colon:
            return i_end, m_fl_colon.end(1), c0, b0
        m_fl_num_dot = re.match(r"^(\s*[^\d\n]{0,35}\d+\s*[.．])", fl_txt)
        if m_fl_num_dot:
            return i_end, m_fl_num_dot.end(1), c0, b0
        return None

    # Pattern B: Bare numbered / Roman / parenthesized subheading marker ('1.', 'I.', '(1)', '（一）')
    if re.match(r"^(?:[IVXLCDM]+\.|[0-9]+[.)]|[（(][0-9一二三四五六七八九十]+[)）])$", pfx_str):
        m_fl_mk = re.match(r"^(\s*(?:[IVXLCDM]+[．.]|[0-9]+[．.)]|[（(][^()（）\n]{1,8}[)）]))", fl_txt)
        if m_fl_mk:
            return i_end, m_fl_mk.end(1), c0, b0

    return None


def _extract_pre_quote_introducer_style(
    en_txt: str,
    en_colors: List[str],
    en_bolds: List[bool],
    s0: int,
    s1: int,
) -> Optional[Tuple[str, bool, int]]:
    """
    Detect a 2-state pre-quote gap where plain narrative is followed immediately by a styled
    quote-introducer ending right at the opening quote (e.g. `... In the Age of Grace, <BL>the Lord Jesus said:</BL> `).
    Returns `(intro_col, intro_bold, en_intro_word_count)` if matched, else None.
    """
    non_ws = [k for k in range(s0, s1) if not en_txt[k].isspace()]
    if len(non_ws) < 2:
        return None

    # Walk transitions across non-whitespace chars in [s0, s1)
    runs: List[Tuple[str, bool, int, int]] = []
    for k in non_ws:
        ck = "blue" if en_colors[k] == "blue" else "none"
        bd = bool(en_bolds[k])
        if not runs or (runs[-1][0], runs[-1][1]) != (ck, bd):
            runs.append((ck, bd, k, k + 1))
        else:
            runs[-1] = (ck, bd, runs[-1][2], k + 1)

    if len(runs) == 2 and runs[0][0] == "none" and not runs[0][1]:
        intro_col, intro_bd, i_s, i_e = runs[1]
        en_intro_str = en_txt[i_s:i_e].strip()
        w_cnt = max(1, len(en_intro_str.split()))
        return intro_col, intro_bd, w_cnt
    return None


def _apply_fl_pre_quote_introducer(
    fl_txt: str,
    s0: int,
    s1: int,
    intro_col: str,
    intro_bd: bool,
    en_intro_words: int,
    out_colors: List[str],
    out_bolds: List[bool],
    preserve_fl_red: List[bool],
) -> None:
    """
    In Col 3 gap `[s0, s1)` right before an opening quote, keep leading narrative plain and apply
    `(intro_col, intro_bd)` ONLY to the trailing quote-introducer clause (e.g. `प्रभु यीशु ने कहा था : `).
    """
    # 1. Initialize entire gap to plain
    for k in range(s0, s1):
        out_colors[k] = "red" if preserve_fl_red[k] else "none"
        out_bolds[k] = False

    gap_sub = fl_txt[s0:s1]
    if not gap_sub.strip():
        return

    # Strip trailing colon / dash / spaces at the very end when searching for preceding clause boundary
    trimmed_end = len(gap_sub.rstrip())
    body_end = trimmed_end
    while body_end > 0 and gap_sub[body_end - 1] in ":：—–-,， ":
        body_end -= 1

    # Prefer last sentence/clause boundary before the introducer (e.g. `,` in `अनुग्रह के युग में, प्रभु यीशु ने कहा था :`)
    boundary_chars = "।.?!;।？！；。，,\n"
    last_brk = -1
    for idx in range(body_end - 1, -1, -1):
        if gap_sub[idx] in boundary_chars:
            last_brk = idx
            break

    if last_brk != -1:
        cand_rel = last_brk + 1
        while cand_rel < len(gap_sub) and gap_sub[cand_rel].isspace():
            cand_rel += 1
        # Sanity check: ensure trailing clause isn't abnormally tiny or empty
        if cand_rel < trimmed_end:
            intro_start = s0 + cand_rel
        else:
            intro_start = s0
    else:
        # Space-delimited or proportional fallback at tail of gap
        words = list(re.finditer(r"\S+", gap_sub[:trimmed_end]))
        if len(words) > en_intro_words + 1:
            take_n = min(len(words), max(2, round(en_intro_words * 1.25)))
            intro_start = s0 + words[-take_n].start()
        else:
            intro_start = s0 + (len(gap_sub) - len(gap_sub.lstrip()))

    b_end = s0 + trimmed_end
    for k in range(intro_start, s1):
        out_colors[k] = intro_col
        out_bolds[k] = bool(intro_bd and not fl_txt[k].isspace() and k < b_end)


def _col3_existing_violates_quote_bounds(
    en_txt: str,
    en_cols: List[str],
    en_bds: List[bool],
    en_start: int,
    fl_txt: str,
    fl_start: int,
    work_cols: List[str],
    work_bds: List[bool],
) -> bool:
    """Return True if Col 3's existing blue/bold bleeds outside quotes or only partially covers inner quotes."""
    en_qs = find_top_level_quotes(en_txt, start=en_start)
    fl_qs = find_top_level_quotes(fl_txt, start=fl_start)
    if not en_qs or not fl_qs:
        return False

    en_cits = find_post_quote_citation_spans(en_txt)
    fl_cits = find_post_quote_citation_spans(fl_txt)

    def _is_word_ch(ch: str) -> bool:
        return ch.isalnum() or unicodedata.category(ch).startswith("L")

    en_word_idxs = [i for i in range(en_start, len(en_txt)) if _is_word_ch(en_txt[i])]
    en_bd_words = [i for i in en_word_idxs if en_bds[i]]
    en_bl_words = [i for i in en_word_idxs if en_cols[i] == "blue"]

    en_bd_in_q = bool(en_bd_words) and all(any(qs <= i < qe for _, qs, qe, _ in en_qs) for i in en_bd_words)
    en_bl_in_qc = bool(en_bl_words) and all(
        any(qo <= i < qc for qo, _, _, qc in en_qs) or any(cs <= i < ce for cs, ce in en_cits)
        for i in en_bl_words
    )

    for k in range(fl_start, len(fl_txt)):
        if not _is_word_ch(fl_txt[k]):
            continue
        in_fq_inner = any(qs <= k < qe for _, qs, qe, _ in fl_qs)
        in_fq_outer = any(qo <= k < qc for qo, _, _, qc in fl_qs)
        in_fcit = any(cs <= k < ce for cs, ce in fl_cits)
        if en_bd_in_q and work_bds[k] and not in_fq_inner:
            return True
        if en_bl_in_qc and work_cols[k] == "blue" and not (in_fq_outer or in_fcit):
            return True

    # Check partial inner coverage inside Col 3 quotes when Col 2 quotes are fully styled
    for _, fqs, fqe, _ in fl_qs:
        q_words = [k for k in range(fqs, fqe) if _is_word_ch(fl_txt[k])]
        if not q_words:
            continue
        bl_cnt = sum(1 for k in q_words if work_cols[k] == "blue")
        bd_cnt = sum(1 for k in q_words if work_bds[k])
        if 0 < bl_cnt < len(q_words) or 0 < bd_cnt < len(q_words):
            return True
    return False


def _enforce_quote_boundary_consistency(
    en_txt: str,
    en_cols: List[str],
    en_bds: List[bool],
    en_start: int,
    fl_txt: str,
    fl_start: int,
    work_cols: List[str],
    work_bds: List[bool],
) -> None:
    """
    Post-alignment deterministic Quote-Boundary Clamping & Inner-Quote Completion:
      1. When all bold in Col 2 sits strictly inside top-level quotes, strip ANY bold outside quotes in Col 3.
      2. When all blue in Col 2 sits strictly inside top-level quotes (+ citations), strip ANY blue outside quotes/citations in Col 3.
      3. When every styled quote in Col 2 is 100% uniformly styled inside `(q_in_s, q_in_e)`, promote any partially
         matched Col 3 quote (`>= 20%` inner words styled) to 100% inner quote coverage so leading/trailing words
         like `sangat` are never missed!
    """
    en_qs = find_top_level_quotes(en_txt, start=en_start)
    fl_qs = find_top_level_quotes(fl_txt, start=fl_start)
    if not en_qs or not fl_qs:
        return

    en_cits = find_post_quote_citation_spans(en_txt)
    fl_cits = find_post_quote_citation_spans(fl_txt)

    def _is_word_ch(ch: str) -> bool:
        return ch.isalnum() or unicodedata.category(ch).startswith("L")

    en_word_idxs = [i for i in range(en_start, len(en_txt)) if _is_word_ch(en_txt[i])]
    en_bd_words = [i for i in en_word_idxs if en_bds[i]]
    en_bl_words = [i for i in en_word_idxs if en_cols[i] == "blue"]

    en_bd_strictly_in_q = bool(en_bd_words) and all(
        any(qs <= i < qe for _, qs, qe, _ in en_qs) for i in en_bd_words
    )
    en_bl_strictly_in_qc = bool(en_bl_words) and all(
        any(qo <= i < qc for qo, _, _, qc in en_qs) or any(cs <= i < ce for cs, ce in en_cits)
        for i in en_bl_words
    )

    n_fl = len(fl_txt)
    for k in range(fl_start, n_fl):
        in_fq_inner = any(qs <= k < qe for _, qs, qe, _ in fl_qs)
        in_fq_outer = any(qo <= k < qc for qo, _, _, qc in fl_qs)
        in_fcit = any(cs <= k < ce for cs, ce in fl_cits) or (
            fl_txt[k] in ".。։।" and k > fl_start and any(cs <= k - 1 < ce for cs, ce in fl_cits)
        )
        if en_bd_strictly_in_q and not in_fq_inner:
            work_bds[k] = False
        if en_bl_strictly_in_qc and work_cols[k] == "blue" and not (in_fq_outer or in_fcit):
            work_cols[k] = "none"

    # Check if every styled quote in Col 2 has uniform full-inner coverage (>= 85% of inner words)
    en_quotes_full = True
    en_any_styled_q = False
    en_open_quote_blue = False
    for qo, qs, qe, _ in en_qs:
        qw = [i for i in range(qs, qe) if _is_word_ch(en_txt[i])]
        if not qw:
            continue
        bl_n = sum(1 for i in qw if en_cols[i] == "blue")
        bd_n = sum(1 for i in qw if en_bds[i])
        if bl_n > 0 or bd_n > 0:
            en_any_styled_q = True
            if en_cols[qo] == "blue":
                en_open_quote_blue = True
            if (0 < bl_n < len(qw) * 0.85) or (0 < bd_n < len(qw) * 0.85):
                en_quotes_full = False
                break

    if en_any_styled_q and en_quotes_full:
        en_has_any_blue = bool(en_bl_words)
        en_has_any_bold = bool(en_bd_words)
        for fqo, fqs, fqe, fqc in fl_qs:
            fqw = [k for k in range(fqs, fqe) if _is_word_ch(fl_txt[k])]
            if not fqw:
                continue
            f_bl = sum(1 for k in fqw if work_cols[k] == "blue")
            f_bd = sum(1 for k in fqw if work_bds[k])
            if (f_bl / len(fqw)) >= 0.20 or (f_bd / len(fqw)) >= 0.20:
                target_bl = (f_bl > 0) or (en_has_any_blue and en_bl_strictly_in_qc and f_bd > 0)
                target_bd = (f_bd > 0) or (en_has_any_bold and en_bd_strictly_in_q and f_bl > 0)
                b_s, b_e = fqs, fqe
                while b_s < b_e and fl_txt[b_s].isspace():
                    b_s += 1
                while b_e > b_s and fl_txt[b_e - 1].isspace():
                    b_e -= 1
                for k in range(fqs, fqe):
                    if target_bl and work_cols[k] != "red":
                        work_cols[k] = "blue"
                    work_bds[k] = bool(target_bd and (b_s <= k < b_e))
                outer_c = "blue" if (target_bl and (en_open_quote_blue or work_cols[fqo] == "blue")) else work_cols[fqo]
                for k in list(range(fqo, fqs)) + list(range(fqe, fqc)):
                    if outer_c == "blue" and work_cols[k] != "red":
                        work_cols[k] = "blue"
                    work_bds[k] = False


def _try_quote_pair_blue_bold_alignment(
    en_txt: str,
    en_colors: List[str],
    en_bolds: List[bool],
    en_start: int,
    fl_txt: str,
    fl_start: int,
    out_colors: List[str],
    out_bolds: List[bool],
    preserve_fl_red: List[bool],
) -> bool:
    """When Col 2's Blue/Bold spans correspond cleanly to top-level quotation spans and `len(en_quotes) == len(fl_quotes)`,
    deterministically map every quote's outer marks, inner speech, inter-quote narrative bridge, pre-quote introducer
    (e.g. `In the Age of Grace, <BL>the Lord Jesus said:</BL>`), and post-quote citation onto Col 3 — guaranteeing 100%
    complete bold/blue inside quotes and zero bleed onto SOV post-quote clauses (e.g. Hindi `का क्या अर्थ है?`)!
    """
    en_quotes = find_top_level_quotes(en_txt, start=en_start)
    fl_quotes = find_top_level_quotes(fl_txt, start=fl_start)
    if not en_quotes or len(en_quotes) != len(fl_quotes):
        return False

    en_cits = find_post_quote_citation_spans(en_txt)
    fl_cits = find_post_quote_citation_spans(fl_txt)

    def _is_in_spans(idx: int, sp_list: List[Tuple[int, int]]) -> bool:
        return any(s0 <= idx < s1 for s0, s1 in sp_list)

    def _region_bb_uniform(s0: int, s1: int, skip_cits: bool = False) -> Optional[Tuple[str, bool]]:
        word_seen: Set[Tuple[str, bool]] = set()
        any_seen: Set[Tuple[str, bool]] = set()
        for k in range(s0, s1):
            ch = en_txt[k]
            if ch.isspace():
                continue
            if skip_cits and _is_in_spans(k, en_cits):
                continue
            if skip_cits and ch in ".。" and k > s0 and _is_in_spans(k - 1, en_cits):
                continue
            ck = "blue" if en_colors[k] == "blue" else "none"
            any_seen.add((ck, en_bolds[k]))
            if ch.isalnum() or unicodedata.category(ch).startswith("L"):
                word_seen.add((ck, en_bolds[k]))
        target_seen = word_seen if word_seen else any_seen
        if not target_seen:
            return ("none", False)
        if len(target_seen) == 1:
            return next(iter(target_seen))
        return None

    # Inspect each non-quote gap: allow either uniform `(col, bold)` OR 2-state `plain + pre-quote introducer`!
    gap_specs: List[Tuple[str, Tuple]] = []
    prev_e = en_start
    for q_open, q_in_s, q_in_e, q_close in en_quotes:
        # Advance prev_e past any post-quote citation attached to the previous quote
        for cs, ce in en_cits:
            if 0 <= cs - prev_e <= 4:
                prev_e = ce + (1 if ce < len(en_txt) and en_txt[ce] in ".。" else 0)
                break
        g_style = _region_bb_uniform(prev_e, q_open, skip_cits=True)
        if g_style is not None:
            gap_specs.append(("uniform", g_style))
        else:
            intro_spec = _extract_pre_quote_introducer_style(en_txt, en_colors, en_bolds, prev_e, q_open)
            if intro_spec is None:
                return False
            gap_specs.append(("intro", intro_spec))
        prev_e = q_close

    for cs, ce in en_cits:
        if 0 <= cs - prev_e <= 4:
            prev_e = ce + (1 if ce < len(en_txt) and en_txt[ce] in ".。" else 0)
            break
    tail_style = _region_bb_uniform(prev_e, len(en_txt), skip_cits=True)
    if tail_style is None:
        return False
    gap_specs.append(("uniform", tail_style))

    # Verify inside each quote in en_txt has a uniform inner (blue/none, bold) style
    quote_inner_styles: List[Tuple[str, bool, str]] = []
    for q_open, q_in_s, q_in_e, q_close in en_quotes:
        in_st = _region_bb_uniform(q_in_s, q_in_e, skip_cits=False)
        if in_st is None:
            return False
        q_open_bl = "blue" if (en_colors[q_open] == "blue" or in_st[0] == "blue") else "none"
        quote_inner_styles.append((in_st[0], in_st[1], q_open_bl))

    # Apply exact structural mapping onto fl_txt[fl_start:]!
    def _set_fl_range(s0: int, s1: int, col_val: str, bd_val: bool):
        if s0 >= s1:
            return
        b_start, b_end = s0, s1
        if bd_val:
            while b_start < b_end and fl_txt[b_start].isspace():
                b_start += 1
            while b_end > b_start and fl_txt[b_end - 1].isspace():
                b_end -= 1
        for k in range(s0, s1):
            if preserve_fl_red[k] and col_val == "none":
                out_colors[k] = "red"
            else:
                out_colors[k] = col_val
            out_bolds[k] = bd_val and (b_start <= k < b_end)

    fl_prev = fl_start
    for qi, (fq_open, fq_in_s, fq_in_e, fq_close) in enumerate(fl_quotes):
        g_kind, g_payload = gap_specs[qi]
        if g_kind == "uniform":
            g_col, g_bd = g_payload
            _set_fl_range(fl_prev, fq_open, g_col, g_bd)
        else:
            i_col, i_bd, i_words = g_payload
            _apply_fl_pre_quote_introducer(
                fl_txt, fl_prev, fq_open, i_col, i_bd, i_words, out_colors, out_bolds, preserve_fl_red
            )

        in_col, in_bd, outer_col = quote_inner_styles[qi]
        _set_fl_range(fq_open, fq_in_s, outer_col, False)
        _set_fl_range(fq_in_s, fq_in_e, in_col, in_bd)
        _set_fl_range(fq_in_e, fq_close, outer_col, False)

        # If a citation in Col 3 immediately follows this quote, match Col 2's post-quote citation color & unbold
        eq_close = en_quotes[qi][3]
        en_post_cit = next(((cs, ce) for cs, ce in en_cits if 0 <= cs - eq_close <= 4), None)
        fl_post_cit = next(((cs, ce) for cs, ce in fl_cits if 0 <= cs - fq_close <= 4), None)
        if fl_post_cit is not None:
            cs_f, ce_f = fl_post_cit
            if en_post_cit is not None:
                cs_e, ce_e = en_post_cit
                cit_col = "blue" if any(en_colors[x] == "blue" for x in range(cs_e, ce_e) if not en_txt[x].isspace()) else "none"
            else:
                cit_col = outer_col
            _set_fl_range(fq_close, ce_f, cit_col, False)
            if ce_f < len(fl_txt) and fl_txt[ce_f] in ".。։।":
                _set_fl_range(ce_f, ce_f + 1, cit_col, False)
                fl_prev = ce_f + 1
            else:
                fl_prev = ce_f
        else:
            fl_prev = fq_close

    _, (t_col, t_bd) = gap_specs[-1]
    _set_fl_range(fl_prev, len(fl_txt), t_col, t_bd)
    return True


def format_annotated_source_v2(
    segments: List[Tuple[str, str, bool]],
    include_blue_bold: bool = True,
    include_red: bool = False,
) -> str:
    parts = []
    for txt, ck, bd in segments:
        if ck == "blue" and include_blue_bold:
            parts.append(f"<BB>{txt}</BB>" if bd else f"<BL>{txt}</BL>")
        elif ck == "red" and include_red:
            parts.append(f"<RB>{txt}</RB>" if bd else f"<RD>{txt}</RD>")
        elif ck == "none" and bd and include_blue_bold:
            parts.append(f"<BD>{txt}</BD>")
        else:
            parts.append(txt)
    return "".join(parts)


def apply_whole_paragraph_style(p_fl: ET.Element, color_kind: str, bold: bool, preserve_existing_red: bool = False):
    for r in p_fl.findall(f"{W}r"):
        txt = "".join(t.text or "" for t in r.findall(f"{W}t"))
        if not txt:
            continue
        if preserve_existing_red and color_kind == "none" and is_run_red(r):
            pass
        else:
            set_run_color_kind(r, color_kind)
        set_run_bold(r, bold if txt.strip() else False)


def _trim_outer_quotes_from_bold_span(full_text: str, s0: int, s1: int) -> Tuple[int, int]:
    """Ensure outer quotation marks and surrounding spaces are NOT bolded, while keeping inner punctuation bold."""
    while s0 < s1 and (full_text[s0].isspace() or full_text[s0] in ALL_OUTER_QUOTE_CHARS):
        s0 += 1
    while s1 > s0 and (full_text[s1 - 1].isspace() or full_text[s1 - 1] in ALL_OUTER_QUOTE_CHARS):
        s1 -= 1
    return s0, s1


def _locate_substring_fuzzy(
    haystack: str,
    needle: str,
    start_cursor: int = 0,
    search_from: Optional[int] = None,
) -> Optional[Tuple[int, int]]:
    if not needle:
        return None
    cur = search_from if search_from is not None else start_cursor
    idx = haystack.find(needle, cur)
    if idx != -1:
        return idx, idx + len(needle)
    idx_any = haystack.find(needle)
    if idx_any != -1 and search_from is None:
        return idx_any, idx_any + len(needle)

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
        idx2 = h_norm.find(n_norm, cur)
        if idx2 != -1:
            return idx2, idx2 + len(n_norm)
        if search_from is None:
            idx3 = h_norm.find(n_norm)
            if idx3 != -1:
                return idx3, idx3 + len(n_norm)
    return None


def _write_char_styles_to_paragraph(
    p_fl: ET.Element,
    char_colors: List[str],
    char_bolds: List[bool],
) -> None:
    full_txt = _get_p_text(p_fl)
    n = len(full_txt)
    if n == 0 or len(char_colors) != n or len(char_bolds) != n:
        return

    cuts: Set[int] = set()
    for i in range(1, n):
        if char_colors[i] != char_colors[i - 1] or char_bolds[i] != char_bolds[i - 1]:
            cuts.add(i)

    sliced = split_paragraph_runs_at_offsets(p_fl, cuts)
    for r_s, _r_e, r_el in sliced:
        set_run_color_kind(r_el, char_colors[r_s])
        set_run_bold(r_el, char_bolds[r_s])


def _apply_span_dicts_onto_arrays(
    fl_txt: str,
    fl_start: int,
    span_items: List[Dict],
    char_colors: List[str],
    char_bolds: List[bool],
    en_allows_bold: bool,
) -> int:
    """Apply [{'text': ..., 'blue': bool, 'red': bool, 'bold': bool}, ...] onto `char_colors`/`char_bolds`."""
    n = len(fl_txt)
    cursor = fl_start
    matched_cnt = 0

    for item in span_items:
        sub = str(item.get("text", ""))
        if not sub.strip():
            continue
        bl = bool(item.get("blue", False))
        rd = bool(item.get("red", False))
        bd = bool(item.get("bold", False)) and en_allows_bold
        if not bl and not rd and not bd:
            continue

        loc = _locate_substring_fuzzy(fl_txt, sub, start_cursor=cursor)
        if loc is None:
            loc = _locate_substring_fuzzy(fl_txt, sub.strip(), start_cursor=cursor)
        if loc is None or loc[0] < fl_start:
            continue

        s0, s1 = loc
        cursor = s1
        matched_cnt += 1

        if rd:
            for idx in range(s0, s1):
                char_colors[idx] = "red"
        elif bl:
            for idx in range(s0, s1):
                char_colors[idx] = "blue"

        if bd:
            b0, b1 = _trim_outer_quotes_from_bold_span(fl_txt, s0, s1)
            for idx in range(s0, s1):
                char_bolds[idx] = b0 <= idx < b1
        else:
            for idx in range(s0, s1):
                char_bolds[idx] = False

    if matched_cnt == 0:
        return 0

    # Auto-extend blue to immediately adjacent outer quotation marks if omitted
    for i in range(max(1, fl_start + 1), n):
        if char_colors[i] == "blue" and char_colors[i - 1] == "none" and fl_txt[i - 1] in ALL_OUTER_QUOTE_CHARS:
            char_colors[i - 1] = "blue"
    for i in range(n - 2, fl_start - 1, -1):
        if char_colors[i] == "blue" and char_colors[i + 1] == "none" and fl_txt[i + 1] in ALL_OUTER_QUOTE_CHARS:
            char_colors[i + 1] = "blue"

    # Outer quotes & post-quote citations are never bold
    for i in range(fl_start, n):
        if fl_txt[i] in ALL_OUTER_QUOTE_CHARS:
            char_bolds[i] = False
    for c0, c1 in find_post_quote_citation_spans(fl_txt):
        if c0 >= fl_start:
            for idx in range(c0, c1):
                char_bolds[idx] = False

    return matched_cnt


def _align_spans_locally_onto_arrays(
    en_txt: str,
    en_colors: List[str],
    en_bolds: List[bool],
    en_start: int,
    fl_txt: str,
    fl_start: int,
    char_colors: List[str],
    char_bolds: List[bool],
    align_blue_bold: bool,
    align_red: bool,
) -> int:
    """Offline proportional clause/word fallback alignment using `local_align_col3_spans`."""
    en_sub_full = en_txt[en_start:]
    fl_sub_full = fl_txt[fl_start:]
    if not en_sub_full.strip() or not fl_sub_full.strip():
        return 0

    raw_segs = _compress_style_segments(en_txt, en_colors, en_bolds, start=en_start)
    spans_req: List[Dict] = []
    meta_by_id: Dict[str, Tuple[str, bool, str]] = {}
    rel_pos = 0
    for seg_txt, ck, bd in raw_segs:
        seg_len = len(seg_txt)
        s0 = rel_pos
        s1 = rel_pos + seg_len
        rel_pos = s1
        if not seg_txt.strip():
            continue
        want = False
        if ck == "red" and align_red:
            want = True
        elif (ck == "blue" or (ck == "none" and bd)) and align_blue_bold:
            want = True
        if not want:
            continue
        sid = f"sp_{len(spans_req)}"
        spans_req.append(
            {
                "span_id": sid,
                "type": "up",
                "old_en_in_col3": seg_txt,
                "old_en_char_start": s0,
                "old_en_char_end": s1,
                "left_old_context": en_sub_full[max(0, s0 - 30) : s0],
                "right_old_context": en_sub_full[s1 : min(len(en_sub_full), s1 + 30)],
            }
        )
        meta_by_id[sid] = (ck, bd, seg_txt)

    if not spans_req:
        return 0

    aligned_map = local_align_col3_spans(fl_sub_full, en_sub_full, spans_req, _locate_substring_fuzzy)
    if not aligned_map:
        return 0

    # Track repeated identical English phrases in the same paragraph (e.g. 'winning over people’s hearts' x2)
    phrase_memory: Dict[str, str] = {}
    span_items: List[Dict] = []
    fl_search_cur = 0

    for req in spans_req:
        sid = req["span_id"]
        ck, bd, en_piece = meta_by_id[sid]
        norm_en_key = re.sub(r"[\s.,?!;:“”\"'’()]+", " ", en_piece.strip().lower()).strip()
        sub_fl = aligned_map.get(sid, "")

        # If this exact English multi-word phrase already aligned earlier in the paragraph and the same foreign phrase
        # appears verbatim near the proportional region, prefer that exact repeated foreign phrase!
        if norm_en_key and " " in norm_en_key and norm_en_key in phrase_memory:
            prev_fl_phrase = phrase_memory[norm_en_key]
            repeat_pos = fl_sub_full.find(prev_fl_phrase, fl_search_cur)
            if repeat_pos != -1:
                sub_fl = prev_fl_phrase
                if en_piece.rstrip().endswith(".") and repeat_pos + len(prev_fl_phrase) < len(fl_sub_full) and fl_sub_full[repeat_pos + len(prev_fl_phrase)] == ".":
                    sub_fl = prev_fl_phrase + "."

        if not sub_fl:
            continue

        # Trim trailing sentence punctuation on local match if Col 2 span did NOT end with punctuation
        en_r = en_piece.rstrip()
        if en_r and en_r[-1] not in ".?!;:。？！；：)" and sub_fl[-1] in ".?!;:。？！；：)":
            trimmed_fl = sub_fl.rstrip(".?!;:。？！；：)")
            if trimmed_fl.strip():
                sub_fl = trimmed_fl

        if norm_en_key and " " in norm_en_key and norm_en_key not in phrase_memory:
            clean_mem = sub_fl.rstrip(".?!;:。？！；：")
            if len(clean_mem) >= 4:
                phrase_memory[norm_en_key] = clean_mem

        loc_adv = _locate_substring_fuzzy(fl_sub_full, sub_fl, start_cursor=fl_search_cur)
        if loc_adv is not None:
            fl_search_cur = loc_adv[1]

        span_items.append(
            {
                "text": sub_fl,
                "blue": ck == "blue",
                "red": ck == "red",
                "bold": bd,
            }
        )

    en_allows_bold = any(en_bolds[k] for k in range(en_start, len(en_txt)))
    return _apply_span_dicts_onto_arrays(
        fl_txt,
        fl_start,
        span_items,
        char_colors,
        char_bolds,
        en_allows_bold=en_allows_bold,
    )


_FELLOWSHIP_ALIGN_SYSTEM_PROMPT = """You are a precision multilingual typesetting alignment engine for trilingual manuscripts (Column 1: Chinese, Column 2: English with formatting tags, Column 3: Target Foreign Language).

In Column 2 (English), styled spans are marked with:
- <BB>...</BB> : Blue font (#0000FF) AND Bold (`"blue": true, "red": false, "bold": true`)
- <BL>...</BL> : Blue font ONLY, NOT Bold (`"blue": true, "red": false, "bold": false`)
- <BD>...</BD> : Bold ONLY, NOT Blue/Red (`"blue": false, "red": false, "bold": true`)
- <RB>...</RB> : Red font (#FF0000) AND Bold (`"blue": false, "red": true, "bold": true`)
- <RD>...</RD> : Red font ONLY, NOT Bold (`"blue": false, "red": true, "bold": false`)

Your task:
For each item, locate the EXACT corresponding verbatim substrings inside `col3_text` in their exact order of appearance, and return JSON:
{
  "results": [
    {
      "id": <item_id>,
      "spans": [
        {"text": "<exact verbatim substring copied from col3_text>", "blue": true, "red": false, "bold": false},
        {"text": "<exact verbatim substring copied from col3_text>", "blue": true, "red": false, "bold": true}
      ]
    }
  ]
}

STRICT RULES:
1. Every `text` MUST be copied character-for-character from `col3_text`. Do NOT translate or alter spelling/punctuation.
2. NEVER set `"bold": true` on `<BL>` or `<RD>` spans! Strictly mirror English's exact `(blue, red, bold)` combination!
3. Narrative text BETWEEN two quotes (e.g. 'Isaiah answered:' / 'Ο Ησαΐας απάντησε:') that is unstyled in English MUST be omitted from styled spans so it stays plain (neither blue nor bold)!
4. Outer quotation marks (“ ” « » „ “ ‘ ’) and citation parentheses like `(Matio 7:23)` MUST have `"bold": false`.
5. Punctuation INSIDE a bolded quote sentence (commas, semicolons, exclamation/question marks inside `<BB>...</BB>`) stays inside the `"bold": true` span.
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
                                "red": {"type": "boolean"},
                                "bold": {"type": "boolean"},
                            },
                            "required": ["text", "blue", "red", "bold"],
                        },
                    },
                },
                "required": ["id", "spans"],
            },
        }
    },
    "required": ["results"],
}


def _count_distinct_true_runs(flags: List[bool], text: str, start: int = 0) -> int:
    cnt = 0
    in_run = False
    for i in range(start, len(flags)):
        if text[i].isspace():
            continue
        if flags[i]:
            if not in_run:
                cnt += 1
                in_run = True
        else:
            in_run = False
    return cnt


def sync_table_col3_blue_and_bold(
    tbl: ET.Element,
    api_pool: Optional[ApiKeyPool],
    model: str,
    api_mode: str,
    enable_ai_col3: bool = True,
    sync_red: bool = False,
    log_fn: Optional[Callable[[str], None]] = None,
) -> Dict:
    """Synchronize Col 3 blue (#0000FF), optional red (#FF0000 when `sync_red=True`), and bold (<w:b/>)
    strictly mirroring Col 2 across all rows of `tbl`.
    """
    rows = tbl.findall(f"{W}tr")
    tier1_synced = 0
    skipped_existing = 0
    local_fallback_synced = 0
    ai_tasks: List[Dict] = []

    for r_idx, tr in enumerate(rows):
        tcs = tr.findall(f"{W}tc")
        if len(tcs) < 3:
            continue
        ps_zh = tcs[0].findall(f"{W}p")
        ps_en = tcs[1].findall(f"{W}p")
        ps_fl = tcs[2].findall(f"{W}p")

        next_tr = rows[r_idx + 1] if r_idx + 1 < len(rows) else None
        if r_idx == 0 and _is_main_title_table_row(tr, next_tr):
            en0_txt, en0_cols, _ = _extract_char_styles(ps_en[0], include_red=sync_red) if ps_en else ("", [], [])
            title_blue = (not en0_txt.strip()) or any(c == "blue" for c in en0_cols)
            title_red = sync_red and any(c == "red" for c in en0_cols) and not title_blue
            ck_title = "blue" if title_blue else ("red" if title_red else "none")
            for p_fl in ps_fl:
                apply_whole_paragraph_style(p_fl, color_kind=ck_title, bold=True)
            tier1_synced += 1
            continue

        pair_count = min(len(ps_en), len(ps_fl))
        for p_i in range(pair_count):
            p_en = ps_en[p_i]
            p_fl = ps_fl[p_i]
            p_zh = ps_zh[p_i] if p_i < len(ps_zh) else None

            en_txt, en_cols, en_bds = _extract_char_styles(p_en, include_red=sync_red)
            fl_txt, fl_orig_cols, fl_orig_bds = _extract_char_styles(p_fl, include_red=True)
            if not fl_txt.strip() or not en_txt.strip():
                continue

            en_non_ws = [i for i, ch in enumerate(en_txt) if not ch.isspace()]
            if not en_non_ws:
                continue

            # 1. Check whole-paragraph uniform style in Col 2
            distinct_en_styles = {(en_cols[i], en_bds[i]) for i in en_non_ws}
            if len(distinct_en_styles) == 1:
                u_col, u_bd = next(iter(distinct_en_styles))
                if (u_col, u_bd) == ("none", False):
                    # Plain Col 2: strip any stray blue/bold on Col 3 (preserving existing Col 3 red if sync_red=False)
                    zh_txt_chk, zh_cols_chk, zh_bds_chk = _extract_char_styles(p_zh, include_red=sync_red) if p_zh is not None else ("", [], [])
                    zh_plain = not any(c != "none" or b for ch, c, b in zip(zh_txt_chk, zh_cols_chk, zh_bds_chk) if not ch.isspace())
                    if zh_plain and (any(c == "blue" for c in fl_orig_cols) or any(fl_orig_bds)):
                        apply_whole_paragraph_style(p_fl, color_kind="none", bold=False, preserve_existing_red=not sync_red)
                    continue
                else:
                    apply_whole_paragraph_style(p_fl, color_kind=u_col, bold=u_bd)
                    # Unbold outer quotes & citations if whole paragraph is a quoted verse
                    fl_cits = find_post_quote_citation_spans(fl_txt)
                    if fl_cits and u_bd:
                        _, cur_c, cur_b = _extract_char_styles(p_fl, include_red=True)
                        for idx_c, ch_c in enumerate(fl_txt):
                            if ch_c in ALL_OUTER_QUOTE_CHARS:
                                cur_b[idx_c] = False
                        for cs, ce in fl_cits:
                            for idx_c in range(cs, ce):
                                cur_b[idx_c] = False
                        _write_char_styles_to_paragraph(p_fl, cur_c, cur_b)
                    tier1_synced += 1
                    continue

            # Prepare working style arrays for Col 3
            n_fl = len(fl_txt)
            preserve_fl_red = [c == "red" for c in fl_orig_cols]
            work_cols = list(fl_orig_cols)
            work_bds = list(fl_orig_bds)

            # 2. Deterministic Leading Prefix Synchronizer ('Question 10:' -> 'Swali la 10:', 'Answer:', '1.', etc.)
            en_rem_start = 0
            fl_rem_start = 0
            pfx_hit = _detect_leading_styled_prefix(en_txt, en_cols, en_bds, fl_txt)
            if pfx_hit is not None:
                en_rem_start, fl_rem_start, pfx_col, pfx_bd = pfx_hit
                fl_lead_ws = 0
                while fl_lead_ws < fl_rem_start and fl_txt[fl_lead_ws].isspace():
                    fl_lead_ws += 1
                for k in range(0, fl_rem_start):
                    if k < fl_lead_ws:
                        work_cols[k] = "none"
                        work_bds[k] = False
                    else:
                        work_cols[k] = pfx_col
                        work_bds[k] = pfx_bd

            # Check what styles exist in Col 2's remainder [en_rem_start:]
            en_rem_non_ws = [i for i in range(en_rem_start, len(en_txt)) if not en_txt[i].isspace()]
            en_rem_has_blue = any(en_cols[i] == "blue" for i in en_rem_non_ws)
            en_rem_has_bold = any(en_bds[i] for i in en_rem_non_ws)
            en_rem_has_red = sync_red and any(en_cols[i] == "red" for i in en_rem_non_ws)

            # If Col 2 remainder is completely plain, clear blue/bold on Col 3 remainder and finish!
            if not en_rem_has_blue and not en_rem_has_bold and not en_rem_has_red:
                for k in range(fl_rem_start, n_fl):
                    work_cols[k] = "red" if (not sync_red and preserve_fl_red[k]) else "none"
                    work_bds[k] = False
                _write_char_styles_to_paragraph(p_fl, work_cols, work_bds)
                tier1_synced += 1
                continue

            # 3. Strict Decoupling Guards on Col 3 remainder:
            #    - If Col 2 remainder has NO bold -> strip ALL bold from Col 3 remainder!
            #    - If Col 2 remainder has NO blue -> strip ALL blue from Col 3 remainder!
            if not en_rem_has_bold:
                for k in range(fl_rem_start, n_fl):
                    work_bds[k] = False
            if not en_rem_has_blue:
                for k in range(fl_rem_start, n_fl):
                    if work_cols[k] == "blue":
                        work_cols[k] = "none"

            # If all bold in Col 2 remainder sits strictly inside top-level quotes, strip any bold OUTSIDE quotes in Col 3!
            if en_rem_has_bold:
                en_qs_chk = find_top_level_quotes(en_txt, start=en_rem_start)
                fl_qs_chk = find_top_level_quotes(fl_txt, start=fl_rem_start)
                if en_qs_chk and all(
                    any(qi_s <= idx < qi_e for _, qi_s, qi_e, _ in en_qs_chk)
                    for idx in en_rem_non_ws
                    if en_bds[idx]
                ):
                    for k in range(fl_rem_start, n_fl):
                        if not any(fqi_s <= k < fqi_e for _, fqi_s, fqi_e, _ in fl_qs_chk):
                            work_bds[k] = False

            # Check Red channel status
            fl_rem_has_red = any(fl_orig_cols[k] == "red" for k in range(fl_rem_start, n_fl) if not fl_txt[k].isspace())
            need_align_red = en_rem_has_red and not fl_rem_has_red

            # 4. Try exact deterministic Quote-Pair Blue/Bold mapping
            need_align_bb = en_rem_has_blue or en_rem_has_bold
            if need_align_bb:
                if _try_quote_pair_blue_bold_alignment(
                    en_txt,
                    en_cols,
                    en_bds,
                    en_rem_start,
                    fl_txt,
                    fl_rem_start,
                    work_cols,
                    work_bds,
                    preserve_fl_red,
                ):
                    need_align_bb = False
                    tier1_synced += 1

            # 5. If Quote-Pair didn't resolve Blue/Bold, check if Col 3's existing Blue/Bold already matches Col 2
            if need_align_bb:
                fl_rem_has_blue = any(work_cols[k] == "blue" for k in range(fl_rem_start, n_fl) if not fl_txt[k].isspace())
                fl_rem_has_bold = any(work_bds[k] for k in range(fl_rem_start, n_fl) if not fl_txt[k].isspace())
                en_bd_runs = _count_distinct_true_runs(en_bds, en_txt, start=en_rem_start)
                fl_bd_runs = _count_distinct_true_runs(work_bds, fl_txt, start=fl_rem_start)

                if (
                    (not en_rem_has_blue or fl_rem_has_blue)
                    and (not en_rem_has_bold or (fl_rem_has_bold and fl_bd_runs == en_bd_runs))
                    and not _col3_existing_violates_quote_bounds(
                        en_txt, en_cols, en_bds, en_rem_start, fl_txt, fl_rem_start, work_cols, work_bds
                    )
                ):
                    # Normalize outer quotes & post-quote citations locally
                    for k in range(fl_rem_start, n_fl):
                        if fl_txt[k] in ALL_OUTER_QUOTE_CHARS:
                            work_bds[k] = False
                    for c0, c1 in find_post_quote_citation_spans(fl_txt):
                        if c0 >= fl_rem_start:
                            for k in range(c0, c1):
                                work_bds[k] = False
                    need_align_bb = False
                    skipped_existing += 1

            if not need_align_bb and not need_align_red:
                _write_char_styles_to_paragraph(p_fl, work_cols, work_bds)
                continue

            # Clear channels that are about to be freshly aligned on [fl_rem_start:] so stale artifacts don't linger
            for k in range(fl_rem_start, n_fl):
                if need_align_bb:
                    if work_cols[k] == "blue":
                        work_cols[k] = "none"
                    if not (fl_rem_has_red and work_cols[k] == "red"):
                        work_bds[k] = False
                if need_align_red and work_cols[k] == "red":
                    work_cols[k] = "none"

            ai_tasks.append(
                {
                    "row_idx": r_idx,
                    "p_fl": p_fl,
                    "en_txt": en_txt,
                    "en_cols": en_cols,
                    "en_bds": en_bds,
                    "en_rem_start": en_rem_start,
                    "en_rem_has_bold": en_rem_has_bold,
                    "fl_txt": fl_txt,
                    "fl_rem_start": fl_rem_start,
                    "work_cols": work_cols,
                    "work_bds": work_bds,
                    "need_align_bb": need_align_bb,
                    "need_align_red": need_align_red,
                    "annotated_en": format_annotated_source_v2(
                        _compress_style_segments(en_txt, en_cols, en_bds, start=en_rem_start),
                        include_blue_bold=need_align_bb,
                        include_red=need_align_red,
                    ),
                    "zh_text": _get_p_text(p_zh) if p_zh is not None else "",
                }
            )

    ai_synced = 0
    failed_rows: List[int] = []
    if ai_tasks:
        row_nums = sorted({t["row_idx"] + 1 for t in ai_tasks})
        can_use_ai = bool(enable_ai_col3 and api_pool is not None and len(api_pool) > 0)

        if not can_use_ai:
            for task in ai_tasks:
                applied_local = _align_spans_locally_onto_arrays(
                    task["en_txt"],
                    task["en_cols"],
                    task["en_bds"],
                    task["en_rem_start"],
                    task["fl_txt"],
                    task["fl_rem_start"],
                    task["work_cols"],
                    task["work_bds"],
                    align_blue_bold=task["need_align_bb"],
                    align_red=task["need_align_red"],
                )
                if applied_local > 0:
                    local_fallback_synced += 1
                _enforce_quote_boundary_consistency(
                    task["en_txt"],
                    task["en_cols"],
                    task["en_bds"],
                    task["en_rem_start"],
                    task["fl_txt"],
                    task["fl_rem_start"],
                    task["work_cols"],
                    task["work_bds"],
                )
                _write_char_styles_to_paragraph(task["p_fl"], task["work_cols"], task["work_bds"])
            if log_fn:
                mode_note = "纯本地模式" if not enable_ai_col3 else "未填写 API Key"
                log_fn(
                    f"   ℹ️ {mode_note}：第 {row_nums} 行段内引文/红蓝字已通过本地分句与引号对齐完成 {local_fallback_synced} 段同步。"
                )
        else:
            BATCH_SIZE = 3
            ai_dead = False
            for b_start in range(0, len(ai_tasks), BATCH_SIZE):
                batch = ai_tasks[b_start : b_start + BATCH_SIZE]
                if ai_dead:
                    for task in batch:
                        applied_local = _align_spans_locally_onto_arrays(
                            task["en_txt"],
                            task["en_cols"],
                            task["en_bds"],
                            task["en_rem_start"],
                            task["fl_txt"],
                            task["fl_rem_start"],
                            task["work_cols"],
                            task["work_bds"],
                            align_blue_bold=task["need_align_bb"],
                            align_red=task["need_align_red"],
                        )
                        if applied_local > 0:
                            local_fallback_synced += 1
                        _enforce_quote_boundary_consistency(
                            task["en_txt"],
                            task["en_cols"],
                            task["en_bds"],
                            task["en_rem_start"],
                            task["fl_txt"],
                            task["fl_rem_start"],
                            task["work_cols"],
                            task["work_bds"],
                        )
                        _write_char_styles_to_paragraph(task["p_fl"], task["work_cols"], task["work_bds"])
                    continue

                payload_items = []
                for local_id, task in enumerate(batch):
                    payload_items.append(
                        {
                            "id": local_id,
                            "col1_chinese_ref": task["zh_text"][:1500],
                            "col2_annotated_english": task["annotated_en"],
                            "col3_text": task["fl_txt"][task["fl_rem_start"] :],
                        }
                    )
                user_prompt = (
                    "Align the <BB>, <BL>, <BD>, <RB>, and <RD> spans from `col2_annotated_english` onto `col3_text`:\n"
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
                        applied_n = _apply_span_dicts_onto_arrays(
                            task["fl_txt"],
                            task["fl_rem_start"],
                            spans,
                            task["work_cols"],
                            task["work_bds"],
                            en_allows_bold=task["en_rem_has_bold"],
                        )
                        if applied_n == 0:
                            applied_n = _align_spans_locally_onto_arrays(
                                task["en_txt"],
                                task["en_cols"],
                                task["en_bds"],
                                task["en_rem_start"],
                                task["fl_txt"],
                                task["fl_rem_start"],
                                task["work_cols"],
                                task["work_bds"],
                                align_blue_bold=task["need_align_bb"],
                                align_red=task["need_align_red"],
                            )
                        _enforce_quote_boundary_consistency(
                            task["en_txt"],
                            task["en_cols"],
                            task["en_bds"],
                            task["en_rem_start"],
                            task["fl_txt"],
                            task["fl_rem_start"],
                            task["work_cols"],
                            task["work_bds"],
                        )
                        _write_char_styles_to_paragraph(task["p_fl"], task["work_cols"], task["work_bds"])
                        if applied_n > 0:
                            ai_synced += 1
                            if log_fn:
                                log_fn(f"   ✓ 第 {task['row_idx'] + 1} 行第 3 列已精准还原 {applied_n} 处红蓝/加粗标记")
                except Exception as ai_ex:
                    ai_dead = True
                    friendly = format_friendly_api_error(ai_ex, total_keys=len(api_pool))
                    if log_fn:
                        log_fn(f"   ⚠️ {friendly} -> 已自动切换「本地位置估算模式」继续完成第 3 列标记")
                    for task in batch:
                        applied_local = _align_spans_locally_onto_arrays(
                            task["en_txt"],
                            task["en_cols"],
                            task["en_bds"],
                            task["en_rem_start"],
                            task["fl_txt"],
                            task["fl_rem_start"],
                            task["work_cols"],
                            task["work_bds"],
                            align_blue_bold=task["need_align_bb"],
                            align_red=task["need_align_red"],
                        )
                        if applied_local > 0:
                            local_fallback_synced += 1
                        _enforce_quote_boundary_consistency(
                            task["en_txt"],
                            task["en_cols"],
                            task["en_bds"],
                            task["en_rem_start"],
                            task["fl_txt"],
                            task["fl_rem_start"],
                            task["work_cols"],
                            task["work_bds"],
                        )
                        _write_char_styles_to_paragraph(task["p_fl"], task["work_cols"], task["work_bds"])

    return {
        "tier1_synced": tier1_synced,
        "skipped_existing": skipped_existing,
        "ai_paragraphs_total": len(ai_tasks),
        "ai_synced": ai_synced,
        "local_fallback_synced": local_fallback_synced,
        "failed_rows": sorted(set(failed_rows)),
    }


# ---------------- Step 7: Spacing, Standard Fonts, Thai Sizes & 3-Column Width Balancing ----------------

def _is_main_title_table_row(tr: ET.Element, next_tr: Optional[ET.Element] = None) -> bool:
    """Detect if Row 0 of a table is the Main Title Row (保留四号 14pt / 泰语 18pt)."""
    tcs = tr.findall(f"{W}tc")
    if len(tcs) < 2:
        return False
    c1_txt = _get_tc_text(tcs[0]).strip()
    c2_txt = _get_tc_text(tcs[1]).strip()
    if not c1_txt or len(c1_txt) > 80 or len(c2_txt) > 180:
        return False
    if "上面的交通" in c1_txt or "Fellowship From" in c2_txt:
        return False
    # Exclude numbered subheadings / questions
    if re.match(r"^(?:[0-9IVX]+[．.、]|第[一二三四五六七八九十0-9]+条|问题\s*\d+|Question\s+\d+)", c1_txt, re.I):
        return False
    p0 = tcs[0].find(f"{W}p")
    if p0 is not None:
        jc = p0.find(f"{W}pPr/{W}jc")
        if jc is not None and jc.attrib.get(f"{W}val") == "center":
            return True
    for r in tcs[0].findall(f".//{W}r"):
        sz = r.find(f"{W}rPr/{W}sz")
        if sz is not None and sz.attrib.get(f"{W}val") in ("28", "36"):
            return True
    # Also recognize Row 0 main title when all non-ws runs in Col 1 & Col 2 are bold and followed by blank Row 1
    c1_runs = [r for r in tcs[0].findall(f".//{W}r") if "".join(t.text or "" for t in r.findall(f"{W}t")).strip()]
    c2_runs = [r for r in tcs[1].findall(f".//{W}r") if "".join(t.text or "" for t in r.findall(f"{W}t")).strip()]
    if c1_runs and all(is_run_bold(r) for r in c1_runs) and (not c2_runs or all(is_run_bold(r) for r in c2_runs)):
        if next_tr is None or _is_all_empty_row(next_tr):
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


def apply_font_sizes_to_paragraph(
    p: ET.Element,
    is_main_title_row: bool = False,
    col_idx: int = 2,
    font_name: str = "",
    title_sz: int = 28,
    body_sz: int = 24,
    citation_sz: int = 21,
    citation_hints: Optional[Dict] = None,
) -> int:
    """Apply standard font family + title/body/citation font sizes onto `p` (and its paragraph mark `<w:pPr>/<w:rPr>`)."""
    title_sz_s = str(title_sz)
    body_sz_s = str(body_sz)
    cit_sz_s = str(citation_sz)

    # Sync paragraph mark <w:pPr>/<w:rPr> font & size so Word's trailing pilcrow matches
    ppr = p.find(f"{W}pPr")
    if ppr is not None:
        p_rpr = ppr.find(f"{W}rPr")
        if p_rpr is None and font_name:
            p_rpr = ET.SubElement(ppr, f"{W}rPr")
        if p_rpr is not None:
            if font_name:
                set_rpr_font_family(p_rpr, font_name, col_idx)
            set_rpr_font_size(p_rpr, title_sz_s if is_main_title_row else body_sz_s)

    if is_main_title_row:
        for r in p.findall(f"{W}r"):
            if font_name:
                set_run_font_family(r, font_name, col_idx)
            set_run_font_size(r, title_sz_s)
        return 0

    full_txt = _get_p_text(p)
    if not full_txt:
        for r in p.findall(f"{W}r"):
            if font_name:
                set_run_font_family(r, font_name, col_idx)
            set_run_font_size(r, body_sz_s)
        return 0

    if col_idx == 2:
        cit_items = find_col3_citation_spans_with_hints(full_txt, hints=citation_hints)
    else:
        base_cits = find_post_quote_citation_spans(full_txt)
        if citation_hints and citation_hints.get("whole_para_citation"):
            cit_items = [(0, len(full_txt), None)]
        else:
            cit_items = [(s0, s1, None) for s0, s1 in base_cits]

    if not cit_items:
        for r in p.findall(f"{W}r"):
            if font_name:
                set_run_font_family(r, font_name, col_idx)
            set_run_font_size(r, body_sz_s)
        return 0

    cuts: Set[int] = set()
    for s0, s1, _ in cit_items:
        cuts.add(s0)
        cuts.add(s1)

    sliced_runs = split_paragraph_runs_at_offsets(p, cuts)
    for r_s, r_e, r_el in sliced_runs:
        if font_name:
            set_run_font_family(r_el, font_name, col_idx)
        matched_cit = next(((c0, c1, bl_ov) for c0, c1, bl_ov in cit_items if r_s >= c0 and r_e <= c1), None)
        if matched_cit is not None:
            set_run_font_size(r_el, cit_sz_s)
            if col_idx == 2:
                # Citations in Col 3 are NEVER bold, and match Col 2's blue/non-blue state
                set_run_bold(r_el, False)
                _, _, bl_ov = matched_cit
                if bl_ov is True:
                    set_run_color_kind(r_el, "blue")
                elif bl_ov is False:
                    set_run_color_kind(r_el, "none")
        else:
            set_run_font_size(r_el, body_sz_s)

    # Also ensure non-text runs (if any) get font family
    if font_name:
        for r in p.findall(f"{W}r"):
            set_run_font_family(r, font_name, col_idx)

    return len(cit_items)


def capture_table_citation_hints(tbl: ET.Element) -> Dict[Tuple[int, int], Dict]:
    """Pre-capture Col 1 & Col 2 citation hints (including original w:sz='21') per `(row_idx, p_idx)`."""
    hints_map: Dict[Tuple[int, int], Dict] = {}
    for r_idx, tr in enumerate(tbl.findall(f"{W}tr")):
        tcs = tr.findall(f"{W}tc")
        if len(tcs) < 3:
            continue
        ps_zh = tcs[0].findall(f"{W}p")
        ps_en = tcs[1].findall(f"{W}p")
        ps_fl = tcs[2].findall(f"{W}p")
        for p_i in range(len(ps_fl)):
            p_zh = ps_zh[p_i] if p_i < len(ps_zh) else None
            p_en = ps_en[p_i] if p_i < len(ps_en) else None
            hints_map[(r_idx, p_i)] = extract_paragraph_source_citation_hints(p_zh, p_en)
    return hints_map


def enforce_table_spacing_and_typography(
    tbl: ET.Element,
    apply_curly_quotes_col2: bool = False,
    apply_curly_quotes_col3: bool = False,
    lang_code: str = "AUTO",
    lang_typo: Optional[Dict] = None,
    citation_hints_map: Optional[Dict[Tuple[int, int], Dict]] = None,
) -> int:
    """Apply standard fonts (Col 1 宋体, Col 2 Georgia, Col 3 get_lang_typography), Thai/standard sizes,
    cross-column & multi-script citation sizing, and exact 20pt line spacing.
    """
    if apply_curly_quotes_col2 or apply_curly_quotes_col3:
        apply_curly_quotes_to_table(tbl, apply_curly_quotes_col2, apply_curly_quotes_col3)

    if citation_hints_map is None:
        citation_hints_map = capture_table_citation_hints(tbl)

    typo = lang_typo if lang_typo is not None else get_lang_typography(lang_code)
    col3_font = str(typo.get("font", "")).strip()
    col3_title_sz = int(typo.get("title_sz", 28))
    col3_body_sz = int(typo.get("body_sz", 24))
    col3_cit_sz = int(typo.get("citation_sz", 21))
    col3_line_dxa = int(typo.get("line_dxa", 400))

    rows = tbl.findall(f"{W}tr")
    citation_total = 0

    for r_idx, tr in enumerate(rows):
        next_tr = rows[r_idx + 1] if r_idx + 1 < len(rows) else None
        is_title_row = r_idx == 0 and _is_main_title_table_row(tr, next_tr)
        tcs = tr.findall(f"{W}tc")
        for c_idx, tc in enumerate(tcs):
            if c_idx == 0:
                c_font, c_tsz, c_bsz, c_csz, c_ldxa = "宋体", 28, 24, 21, 400
            elif c_idx == 1:
                c_font, c_tsz, c_bsz, c_csz, c_ldxa = "Georgia", 28, 24, 21, 400
            else:
                c_font, c_tsz, c_bsz, c_csz, c_ldxa = col3_font, col3_title_sz, col3_body_sz, col3_cit_sz, col3_line_dxa

            for p_idx, p in enumerate(tc.findall(f"{W}p")):
                ppr = p.find(f"{W}pPr")
                if ppr is None:
                    ppr = ET.Element(f"{W}pPr")
                    p.insert(0, ppr)
                sp = _upsert_ppr_child(ppr, "spacing")
                for rm_attr in ("beforeLines", "afterLines", "beforeAutospacing", "afterAutospacing"):
                    sp.attrib.pop(f"{W}{rm_attr}", None)
                sp.attrib[f"{W}before"] = "0"
                sp.attrib[f"{W}after"] = "0"
                sp.attrib[f"{W}line"] = str(c_ldxa)
                sp.attrib[f"{W}lineRule"] = "exact"
                _sort_ppr_children(ppr)

                p_hints = citation_hints_map.get((r_idx, p_idx))
                citation_total += apply_font_sizes_to_paragraph(
                    p,
                    is_main_title_row=is_title_row,
                    col_idx=c_idx,
                    font_name=c_font,
                    title_sz=c_tsz,
                    body_sz=c_bsz,
                    citation_sz=c_csz,
                    citation_hints=p_hints,
                )

    return citation_total


def _estimate_visual_dxa_length(text: str, half_pt_sz: int = 24) -> float:
    """
    Estimate effective horizontal unwrapped line length (in dxa, including word-wrap raggedness)
    for a single paragraph at `half_pt_sz` (default 24 = 12pt) to equalize vertical column heights
    across ALL scripts (CJK, Latin/Cyrillic, Khmer, Lao, Thai, Myanmar, Devanagari, Arabic, etc.).
    """
    if not text:
        return 0.0
    sz_scale = max(0.6, float(half_pt_sz) / 24.0)
    total = 0.0
    se_asian_chars = 0
    latin_chars = 0
    non_ws_chars = 0

    prev_is_coeng = False
    for ch in text:
        if ch in "\r\n":
            continue
        cat = unicodedata.category(ch)
        cp = ord(ch)
        if cp == 0x17D2:  # Khmer coeng (makes next consonant subscript, 0 horizontal advance)
            prev_is_coeng = True
            continue
        if prev_is_coeng:
            prev_is_coeng = False
            continue
        if cat in ("Mn", "Me", "Cf"):
            continue
        if ch == " ":
            total += 68.0 * sz_scale
            continue
        non_ws_chars += 1
        if cat == "Mc":
            # Spacing vowel signs (Khmer ា, ើ, ៅ, Devanagari ी, ा, etc.) DO occupy horizontal width!
            total += 82.0 * sz_scale
        elif (
            0x4E00 <= cp <= 0x9FFF
            or 0x3400 <= cp <= 0x4DBF
            or 0x3000 <= cp <= 0x303F
            or 0xFF00 <= cp <= 0xFFEF
            or 0xAC00 <= cp <= 0xD7AF
        ):
            total += 240.0 * sz_scale
        elif 0x1780 <= cp <= 0x17FF:  # Khmer (Leelawadee UI wide glyphs + cluster wrapping)
            total += 168.0 * sz_scale
            se_asian_chars += 1
        elif 0x0E80 <= cp <= 0x0EFF:  # Lao (DokChampa wide rounded glyphs + phrase wrapping)
            total += 162.0 * sz_scale
            se_asian_chars += 1
        elif 0x1000 <= cp <= 0x109F:  # Myanmar
            total += 155.0 * sz_scale
            se_asian_chars += 1
        elif 0x0E00 <= cp <= 0x0E7F:  # Thai (Note: Angsana New is condensed per pt)
            total += 112.0 * sz_scale
            se_asian_chars += 1
        elif 0x0900 <= cp <= 0x0DFF:  # Devanagari / Indic scripts
            total += 145.0 * sz_scale
            latin_chars += 1
        elif ch.isupper():
            total += 152.0 * sz_scale
            latin_chars += 1
        else:
            total += 122.0 * sz_scale
            latin_chars += 1

    if non_ws_chars > 0:
        # Word-wrap ragged right margin factor: CJK wraps at any char (~1.03x); space/cluster scripts
        # leave unused space at the end of wrapped lines (~1.12x for Latin, ~1.22x for Khmer/Lao phrases).
        if se_asian_chars > non_ws_chars * 0.35:
            total *= 1.22
        elif latin_chars > non_ws_chars * 0.35:
            total *= 1.11
        else:
            total *= 1.03
    return total


def _solve_equal_height_widths(
    row_para_dxa: List[Tuple[List[float], List[float], List[float]]],
    line_dxas: Tuple[int, int, int],
    total_w: int,
) -> Tuple[int, int, int]:
    """
    Find `(w1, w2, w3)` summing to `total_w` that minimizes total row-by-row vertical height
    discrepancy across all 3 columns after content wrapping — adapting automatically to ANY
    language (whether `20:35:45`, `18:27:55`, or `30:33:37`)!
    """
    CELL_PAD_DXA = 216
    if not row_para_dxa:
        w1 = int(round(total_w * 0.26))
        w2 = int(round(total_w * 0.37))
        return w1, w2, total_w - w1 - w2

    def _loss(r1: float, r2: float) -> float:
        r3 = 1.0 - r1 - r2
        u1 = max(1200.0, total_w * r1 - CELL_PAD_DXA)
        u2 = max(1400.0, total_w * r2 - CELL_PAD_DXA)
        u3 = max(1500.0, total_w * r3 - CELL_PAD_DXA)
        usables = (u1, u2, u3)
        total_loss = 0.0
        col_totals = [0.0, 0.0, 0.0]
        for p_lists in row_para_dxa:
            h_cols = [0.0, 0.0, 0.0]
            for c_i in range(3):
                u_c = usables[c_i]
                ld = line_dxas[c_i]
                lines_c = 0
                for pdxa in p_lists[c_i]:
                    if pdxa <= 1.0:
                        lines_c += 1
                    else:
                        lines_c += max(1, int((pdxa + u_c - 1.0) // u_c))
                h_cols[c_i] = lines_c * ld
                col_totals[c_i] += h_cols[c_i]
            gap = max(h_cols) - min(h_cols)
            total_loss += gap * gap
        # Also penalize overall cumulative column height imbalance across the whole table
        tot_gap = max(col_totals) - min(col_totals)
        return total_loss + 0.35 * (tot_gap * tot_gap) / max(1, len(row_para_dxa))

    # Coarse grid search (step = 0.01) within readable column bounds:
    # Col 1 Chinese: 16% .. 38% | Col 2 English: 20% .. 48% | Col 3 Foreign: 22% .. 60%
    best_r1, best_r2 = 0.25, 0.36
    best_score = float("inf")
    for i1 in range(16, 39):
        r1 = i1 / 100.0
        for i2 in range(20, 49):
            r2 = i2 / 100.0
            r3 = 1.0 - r1 - r2
            if r3 < 0.22 or r3 > 0.60:
                continue
            sc = _loss(r1, r2)
            if sc < best_score:
                best_score = sc
                best_r1, best_r2 = r1, r2

    # Fine local refinement (step = 0.0025)
    for d1 in range(-4, 5):
        r1 = best_r1 + d1 * 0.0025
        if not (0.16 <= r1 <= 0.38):
            continue
        for d2 in range(-4, 5):
            r2 = best_r2 + d2 * 0.0025
            r3 = 1.0 - r1 - r2
            if not (0.20 <= r2 <= 0.48 and 0.22 <= r3 <= 0.60):
                continue
            sc = _loss(r1, r2)
            if sc < best_score:
                best_score = sc
                best_r1, best_r2 = r1, r2

    w1 = int(round(total_w * best_r1))
    w2 = int(round(total_w * best_r2))
    w3 = total_w - w1 - w2
    return w1, w2, w3


def _ensure_standard_table_borders(tblPr: ET.Element) -> None:
    tblBorders = tblPr.find(f"{W}tblBorders")
    if tblBorders is not None:
        return
    tblBorders = ET.Element(f"{W}tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b_el = ET.SubElement(tblBorders, f"{W}{side}")
        b_el.attrib[f"{W}val"] = "single"
        b_el.attrib[f"{W}sz"] = "4"
        b_el.attrib[f"{W}space"] = "0"
        b_el.attrib[f"{W}color"] = "auto"
    tblLayout = tblPr.find(f"{W}tblLayout")
    if tblLayout is not None:
        tblPr.insert(list(tblPr).index(tblLayout), tblBorders)
    else:
        tblPr.append(tblBorders)


def balance_three_column_widths(
    root: ET.Element,
    tbl: ET.Element,
    lang_typo: Optional[Dict] = None,
) -> Tuple[int, int, int]:
    """Strip any staggered gridSpan/gridBefore/vMerge props, rebuild a strict 3-column <w:tblGrid>,
    and dynamically solve for `(w1, w2, w3)` so the rendered heights of all 3 columns match closely.
    """
    strip_staggered_grid_props_from_table(tbl)

    total_w = 0
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

    tblPr = tbl.find(f"{W}tblPr")
    if tblPr is None:
        tblPr = ET.Element(f"{W}tblPr")
        tbl.insert(0, tblPr)

    if total_w < 6000:
        tblW_existing = tblPr.find(f"{W}tblW")
        if tblW_existing is not None and tblW_existing.attrib.get(f"{W}type") == "dxa":
            try:
                total_w = int(float(tblW_existing.attrib.get(f"{W}w", "0")))
            except ValueError:
                total_w = 0
    if total_w < 6000:
        total_w = 14570

    col3_bsz = int((lang_typo or {}).get("body_sz", 24))
    col3_ldxa = int((lang_typo or {}).get("line_dxa", 400))
    col_szs = (24, 24, col3_bsz)
    line_dxas = (400, 400, col3_ldxa)

    row_para_dxa: List[Tuple[List[float], List[float], List[float]]] = []
    for tr in tbl.findall(f"{W}tr"):
        tcs = tr.findall(f"{W}tc")
        if len(tcs) < 3:
            continue
        p_lists: List[List[float]] = [[], [], []]
        has_any = False
        for c_i in range(3):
            for p in tcs[c_i].findall(f"{W}p"):
                ptxt = _get_p_text(p)
                if ptxt.strip():
                    has_any = True
                p_lists[c_i].append(_estimate_visual_dxa_length(ptxt, half_pt_sz=col_szs[c_i]))
        if has_any:
            row_para_dxa.append((p_lists[0], p_lists[1], p_lists[2]))

    w1, w2, w3 = _solve_equal_height_widths(row_para_dxa, line_dxas=line_dxas, total_w=total_w)

    tblW = tblPr.find(f"{W}tblW")
    if tblW is None:
        tblW = ET.SubElement(tblPr, f"{W}tblW")
    tblW.attrib[f"{W}w"] = str(total_w)
    tblW.attrib[f"{W}type"] = "dxa"

    # Normalize tblInd so merged rows never shift horizontally
    tblInd = tblPr.find(f"{W}tblInd")
    if tblInd is not None:
        tblInd.attrib[f"{W}w"] = "0"
        tblInd.attrib[f"{W}type"] = "dxa"

    _ensure_standard_table_borders(tblPr)

    tblLayout = tblPr.find(f"{W}tblLayout")
    if tblLayout is None:
        tblLayout = ET.SubElement(tblPr, f"{W}tblLayout")
    tblLayout.attrib[f"{W}type"] = "fixed"

    # Rebuild strict 3-column <w:tblGrid>
    tblGrid = tbl.find(f"{W}tblGrid")
    if tblGrid is None:
        tblGrid = ET.Element(f"{W}tblGrid")
        tbl.insert(1, tblGrid)
    for ch in list(tblGrid):
        tblGrid.remove(ch)
    for col_w in (w1, w2, w3):
        gc = ET.SubElement(tblGrid, f"{W}gridCol")
        gc.attrib[f"{W}w"] = str(col_w)

    # Set uniform <w:tcW> at top of <w:tcPr> across all 3 cells in every row
    widths_tuple = (w1, w2, w3)
    for tr in tbl.findall(f"{W}tr"):
        tcs = tr.findall(f"{W}tc")
        if len(tcs) == 3:
            for c_i, tc in enumerate(tcs):
                tcPr = tc.find(f"{W}tcPr")
                if tcPr is None:
                    tcPr = ET.Element(f"{W}tcPr")
                    tc.insert(0, tcPr)
                for nw in list(tcPr.findall(f"{W}noWrap")):
                    tcPr.remove(nw)
                tcW = tcPr.find(f"{W}tcW")
                if tcW is None:
                    tcW = ET.Element(f"{W}tcW")
                    insert_idx = 1 if (len(tcPr) > 0 and _local_tag(tcPr[0]) == "cnfStyle") else 0
                    tcPr.insert(insert_idx, tcW)
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

    fd, tmp_name = tempfile.mkstemp(prefix=".fel_", suffix=".docx", dir=str(dst_docx.parent))
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        with zipfile.ZipFile(src_docx, "r") as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            cleaned_rels_bytes: Optional[bytes] = None
            valid_rel_ids: Optional[Set[str]] = None
            if "word/_rels/document.xml.rels" in zin.namelist():
                raw_rels = _read_zip_member_bounded(zin, "word/_rels/document.xml.rels")
                cleaned_rels_bytes, valid_rel_ids = strip_comment_and_people_rels_bytes(raw_rels)

            sanitize_fellowship_ooxml(doc_root, valid_rel_ids=valid_rel_ids)
            for _pfx, _uri in COMMON_NAMESPACES.items():
                ET.register_namespace(_pfx, _uri)
            doc_xml_body = ET.tostring(doc_root, encoding="unicode")
            doc_bytes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + doc_xml_body).encode("utf-8")

            kept_names_low = [
                it.filename.replace("\\", "/").lower()
                for it in zin.infolist()
                if not docx_xml_utils._is_slimmed_bloat_member(it.filename)
                and not it.filename.replace("\\", "/").lower().startswith("word/comments")
                and it.filename.replace("\\", "/").lower() != "word/people.xml"
            ]
            has_emf = any(n.endswith(".emf") for n in kept_names_low)
            has_wmf = any(n.endswith(".wmf") for n in kept_names_low)

            for item in zin.infolist():
                fn = item.filename
                fn_low = fn.replace("\\", "/").lower()
                # Drop all comment XML parts, people.xml, embedded fonts, and thumbnail bloat
                if (
                    fn_low.startswith("word/comments")
                    or fn_low == "word/people.xml"
                    or docx_xml_utils._is_slimmed_bloat_member(fn)
                ):
                    continue

                zi = docx_xml_utils._clone_deflated_zipinfo(item)
                if fn_low == "word/document.xml":
                    zout.writestr(zi, doc_bytes, compress_type=zipfile.ZIP_DEFLATED)
                    continue

                raw = _read_zip_member_bounded(zin, item)
                if fn_low == "word/settings.xml":
                    zout.writestr(
                        zi,
                        docx_xml_utils._slim_settings_xml(raw, disable_track_revisions=True),
                        compress_type=zipfile.ZIP_DEFLATED,
                    )
                elif fn_low == "word/fonttable.xml":
                    zout.writestr(
                        zi,
                        docx_xml_utils._slim_font_table_xml(raw),
                        compress_type=zipfile.ZIP_DEFLATED,
                    )
                elif fn_low == "_rels/.rels":
                    zout.writestr(
                        zi,
                        docx_xml_utils._slim_root_rels_xml(raw),
                        compress_type=zipfile.ZIP_DEFLATED,
                    )
                elif fn_low == "word/_rels/document.xml.rels" and cleaned_rels_bytes is not None:
                    zout.writestr(zi, cleaned_rels_bytes, compress_type=zipfile.ZIP_DEFLATED)
                elif fn_low == "[content_types].xml":
                    ct_slim = docx_xml_utils._patch_content_types_xml(
                        raw,
                        ensure_comments_ct=False,
                        keep_people=False,
                        has_emf_media=has_emf,
                        has_wmf_media=has_wmf,
                    )
                    zout.writestr(zi, ct_slim, compress_type=zipfile.ZIP_DEFLATED)
                else:
                    zout.writestr(zi, raw, compress_type=zipfile.ZIP_DEFLATED)
        os.replace(str(tmp_path), str(dst_docx))
    except Exception:
        try:
            if tmp_path.exists():
                tmp_path.unlink()
        except Exception:
            pass
        raise


def inspect_fellowship_docx_quick(docx_path: Path) -> Dict:
    """Fast local inspection for GUI preview table (title, part_num, row_count, internal_parts)."""
    validate_docx_archive(docx_path)
    with zipfile.ZipFile(docx_path, "r") as zf:
        root = _parse_safe_xml_bytes(_read_zip_member_bounded(zf, "word/document.xml"))
    body = root.find(f"{W}body")
    tbls = body.findall(f"{W}tbl") if body is not None else []
    row_cnt = sum(len(t.findall(f"{W}tr")) for t in tbls)

    raw_title = ""
    if body is not None:
        for el in body:
            if _local_tag(el) == "tbl":
                break
            if _local_tag(el) == "p":
                tx = _get_p_text(el).strip()
                if tx and not _is_boilerplate_header_paragraph(tx) and not re.match(r"^\d+\.\s+[a-z]", tx):
                    raw_title = tx
    if not raw_title and tbls and len(tbls[0].findall(f"{W}tr")) > 0:
        trs0 = tbls[0].findall(f"{W}tr")
        if _is_main_title_table_row(trs0[0], trs0[1] if len(trs0) > 1 else None):
            tcs0 = trs0[0].findall(f"{W}tc")
            if tcs0:
                raw_title = _get_tc_text(tcs0[0]).strip()
    if not raw_title:
        raw_title = clean_title_noise(docx_path.stem, is_filename_stem=True)

    base_title, part_num = parse_title_and_part(raw_title, docx_path.name)
    if len(tbls) >= 2:
        part_num = None
    return {
        "path": str(docx_path),
        "filename": docx_path.name,
        "base_title": base_title or clean_title_noise(docx_path.stem, is_filename_stem=True),
        "part_num": part_num,
        "row_count": row_cnt,
        "table_count": len(tbls),
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
    clean_title = sanitize_filename_part(clean_title_noise(base_title, is_filename_stem=False) or base_title)
    dt = re.sub(r"[^0-9]", "", date_yyyymmdd or "") or datetime.now().strftime("%Y%m%d")
    if part_num is not None:
        return f"{pfx}_{clean_title}-{part_num:02d}-三语稿-{dt}.docx"
    return f"{pfx}_{clean_title}-三语稿-{dt}.docx"


def run_fellowship_pipeline(
    input_files: List[str],
    output_dir: str,
    output_mode: str = "single_only",  # "single_only", "merged_only", "both"
    lang_code: str = "AUTO",
    tri_prefix_override: str = "",
    quote_mode: str = "auto",  # "auto", "curly", "skip"
    date_yyyymmdd: str = "",
    enable_ai_col3: bool = True,
    sync_red: bool = False,
    api_key=None,
    model: str = "gemini-3.5-flash-lite",
    api_mode: str = "auto",
    rotation_mode: str = "round_robin",
    log_fn: Optional[Callable[[str], None]] = None,
    progress_fn: Optional[Callable[[float, str], None]] = None,
) -> Dict:
    """Execute Multi-Part / Fellowship Manuscript cleaning, styling, typography, and merging."""

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

    # Fallback script detection from first document's Col 3 text when filename had no language keyword/code
    if resolved_code == "AUTO":
        try:
            validate_docx_archive(sorted_paths[0])
            with zipfile.ZipFile(sorted_paths[0], "r") as zf0:
                r0_xml = _parse_safe_xml_bytes(_read_zip_member_bounded(zf0, "word/document.xml"))
            tbl0 = r0_xml.find(f".//{W}tbl")
            if tbl0 is not None:
                c3_sample = "".join(
                    _get_tc_text(tr.findall(f"{W}tc")[2])[:400]
                    for tr in tbl0.findall(f"{W}tr")[:6]
                    if len(tr.findall(f"{W}tc")) >= 3
                )
                det_script = detect_lang_code_from_col3_script(c3_sample)
                if det_script:
                    resolved_code = det_script
        except Exception:
            pass

    lang_typo = get_lang_typography(resolved_code, langs)

    tri_prefix = (tri_prefix_override or "").strip().rstrip("_")
    if not tri_prefix:
        tri_prefix = get_tri_prefix(resolved_code, langs)

    dt_str = re.sub(r"[^0-9]", "", date_yyyymmdd or "") or datetime.now().strftime("%Y%m%d")

    pool = None
    if enable_ai_col3:
        if isinstance(api_key, ApiKeyPool):
            pool = api_key
        elif api_key:
            try:
                pool = ApiKeyPool(api_key, rotation_mode=rotation_mode)
            except ValueError:
                pool = None

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
            doc_root = _parse_safe_xml_bytes(_read_zip_member_bounded(zf, "word/document.xml"))

        # 1. Unwrap <w:sdt> content controls
        sdt_n = unwrap_all_sdts(doc_root)
        # 2. Accept track changes & strip comments + highlights
        rev_n, cmt_n = accept_revisions_and_strip_comments(doc_root)
        # 3. Single-doc multi-part detection & internal table merge (Feedback #3)
        was_internally_merged, merged_extra_tbls = merge_internal_part_tables_in_body(
            doc_root, output_mode=output_mode
        )
        # 4. Enforce Landscape A4 page setup
        enforce_landscape_page_setup(doc_root)
        # 5. Clean boilerplate header & normalize title (stripping part1 if single-doc multi-part merged)
        base_title, part_num, title_p = clean_body_header_and_extract_title(
            doc_root,
            src_p.name,
            force_strip_part=(was_internally_merged and total_files == 1),
        )

        if part_num is None and total_files > 1:
            part_num = idx + 1
            set_clean_top_title_paragraph(title_p, f"{base_title}-{part_num:02d}")

        body = doc_root.find(f"{W}body")
        tbl = body.find(f"{W}tbl")

        # 6. Strip trailing spaces & trailing empty paragraphs at bottom of non-empty cells (Feedback #6)
        clean_table_trailing_spaces_and_empty_paragraphs(tbl)

        # 7. Pre-capture Col 1 & Col 2 source citation hints (including original sz=21) BEFORE size overwrite
        cit_hints_map = capture_table_citation_hints(tbl)

        # Auto-detect guillemets « » in auto quote mode so French/Greek guillemets are preserved cleanly
        use_q_col3 = apply_q_col3
        if quote_mode == "auto" and use_q_col3:
            sample_c3 = "".join(
                _get_tc_text(tr.findall(f"{W}tc")[2])[:300]
                for tr in tbl.findall(f"{W}tr")[:5]
                if len(tr.findall(f"{W}tc")) >= 3
            )
            if "«" in sample_c3 and "»" in sample_c3:
                use_q_col3 = False

        # 8. Curly quotes on Col 2 & Col 3 BEFORE span alignment
        apply_curly_quotes_to_table(
            tbl=tbl,
            apply_curly_quotes_col2=apply_q_col2,
            apply_curly_quotes_col3=use_q_col3,
        )

        # 9. Sync Col 3 blue (#0000FF), optional red (#FF0000), & bold (<w:b/>) BEFORE citation font-size slicing
        sync_stats = sync_table_col3_blue_and_bold(
            tbl=tbl,
            api_pool=pool,
            model=model,
            api_mode=api_mode,
            enable_ai_col3=enable_ai_col3,
            sync_red=sync_red,
            log_fn=log,
        )
        if sync_stats.get("failed_rows"):
            ai_failed_items.append(f"{src_p.name}（第 {sync_stats['failed_rows']} 行）")

        # 10. Standard fonts + Thai/default sizes + multi-script & God's Word citations + exact 20pt spacing
        cit_cnt = enforce_table_spacing_and_typography(
            tbl=tbl,
            lang_code=resolved_code,
            lang_typo=lang_typo,
            citation_hints_map=cit_hints_map,
        )

        # 11. Strip staggered gridSpan props & balance 3-column widths
        w_tuple = balance_three_column_widths(doc_root, tbl, lang_typo=lang_typo)

        merge_note = f" | 文档内合并 {merged_extra_tbls + 1} 个part表" if was_internally_merged else ""
        font_note = f" | 外文字体 {lang_typo['font']}" if lang_typo.get("font") else ""
        log(
            f"   ✓ 拆控件 {sdt_n} 处 | 清批注/修订 {cmt_n + rev_n} 处{merge_note} | "
            f"第3列样式同步(确定性 {sync_stats['tier1_synced']} 处, AI/本地 {sync_stats['ai_synced'] + sync_stats.get('local_fallback_synced', 0)} 段) | "
            f"摘自字号 {cit_cnt} 处{font_note} | 列宽 {w_tuple}"
        )

        processed_docs.append(
            {
                "src_path": src_p,
                "root": doc_root,
                "base_title": base_title,
                "part_num": part_num,
                "title_p": title_p,
                "tbl": tbl,
                "was_internally_merged": was_internally_merged,
            }
        )

    saved_files: List[str] = []
    saved_paths_seen: Set[str] = set()

    if output_mode in ("single_only", "both"):
        prog(84.0, "正在导出单篇整理定稿（修订模式：关闭）...")
        for item in processed_docs:
            # When output_mode == 'both' and there is only 1 already-internally-merged file (part_num=None),
            # defer to the merged output block so we don't write the identical filename twice.
            if output_mode == "both" and total_files == 1 and item["part_num"] is None:
                continue
            out_name = build_fellowship_output_name(
                tri_prefix=tri_prefix,
                base_title=item["base_title"],
                date_yyyymmdd=dt_str,
                part_num=item["part_num"],
            )
            dst_p = out_dir_p / out_name
            if item["part_num"] is not None:
                set_clean_top_title_paragraph(item["title_p"], f"{item['base_title']}-{item['part_num']:02d}")
            else:
                set_clean_top_title_paragraph(item["title_p"], item["base_title"])
            write_cleaned_docx(item["src_path"], dst_p, item["root"])
            saved_files.append(str(dst_p))
            saved_paths_seen.add(str(dst_p))
            log(f"✅ 已输出单篇：{dst_p.name}")

    if output_mode in ("merged_only", "both"):
        prog(92.0, "正在按顺序合并总文档并全局均衡三列宽度...")
        first = processed_docs[0]
        merged_root = copy.deepcopy(first["root"])
        enforce_landscape_page_setup(merged_root)
        merged_body = merged_root.find(f"{W}body")
        merged_title_p = merged_body.find(f"{W}p")
        merged_tbl = merged_body.find(f"{W}tbl")

        if merged_title_p is not None:
            set_clean_top_title_paragraph(merged_title_p, first["base_title"])

        p1_rows = merged_tbl.findall(f"{W}tr")
        p1_headers = p1_rows[:2] if len(p1_rows) >= 2 else p1_rows[:1]

        for extra in processed_docs[1:]:
            _append_part_rows_into_table(merged_tbl, extra["tbl"], p1_headers, deep_copy=True)

        merged_widths = balance_three_column_widths(merged_root, merged_tbl, lang_typo=lang_typo)
        total_merged_rows = len(merged_tbl.findall(f"{W}tr"))

        merged_name = build_fellowship_output_name(
            tri_prefix=tri_prefix,
            base_title=first["base_title"],
            date_yyyymmdd=dt_str,
            part_num=None,
        )
        merged_dst = out_dir_p / merged_name
        write_cleaned_docx(first["src_path"], merged_dst, merged_root)
        if str(merged_dst) not in saved_paths_seen:
            saved_files.append(str(merged_dst))
            saved_paths_seen.add(str(merged_dst))
        log(f"🎉 已输出汇总合并总文档（共 {total_merged_rows} 行 | 列宽 {merged_widths}）：{merged_dst.name}")

    prog(100.0, f"整理完成！共输出 {len(saved_files)} 份文档")
    return {
        "saved_files": saved_files,
        "tri_prefix": tri_prefix,
        "lang_code": resolved_code,
        "font": lang_typo.get("font", ""),
        "ai_failed_items": ai_failed_items,
    }
