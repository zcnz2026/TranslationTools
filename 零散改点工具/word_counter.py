import copy
from datetime import datetime
import os
from pathlib import Path
import re
import shutil
import sys
from typing import Optional
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
from lang_registry import load_languages  # noqa: E402

# Matches new & legacy word-count tags:
#   _(UP 120 words, RV 45 words)
#   _(UP  words, RV  words)
#   _(UP 120 words) / _(RV 45 words)
#   _(up452 rv609) / _(up0)
_WORD_COUNT_TAG_RE = re.compile(
    r"[\s_]*\(\s*(?:"
    r"(?:UP|RV)\s*\d*\s*(?:words?)?(?:\s*[,，/&\s]\s*(?:UP|RV)\s*\d*\s*(?:words?)?)*"
    r"|"
    r"(?:up|rv)\s*\d+(?:\s+(?:up|rv)\s*\d+)*"
    r")\s*\)",
    re.IGNORECASE,
)

_TRAILING_DATE_RE = re.compile(r"_((?:19|20)\d{6}|\d{4})(?:--[^\s/\\]+)?$")
_CJK_CHAR_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")

_TOP_INSTRUCTION_NORMS = frozenset({
    "blue highlighted parts need to be reviewed",
    "yellow highlighted parts need to be updated",
})


# ---------------------------------------------------------------------------
# Cross-branch XML helpers (delegate to docx_xml_utils when available, with
# identical local fallbacks so standalone tests pass before branch merge)
# ---------------------------------------------------------------------------

def format_up_rv_count_tag(up_words: int, rv_words: int) -> str:
    """
    Format UP/RV word count tag omitting zero counts:
      - up > 0, rv > 0 -> '(UP X words, RV Y words)'
      - up > 0, rv == 0 -> '(UP X words)'
      - up == 0, rv > 0 -> '(RV Y words)'
      - up == 0, rv == 0 -> ''
    """
    fn = getattr(docx_xml_utils, "format_up_rv_count_tag", None)
    if callable(fn):
        return fn(int(up_words or 0), int(rv_words or 0))
    u = max(0, int(up_words or 0))
    r = max(0, int(rv_words or 0))
    if u > 0 and r > 0:
        return f"(UP {u} words, RV {r} words)"
    if u > 0:
        return f"(UP {u} words)"
    if r > 0:
        return f"(RV {r} words)"
    return ""


def is_top_highlight_instruction_text(text: str) -> bool:
    norm = re.sub(r"\s+", " ", (text or "").strip()).rstrip(".").strip().lower()
    return norm in _TOP_INSTRUCTION_NORMS


def extract_full_article_en_title(root: ET.Element) -> str:
    """
    Extract clean English chapter title from Row 0 Col 2 of the first <w:tbl> in `root`
    when Row 0 looks like a chapter title row.
    """
    fn = getattr(docx_xml_utils, "extract_full_article_en_title", None)
    if callable(fn):
        res = fn(root)
        if res:
            return re.sub(r"\s+", " ", res).strip()

    ns = docx_xml_utils.NS
    tbl = root.find(".//w:tbl", ns)
    if tbl is None:
        return ""
    trs = tbl.findall("w:tr", ns)
    if not trs:
        return ""
    tcs0 = trs[0].findall("w:tc", ns)
    if len(tcs0) < 2:
        return ""

    cn0 = re.sub(r"\s+", " ", docx_xml_utils.get_cell_plain_text(tcs0[0])).strip()
    en0 = re.sub(r"\s+", " ", docx_xml_utils.get_cell_plain_text(tcs0[1])).strip()
    if not en0 or len(en0) > 200:
        return ""

    # Handle 2-row Appendix headers (e.g. Row 0 = 'Appendix 1:', Row 1 = 'The Appearance of God...', Row 2 = '')
    if len(trs) >= 3 and en0.endswith(":") and len(en0) <= 24:
        tcs2 = trs[2].findall("w:tc", ns)
        row2_blank = bool(tcs2) and all(not docx_xml_utils.get_cell_plain_text(c).strip() for c in tcs2)
        if row2_blank:
            tcs1 = trs[1].findall("w:tc", ns)
            if len(tcs1) >= 2:
                en1 = re.sub(r"\s+", " ", docx_xml_utils.get_cell_plain_text(tcs1[1])).strip()
                if en1 and len(en1) <= 180:
                    return f"{en0} {en1}"

    has_bold = (
        tcs0[1].find(".//w:b", ns) is not None
        or tcs0[0].find(".//w:b", ns) is not None
        or tcs0[0].find(".//w:outlineLvl", ns) is not None
        or tcs0[1].find(".//w:outlineLvl", ns) is not None
    )
    row1_blank = False
    if len(trs) >= 2:
        tcs1 = trs[1].findall("w:tc", ns)
        row1_blank = bool(tcs1) and all(not docx_xml_utils.get_cell_plain_text(c).strip() for c in tcs1)

    short_heading_shape = (
        len(en0) <= 180
        and len(cn0) <= 60
        and "。" not in cn0
        and not en0.endswith(".")
    )
    if has_bold or row1_blank or short_heading_shape:
        return en0
    return ""


def _make_top_instruction_paragraph(text: str, highlight_val: str) -> ET.Element:
    W = docx_xml_utils.W
    p_el = ET.Element(f"{W}p")
    r_el = ET.Element(f"{W}r")
    rPr = ET.Element(f"{W}rPr")
    ET.SubElement(
        rPr,
        f"{W}rFonts",
        {
            f"{W}ascii": "Georgia",
            f"{W}eastAsia": "宋体",
            f"{W}hAnsi": "Georgia",
            f"{W}cs": "Georgia",
        },
    )
    ET.SubElement(rPr, f"{W}sz", {f"{W}val": "28"})
    ET.SubElement(rPr, f"{W}szCs", {f"{W}val": "28"})
    ET.SubElement(rPr, f"{W}highlight", {f"{W}val": highlight_val})
    r_el.append(rPr)
    t_el = ET.Element(f"{W}t")
    t_el.text = text
    r_el.append(t_el)
    p_el.append(r_el)
    return p_el


def update_top_highlight_instruction_lines(root: ET.Element, has_up: bool, has_rv: bool) -> None:
    """
    Synchronize leading 'Blue/Yellow highlighted parts...' instruction lines at the top of `w:body`.
    """
    fn = getattr(docx_xml_utils, "update_top_highlight_instruction_lines", None)
    if callable(fn):
        fn(root, bool(has_up), bool(has_rv))
        return

    ns = docx_xml_utils.NS
    body = root.find("w:body", ns)
    if body is None:
        return

    for ch in list(body):
        loc = ch.tag.split("}")[-1]
        if loc == "tbl":
            break
        if loc == "p":
            ptxt = "".join(t.text or "" for t in ch.findall(".//w:t", ns)).strip()
            if is_top_highlight_instruction_text(ptxt):
                body.remove(ch)

    to_insert = []
    if has_rv:
        to_insert.append(_make_top_instruction_paragraph("Blue highlighted parts need to be reviewed.", "cyan"))
    if has_up:
        to_insert.append(_make_top_instruction_paragraph("Yellow highlighted parts need to be updated.", "yellow"))

    for idx, p_el in enumerate(to_insert):
        body.insert(idx, p_el)


def balance_trilingual_table_widths(root: ET.Element) -> int:
    """
    Adaptively balance 3-column table widths (`三语拉齐`) across all 3-cell tables in `root`.
    """
    fn = getattr(docx_xml_utils, "balance_trilingual_table_widths", None)
    if callable(fn):
        return int(fn(root))

    ns = docx_xml_utils.NS
    W = docx_xml_utils.W
    span_tags = {"gridSpan", "gridBefore", "gridAfter", "wBefore", "wAfter"}
    total_dxa = 14560
    balanced_cnt = 0

    def _text_weight(txt: str) -> float:
        if not txt:
            return 0.0
        non_ws = re.sub(r"\s+", "", txt)
        if not non_ws:
            return 0.0
        cjk = len(_CJK_CHAR_RE.findall(non_ws))
        other = len(non_ws) - cjk
        return cjk * 1.8 + other * 1.0

    for tbl in root.findall(".//w:tbl", ns):
        trs = tbl.findall("w:tr", ns)
        three_cell_trs = [tr for tr in trs if len(tr.findall("w:tc", ns)) == 3]
        if not three_cell_trs:
            continue

        weights = [0.0, 0.0, 0.0]
        for tr in three_cell_trs:
            tcs = tr.findall("w:tc", ns)
            for c_i in range(3):
                weights[c_i] += _text_weight(docx_xml_utils.get_cell_plain_text(tcs[c_i]))

        tot_w = sum(weights)
        if tot_w <= 0:
            pcts = [1100, 1900, 2000]
        else:
            raw = [round(5000.0 * w / tot_w) for w in weights]
            bounds = [(950, 1250), (1700, 2050), (1800, 2150)]
            pcts = [max(lo, min(hi, val)) for val, (lo, hi) in zip(raw, bounds)]
            diff = 5000 - sum(pcts)
            # Distribute remainder across columns that still have headroom within bounds
            for _ in range(10):
                if diff == 0:
                    break
                for idx_c in (2, 1, 0):
                    if diff == 0:
                        break
                    lo, hi = bounds[idx_c]
                    if diff > 0 and pcts[idx_c] < hi:
                        add = min(diff, hi - pcts[idx_c])
                        pcts[idx_c] += add
                        diff -= add
                    elif diff < 0 and pcts[idx_c] > lo:
                        sub = min(-diff, pcts[idx_c] - lo)
                        pcts[idx_c] -= sub
                        diff += sub
            if sum(pcts) != 5000:
                pcts[2] += 5000 - sum(pcts)

        dxa_cols = [round(total_dxa * pcts[0] / 5000), round(total_dxa * pcts[1] / 5000)]
        dxa_cols.append(total_dxa - dxa_cols[0] - dxa_cols[1])

        tblPr = tbl.find("w:tblPr", ns)
        if tblPr is None:
            tblPr = ET.Element(f"{W}tblPr")
            tbl.insert(0, tblPr)
        tblW = tblPr.find("w:tblW", ns)
        if tblW is None:
            tblW = ET.SubElement(tblPr, f"{W}tblW")
        tblW.attrib.clear()
        tblW.set(f"{W}w", "5000")
        tblW.set(f"{W}type", "pct")

        tblGrid = tbl.find("w:tblGrid", ns)
        if tblGrid is None:
            tblGrid = ET.Element(f"{W}tblGrid")
            ins_idx = 1 if len(tbl) > 0 and tbl[0] is tblPr else 0
            tbl.insert(ins_idx, tblGrid)
        else:
            for ch in list(tblGrid):
                tblGrid.remove(ch)
        for dxa_v in dxa_cols:
            ET.SubElement(tblGrid, f"{W}gridCol", {f"{W}w": str(dxa_v)})

        for tr in three_cell_trs:
            trPr = tr.find("w:trPr", ns)
            if trPr is not None:
                for ch in list(trPr):
                    if ch.tag.split("}")[-1] in span_tags:
                        trPr.remove(ch)
            for idx_c, tc in enumerate(tr.findall("w:tc", ns)):
                tcPr = tc.find("w:tcPr", ns)
                if tcPr is None:
                    tcPr = ET.Element(f"{W}tcPr")
                    tc.insert(0, tcPr)
                for ch in list(tcPr):
                    if ch.tag.split("}")[-1] in span_tags:
                        tcPr.remove(ch)
                tcW = tcPr.find("w:tcW", ns)
                if tcW is None:
                    tcW = ET.Element(f"{W}tcW")
                    tcPr.insert(0, tcW)
                tcW.attrib.clear()
                tcW.set(f"{W}w", str(pcts[idx_c]))
                tcW.set(f"{W}type", "pct")

        balanced_cnt += 1

    return balanced_cnt


def ensure_three_column_tables_and_landscape(root: ET.Element) -> int:
    """
    Ensure Landscape A4 page orientation and auto-expand any 2-column <w:tbl> into 3 columns.
    """
    fn = getattr(docx_xml_utils, "ensure_three_column_tables_and_landscape", None)
    if callable(fn):
        return int(fn(root))

    ns = docx_xml_utils.NS
    W = docx_xml_utils.W
    body = root.find("w:body", ns)
    if body is None:
        return 0

    # 1. Ensure Landscape A4 on all sectPr elements (and body trailing sectPr)
    sect_list = root.findall(".//w:sectPr", ns)
    if not sect_list:
        sect_el = ET.SubElement(body, f"{W}sectPr")
        sect_list = [sect_el]
    for sect in sect_list:
        pgSz = sect.find("w:pgSz", ns)
        if pgSz is None:
            pgSz = ET.Element(f"{W}pgSz")
            sect.insert(0, pgSz)
        try:
            cur_w = int(pgSz.attrib.get(f"{W}w", "0"))
            cur_h = int(pgSz.attrib.get(f"{W}h", "0"))
        except ValueError:
            cur_w, cur_h = 0, 0
        if cur_w <= 0 or cur_h <= 0 or cur_w < cur_h:
            pgSz.set(f"{W}w", "16838")
            pgSz.set(f"{W}h", "11906")
            pgSz.set(f"{W}orient", "landscape")

    # 2. Strip stray gridSpan/gridBefore/wBefore and expand 2-cell tables to 3 cells
    span_tags = {"gridSpan", "gridBefore", "gridAfter", "wBefore", "wAfter"}
    expanded = 0
    for tbl in body.findall("w:tbl", ns):
        trs = tbl.findall("w:tr", ns)
        if not trs:
            continue
        for tr in trs:
            tcs = tr.findall("w:tc", ns)
            if len(tcs) in (2, 3):
                trPr = tr.find("w:trPr", ns)
                if trPr is not None:
                    for ch in list(trPr):
                        if ch.tag.split("}")[-1] in span_tags:
                            trPr.remove(ch)
                for tc in tcs:
                    tcPr = tc.find("w:tcPr", ns)
                    if tcPr is not None:
                        for ch in list(tcPr):
                            if ch.tag.split("}")[-1] in span_tags:
                                tcPr.remove(ch)

        has_2col = any(len(tr.findall("w:tc", ns)) == 2 for tr in trs)
        if not has_2col:
            continue

        tblGrid = tbl.find("w:tblGrid", ns)
        if tblGrid is None:
            tblGrid = ET.Element(f"{W}tblGrid")
            tblPr = tbl.find("w:tblPr", ns)
            ins_pos = 1 if (tblPr is not None and len(tbl) > 0 and tbl[0] is tblPr) else 0
            tbl.insert(ins_pos, tblGrid)
        else:
            for ch in list(tblGrid):
                tblGrid.remove(ch)
        for col_w in ("3057", "5533", "5970"):
            ET.SubElement(tblGrid, f"{W}gridCol", {f"{W}w": col_w})

        for tr in trs:
            tcs = tr.findall("w:tc", ns)
            if len(tcs) == 2:
                donor_tcPr = tcs[1].find("w:tcPr", ns)
                new_tc = ET.Element(f"{W}tc")
                new_tcPr = ET.SubElement(new_tc, f"{W}tcPr")
                ET.SubElement(new_tcPr, f"{W}tcW", {f"{W}w": "5970", f"{W}type": "dxa"})
                if donor_tcPr is not None:
                    bdr = donor_tcPr.find("w:tcBorders", ns)
                    if bdr is not None:
                        new_tcPr.append(copy.deepcopy(bdr))
                ET.SubElement(new_tc, f"{W}p")
                tr.append(new_tc)
        expanded += 1

    return expanded


# ---------------------------------------------------------------------------
# Word Counter & Standardized Naming Functions
# ---------------------------------------------------------------------------

def parse_doc_filename(
    stem: str,
    known_codes: Optional[set[str]] = None,
    strip_draft_prefix: bool = True,
) -> dict:
    """
    Parse a .docx filename stem into structured components:
      - existing_lang: e.g. 'GR' if prefixed with 'GR_', else ''
      - orig_name: clean original document title (without lang prefix, word count tag, or trailing date)
      - existing_date: 8-digit YYYYMMDD if present at the end, else ''
      - had_word_count: True if stem already contained an UP/RV word count block
    """
    s = (stem or "").strip()
    if strip_draft_prefix and s.startswith("初筛_"):
        s = s[len("初筛_") :].strip()

    codes_upper = {c.upper() for c in (known_codes or set())}

    # 1. Extract trailing date before or after word count tag
    existing_date = ""
    m_date = _TRAILING_DATE_RE.search(s)
    if m_date:
        raw_dt = m_date.group(1)
        if len(raw_dt) == 8:
            existing_date = raw_dt
        s = s[: m_date.start()].strip(" _")

    # 2. Detect & strip all word-count parenthetical tags
    had_word_count = bool(_WORD_COUNT_TAG_RE.search(s))
    if had_word_count:
        s = _WORD_COUNT_TAG_RE.sub("", s).strip(" _")

    # Check again if a trailing date was sitting before the word count tag
    m_date2 = _TRAILING_DATE_RE.search(s)
    if m_date2:
        raw_dt2 = m_date2.group(1)
        if len(raw_dt2) == 8 and not existing_date:
            existing_date = raw_dt2
        s = s[: m_date2.start()].strip(" _")

    # 3. Extract leading language code (e.g. 'GR_God’s Words')
    existing_lang = ""
    m_lang = re.match(r"^([A-Za-z]{2,5})_(.+)$", s)
    if m_lang:
        cand_code = m_lang.group(1)
        rest_name = m_lang.group(2).strip()
        if rest_name and (cand_code.upper() in codes_upper or re.fullmatch(r"[A-Z]{2,3}", cand_code)):
            existing_lang = cand_code.upper()
            s = rest_name

    orig_name = s.strip() or (stem or "Document").strip()
    return {
        "existing_lang": existing_lang,
        "orig_name": orig_name,
        "existing_date": existing_date,
        "had_word_count": had_word_count,
    }


def detect_lang_from_text(text: str, lang_entries: list[dict]) -> str:
    t_low = (text or "").lower()
    best_code = ""
    best_len = 0
    for entry in lang_entries:
        code = entry.get("code", "")
        for kw in entry.get("keywords", []):
            k_low = kw.lower().strip()
            if k_low and k_low in t_low and len(k_low) > best_len:
                best_code = code
                best_len = len(k_low)
    return best_code


def format_counted_docx_name(
    lang_code: str,
    orig_name: str,
    up_words: int,
    rv_words: int,
    date_yyyymmdd: str,
) -> str:
    """
    Format standard output filename (omitting zero word counts):
      - up > 0, rv > 0 -> {LANG}_{ORIG_NAME}_(UP {up} words, RV {rv} words)_{YYYYMMDD}.docx
      - up > 0, rv == 0 -> {LANG}_{ORIG_NAME}_(UP {up} words)_{YYYYMMDD}.docx
      - up == 0, rv > 0 -> {LANG}_{ORIG_NAME}_(RV {rv} words)_{YYYYMMDD}.docx
      - up == 0, rv == 0 -> {LANG}_{ORIG_NAME}_{YYYYMMDD}.docx
    """
    safe_lang = docx_xml_utils.sanitize_filename_part((lang_code or "XX").strip().upper(), fallback="XX", max_len=16)
    safe_orig = docx_xml_utils.sanitize_filename_part((orig_name or "Document").strip(), fallback="Document", max_len=140)
    today_str = datetime.now().strftime("%Y%m%d")
    safe_date = docx_xml_utils.sanitize_filename_part((date_yyyymmdd or today_str).strip(), fallback=today_str, max_len=16)
    count_tag = format_up_rv_count_tag(up_words, rv_words)
    if count_tag:
        return f"{safe_lang}_{safe_orig}_{count_tag}_{safe_date}.docx"
    return f"{safe_lang}_{safe_orig}_{safe_date}.docx"


def resolve_doc_naming_label(
    naming_mode: str,
    fallback_orig_name: str,
    scattered_keyword: str = "",
    doc_root: Optional[ET.Element] = None,
    docx_path: Optional[Path] = None,
) -> str:
    """
    Compute document label according to `naming_mode` ('keep' | 'full' | 'scattered'):
      - 'full': Row 0 Col 2 English title from `doc_root` / `docx_path` (fallback to `fallback_orig_name`)
      - 'scattered': f"God's words in {kw}" if `scattered_keyword.strip()` else "God's words"
      - 'keep': `fallback_orig_name`
    """
    mode = (naming_mode or "keep").strip().lower()
    if mode == "scattered":
        kw = (scattered_keyword or "").strip()
        return f"God's words in {kw}" if kw else "God's words"
    if mode == "full":
        root_el = doc_root
        if root_el is None and docx_path is not None and Path(docx_path).is_file():
            try:
                root_el, _, _ = docx_xml_utils.load_docx_xml(Path(docx_path))
            except Exception:
                root_el = None
        if root_el is not None:
            en_title = extract_full_article_en_title(root_el)
            if en_title:
                return en_title
        return fallback_orig_name
    return fallback_orig_name


def analyze_docx_word_count(
    docx_path: Path,
    default_lang_code: str = "AUTO",
    keep_existing_lang: bool = True,
    date_yyyymmdd: str = "",
    keep_existing_date: bool = False,
    strip_draft_prefix: bool = True,
    lang_entries: Optional[list[dict]] = None,
    naming_mode: str = "keep",
    scattered_keyword: str = "",
    update_header_lines: bool = False,
) -> dict:
    """
    Count Col 2 English words (UP=yellow, RV=cyan; punctuation excluded) in `docx_path`
    and compute its standardized target filename under `naming_mode` ('keep' | 'full' | 'scattered').
    """
    p = Path(docx_path).resolve()
    if not p.is_file():
        raise FileNotFoundError(f"找不到文件：{p}")

    entries = lang_entries if lang_entries is not None else load_languages()
    known_codes = {e["code"].upper() for e in entries if e.get("code")}

    cfg = docx_xml_utils.load_config()
    col2_up = cfg.get("col2_up_highlight", "yellow")
    col2_rv = cfg.get("col2_rv_highlight", "cyan")

    root, _, _ = docx_xml_utils.load_docx_xml(p)
    up_words, rv_words = docx_xml_utils.compute_col2_up_rv_words(root, up_color=col2_up, rv_color=col2_rv)

    parsed = parse_doc_filename(p.stem, known_codes=known_codes, strip_draft_prefix=strip_draft_prefix)

    # Resolve target language code
    sel_code = (default_lang_code or "AUTO").strip().upper()
    if keep_existing_lang and parsed["existing_lang"]:
        final_lang = parsed["existing_lang"]
    elif sel_code and sel_code != "AUTO":
        final_lang = sel_code
    else:
        det = detect_lang_from_text(f"{p.parent.name} {p.stem}", entries)
        final_lang = det or parsed["existing_lang"] or "XX"

    # Resolve target date (YYYYMMDD)
    today_str = datetime.now().strftime("%Y%m%d")
    if keep_existing_date and parsed["existing_date"]:
        final_date = parsed["existing_date"]
    else:
        final_date = (date_yyyymmdd or "").strip() or today_str

    target_orig = resolve_doc_naming_label(
        naming_mode=naming_mode,
        fallback_orig_name=parsed["orig_name"],
        scattered_keyword=scattered_keyword,
        doc_root=root,
    )

    new_filename = format_counted_docx_name(
        lang_code=final_lang,
        orig_name=target_orig,
        up_words=up_words,
        rv_words=rv_words,
        date_yyyymmdd=final_date,
    )

    return {
        "path": str(p),
        "old_name": p.name,
        "new_name": new_filename,
        "lang_code": final_lang,
        "orig_name": target_orig,
        "raw_orig_name": parsed["orig_name"],
        "naming_mode": (naming_mode or "keep").strip().lower(),
        "update_header_lines": bool(update_header_lines),
        "date_str": final_date,
        "up_words": up_words,
        "rv_words": rv_words,
        "total_words": up_words + rv_words,
        "had_word_count": parsed["had_word_count"],
    }


def apply_word_count_filename(
    docx_path: Path,
    new_filename: str,
    copy_to_dir: Optional[Path] = None,
    update_header_lines: bool = False,
) -> tuple[Path, str]:
    """
    Rename `docx_path` in place (or copy to `copy_to_dir`) with `new_filename`.
    When `update_header_lines=True`, also updates the top 'Blue/Yellow highlighted parts...'
    instruction lines inside the saved `.docx`.
    Returns (final_path, action_description).
    """
    src = Path(docx_path).resolve()
    if not src.is_file():
        raise FileNotFoundError(f"原文档不存在：{src}")

    if copy_to_dir is not None:
        out_dir = Path(copy_to_dir).resolve()
        out_dir.mkdir(parents=True, exist_ok=True)
        dst = docx_xml_utils.safe_join_output_file(out_dir, new_filename)
        if update_header_lines:
            cfg = docx_xml_utils.load_config()
            col2_up = cfg.get("col2_up_highlight", "yellow")
            col2_rv = cfg.get("col2_rv_highlight", "cyan")
            root, _, ns_pairs = docx_xml_utils.load_docx_xml(src)
            comments_root = docx_xml_utils.load_docx_comments_xml(src)
            up_w, rv_w = docx_xml_utils.compute_col2_up_rv_words(root, up_color=col2_up, rv_color=col2_rv)
            update_top_highlight_instruction_lines(root, has_up=(up_w > 0), has_rv=(rv_w > 0))
            docx_xml_utils.save_docx_xml(src, dst, root, ns_pairs, comments_root=comments_root)
            return dst, "✅ 已更新说明行并另存副本"
        shutil.copy2(src, dst)
        return dst, "✅ 已另存副本"

    dst = docx_xml_utils.safe_join_output_file(src.parent, new_filename)
    same_name = src.name == dst.name

    # Check collision with a DIFFERENT file in the same directory
    if not same_name and dst.exists() and os.path.normcase(str(src)) != os.path.normcase(str(dst)):
        raise FileExistsError(f"同目录下已存在同名文件，为防覆盖已跳过：{dst.name}")

    if update_header_lines:
        cfg = docx_xml_utils.load_config()
        col2_up = cfg.get("col2_up_highlight", "yellow")
        col2_rv = cfg.get("col2_rv_highlight", "cyan")
        root, _, ns_pairs = docx_xml_utils.load_docx_xml(src)
        comments_root = docx_xml_utils.load_docx_comments_xml(src)
        up_w, rv_w = docx_xml_utils.compute_col2_up_rv_words(root, up_color=col2_up, rv_color=col2_rv)
        update_top_highlight_instruction_lines(root, has_up=(up_w > 0), has_rv=(rv_w > 0))
        try:
            docx_xml_utils.save_docx_xml(src, src, root, ns_pairs, comments_root=comments_root)
            if not same_name:
                src.rename(dst)
        except PermissionError as pe:
            raise PermissionError(f"文件正在被 Word 或其他程序占用，请先关闭文档后重试：{src.name}") from pe
        if same_name:
            return src, "✅ 已更新开头说明行（文件名已是最新）"
        return dst, "✅ 已更新说明行与文件名"

    if same_name:
        return src, "✅ 文件名已是最新"

    try:
        src.rename(dst)
    except PermissionError as pe:
        raise PermissionError(f"文件正在被 Word 或其他程序占用，请先关闭文档后重试：{src.name}") from pe

    return dst, "✅ 已更新文件名"
