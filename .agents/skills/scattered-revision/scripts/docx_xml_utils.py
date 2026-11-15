import copy
import difflib
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path

W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_URI = "http://www.w3.org/XML/1998/namespace"
W = f"{{{W_URI}}}"
XML_SPACE = f"{{{XML_URI}}}space"
NS = {"w": W_URI}

# Strict ECMA-376 CT_RPr child tag order
RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
    "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof",
    "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
    "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
    "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
    "specVanish", "oMath", "rPrChange"
]
RPR_RANK = {name: idx for idx, name in enumerate(RPR_ORDER)}


DEFAULT_CONFIG = {
    "default_author": "整理组",
    "col2_up_highlight": "yellow",
    "col2_rv_highlight": "cyan",
    "col3_up_highlight": "yellow",
    "col3_rv_highlight": "lightGray",
    "accept_col2_rv_revisions": False,
    "split_col3_rv_paragraph": False,
    "rv_word_change_ratio_threshold": 0.42,
    "header_paragraphs": [],
    "ensure_top_blank_lines": 2,
    "language_code_map": {
        "法语": "FR", "中英法": "FR",
        "葡萄牙语": "PT", "葡语": "PT", "中英葡": "PT",
        "西班牙语": "ES", "西语": "ES", "中英西": "ES",
        "德语": "DE", "中英德": "DE",
        "意大利语": "IT", "意语": "IT", "中英意": "IT",
        "希腊语": "GR", "腊语": "GR", "中英腊": "GR",
        "俄语": "RU", "中英俄": "RU",
        "乌克兰语": "UK", "乌语": "UK", "中英乌": "UK",
        "韩语": "KO", "朝鲜语": "KO", "中英韩": "KO",
        "日语": "JA", "日本语": "JA", "中英日": "JA",
        "泰语": "TH", "中英泰": "TH",
        "越南语": "VI", "越语": "VI", "中英越": "VI",
        "印度尼西亚语": "ID", "印尼语": "ID", "中英印尼": "ID",
        "马来语": "MS", "中英马来": "MS",
        "塔加洛语": "TL", "他加禄语": "TL", "菲律宾语": "TL", "中英塔": "TL",
        "缅甸语": "MY", "缅语": "MY", "中英缅": "MY",
        "高棉语": "KM", "柬埔寨语": "KM", "中英柬": "KM",
        "老挝语": "LO", "中英老": "LO",
        "阿拉伯语": "AR", "阿语": "AR", "中英阿": "AR",
        "土耳其语": "TR", "土语": "TR", "中英土": "TR",
        "波斯语": "FA", "中英波斯": "FA",
        "希伯来语": "HE", "印地语": "HI", "乌尔都语": "UR", "孟加拉语": "BN",
        "波兰语": "PL", "中英波": "PL",
        "荷兰语": "NL", "中英荷": "NL",
        "罗马尼亚语": "RO", "中英罗": "RO",
        "捷克语": "CS", "匈牙利语": "HU", "保加利亚语": "BG", "瑞典语": "SV",
        "蒙古语": "MN", "斯瓦希里语": "SW", "阿姆哈拉语": "AM",
    },
    "exclude_library_dir_keywords": ["标记结果", "初筛", "答案", "output", "temp"],
}


def load_config():
    import os
    import sys
    candidates = [
        Path(__file__).resolve().parent.parent / "resources" / "config.json",
        Path(__file__).resolve().parent / "resources" / "config.json",
    ]
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        candidates.insert(0, Path(sys._MEIPASS) / "resources" / "config.json")
        candidates.insert(1, Path(sys._MEIPASS) / "skill_scripts" / "resources" / "config.json")
    if getattr(sys, "frozen", False):
        candidates.insert(0, Path(sys.executable).resolve().parent / "config.json")

    cfg = None
    for cfg_path in candidates:
        if cfg_path.exists():
            with open(cfg_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            break
    if cfg is None:
        cfg = copy.deepcopy(DEFAULT_CONFIG)

    # Merge user-extensible languages.json if present
    lang_json_candidates = []
    env_lp = os.environ.get("SCATTERED_REV_LANGUAGES_JSON", "").strip()
    if env_lp:
        lang_json_candidates.append(Path(env_lp))
    if getattr(sys, "frozen", False):
        lang_json_candidates.append(Path(sys.executable).resolve().parent / "languages.json")
    lang_json_candidates.append(Path(__file__).resolve().parent.parent.parent.parent.parent / "零散改点工具" / "languages.json")
    lang_json_candidates.append(Path(__file__).resolve().parent.parent / "resources" / "languages.json")

    raw_pairs = list(cfg.get("language_code_map", {}).items())
    for lp in lang_json_candidates:
        if lp.exists():
            try:
                with open(lp, "r", encoding="utf-8") as f:
                    ldata = json.load(f)
                entries = ldata.get("languages", ldata) if isinstance(ldata, dict) else ldata
                if isinstance(entries, list):
                    custom_pairs = []
                    for item in entries:
                        code = str(item.get("code", "")).strip().upper()
                        name = str(item.get("name", "")).strip()
                        kws = item.get("keywords") or []
                        if not code:
                            continue
                        if name and name not in kws:
                            kws = [name] + list(kws)
                        for kw in kws:
                            kw_s = str(kw).strip()
                            if kw_s:
                                custom_pairs.append((kw_s, code))
                    # User custom entries override built-in ones if same keyword
                    pair_dict = dict(raw_pairs)
                    for k, v in custom_pairs:
                        pair_dict[k] = v
                    raw_pairs = list(pair_dict.items())
                break
            except Exception:
                pass

    # Sort longest keyword first to prevent short substrings shadowing specific language names
    raw_pairs.sort(key=lambda kv: len(kv[0]), reverse=True)
    cfg["language_code_map"] = dict(raw_pairs)
    return cfg


def register_xml_namespaces(xml_bytes: bytes):
    head = xml_bytes[:8192].decode("utf-8", errors="ignore")
    m = re.search(r"<w:document\b[^>]*>", head)
    open_tag = m.group(0) if m else ""
    ns_pairs = re.findall(r'xmlns:([A-Za-z0-9_.-]+)="([^"]+)"', open_tag)
    for prefix, uri in ns_pairs:
        ET.register_namespace(prefix, uri)
    ET.register_namespace("w", W_URI)
    return open_tag, ns_pairs


MAX_DOCX_MEMBER_BYTES = 100 * 1024 * 1024   # 100 MB single XML/media member limit
MAX_DOCX_TOTAL_BYTES = 300 * 1024 * 1024    # 300 MB total uncompressed docx archive limit
_DTD_RE = re.compile(rb"<!\s*(?:DOCTYPE|ENTITY)\b", re.IGNORECASE)
_WIN_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def sanitize_filename_part(value: str, fallback: str = "doc", max_len: int = 120) -> str:
    """Sanitize a user/config string for safe single-filename usage (prevents CWE-22 path traversal)."""
    s = re.sub(r'[\x00-\x1f/\\:*?"<>|]+', "_", (value or "").strip())
    s = s.strip(" .")
    while ".." in s:
        s = s.replace("..", "_")
    if not s:
        s = fallback
    stem_upper = s.split(".")[0].upper()
    if stem_upper in _WIN_RESERVED_NAMES:
        s = f"_{s}"
    return s[:max_len]


def safe_join_output_file(outdir: Path, filename: str) -> Path:
    """Resolve output file inside outdir and strictly block directory traversal outside outdir."""
    resolved_dir = outdir.resolve()
    clean_name = Path(filename).name
    if clean_name != filename or ".." in Path(filename).parts or "/" in filename or "\\" in filename:
        raise ValueError(f"Unsafe output filename rejected: {filename!r}")
    candidate = (resolved_dir / clean_name).resolve()
    if candidate.parent != resolved_dir:
        raise ValueError(f"Path traversal blocked: {candidate} escapes {resolved_dir}")
    return candidate


def _validate_docx_zip(zf: zipfile.ZipFile) -> None:
    """Validate .docx archive against Zip Slip path traversal and Zip Bomb decompression DoS."""
    total_uncompressed = 0
    for info in zf.infolist():
        fname = info.filename
        if "\x00" in fname or fname.startswith(("/", "\\")) or ":" in fname:
            raise ValueError(f"Unsafe ZIP entry path in .docx: {fname!r}")
        norm_parts = fname.replace("\\", "/").split("/")
        if ".." in norm_parts:
            raise ValueError(f"Zip Slip path traversal detected in .docx entry: {fname!r}")
        if info.file_size < 0 or info.file_size > MAX_DOCX_MEMBER_BYTES:
            raise ValueError(
                f"ZIP member {fname!r} uncompressed size ({info.file_size} bytes) exceeds safety limit."
            )
        total_uncompressed += info.file_size
        if total_uncompressed > MAX_DOCX_TOTAL_BYTES:
            raise ValueError(
                f"Total uncompressed .docx size ({total_uncompressed} bytes) exceeds safety limit."
            )


def _assert_safe_openxml_bytes(xml_bytes: bytes) -> None:
    """Block DTD / Entity declarations (Billion Laughs / XXE) before XML parsing."""
    if len(xml_bytes) > MAX_DOCX_MEMBER_BYTES:
        raise ValueError("word/document.xml exceeds maximum allowed byte size.")
    if _DTD_RE.search(xml_bytes):
        raise ValueError("Unsafe XML DOCTYPE/ENTITY declaration rejected in OpenXML document.")


def load_docx_xml(docx_path: Path):
    with zipfile.ZipFile(docx_path, "r") as z:
        _validate_docx_zip(z)
        info = z.getinfo("word/document.xml")
        with z.open(info, "r") as fp:
            xml_bytes = fp.read(MAX_DOCX_MEMBER_BYTES + 1)
    _assert_safe_openxml_bytes(xml_bytes)
    open_tag, ns_pairs = register_xml_namespaces(xml_bytes)
    root = ET.fromstring(xml_bytes)
    return root, open_tag, ns_pairs


def save_docx_xml(src_docx_path: Path, dst_docx_path: Path, root: ET.Element, ns_pairs):
    register_xml_namespaces(b"")
    for prefix, uri in ns_pairs:
        ET.register_namespace(prefix, uri)
    ET.register_namespace("w", W_URI)

    xml_str = ET.tostring(root, encoding="unicode")
    # Re-inject any namespace declarations dropped by ElementTree so mc:Ignorable stays valid in Word
    m = re.match(r"<w:document\b[^>]*>", xml_str)
    if m:
        cur_open = m.group(0)
        missing = []
        for prefix, uri in ns_pairs:
            decl = f'xmlns:{prefix}="{uri}"'
            if f"xmlns:{prefix}=" not in cur_open:
                missing.append(decl)
        if missing:
            patched_open = cur_open[:-1] + " " + " ".join(missing) + ">"
            xml_str = patched_open + xml_str[m.end():]

    full_bytes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml_str).encode("utf-8")
    dst_docx_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src_docx_path, "r") as zin:
        _validate_docx_zip(zin)
        with zipfile.ZipFile(dst_docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, full_bytes)
                else:
                    with zin.open(item, "r") as mfp:
                        data = mfp.read(MAX_DOCX_MEMBER_BYTES + 1)
                    if len(data) > MAX_DOCX_MEMBER_BYTES:
                        raise ValueError(f"Member {item.filename!r} exceeded actual read byte limit.")
                    zout.writestr(item, data)


def insert_rpr_child_ordered(rPr: ET.Element, new_child: ET.Element):
    new_local = new_child.tag.split("}")[-1]
    for existing in list(rPr):
        if existing.tag.split("}")[-1] == new_local:
            rPr.remove(existing)
    target_rank = RPR_RANK.get(new_local, 999)
    for idx, child in enumerate(list(rPr)):
        c_local = child.tag.split("}")[-1]
        if RPR_RANK.get(c_local, 999) > target_rank:
            rPr.insert(idx, new_child)
            return
    rPr.append(new_child)


def remove_rpr_child(rPr: ET.Element, local_name: str):
    if rPr is None:
        return
    for existing in list(rPr):
        if existing.tag.split("}")[-1] == local_name:
            rPr.remove(existing)


def build_rpr(base_rPr: ET.Element = None, highlight: str = None, shd_fill: str = None) -> ET.Element:
    rPr = copy.deepcopy(base_rPr) if base_rPr is not None else ET.Element(f"{W}rPr")
    if highlight:
        hl_el = ET.Element(f"{W}highlight", {f"{W}val": highlight})
        insert_rpr_child_ordered(rPr, hl_el)
    else:
        remove_rpr_child(rPr, "highlight")

    if shd_fill:
        shd_el = ET.Element(f"{W}shd", {f"{W}val": "clear", f"{W}color": "auto", f"{W}fill": shd_fill})
        insert_rpr_child_ordered(rPr, shd_el)
    else:
        remove_rpr_child(rPr, "shd")
    return rPr


def make_run(text: str, base_rPr: ET.Element = None, highlight: str = None, shd_fill: str = None, is_del: bool = False) -> ET.Element:
    r = ET.Element(f"{W}r")
    rPr = build_rpr(base_rPr, highlight=highlight, shd_fill=shd_fill)
    if len(rPr) > 0:
        r.append(rPr)
    t_tag = f"{W}delText" if is_del else f"{W}t"
    t = ET.Element(t_tag)
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        t.set(XML_SPACE, "preserve")
    t.text = text
    r.append(t)
    return r


def make_track_el(kind: str, rev_id: int, author: str, date_iso: str, run_el: ET.Element) -> ET.Element:
    el = ET.Element(f"{W}{kind}", {
        f"{W}id": str(rev_id),
        f"{W}author": author,
        f"{W}date": date_iso,
    })
    el.append(run_el)
    return el


def get_cell_plain_text(tc: ET.Element) -> str:
    return "\n".join(
        "".join(t.text or "" for t in p.findall(".//w:t", NS))
        for p in tc.findall("w:p", NS)
    )


def get_first_run_rpr(tc: ET.Element, fallback_tc: ET.Element = None) -> ET.Element:
    """
    Extract the representative clean <w:rPr> for a table cell, preserving the input document's
    exact font family, font size, color, etc.
    Prefers the dominant (longest text) non-whitespace run in `tc` so a stray 1-char kerning/punctuation
    run never hijacks cell formatting; falls back to `tc`'s <w:pPr>/<w:rPr>, then `fallback_tc`.
    """
    best_rPr = None
    best_len = -1
    for r in tc.findall(".//w:r", NS):
        rPr = r.find("w:rPr", NS)
        if rPr is None:
            continue
        txt = "".join(t.text or "" for t in r.findall("w:t", NS)).strip()
        if len(txt) > best_len:
            best_len = len(txt)
            best_rPr = rPr

    pPr_rPr = tc.find("./w:p/w:pPr/w:rPr", NS)
    fb_rPr = None
    if fallback_tc is not None:
        fb_rPr = get_first_run_rpr(fallback_tc, fallback_tc=None)

    if best_rPr is not None:
        cp = copy.deepcopy(best_rPr)
    elif pPr_rPr is not None:
        cp = copy.deepcopy(pPr_rPr)
    elif fb_rPr is not None:
        cp = copy.deepcopy(fb_rPr)
    else:
        cp = ET.Element(f"{W}rPr")
        ET.SubElement(cp, f"{W}rFonts", {
            f"{W}ascii": "Georgia", f"{W}eastAsia": "宋体",
            f"{W}hAnsi": "Georgia", f"{W}cs": "Georgia",
        })
        ET.SubElement(cp, f"{W}sz", {f"{W}val": "24"})
        ET.SubElement(cp, f"{W}szCs", {f"{W}val": "24"})

    remove_rpr_child(cp, "highlight")
    remove_rpr_child(cp, "shd")

    # Fill in missing rFonts / sz / szCs from paragraph rPr or fallback cell rPr so runs never fall back to Normal style
    for donor in (pPr_rPr, fb_rPr):
        if donor is None:
            continue
        for tag_local in ("rFonts", "sz", "szCs"):
            if cp.find(f"w:{tag_local}", NS) is None:
                d_el = donor.find(f"w:{tag_local}", NS)
                if d_el is not None:
                    insert_rpr_child_ordered(cp, copy.deepcopy(d_el))

    return cp


def get_cell_template_ppr(tc: ET.Element, fallback_tc: ET.Element = None) -> ET.Element:
    """Extract clean template <w:pPr> from target cell (or fallback cell) to preserve original spacing/indent/font."""
    p = tc.find("w:p", NS)
    if p is not None and p.find("w:pPr", NS) is not None:
        cp = copy.deepcopy(p.find("w:pPr", NS))
    elif fallback_tc is not None:
        return get_cell_template_ppr(fallback_tc, fallback_tc=None)
    else:
        return None
    p_rPr = cp.find("w:rPr", NS)
    if p_rPr is not None:
        remove_rpr_child(p_rPr, "highlight")
        remove_rpr_child(p_rPr, "shd")
    return cp


def extract_char_shading(tc: ET.Element):
    """Return (full_text, list_of_fill_hex_or_None_per_char)."""
    chars = []
    fills = []
    for p_i, p in enumerate(tc.findall("w:p", NS)):
        if p_i > 0:
            chars.append("\n")
            fills.append(None)
        for r in p.findall("w:r", NS):
            txt = "".join(x.text or "" for x in r.findall("w:t", NS))
            if not txt:
                continue
            shd = r.find("./w:rPr/w:shd", NS)
            fill = shd.attrib.get(f"{W}fill") if shd is not None else None
            if fill in (None, "auto", "FFFFFF", "ffffff"):
                fill = None
            for ch in txt:
                chars.append(ch)
                fills.append(fill)
    return "".join(chars), fills


def split_into_bg_regions(src_text: str, inp_text: str, inp_char_fills: list):
    """
    Align old library English (src_text) with new input English (inp_text)
    and partition into [(fill_hex_or_None, src_segment, inp_segment), ...].
    """
    inp_matches = list(re.finditer(r"\S+", inp_text))
    src_matches = list(re.finditer(r"\S+", src_text))
    inp_words = [m.group() for m in inp_matches]
    src_words = [m.group() for m in src_matches]
    inp_w_fills = [inp_char_fills[m.start()] for m in inp_matches]

    if not inp_w_fills:
        return [(None, src_text, inp_text)]
    if all(f is None for f in inp_w_fills):
        return [(None, src_text, inp_text)]
    if all(f is not None for f in inp_w_fills):
        dom_fill = next(f for f in inp_w_fills if f is not None)
        return [(dom_fill, src_text, inp_text)]

    sm = difflib.SequenceMatcher(None, src_words, inp_words, autojunk=False)
    opcodes = sm.get_opcodes()

    def map_j_to_i(j: int) -> int:
        if j <= 0:
            return 0
        if j >= len(inp_words):
            return len(src_words)
        for tag, i1, i2, j1, j2 in opcodes:
            if j1 <= j <= j2:
                if tag == "equal":
                    return i1 + (j - j1)
                if j == j1:
                    return i1
                if j == j2:
                    return i2
                # Shading boundaries occur at sentence boundaries: align sentence-ending words inside replace blocks
                def _ends_sent(w: str) -> bool:
                    s = w.rstrip("”’\"')")
                    return bool(s) and s[-1] in ".?!"

                inp_sent_ends = sum(1 for w in inp_words[j1:j] if _ends_sent(w))
                src_end_positions = [idx_s + 1 for idx_s in range(i1, i2) if _ends_sent(src_words[idx_s])]
                if inp_sent_ends > 0 and len(src_end_positions) >= inp_sent_ends:
                    return src_end_positions[inp_sent_ends - 1]
                if inp_sent_ends == 0 and src_end_positions and _ends_sent(inp_words[j - 1]):
                    return src_end_positions[0]
                ratio = (j - j1) / max(1, (j2 - j1))
                return i1 + round(ratio * (i2 - i1))
        return len(src_words)

    runs = []
    cur_f = inp_w_fills[0]
    start_j = 0
    for j in range(1, len(inp_w_fills)):
        if (inp_w_fills[j] is not None) != (cur_f is not None):
            runs.append((cur_f, start_j, j))
            cur_f = inp_w_fills[j]
            start_j = j
    runs.append((cur_f, start_j, len(inp_w_fills)))

    regions = []
    for idx, (fill_val, j1, j2) in enumerate(runs):
        i1 = map_j_to_i(j1)
        i2 = map_j_to_i(j2)
        s_c1 = 0 if idx == 0 else (src_matches[i1].start() if i1 < len(src_matches) else len(src_text))
        s_c2 = len(src_text) if idx == len(runs) - 1 else (src_matches[i2].start() if i2 < len(src_matches) else len(src_text))
        t_c1 = 0 if idx == 0 else (inp_matches[j1].start() if j1 < len(inp_matches) else len(inp_text))
        t_c2 = len(inp_text) if idx == len(runs) - 1 else (inp_matches[j2].start() if j2 < len(inp_matches) else len(inp_text))
        regions.append((fill_val, src_text[s_c1:s_c2], inp_text[t_c1:t_c2]))
    return regions


# Fine-grained tokenizer keeping you(pl)/you(s)/God's together while splitting punctuation & spaces
TOKEN_RE = re.compile(
    r"[A-Za-z0-9]+(?:[’'][A-Za-z0-9]+)*(?:-[A-Za-z0-9]+)*(?:\([A-Za-z]+\))?"
    r"|\s+"
    r"|[^\w\s]"
)


def tokenize_en_fine(text: str) -> list:
    return TOKEN_RE.findall(text)


def is_word_token(tok: str) -> bool:
    return bool(re.search(r"[A-Za-z0-9]", tok))


PRON_TAG_RE = re.compile(r"^(you|your|yours)(?:\((s|pl)\))?$", re.IGNORECASE)


def is_non_revision_pronoun_pair(old_tok: str, new_tok: str) -> bool:
    """
    Return True when old_tok and new_tok are the exact same pronoun (you/your/yours, same case)
    and differ ONLY by adding or removing (s)/(pl), e.g. 'you' <-> 'you(pl)', 'Your(s)' <-> 'Your'.
    Switching singular <-> plural ('You(s)' <-> 'you(pl)', 'you(s)' <-> 'you(pl)') returns False.
    """
    m1 = PRON_TAG_RE.match(old_tok)
    m2 = PRON_TAG_RE.match(new_tok)
    if not m1 or not m2:
        return False
    stem1, tag1 = m1.group(1), (m1.group(2) or "").lower()
    stem2, tag2 = m2.group(1), (m2.group(2) or "").lower()
    return stem1 == stem2 and ((tag1 == "") != (tag2 == ""))


def diff_tokens(old_text: str, new_text: str) -> list:
    """
    Produce fine-grained diff operations [('eq'|'del'|'ins', text), ...]
    auto-accepting pure (s)/(pl) tag additions/removals on you/your/yours into 'eq'.
    """
    a = tokenize_en_fine(old_text)
    b = tokenize_en_fine(new_text)
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    ops = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for tok in a[i1:i2]:
                ops.append(("eq", tok))
        elif tag == "delete":
            ops.append(("del", "".join(a[i1:i2])))
        elif tag == "insert":
            ops.append(("ins", "".join(b[j1:j2])))
        elif tag == "replace":
            # Peel leading non-revision pronoun pairs (e.g. 'you' <-> 'you(pl)') directly as accepted 'eq'
            while i1 < i2 and j1 < j2 and is_non_revision_pronoun_pair(a[i1], b[j1]):
                ops.append(("eq", b[j1]))
                i1 += 1
                j1 += 1
                if i1 < i2 and j1 < j2 and a[i1] == b[j1] and a[i1].isspace():
                    ops.append(("eq", b[j1]))
                    i1 += 1
                    j1 += 1
            # Peel trailing non-revision pronoun pairs
            trailing_eq = []
            while i1 < i2 and j1 < j2 and is_non_revision_pronoun_pair(a[i2 - 1], b[j2 - 1]):
                trailing_eq.append(("eq", b[j2 - 1]))
                i2 -= 1
                j2 -= 1
                if i1 < i2 and j1 < j2 and a[i2 - 1] == b[j2 - 1] and a[i2 - 1].isspace():
                    trailing_eq.append(("eq", b[j2 - 1]))
                    i2 -= 1
                    j2 -= 1
            if i1 < i2:
                ops.append(("del", "".join(a[i1:i2])))
            if j1 < j2:
                ops.append(("ins", "".join(b[j1:j2])))
            for eq_op in reversed(trailing_eq):
                ops.append(eq_op)
    return ops


SENT_END_CHARS = {".", "?", "!"}
OPEN_QUOTES = {"“", '"'}
CLOSE_QUOTES = {"”", '"'}
CLAUSE_SEPS = {",", ";", ":", ".", "?", "!", "—"}

WORD_EXTRACT_RE = re.compile(r"[A-Za-z0-9]+(?:[’'][A-Za-z0-9]+)*(?:-[A-Za-z0-9]+)*(?:\([A-Za-z]+\))?")
HARD_DELIMS = {",", ";", ":", ".", "?", "!", "—", "”"}
OPEN_BRACKETS = {"(", "（", "["}
CLOSE_BRACKETS = {")", "）", "]"}
SUBCLAUSE_CONJ = {"and", "or", "but", "yet", "while", "nor"}
PREPOSITIONS = {"of", "in", "to", "for", "with", "on", "at", "from", "by", "between", "among", "into"}


def segment_diff_into_sentences(ops: list) -> list:
    """
    Split a region's diff op stream into synchronized sentence units.
    Only splits at true sentence-ending punctuation (outside open quotes, or right after closing quote)
    where old and new text remain synchronized.
    """
    sentences = []
    cur = []
    quote_depth = 0

    i = 0
    n = len(ops)
    while i < n:
        kind, txt = ops[i]
        cur.append((kind, txt))
        if kind != "del":
            for ch in txt:
                if ch in ("“",):
                    quote_depth += 1
                elif ch in ("”",) and quote_depth > 0:
                    quote_depth -= 1

        # Check if this position ends a synchronized sentence
        split_here = False
        if kind == "eq" and txt in SENT_END_CHARS:
            # Consume immediately following closing quote if eq
            while i + 1 < n and ops[i + 1][0] == "eq" and ops[i + 1][1] in ("”", "’", '"', "'"):
                i += 1
                cur.append(ops[i])
                if ops[i][1] == "”" and quote_depth > 0:
                    quote_depth -= 1
            if quote_depth == 0:
                split_here = True
        elif kind == "ins" and len(cur) >= 2 and cur[-2][0] == "del":
            # Both del and ins end with sentence-ending punctuation (e.g. 'God.' -> 'God!')
            del_str = cur[-2][1].rstrip("”’\"' ")
            ins_str = txt.rstrip("”’\"' ")
            if del_str and ins_str and del_str[-1] in SENT_END_CHARS and ins_str[-1] in SENT_END_CHARS and quote_depth == 0:
                split_here = True

        if split_here:
            sentences.append(cur)
            cur = []
        i += 1

    if cur:
        sentences.append(cur)
    return sentences


def sentence_unit_stats(sent_ops: list) -> dict:
    has_edit = any(k in ("del", "ins") and not t.isspace() for k, t in sent_ops)
    old_txt = "".join(t for k, t in sent_ops if k in ("eq", "del"))
    new_txt = "".join(t for k, t in sent_ops if k in ("eq", "ins"))
    preview = "".join(
        t if k == "eq" else (f"{{-{t}-}}" if k == "del" else f"{{+{t}+}}")
        for k, t in sent_ops
    )
    eq_words = sum(len(re.findall(r"[A-Za-z0-9’']+", t)) for k, t in sent_ops if k == "eq")
    del_words = sum(len(re.findall(r"[A-Za-z0-9’']+", t)) for k, t in sent_ops if k == "del")
    ins_words = sum(len(re.findall(r"[A-Za-z0-9’']+", t)) for k, t in sent_ops if k == "ins")
    total_new_words = max(1, eq_words + ins_words)
    change_ratio = round(max(del_words, ins_words) / total_new_words, 3)
    edit_hunks = sum(1 for idx, (k, _) in enumerate(sent_ops) if k in ("del", "ins") and (idx == 0 or sent_ops[idx - 1][0] == "eq"))
    return {
        "has_edit": has_edit,
        "old_text": old_txt.strip(),
        "new_text": new_txt.strip(),
        "diff_preview": preview.strip(),
        "new_words": eq_words + ins_words,
        "del_words": del_words,
        "ins_words": ins_words,
        "change_ratio": change_ratio,
        "edit_hunks": edit_hunks,
    }


def _extract_words_lower(s: str) -> list:
    return [m.group(0).lower() for m in WORD_EXTRACT_RE.finditer(s)]


def _count_ops_words(ops_slice: list) -> int:
    txt = "".join(t for k, t in ops_slice if k in ("eq", "ins"))
    return count_msword_en_words(txt)


def compute_precision_yellow(ops: list) -> list:
    """
    Given diff ops [(kind, text), ...] for a contiguous non-cyan block inside a shaded region,
    return [(kind, text, 'yellow'|None), ...] implementing the 5 precision rules:
      1. Pure punctuation / case-after-punct edits -> 3-word window around edit.
      2. Long clauses (>12 words) with localized edit -> split at intact conjunction sub-clauses (>=5w each side).
      3. Tighten after unchanged/inserted ':' or ';', restrict parenthetical edits to '(...)', stop at cluster-ending delimiters.
      4. Bridge fragmented internal gaps when surrounding yellow spans outnumber unmarked internal gaps (outer ends untouched).
    """
    n = len(ops)
    if n == 0:
        return []

    hl = [False] * n

    # Step A: Identify edit clusters
    clusters = []
    i = 0
    while i < n:
        if ops[i][0] in ("del", "ins") and not ops[i][1].isspace():
            start = i
            end = i
            while True:
                if ops[end][1].strip() in (":", ";"):
                    break
                j = end + 1
                while j < n and ops[j][0] == "eq" and ops[j][1].isspace():
                    j += 1
                if (
                    j < n and ops[j][0] == "eq" and ops[j][1] == ","
                    and j + 1 < n and ops[j + 1][0] in ("del", "ins")
                ):
                    end = j + 1
                    continue
                if j < n and ops[j][0] in ("del", "ins") and not ops[j][1].isspace():
                    end = j
                    continue
                if (
                    ops[end][1].strip() in ("“", '"', "‘", "'")
                    and end + 2 < n
                    and ops[end + 1][0] == "eq"
                    and is_word_token(ops[end + 1][1])
                    and ops[end + 2][0] in ("del", "ins")
                    and ops[end + 2][1].strip() in ("”", '"', "’", "'")
                ):
                    end = end + 2
                    continue
                break
            clusters.append((start, end))
            i = end + 1
        else:
            i += 1

    if not clusters:
        return [(k, t, None) for k, t in ops]

    cluster_infos = []
    for s, e in clusters:
        del_txt = "".join(ops[k][1] for k in range(s, e + 1) if ops[k][0] == "del")
        ins_txt = "".join(ops[k][1] for k in range(s, e + 1) if ops[k][0] == "ins")
        ctype = "punct_only" if _extract_words_lower(del_txt) == _extract_words_lower(ins_txt) else "word_edit"
        cluster_infos.append({"start": s, "end": e, "type": ctype, "del_txt": del_txt, "ins_txt": ins_txt})

    # Rule 3a: Newly inserted ':'/';'/',' immediately preceding a word edit should NOT trigger a 3-word window before ':'
    for idx_c, ci in enumerate(cluster_infos):
        if ci["type"] == "punct_only" and ci["ins_txt"].strip() in (":", ";", ",") and idx_c + 1 < len(cluster_infos):
            nxt = cluster_infos[idx_c + 1]
            between = ops[ci["end"] + 1 : nxt["start"]]
            if nxt["type"] == "word_edit" and all(k == "eq" and t.isspace() for k, t in between):
                ci["type"] = "leading_delim_skip"

    # Pass 1: Expand 'word_edit' clusters to tight clause / parentheses / conjunction boundaries (Rules 2 & 3)
    for ci in cluster_infos:
        if ci["type"] != "word_edit":
            continue
        s, e = ci["start"], ci["end"]

        L = s
        paren_depth = 0
        while L > 0:
            pk, pt = ops[L - 1]
            if pk == "eq":
                if pt in CLOSE_BRACKETS:
                    paren_depth += 1
                elif pt in OPEN_BRACKETS:
                    if paren_depth > 0:
                        paren_depth -= 1
                    else:
                        L -= 1
                        break
                elif paren_depth == 0 and pt in HARD_DELIMS:
                    break
                elif paren_depth == 0 and pt in ("“",):
                    break
            elif pk in ("del", "ins"):
                if pt.strip() in (":", ";"):
                    break
            L -= 1

        R = e
        effective_new_tail = "".join(ops[k][1] for k in range(s, e + 1) if ops[k][0] in ("eq", "ins")).rstrip()
        if effective_new_tail and effective_new_tail[-1] in HARD_DELIMS:
            R = e
        else:
            paren_depth = 0
            started_at_open_paren = (ops[L][0] == "eq" and ops[L][1] in OPEN_BRACKETS)
            while R + 1 < n:
                nk, nt = ops[R + 1]
                if nk == "eq":
                    if nt in OPEN_BRACKETS:
                        if not started_at_open_paren and paren_depth == 0:
                            break
                        paren_depth += 1
                    elif nt in CLOSE_BRACKETS:
                        if started_at_open_paren and paren_depth == 0:
                            R += 1
                            break
                        if paren_depth > 0:
                            paren_depth -= 1
                    elif paren_depth == 0 and nt in HARD_DELIMS:
                        R += 1
                        while R + 1 < n and ops[R + 1][0] == "eq" and ops[R + 1][1] in ("”", "’", '"', "'"):
                            R += 1
                        break
                R += 1

        while L < s and ops[L][0] == "eq" and ops[L][1].isspace():
            L += 1
        while R > e and ops[R][0] == "eq" and ops[R][1].isspace():
            R -= 1

        # Rule 2: Conjunction sub-splitting for long clauses (> 12 words)
        if _count_ops_words(ops[L : R + 1]) > 12:
            best_L, best_R = L, R
            p_depth = 0
            for k_idx in range(L, R + 1):
                ok, ot = ops[k_idx]
                if ok == "eq":
                    if ot in OPEN_BRACKETS:
                        p_depth += 1
                    elif ot in CLOSE_BRACKETS and p_depth > 0:
                        p_depth -= 1
                    elif p_depth == 0 and ot.lower() in SUBCLAUSE_CONJ:
                        prev_words = [
                            ops[p][1].lower()
                            for p in range(max(L, k_idx - 5), k_idx)
                            if ops[p][0] == "eq" and is_word_token(ops[p][1])
                        ]
                        if len(prev_words) >= 2 and prev_words[-2] in PREPOSITIONS:
                            continue
                        if prev_words and prev_words[-1] == "both":
                            continue
                        w_left = _count_ops_words(ops[best_L:k_idx])
                        w_right = _count_ops_words(ops[k_idx : best_R + 1])
                        if w_left >= 5 and w_right >= 5:
                            if k_idx > e:
                                has_other = any(
                                    ci2["type"] == "word_edit" and k_idx <= ci2["start"] <= best_R
                                    for ci2 in cluster_infos
                                )
                                if not has_other:
                                    best_R = k_idx - 1
                                    while best_R > e and ops[best_R][0] == "eq" and ops[best_R][1].isspace():
                                        best_R -= 1
                                    break
                            elif k_idx < s:
                                has_other = any(
                                    ci2["type"] == "word_edit" and best_L <= ci2["end"] <= k_idx
                                    for ci2 in cluster_infos
                                )
                                if not has_other:
                                    best_L = k_idx
            L, R = best_L, best_R

        for k_idx in range(L, R + 1):
            hl[k_idx] = True

    # Pass 2: Rule 1 — 3-word window around pure punctuation mismatches
    for ci in cluster_infos:
        if ci["type"] != "punct_only":
            continue
        s, e = ci["start"], ci["end"]
        inner_words = _count_ops_words(ops[s : e + 1])
        L, R = s, e
        need_left = 1
        while L > 0 and need_left > 0:
            L -= 1
            if ops[L][0] != "del" and is_word_token(ops[L][1]):
                need_left -= 1
        need_right = max(1, 3 - inner_words - 1)
        while R + 1 < n and need_right > 0:
            R += 1
            if ops[R][0] != "del" and is_word_token(ops[R][1]):
                need_right -= 1
        for k_idx in range(L, R + 1):
            hl[k_idx] = True

    # Bridge pure whitespace between already-highlighted tokens
    for k_idx in range(1, n - 1):
        if not hl[k_idx] and ops[k_idx][0] == "eq" and ops[k_idx][1].isspace() and hl[k_idx - 1] and hl[k_idx + 1]:
            hl[k_idx] = True

    # Pass 3: Rule 4 — Anti-fragmentation gap bridging (internal gaps only; outer ends never expanded)
    changed = True
    while changed:
        changed = False
        spans = []
        cur_val = hl[0]
        st = 0
        for k_idx in range(1, n):
            if hl[k_idx] != cur_val:
                spans.append((cur_val, st, k_idx - 1, _count_ops_words(ops[st:k_idx])))
                cur_val = hl[k_idx]
                st = k_idx
        spans.append((cur_val, st, n - 1, _count_ops_words(ops[st:n])))

        def _is_new_sent_comma_clause(idx_sp: int) -> bool:
            _, gs, ge, _ = spans[idx_sp]
            prev_end = "".join(t for k, t in ops[spans[idx_sp - 1][1] : spans[idx_sp - 1][2] + 1] if k in ("eq", "ins")).rstrip("”’\"') ")
            g_txt = "".join(t for k, t in ops[gs : ge + 1] if k in ("eq", "ins")).strip()
            return (
                bool(prev_end) and prev_end[-1] in ".?!"
                and bool(g_txt) and g_txt[-1] in (",", ";")
                and not any(ch in g_txt for ch in ".?!")
            )

        for sp_i in range(1, len(spans) - 1):
            is_h, g_s, g_e, g_words = spans[sp_i]
            if is_h:
                continue
            left_w = spans[sp_i - 1][3]
            right_w = spans[sp_i + 1][3]
            # Condition A1: Micro-gap (<= 4 words, e.g. '“God,' or 'I swear to You(s):') tightly sandwiched between yellow spans
            if (
                not _is_new_sent_comma_clause(sp_i)
                and g_words <= 4
                and (left_w + right_w) >= max(8, g_words * 3)
                and max(left_w, right_w) >= 4
            ):
                for k_idx in range(g_s, g_e + 1):
                    hl[k_idx] = True
                changed = True
                break
            # Condition A2: Intra-sentence or short whole-sentence gap (<= 10 words, e.g. 'sin does not represent God, and God is sinless.')
            # Never bridge across a sentence break (.?!) into an unchanged opening comma/semicolon clause of the next sentence!
            if (
                not _is_new_sent_comma_clause(sp_i)
                and g_words <= 10
                and min(left_w, right_w) >= 4
                and (left_w + right_w) >= max(14, g_words * 3)
            ):
                for k_idx in range(g_s, g_e + 1):
                    hl[k_idx] = True
                changed = True
                break

        if not changed and len(spans) >= 7:
            true_indices = [idx_sp for idx_sp, sp in enumerate(spans) if sp[0]]
            for a_i in range(len(true_indices)):
                for b_i in range(len(true_indices) - 1, a_i + 2, -1):
                    si, sj = true_indices[a_i], true_indices[b_i]
                    sub_indices = list(range(si, sj + 1))
                    if any(not spans[idx_s][0] and _is_new_sent_comma_clause(idx_s) for idx_s in sub_indices):
                        continue
                    sub = [spans[idx_s] for idx_s in sub_indices]
                    false_gaps = [sp[3] for sp in sub if not sp[0]]
                    if not false_gaps or max(false_gaps) > 15:
                        continue
                    y_sum = sum(sp[3] for sp in sub if sp[0])
                    u_sum = sum(false_gaps)
                    if y_sum >= max(30, u_sum * 2.2):
                        for sp in sub:
                            if not sp[0]:
                                for k_idx in range(sp[1], sp[2] + 1):
                                    hl[k_idx] = True
                        changed = True
                        break
                if changed:
                    break

    return [(ops[k][0], ops[k][1], "yellow" if hl[k] else None) for k in range(n)]


def apply_clause_yellow_highlights(sent_ops: list) -> list:
    """
    Given a sentence unit's ops, split at unchanged clause delimiters [,;:.?!—]
    and mark only clauses containing edits with highlight='yellow' (trimming outer spaces).
    Returns list of (kind, text, highlight_or_None).
    """
    # Group into clauses split AFTER eq clause delimiters (including trailing quote if any)
    clauses = []
    cur = []
    n = len(sent_ops)
    i = 0
    while i < n:
        k, t = sent_ops[i]
        cur.append((k, t))
        if k == "eq" and t in CLAUSE_SEPS:
            while i + 1 < n and sent_ops[i + 1][0] == "eq" and sent_ops[i + 1][1] in ("”", "’", '"', "'"):
                i += 1
                cur.append(sent_ops[i])
            clauses.append(cur)
            cur = []
        i += 1
    if cur:
        clauses.append(cur)

    edited_flags = [any(k in ("del", "ins") and not t.isspace() for k, t in cl) for cl in clauses]
    result = []
    for c_idx, cl in enumerate(clauses):
        if not edited_flags[c_idx]:
            for k, t in cl:
                result.append((k, t, None))
        else:
            prev_edited = c_idx > 0 and edited_flags[c_idx - 1]
            next_edited = c_idx + 1 < len(clauses) and edited_flags[c_idx + 1]
            first_non_ws = 0
            if not prev_edited:
                while first_non_ws < len(cl) and cl[first_non_ws][0] == "eq" and cl[first_non_ws][1].isspace():
                    first_non_ws += 1
            last_non_ws = len(cl) - 1
            if not next_edited:
                while last_non_ws >= first_non_ws and cl[last_non_ws][0] == "eq" and cl[last_non_ws][1].isspace():
                    last_non_ws -= 1
            for idx, (k, t) in enumerate(cl):
                hl = "yellow" if first_non_ws <= idx <= last_non_ws else None
                result.append((k, t, hl))
    return result


def apply_full_sentence_highlight(sent_ops: list, color: str) -> list:
    """Highlight the full sentence from first non-space token to last non-space token."""
    first_non_ws = 0
    while first_non_ws < len(sent_ops) and sent_ops[first_non_ws][0] == "eq" and sent_ops[first_non_ws][1].isspace():
        first_non_ws += 1
    last_non_ws = len(sent_ops) - 1
    while last_non_ws >= first_non_ws and sent_ops[last_non_ws][0] == "eq" and sent_ops[last_non_ws][1].isspace():
        last_non_ws -= 1
    result = []
    for idx, (k, t) in enumerate(sent_ops):
        hl = color if first_non_ws <= idx <= last_non_ws else None
        result.append((k, t, hl))
    return result


def count_msword_en_words(text: str) -> int:
    """Exact MS Word word count formula verified against test ground truth (RV 609 / UP 452)."""
    tokens = re.split(r"[\s—–]+", text.strip())
    return sum(1 for tok in tokens if re.search(r"[A-Za-z0-9]", tok))


def find_fuzzy_substring_span(haystack: str, needle: str, search_from: int = 0):
    """
    Locate needle inside haystack tolerating French/typographic differences:
    non-breaking spaces (\xa0), multiple spaces, straight vs curly quotes/apostrophes.
    Returns (start_char_idx, end_char_idx) in original haystack, or None.
    """
    needle = needle.strip()
    if not needle:
        return None
    # 1. Direct match
    pos = haystack.find(needle, search_from)
    if pos != -1:
        return (pos, pos + len(needle))
    pos = haystack.find(needle)
    if pos != -1:
        return (pos, pos + len(needle))

    # 2. Normalized regex match preserving exact character positions in haystack
    has_zw = any(z in haystack for z in ("\u200b", "\u200c", "\ufeff"))
    zw_sep = r"[\u200b\u200c\ufeff]*" if has_zw else ""
    pattern_parts = []
    i = 0
    while i < len(needle):
        ch = needle[i]
        if ch in ("\u200b", "\u200c", "\ufeff"):
            i += 1
            continue
        if ch.isspace() or ch in ("\xa0", "\u3000"):
            while i < len(needle) and (needle[i].isspace() or needle[i] in ("\xa0", "\u3000", "\u200b", "\u200c", "\ufeff")):
                i += 1
            pattern_parts.append(r"[\s\xa0\u3000\u200b\u200c\ufeff]+")
        elif ch in ("'", "’", "‘", "‚", "‹", "›"):
            pattern_parts.append(r"['’‘‚‹›]")
            i += 1
        elif ch in ('"', "“", "”", "„", "‟", "«", "»", "「", "」", "『", "』"):
            pattern_parts.append(r'["“”„‟«»「」『』]')
            i += 1
        elif ch in ("-", "–", "—", "―"):
            pattern_parts.append(r"[-–—―]")
            i += 1
        elif ch in ("·", "·", ";", "；", "؛"):
            pattern_parts.append(r"[··;；؛]")
            i += 1
        elif ch in ("?", "？", "؟"):
            pattern_parts.append(r"[?？؟]")
            i += 1
        elif ch in (",", "，", "、", "،"):
            pattern_parts.append(r"[,，、،]")
            i += 1
        else:
            pattern_parts.append(re.escape(ch))
            i += 1
    joiner = zw_sep if zw_sep else ""
    regex = re.compile(joiner.join(pattern_parts), flags=re.UNICODE)
    m = regex.search(haystack, search_from)
    if m:
        return (m.start(), m.end())
    m = regex.search(haystack)
    if m:
        return (m.start(), m.end())
    return None


def apply_highlights_to_col3_cell(tc: ET.Element, char_spans: list):
    """
    Apply character-offset highlights [(start, end, hl_color), ...] onto existing runs of tc
    WITHOUT adding extra line breaks/paragraph splits, preserving original run rPr formatting.
    """
    if not char_spans:
        return
    # Sort spans and resolve any overlap (later or higher-priority RV wins if overlapping, though normally disjoint)
    total_text = get_cell_plain_text(tc)
    char_hl = [None] * len(total_text)
    for s, e, color in char_spans:
        s = max(0, min(len(total_text), s))
        e = max(s, min(len(total_text), e))
        for idx in range(s, e):
            char_hl[idx] = color

    global_offset = 0
    paragraphs = tc.findall("w:p", NS)
    for p_i, p in enumerate(paragraphs):
        if p_i > 0:
            global_offset += 1  # accounts for '\n' between paragraphs in get_cell_plain_text
        new_children = []
        for ch in list(p):
            local = ch.tag.split("}")[-1]
            if local != "r":
                new_children.append(ch)
                continue
            r_text = "".join(t.text or "" for t in ch.findall("w:t", NS))
            if not r_text:
                new_children.append(ch)
                continue
            base_rPr = ch.find("w:rPr", NS)
            r_len = len(r_text)
            r_hls = char_hl[global_offset: global_offset + r_len]
            global_offset += r_len

            # Group contiguous characters with same highlight inside this run
            seg_start = 0
            cur_hl = r_hls[0]
            for k in range(1, r_len):
                if r_hls[k] != cur_hl:
                    seg_txt = r_text[seg_start:k]
                    shd_el = base_rPr.find("w:shd", NS) if base_rPr is not None else None
                    shd_fill = shd_el.attrib.get(f"{W}fill") if shd_el is not None else None
                    if shd_fill in ("auto", "FFFFFF", "ffffff"):
                        shd_fill = None
                    new_children.append(make_run(seg_txt, base_rPr=base_rPr, highlight=cur_hl, shd_fill=shd_fill))
                    seg_start = k
                    cur_hl = r_hls[k]
            seg_txt = r_text[seg_start:r_len]
            shd_el = base_rPr.find("w:shd", NS) if base_rPr is not None else None
            shd_fill = shd_el.attrib.get(f"{W}fill") if shd_el is not None else None
            if shd_fill in ("auto", "FFFFFF", "ffffff"):
                shd_fill = None
            new_children.append(make_run(seg_txt, base_rPr=base_rPr, highlight=cur_hl, shd_fill=shd_fill))

        for ch in list(p):
            p.remove(ch)
        for ch in new_children:
            p.append(ch)
