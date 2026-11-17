import copy
import difflib
import json
import re
import threading
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import quoteattr

W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_URI = "http://www.w3.org/XML/1998/namespace"
PKG_REL_URI = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_URI = "http://schemas.openxmlformats.org/package/2006/content-types"

W = f"{{{W_URI}}}"
R_NS = f"{{{R_URI}}}"
XML_SPACE = f"{{{XML_URI}}}space"
NS = {"w": W_URI, "r": R_URI}

COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
SKIPPED_ROW_MARKER = "【未在总库匹配到原文，已跳过】"

_STRIPPED_COMMENT_PART_SUFFIXES = (
    "commentsextended.xml",
    "commentsids.xml",
    "commentsextensible.xml",
)
_STRIPPED_COMMENT_ZIP_NAMES = frozenset({
    "word/commentsExtended.xml",
    "word/commentsIds.xml",
    "word/commentsExtensible.xml",
    "word/_rels/comments.xml.rels",
    "word/_rels/commentsExtended.xml.rels",
    "word/_rels/commentsIds.xml.rels",
    "word/_rels/commentsExtensible.xml.rels",
})
_STRIPPED_COMMENT_REL_TYPES = (
    "/commentsExtended",
    "/commentsIds",
    "/commentsExtensible",
)
_PEOPLE_ZIP_NAMES = frozenset({
    "word/people.xml",
    "word/_rels/people.xml.rels",
})

_ET_NS_LOCK = threading.Lock()

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

# Strict ECMA-376 CT_PPr child tag order
PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd",
    "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
    "sectPr", "pPrChange",
]
PPR_RANK = {name: idx for idx, name in enumerate(PPR_ORDER)}

# Strict ECMA-376 CT_TcPr child tag order
TCPR_ORDER = [
    "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders",
    "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign",
    "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange",
]
TCPR_RANK = {name: idx for idx, name in enumerate(TCPR_ORDER)}

# Strict ECMA-376 CT_TblPr child tag order
TBLPR_ORDER = [
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
    "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
    "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
    "tblCaption", "tblDescription", "tblPrChange",
]
TBLPR_RANK = {name: idx for idx, name in enumerate(TBLPR_ORDER)}

# Strict ECMA-376 CT_SectPr child tag order
SECTPR_ORDER = [
    "headerReference", "footerReference", "footnotePr", "endnotePr", "type",
    "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType",
    "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection",
    "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
]
SECTPR_RANK = {name: idx for idx, name in enumerate(SECTPR_ORDER)}

INLINE_SEMANTIC_RPR_TAGS = frozenset({
    "b", "bCs", "i", "iCs", "u", "strike", "dstrike", "vertAlign", "caps", "smallCaps", "rtl", "cs"
})

ALLOWED_COL3_RPR_TAGS = frozenset({
    "rFonts", "sz", "szCs", "lang", "b", "bCs", "i", "iCs", "u",
    "rtl", "cs", "color", "vertAlign", "caps", "smallCaps", "strike", "dstrike",
})

_GLOBAL_MARKUP_ID_TAGS = frozenset({
    "ins", "del",
    "rPrChange", "pPrChange", "tblPrChange", "trPrChange", "tcPrChange",
    "sectPrChange", "tblGridChange", "numberingChange",
    "bookmarkStart", "bookmarkEnd",
    "commentRangeStart", "commentRangeEnd", "commentReference", "comment",
    "moveFrom", "moveTo",
    "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
    "customXmlInsRangeStart", "customXmlInsRangeEnd",
    "customXmlDelRangeStart", "customXmlDelRangeEnd",
    "customXmlMoveFromRangeStart", "customXmlMoveFromRangeEnd",
    "customXmlMoveToRangeStart", "customXmlMoveToRangeEnd",
    "cellIns", "cellDel", "cellMerge",
    "permStart", "permEnd",
})


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
    m = re.search(r"<w:(?:document|comments)\b[^>]*>", head)
    open_tag = m.group(0) if m else ""
    ns_pairs = re.findall(r'xmlns:([A-Za-z0-9_.-]+)="([^"]+)"', open_tag)
    with _ET_NS_LOCK:
        for prefix, uri in ns_pairs:
            if prefix != "xml":
                ET.register_namespace(prefix, uri)
        ET.register_namespace("w", W_URI)
        ET.register_namespace("r", R_URI)
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
        raise ValueError("OpenXML member exceeds maximum allowed byte size.")
    if _DTD_RE.search(xml_bytes):
        raise ValueError("Unsafe XML DOCTYPE/ENTITY declaration rejected in OpenXML document.")


def _local_name(tag_or_attr: str) -> str:
    return tag_or_attr.split("}")[-1] if "}" in tag_or_attr else tag_or_attr


def _ns_uri(tag_or_attr: str) -> str:
    if tag_or_attr.startswith("{") and "}" in tag_or_attr:
        return tag_or_attr[1:tag_or_attr.index("}")]
    return ""


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


def new_comments_root() -> ET.Element:
    return ET.Element(f"{W}comments")


def load_docx_comments_xml(docx_path: Path):
    """Load word/comments.xml from a .docx if present, returning sanitized <w:comments> Element or None."""
    with zipfile.ZipFile(docx_path, "r") as z:
        _validate_docx_zip(z)
        if "word/comments.xml" not in z.namelist():
            return None
        info = z.getinfo("word/comments.xml")
        with z.open(info, "r") as fp:
            xml_bytes = fp.read(MAX_DOCX_MEMBER_BYTES + 1)
    _assert_safe_openxml_bytes(xml_bytes)
    raw_root = ET.fromstring(xml_bytes)
    clean_root = new_comments_root()
    seen_ids = set()
    for c_el in raw_root.findall(f"{W}comment"):
        cid_raw = c_el.attrib.get(f"{W}id", "").strip()
        try:
            cid_int = int(cid_raw)
        except ValueError:
            continue
        if cid_int in seen_ids:
            continue
        seen_ids.add(cid_int)
        clean_root.append(sanitize_comment_element(c_el, cid_int))
    return clean_root


def load_docx_comments_map(docx_path: Path) -> dict:
    """Return {str(w:id): <w:comment Element>} from word/comments.xml of docx_path (empty dict if none)."""
    c_root = load_docx_comments_xml(docx_path)
    if c_root is None:
        return {}
    out = {}
    for c_el in c_root.findall(f"{W}comment"):
        cid = c_el.attrib.get(f"{W}id", "").strip()
        if cid and cid not in out:
            out[cid] = c_el
    return out


def load_docx_styles_xml(docx_path: Path):
    """Load word/styles.xml from a .docx if present, returning <w:styles> Element or None."""
    with zipfile.ZipFile(docx_path, "r") as z:
        _validate_docx_zip(z)
        if "word/styles.xml" not in z.namelist():
            return None
        info = z.getinfo("word/styles.xml")
        with z.open(info, "r") as fp:
            xml_bytes = fp.read(MAX_DOCX_MEMBER_BYTES + 1)
    _assert_safe_openxml_bytes(xml_bytes)
    return ET.fromstring(xml_bytes)


class MarkupIdAllocator:
    """
    Unified global OOXML CT_Markup integer w:id allocator.
    Scans ALL w:id attributes across ins, del, *PrChange, bookmarkStart/End,
    commentRangeStart/End/Reference, moveFrom/To, etc. in both doc_root and comments_root
    so newly created revisions, bookmarks, and migrated comments never collide!
    """

    def __init__(self, doc_root: ET.Element = None, comments_root: ET.Element = None):
        self._lock = threading.Lock()
        max_id = 0
        found_any = False
        for tree in (doc_root, comments_root):
            if tree is None:
                continue
            for elem in tree.iter():
                loc = _local_name(elem.tag)
                val = elem.attrib.get(f"{W}id", "").strip()
                if not val and loc in _GLOBAL_MARKUP_ID_TAGS:
                    val = elem.attrib.get("id", "").strip()
                if val:
                    try:
                        iv = int(val)
                        found_any = True
                        if iv > max_id:
                            max_id = iv
                    except ValueError:
                        pass
        self._next = (max_id + 1) if found_any else 1

    def next_id(self) -> int:
        with self._lock:
            nid = self._next
            self._next += 1
            return nid


# Backward-compatible alias
CommentIdAllocator = MarkupIdAllocator


def _clean_comment_subtree_children(parent_in: ET.Element) -> list:
    """Clean children inside a <w:comment> paragraph/run so zero broken external rels, styles, bookmarks, or paraIds remain."""
    UNWRAP_TAGS = {"hyperlink", "smartTag", "customXml", "ins", "moveTo", "fldSimple"}
    DROP_TAGS = {
        "del", "moveFrom", "moveFromRangeStart", "moveFromRangeEnd",
        "moveToRangeStart", "moveToRangeEnd",
        "bookmarkStart", "bookmarkEnd", "permStart", "permEnd", "proofErr",
        "footnoteReference", "endnoteReference",
        "commentRangeStart", "commentRangeEnd", "commentReference",
        "drawing", "pict", "object",
    }
    ALLOWED_RUN_CHILDREN = {
        "rPr", "t", "tab", "br", "cr", "noBreakHyphen", "softHyphen", "annotationRef"
    }

    result = []
    for ch in list(parent_in):
        ns = _ns_uri(ch.tag)
        loc = _local_name(ch.tag)
        if loc == "sdt" and ns == W_URI:
            sdt_content = ch.find(f"{W}sdtContent")
            if sdt_content is not None:
                result.extend(_clean_comment_subtree_children(sdt_content))
            continue
        if ns != W_URI:
            continue
        if loc in DROP_TAGS:
            continue
        if loc in UNWRAP_TAGS:
            result.extend(_clean_comment_subtree_children(ch))
            continue
        if loc == "pPr":
            new_ppr = ET.Element(f"{W}pPr")
            for sub in list(ch):
                if _ns_uri(sub.tag) != W_URI:
                    continue
                sloc = _local_name(sub.tag)
                if sloc in ("pStyle", "numPr", "sectPr", "pPrChange"):
                    continue
                cp_sub = copy.deepcopy(sub)
                _strip_unsafe_attrs_recursive(cp_sub)
                if sloc == "rPr":
                    remove_rpr_child(cp_sub, "rStyle")
                    remove_rpr_child(cp_sub, "rPrChange")
                new_ppr.append(cp_sub)
            if len(new_ppr) > 0:
                result.append(new_ppr)
            continue
        if loc == "r":
            new_r = ET.Element(f"{W}r")
            has_content = False
            for sub in list(ch):
                if _ns_uri(sub.tag) != W_URI:
                    continue
                sloc = _local_name(sub.tag)
                if sloc not in ALLOWED_RUN_CHILDREN:
                    continue
                if sloc == "rPr":
                    cp_rpr = copy.deepcopy(sub)
                    _strip_unsafe_attrs_recursive(cp_rpr)
                    remove_rpr_child(cp_rpr, "rStyle")
                    remove_rpr_child(cp_rpr, "rPrChange")
                    if len(cp_rpr) > 0:
                        new_r.insert(0, cp_rpr)
                else:
                    cp_sub = copy.deepcopy(sub)
                    _strip_unsafe_attrs_recursive(cp_sub)
                    new_r.append(cp_sub)
                    has_content = True
            if has_content:
                result.append(new_r)
            continue
    return result


def _strip_unsafe_attrs_recursive(elem: ET.Element) -> None:
    """Keep only W_URI (excluding rsid/paraId/textId) and XML_URI attributes across elem subtree."""
    for node in elem.iter():
        for attr_k in list(node.attrib.keys()):
            ans = _ns_uri(attr_k)
            aloc = _local_name(attr_k)
            if aloc in ("paraId", "textId") or aloc.startswith("rsid"):
                del node.attrib[attr_k]
            elif ans not in (W_URI, XML_URI, ""):
                del node.attrib[attr_k]


def sanitize_comment_element(comment_el: ET.Element, new_id: int) -> ET.Element:
    """Build a 100% self-contained, valid <w:comment> element with w:id=new_id."""
    author = comment_el.attrib.get(f"{W}author") or "整理组"
    date_val = comment_el.attrib.get(f"{W}date")
    initials = comment_el.attrib.get(f"{W}initials")

    attrs = {f"{W}id": str(new_id), f"{W}author": author}
    if date_val:
        attrs[f"{W}date"] = date_val
    if initials is not None:
        attrs[f"{W}initials"] = initials

    out_c = ET.Element(f"{W}comment", attrs)
    for p in comment_el.findall(f".//{W}p"):
        # Avoid nested paragraphs if any table existed inside comment; flatten paragraphs cleanly
        new_p = ET.Element(f"{W}p")
        for cleaned_ch in _clean_comment_subtree_children(p):
            new_p.append(cleaned_ch)
        out_c.append(new_p)

    if len(out_c) == 0:
        out_c.append(ET.Element(f"{W}p"))
    return out_c


def _make_comment_ref_run(cid_str: str) -> ET.Element:
    r = ET.Element(f"{W}r")
    ET.SubElement(r, f"{W}commentReference", {f"{W}id": cid_str})
    return r


def _insert_after_ppr(p_el: ET.Element, new_el: ET.Element) -> None:
    """Insert new_el at the start of <w:p> immediately AFTER <w:pPr> (preserving ECMA-376 CT_P order)."""
    if len(p_el) > 0 and _local_name(p_el[0].tag) == "pPr":
        p_el.insert(1, new_el)
    else:
        p_el.insert(0, new_el)


def _clean_src_col3_p_children(node: ET.Element) -> list:
    """
    Flatten & sanitize children of a cloned library Col 3 <w:p>:
      - Unwrap sdt/sdtContent, hyperlink, ins, moveTo, smartTag, customXml, fldSimple
      - Drop del, moveFrom*, bookmarkStart/End, footnoteReference, endnoteReference, permStart/End, proofErr, drawing/pict/object
      - Keep commentRangeStart, commentRangeEnd, and <w:r> (splitting out commentReference into its own clean <w:r> if mixed with text)
      - Strip w14:paraId / w14:textId and external r:* attributes.
    """
    UNWRAP_TAGS = {"hyperlink", "smartTag", "customXml", "ins", "moveTo", "fldSimple"}
    DROP_TAGS = {
        "pPr", "del", "moveFrom", "moveFromRangeStart", "moveFromRangeEnd",
        "moveToRangeStart", "moveToRangeEnd",
        "bookmarkStart", "bookmarkEnd", "permStart", "permEnd", "proofErr",
        "footnoteReference", "endnoteReference",
        "drawing", "pict", "object",
    }
    TEXT_RUN_CHILDREN = {"t", "tab", "br", "cr", "noBreakHyphen", "softHyphen"}

    out = []
    for ch in list(node):
        ns = _ns_uri(ch.tag)
        loc = _local_name(ch.tag)
        if loc == "sdt" and ns == W_URI:
            sc = ch.find(f"{W}sdtContent")
            if sc is not None:
                out.extend(_clean_src_col3_p_children(sc))
            continue
        if ns != W_URI:
            continue
        if loc in DROP_TAGS:
            continue
        if loc in UNWRAP_TAGS:
            out.extend(_clean_src_col3_p_children(ch))
            continue
        if loc in ("commentRangeStart", "commentRangeEnd"):
            cid = ch.attrib.get(f"{W}id", "").strip()
            if cid:
                out.append(ET.Element(f"{W}{loc}", {f"{W}id": cid}))
            continue
        if loc == "commentReference":
            cid = ch.attrib.get(f"{W}id", "").strip()
            if cid:
                out.append(_make_comment_ref_run(cid))
            continue
        if loc == "r":
            old_rpr = ch.find(f"{W}rPr")
            # Check for nested commentRangeStart/End or text/commentReference inside <w:r>
            text_subs = []
            cref_ids = []
            for sub in list(ch):
                if _ns_uri(sub.tag) != W_URI:
                    continue
                sloc = _local_name(sub.tag)
                if sloc in TEXT_RUN_CHILDREN:
                    cp_s = copy.deepcopy(sub)
                    _strip_unsafe_attrs_recursive(cp_s)
                    text_subs.append(cp_s)
                elif sloc == "commentReference":
                    cid = sub.attrib.get(f"{W}id", "").strip()
                    if cid:
                        cref_ids.append(cid)
                elif sloc in ("commentRangeStart", "commentRangeEnd"):
                    cid = sub.attrib.get(f"{W}id", "").strip()
                    if cid:
                        out.append(ET.Element(f"{W}{sloc}", {f"{W}id": cid}))
            if text_subs:
                nr = ET.Element(f"{W}r")
                if old_rpr is not None:
                    cp_rpr = copy.deepcopy(old_rpr)
                    _strip_unsafe_attrs_recursive(cp_rpr)
                    nr.append(cp_rpr)
                for ts in text_subs:
                    nr.append(ts)
                out.append(nr)
            for cid in cref_ids:
                out.append(_make_comment_ref_run(cid))
    return out


_SAFE_COL3_PPR_TAGS = frozenset({
    "keepNext", "keepLines", "widowControl", "suppressAutoHyphens",
    "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
    "spacing", "ind", "contextualSpacing", "jc", "textAlignment", "outlineLvl",
})

_DEFAULT_STYLE_SAFE_PPR_TAGS = frozenset({
    "widowControl", "suppressAutoHyphens", "autoSpaceDE", "autoSpaceDN",
    "bidi", "adjustRightInd", "snapToGrid", "ind", "jc", "outlineLvl",
})

_BOOLEAN_TOGGLE_RPR_TAGS = frozenset({
    "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "rtl", "cs",
})


def insert_ppr_child_ordered(pPr: ET.Element, new_child: ET.Element):
    new_local = _local_name(new_child.tag)
    for existing in list(pPr):
        if _local_name(existing.tag) == new_local:
            pPr.remove(existing)
    target_rank = PPR_RANK.get(new_local, 999)
    for idx, child in enumerate(list(pPr)):
        c_local = _local_name(child.tag)
        if PPR_RANK.get(c_local, 999) > target_rank:
            pPr.insert(idx, new_child)
            return
    pPr.append(new_child)


def remove_ppr_child(pPr: ET.Element, local_name: str):
    if pPr is None:
        return
    for existing in list(pPr):
        if _local_name(existing.tag) == local_name:
            pPr.remove(existing)


def insert_tcpr_child_ordered(tcPr: ET.Element, new_child: ET.Element):
    new_local = _local_name(new_child.tag)
    for existing in list(tcPr):
        if _local_name(existing.tag) == new_local:
            tcPr.remove(existing)
    target_rank = TCPR_RANK.get(new_local, 999)
    for idx, child in enumerate(list(tcPr)):
        c_local = _local_name(child.tag)
        if TCPR_RANK.get(c_local, 999) > target_rank:
            tcPr.insert(idx, new_child)
            return
    tcPr.append(new_child)


def insert_tblpr_child_ordered(tblPr: ET.Element, new_child: ET.Element):
    new_local = _local_name(new_child.tag)
    for existing in list(tblPr):
        if _local_name(existing.tag) == new_local:
            tblPr.remove(existing)
    target_rank = TBLPR_RANK.get(new_local, 999)
    for idx, child in enumerate(list(tblPr)):
        c_local = _local_name(child.tag)
        if TBLPR_RANK.get(c_local, 999) > target_rank:
            tblPr.insert(idx, new_child)
            return
    tblPr.append(new_child)


def insert_sectpr_child_ordered(sectPr: ET.Element, new_child: ET.Element):
    new_local = _local_name(new_child.tag)
    if new_local not in ("headerReference", "footerReference"):
        for existing in list(sectPr):
            if _local_name(existing.tag) == new_local:
                sectPr.remove(existing)
    target_rank = SECTPR_RANK.get(new_local, 999)
    for idx, child in enumerate(list(sectPr)):
        c_local = _local_name(child.tag)
        if SECTPR_RANK.get(c_local, 999) > target_rank:
            sectPr.insert(idx, new_child)
            return
    sectPr.append(new_child)


def _parse_styles_info(src_styles_root: ET.Element = None) -> dict:
    info = {
        "doc_pPr": None,
        "doc_rPr": None,
        "p_styles": {},
        "c_styles": {},
        "default_p_id": None,
    }
    if src_styles_root is None:
        return info
    info["doc_pPr"] = src_styles_root.find("./w:docDefaults/w:pPrDefault/w:pPr", NS)
    info["doc_rPr"] = src_styles_root.find("./w:docDefaults/w:rPrDefault/w:rPr", NS)
    for st in src_styles_root.findall(f"{W}style"):
        stype = st.attrib.get(f"{W}type", "").strip()
        sid = st.attrib.get(f"{W}styleId", "").strip()
        if not sid:
            continue
        is_def = st.attrib.get(f"{W}default", "").strip() in ("1", "true")
        if stype == "paragraph":
            info["p_styles"][sid] = st
            if is_def and info["default_p_id"] is None:
                info["default_p_id"] = sid
        elif stype == "character":
            info["c_styles"][sid] = st
    return info


def _get_style_chain(style_id: str, styles_dict: dict) -> list:
    if not style_id or style_id not in styles_dict:
        return []
    chain = []
    seen = set()
    cur = style_id
    while cur and cur in styles_dict and cur not in seen:
        seen.add(cur)
        st = styles_dict[cur]
        chain.append(st)
        based = st.find(f"{W}basedOn")
        cur = based.attrib.get(f"{W}val", "").strip() if based is not None else ""
    chain.reverse()
    return chain


def _merge_rpr_layer(dst_rPr: ET.Element, src_rPr: ET.Element, allowed_tags=ALLOWED_COL3_RPR_TAGS) -> None:
    if src_rPr is None:
        return
    for ch in src_rPr:
        if _ns_uri(ch.tag) != W_URI:
            continue
        loc = _local_name(ch.tag)
        if loc not in allowed_tags:
            continue
        if loc == "rFonts":
            dst_rf = dst_rPr.find(f"{W}rFonts")
            if dst_rf is None:
                dst_rf = ET.Element(f"{W}rFonts")
                insert_rpr_child_ordered(dst_rPr, dst_rf)
            theme_clear_map = {
                f"{W}ascii": f"{W}asciiTheme",
                f"{W}hAnsi": f"{W}hAnsiTheme",
                f"{W}eastAsia": f"{W}eastAsiaTheme",
                f"{W}cs": f"{W}cstheme",
            }
            for ak, av in ch.attrib.items():
                if _ns_uri(ak) not in (W_URI, ""):
                    continue
                dst_rf.set(ak, av)
                if ak in theme_clear_map and theme_clear_map[ak] in dst_rf.attrib:
                    del dst_rf.attrib[theme_clear_map[ak]]
        else:
            cp = copy.deepcopy(ch)
            _strip_unsafe_attrs_recursive(cp)
            insert_rpr_child_ordered(dst_rPr, cp)


def _finalize_col3_rpr(rPr: ET.Element, fallback_mark_rPr: ET.Element = None) -> ET.Element:
    """Clean disabled boolean toggles, fill missing font/size from paragraph-mark fallback or defaults, and enforce RPR_ORDER."""
    for ch in list(rPr):
        if _ns_uri(ch.tag) != W_URI:
            rPr.remove(ch)
            continue
        loc = _local_name(ch.tag)
        if loc in _BOOLEAN_TOGGLE_RPR_TAGS:
            val = ch.attrib.get(f"{W}val", "1").strip().lower()
            if val in ("0", "false", "off"):
                rPr.remove(ch)
        elif loc == "u":
            uval = ch.attrib.get(f"{W}val", "single").strip().lower()
            if uval in ("none", "0", "false"):
                rPr.remove(ch)
        elif loc in ("rStyle", "rPrChange", "highlight", "shd"):
            rPr.remove(ch)

    if fallback_mark_rPr is not None:
        for fb_tag in ("rFonts", "sz", "szCs", "lang"):
            if rPr.find(f"{W}{fb_tag}") is None:
                fb_el = fallback_mark_rPr.find(f"{W}{fb_tag}")
                if fb_el is not None:
                    cp_fb = copy.deepcopy(fb_el)
                    _strip_unsafe_attrs_recursive(cp_fb)
                    if fb_tag == "rFonts":
                        for t_attr in (f"{W}asciiTheme", f"{W}hAnsiTheme", f"{W}eastAsiaTheme", f"{W}cstheme"):
                            cp_fb.attrib.pop(t_attr, None)
                    insert_rpr_child_ordered(rPr, cp_fb)

    rf = rPr.find(f"{W}rFonts")
    if rf is None or not rf.attrib:
        insert_rpr_child_ordered(rPr, ET.Element(f"{W}rFonts", {
            f"{W}ascii": "Georgia",
            f"{W}hAnsi": "Georgia",
            f"{W}cs": "Georgia",
        }))
    else:
        asc = rf.attrib.get(f"{W}ascii")
        hansi = rf.attrib.get(f"{W}hAnsi")
        if asc and not hansi:
            rf.set(f"{W}hAnsi", asc)
            rf.attrib.pop(f"{W}hAnsiTheme", None)
        elif hansi and not asc:
            rf.set(f"{W}ascii", hansi)
            rf.attrib.pop(f"{W}asciiTheme", None)

    sz_el = rPr.find(f"{W}sz")
    szcs_el = rPr.find(f"{W}szCs")
    if sz_el is None and szcs_el is not None:
        insert_rpr_child_ordered(rPr, ET.Element(f"{W}sz", dict(szcs_el.attrib)))
    elif szcs_el is None and sz_el is not None:
        insert_rpr_child_ordered(rPr, ET.Element(f"{W}szCs", dict(sz_el.attrib)))
    elif sz_el is None and szcs_el is None:
        insert_rpr_child_ordered(rPr, ET.Element(f"{W}sz", {f"{W}val": "24"}))
        insert_rpr_child_ordered(rPr, ET.Element(f"{W}szCs", {f"{W}val": "24"}))

    return rPr


def _resolve_src_col3_ppr(old_pPr: ET.Element, styles_info: dict, fallback_spacing_el: ET.Element = None) -> ET.Element:
    """
    Resolve library Col 3 paragraph <w:pPr> preserving its OWN alignment (jc), indentation (ind),
    spacing, bidi, and outlineLvl (including inheritance from src_styles_root pStyle / docDefaults).
    Removes pStyle after resolving so target styles.xml never remaps style IDs.
    """
    new_pPr = ET.Element(f"{W}pPr")

    # 1. docDefaults pPr
    doc_pPr = styles_info.get("doc_pPr")
    if doc_pPr is not None:
        for ch in doc_pPr:
            if _ns_uri(ch.tag) == W_URI and _local_name(ch.tag) in _DEFAULT_STYLE_SAFE_PPR_TAGS:
                cp = copy.deepcopy(ch)
                _strip_unsafe_attrs_recursive(cp)
                insert_ppr_child_ordered(new_pPr, cp)

    # 2. Paragraph style chain (explicit pStyle or default Normal style)
    explicit_pstyle = None
    if old_pPr is not None:
        ps_el = old_pPr.find(f"{W}pStyle")
        if ps_el is not None:
            explicit_pstyle = ps_el.attrib.get(f"{W}val", "").strip() or None

    eff_pstyle = explicit_pstyle or styles_info.get("default_p_id")
    is_explicit_non_default = bool(explicit_pstyle and explicit_pstyle != styles_info.get("default_p_id"))
    style_allowed = _SAFE_COL3_PPR_TAGS if is_explicit_non_default else _DEFAULT_STYLE_SAFE_PPR_TAGS
    for st_el in _get_style_chain(eff_pstyle, styles_info.get("p_styles", {})):
        st_ppr = st_el.find(f"{W}pPr")
        if st_ppr is None:
            continue
        for ch in st_ppr:
            if _ns_uri(ch.tag) == W_URI and _local_name(ch.tag) in style_allowed:
                cp = copy.deepcopy(ch)
                _strip_unsafe_attrs_recursive(cp)
                insert_ppr_child_ordered(new_pPr, cp)

    # 3. Direct paragraph pPr properties from src_tc3
    if old_pPr is not None:
        for ch in old_pPr:
            if _ns_uri(ch.tag) == W_URI and _local_name(ch.tag) in _SAFE_COL3_PPR_TAGS:
                cp = copy.deepcopy(ch)
                _strip_unsafe_attrs_recursive(cp)
                insert_ppr_child_ordered(new_pPr, cp)

    if new_pPr.find(f"{W}spacing") is None and fallback_spacing_el is not None:
        cp_sp = copy.deepcopy(fallback_spacing_el)
        _strip_unsafe_attrs_recursive(cp_sp)
        insert_ppr_child_ordered(new_pPr, cp_sp)

    # 4. If non-RTL body paragraph still has no explicit <w:jc>, default to justified ("both")
    bidi_el = new_pPr.find(f"{W}bidi")
    is_rtl = bidi_el is not None and bidi_el.attrib.get(f"{W}val", "1").strip().lower() not in ("0", "false", "off")
    if not is_rtl and new_pPr.find(f"{W}jc") is None:
        insert_ppr_child_ordered(new_pPr, ET.Element(f"{W}jc", {f"{W}val": "both"}))

    return new_pPr


def clean_and_migrate_col3_paragraphs(
    src_tc3: ET.Element,
    target_tc3: ET.Element,
    target_tc2: ET.Element,
    src_comments_map: dict = None,
    target_comments_root: ET.Element = None,
    id_allocator: CommentIdAllocator = None,
    src_styles_root: ET.Element = None,
) -> list:
    """
    Clone text, inline formatting, and Word comments from library Col 3 (`src_tc3`) while:
      1. 100% preserving `src_tc3`'s OWN paragraph formatting (`jc`, `ind`, `spacing`, `bidi`, `outlineLvl`,
         resolving inherited `pStyle` / `docDefaults` from `src_styles_root` and removing `pStyle`)
         and `src_tc3`'s OWN per-run `<w:rPr>` (`rFonts`, `sz`, `szCs`, `lang`, `b`, `bCs`, `i`, `iCs`, `u`,
         `rtl`, `cs`, `color` — NEVER inheriting bold `<w:b/>` or alignment from target cells!).
      2. Stripping all non-comment OOXML troublemakers (bookmarks, footnotes, permissions, proofErr, smartTag,
         customXml, sdt, hyperlinks, track changes ins/del, w14:paraId/textId).
      3. Migrating Word comments from `src_comments_map` into `target_comments_root` with globally unique integer `w:id`s
         and guaranteeing complete, properly ordered `(commentRangeStart, commentRangeEnd, commentReference)` triplets.
    """
    styles_info = _parse_styles_info(src_styles_root)
    tpl_pPr = get_cell_template_ppr(target_tc3, fallback_tc=target_tc2)
    fallback_spacing = tpl_pPr.find(f"{W}spacing") if tpl_pPr is not None else None
    src_cmap = src_comments_map if (src_comments_map and target_comments_root is not None and id_allocator is not None) else {}

    paras = []
    for p in src_tc3.findall("w:p", NS):
        old_pPr = p.find("w:pPr", NS)
        p_mark_rPr = old_pPr.find("w:rPr", NS) if old_pPr is not None else None

        # Build paragraph-level inherited rPr (docDefaults -> pStyle chain)
        para_base_rPr = ET.Element(f"{W}rPr")
        _merge_rpr_layer(para_base_rPr, styles_info.get("doc_rPr"))
        explicit_pstyle = None
        if old_pPr is not None:
            ps_el = old_pPr.find(f"{W}pStyle")
            if ps_el is not None:
                explicit_pstyle = ps_el.attrib.get(f"{W}val", "").strip() or None
        eff_pstyle = explicit_pstyle or styles_info.get("default_p_id")
        for st_el in _get_style_chain(eff_pstyle, styles_info.get("p_styles", {})):
            _merge_rpr_layer(para_base_rPr, st_el.find(f"{W}rPr"))

        new_pPr = _resolve_src_col3_ppr(old_pPr, styles_info, fallback_spacing_el=fallback_spacing)
        clean_mark_rPr = _finalize_col3_rpr(copy.deepcopy(para_base_rPr), fallback_mark_rPr=p_mark_rPr)
        for non_mark in ("b", "bCs", "i", "iCs", "u", "strike", "dstrike", "color"):
            remove_rpr_child(clean_mark_rPr, non_mark)
        if len(clean_mark_rPr) > 0:
            insert_ppr_child_ordered(new_pPr, clean_mark_rPr)

        cp = ET.Element(f"{W}p")
        if len(new_pPr) > 0:
            cp.append(new_pPr)

        cleaned_children = _clean_src_col3_p_children(p)
        for ch in cleaned_children:
            if _local_name(ch.tag) == "r":
                cref = ch.find(f"{W}commentReference")
                if cref is not None and ch.find(f"{W}t") is None:
                    old_rPr = ch.find(f"{W}rPr")
                    if old_rPr is not None:
                        ch.remove(old_rPr)
                    if len(clean_mark_rPr) > 0:
                        ch.insert(0, copy.deepcopy(clean_mark_rPr))
                else:
                    old_rPr = ch.find("w:rPr", NS)
                    new_rPr = copy.deepcopy(para_base_rPr)
                    if old_rPr is not None:
                        rs_el = old_rPr.find(f"{W}rStyle")
                        if rs_el is not None:
                            rs_id = rs_el.attrib.get(f"{W}val", "").strip()
                            for cst in _get_style_chain(rs_id, styles_info.get("c_styles", {})):
                                _merge_rpr_layer(new_rPr, cst.find(f"{W}rPr"))
                        _merge_rpr_layer(new_rPr, old_rPr)
                        ch.remove(old_rPr)
                    _finalize_col3_rpr(new_rPr, fallback_mark_rPr=p_mark_rPr)
                    if len(new_rPr) > 0:
                        ch.insert(0, new_rPr)
            cp.append(ch)
        paras.append(cp)

    if not paras:
        empty_p = ET.Element(f"{W}p")
        def_pPr = _resolve_src_col3_ppr(None, styles_info, fallback_spacing_el=fallback_spacing)
        if len(def_pPr) > 0:
            empty_p.append(def_pPr)
        paras.append(empty_p)

    # Remap & repair Word comment markers across `paras`
    _remap_and_repair_cell_comments(paras, src_cmap, target_comments_root, id_allocator)
    return paras


def extract_cell_comment_anchors(tc: ET.Element, comments_root: ET.Element = None) -> list:
    """
    Extract all Word comments anchored in `tc` along with their visible text snippet,
    surrounding context, relative position, and cloned <w:comment> element so they can be
    re-anchored losslessly onto a rebuilt cell (e.g. Col 2 diff runs in step1_render_draft.py).
    """
    if tc is None or comments_root is None:
        return []
    cmap = {}
    for c_el in comments_root.findall(f"{W}comment"):
        cid = c_el.attrib.get(f"{W}id", "").strip()
        if cid:
            cmap[cid] = c_el
    if not cmap:
        return []

    chars = []
    open_starts = {}   # cid -> start_char_idx
    spans = []         # list of (cid, start_idx, end_idx)
    seen_cids = set()
    ref_only = []      # list of (cid, char_idx)

    def _walk_node(node: ET.Element):
        for ch in node:
            loc = _local_name(ch.tag)
            if loc in ("del", "moveFrom"):
                continue
            if loc == "commentRangeStart":
                cid = ch.attrib.get(f"{W}id", "").strip()
                if cid in cmap:
                    if cid in open_starts:
                        spans.append((cid, open_starts[cid], len(chars)))
                        seen_cids.add(cid)
                    open_starts[cid] = len(chars)
            elif loc == "commentRangeEnd":
                cid = ch.attrib.get(f"{W}id", "").strip()
                if cid in cmap:
                    s_off = open_starts.pop(cid, 0)
                    spans.append((cid, s_off, len(chars)))
                    seen_cids.add(cid)
            elif loc == "commentReference":
                cid = ch.attrib.get(f"{W}id", "").strip()
                if cid in cmap:
                    ref_only.append((cid, len(chars)))
            elif loc == "t":
                if ch.text:
                    chars.extend(ch.text)
            else:
                _walk_node(ch)

    for p_i, p in enumerate(tc.findall("w:p", NS)):
        if p_i > 0:
            chars.append("\n")
        _walk_node(p)

    full_text = "".join(chars)
    total_len = len(full_text)
    for cid, s_off in open_starts.items():
        spans.append((cid, s_off, total_len))
        seen_cids.add(cid)
    for cid, r_off in ref_only:
        if cid not in seen_cids:
            spans.append((cid, r_off, r_off))
            seen_cids.add(cid)

    anchors = []
    for cid, s_off, e_off in spans:
        anchors.append({
            "old_cid": cid,
            "anchor_text": full_text[s_off:e_off],
            "left_ctx": full_text[max(0, s_off - 45):s_off],
            "right_ctx": full_text[e_off:min(total_len, e_off + 45)],
            "rel_pos": s_off / max(1, total_len),
            "comment_el": copy.deepcopy(cmap[cid]),
        })
    return anchors


def _locate_anchor_in_rebuilt_text(new_text: str, anc: dict) -> tuple:
    """Locate best (start_char, end_char) in `new_text` for a saved comment anchor."""
    total_len = len(new_text)
    if total_len == 0:
        return (0, 0)

    atxt = anc.get("anchor_text") or ""
    rel_pos = float(anc.get("rel_pos", 0.0))
    left_ctx = anc.get("left_ctx") or ""
    right_ctx = anc.get("right_ctx") or ""
    est_pos = int(rel_pos * total_len)

    def _score_candidate(cs: int, ce: int) -> float:
        l_sub = new_text[max(0, cs - 45):cs]
        r_sub = new_text[ce:min(total_len, ce + 45)]
        l_sim = difflib.SequenceMatcher(None, left_ctx, l_sub).ratio() if left_ctx else 0.5
        r_sim = difflib.SequenceMatcher(None, right_ctx, r_sub).ratio() if right_ctx else 0.5
        pos_pen = 0.35 * abs((cs / max(1, total_len)) - rel_pos)
        return l_sim + r_sim - pos_pen

    for cand_needle in ([atxt, atxt.strip()] if atxt.strip() != atxt else [atxt]):
        if not cand_needle:
            continue
        exact_locs = []
        pos = new_text.find(cand_needle)
        while pos != -1:
            exact_locs.append((pos, pos + len(cand_needle)))
            pos = new_text.find(cand_needle, pos + 1)
        if len(exact_locs) == 1:
            return exact_locs[0]
        if len(exact_locs) > 1:
            return max(exact_locs, key=lambda pair: _score_candidate(pair[0], pair[1]))

    if atxt.strip():
        f_loc = find_fuzzy_substring_span(new_text, atxt, search_from=max(0, est_pos - 120))
        if f_loc is not None:
            return f_loc
        # Windowed SequenceMatcher fallback when words inside anchor_text were edited
        win_rad = max(160, len(atxt) * 2)
        w_s = max(0, est_pos - win_rad)
        w_e = min(total_len, est_pos + len(atxt) + win_rad)
        win_str = new_text[w_s:w_e]
        sm = difflib.SequenceMatcher(None, atxt.strip(), win_str, autojunk=False)
        blocks = [b for b in sm.get_matching_blocks() if b.size > 0]
        matched_chars = sum(b.size for b in blocks)
        if blocks and matched_chars >= max(3, int(len(atxt.strip()) * 0.45)):
            return (w_s + blocks[0].b, w_s + blocks[-1].b + blocks[-1].size)

    # Paragraph fallback (first / closest / last paragraph)
    p_ranges = []
    cur_off = 0
    for part in new_text.split("\n"):
        p_ranges.append((cur_off, cur_off + len(part)))
        cur_off += len(part) + 1
    if p_ranges:
        p_idx = min(len(p_ranges) - 1, max(0, round(rel_pos * max(0, len(p_ranges) - 1))))
        return p_ranges[p_idx]
    return (0, total_len)


def reanchor_comments_to_cell(
    tc: ET.Element,
    anchors: list,
    target_comments_root: ET.Element,
    id_allocator: MarkupIdAllocator,
) -> int:
    """
    Re-anchor extracted comment anchors onto `tc` (supporting plain <w:r> AND <w:ins>/<w:del> diff runs)
    with fresh globally unique `w:id`s allocated from `id_allocator`.
    Returns count of re-anchored comments.
    """
    if not anchors or tc is None or target_comments_root is None or id_allocator is None:
        return 0

    paras = tc.findall("w:p", NS)
    if not paras:
        return 0

    new_text = get_cell_plain_text(tc)
    p_bounds = []
    cur_p_off = 0
    for p in paras:
        ptxt = "".join(t.text or "" for t in p.findall(".//w:t", NS))
        p_bounds.append((cur_p_off, cur_p_off + len(ptxt)))
        cur_p_off += len(ptxt) + 1

    resolved = []  # list of (nid_s, s_pos, e_pos)
    cut_points = set()
    for anc in anchors:
        nid = id_allocator.next_id()
        nid_s = str(nid)
        target_comments_root.append(sanitize_comment_element(anc["comment_el"], nid))
        s_pos, e_pos = _locate_anchor_in_rebuilt_text(new_text, anc)
        # Snap away from '\n' inter-paragraph virtual separator if needed
        for idx_p in range(len(p_bounds) - 1):
            sep_idx = p_bounds[idx_p][1]
            if s_pos == sep_idx and e_pos > sep_idx:
                s_pos = p_bounds[idx_p + 1][0]
            if e_pos == sep_idx + 1 and s_pos <= sep_idx:
                e_pos = sep_idx
        if e_pos < s_pos:
            e_pos = s_pos
        resolved.append((nid_s, s_pos, e_pos))
        cut_points.add(s_pos)
        cut_points.add(e_pos)

    def _clone_sub_run(orig_r: ET.Element, sub_txt: str) -> ET.Element:
        nr = ET.Element(f"{W}r")
        rpr = orig_r.find(f"{W}rPr")
        if rpr is not None:
            nr.append(copy.deepcopy(rpr))
        t_el = ET.Element(f"{W}t")
        if sub_txt.startswith(" ") or sub_txt.endswith(" ") or "  " in sub_txt:
            t_el.set(XML_SPACE, "preserve")
        t_el.text = sub_txt
        nr.append(t_el)
        return nr

    def _make_ref_run_clean(nid_s: str, donor_rpr: ET.Element = None) -> ET.Element:
        r_el = _make_comment_ref_run(nid_s)
        if donor_rpr is not None:
            clean_rpr = build_rpr(donor_rpr, highlight=None, shd_fill=None)
            remove_rpr_child(clean_rpr, "rStyle")
            if len(clean_rpr) > 0:
                r_el.insert(0, clean_rpr)
        return r_el

    starts_at = {}      # char_pos -> list of (nid_s, is_zero_len)
    ends_at = {}        # char_pos -> list of nid_s (for non-zero-len)
    for nid_s, s_pos, e_pos in resolved:
        if s_pos == e_pos:
            starts_at.setdefault(s_pos, []).append((nid_s, True))
        else:
            starts_at.setdefault(s_pos, []).append((nid_s, False))
            ends_at.setdefault(e_pos, []).append(nid_s)

    for p_idx, p in enumerate(paras):
        p_start, p_end = p_bounds[p_idx]
        donor_rpr = None
        for r_cand in p.findall(".//w:r", NS):
            if r_cand.find(f"{W}rPr") is not None:
                donor_rpr = r_cand.find(f"{W}rPr")
                break

        # Step 1: Split plain <w:r> and <w:ins> elements at cut_points inside [p_start, p_end]
        split_items = []  # list of (el, vis_start, vis_end, is_ppr)
        cursor = p_start
        for ch in list(p):
            loc = _local_name(ch.tag)
            if loc == "pPr":
                split_items.append((ch, cursor, cursor, True))
                continue
            if loc == "r":
                rtxt = "".join(t.text or "" for t in ch.findall(f"{W}t"))
                if not rtxt:
                    split_items.append((ch, cursor, cursor, False))
                    continue
                r_s, r_e = cursor, cursor + len(rtxt)
                cursor = r_e
                inner_cuts = sorted(c for c in cut_points if r_s < c < r_e)
                if not inner_cuts:
                    split_items.append((ch, r_s, r_e, False))
                else:
                    pts = [r_s] + inner_cuts + [r_e]
                    for k_i in range(len(pts) - 1):
                        a_off, b_off = pts[k_i], pts[k_i + 1]
                        sub_r = _clone_sub_run(ch, rtxt[a_off - r_s : b_off - r_s])
                        split_items.append((sub_r, a_off, b_off, False))
                continue
            if loc == "ins":
                sub_rs = ch.findall(f"{W}r")
                itxt = "".join(t.text or "" for t in ch.findall(f".//{W}t"))
                if not itxt or len(sub_rs) != 1:
                    i_s, i_e = cursor, cursor + len(itxt)
                    cursor = i_e
                    split_items.append((ch, i_s, i_e, False))
                    continue
                i_s, i_e = cursor, cursor + len(itxt)
                cursor = i_e
                inner_cuts = sorted(c for c in cut_points if i_s < c < i_e)
                if not inner_cuts:
                    split_items.append((ch, i_s, i_e, False))
                else:
                    pts = [i_s] + inner_cuts + [i_e]
                    orig_r0 = sub_rs[0]
                    for k_i in range(len(pts) - 1):
                        a_off, b_off = pts[k_i], pts[k_i + 1]
                        ins_attrs = dict(ch.attrib)
                        if k_i > 0:
                            ins_attrs[f"{W}id"] = str(id_allocator.next_id())
                        new_ins = ET.Element(f"{W}ins", ins_attrs)
                        new_ins.append(_clone_sub_run(orig_r0, itxt[a_off - i_s : b_off - i_s]))
                        split_items.append((new_ins, a_off, b_off, False))
                continue
            # <w:del> or other zero-visible-length child
            split_items.append((ch, cursor, cursor, False))

        # Step 2: Rebuild paragraph children inserting Start / End / Reference at exact offsets
        emitted_positions = set()
        new_p_children = []

        def _emit_markers_at(pos: int):
            if pos in emitted_positions:
                return
            emitted_positions.add(pos)
            for nid_s in ends_at.get(pos, ()):
                new_p_children.append(ET.Element(f"{W}commentRangeEnd", {f"{W}id": nid_s}))
                new_p_children.append(_make_ref_run_clean(nid_s, donor_rpr))
            for nid_s, is_zero in starts_at.get(pos, ()):
                new_p_children.append(ET.Element(f"{W}commentRangeStart", {f"{W}id": nid_s}))
                if is_zero:
                    new_p_children.append(ET.Element(f"{W}commentRangeEnd", {f"{W}id": nid_s}))
                    new_p_children.append(_make_ref_run_clean(nid_s, donor_rpr))

        for el, v_s, v_e, is_ppr in split_items:
            if is_ppr:
                new_p_children.append(el)
                continue
            # Only emit markers at v_s once we are about to emit visible text at v_s (or if paragraph has no visible text)
            if v_e > v_s:
                _emit_markers_at(v_s)
            new_p_children.append(el)

        if p_start == p_end:
            _emit_markers_at(p_start)
        _emit_markers_at(p_end)

        for ch in list(p):
            p.remove(ch)
        for ch in new_p_children:
            p.append(ch)

    return len(resolved)


def _remap_and_repair_cell_comments(
    paras: list,
    src_cmap: dict,
    target_comments_root: ET.Element,
    id_allocator: CommentIdAllocator,
) -> None:
    """
    Ensure every comment in `paras`:
      - Drops markers whose old_cid is absent from `src_cmap` (or all markers if `src_cmap` is empty);
      - Allocates a fresh globally unique integer `new_id` per comment instance and appends sanitized <w:comment> to `target_comments_root`;
      - Auto-opens at start of `paras[0]` (after pPr) if half-open (ended in cell without Start);
      - Auto-closes at end of `paras[-1]` if half-open (started in cell without End);
      - Guarantees exact 1 Start -> 1 End -> 1 Reference order for every `new_id`.
    """
    open_inst = {}           # old_cid -> (new_id_str, saw_text_since_start)
    waiting_ref = {}         # old_cid -> new_id_str
    completed_old = set()    # old_cids that have completed at least one triplet in this cell
    needs_ref_after_end = {} # new_id_str -> (p_el, end_el)

    def _alloc_and_copy(old_cid: str) -> str:
        nid = id_allocator.next_id()
        nid_s = str(nid)
        target_comments_root.append(sanitize_comment_element(src_cmap[old_cid], nid))
        return nid_s

    def _make_ref_with_rpr(nid_s: str, donor_p: ET.Element) -> ET.Element:
        r_el = _make_comment_ref_run(nid_s)
        for existing_r in donor_p.findall(f"{W}r"):
            rpr = existing_r.find(f"{W}rPr")
            if rpr is not None:
                r_el.insert(0, copy.deepcopy(rpr))
                break
        return r_el

    prepend_starts = []  # new_id_str list to insert at start of paras[0] (in order)

    for p_el in paras:
        new_children = []
        for ch in list(p_el):
            loc = _local_name(ch.tag)
            if loc == "pPr":
                new_children.append(ch)
                continue
            if loc == "commentRangeStart":
                old_cid = ch.attrib.get(f"{W}id", "").strip()
                if old_cid not in src_cmap:
                    continue
                if old_cid in open_inst:
                    prev_nid, saw_txt = open_inst[old_cid]
                    if not saw_txt:
                        # Duplicate adjacent commentRangeStart with same ID -> ignore duplicate
                        continue
                    # Close previous unclosed instance before opening a new one
                    end_el = ET.Element(f"{W}commentRangeEnd", {f"{W}id": prev_nid})
                    new_children.append(end_el)
                    new_children.append(_make_ref_with_rpr(prev_nid, p_el))
                    needs_ref_after_end.pop(prev_nid, None)
                    completed_old.add(old_cid)
                nid_s = _alloc_and_copy(old_cid)
                open_inst[old_cid] = (nid_s, False)
                waiting_ref.pop(old_cid, None)
                new_children.append(ET.Element(f"{W}commentRangeStart", {f"{W}id": nid_s}))
                continue

            if loc == "commentRangeEnd":
                old_cid = ch.attrib.get(f"{W}id", "").strip()
                if old_cid not in src_cmap:
                    continue
                if old_cid in open_inst:
                    nid_s, _ = open_inst.pop(old_cid)
                elif old_cid in waiting_ref or old_cid in completed_old:
                    # Stray duplicate commentRangeEnd -> drop
                    continue
                else:
                    # Half-open comment that started BEFORE this cell!
                    nid_s = _alloc_and_copy(old_cid)
                    prepend_starts.append(nid_s)
                end_el = ET.Element(f"{W}commentRangeEnd", {f"{W}id": nid_s})
                new_children.append(end_el)
                waiting_ref[old_cid] = nid_s
                needs_ref_after_end[nid_s] = (p_el, end_el)
                continue

            if loc == "r" and ch.find(f"{W}commentReference") is not None and ch.find(f"{W}t") is None:
                cref = ch.find(f"{W}commentReference")
                old_cid = cref.attrib.get(f"{W}id", "").strip()
                if old_cid not in src_cmap:
                    continue
                if old_cid in waiting_ref:
                    nid_s = waiting_ref.pop(old_cid)
                    cref.set(f"{W}id", nid_s)
                    needs_ref_after_end.pop(nid_s, None)
                    completed_old.add(old_cid)
                    new_children.append(ch)
                elif old_cid in open_inst:
                    nid_s, _ = open_inst.pop(old_cid)
                    end_el = ET.Element(f"{W}commentRangeEnd", {f"{W}id": nid_s})
                    new_children.append(end_el)
                    cref.set(f"{W}id", nid_s)
                    completed_old.add(old_cid)
                    new_children.append(ch)
                elif old_cid not in completed_old:
                    nid_s = _alloc_and_copy(old_cid)
                    prepend_starts.append(nid_s)
                    end_el = ET.Element(f"{W}commentRangeEnd", {f"{W}id": nid_s})
                    new_children.append(end_el)
                    cref.set(f"{W}id", nid_s)
                    completed_old.add(old_cid)
                    new_children.append(ch)
                # else: duplicate commentReference -> drop
                continue

            # Regular run / element -> mark open comments as having seen content
            if open_inst:
                for k_cid in list(open_inst.keys()):
                    open_inst[k_cid] = (open_inst[k_cid][0], True)
            new_children.append(ch)

        for ch in list(p_el):
            p_el.remove(ch)
        for ch in new_children:
            p_el.append(ch)

    # Prepend auto-opened Start markers at beginning of paras[0] (immediately after pPr)
    if prepend_starts and paras:
        p0 = paras[0]
        for nid_s in reversed(prepend_starts):
            _insert_after_ppr(p0, ET.Element(f"{W}commentRangeStart", {f"{W}id": nid_s}))

    # Insert missing commentReference runs for any End that didn't have a Reference
    for nid_s, (p_owner, end_el) in list(needs_ref_after_end.items()):
        children_list = list(p_owner)
        if end_el in children_list:
            idx_e = children_list.index(end_el)
            p_owner.insert(idx_e + 1, _make_ref_with_rpr(nid_s, p_owner))

    # Auto-close any still-open commentRangeStart at the end of paras[-1]
    if open_inst and paras:
        plast = paras[-1]
        for _, (nid_s, _) in open_inst.items():
            plast.append(ET.Element(f"{W}commentRangeEnd", {f"{W}id": nid_s}))
            plast.append(_make_ref_with_rpr(nid_s, plast))


_STANDALONE_TRACK_CHANGE_TAGS = frozenset({
    "ins", "del",
    "rPrChange", "pPrChange", "tblPrChange", "trPrChange", "tcPrChange",
    "sectPrChange", "tblGridChange", "numberingChange",
    "moveFrom", "moveTo", "cellIns", "cellDel", "cellMerge",
})


def sanitize_document_ooxml(root: ET.Element, comments_root: ET.Element = None):
    """
    Whole-document OOXML integrity pass run before saving any .docx:
      1. Strip invalid (>= 0x80000000 or <= 0) or duplicate w14:paraId / w14:textId attributes.
      2. Remove orphaned, half-open, reversed, duplicate, or hidden/internal (`_*` such as `_GoBack`,
         `_Hlk*`, `_Toc*`, `_3dy6vkm`) w:bookmarkStart / w:bookmarkEnd elements.
      3. Repair/clean w:commentRangeStart / w:commentRangeEnd / w:commentReference in `root` and prune
         any unreferenced or duplicate <w:comment> elements in `comments_root`.
      4. Enforce global uniqueness across ALL CT_Markup `w:id` attributes (comments, remaining user bookmarks,
         ins, del, *PrChange, moveFrom/To), re-numbering any colliding bookmarks or revisions!
    Returns sanitized `comments_root` (or None).
    """
    # 1. Clean w14:paraId / w14:textId across root and comments_root
    seen_para_ids = set()
    trees = [root] + ([comments_root] if comments_root is not None else [])
    for tree in trees:
        for elem in tree.iter():
            for attr_k in list(elem.attrib.keys()):
                aloc = _local_name(attr_k)
                if aloc == "paraId":
                    val = elem.attrib[attr_k].strip()
                    try:
                        iv = int(val, 16)
                    except ValueError:
                        iv = -1
                    norm_hex = val.upper()
                    if len(val) != 8 or iv <= 0 or iv >= 0x80000000 or norm_hex in seen_para_ids:
                        del elem.attrib[attr_k]
                    else:
                        seen_para_ids.add(norm_hex)
                elif aloc == "textId":
                    val = elem.attrib[attr_k].strip()
                    try:
                        iv = int(val, 16)
                    except ValueError:
                        iv = -1
                    if len(val) != 8 or iv <= 0 or iv >= 0x80000000:
                        del elem.attrib[attr_k]

    # Build parent map for root
    parent_map = {child: parent for parent in root.iter() for child in parent}

    def _remove_elem(el: ET.Element):
        par = parent_map.get(el)
        if par is not None and el in list(par):
            par.remove(el)
            if _local_name(par.tag) == "r" and len(par) == 0:
                _remove_elem(par)
            elif _local_name(par.tag) == "r" and all(_local_name(x.tag) == "rPr" for x in par):
                _remove_elem(par)

    # 2. Validate bookmarkStart / bookmarkEnd pairs across root in document order
    bm_starts = {}   # id -> list[(order_idx, elem, name)]
    bm_ends = {}     # id -> list[(order_idx, elem)]
    order_counter = 0
    for elem in list(root.iter()):
        loc = _local_name(elem.tag)
        if loc == "bookmarkStart":
            bid = elem.attrib.get(f"{W}id", "").strip()
            bname = elem.attrib.get(f"{W}name", "").strip()
            bm_starts.setdefault(bid, []).append((order_counter, elem, bname))
            order_counter += 1
        elif loc == "bookmarkEnd":
            bid = elem.attrib.get(f"{W}id", "").strip()
            bm_ends.setdefault(bid, []).append((order_counter, elem))
            order_counter += 1

    used_bm_names = set()
    valid_bm_ids = set()
    for bid, s_list in bm_starts.items():
        e_list = bm_ends.get(bid, [])
        if not bid or len(s_list) != 1 or len(e_list) != 1:
            continue
        s_ord, _, bname = s_list[0]
        e_ord, _ = e_list[0]
        # Strip ALL hidden/internal Word/GoogleDocs bookmarks starting with '_' (_GoBack, _Hlk*, _Toc*, _3dy6vkm, etc.)
        if not bname or bname.startswith("_") or s_ord >= e_ord or bname in used_bm_names:
            continue
        try:
            int(bid)
        except ValueError:
            continue
        used_bm_names.add(bname)
        valid_bm_ids.add(bid)

    for bid, s_list in bm_starts.items():
        if bid not in valid_bm_ids:
            for _, el, _ in s_list:
                _remove_elem(el)
    for bid, e_list in bm_ends.items():
        if bid not in valid_bm_ids:
            for _, el in e_list:
                _remove_elem(el)

    # 3. Validate & synchronize comments between root and comments_root
    available_comments = {}
    if comments_root is not None:
        for c_el in list(comments_root.findall(f"{W}comment")):
            cid = c_el.attrib.get(f"{W}id", "").strip()
            try:
                cid_int = int(cid)
            except ValueError:
                comments_root.remove(c_el)
                continue
            cid_norm = str(cid_int)
            if cid_norm in available_comments:
                comments_root.remove(c_el)
            else:
                available_comments[cid_norm] = sanitize_comment_element(c_el, cid_int)

    # Refresh parent_map after bookmark cleanup
    parent_map = {child: parent for parent in root.iter() for child in parent}

    c_starts = {}  # cid -> list[(ord, elem)]
    c_ends = {}    # cid -> list[(ord, elem)]
    c_refs = {}    # cid -> list[(ord, elem)]
    ord_c = 0
    for elem in list(root.iter()):
        loc = _local_name(elem.tag)
        if loc in ("commentRangeStart", "commentRangeEnd", "commentReference"):
            cid = elem.attrib.get(f"{W}id", "").strip()
            if cid not in available_comments:
                _remove_elem(elem)
                continue
            if loc == "commentRangeStart":
                c_starts.setdefault(cid, []).append((ord_c, elem))
            elif loc == "commentRangeEnd":
                c_ends.setdefault(cid, []).append((ord_c, elem))
            else:
                c_refs.setdefault(cid, []).append((ord_c, elem))
            ord_c += 1

    all_doc_cids = set(c_starts.keys()) | set(c_ends.keys()) | set(c_refs.keys())
    referenced_cids = set()

    for cid in all_doc_cids:
        s_list = c_starts.get(cid, [])
        e_list = c_ends.get(cid, [])
        r_list = c_refs.get(cid, [])

        # Deduplicate extra markers keeping first valid occurrence
        for _, dup_el in s_list[1:]:
            _remove_elem(dup_el)
        s_list = s_list[:1]
        for _, dup_el in e_list[1:]:
            _remove_elem(dup_el)
        e_list = e_list[:1]
        for _, dup_el in r_list[1:]:
            _remove_elem(dup_el)
        r_list = r_list[:1]

        # Find an anchor <w:p> if repair is needed
        def _enclosing_p(el: ET.Element):
            cur = el
            while cur is not None:
                if _local_name(cur.tag) == "p":
                    return cur
                cur = parent_map.get(cur)
            return None

        if len(s_list) == 1 and len(e_list) == 1 and s_list[0][0] < e_list[0][0]:
            if len(r_list) == 0:
                p_end = _enclosing_p(e_list[0][1])
                if p_end is not None and e_list[0][1] in list(p_end):
                    idx_e = list(p_end).index(e_list[0][1])
                    p_end.insert(idx_e + 1, _make_comment_ref_run(cid))
                    referenced_cids.add(cid)
                    continue
            else:
                referenced_cids.add(cid)
                continue

        # Half-open or out-of-order comment remaining in root (e.g. cross-row comment after row deletion)
        anchor_el = (s_list[0][1] if s_list else (e_list[0][1] if e_list else (r_list[0][1] if r_list else None)))
        anchor_p = _enclosing_p(anchor_el) if anchor_el is not None else None
        for _, el in s_list + e_list + r_list:
            _remove_elem(el)
        if anchor_p is not None:
            _insert_after_ppr(anchor_p, ET.Element(f"{W}commentRangeStart", {f"{W}id": cid}))
            anchor_p.append(ET.Element(f"{W}commentRangeEnd", {f"{W}id": cid}))
            anchor_p.append(_make_comment_ref_run(cid))
            referenced_cids.add(cid)

    if comments_root is not None:
        for ch in list(comments_root):
            comments_root.remove(ch)
        for cid in sorted(referenced_cids, key=lambda x: int(x)):
            comments_root.append(available_comments[cid])

    # 4. Enforce global CT_Markup w:id uniqueness across comments, remaining user bookmarks, and revisions
    max_id = 0
    for tree in (root, comments_root):
        if tree is None:
            continue
        for elem in tree.iter():
            v_str = elem.attrib.get(f"{W}id", "").strip()
            if v_str:
                try:
                    iv = int(v_str)
                    if iv > max_id:
                        max_id = iv
                except ValueError:
                    pass
    next_global_id = max_id + 1
    used_global_ids = {int(cid) for cid in referenced_cids}

    # Register standalone track-change IDs first (deduplicating any colliding track-change IDs),
    # then re-number any remaining valid user bookmarks whose w:id collides with anything!
    for elem in root.iter():
        loc = _local_name(elem.tag)
        if loc in _STANDALONE_TRACK_CHANGE_TAGS:
            raw_id = elem.attrib.get(f"{W}id", "").strip()
            needs_new = False
            try:
                iv = int(raw_id)
                if iv < 0 or iv in used_global_ids:
                    needs_new = True
            except ValueError:
                needs_new = True
            if needs_new:
                new_iv = next_global_id
                next_global_id += 1
                elem.set(f"{W}id", str(new_iv))
                used_global_ids.add(new_iv)
            else:
                used_global_ids.add(iv)

    for bid in sorted(valid_bm_ids, key=lambda x: int(x)):
        bid_int = int(bid)
        s_el = bm_starts[bid][0][1]
        e_el = bm_ends[bid][0][1]
        if bid_int < 0 or bid_int in used_global_ids:
            new_bid = next_global_id
            next_global_id += 1
            s_el.set(f"{W}id", str(new_bid))
            e_el.set(f"{W}id", str(new_bid))
            used_global_ids.add(new_bid)
        else:
            used_global_ids.add(bid_int)

    return comments_root


def _patch_document_rels_xml(rels_bytes: bytes, ensure_comments_rel: bool, keep_people: bool = False) -> bytes:
    """Update word/_rels/document.xml.rels: strip extended comment / orphaned people rels and sync comments.xml rel."""
    _assert_safe_openxml_bytes(rels_bytes)
    rels_root = ET.fromstring(rels_bytes)
    kept = []
    used_ids = set()
    has_comments = False

    for rel in list(rels_root):
        if _local_name(rel.tag) != "Relationship":
            continue
        rid = rel.attrib.get("Id", "").strip()
        rtype = rel.attrib.get("Type", "").strip()
        target = rel.attrib.get("Target", "").strip()
        target_low = target.replace("\\", "/").lower()
        rtype_low = rtype.lower()

        if any(target_low.endswith(sfx) for sfx in _STRIPPED_COMMENT_PART_SUFFIXES):
            continue
        if any(rtype.endswith(rt_sfx) for rt_sfx in _STRIPPED_COMMENT_REL_TYPES):
            continue
        if not keep_people and (target_low.endswith("people.xml") or rtype_low.endswith("/people")):
            continue

        is_cmt = (rtype == COMMENTS_REL_TYPE or target_low.endswith("comments.xml"))
        if is_cmt:
            if not ensure_comments_rel or has_comments:
                continue
            has_comments = True
            rel.set("Type", COMMENTS_REL_TYPE)
            rel.set("Target", "comments.xml")

        if rid:
            used_ids.add(rid)
        kept.append(rel)

    if ensure_comments_rel and not has_comments:
        n = 1
        while f"rId{n}" in used_ids:
            n += 1
        new_rel = ET.Element(f"{{{PKG_REL_URI}}}Relationship", {
            "Id": f"rId{n}",
            "Type": COMMENTS_REL_TYPE,
            "Target": "comments.xml",
        })
        kept.append(new_rel)

    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n',
        f'<Relationships xmlns="{PKG_REL_URI}">',
    ]
    for rel in kept:
        attr_str = " ".join(f"{_local_name(k)}={quoteattr(v)}" for k, v in rel.attrib.items())
        parts.append(f"<Relationship {attr_str}/>")
    parts.append("</Relationships>")
    return "".join(parts).encode("utf-8")


def _patch_content_types_xml(ct_bytes: bytes, ensure_comments_ct: bool, keep_people: bool = False) -> bytes:
    """Update [Content_Types].xml: strip extended comment / orphaned people overrides and sync /word/comments.xml Override."""
    _assert_safe_openxml_bytes(ct_bytes)
    ct_root = ET.fromstring(ct_bytes)
    defaults = []
    overrides = []
    has_comments_ov = False

    for ch in list(ct_root):
        loc = _local_name(ch.tag)
        if loc == "Default":
            defaults.append(ch)
        elif loc == "Override":
            pname = ch.attrib.get("PartName", "").strip()
            pname_low = pname.replace("\\", "/").lower()
            if any(pname_low.endswith(sfx) for sfx in _STRIPPED_COMMENT_PART_SUFFIXES):
                continue
            if not keep_people and pname_low.endswith("/people.xml"):
                continue
            if pname_low == "/word/comments.xml":
                if not ensure_comments_ct or has_comments_ov:
                    continue
                has_comments_ov = True
                ch.set("PartName", "/word/comments.xml")
                ch.set("ContentType", COMMENTS_CONTENT_TYPE)
            overrides.append(ch)

    if ensure_comments_ct and not has_comments_ov:
        overrides.append(ET.Element(f"{{{CONTENT_TYPES_URI}}}Override", {
            "PartName": "/word/comments.xml",
            "ContentType": COMMENTS_CONTENT_TYPE,
        }))

    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n',
        f'<Types xmlns="{CONTENT_TYPES_URI}">',
    ]
    for d_el in defaults:
        attr_str = " ".join(f"{_local_name(k)}={quoteattr(v)}" for k, v in d_el.attrib.items())
        parts.append(f"<Default {attr_str}/>")
    for o_el in overrides:
        attr_str = " ".join(f"{_local_name(k)}={quoteattr(v)}" for k, v in o_el.attrib.items())
        parts.append(f"<Override {attr_str}/>")
    parts.append("</Types>")
    return "".join(parts).encode("utf-8")


def _serialize_comments_xml(comments_root: ET.Element) -> bytes:
    with _ET_NS_LOCK:
        ET.register_namespace("w", W_URI)
        ET.register_namespace("r", R_URI)
        xml_str = ET.tostring(comments_root, encoding="unicode")
    m = re.match(r"<w:comments\b[^>]*>", xml_str)
    if m:
        cur_open = m.group(0)
        if f'xmlns:w="{W_URI}"' not in cur_open:
            if cur_open.endswith("/>"):
                xml_str = cur_open[:-2] + f' xmlns:w="{W_URI}"/>'
            else:
                xml_str = cur_open[:-1] + f' xmlns:w="{W_URI}">' + xml_str[m.end():]
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml_str).encode("utf-8")


def save_docx_xml(
    src_docx_path: Path,
    dst_docx_path: Path,
    root: ET.Element,
    ns_pairs,
    comments_root: ET.Element = None,
):
    if comments_root is None:
        comments_root = load_docx_comments_xml(src_docx_path)

    comments_root = sanitize_document_ooxml(root, comments_root)
    has_active_comments = comments_root is not None and len(comments_root.findall(f"{W}comment")) > 0

    with _ET_NS_LOCK:
        for prefix, uri in ns_pairs:
            if prefix != "xml":
                ET.register_namespace(prefix, uri)
        ET.register_namespace("w", W_URI)
        ET.register_namespace("r", R_URI)
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

    full_doc_bytes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml_str).encode("utf-8")
    comments_bytes = _serialize_comments_xml(comments_root) if has_active_comments else None

    # Support saving when src_docx_path == dst_docx_path safely by buffering if needed
    same_file = src_docx_path.resolve() == dst_docx_path.resolve()
    tmp_dst = dst_docx_path.with_suffix(".docx.tmp_save") if same_file else dst_docx_path
    dst_docx_path.parent.mkdir(parents=True, exist_ok=True)

    wrote_comments = False
    wrote_rels = False
    wrote_ct = False

    with zipfile.ZipFile(src_docx_path, "r") as zin:
        _validate_docx_zip(zin)
        zin_names = {it.filename.replace("\\", "/") for it in zin.infolist()}
        keep_people = has_active_comments and ("word/people.xml" in zin_names)
        with zipfile.ZipFile(tmp_dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                norm_name = item.filename.replace("\\", "/")
                if norm_name in _STRIPPED_COMMENT_ZIP_NAMES or any(
                    norm_name.lower().endswith(sfx) for sfx in _STRIPPED_COMMENT_PART_SUFFIXES
                ):
                    continue
                if not keep_people and (
                    norm_name in _PEOPLE_ZIP_NAMES or norm_name.lower().endswith("/people.xml")
                ):
                    continue
                if norm_name == "word/document.xml":
                    zout.writestr(item, full_doc_bytes)
                elif norm_name == "word/comments.xml":
                    if has_active_comments and comments_bytes is not None:
                        zout.writestr(item, comments_bytes)
                        wrote_comments = True
                elif norm_name == "word/_rels/document.xml.rels":
                    with zin.open(item, "r") as mfp:
                        raw_rels = mfp.read(MAX_DOCX_MEMBER_BYTES + 1)
                    zout.writestr(
                        item,
                        _patch_document_rels_xml(raw_rels, ensure_comments_rel=has_active_comments, keep_people=keep_people),
                    )
                    wrote_rels = True
                elif norm_name == "[Content_Types].xml":
                    with zin.open(item, "r") as mfp:
                        raw_ct = mfp.read(MAX_DOCX_MEMBER_BYTES + 1)
                    zout.writestr(
                        item,
                        _patch_content_types_xml(raw_ct, ensure_comments_ct=has_active_comments, keep_people=keep_people),
                    )
                    wrote_ct = True
                else:
                    with zin.open(item, "r") as mfp:
                        data = mfp.read(MAX_DOCX_MEMBER_BYTES + 1)
                    if len(data) > MAX_DOCX_MEMBER_BYTES:
                        raise ValueError(f"Member {item.filename!r} exceeded actual read byte limit.")
                    zout.writestr(item, data)

            if has_active_comments and not wrote_comments and comments_bytes is not None:
                zout.writestr("word/comments.xml", comments_bytes)
            if has_active_comments and not wrote_rels:
                empty_rels = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="{PKG_REL_URI}"/>'.encode("utf-8")
                zout.writestr(
                    "word/_rels/document.xml.rels",
                    _patch_document_rels_xml(empty_rels, ensure_comments_rel=True, keep_people=keep_people),
                )
            if has_active_comments and not wrote_ct:
                empty_ct = (
                    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                    f'<Types xmlns="{CONTENT_TYPES_URI}">'
                    f'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                    f'<Default Extension="xml" ContentType="application/xml"/>'
                    f'<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                    f'</Types>'
                ).encode("utf-8")
                zout.writestr(
                    "[Content_Types].xml",
                    _patch_content_types_xml(empty_ct, ensure_comments_ct=True, keep_people=keep_people),
                )

    if same_file:
        tmp_dst.replace(dst_docx_path)


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
    """Exact MS Word word count formula verified against test ground truth (RV 609 / UP 452). Punctuation excluded."""
    tokens = re.split(r"[\s—–]+", text.strip())
    return sum(1 for tok in tokens if re.search(r"[A-Za-z0-9]", tok))


RV_HIGHLIGHT_ALIASES = frozenset({"cyan", "blue", "darkCyan"})
_WORD_COUNT_SKIP_TAGS = frozenset({
    "del", "moveFrom", "rPrChange", "pPrChange", "tblPrChange", "trPrChange", "tcPrChange"
})


def compute_col2_up_rv_words(root: ET.Element, up_color: str = "yellow", rv_color: str = "cyan") -> tuple[int, int]:
    """
    Count English words (punctuation excluded; words, not characters) in Column 2 (index 1)
    across all tables in `root` (excluding skipped rows marked in Column 3):
      - UP = yellow highlight (`up_color`)
      - RV = cyan highlight (`rv_color` / cyan / blue / darkCyan)
    Deleted track-change runs (<w:del>) are strictly ignored.
    """
    rv_vals = {rv_color} | RV_HIGHLIGHT_ALIASES
    up_chunks: list[str] = []
    rv_chunks: list[str] = []

    def _walk_runs(node: ET.Element, u_buf: list[str], r_buf: list[str]) -> None:
        for ch in node:
            local = ch.tag.split("}")[-1]
            if local in _WORD_COUNT_SKIP_TAGS:
                continue
            if local == "r":
                hl = ch.find("./w:rPr/w:highlight", NS)
                hval = hl.attrib.get(f"{W}val") if hl is not None else None
                parts = []
                for sub in ch:
                    slocal = sub.tag.split("}")[-1]
                    if slocal == "t" and sub.text:
                        parts.append(sub.text)
                    elif slocal in ("tab", "br", "cr"):
                        parts.append(" ")
                wt = "".join(parts)
                if not wt:
                    continue
                if hval == up_color:
                    u_buf.append(wt)
                elif u_buf:
                    u_buf.append(" ")
                if hval in rv_vals:
                    r_buf.append(wt)
                elif r_buf:
                    r_buf.append(" ")
            else:
                _walk_runs(ch, u_buf, r_buf)

    for tbl in root.findall(".//w:tbl", NS):
        for tr in tbl.findall("w:tr", NS):
            tcs = tr.findall("w:tc", NS)
            if len(tcs) < 2:
                continue
            if len(tcs) >= 3 and get_cell_plain_text(tcs[2]).strip() == SKIPPED_ROW_MARKER:
                continue
            tc2 = tcs[1]
            for p in tc2.findall("w:p", NS):
                u_buf: list[str] = []
                r_buf: list[str] = []
                _walk_runs(p, u_buf, r_buf)
                if u_buf:
                    up_chunks.append("".join(u_buf))
                if r_buf:
                    rv_chunks.append("".join(r_buf))

    return count_msword_en_words(" ".join(up_chunks)), count_msword_en_words(" ".join(rv_chunks))


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
    WITHOUT adding extra line breaks/paragraph splits, preserving original run rPr formatting
    and any Word comment markers (<w:commentRangeStart>, <w:commentRangeEnd>, <w:commentReference>).
    """
    if not char_spans:
        return
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
            extra_crefs = [
                sub.attrib.get(f"{W}id", "").strip()
                for sub in ch.findall("w:commentReference", NS)
                if sub.attrib.get(f"{W}id", "").strip()
            ]
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
            for cid in extra_crefs:
                cref_r = _make_comment_ref_run(cid)
                if base_rPr is not None:
                    cref_r.insert(0, build_rpr(base_rPr, highlight=None, shd_fill=None))
                new_children.append(cref_r)

        for ch in list(p):
            p.remove(ch)
        for ch in new_children:
            p.append(ch)


# ==============================================================================
# v1.7.0 Public Helpers: 2->3 Col Expansion, Duplicate Row Merge, Column Balance,
# Naming Tags, Full-Article English Title Extraction, & Top Instruction Lines
# ==============================================================================

_STRAY_SPAN_TAGS = ("gridSpan", "gridBefore", "gridAfter", "wBefore", "wAfter")
_TOP_INSTRUCTION_NORMS = frozenset({
    "blue highlighted parts need to be reviewed",
    "yellow highlighted parts need to be updated",
})
_CJK_WEIGHT_RE = re.compile(r"[\u3000-\u303f\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uac00-\ud7af\uff00-\uffef]")
_WS_COLLAPSE_RE = re.compile(r"[\s\u3000\u200b\u200c\ufeff]+")


def _strip_stray_grid_spans_on_row(tr: ET.Element) -> bool:
    changed = False
    trPr = tr.find(f"{W}trPr")
    if trPr is not None:
        for tname in _STRAY_SPAN_TAGS:
            for el in list(trPr.findall(f"{W}{tname}")):
                trPr.remove(el)
                changed = True
    for tc in tr.findall(f"{W}tc"):
        tcPr = tc.find(f"{W}tcPr")
        if tcPr is not None:
            for tname in _STRAY_SPAN_TAGS:
                for el in list(tcPr.findall(f"{W}{tname}")):
                    tcPr.remove(el)
                    changed = True
    return changed


def _ensure_landscape_a4_on_root(root: ET.Element) -> bool:
    changed = False
    body = root.find(f"{W}body")
    sect_prs = root.findall(f".//{W}sectPr")
    if not sect_prs and body is not None:
        sp = ET.SubElement(body, f"{W}sectPr")
        sect_prs = [sp]
        changed = True

    for sp in sect_prs:
        pgSz = sp.find(f"{W}pgSz")
        if pgSz is None:
            insert_sectpr_child_ordered(sp, ET.Element(f"{W}pgSz", {
                f"{W}w": "16838",
                f"{W}h": "11906",
                f"{W}orient": "landscape",
            }))
            changed = True
        else:
            try:
                cur_w = int(pgSz.attrib.get(f"{W}w", "0") or 0)
            except ValueError:
                cur_w = 0
            try:
                cur_h = int(pgSz.attrib.get(f"{W}h", "0") or 0)
            except ValueError:
                cur_h = 0
            if cur_w <= 0 or cur_h <= 0 or cur_w < cur_h:
                pgSz.set(f"{W}w", "16838")
                pgSz.set(f"{W}h", "11906")
                pgSz.set(f"{W}orient", "landscape")
                changed = True
    return changed


def ensure_three_column_tables_and_landscape(root: ET.Element) -> int:
    """
    1. Sets <w:sectPr>/<w:pgSz> to Landscape A4 (w="16838" h="11906" orient="landscape") whenever width < height.
    2. Strips stray <w:gridSpan> / <w:gridBefore> / <w:wBefore> on 2-cell or 3-cell rows.
    3. Expands any 2-column <w:tbl> to 3 <w:gridCol>s (3057, 5533, 5970) and appends a 3rd <w:tc>
       (copying <w:tcBorders> from tc[1] + an empty <w:p>) to every 2-cell <w:tr>.
    Returns count of expanded tables.
    """
    _ensure_landscape_a4_on_root(root)

    expanded_tables = 0
    for tbl in root.findall(f".//{W}tbl"):
        trs = tbl.findall(f"{W}tr")
        if not trs:
            continue
        has_2cell = False
        for tr in trs:
            tcs = tr.findall(f"{W}tc")
            if len(tcs) in (2, 3):
                _strip_stray_grid_spans_on_row(tr)
            if len(tcs) == 2:
                has_2cell = True

        if not has_2cell:
            continue

        # Ensure tblPr with 5000 pct width
        tblPr = tbl.find(f"{W}tblPr")
        if tblPr is None:
            tblPr = ET.Element(f"{W}tblPr")
            tbl.insert(0, tblPr)
        insert_tblpr_child_ordered(tblPr, ET.Element(f"{W}tblW", {f"{W}w": "5000", f"{W}type": "pct"}))

        # Rebuild tblGrid with 3 gridCols: 3057, 5533, 5970
        old_grid = tbl.find(f"{W}tblGrid")
        new_grid = ET.Element(f"{W}tblGrid")
        for gw in ("3057", "5533", "5970"):
            ET.SubElement(new_grid, f"{W}gridCol", {f"{W}w": gw})
        if old_grid is not None:
            idx_g = list(tbl).index(old_grid)
            tbl.remove(old_grid)
            tbl.insert(idx_g, new_grid)
        else:
            idx_pr = list(tbl).index(tblPr)
            tbl.insert(idx_pr + 1, new_grid)

        pct_widths = ("1050", "1900", "2050")
        for tr in trs:
            tcs = tr.findall(f"{W}tc")
            if len(tcs) != 2:
                continue
            for col_i in (0, 1):
                tcPr_i = tcs[col_i].find(f"{W}tcPr")
                if tcPr_i is None:
                    tcPr_i = ET.Element(f"{W}tcPr")
                    tcs[col_i].insert(0, tcPr_i)
                insert_tcpr_child_ordered(
                    tcPr_i,
                    ET.Element(f"{W}tcW", {f"{W}w": pct_widths[col_i], f"{W}type": "pct"}),
                )

            tc1_pr = tcs[1].find(f"{W}tcPr")
            new_tc = ET.Element(f"{W}tc")
            new_tcPr = ET.Element(f"{W}tcPr")
            insert_tcpr_child_ordered(new_tcPr, ET.Element(f"{W}tcW", {f"{W}w": pct_widths[2], f"{W}type": "pct"}))
            if tc1_pr is not None:
                for copy_tag in ("tcBorders", "tcMar", "vAlign"):
                    sub_el = tc1_pr.find(f"{W}{copy_tag}")
                    if sub_el is not None:
                        cp_sub = copy.deepcopy(sub_el)
                        _strip_unsafe_attrs_recursive(cp_sub)
                        insert_tcpr_child_ordered(new_tcPr, cp_sub)
            new_tc.append(new_tcPr)
            new_tc.append(ET.Element(f"{W}p"))
            tr.append(new_tc)

        expanded_tables += 1

    return expanded_tables


def _normalize_shd_fill_val(shd_el: ET.Element = None):
    if shd_el is None:
        return None
    fill = (shd_el.attrib.get(f"{W}fill") or "").strip()
    if not fill or fill.lower() in ("auto", "ffffff", "none"):
        return None
    return fill.upper()


def _normalize_hl_val(hl_el: ET.Element = None):
    if hl_el is None:
        return None
    hval = (hl_el.attrib.get(f"{W}val") or "").strip()
    if not hval or hval.lower() in ("none", "white"):
        return None
    return hval


def _normalize_color_val(col_el: ET.Element = None):
    if col_el is None:
        return None
    cval = (col_el.attrib.get(f"{W}val") or "").strip()
    if not cval or cval.lower() in ("auto", "000000", "black", "ffffff"):
        return None
    return cval.upper()


def _extract_cell_char_visual_props(tc: ET.Element):
    """Return (text, list_of_(shd_fill, hl_val, color_val)_per_char)."""
    chars = []
    props = []
    tcPr = tc.find(f"{W}tcPr")
    tc_shd = _normalize_shd_fill_val(tcPr.find(f"{W}shd") if tcPr is not None else None)

    for p_i, p in enumerate(tc.findall(f"{W}p")):
        if p_i > 0:
            chars.append("\n")
            props.append((None, None, None))
        pPr = p.find(f"{W}pPr")
        p_shd = _normalize_shd_fill_val(pPr.find(f"{W}shd") if pPr is not None else None) or tc_shd
        for r in p.findall(f".//{W}r"):
            txt = "".join(t.text or "" for t in r.findall(f"{W}t"))
            if not txt:
                continue
            rPr = r.find(f"{W}rPr")
            r_shd = _normalize_shd_fill_val(rPr.find(f"{W}shd") if rPr is not None else None) or p_shd
            r_hl = _normalize_hl_val(rPr.find(f"{W}highlight") if rPr is not None else None)
            r_col = _normalize_color_val(rPr.find(f"{W}color") if rPr is not None else None)
            prop_tuple = (r_shd, r_hl, r_col)
            for ch in txt:
                chars.append(ch)
                props.append(prop_tuple)
    return "".join(chars), props


def _rebuild_cell_with_visual_props(tc: ET.Element, full_text: str, char_props: list) -> None:
    """Rebuild paragraphs/runs of `tc` preserving its template pPr & base_rPr with per-char (shd, hl, color)."""
    base_rPr = get_first_run_rpr(tc)
    remove_rpr_child(base_rPr, "color")
    orig_pPrs = [
        copy.deepcopy(p_item.find(f"{W}pPr")) if p_item.find(f"{W}pPr") is not None else None
        for p_item in tc.findall(f"{W}p")
    ]
    for ppr in orig_pPrs:
        if ppr is not None:
            remove_ppr_child(ppr, "shd")

    para_slices = [[]]
    for ch, prop in zip(full_text, char_props):
        if ch == "\n":
            para_slices.append([])
        else:
            para_slices[-1].append((ch, prop))

    tcPr = tc.find(f"{W}tcPr")
    if tcPr is not None:
        for shd_c in list(tcPr.findall(f"{W}shd")):
            tcPr.remove(shd_c)
    for ch in list(tc):
        if ch is not tcPr:
            tc.remove(ch)

    for p_idx, items in enumerate(para_slices):
        p_el = ET.Element(f"{W}p")
        chosen_pPr = (
            orig_pPrs[p_idx]
            if p_idx < len(orig_pPrs)
            else (orig_pPrs[-1] if orig_pPrs else None)
        )
        if chosen_pPr is not None:
            p_el.append(copy.deepcopy(chosen_pPr))

        if items:
            seg_chars = [items[0][0]]
            cur_prop = items[0][1]
            for ch, prop in items[1:]:
                if prop == cur_prop:
                    seg_chars.append(ch)
                else:
                    shd_f, hl_v, col_v = cur_prop
                    r_el = make_run("".join(seg_chars), base_rPr=base_rPr, highlight=hl_v, shd_fill=shd_f)
                    if col_v:
                        rPr_el = r_el.find(f"{W}rPr")
                        if rPr_el is None:
                            rPr_el = ET.Element(f"{W}rPr")
                            r_el.insert(0, rPr_el)
                        insert_rpr_child_ordered(rPr_el, ET.Element(f"{W}color", {f"{W}val": col_v}))
                    p_el.append(r_el)
                    seg_chars = [ch]
                    cur_prop = prop
            shd_f, hl_v, col_v = cur_prop
            r_el = make_run("".join(seg_chars), base_rPr=base_rPr, highlight=hl_v, shd_fill=shd_f)
            if col_v:
                rPr_el = r_el.find(f"{W}rPr")
                if rPr_el is None:
                    rPr_el = ET.Element(f"{W}rPr")
                    r_el.insert(0, rPr_el)
                insert_rpr_child_ordered(rPr_el, ET.Element(f"{W}color", {f"{W}val": col_v}))
            p_el.append(r_el)

        tc.append(p_el)


def _merge_col1_visual_runs(
    earlier_tc0: ET.Element,
    later_tc0: ET.Element,
    comments_root: ET.Element = None,
    id_allocator: MarkupIdAllocator = None,
) -> bool:
    txt_a, props_a = _extract_cell_char_visual_props(earlier_tc0)
    txt_b, props_b = _extract_cell_char_visual_props(later_tc0)
    if not txt_b or all(p == (None, None, None) for p in props_b):
        return False

    sm = difflib.SequenceMatcher(None, txt_a, txt_b, autojunk=False)
    out_chars = []
    out_props = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                pa = props_a[i1 + k]
                pb = props_b[j1 + k]
                out_chars.append(txt_a[i1 + k])
                out_props.append((pa[0] or pb[0], pa[1] or pb[1], pa[2] or pb[2]))
        elif tag == "replace":
            b_has_style = any(p != (None, None, None) for p in props_b[j1:j2])
            a_has_style = any(p != (None, None, None) for p in props_a[i1:i2])
            if b_has_style and not a_has_style:
                for k in range(j1, j2):
                    out_chars.append(txt_b[k])
                    out_props.append(props_b[k])
            else:
                for k in range(i1, i2):
                    out_chars.append(txt_a[k])
                    out_props.append(props_a[k])
        elif tag == "delete":
            for k in range(i1, i2):
                out_chars.append(txt_a[k])
                out_props.append(props_a[k])
        elif tag == "insert":
            if any(p != (None, None, None) for p in props_b[j1:j2]):
                for k in range(j1, j2):
                    out_chars.append(txt_b[k])
                    out_props.append(props_b[k])

    merged_txt = "".join(out_chars)
    if merged_txt == txt_a and out_props == props_a:
        return False

    anchors = (
        extract_cell_comment_anchors(earlier_tc0, comments_root)
        + extract_cell_comment_anchors(later_tc0, comments_root)
    )
    _rebuild_cell_with_visual_props(earlier_tc0, merged_txt, out_props)
    if anchors and comments_root is not None and id_allocator is not None:
        reanchor_comments_to_cell(earlier_tc0, anchors, comments_root, id_allocator)
    return True


_EN_SENT_END_SPLIT_RE = re.compile(r'[.?!]+(?:[”’"\')\]]+)?')


def _split_col2_into_sentence_units(text: str, props: list) -> list:
    """
    Partition (text, props) into ordered sentence/newline dicts:
      {"is_nl": bool, "text": str, "props": list, "words": list[str], "has_shd": bool, "shd_chars": int}
    """
    units = []
    n = len(text)
    idx = 0

    def _make_sent_unit(s_i: int, e_i: int) -> dict:
        sub_t = text[s_i:e_i]
        sub_p = props[s_i:e_i]
        shd_c = sum(1 for ch, p in zip(sub_t, sub_p) if (not ch.isspace()) and p[0] is not None)
        words = [m.group(0).lower() for m in WORD_EXTRACT_RE.finditer(sub_t)]
        return {
            "is_nl": False,
            "text": sub_t,
            "props": sub_p,
            "words": words,
            "has_shd": shd_c > 0,
            "shd_chars": shd_c,
        }

    while idx < n:
        nl_pos = text.find("\n", idx)
        p_end = nl_pos if nl_pos != -1 else n
        p_sub = text[idx:p_end]
        cur_s = 0
        for m in _EN_SENT_END_SPLIT_RE.finditer(p_sub):
            end_local = m.end()
            if p_sub[cur_s:end_local].strip():
                units.append(_make_sent_unit(idx + cur_s, idx + end_local))
                cur_s = end_local
        if cur_s < len(p_sub):
            if p_sub[cur_s:].strip() or not units or units[-1]["is_nl"]:
                units.append(_make_sent_unit(idx + cur_s, p_end))
            else:
                # Attach trailing spaces at end of paragraph to last sentence
                units[-1]["text"] += p_sub[cur_s:]
                units[-1]["props"].extend(props[idx + cur_s : p_end])
        if nl_pos != -1:
            units.append({
                "is_nl": True,
                "text": "\n",
                "props": [props[nl_pos]],
                "words": [],
                "has_shd": False,
                "shd_chars": 0,
            })
            idx = nl_pos + 1
        else:
            break
    return units


def _sent_similarity(u_a: dict, u_b: dict) -> float:
    if u_a["is_nl"] != u_b["is_nl"]:
        return 0.0
    if u_a["is_nl"]:
        return 1.0
    wa, wb = u_a["words"], u_b["words"]
    if not wa and not wb:
        return 1.0 if u_a["text"].strip() == u_b["text"].strip() else 0.0
    if not wa or not wb:
        return 0.0
    if wa == wb:
        return 1.0
    return difflib.SequenceMatcher(None, wa, wb, autojunk=False).ratio()


def _merge_two_shaded_sentence_units(sa: dict, sb: dict) -> tuple:
    """Merge two aligned sentence units when BOTH contain shading (union spans / prefer shaded edits)."""
    ta, pa = sa["text"], sa["props"]
    tb, pb = sb["text"], sb["props"]
    if ta == tb:
        merged_p = [(x[0] or y[0], x[1] or y[1], x[2] or y[2]) for x, y in zip(pa, pb)]
        return ta, merged_p

    # Token-level alignment to union shaded spans or prefer shaded edits
    toks_a = tokenize_en_fine(ta)
    toks_b = tokenize_en_fine(tb)

    def _slice_tok_props(toks, all_p):
        res = []
        c = 0
        for tk in toks:
            ln = len(tk)
            res.append(all_p[c : c + ln])
            c += ln
        return res

    tprops_a = _slice_tok_props(toks_a, pa)
    tprops_b = _slice_tok_props(toks_b, pb)
    sm = difflib.SequenceMatcher(None, toks_a, toks_b, autojunk=False)
    out_t = []
    out_p = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                tk = toks_a[i1 + k]
                out_t.append(tk)
                for xa, xb in zip(tprops_a[i1 + k], tprops_b[j1 + k]):
                    out_p.append((xa[0] or xb[0], xa[1] or xb[1], xa[2] or xb[2]))
        else:
            a_has = any(p[0] is not None for sl in tprops_a[i1:i2] for p in sl)
            b_has = any(p[0] is not None for sl in tprops_b[j1:j2] for p in sl)
            choose_b = (b_has and not a_has) or (b_has and a_has and sb["shd_chars"] > sa["shd_chars"])
            if choose_b:
                for k in range(j1, j2):
                    out_t.append(toks_b[k])
                    out_p.extend(tprops_b[k])
            else:
                for k in range(i1, i2):
                    out_t.append(toks_a[k])
                    out_p.extend(tprops_a[k])
    return "".join(out_t), out_p


def _merge_col2_shaded_sentences(
    earlier_tc1: ET.Element,
    later_tc1: ET.Element,
    comments_root: ET.Element = None,
    id_allocator: MarkupIdAllocator = None,
) -> bool:
    txt_a, props_a = _extract_cell_char_visual_props(earlier_tc1)
    txt_b, props_b = _extract_cell_char_visual_props(later_tc1)
    later_anchors = extract_cell_comment_anchors(later_tc1, comments_root)

    if not txt_b.strip():
        if later_anchors and comments_root is not None and id_allocator is not None:
            reanchor_comments_to_cell(earlier_tc1, later_anchors, comments_root, id_allocator)
            return True
        return False

    units_a = _split_col2_into_sentence_units(txt_a, props_a)
    units_b = _split_col2_into_sentence_units(txt_b, props_b)
    m_len, n_len = len(units_a), len(units_b)

    # Monotonic DP alignment to anchor matching sentences (>= 0.68 word similarity)
    dp = [[0.0] * (n_len + 1) for _ in range(m_len + 1)]
    sim_mat = [[_sent_similarity(units_a[i], units_b[j]) for j in range(n_len)] for i in range(m_len)]
    for i in range(m_len - 1, -1, -1):
        for j in range(n_len - 1, -1, -1):
            best = max(dp[i + 1][j], dp[i][j + 1])
            sc = sim_mat[i][j]
            if sc >= 0.68:
                cand = sc + dp[i + 1][j + 1]
                if cand >= best:
                    best = cand
            dp[i][j] = best

    anchors_ij = []
    ci, cj = 0, 0
    while ci < m_len and cj < n_len:
        sc = sim_mat[ci][cj]
        if sc >= 0.68 and abs(dp[ci][cj] - (sc + dp[ci + 1][cj + 1])) < 1e-9:
            anchors_ij.append((ci, cj))
            ci += 1
            cj += 1
        elif dp[ci + 1][cj] >= dp[ci][cj + 1]:
            ci += 1
        else:
            cj += 1

    out_chunks = []  # list of (text_str, props_list)

    def _append_chunk(c_txt: str, c_props: list):
        if not c_txt:
            return
        if c_txt == "\n":
            out_chunks.append((c_txt, c_props))
            return
        at_para_start = (not out_chunks) or (out_chunks[-1][0] == "\n")
        if at_para_start:
            # Strip leading horizontal spaces at start of paragraph
            k = 0
            while k < len(c_txt) and c_txt[k] == " ":
                k += 1
            c_txt = c_txt[k:]
            c_props = c_props[k:]
        else:
            prev_t = out_chunks[-1][0]
            if prev_t and not prev_t[-1].isspace() and not c_txt[0].isspace():
                c_txt = " " + c_txt
                c_props = [(None, None, None)] + list(c_props)
        if c_txt:
            out_chunks.append((c_txt, list(c_props)))

    def _emit_gap(a_slice: list, b_slice: list):
        if not a_slice and not b_slice:
            return
        a_shd = any(u["has_shd"] for u in a_slice)
        b_shd = any(u["has_shd"] for u in b_slice)
        if b_shd and not a_shd:
            chosen = b_slice
        elif a_shd and not b_shd:
            chosen = a_slice
        elif a_shd and b_shd:
            if len(a_slice) == len(b_slice):
                for ua, ub in zip(a_slice, b_slice):
                    mt, mp = _merge_two_shaded_sentence_units(ua, ub)
                    _append_chunk(mt, mp)
                return
            a_sc = sum(u["shd_chars"] for u in a_slice)
            b_sc = sum(u["shd_chars"] for u in b_slice)
            chosen = b_slice if b_sc > a_sc else a_slice
        else:
            chosen = a_slice
        for u in chosen:
            _append_chunk(u["text"], u["props"])

    prev_i, prev_j = 0, 0
    for ai, bj in anchors_ij:
        _emit_gap(units_a[prev_i:ai], units_b[prev_j:bj])
        ua, ub = units_a[ai], units_b[bj]
        if ua["is_nl"]:
            _append_chunk("\n", [(None, None, None)])
        elif ub["has_shd"] and not ua["has_shd"]:
            _append_chunk(ub["text"], ub["props"])
        elif ua["has_shd"] and not ub["has_shd"]:
            _append_chunk(ua["text"], ua["props"])
        elif ua["has_shd"] and ub["has_shd"]:
            mt, mp = _merge_two_shaded_sentence_units(ua, ub)
            _append_chunk(mt, mp)
        else:
            _append_chunk(ua["text"], ua["props"])
        prev_i, prev_j = ai + 1, bj + 1

    _emit_gap(units_a[prev_i:], units_b[prev_j:])

    merged_txt = "".join(t for t, _ in out_chunks)
    merged_props = [p for _, plist in out_chunks for p in plist]

    if merged_txt == txt_a and merged_props == props_a:
        if later_anchors and comments_root is not None and id_allocator is not None:
            reanchor_comments_to_cell(earlier_tc1, later_anchors, comments_root, id_allocator)
            return True
        return False

    all_anchors = extract_cell_comment_anchors(earlier_tc1, comments_root) + later_anchors
    _rebuild_cell_with_visual_props(earlier_tc1, merged_txt, merged_props)
    if all_anchors and comments_root is not None and id_allocator is not None:
        reanchor_comments_to_cell(earlier_tc1, all_anchors, comments_root, id_allocator)
    return True


def _is_top_instruction_paragraph_text(txt: str) -> bool:
    cleaned = re.sub(r"\s+", " ", (txt or "").strip()).rstrip(".").strip().lower()
    return cleaned in _TOP_INSTRUCTION_NORMS


def merge_same_doc_duplicate_rows(
    root: ET.Element,
    log_fn=print,
    comments_root: ET.Element = None,
) -> int:
    """
    Scan all <w:tr>s across <w:tbl>s in `body`. When a later row has the same normalized Col 1
    Chinese full paragraph (`len(cn_norm) >= 20`, exact match or similarity >= 0.96 with length diff < 5%)
    as an earlier row:
      - Merges Col 1 visual shading/highlight/color runs onto the earlier row's tc[0];
      - Merges Col 2 shaded regions onto the earlier row's tc[1] sentence-by-sentence;
      - Deletes the duplicate later <w:tr>; if its <w:tbl> now has 0 rows, deletes that <w:tbl>
        plus its immediately preceding title <w:p>(s) and trailing blank <w:p>(s).
    Returns count of merged duplicate rows.
    """
    body = root.find(f"{W}body")
    if body is None:
        return 0

    id_allocator = MarkupIdAllocator(doc_root=root, comments_root=comments_root) if comments_root is not None else None
    seen_rows = []  # list of {"cn_norm": str, "tr": Element, "body_idx": int}
    emptied_tables = set()
    merged_count = 0

    for b_idx, ch in enumerate(list(body)):
        if _local_name(ch.tag) != "tbl":
            continue
        for tr in list(ch.findall(f"{W}tr")):
            tcs = tr.findall(f"{W}tc")
            if len(tcs) < 2:
                continue
            cn_raw = get_cell_plain_text(tcs[0])
            cn_norm = _WS_COLLAPSE_RE.sub("", cn_raw)
            if len(cn_norm) < 20:
                continue

            matched_prev = None
            for prev in seen_rows:
                p_norm = prev["cn_norm"]
                if p_norm == cn_norm:
                    matched_prev = prev
                    break
                max_l = max(len(p_norm), len(cn_norm))
                if abs(len(p_norm) - len(cn_norm)) / max_l < 0.05:
                    if difflib.SequenceMatcher(None, p_norm, cn_norm, autojunk=False).ratio() >= 0.96:
                        matched_prev = prev
                        break

            if matched_prev is None:
                seen_rows.append({"cn_norm": cn_norm, "tr": tr, "body_idx": b_idx})
                continue

            prev_tcs = matched_prev["tr"].findall(f"{W}tc")
            _merge_col1_visual_runs(prev_tcs[0], tcs[0], comments_root=comments_root, id_allocator=id_allocator)
            _merge_col2_shaded_sentences(prev_tcs[1], tcs[1], comments_root=comments_root, id_allocator=id_allocator)

            ch.remove(tr)
            merged_count += 1
            if log_fn:
                log_fn(f"🔀 已自动合并同稿重复段落底纹并移除重复行：{cn_norm[:36]}...")

        if len(ch.findall(f"{W}tr")) == 0:
            emptied_tables.add(ch)

    if emptied_tables:
        children = list(body)
        n = len(children)
        to_remove = set()

        def _is_blank_p(el: ET.Element) -> bool:
            return _local_name(el.tag) == "p" and "".join(t.text or "" for t in el.findall(f".//{W}t")).strip() == ""

        def _is_title_p(el: ET.Element) -> bool:
            if _local_name(el.tag) != "p":
                return False
            ptxt = "".join(t.text or "" for t in el.findall(f".//{W}t")).strip()
            return bool(ptxt) and not _is_top_instruction_paragraph_text(ptxt)

        for idx, ch in enumerate(children):
            if ch not in emptied_tables:
                continue
            to_remove.add(idx)
            b_i = idx - 1
            if b_i >= 1 and _is_blank_p(children[b_i]) and _is_title_p(children[b_i - 1]):
                to_remove.add(b_i)
                b_i -= 1
            while b_i >= 0 and _is_title_p(children[b_i]):
                to_remove.add(b_i)
                b_i -= 1
            f_i = idx + 1
            while f_i < n and _is_blank_p(children[f_i]):
                to_remove.add(f_i)
                f_i += 1

        for idx in sorted(to_remove, reverse=True):
            body.remove(children[idx])

        non_sect = [c for c in list(body) if _local_name(c.tag) != "sectPr"]
        if not non_sect:
            body.insert(0, ET.Element(f"{W}p"))

    return merged_count


def _compute_text_visual_weight(text: str, cjk_mult: float = 1.0) -> float:
    cleaned = re.sub(r"\s+", " ", (text or "").strip())
    if not cleaned:
        return 0.0
    cjk_cnt = len(_CJK_WEIGHT_RE.findall(cleaned))
    other_cnt = len(cleaned) - cjk_cnt
    return cjk_cnt * cjk_mult + other_cnt * 1.0


def _project_bounded_pct_triple(weights: tuple, bounds: tuple, target_sum: int = 5000) -> tuple:
    """Project 3 positive weights onto exact integer triple summing to `target_sum` within `bounds`."""
    w_list = [max(1.0, float(w)) for w in weights]
    vals = [0.0, 0.0, 0.0]
    fixed = [False, False, False]

    for _ in range(5):
        rem_target = target_sum - sum(vals[i] for i in range(3) if fixed[i])
        free_idxs = [i for i in range(3) if not fixed[i]]
        if not free_idxs:
            break
        free_w = sum(w_list[i] for i in free_idxs)
        clamped_any = False
        for i in free_idxs:
            raw_v = rem_target * (w_list[i] / free_w) if free_w > 0 else (rem_target / len(free_idxs))
            lo, hi = bounds[i]
            if raw_v < lo:
                vals[i] = float(lo)
                fixed[i] = True
                clamped_any = True
            elif raw_v > hi:
                vals[i] = float(hi)
                fixed[i] = True
                clamped_any = True
            else:
                vals[i] = raw_v
        if not clamped_any:
            break

    ints = [min(bounds[i][1], max(bounds[i][0], round(vals[i]))) for i in range(3)]
    diff = target_sum - sum(ints)
    while diff != 0:
        step = 1 if diff > 0 else -1
        moved = False
        # Prefer adjusting Col 3, then Col 2, then Col 1 within bounds
        for idx in (2, 1, 0):
            cand = ints[idx] + step
            if bounds[idx][0] <= cand <= bounds[idx][1]:
                ints[idx] = cand
                diff -= step
                moved = True
                if diff == 0:
                    break
        if not moved:
            break
    return (ints[0], ints[1], ints[2])


def balance_trilingual_table_widths(root: ET.Element) -> int:
    """
    Adaptive 3-column table width balancer (`三语拉齐`):
      - Strips stray gridSpan / gridBefore / wBefore on 3-cell rows;
      - Computes visual text weight of Col 1 (CJK × 1.8), Col 2, and Col 3 per table;
      - Sets <w:tblW w:w="5000" w:type="pct"/>, <w:tblGrid> summing to printable width (~14560 dxa),
        and <w:tcW w:type="pct"> on all 3 cells (Col 1 [950, 1250], Col 2 [1700, 2050], Col 3 [1800, 2150], sum=5000).
    Returns count of balanced tables.
    """
    _ensure_landscape_a4_on_root(root)

    printable_dxa = 14560
    sect = root.find(f".//{W}sectPr")
    if sect is not None:
        pgSz = sect.find(f"{W}pgSz")
        pgMar = sect.find(f"{W}pgMar")
        try:
            pw = int(pgSz.attrib.get(f"{W}w", "16838")) if pgSz is not None else 16838
            ph = int(pgSz.attrib.get(f"{W}h", "11906")) if pgSz is not None else 11906
            eff_w = max(pw, ph)
            ml = int(pgMar.attrib.get(f"{W}left", "1134")) if pgMar is not None else 1134
            mr = int(pgMar.attrib.get(f"{W}right", "1134")) if pgMar is not None else 1134
            cand_dxa = eff_w - ml - mr
            if 9000 <= cand_dxa <= 19000:
                printable_dxa = cand_dxa
        except ValueError:
            printable_dxa = 14560

    bounds = ((950, 1250), (1700, 2050), (1800, 2150))
    balanced_count = 0

    for tbl in root.findall(f".//{W}tbl"):
        trs = tbl.findall(f"{W}tr")
        rows_3 = [tr for tr in trs if len(tr.findall(f"{W}tc")) == 3]
        if not rows_3:
            continue

        w0, w1, w2 = 0.0, 0.0, 0.0
        for tr in rows_3:
            _strip_stray_grid_spans_on_row(tr)
            tcs = tr.findall(f"{W}tc")
            c3_txt = get_cell_plain_text(tcs[2])
            if c3_txt.strip() == SKIPPED_ROW_MARKER:
                c3_txt = ""
            w0 += _compute_text_visual_weight(get_cell_plain_text(tcs[0]), cjk_mult=1.8)
            w1 += _compute_text_visual_weight(get_cell_plain_text(tcs[1]), cjk_mult=1.0)
            w2 += _compute_text_visual_weight(c3_txt, cjk_mult=1.4)

        if w1 <= 0 and w2 <= 0:
            pcts = (1050, 1950, 2000)
        elif w2 <= 0:
            pcts = (1050, 1950, 2000)
        else:
            pcts = _project_bounded_pct_triple((w0, w1, w2), bounds, target_sum=5000)

        g0 = round(printable_dxa * pcts[0] / 5000)
        g1 = round(printable_dxa * pcts[1] / 5000)
        g2 = printable_dxa - g0 - g1

        tblPr = tbl.find(f"{W}tblPr")
        if tblPr is None:
            tblPr = ET.Element(f"{W}tblPr")
            tbl.insert(0, tblPr)
        insert_tblpr_child_ordered(tblPr, ET.Element(f"{W}tblW", {f"{W}w": "5000", f"{W}type": "pct"}))

        old_grid = tbl.find(f"{W}tblGrid")
        new_grid = ET.Element(f"{W}tblGrid")
        for gw in (g0, g1, g2):
            ET.SubElement(new_grid, f"{W}gridCol", {f"{W}w": str(gw)})
        if old_grid is not None:
            idx_g = list(tbl).index(old_grid)
            tbl.remove(old_grid)
            tbl.insert(idx_g, new_grid)
        else:
            idx_pr = list(tbl).index(tblPr)
            tbl.insert(idx_pr + 1, new_grid)

        for tr in rows_3:
            tcs = tr.findall(f"{W}tc")
            for col_i in range(3):
                tcPr = tcs[col_i].find(f"{W}tcPr")
                if tcPr is None:
                    tcPr = ET.Element(f"{W}tcPr")
                    tcs[col_i].insert(0, tcPr)
                insert_tcpr_child_ordered(
                    tcPr,
                    ET.Element(f"{W}tcW", {f"{W}w": str(pcts[col_i]), f"{W}type": "pct"}),
                )

        balanced_count += 1

    return balanced_count


def format_up_rv_count_tag(up_words: int, rv_words: int) -> str:
    """
    Format UP/RV English word count tag omitting zero counts:
      - up > 0 and rv > 0 -> "(UP {up} words, RV {rv} words)"
      - up > 0 and rv == 0 -> "(UP {up} words)"
      - up == 0 and rv > 0 -> "(RV {rv} words)"
      - up == 0 and rv == 0 -> ""
    """
    up = max(0, int(up_words or 0))
    rv = max(0, int(rv_words or 0))
    if up > 0 and rv > 0:
        return f"(UP {up} words, RV {rv} words)"
    if up > 0:
        return f"(UP {up} words)"
    if rv > 0:
        return f"(RV {rv} words)"
    return ""


def _cell_is_dominantly_bold(tc: ET.Element) -> bool:
    bold_chars = 0
    total_chars = 0
    for p in tc.findall(f"{W}p"):
        p_b_el = p.find(f"./{W}pPr/{W}rPr/{W}b")
        p_bold = p_b_el is not None and p_b_el.attrib.get(f"{W}val", "1").strip().lower() not in ("0", "false", "off")
        for r in p.findall(f".//{W}r"):
            if _local_name(r.tag) != "r":
                continue
            rtxt = "".join(t.text or "" for t in r.findall(f"{W}t")).strip()
            if not rtxt:
                continue
            total_chars += len(rtxt)
            b_el = r.find(f"./{W}rPr/{W}b")
            if b_el is not None:
                is_b = b_el.attrib.get(f"{W}val", "1").strip().lower() not in ("0", "false", "off")
            else:
                is_b = p_bold
            if is_b:
                bold_chars += len(rtxt)
    return total_chars > 0 and (bold_chars / total_chars) >= 0.5


def extract_full_article_en_title(root: ET.Element) -> str:
    """
    Inspect Row 0 of the first <w:tbl>: if Row 0 has >= 2 cells and looks like a chapter title row
    (bold or <= 180 chars with Row 1 blank or body text below), return clean `get_cell_plain_text(tcs[1]).strip()`.
    """
    tbl = root.find(f".//{W}tbl")
    if tbl is None:
        return ""
    trs = tbl.findall(f"{W}tr")
    if not trs:
        return ""
    tcs0 = trs[0].findall(f"{W}tc")
    if len(tcs0) < 2:
        return ""

    en_title = re.sub(r"\s+", " ", get_cell_plain_text(tcs0[1])).strip()
    if not en_title:
        return ""

    is_bold = _cell_is_dominantly_bold(tcs0[1]) or _cell_is_dominantly_bold(tcs0[0])
    is_centered = any(
        (jc.attrib.get(f"{W}val") or "").strip().lower() == "center"
        for jc in tcs0[1].findall(f".//{W}pPr/{W}jc")
    )

    if (is_bold or is_centered) and len(en_title) <= 240:
        return en_title

    if len(en_title) <= 180 and len(trs) >= 2:
        tcs1 = trs[1].findall(f"{W}tc")
        row1_en = get_cell_plain_text(tcs1[1]).strip() if len(tcs1) >= 2 else ""
        row1_cn = get_cell_plain_text(tcs1[0]).strip() if len(tcs1) >= 1 else ""
        # Row 1 is blank separator row or longer body paragraph
        if (not row1_en and not row1_cn) or len(row1_en) > len(en_title):
            return en_title

    return ""


def _build_top_instruction_paragraph(text: str, highlight_color: str) -> ET.Element:
    p_el = ET.Element(f"{W}p")
    pPr = ET.SubElement(p_el, f"{W}pPr")
    insert_ppr_child_ordered(pPr, ET.Element(f"{W}spacing", {
        f"{W}before": "0",
        f"{W}after": "0",
        f"{W}line": "400",
        f"{W}lineRule": "exact",
    }))
    insert_ppr_child_ordered(pPr, ET.Element(f"{W}ind", {
        f"{W}firstLine": "0",
        f"{W}firstLineChars": "0",
    }))
    insert_ppr_child_ordered(pPr, ET.Element(f"{W}jc", {f"{W}val": "left"}))
    p_rPr = ET.Element(f"{W}rPr")
    insert_rpr_child_ordered(p_rPr, ET.Element(f"{W}rFonts", {
        f"{W}ascii": "Georgia",
        f"{W}eastAsia": "宋体",
        f"{W}hAnsi": "Georgia",
        f"{W}cs": "Georgia",
    }))
    insert_rpr_child_ordered(p_rPr, ET.Element(f"{W}sz", {f"{W}val": "28"}))
    insert_rpr_child_ordered(p_rPr, ET.Element(f"{W}szCs", {f"{W}val": "28"}))
    insert_ppr_child_ordered(pPr, p_rPr)

    r_el = ET.SubElement(p_el, f"{W}r")
    rPr = ET.SubElement(r_el, f"{W}rPr")
    insert_rpr_child_ordered(rPr, ET.Element(f"{W}rFonts", {
        f"{W}ascii": "Georgia",
        f"{W}eastAsia": "宋体",
        f"{W}hAnsi": "Georgia",
        f"{W}cs": "Georgia",
    }))
    insert_rpr_child_ordered(rPr, ET.Element(f"{W}sz", {f"{W}val": "28"}))
    insert_rpr_child_ordered(rPr, ET.Element(f"{W}szCs", {f"{W}val": "28"}))
    insert_rpr_child_ordered(rPr, ET.Element(f"{W}highlight", {f"{W}val": highlight_color}))
    t_el = ET.SubElement(r_el, f"{W}t")
    t_el.text = text
    return p_el


def update_top_highlight_instruction_lines(root: ET.Element, has_up: bool, has_rv: bool) -> None:
    """
    Synchronize leading instruction lines at the very top of `body` (index 0):
      - Removes any existing leading <w:p> matching "blue highlighted parts need to be reviewed"
        or "yellow highlighted parts need to be updated" (case-insensitive, trailing '.' ignored).
      - Inserts at index 0:
          * If `has_rv`: Georgia sz="28" szCs="28" highlight="cyan", "Blue highlighted parts need to be reviewed."
          * If `has_up`: Georgia sz="28" szCs="28" highlight="yellow", "Yellow highlighted parts need to be updated."
          * If neither: inserts nothing!
    """
    body = root.find(f"{W}body")
    if body is None:
        return

    # Remove any existing top instruction lines appearing before the first table
    for ch in list(body):
        loc = _local_name(ch.tag)
        if loc in ("tbl", "sectPr"):
            break
        if loc == "p":
            ptxt = "".join(t.text or "" for t in ch.findall(f".//{W}t")).strip()
            if _is_top_instruction_paragraph_text(ptxt):
                body.remove(ch)

    to_insert = []
    if has_rv:
        to_insert.append(_build_top_instruction_paragraph(
            "Blue highlighted parts need to be reviewed.",
            "cyan",
        ))
    if has_up:
        to_insert.append(_build_top_instruction_paragraph(
            "Yellow highlighted parts need to be updated.",
            "yellow",
        ))

    for idx, p_el in enumerate(to_insert):
        body.insert(idx, p_el)

    non_sect = [c for c in list(body) if _local_name(c.tag) != "sectPr"]
    if not non_sect:
        body.insert(0, ET.Element(f"{W}p"))

