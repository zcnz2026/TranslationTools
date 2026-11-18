import json
import os
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional


AUTO_COMBO_LABEL = "AUTO - 自动识别（按总库目录/文件名）"

DEFAULT_TITLE_PT = 14.0
DEFAULT_BODY_PT = 12.0
DEFAULT_CITATION_PT = 10.5

# Common alternate language codes used in filenames / legacy configs -> canonical DEFAULT_LANGUAGES code
_CODE_ALIASES: Dict[str, str] = {
    "PO": "PL",
    "UA": "UK",
    "EL": "GR",
    "FIL": "TL",
}

DEFAULT_LANGUAGES: List[Dict] = [
    {"code": "FR", "name": "法语", "keywords": ["法语", "中英法"], "font": "Georgia"},
    {"code": "PT", "name": "葡萄牙语", "keywords": ["葡萄牙语", "葡语", "中英葡"], "font": "Georgia"},
    {"code": "ES", "name": "西班牙语", "keywords": ["西班牙语", "西语", "中英西"], "font": "Georgia"},
    {"code": "DE", "name": "德语", "keywords": ["德语", "中英德"], "font": "Times New Roman"},
    {"code": "IT", "name": "意大利语", "keywords": ["意大利语", "意语", "意大利", "中英意"], "font": "Georgia"},
    {"code": "GR", "name": "希腊语", "keywords": ["希腊语", "腊语", "希腊", "中英腊"], "font": "Georgia"},
    {"code": "RU", "name": "俄语", "keywords": ["俄语", "中英俄"], "font": "Times New Roman"},
    {"code": "UK", "name": "乌克兰语", "keywords": ["乌克兰语", "乌语", "乌克兰", "中英乌"], "font": "Arial"},
    {"code": "KO", "name": "韩语", "keywords": ["韩语", "朝鲜语", "中英韩"], "font": "Malgun Gothic"},
    {"code": "JA", "name": "日语", "keywords": ["日语", "日本语", "中英日"], "font": ""},
    {
        "code": "TH",
        "name": "泰语",
        "keywords": ["泰语", "中英泰"],
        "font": "Angsana New",
        "title_pt": 18,
        "body_pt": 16,
        "citation_pt": 14,
    },
    {"code": "VI", "name": "越南语", "keywords": ["越南语", "越语", "越南", "中英越"], "font": "Times New Roman"},
    {"code": "ID", "name": "印尼语", "keywords": ["印尼语", "印度尼西亚语", "印尼", "中英印尼"], "font": "Times New Roman"},
    {"code": "MS", "name": "马来语", "keywords": ["马来语", "中英马来", "中英马"], "font": ""},
    {"code": "TL", "name": "塔加洛语", "keywords": ["塔加洛语", "他加禄语", "菲律宾语", "菲语", "中英塔", "中英菲"], "font": "Georgia"},
    {"code": "MY", "name": "缅甸语", "keywords": ["缅甸语", "缅语", "缅甸", "中英缅"], "font": "Myanmar Text"},
    {"code": "KM", "name": "高棉语", "keywords": ["高棉语", "柬埔寨语", "柬语", "中英柬", "中英高棉"], "font": "Leelawadee UI"},
    {"code": "LO", "name": "老挝语", "keywords": ["老挝语", "老语", "老挝", "中英老"], "font": "DokChampa"},
    {"code": "AR", "name": "阿拉伯语", "keywords": ["阿拉伯语", "阿语", "中英阿"], "font": "Simplified Arabic"},
    {"code": "TR", "name": "土耳其语", "keywords": ["土耳其语", "土语", "中英土"], "font": ""},
    {"code": "FA", "name": "波斯语", "keywords": ["波斯语", "中英波斯"], "font": ""},
    {"code": "HE", "name": "希伯来语", "keywords": ["希伯来语", "希伯来", "中英希伯来"], "font": "FrankRuehl"},
    {"code": "HI", "name": "印地语", "keywords": ["印地语", "中英印地"], "font": "Mangal"},
    {"code": "NE", "name": "尼泊尔语", "keywords": ["尼泊尔语", "尼语", "中英尼"], "font": "Mangal"},
    {"code": "UR", "name": "乌尔都语", "keywords": ["乌尔都语", "中英乌尔都"], "font": ""},
    {"code": "BN", "name": "孟加拉语", "keywords": ["孟加拉语", "中英孟"], "font": ""},
    {"code": "TA", "name": "泰米尔语", "keywords": ["泰米尔语", "中英泰米尔"], "font": ""},
    {"code": "PA", "name": "旁遮普语", "keywords": ["旁遮普语", "中英旁遮普"], "font": ""},
    {"code": "TE", "name": "泰卢固语", "keywords": ["泰卢固语", "中英泰卢固"], "font": ""},
    {"code": "MR", "name": "马拉地语", "keywords": ["马拉地语", "中英马拉地"], "font": ""},
    {"code": "PL", "name": "波兰语", "keywords": ["波兰语", "波语", "中英波", "PO"], "font": "Georgia"},
    {"code": "NL", "name": "荷兰语", "keywords": ["荷兰语", "荷语", "中英荷"], "font": "Georgia"},
    {"code": "RO", "name": "罗马尼亚语", "keywords": ["罗马尼亚语", "罗语", "罗马尼亚", "中英罗"], "font": "Georgia"},
    {"code": "CS", "name": "捷克语", "keywords": ["捷克语", "中英捷"], "font": ""},
    {"code": "SK", "name": "斯洛伐克语", "keywords": ["斯洛伐克语", "斯洛伐克", "中英斯洛伐克"], "font": "Times New Roman"},
    {"code": "HU", "name": "匈牙利语", "keywords": ["匈牙利语", "马扎尔语", "匈语", "中英匈", "中英马扎尔"], "font": "Times New Roman"},
    {"code": "BG", "name": "保加利亚语", "keywords": ["保加利亚语", "保语", "中英保"], "font": "Georgia"},
    {"code": "SR", "name": "塞尔维亚语", "keywords": ["塞尔维亚语", "中英塞"], "font": "Georgia"},
    {"code": "HR", "name": "克罗地亚语", "keywords": ["克罗地亚语", "中英克"], "font": "Georgia"},
    {"code": "SL", "name": "斯洛文尼亚语", "keywords": ["斯洛文尼亚语", "斯洛文尼亚", "中英斯洛文尼亚"], "font": ""},
    {"code": "SV", "name": "瑞典语", "keywords": ["瑞典语", "中英瑞"], "font": ""},
    {"code": "MN", "name": "蒙古语", "keywords": ["蒙古语", "蒙语", "中英蒙"], "font": "Times New Roman"},
    {"code": "HMN", "name": "苗语", "keywords": ["苗语", "中英苗"], "font": "Georgia"},
    {"code": "SW", "name": "斯瓦希里语", "keywords": ["斯瓦希里语", "中英斯"], "font": "Georgia"},
    {"code": "AM", "name": "阿姆哈拉语", "keywords": ["阿姆哈拉语", "中英阿姆"], "font": ""},
    {"code": "HY", "name": "亚美尼亚语", "keywords": ["亚美尼亚语", "亚美尼亚", "中英亚美尼亚"], "font": "Noto Sans Armenian"},
    {"code": "KA", "name": "格鲁吉亚语", "keywords": ["格鲁吉亚语", "格鲁吉亚", "中英格鲁吉亚"], "font": "Sylfaen"},
]

_DEFAULT_BY_CODE: Dict[str, Dict] = {item["code"]: item for item in DEFAULT_LANGUAGES}


def _lookup_default_entry(code: str) -> Optional[Dict]:
    c = (code or "").strip().upper()
    if c in _DEFAULT_BY_CODE:
        return _DEFAULT_BY_CODE[c]
    canon = _CODE_ALIASES.get(c)
    if canon and canon in _DEFAULT_BY_CODE:
        return _DEFAULT_BY_CODE[canon]
    return None


def _clone_default_languages() -> List[Dict]:
    out = []
    for item in DEFAULT_LANGUAGES:
        d = dict(item)
        d["keywords"] = list(item.get("keywords", []))
        out.append(d)
    return out


def get_app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_languages_file() -> Path:
    p = get_app_base_dir() / "languages.json"
    os.environ["SCATTERED_REV_LANGUAGES_JSON"] = str(p)
    return p


def load_languages() -> List[Dict]:
    p = get_languages_file()
    if not p.exists():
        defaults = _clone_default_languages()
        save_languages(defaults)
        return defaults

    try:
        with open(p, "r", encoding="utf-8") as f:
            data = json.load(f)
        raw_list = data.get("languages", data) if isinstance(data, dict) else data
        had_any_font_field = any(
            isinstance(it, dict) and "font" in it for it in (raw_list if isinstance(raw_list, list) else [])
        )
        cleaned: List[Dict] = []
        for item in raw_list:
            if not isinstance(item, dict):
                continue
            code = str(item.get("code", "")).strip().upper()
            name = str(item.get("name", "")).strip()
            kws = [str(k).strip() for k in (item.get("keywords") or []) if str(k).strip()]
            if not code:
                continue
            def_entry = _lookup_default_entry(code)
            if not name:
                name = def_entry["name"] if def_entry else code
            if name not in kws:
                kws.insert(0, name)
            # Backfill newly added standard keywords (e.g. '马扎尔语', '菲语', 'PO') without losing user keywords
            if def_entry:
                for d_kw in def_entry.get("keywords", []):
                    if d_kw not in kws:
                        kws.append(d_kw)

            # Preserve user-configured font if explicitly present in JSON; otherwise backfill from DEFAULT_LANGUAGES
            if "font" in item and item["font"] is not None:
                font_val = str(item.get("font", "")).strip()
            else:
                font_val = str(def_entry.get("font", "")).strip() if def_entry else ""

            entry_out: Dict = {
                "code": code,
                "name": name,
                "keywords": kws,
                "font": font_val,
            }

            if item.get("tri_prefix"):
                entry_out["tri_prefix"] = str(item["tri_prefix"]).strip()

            if "favorite" in item and item["favorite"] is not None:
                fav_raw = item["favorite"]
                if isinstance(fav_raw, str):
                    entry_out["favorite"] = fav_raw.strip().lower() in ("1", "true", "yes", "on", "★")
                else:
                    entry_out["favorite"] = bool(fav_raw)

            for pt_key in ("title_pt", "body_pt", "citation_pt"):
                if pt_key in item and item[pt_key] is not None:
                    try:
                        val_num = float(item[pt_key])
                        entry_out[pt_key] = int(val_num) if val_num.is_integer() else val_num
                    except (TypeError, ValueError):
                        pass
                elif def_entry and pt_key in def_entry:
                    entry_out[pt_key] = def_entry[pt_key]

            cleaned.append(entry_out)
        if cleaned:
            # Seamless upgrade from pre-v1.7.0 languages.json: append any newly added standard languages
            if not had_any_font_field:
                seen_codes = {e["code"] for e in cleaned}
                for d in DEFAULT_LANGUAGES:
                    if d["code"] not in seen_codes:
                        item_copy = dict(d)
                        item_copy["keywords"] = list(d.get("keywords", []))
                        cleaned.append(item_copy)
            return cleaned
    except Exception:
        pass
    return _clone_default_languages()


def save_languages(entries: List[Dict]) -> None:
    p = get_languages_file()
    serialized: List[Dict] = []
    for item in entries:
        if not isinstance(item, dict):
            continue
        d = dict(item)
        if "favorite" in d and d["favorite"] is not None:
            fav_raw = d["favorite"]
            if isinstance(fav_raw, str):
                d["favorite"] = fav_raw.strip().lower() in ("1", "true", "yes", "on", "★")
            else:
                d["favorite"] = bool(fav_raw)
        serialized.append(d)
    payload = {
        "_comment": "可在此自由新增或修改语种配置：code 为输出定稿文件名前缀，name 为显示名称，keywords 为从总库目录/文件名自动识别该语种的关键词，font 为第 3 列标准外文字体（留空则保留原文档字体），favorite 为可选常用语种置顶布尔标记。",
        "languages": serialized,
    }
    with open(p, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def get_lang_typography(code: str, entries: Optional[List[Dict]] = None) -> Dict:
    """Return standard Col 3 font & OOXML half-point sizes (`w:sz`) for `code`:
    `{"font": str, "title_sz": int, "body_sz": int, "citation_sz": int}`.
    Defaults: 28 / 24 / 21 (14pt / 12pt / 10.5pt); Thai (`TH`): 36 / 32 / 28 (18pt / 16pt / 14pt).
    """
    c = (code or "").strip().upper()
    active_entries = entries if entries is not None else load_languages()

    matched: Optional[Dict] = None
    if c and c != "AUTO":
        for e in active_entries:
            if str(e.get("code", "")).strip().upper() == c:
                matched = e
                break
        if matched is None:
            canon = _CODE_ALIASES.get(c)
            if canon:
                for e in active_entries:
                    if str(e.get("code", "")).strip().upper() == canon:
                        matched = e
                        break
        if matched is None:
            for e in active_entries:
                kws_upper = [str(k).strip().upper() for k in e.get("keywords", [])]
                if c in kws_upper:
                    matched = e
                    break

    def_entry = _lookup_default_entry(c) if c and c != "AUTO" else None

    font_val = ""
    if matched is not None:
        if "font" in matched and matched["font"] is not None:
            font_val = str(matched.get("font", "")).strip()
        elif def_entry is not None:
            font_val = str(def_entry.get("font", "")).strip()
    elif def_entry is not None:
        font_val = str(def_entry.get("font", "")).strip()

    def _resolve_sz(pt_key: str, default_pt: float) -> int:
        raw_pt = None
        if matched is not None and matched.get(pt_key) is not None:
            raw_pt = matched.get(pt_key)
        elif def_entry is not None and def_entry.get(pt_key) is not None:
            raw_pt = def_entry.get(pt_key)
        if raw_pt is not None:
            try:
                return int(round(float(raw_pt) * 2))
            except (TypeError, ValueError):
                pass
        return int(round(default_pt * 2))

    return {
        "font": font_val,
        "title_sz": _resolve_sz("title_pt", DEFAULT_TITLE_PT),
        "body_sz": _resolve_sz("body_pt", DEFAULT_BODY_PT),
        "citation_sz": _resolve_sz("citation_pt", DEFAULT_CITATION_PT),
    }


def build_keywords_list(name: str, raw_keywords: str) -> List[str]:
    parts = [x.strip() for x in re.split(r"[,，、;；]+", raw_keywords or "") if x.strip()]
    clean_name = name.strip()
    if clean_name and clean_name not in parts:
        parts.insert(0, clean_name)
    # If user left keywords blank and name ends with '语' (e.g. '塞尔维亚语'), also add '中英X' helper if 2-char like '泰语'
    if not raw_keywords.strip() and clean_name.endswith("语") and len(clean_name) >= 2:
        short = clean_name[:-1]
        cand = f"中英{short}" if len(short) <= 2 else f"中英{short[0]}"
        if cand not in parts:
            parts.append(cand)
    return parts


_COMBO_MARKER_PREFIX_RE = re.compile(r"^(?:[★☆⏱⭐📌•·\ufe0f\u200b]+\s*)+")


def format_combo_label(entry: Dict, prefix: str = "") -> str:
    p = (prefix or "").strip()
    if p:
        return f"{p} {entry['code']} - {entry['name']}"
    return f"{entry['code']} - {entry['name']}"


def get_combo_values(entries: List[Dict]) -> List[str]:
    return [AUTO_COMBO_LABEL] + [format_combo_label(e) for e in entries]


def get_prioritized_combo_values(
    entries: List[Dict],
    recent_codes: Optional[List[str]] = None,
    favorite_codes: Optional[List[str]] = None,
) -> List[str]:
    """
    Build language combobox values with pinned groups:
      0. `AUTO_COMBO_LABEL`
      1. Favorite languages (`entry.get("favorite")` is True or code in `favorite_codes`): `"★ CODE - Name"`
      2. Recently used languages (`recent_codes` in MRU order, excluding favorites): `"⏱ CODE - Name"`
      3. Remaining languages in registry order: `"CODE - Name"`
    """
    out: List[str] = [AUTO_COMBO_LABEL]
    if not entries:
        return out

    by_code: Dict[str, Dict] = {}
    ordered_entries: List[Dict] = []
    for e in entries:
        if not isinstance(e, dict):
            continue
        c = str(e.get("code", "")).strip().upper()
        if not c or c in by_code:
            continue
        by_code[c] = e
        ordered_entries.append(e)

    fav_set = set()
    for raw_fc in (favorite_codes or []):
        fc = _COMBO_MARKER_PREFIX_RE.sub("", str(raw_fc or "").strip())
        if " - " in fc:
            fc = fc.split(" - ", 1)[0].strip()
        fc_up = re.sub(r"[^A-Za-z0-9_-]", "", fc).upper()
        if not fc_up or fc_up in ("AUTO", "XX"):
            continue
        fav_set.add(fc_up)
        if fc_up not in by_code and fc_up in _CODE_ALIASES:
            fav_set.add(_CODE_ALIASES[fc_up])

    seen_codes = set()

    # 1. Favorites
    for e in ordered_entries:
        c = str(e["code"]).strip().upper()
        if bool(e.get("favorite")) or c in fav_set:
            out.append(format_combo_label(e, "★ "))
            seen_codes.add(c)

    # 2. Recently used (MRU order, excluding already-listed favorites)
    for raw_rc in (recent_codes or []):
        rc = _COMBO_MARKER_PREFIX_RE.sub("", str(raw_rc or "").strip())
        if " - " in rc:
            rc = rc.split(" - ", 1)[0].strip()
        rc_up = re.sub(r"[^A-Za-z0-9_-]", "", rc).upper()
        if not rc_up or rc_up in ("AUTO", "XX"):
            continue
        target_code = rc_up
        if target_code not in by_code and target_code in _CODE_ALIASES and _CODE_ALIASES[target_code] in by_code:
            target_code = _CODE_ALIASES[target_code]
        if target_code in by_code and target_code not in seen_codes:
            out.append(format_combo_label(by_code[target_code], "⏱ "))
            seen_codes.add(target_code)

    # 3. Remaining languages
    for e in ordered_entries:
        c = str(e["code"]).strip().upper()
        if c not in seen_codes:
            out.append(format_combo_label(e))
            seen_codes.add(c)

    return out


def resolve_combo_to_code(text: str, entries: List[Dict]) -> str:
    s = _COMBO_MARKER_PREFIX_RE.sub("", (text or "").strip()).strip()
    if not s or s.upper() == "AUTO" or s.upper().startswith("AUTO ") or s.upper().startswith("AUTO-"):
        return "AUTO"
    if " - " in s:
        lhs = _COMBO_MARKER_PREFIX_RE.sub("", s.split(" - ", 1)[0]).strip()
        lhs_clean = re.sub(r"[^A-Za-z0-9_-]", "", lhs).upper()
        return lhs_clean or "AUTO"
    s_up = s.upper()
    for e in entries:
        if (
            s_up == e["code"]
            or s == e["name"]
            or s in e.get("keywords", [])
            or s_up in [str(k).upper() for k in e.get("keywords", [])]
        ):
            return e["code"]
    canon = _CODE_ALIASES.get(s_up)
    if canon:
        for e in entries:
            if e["code"] == canon:
                return canon
    return re.sub(r"[^A-Za-z0-9_-]", "", s).upper() or "AUTO"


def record_recent_language(
    code: str,
    recent_codes: Optional[List[str]] = None,
    max_items: int = 5,
) -> List[str]:
    """Return updated MRU language code list with `code` moved/inserted at index 0 (ignoring AUTO / XX / empty)."""
    limit = max(0, int(max_items))
    if limit == 0:
        return []

    def _norm_code(raw_val: str) -> str:
        val = _COMBO_MARKER_PREFIX_RE.sub("", str(raw_val or "").strip()).strip()
        if " - " in val:
            val = val.split(" - ", 1)[0].strip()
        return re.sub(r"[^A-Za-z0-9_-]", "", val).upper()

    cleaned_existing: List[str] = []
    for item in (recent_codes or []):
        c = _norm_code(item)
        if c and c not in ("AUTO", "XX") and c not in cleaned_existing:
            cleaned_existing.append(c)

    target = _norm_code(code)
    if not target or target in ("AUTO", "XX"):
        return cleaned_existing[:limit]

    updated = [target] + [c for c in cleaned_existing if c != target]
    return updated[:limit]


def find_lang_display_name(code: str, entries: List[Dict]) -> str:
    c = (code or "").strip().upper()
    for e in entries:
        if e["code"] == c:
            return f"{e['name']} ({c})"
    def_entry = _lookup_default_entry(c)
    if def_entry:
        return f"{def_entry['name']} ({c})"
    return c or "auto"


def _keyword_matches_hint(kw: str, hint_text: str) -> bool:
    if not kw:
        return False
    if kw.isascii():
        # ASCII keywords (e.g. 'PO') must match on non-letter boundaries so 'PORTUGUESE' / 'REPORT' never false-match
        pat = rf"(?<![A-Za-z]){re.escape(kw)}(?![A-Za-z])"
        return re.search(pat, hint_text) is not None
    return kw in hint_text


def detect_lang_code_from_hint(hint_text: str, entries: List[Dict]) -> str:
    """Detect language code from folder name, filename, or title text using configured keywords."""
    s = (hint_text or "").strip()
    if not s:
        return ""

    # Prefer longest non-ASCII keyword match first (so '中英波斯' beats '中英波' and '中英斯洛伐克' beats '中英斯'),
    # preserving entry order on ties.
    best_code = ""
    best_score: tuple = (-1, -1, 0)  # (is_non_ascii, kw_len, -entry_idx)
    for idx, e in enumerate(entries):
        for kw in e.get("keywords", []):
            kw_s = (kw or "").strip()
            if not kw_s:
                continue
            if _keyword_matches_hint(kw_s, s):
                cand_score = (0 if kw_s.isascii() else 1, len(kw_s), -idx)
                if cand_score > best_score:
                    best_score = cand_score
                    best_code = e["code"]
    if best_code and best_score[0] == 1:
        return best_code

    # Check explicit code prefixes/tokens like "ES_...", "[ES]", "71 HI_38_...", "PO_..."
    known_codes = {e["code"]: e["code"] for e in entries}
    for alias, canon in _CODE_ALIASES.items():
        if alias not in known_codes:
            if canon in known_codes:
                known_codes[alias] = canon
            elif canon in _DEFAULT_BY_CODE:
                known_codes[alias] = canon

    for m in re.finditer(r"(?:^|[\s\[\(\\/])([A-Za-z]{2,4})(?=[_\]\-])", s):
        cand_raw = m.group(1)
        cand = cand_raw.upper()
        # Require uppercase in filename when preceded by space (so '71 HI_38_' matches 'HI', not lowercase words)
        if m.start() > 0 and s[m.start()] == " " and not cand_raw.isupper():
            continue
        if cand in known_codes:
            return known_codes[cand]

    m_lead = re.match(r"^\[?([A-Za-z]{2,4})\]?[_\-\s]", s)
    if m_lead:
        cand = m_lead.group(1).upper()
        if cand in known_codes:
            return known_codes[cand]

    return best_code


def get_tri_prefix(code: str, entries: List[Dict]) -> str:
    """Return '中英X' trilingual prefix for a language code (e.g. GR -> '中英腊', MY -> '中英缅', ES -> '中英西')."""
    c = (code or "").strip().upper()
    if not c or c == "AUTO":
        return "中英外"
    search_list = list(entries) + [d for d in DEFAULT_LANGUAGES if all(x.get("code") != d["code"] for x in entries)]
    target_codes = [c]
    if c in _CODE_ALIASES:
        target_codes.append(_CODE_ALIASES[c])
    for tc in target_codes:
        for e in search_list:
            if e["code"] == tc:
                explicit = str(e.get("tri_prefix", "")).strip()
                if explicit:
                    return explicit
                tri_kws = [kw.strip() for kw in e.get("keywords", []) if kw.strip().startswith("中英")]
                if tri_kws:
                    tri_kws.sort(key=len)
                    return tri_kws[0]
                nm = e.get("name", "").strip()
                if nm.endswith("语") and len(nm) >= 2:
                    core = nm[:-1]
                    return f"中英{core}" if len(core) <= 2 else f"中英{core[0]}"
                return f"中英{c}"
    return f"中英{c}"
