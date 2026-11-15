import json
import os
import re
import sys
from pathlib import Path
from typing import Dict, List


AUTO_COMBO_LABEL = "AUTO - 自动识别（按总库目录/文件名）"

DEFAULT_LANGUAGES: List[Dict] = [
    {"code": "FR", "name": "法语", "keywords": ["法语", "中英法"]},
    {"code": "PT", "name": "葡萄牙语", "keywords": ["葡萄牙语", "葡语", "中英葡"]},
    {"code": "ES", "name": "西班牙语", "keywords": ["西班牙语", "西语", "中英西"]},
    {"code": "DE", "name": "德语", "keywords": ["德语", "中英德"]},
    {"code": "IT", "name": "意大利语", "keywords": ["意大利语", "意语", "中英意"]},
    {"code": "GR", "name": "希腊语", "keywords": ["希腊语", "腊语", "中英腊"]},
    {"code": "RU", "name": "俄语", "keywords": ["俄语", "中英俄"]},
    {"code": "UK", "name": "乌克兰语", "keywords": ["乌克兰语", "乌语", "中英乌"]},
    {"code": "KO", "name": "韩语", "keywords": ["韩语", "朝鲜语", "中英韩"]},
    {"code": "JA", "name": "日语", "keywords": ["日语", "日本语", "中英日"]},
    {"code": "TH", "name": "泰语", "keywords": ["泰语", "中英泰"]},
    {"code": "VI", "name": "越南语", "keywords": ["越南语", "越语", "中英越"]},
    {"code": "ID", "name": "印尼语", "keywords": ["印尼语", "印度尼西亚语", "中英印尼"]},
    {"code": "MS", "name": "马来语", "keywords": ["马来语", "中英马来", "中英马"]},
    {"code": "TL", "name": "塔加洛语", "keywords": ["塔加洛语", "他加禄语", "菲律宾语", "中英塔", "中英菲"]},
    {"code": "MY", "name": "缅甸语", "keywords": ["缅甸语", "缅语", "中英缅"]},
    {"code": "KM", "name": "高棉语", "keywords": ["高棉语", "柬埔寨语", "柬语", "中英柬", "中英高棉"]},
    {"code": "LO", "name": "老挝语", "keywords": ["老挝语", "老语", "中英老"]},
    {"code": "AR", "name": "阿拉伯语", "keywords": ["阿拉伯语", "阿语", "中英阿"]},
    {"code": "TR", "name": "土耳其语", "keywords": ["土耳其语", "土语", "中英土"]},
    {"code": "FA", "name": "波斯语", "keywords": ["波斯语", "中英波斯"]},
    {"code": "HE", "name": "希伯来语", "keywords": ["希伯来语", "中英希伯来"]},
    {"code": "HI", "name": "印地语", "keywords": ["印地语", "中英印地"]},
    {"code": "UR", "name": "乌尔都语", "keywords": ["乌尔都语", "中英乌尔都"]},
    {"code": "BN", "name": "孟加拉语", "keywords": ["孟加拉语", "中英孟"]},
    {"code": "PL", "name": "波兰语", "keywords": ["波兰语", "波语", "中英波"]},
    {"code": "NL", "name": "荷兰语", "keywords": ["荷兰语", "荷语", "中英荷"]},
    {"code": "RO", "name": "罗马尼亚语", "keywords": ["罗马尼亚语", "罗语", "中英罗"]},
    {"code": "CS", "name": "捷克语", "keywords": ["捷克语", "中英捷"]},
    {"code": "HU", "name": "匈牙利语", "keywords": ["匈牙利语", "匈语", "中英匈"]},
    {"code": "BG", "name": "保加利亚语", "keywords": ["保加利亚语", "保语", "中英保"]},
    {"code": "SV", "name": "瑞典语", "keywords": ["瑞典语", "中英瑞"]},
    {"code": "MN", "name": "蒙古语", "keywords": ["蒙古语", "蒙语", "中英蒙"]},
    {"code": "SW", "name": "斯瓦希里语", "keywords": ["斯瓦希里语", "中英斯"]},
    {"code": "AM", "name": "阿姆哈拉语", "keywords": ["阿姆哈拉语", "中英阿姆"]},
]


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
        save_languages(DEFAULT_LANGUAGES)
        return [dict(item) for item in DEFAULT_LANGUAGES]

    try:
        with open(p, "r", encoding="utf-8") as f:
            data = json.load(f)
        raw_list = data.get("languages", data) if isinstance(data, dict) else data
        cleaned = []
        for item in raw_list:
            code = str(item.get("code", "")).strip().upper()
            name = str(item.get("name", "")).strip()
            kws = [str(k).strip() for k in (item.get("keywords") or []) if str(k).strip()]
            if not code:
                continue
            if not name:
                name = code
            if name not in kws:
                kws.insert(0, name)
            cleaned.append({"code": code, "name": name, "keywords": kws})
        if cleaned:
            return cleaned
    except Exception:
        pass
    return [dict(item) for item in DEFAULT_LANGUAGES]


def save_languages(entries: List[Dict]) -> None:
    p = get_languages_file()
    payload = {
        "_comment": "可在此自由新增或修改语种配置：code 为输出定稿文件名前缀，name 为显示名称，keywords 为从总库目录/文件名自动识别该语种的关键词。",
        "languages": entries,
    }
    with open(p, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


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


def format_combo_label(entry: Dict) -> str:
    return f"{entry['code']} - {entry['name']}"


def get_combo_values(entries: List[Dict]) -> List[str]:
    return [AUTO_COMBO_LABEL] + [format_combo_label(e) for e in entries]


def resolve_combo_to_code(text: str, entries: List[Dict]) -> str:
    s = (text or "").strip()
    if not s or s.upper() == "AUTO" or s.startswith("AUTO "):
        return "AUTO"
    if " - " in s:
        return s.split(" - ", 1)[0].strip().upper()
    # Check if user typed a Chinese name or keyword
    for e in entries:
        if s.upper() == e["code"] or s == e["name"] or s in e.get("keywords", []):
            return e["code"]
    return re.sub(r"[^A-Za-z0-9_-]", "", s).upper() or "AUTO"


def find_lang_display_name(code: str, entries: List[Dict]) -> str:
    c = (code or "").strip().upper()
    for e in entries:
        if e["code"] == c:
            return f"{e['name']} ({c})"
    return c or "auto"


def detect_lang_code_from_hint(hint_text: str, entries: List[Dict]) -> str:
    """Detect language code from folder name, filename, or title text using configured keywords."""
    s = (hint_text or "").strip()
    if not s:
        return ""
    for e in entries:
        for kw in e.get("keywords", []):
            if kw and kw in s:
                return e["code"]
    # Also check leading code prefix like "ES_..." or "[ES]"
    m = re.match(r"^\[?([A-Za-z]{2,4})\]?[_\-\s]", s)
    if m:
        cand = m.group(1).upper()
        for e in entries:
            if e["code"] == cand:
                return cand
    return ""


def get_tri_prefix(code: str, entries: List[Dict]) -> str:
    """Return '中英X' trilingual prefix for a language code (e.g. GR -> '中英腊', MY -> '中英缅', ES -> '中英西')."""
    c = (code or "").strip().upper()
    if not c or c == "AUTO":
        return "中英外"
    for e in entries:
        if e["code"] == c:
            explicit = str(e.get("tri_prefix", "")).strip()
            if explicit:
                return explicit
            tri_kws = [kw.strip() for kw in e.get("keywords", []) if kw.strip().startswith("中英")]
            if tri_kws:
                # Prefer concise 3-char '中英X' (e.g. '中英腊', '中英缅', '中英马') when available
                tri_kws.sort(key=len)
                return tri_kws[0]
            nm = e.get("name", "").strip()
            if nm.endswith("语") and len(nm) >= 2:
                core = nm[:-1]
                return f"中英{core}" if len(core) <= 2 else f"中英{core[0]}"
            return f"中英{c}"
    return f"中英{c}"
