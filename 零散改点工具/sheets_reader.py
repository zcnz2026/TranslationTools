"""
Google Sheets (Public CSV + OAuth 2.0 PKCE API v4) & Local Excel (.xlsx / .csv) Reader
for Scenario 3 (改点入库 + 通读定稿入库).

Security guarantees:
  - Pure Python stdlib (zero extra third-party build dependencies).
  - Bounded ZIP reads + XML DTD rejection for local `.xlsx` files.
  - OAuth 2.0 Desktop PKCE (S256 + CSRF state nonce + ephemeral 127.0.0.1 loopback).
  - OAuth client credentials & refresh tokens encrypted at rest via Windows DPAPI.
  - Strict `gid` enforcement: ONLY the sheet tab referenced by the pasted link's `gid` is read.
"""

from __future__ import annotations

import base64
import csv
from dataclasses import dataclass, field
import hashlib
from http.server import BaseHTTPRequestHandler, HTTPServer
import io
import json
from pathlib import Path
import re
import secrets
import sys
import threading
import time
from typing import Callable, Dict, List, Optional, Sequence, Tuple
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
import xml.etree.ElementTree as ET
import zipfile


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
from security_utils import dpapi_decrypt_str, dpapi_encrypt_str  # noqa: E402

GOOGLE_OAUTH_AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
GOOGLE_OAUTH_TOKEN_URL = "https://oauth2.googleapis.com/token"
GOOGLE_SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets.readonly"
GOOGLE_USERINFO_URL = "https://www.googleapis.com/oauth2/v3/userinfo"
MAX_HTTP_BYTES = 25 * 1024 * 1024  # 25 MiB cap for sheet downloads

_GSHEET_ID_RE = re.compile(r"/spreadsheets/(?:u/\d+/)?d/([a-zA-Z0-9_-]{15,})")
_GSHEET_GID_RE = re.compile(r"(?:[?#&]|#)gid=(\d+)")
_CELL_REF_COL_RE = re.compile(r"^([A-Za-z]+)")

DEFAULT_COLUMN_ALIASES: Dict[str, List[str]] = {
    "cn": ["中文原文", "Chinese", "中文", "原文", "CN"],
    "en": ["英文", "English", "英语", "EN"],
    "target": ["目标语言", "Target language", "Target Language", "目标语", "外文", "译文", "菲语"],
    "changes": ["改动列", "改点列", "Final Version-改动行", "改动行", "修改列"],
    "final": ["定稿列", "定稿", "Final Version-全", "Final Version", "最终定稿"],
}


@dataclass
class ColumnMappingConfig:
    """User-customizable header name aliases (comma/semicolon/slash/newline separated)."""
    cn: str = "中文原文, Chinese, 中文, 原文"
    en: str = "英文, English, 英语"
    target: str = "目标语言, Target language, 目标语, 外文, 译文, 菲语"
    changes: str = "改动列, 改点列, Final Version-改动行, 改动行"
    final: str = "定稿列, 定稿, Final Version-全, Final Version"

    def aliases_for(self, key: str) -> List[str]:
        raw = str(getattr(self, key, "") or "").strip()
        parts = [p.strip() for p in re.split(r"[,，;；/\n]+", raw) if p.strip()]
        defaults = DEFAULT_COLUMN_ALIASES.get(key, [])
        out: List[str] = []
        seen = set()
        for item in (*parts, *defaults):
            low = _norm_header(item)
            if low and low not in seen:
                seen.add(low)
                out.append(item)
        return out

    def to_dict(self) -> Dict[str, str]:
        return {
            "cn": self.cn,
            "en": self.en,
            "target": self.target,
            "changes": self.changes,
            "final": self.final,
        }

    @classmethod
    def from_dict(cls, data: Optional[dict]) -> "ColumnMappingConfig":
        if not isinstance(data, dict):
            return cls()
        return cls(
            cn=str(data.get("cn") or cls.cn),
            en=str(data.get("en") or cls.en),
            target=str(data.get("target") or cls.target),
            changes=str(data.get("changes") or cls.changes),
            final=str(data.get("final") or cls.final),
        )


@dataclass
class ProofreadRow:
    row_number: int
    cn: str
    en: str
    target: str
    changes: str
    final: str
    source_label: str = ""
    sheet_name: str = ""


@dataclass
class ParsedSheetData:
    source: str
    sheet_name: str
    header_row_idx: int
    column_indices: Dict[str, int]
    rows: List[ProofreadRow] = field(default_factory=list)


def _norm_header(s: str) -> str:
    return re.sub(r"[\s\u3000_:：\-—()（）【】\[\]]+", "", (s or "").strip()).lower()


def clean_cell_spaces(text: str) -> str:
    """Normalize redundant spaces inside a spreadsheet cell without collapsing intentional newlines."""
    if not text:
        return ""
    lines = str(text).replace("\r\n", "\n").replace("\r", "\n").split("\n")
    cleaned = [re.sub(r"[ \t\u00A0]{2,}", " ", ln).strip() for ln in lines]
    return "\n".join(cleaned).strip()


# ---------------------------------------------------------------------------
# Google Sheets URL & Batch Title-Link Parser
# ---------------------------------------------------------------------------

def is_google_sheet_url(text: str) -> bool:
    s = (text or "").strip()
    return s.startswith(("https://docs.google.com/spreadsheets/", "http://docs.google.com/spreadsheets/"))


def parse_google_sheet_url(url: str) -> Tuple[str, str]:
    """
    Extract `(spreadsheet_id, gid)` from a Google Sheets URL.
    Raises ValueError if the URL is not a valid Google Sheets link.
    """
    u = (url or "").strip()
    m_id = _GSHEET_ID_RE.search(u)
    if not m_id:
        raise ValueError(f"无法识别 Google Sheets 表格链接：{u}")
    sheet_id = m_id.group(1)
    m_gid = _GSHEET_GID_RE.search(u)
    gid = m_gid.group(1) if m_gid else "0"
    return sheet_id, gid


def parse_title_link_table(raw_text: str) -> List[Tuple[str, List[str]]]:
    """
    Parse pasted 2-column text (`左列篇目标题` + `右列通读表格链接或本地路径`) into ordered
    `[(chapter_title, [link_or_path_1, link_or_path_2, ...]), ...]`.

    Supports:
      - Excel / Google Sheets copy-paste (Tab-separated: `Title\\tURL`)
      - Multiple rows for the same title (e.g. Part 1..7), OR blank left title inheriting the previous title
      - Multiple URLs on the right separated by spaces, commas, or semicolons
      - Pipe `|` separated lines (`Title | URL`)
    """
    grouped: List[Tuple[str, List[str]]] = []
    title_idx_map: Dict[str, int] = {}
    last_title = ""

    for raw_line in (raw_text or "").splitlines():
        line = raw_line.strip()
        if not line:
            continue

        title_part = ""
        right_part = ""
        if "\t" in raw_line:
            parts = raw_line.split("\t")
            title_part = parts[0].strip()
            right_part = " ".join(p.strip() for p in parts[1:] if p.strip())
        elif "|" in line and not line.startswith(("http://", "https://")):
            left, right = line.split("|", 1)
            title_part = left.strip()
            right_part = right.strip()
        else:
            m_http = re.search(r"(https?://\S+)", line)
            if m_http:
                title_part = line[: m_http.start()].strip(" \t:：-—|")
                right_part = line[m_http.start() :].strip()
            elif line.lower().endswith((".xlsx", ".csv")):
                title_part = ""
                right_part = line
            else:
                # Standalone title line (links may follow on next lines)
                last_title = line
                continue

        if title_part:
            last_title = title_part
        eff_title = title_part or last_title

        # Extract URLs or local file paths from right_part
        urls = re.findall(r"https?://[^\s,，;；]+", right_part)
        sources: List[str] = []
        if urls:
            sources.extend(urls)
        else:
            for cand in re.split(r"[;；\n]+", right_part):
                c = cand.strip().strip('"').strip("'")
                if c:
                    sources.append(c)

        if not sources:
            continue
        if not eff_title:
            eff_title = Path(sources[0]).stem if not is_google_sheet_url(sources[0]) else "未命名篇目"

        t_key = re.sub(r"\s+", " ", eff_title).strip()
        if t_key not in title_idx_map:
            title_idx_map[t_key] = len(grouped)
            grouped.append((eff_title, []))
        idx = title_idx_map[t_key]
        for s in sources:
            if s not in grouped[idx][1]:
                grouped[idx][1].append(s)

    return grouped


# ---------------------------------------------------------------------------
# Pure-Python Bounded & DTD-Safe .xlsx / .csv Reader
# ---------------------------------------------------------------------------

def _col_letters_to_index(letters: str) -> int:
    idx = 0
    for ch in letters.upper():
        if "A" <= ch <= "Z":
            idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return max(0, idx - 1)


def read_csv_rows(csv_text: str) -> List[List[str]]:
    reader = csv.reader(io.StringIO(csv_text))
    return [[str(cell or "") for cell in row] for row in reader]


def list_xlsx_sheet_names(xlsx_path: Path) -> List[str]:
    p = Path(xlsx_path).resolve()
    with zipfile.ZipFile(p, "r") as zf:
        docx_xml_utils._validate_docx_zip(zf)
        wb_bytes = docx_xml_utils.read_docx_member_bounded(zf, "xl/workbook.xml")
        wb_root = docx_xml_utils.parse_safe_openxml_bytes(wb_bytes)
    names: List[str] = []
    for sh in wb_root.iter():
        if sh.tag.split("}")[-1] == "sheet":
            nm = sh.attrib.get("name", "").strip()
            if nm:
                names.append(nm)
    return names


def read_xlsx_rows(
    xlsx_path: Path,
    sheet_name: Optional[str] = None,
    preferred_hints: Sequence[str] = ("定稿", "Final", "final", "✅"),
) -> Tuple[str, List[List[str]]]:
    """
    Read a worksheet from `.xlsx` safely using pure Python stdlib.
    Returns `(chosen_sheet_name, rows_as_list_of_str_lists)`.
    """
    p = Path(xlsx_path).resolve()
    if not p.is_file():
        raise FileNotFoundError(f"找不到表格文件：{p}")

    if p.suffix.lower() == ".csv":
        raw = p.read_bytes()
        for enc in ("utf-8-sig", "utf-8", "gb18030", "utf-16"):
            try:
                return p.stem, read_csv_rows(raw.decode(enc))
            except UnicodeDecodeError:
                continue
        return p.stem, read_csv_rows(raw.decode("utf-8", errors="ignore"))

    with zipfile.ZipFile(p, "r") as zf:
        docx_xml_utils._validate_docx_zip(zf)
        names_set = set(zf.namelist())

        # 1. Shared strings table (optional)
        shared_strings: List[str] = []
        if "xl/sharedStrings.xml" in names_set:
            ss_bytes = docx_xml_utils.read_docx_member_bounded(zf, "xl/sharedStrings.xml")
            ss_root = docx_xml_utils.parse_safe_openxml_bytes(ss_bytes)
            for si in ss_root:
                if si.tag.split("}")[-1] != "si":
                    continue
                parts = [el.text or "" for el in si.iter() if el.tag.split("}")[-1] == "t"]
                shared_strings.append("".join(parts))

        # 2. Workbook sheets & relationships
        wb_bytes = docx_xml_utils.read_docx_member_bounded(zf, "xl/workbook.xml")
        wb_root = docx_xml_utils.parse_safe_openxml_bytes(wb_bytes)
        rels_map: Dict[str, str] = {}
        if "xl/_rels/workbook.xml.rels" in names_set:
            rels_bytes = docx_xml_utils.read_docx_member_bounded(zf, "xl/_rels/workbook.xml.rels")
            rels_root = docx_xml_utils.parse_safe_openxml_bytes(rels_bytes)
            for rel in rels_root:
                if rel.tag.split("}")[-1] == "Relationship":
                    rid = rel.attrib.get("Id", "")
                    tgt = rel.attrib.get("Target", "")
                    if rid and tgt:
                        rels_map[rid] = tgt.lstrip("/")

        sheets_info: List[Tuple[str, str]] = []
        for idx_s, sh in enumerate(wb_root.iter()):
            if sh.tag.split("}")[-1] != "sheet":
                continue
            nm = sh.attrib.get("name", f"Sheet{idx_s + 1}")
            rid = ""
            for k, v in sh.attrib.items():
                if k.endswith("}id") or k == "r:id":
                    rid = v
                    break
            target = rels_map.get(rid, f"worksheets/sheet{idx_s + 1}.xml")
            if not target.startswith("xl/"):
                target = "xl/" + target.lstrip("/")
            sheets_info.append((nm, target))

        if not sheets_info:
            raise ValueError(f"Excel 文件中未找到任何工作表：{p.name}")

        chosen_name, chosen_target = sheets_info[0]
        if sheet_name:
            want = sheet_name.strip().lower()
            for nm, tgt in sheets_info:
                if nm.strip().lower() == want:
                    chosen_name, chosen_target = nm, tgt
                    break
        else:
            for hint in preferred_hints:
                hit = next(((nm, tgt) for nm, tgt in sheets_info if hint.lower() in nm.lower()), None)
                if hit is not None:
                    chosen_name, chosen_target = hit
                    break

        if chosen_target not in names_set:
            raise FileNotFoundError(f"工作表 XML 缺失：{chosen_target}")

        sh_bytes = docx_xml_utils.read_docx_member_bounded(zf, chosen_target)
        sh_root = docx_xml_utils.parse_safe_openxml_bytes(sh_bytes)

    rows_out: List[List[str]] = []
    for row_el in sh_root.iter():
        if row_el.tag.split("}")[-1] != "row":
            continue
        row_dict: Dict[int, str] = {}
        max_col = -1
        auto_col = 0
        for c_el in row_el:
            if c_el.tag.split("}")[-1] != "c":
                continue
            ref = c_el.attrib.get("r", "")
            m_col = _CELL_REF_COL_RE.match(ref)
            col_idx = _col_letters_to_index(m_col.group(1)) if m_col else auto_col
            auto_col = col_idx + 1
            ctype = c_el.attrib.get("t", "")

            val_str = ""
            if ctype == "inlineStr":
                val_str = "".join(el.text or "" for el in c_el.iter() if el.tag.split("}")[-1] == "t")
            else:
                v_el = next((ch for ch in c_el if ch.tag.split("}")[-1] == "v"), None)
                raw_v = (v_el.text or "") if v_el is not None else ""
                if ctype == "s":
                    try:
                        s_i = int(raw_v)
                        val_str = shared_strings[s_i] if 0 <= s_i < len(shared_strings) else ""
                    except ValueError:
                        val_str = ""
                elif ctype == "b":
                    val_str = "TRUE" if raw_v == "1" else "FALSE"
                else:
                    val_str = raw_v

            row_dict[col_idx] = val_str
            if col_idx > max_col:
                max_col = col_idx

        if max_col < 0:
            rows_out.append([])
        else:
            rows_out.append([row_dict.get(i, "") for i in range(max_col + 1)])

    return chosen_name, rows_out


# ---------------------------------------------------------------------------
# Google OAuth 2.0 Desktop PKCE Manager & Google Sheets Fetcher
# ---------------------------------------------------------------------------

def _http_read_bounded(req: urllib.request.Request, timeout: float = 30.0) -> Tuple[int, dict, bytes]:
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        status = getattr(resp, "status", 200)
        headers = {k.lower(): v for k, v in resp.headers.items()}
        chunks: List[bytes] = []
        total = 0
        while True:
            buf = resp.read(65536)
            if not buf:
                break
            total += len(buf)
            if total > MAX_HTTP_BYTES:
                raise ValueError(f"远程表格响应超过上限 ({MAX_HTTP_BYTES // (1024 * 1024)} MB)")
            chunks.append(buf)
        return status, headers, b"".join(chunks)


class GoogleOAuthManager:
    """
    Desktop OAuth 2.0 (JSON credentials import + PKCE S256 browser login) manager for Google Sheets.
    Supports importing Google Cloud `credentials.json` / `client_secret_*.json` (`installed` or `web`)
    as well as pre-authorized `token.json`. All secrets & tokens are encrypted via Windows DPAPI.
    """

    def __init__(
        self,
        client_id: str = "",
        client_secret: str = "",
        access_token: str = "",
        refresh_token: str = "",
        expires_at: float = 0.0,
        user_email: str = "",
        json_filename: str = "",
        project_id: str = "",
        auth_uri: str = "",
        token_uri: str = "",
        redirect_uris: Optional[List[str]] = None,
    ):
        self.client_id = (client_id or "").strip()
        self.client_secret = (client_secret or "").strip()
        self.access_token = (access_token or "").strip()
        self.refresh_token = (refresh_token or "").strip()
        self.expires_at = float(expires_at or 0.0)
        self.user_email = (user_email or "").strip()
        self.json_filename = (json_filename or "").strip()
        self.project_id = (project_id or "").strip()
        self.auth_uri = (auth_uri or "").strip() or GOOGLE_OAUTH_AUTH_URL
        self.token_uri = (token_uri or "").strip() or GOOGLE_OAUTH_TOKEN_URL
        self.redirect_uris: List[str] = list(redirect_uris or [])
        self._lock = threading.Lock()

    @property
    def has_client_config(self) -> bool:
        return bool(self.client_id)

    @property
    def is_authorized(self) -> bool:
        return bool(self.refresh_token or (self.access_token and time.time() < self.expires_at - 30))

    def import_json_file(self, json_path: Union[str, Path]) -> dict:
        """
        Import a Google OAuth 2.0 `.json` credential file (`client_secret_*.json` / `credentials.json` / `token.json`).
        Returns summary dict `{"filename": ..., "project_id": ..., "client_id": ..., "already_authorized": bool}`.
        """
        p = Path(json_path).resolve()
        if not p.is_file():
            raise FileNotFoundError(f"找不到 JSON 授权文件：{p}")
        if p.stat().st_size > 2 * 1024 * 1024:
            raise ValueError("JSON 授权文件大小异常（超过 2MB），请确认文件正确。")

        raw_txt = p.read_text(encoding="utf-8-sig")
        try:
            data = json.loads(raw_txt)
        except Exception as ex:
            raise ValueError(f"JSON 文件格式无效：{ex}") from ex

        if not isinstance(data, dict):
            raise ValueError("JSON 授权文件顶层必须是对象格式。")

        if data.get("type") == "service_account":
            raise ValueError(
                "当前导入的是「服务账号 (service_account)」密钥 JSON。\n"
                "请在 Google Cloud Console → 凭据 → 创建「OAuth 客户端 ID（桌面应用 Desktop app）」并下载对应的 JSON 文件导入，即可弹窗登录个人 Google 账户。"
            )

        # Case 1: Standard Google Cloud Console OAuth client JSON (`installed` or `web`)
        block = data.get("installed") or data.get("web")
        if isinstance(block, dict):
            cid = str(block.get("client_id") or "").strip()
            csec = str(block.get("client_secret") or "").strip()
            proj = str(block.get("project_id") or "").strip()
            auri = str(block.get("auth_uri") or GOOGLE_OAUTH_AUTH_URL).strip()
            turi = str(block.get("token_uri") or GOOGLE_OAUTH_TOKEN_URL).strip()
            ruris = [str(u).strip() for u in (block.get("redirect_uris") or []) if str(u).strip()]
            if not cid:
                raise ValueError("该 JSON 文件缺少 client_id 字段。")
            with self._lock:
                # If switching to a different OAuth client, clear stale tokens from previous client
                if self.client_id and self.client_id != cid:
                    self.access_token = ""
                    self.refresh_token = ""
                    self.expires_at = 0.0
                    self.user_email = ""
                self.client_id = cid
                self.client_secret = csec
                self.project_id = proj
                self.auth_uri = auri
                self.token_uri = turi
                self.redirect_uris = ruris
                self.json_filename = p.name
            return {
                "filename": p.name,
                "project_id": proj,
                "client_id": cid,
                "already_authorized": self.is_authorized,
            }

        # Case 2: Pre-authorized `token.json` or flat `{client_id, client_secret, refresh_token, token}`
        cid = str(data.get("client_id") or "").strip()
        csec = str(data.get("client_secret") or "").strip()
        ref = str(data.get("refresh_token") or "").strip()
        acc = str(data.get("token") or data.get("access_token") or "").strip()
        turi = str(data.get("token_uri") or GOOGLE_OAUTH_TOKEN_URL).strip()
        if cid:
            with self._lock:
                self.client_id = cid
                self.client_secret = csec
                self.token_uri = turi
                self.json_filename = p.name
                if ref:
                    self.refresh_token = ref
                if acc:
                    self.access_token = acc
                    self.expires_at = time.time() + 1800
                if ref and not self.user_email:
                    self.user_email = "已通过 token.json 授权"
            return {
                "filename": p.name,
                "project_id": self.project_id,
                "client_id": cid,
                "already_authorized": self.is_authorized,
            }

        raise ValueError(
            "无法识别该 JSON 授权文件结构（未找到 installed / web / client_id 字段）。\n"
            "请确认导入的是从 Google Cloud 下载的 OAuth 2.0 客户端 JSON 文件。"
        )

    def to_dpapi_token(self) -> str:
        payload = json.dumps(
            {
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "access_token": self.access_token,
                "refresh_token": self.refresh_token,
                "expires_at": self.expires_at,
                "user_email": self.user_email,
                "json_filename": self.json_filename,
                "project_id": self.project_id,
                "auth_uri": self.auth_uri,
                "token_uri": self.token_uri,
                "redirect_uris": self.redirect_uris,
            },
            ensure_ascii=False,
        )
        return dpapi_encrypt_str(payload)

    @classmethod
    def from_dpapi_token(cls, enc_token: str) -> "GoogleOAuthManager":
        if not enc_token:
            return cls()
        try:
            raw = dpapi_decrypt_str(enc_token)
            if not raw:
                return cls()
            data = json.loads(raw)
            if isinstance(data, dict):
                return cls(
                    client_id=data.get("client_id", ""),
                    client_secret=data.get("client_secret", ""),
                    access_token=data.get("access_token", ""),
                    refresh_token=data.get("refresh_token", ""),
                    expires_at=data.get("expires_at", 0.0),
                    user_email=data.get("user_email", ""),
                    json_filename=data.get("json_filename", ""),
                    project_id=data.get("project_id", ""),
                    auth_uri=data.get("auth_uri", ""),
                    token_uri=data.get("token_uri", ""),
                    redirect_uris=data.get("redirect_uris") or [],
                )
        except Exception:
            pass
        return cls()

    def clear_tokens(self, clear_client_json: bool = False) -> None:
        with self._lock:
            self.access_token = ""
            self.refresh_token = ""
            self.expires_at = 0.0
            self.user_email = ""
            if clear_client_json:
                self.client_id = ""
                self.client_secret = ""
                self.json_filename = ""
                self.project_id = ""
                self.redirect_uris = []

    def authorize_interactive(
        self,
        client_id: str = "",
        client_secret: str = "",
        timeout_sec: float = 180.0,
        log_fn: Optional[Callable[[str], None]] = None,
    ) -> str:
        """Run interactive browser OAuth 2.0 PKCE login using the imported JSON config and return authorized email."""
        cid = (client_id or self.client_id or "").strip()
        csec = (client_secret or self.client_secret or "").strip()
        if not cid:
            raise ValueError("请先导入 Google OAuth 2.0 JSON 授权文件（credentials.json / client_secret_*.json）。")

        code_verifier = secrets.token_urlsafe(64)
        digest = hashlib.sha256(code_verifier.encode("ascii")).digest()
        code_challenge = base64.urlsafe_b64encode(digest).rstrip(b"=").decode("ascii")
        state = secrets.token_urlsafe(24)

        result_holder: Dict[str, str] = {}
        done_ev = threading.Event()

        class _CallbackHandler(BaseHTTPRequestHandler):
            def log_message(self, format, *args):  # noqa: A002
                return

            def do_GET(self):  # noqa: N802
                parsed = urllib.parse.urlparse(self.path)
                if parsed.path == "/favicon.ico":
                    self.send_response(204)
                    self.end_headers()
                    return
                qs = urllib.parse.parse_qs(parsed.query)
                if qs.get("state", [""])[0] != state:
                    result_holder["error"] = "OAuth state 校验失败，请重试。"
                    self.send_response(400)
                    self.send_header("Content-Type", "text/html; charset=utf-8")
                    self.end_headers()
                    self.wfile.write("<h3>授权校验失败（state 不匹配），请关闭页面重试。</h3>".encode("utf-8"))
                    done_ev.set()
                    return
                if "error" in qs:
                    result_holder["error"] = f"Google 授权被拒绝：{qs['error'][0]}"
                    self.send_response(200)
                    self.send_header("Content-Type", "text/html; charset=utf-8")
                    self.end_headers()
                    self.wfile.write("<h3>已取消授权，您可以关闭此浏览器标签页。</h3>".encode("utf-8"))
                    done_ev.set()
                    return
                code = qs.get("code", [""])[0]
                if not code:
                    self.send_response(200)
                    self.end_headers()
                    return
                result_holder["code"] = code
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(
                    "<html><body style='font-family:Microsoft YaHei,sans-serif;padding:32px;'>"
                    "<h2 style='color:#15803d;'>✅ Google 账户授权成功！</h2>"
                    "<p>您可以直接关闭此浏览器页面并返回「多语种整理工作台」。</p>"
                    "</body></html>".encode("utf-8")
                )
                done_ev.set()

        # Respect explicit fixed port/path if `web` JSON specified one (e.g. http://localhost:8080/callback);
        # otherwise use standard Desktop loopback `http://127.0.0.1:{ephemeral_port}` (RFC 8252).
        bind_port = 0
        redirect_uri = ""
        for cand_u in self.redirect_uris:
            pu = urllib.parse.urlparse(cand_u)
            if pu.hostname in ("localhost", "127.0.0.1") and pu.port:
                bind_port = int(pu.port)
                redirect_uri = cand_u
                break

        httpd = HTTPServer(("127.0.0.1", bind_port), _CallbackHandler)
        port = httpd.server_address[1]
        if not redirect_uri:
            redirect_uri = f"http://127.0.0.1:{port}"

        srv_thread = threading.Thread(target=httpd.serve_forever, daemon=True)
        srv_thread.start()
        try:
            auth_endpoint = self.auth_uri or GOOGLE_OAUTH_AUTH_URL
            token_endpoint = self.token_uri or GOOGLE_OAUTH_TOKEN_URL
            params = {
                "client_id": cid,
                "redirect_uri": redirect_uri,
                "response_type": "code",
                "scope": f"{GOOGLE_SHEETS_SCOPE} openid email",
                "code_challenge": code_challenge,
                "code_challenge_method": "S256",
                "access_type": "offline",
                "prompt": "consent",
                "state": state,
            }
            auth_url = f"{auth_endpoint}?{urllib.parse.urlencode(params)}"
            if log_fn:
                log_fn("🌐 正在打开浏览器进行 Google 账户授权...")
            webbrowser.open(auth_url)

            if not done_ev.wait(timeout=timeout_sec):
                raise TimeoutError("等待浏览器 Google 授权超时，请重试。")
            if result_holder.get("error"):
                raise RuntimeError(result_holder["error"])
            code = result_holder.get("code", "")
            if not code:
                raise RuntimeError("未收到 Google 授权回调 code。")

            token_form = {
                "client_id": cid,
                "code": code,
                "code_verifier": code_verifier,
                "grant_type": "authorization_code",
                "redirect_uri": redirect_uri,
            }
            if csec:
                token_form["client_secret"] = csec

            req = urllib.request.Request(
                token_endpoint,
                data=urllib.parse.urlencode(token_form).encode("utf-8"),
                headers={"Content-Type": "application/x-www-form-urlencoded"},
                method="POST",
            )
            _, _, body = _http_read_bounded(req, timeout=30.0)
            tok_data = json.loads(body.decode("utf-8"))
            acc = str(tok_data.get("access_token") or "")
            ref = str(tok_data.get("refresh_token") or "")
            exp_in = float(tok_data.get("expires_in") or 3600)
            if not acc:
                raise RuntimeError(f"换取 Google Access Token 失败：{tok_data}")

            email = ""
            try:
                u_req = urllib.request.Request(
                    GOOGLE_USERINFO_URL,
                    headers={"Authorization": f"Bearer {acc}"},
                )
                _, _, u_body = _http_read_bounded(u_req, timeout=15.0)
                email = str(json.loads(u_body.decode("utf-8")).get("email") or "")
            except Exception:
                email = ""

            with self._lock:
                self.client_id = cid
                self.client_secret = csec
                self.access_token = acc
                if ref:
                    self.refresh_token = ref
                self.expires_at = time.time() + exp_in
                self.user_email = email or "已授权 Google 账户"
            return self.user_email
        finally:
            try:
                httpd.shutdown()
                httpd.server_close()
            except Exception:
                pass

    def get_valid_access_token(self) -> str:
        with self._lock:
            if self.access_token and time.time() < self.expires_at - 60:
                return self.access_token
            if not self.refresh_token or not self.client_id:
                return ""
            token_endpoint = self.token_uri or GOOGLE_OAUTH_TOKEN_URL
            form = {
                "client_id": self.client_id,
                "refresh_token": self.refresh_token,
                "grant_type": "refresh_token",
            }
            if self.client_secret:
                form["client_secret"] = self.client_secret
            req = urllib.request.Request(
                token_endpoint,
                data=urllib.parse.urlencode(form).encode("utf-8"),
                headers={"Content-Type": "application/x-www-form-urlencoded"},
                method="POST",
            )
            _, _, body = _http_read_bounded(req, timeout=25.0)
            tok_data = json.loads(body.decode("utf-8"))
            acc = str(tok_data.get("access_token") or "")
            exp_in = float(tok_data.get("expires_in") or 3600)
            if acc:
                self.access_token = acc
                self.expires_at = time.time() + exp_in
            return self.access_token


def fetch_google_sheet_rows(
    url: str,
    oauth: Optional[GoogleOAuthManager] = None,
    timeout: float = 30.0,
) -> Tuple[str, List[List[str]]]:
    """
    Fetch ONLY the worksheet tab matching the URL's `gid` from Google Sheets.
    Uses OAuth 2.0 Sheets API v4 when authorized, otherwise public CSV export (`export?format=csv&gid=...`).
    Returns `(sheet_title, rows)`.
    """
    spreadsheet_id, gid = parse_google_sheet_url(url)

    # 1. Try Google Sheets API v4 via OAuth 2.0 if authorized
    if oauth is not None and oauth.is_authorized:
        try:
            token = oauth.get_valid_access_token()
            if token:
                meta_url = (
                    f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}"
                    "?fields=sheets(properties(sheetId,title))"
                )
                req_meta = urllib.request.Request(
                    meta_url,
                    headers={"Authorization": f"Bearer {token}"},
                )
                _, _, meta_bytes = _http_read_bounded(req_meta, timeout=timeout)
                meta = json.loads(meta_bytes.decode("utf-8"))
                target_title = ""
                for sh in meta.get("sheets") or []:
                    props = sh.get("properties") or {}
                    if str(props.get("sheetId", "")) == str(gid):
                        target_title = str(props.get("title") or "")
                        break
                if not target_title and meta.get("sheets"):
                    target_title = str((meta["sheets"][0].get("properties") or {}).get("title") or "Sheet1")

                quoted_rng = urllib.parse.quote(f"'{target_title}'", safe="")
                val_url = (
                    f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}"
                    f"/values/{quoted_rng}?valueRenderOption=FORMATTED_VALUE"
                )
                req_val = urllib.request.Request(
                    val_url,
                    headers={"Authorization": f"Bearer {token}"},
                )
                _, _, val_bytes = _http_read_bounded(req_val, timeout=timeout)
                val_data = json.loads(val_bytes.decode("utf-8"))
                raw_vals = val_data.get("values") or []
                return target_title, [[str(c or "") for c in r] for r in raw_vals]
        except Exception as oauth_ex:
            # Fall through to public CSV export if OAuth fails (e.g. publicly shared doc)
            _ = oauth_ex

    # 2. Public / Link-shared CSV export for exact `gid`
    export_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid}"
    req = urllib.request.Request(
        export_url,
        headers={"User-Agent": "Mozilla/5.0 ScatteredRevisionTool/1.9"},
    )
    try:
        _, headers, body = _http_read_bounded(req, timeout=timeout)
    except urllib.error.HTTPError as he:
        if he.code in (401, 403):
            raise PermissionError(
                f"无法直接访问该 Google Sheets（HTTP {he.code}）：请开启「知道链接的任何人可查看」权限，或在界面点击「Google OAuth 2.0 授权登录」。\n链接：{url}"
            ) from he
        raise RuntimeError(f"下载 Google Sheets 失败（HTTP {he.code}）：{url}") from he

    ctype = headers.get("content-type", "").lower()
    text = body.decode("utf-8-sig", errors="replace")
    if "text/html" in ctype or text.lstrip().lower().startswith(("<!doctype html", "<html")):
        raise PermissionError(
            f"该 Google Sheets 需要登录权限才能查看（返回了登录页面）。请开启链接查看权限，或在界面完成「Google OAuth 2.0 授权登录」。\n链接：{url}"
        )
    return f"gid={gid}", read_csv_rows(text)


# ---------------------------------------------------------------------------
# Column Detection & Structured Proofread Table Parsing
# ---------------------------------------------------------------------------

def detect_sheet_columns(
    rows: Sequence[Sequence[str]],
    col_config: Optional[ColumnMappingConfig] = None,
) -> Tuple[int, Dict[str, int]]:
    """
    Scan the top 15 rows to locate the header row and map logical columns:
    `cn`, `en`, `target`, `changes`, `final`.
    """
    cfg = col_config or ColumnMappingConfig()
    alias_sets = {
        k: {_norm_header(a) for a in cfg.aliases_for(k) if _norm_header(a)}
        for k in ("cn", "en", "target", "changes", "final")
    }

    best_row = -1
    best_map: Dict[str, int] = {}
    best_score = -1

    scan_limit = min(len(rows), 15)
    for r_i in range(scan_limit):
        row = rows[r_i]
        cur_map: Dict[str, int] = {}
        for c_i, cell in enumerate(row):
            nh = _norm_header(cell)
            if not nh:
                continue
            for key in ("changes", "final", "cn", "en", "target"):
                if key in cur_map:
                    continue
                if nh in alias_sets[key]:
                    cur_map[key] = c_i
                    break
            else:
                # Substring fallback for composite headers like "Final Version-改动行" or "目标语言(菲语)"
                for key in ("changes", "final", "target", "cn", "en"):
                    if key in cur_map:
                        continue
                    if any(a and len(a) >= 2 and a in nh for a in alias_sets[key]):
                        cur_map[key] = c_i
                        break

        # Smart fallback: if `en` and (`final` or `changes`) are found, and `target` is missing,
        # the column right after `en` is usually the language-named target column (e.g. "法语", "高棉语").
        if "en" in cur_map and "target" not in cur_map:
            cand_c = cur_map["en"] + 1
            if cand_c < len(row) and cand_c not in cur_map.values() and (row[cand_c] or "").strip():
                cur_map["target"] = cand_c

        score = len(cur_map)
        if score > best_score and score >= 2:
            best_score = score
            best_row = r_i
            best_map = cur_map

    if best_row < 0:
        raise ValueError(
            "未能在表格前 15 行识别出通读列名（如 中文原文 / 英文 / 目标语言 / 改动列 / 定稿列）。"
            "请在界面「自定义列名」中核对并补充表头名称。"
        )
    return best_row, best_map


def load_proofread_table(
    source: str,
    col_config: Optional[ColumnMappingConfig] = None,
    oauth: Optional[GoogleOAuthManager] = None,
    sheet_name: Optional[str] = None,
) -> ParsedSheetData:
    """
    Load either a Google Sheets URL (strictly by `gid`) or a local `.xlsx`/`.csv` file,
    clean redundant spaces, and return structured `ParsedSheetData`.
    """
    src = (source or "").strip()
    if is_google_sheet_url(src):
        sh_name, raw_rows = fetch_google_sheet_rows(src, oauth=oauth)
    else:
        sh_name, raw_rows = read_xlsx_rows(Path(src), sheet_name=sheet_name)

    hdr_idx, col_map = detect_sheet_columns(raw_rows, col_config=col_config)

    def _cell(row: Sequence[str], key: str) -> str:
        ci = col_map.get(key)
        if ci is None or ci >= len(row):
            return ""
        return clean_cell_spaces(row[ci])

    parsed_rows: List[ProofreadRow] = []
    for r_i in range(hdr_idx + 1, len(raw_rows)):
        row = raw_rows[r_i]
        cn_v = _cell(row, "cn")
        en_v = _cell(row, "en")
        tgt_v = _cell(row, "target")
        chg_v = _cell(row, "changes")
        fin_v = _cell(row, "final")
        if not any((cn_v, en_v, tgt_v, chg_v, fin_v)):
            continue
        parsed_rows.append(
            ProofreadRow(
                row_number=r_i + 1,
                cn=cn_v,
                en=en_v,
                target=tgt_v,
                changes=chg_v,
                final=fin_v,
                source_label=src,
                sheet_name=sh_name,
            )
        )

    return ParsedSheetData(
        source=src,
        sheet_name=sh_name,
        header_row_idx=hdr_idx,
        column_indices=col_map,
        rows=parsed_rows,
    )
