import base64
import ctypes
from ctypes import wintypes
import hashlib
import os
from pathlib import Path
import re
import shutil
import sys
import tempfile
import time
import urllib.parse

ALLOWED_API_HOSTS = frozenset({
    "aiplatform.googleapis.com",
    "generativelanguage.googleapis.com",
})

_MODEL_ID_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$")
_KEY_PATTERNS = [
    re.compile(r"AIza[0-9A-Za-z_-]{16,}"),
    re.compile(r"AQ\.[0-9A-Za-z._-]{16,}"),
    re.compile(r"ya29\.[0-9A-Za-z._-]{16,}"),
    re.compile(r"(?i)(x-goog-api-key|authorization)\s*[:=]\s*['\"]?[^\s,'\"]+"),
]

_DPAPI_PREFIX = "dpapi:v1:"
_DPAPI_ENTROPY = b"ScatteredRevisionTool::APIKey::v1"
_CRYPTPROTECT_UI_FORBIDDEN = 0x01
_FILE_ATTRIBUTE_REPARSE_POINT = 0x0400


class _DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ("cbData", wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_char)),
    ]


def _make_blob(data: bytes):
    buf = ctypes.create_string_buffer(data, len(data))
    blob = _DATA_BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    return blob, buf


def dpapi_encrypt_str(plaintext: str) -> str:
    """Encrypt string with Windows CurrentUser DPAPI (bound to current Windows login account)."""
    s = (plaintext or "").strip()
    if not s:
        return ""
    if sys.platform != "win32":
        raise RuntimeError("DPAPI encryption is only supported on Windows.")

    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    crypt32.CryptProtectData.argtypes = [
        ctypes.POINTER(_DATA_BLOB),
        wintypes.LPCWSTR,
        ctypes.POINTER(_DATA_BLOB),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(_DATA_BLOB),
    ]
    crypt32.CryptProtectData.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree.restype = ctypes.c_void_p

    in_blob, _in_buf = _make_blob(s.encode("utf-8"))
    ent_blob, _ent_buf = _make_blob(_DPAPI_ENTROPY)
    out_blob = _DATA_BLOB()

    ok = crypt32.CryptProtectData(
        ctypes.byref(in_blob),
        "ScatteredRevisionTool",
        ctypes.byref(ent_blob),
        None,
        None,
        _CRYPTPROTECT_UI_FORBIDDEN,
        ctypes.byref(out_blob),
    )
    if not ok:
        raise OSError("Windows DPAPI CryptProtectData failed.")
    try:
        raw = ctypes.string_at(out_blob.pbData, out_blob.cbData)
        return _DPAPI_PREFIX + base64.b64encode(raw).decode("ascii")
    finally:
        if out_blob.pbData:
            kernel32.LocalFree(ctypes.cast(out_blob.pbData, ctypes.c_void_p))


def dpapi_decrypt_str(token: str) -> str:
    """Decrypt a dpapi:v1:... token bound to current Windows user. Returns '' on failure/other machine."""
    t = (token or "").strip()
    if not t or not t.startswith(_DPAPI_PREFIX) or sys.platform != "win32":
        return ""
    try:
        enc_bytes = base64.b64decode(t[len(_DPAPI_PREFIX) :], validate=True)
    except Exception:
        return ""
    if not enc_bytes:
        return ""

    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    crypt32.CryptUnprotectData.argtypes = [
        ctypes.POINTER(_DATA_BLOB),
        ctypes.c_void_p,
        ctypes.POINTER(_DATA_BLOB),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(_DATA_BLOB),
    ]
    crypt32.CryptUnprotectData.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree.restype = ctypes.c_void_p

    in_blob, _in_buf = _make_blob(enc_bytes)
    ent_blob, _ent_buf = _make_blob(_DPAPI_ENTROPY)
    out_blob = _DATA_BLOB()

    ok = crypt32.CryptUnprotectData(
        ctypes.byref(in_blob),
        None,
        ctypes.byref(ent_blob),
        None,
        None,
        _CRYPTPROTECT_UI_FORBIDDEN,
        ctypes.byref(out_blob),
    )
    if not ok:
        return ""
    try:
        raw = ctypes.string_at(out_blob.pbData, out_blob.cbData)
        return raw.decode("utf-8")
    except Exception:
        return ""
    finally:
        if out_blob.pbData:
            kernel32.LocalFree(ctypes.cast(out_blob.pbData, ctypes.c_void_p))


def load_api_key_from_env(app_dir: Path) -> str:
    """Load API key from environment variables or local .env file if configured."""
    for env_name in ("VERTEX_API_KEY", "GEMINI_API_KEY", "SCATTERED_REV_API_KEY"):
        val = os.environ.get(env_name, "").strip()
        if val:
            return val

    dotenv_path = app_dir / ".env"
    if dotenv_path.is_file():
        try:
            for line in dotenv_path.read_text(encoding="utf-8", errors="ignore").splitlines():
                s = line.strip()
                if not s or s.startswith("#") or "=" not in s:
                    continue
                k, v = s.split("=", 1)
                k = k.strip()
                v = v.strip().strip("'\"")
                if k in ("VERTEX_API_KEY", "GEMINI_API_KEY", "GOOGLE_API_KEY", "SCATTERED_REV_API_KEY") and v:
                    return v
        except Exception:
            pass
    return ""


def validate_model_id(model: str) -> str:
    """Strictly validate model identifier to block URL path traversal / query / CRLF injection."""
    m = (model or "").strip()
    if not _MODEL_ID_RE.fullmatch(m) or ".." in m:
        raise ValueError(f"非法的模型名称格式：{m!r}（仅允许字母、数字、点、下划线与连字符）。")
    return m


_REGIONAL_VERTEX_HOST_RE = re.compile(r"^[a-z0-9-]+-aiplatform\.googleapis\.com$")


def validate_https_endpoint(url: str) -> None:
    """Ensure outbound API requests only target allowlisted Google HTTPS endpoints."""
    parsed = urllib.parse.urlsplit(url)
    host = (parsed.hostname or "").lower()
    is_trusted_host = host in ALLOWED_API_HOSTS or bool(_REGIONAL_VERTEX_HOST_RE.fullmatch(host))
    if parsed.scheme != "https" or not is_trusted_host:
        raise ValueError(f"已拦截非受信任的 API 请求地址：{url!r}")


def parse_api_keys(raw_input) -> list[str]:
    """
    Parse an unlimited number of API keys from multi-line text, comma/semicolon/pipe-separated
    strings, or an iterable of strings. Supports '# comment' remarks and deduplicates in order.
    """
    if not raw_input:
        return []
    if isinstance(raw_input, (list, tuple, set)):
        raw_text = "\n".join(str(x) for x in raw_input if x)
    else:
        raw_text = str(raw_input)

    keys: list[str] = []
    seen: set[str] = set()
    for raw_line in raw_text.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if "#" in line:
            line = line.split("#", 1)[0].strip()
        for tok in re.split(r"[\s,，;；|]+", line):
            clean = tok.strip().strip("'\"")
            if not clean:
                continue
            if any(ch in clean for ch in ("\r", "\n", "\x00")):
                continue
            if clean not in seen:
                seen.add(clean)
                keys.append(clean)
    return keys


def mask_key_label(idx: int, key: str) -> str:
    """Return a safe display label for a key (e.g. 'Key #1(尾号9xK2)') without leaking key prefix/body."""
    k = (key or "").strip()
    if len(k) >= 12:
        return f"Key #{idx + 1}(尾号{k[-4:]})"
    return f"Key #{idx + 1}"


def redact_secrets(text: str, extra_secret="") -> str:
    """Mask single or multiple API keys and auth headers before displaying logs or error popups."""
    if not text:
        return ""
    out = str(text)
    for sec in parse_api_keys(extra_secret):
        if len(sec) >= 8 and sec in out:
            out = out.replace(sec, "[REDACTED_KEY]")
    for rx in _KEY_PATTERNS:
        out = rx.sub(
            lambda m: f"{m.group(1)}: [REDACTED_KEY]" if m.lastindex else "[REDACTED_KEY]",
            out,
        )
    return out


def _is_reparse_or_symlink(p: Path) -> bool:
    if p.is_symlink():
        return True
    if sys.platform == "win32" and p.exists():
        try:
            attrs = ctypes.windll.kernel32.GetFileAttributesW(str(p))
            if attrs != 0xFFFFFFFF and (attrs & _FILE_ATTRIBUTE_REPARSE_POINT):
                return True
        except Exception:
            pass
    return False


def _resolve_secure_work_root() -> Path:
    local_app = os.environ.get("LOCALAPPDATA", "").strip()
    base_root = Path(local_app) if local_app else Path(tempfile.gettempdir())
    tool_root = base_root / "ScatteredRevisionTool"
    app_root = tool_root / "work"

    for check_p in (tool_root, app_root):
        if check_p.exists() and _is_reparse_or_symlink(check_p):
            raise RuntimeError(f"检测到不安全的符号链接或重解析点工作目录，已拒绝访问：{check_p}")

    app_root.mkdir(parents=True, exist_ok=True)
    if _is_reparse_or_symlink(app_root):
        raise RuntimeError(f"工作目录校验异常（疑似重解析点）：{app_root}")
    return app_root


def _prune_stale_workdirs(app_root: Path, max_age_seconds: float = 6 * 3600) -> int:
    """Opportunistically remove sibling workdirs in `app_root` older than `max_age_seconds` (default 6h)."""
    if max_age_seconds <= 0 or not app_root.is_dir():
        return 0
    cutoff = time.time() - float(max_age_seconds)
    pruned = 0
    try:
        with os.scandir(app_root) as it:
            entries = list(it)
    except OSError:
        return 0

    for entry in entries:
        try:
            p = Path(entry.path)
            if not entry.is_dir(follow_symlinks=False) or _is_reparse_or_symlink(p):
                continue
            st = os.stat(p)
            latest_mtime = st.st_mtime
            if latest_mtime < cutoff:
                try:
                    with os.scandir(p) as sub_it:
                        for sub in sub_it:
                            try:
                                sub_mt = os.stat(sub.path).st_mtime
                                if sub_mt > latest_mtime:
                                    latest_mtime = sub_mt
                            except OSError:
                                pass
                except OSError:
                    pass
            if latest_mtime < cutoff:
                shutil.rmtree(p, ignore_errors=True)
                pruned += 1
        except OSError:
            continue
    return pruned


def get_secure_user_workdir(stem: str) -> Path:
    """
    Create and return a per-user private working directory under %LOCALAPPDATA%,
    rejecting symlink or NTFS junction reparse points (mitigates CWE-377 / CWE-59).
    """
    app_root = _resolve_secure_work_root()
    safe_stem = re.sub(r"[^\w\u4e00-\u9fff-]+", "_", stem or "").strip("_")[:36] or "doc"
    digest = hashlib.sha256((stem or "").encode("utf-8")).hexdigest()[:12]
    target = app_root / f"{safe_stem}_{digest}"

    if target.exists() and _is_reparse_or_symlink(target):
        raise RuntimeError(f"检测到不安全的符号链接或重解析点工作目录，已拒绝访问：{target}")

    target.mkdir(parents=True, exist_ok=True)
    if _is_reparse_or_symlink(target):
        raise RuntimeError(f"工作目录校验异常（疑似重解析点）：{target}")
    return target


def create_ephemeral_workdir(stem: str, max_age_hours: float = 6.0) -> Path:
    """
    Create and return a 100% fresh, unique per-run working directory under
    `%LOCALAPPDATA%\\ScatteredRevisionTool\\work\\` using timestamp + `os.urandom(4).hex()`,
    while opportunistically pruning stale sibling workdirs older than `max_age_hours` (default 6h).
    """
    app_root = _resolve_secure_work_root()
    _prune_stale_workdirs(app_root, max_age_seconds=max(0.0, float(max_age_hours)) * 3600.0)

    safe_stem = re.sub(r"[^\w\u4e00-\u9fff-]+", "_", stem or "").strip("_")[:28] or "doc"
    for _ in range(8):
        ts_str = time.strftime("%Y%m%d_%H%M%S")
        rand_hex = os.urandom(4).hex()
        target = app_root / f"{safe_stem}_{ts_str}_{rand_hex}"
        if target.exists():
            continue
        target.mkdir(parents=True, exist_ok=False)
        if _is_reparse_or_symlink(target):
            raise RuntimeError(f"工作目录校验异常（疑似重解析点）：{target}")
        return target

    raise RuntimeError(f"无法创建唯一临时工作目录：{app_root}")


def create_secure_update_stagedir(max_age_hours: float = 2.0) -> Path:
    """
    Create and return a per-user update staging directory under
    `%LOCALAPPDATA%\\ScatteredRevisionTool\\updates\\`, rejecting symlink or NTFS
    reparse points on every ancestor/target directory (mitigates CWE-377 / CWE-59).
    """
    local_app = os.environ.get("LOCALAPPDATA", "").strip()
    base_root = Path(local_app) if local_app else Path(tempfile.gettempdir())
    tool_root = base_root / "ScatteredRevisionTool"
    updates_root = tool_root / "updates"

    for check_p in (tool_root, updates_root):
        if check_p.exists() and _is_reparse_or_symlink(check_p):
            raise RuntimeError(f"检测到不安全的符号链接或重解析点更新目录，已拒绝访问：{check_p}")

    updates_root.mkdir(parents=True, exist_ok=True)
    if _is_reparse_or_symlink(updates_root):
        raise RuntimeError(f"更新暂存根目录校验异常（疑似重解析点）：{updates_root}")

    _prune_stale_workdirs(updates_root, max_age_seconds=max(0.0, float(max_age_hours)) * 3600.0)

    for _ in range(8):
        ts_str = time.strftime("%Y%m%d_%H%M%S")
        rand_hex = os.urandom(6).hex()
        target = updates_root / f"upd_{ts_str}_{rand_hex}"
        if target.exists():
            continue
        target.mkdir(parents=True, exist_ok=False)
        if _is_reparse_or_symlink(target):
            raise RuntimeError(f"更新暂存目录校验异常（疑似重解析点）：{target}")
        return target.resolve()

    raise RuntimeError(f"无法创建安全的更新暂存目录：{updates_root}")


def cleanup_stale_ephemeral_workdirs(max_age_seconds: float = 0.0) -> int:
    """Prune ephemeral workdirs under `%LOCALAPPDATA%\\ScatteredRevisionTool\\work`."""
    try:
        app_root = _resolve_secure_work_root()
        return _prune_stale_workdirs(app_root, max_age_seconds=max(0.0, float(max_age_seconds)))
    except Exception:
        return 0
