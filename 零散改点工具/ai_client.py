from concurrent.futures import ThreadPoolExecutor
import copy
import json
import re
import ssl
import threading
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from typing import Callable, Dict, List, Optional, Tuple, Union

from security_utils import (
    mask_key_label,
    parse_api_keys,
    redact_secrets,
    validate_https_endpoint,
    validate_model_id,
)

MAX_API_RESP_BYTES = 10 * 1024 * 1024  # 10 MB upper bound for JSON API responses

SUPPORTED_MODELS = [
    ("Gemini 3.5 Flash-Lite (推荐·极速高性价比) [gemini-3.5-flash-lite]", "gemini-3.5-flash-lite"),
    ("Gemini 3.6 Flash (新一代主力·高速均衡) [gemini-3.6-flash]", "gemini-3.6-flash"),
    ("Gemini 3.8 Flash (高阶推理·稳定版) [gemini-3.8-flash]", "gemini-3.8-flash"),
    ("Gemini 3.5 Flash (通用均衡) [gemini-3.5-flash]", "gemini-3.5-flash"),
    ("Gemini 3.1 Flash-Lite (超低成本) [gemini-3.1-flash-lite]", "gemini-3.1-flash-lite"),
    ("Gemini 3.1 Pro Preview (复杂深度推理) [gemini-3.1-pro-preview]", "gemini-3.1-pro-preview"),
    ("Gemini 4 Hydrogen (Gemini 4 极速轻量) [gemini-4-hydrogen]", "gemini-4-hydrogen"),
    ("Gemini 4 Argon (Gemini 4 高效推理) [gemini-4-argon]", "gemini-4-argon"),
    ("Gemini 4 Neon (Gemini 4 旗舰高速) [gemini-4-neon]", "gemini-4-neon"),
    ("Gemini 4 Carbon (Gemini 4 最强推理) [gemini-4-carbon]", "gemini-4-carbon"),
    ("Gemini 2.5 Flash (经典兼容版) [gemini-2.5-flash]", "gemini-2.5-flash"),
    ("Gemini 2.5 Pro (经典深度版) [gemini-2.5-pro]", "gemini-2.5-pro"),
]

_BUILTIN_MODEL_LABELS: Dict[str, str] = {mid: disp for disp, mid in SUPPORTED_MODELS}
_EXCLUDED_MODEL_SUBSTRINGS = (
    "embedding",
    "imagen",
    "veo",
    "tts",
    "transcribe",
    "live",
    "native-audio",
    "aqa",
    "nano-banana",
    "image",
    "chirp",
    "lyria",
    "robotics",
    "computer-use",
    "gemma",
)

API_MODES = [
    ("自动识别 (支持混填 Vertex 收费 Key 与 Gemini Key)", "auto"),
    ("Vertex AI 收费版 API Key (Google Cloud 计费通道)", "vertex"),
    ("Gemini API Key (AI Studio 通道)", "gemini"),
]

ROTATION_MODES = [
    ("均衡轮换 + 报错秒切 (推荐：多 Key 轮流分摊防超速)", "round_robin"),
    ("顺序主备 + 报错秒切 (优先第 1 个，超额/报错切下一个)", "failover"),
]

_ALREADY_FRIENDLY_SINGLE_RE = re.compile(r"^API Key 调用失败：(.+)$")
_ALREADY_FRIENDLY_POOL_RE = re.compile(r"^密钥池中全部\s+\d+\s+个\s+API Key\s+均调用失败：(.+)$")
_HTTP_STATUS_RE = re.compile(r"(?:\bHTTP\s+(\d{3})\b|\[HTTP\s+(\d{3})(?:/\d{3})?\]|\"code\"\s*:\s*(\d{3}))")


def _extract_status_code(err: Union[BaseException, str, None]) -> Optional[int]:
    if isinstance(err, BaseException):
        st = getattr(err, "http_status", None)
        if isinstance(st, int) and 100 <= st <= 599:
            return st
    raw = str(err or "")
    m = _HTTP_STATUS_RE.search(raw)
    if m:
        for grp in m.groups():
            if grp and grp.isdigit():
                return int(grp)
    return None


def format_friendly_api_error(err: Union[BaseException, str], total_keys: int = 1) -> str:
    """
    Convert any raw HTTP / JSON / network / Python exception into ONE concise,
    human-readable Chinese sentence without raw JSON blocks or tracebacks.
    """
    prefix = (
        f"密钥池中全部 {total_keys} 个 API Key 均调用失败："
        if total_keys and total_keys > 1
        else "API Key 调用失败："
    )
    raw = redact_secrets(str(err or "")).strip()
    first_line = raw.splitlines()[0].strip() if raw else ""

    # Idempotency: if already formatted by format_friendly_api_error, only adjust prefix for total_keys
    if "\n" not in raw and "{" not in raw and "Traceback" not in raw:
        m_single = _ALREADY_FRIENDLY_SINGLE_RE.match(first_line)
        if m_single:
            return f"{prefix}{m_single.group(1).strip()}"
        m_pool = _ALREADY_FRIENDLY_POOL_RE.match(first_line)
        if m_pool:
            return f"{prefix}{m_pool.group(1).strip()}"

    status_code = _extract_status_code(err)
    low = raw.lower()

    if (
        status_code in (401, 403)
        or "api_key_invalid" in low
        or "api key not valid" in low
        or "api_key_service_blocked" in low
        or "expected oauth2" in low
        or "api keys are not supported" in low
        or "unauthenticated" in low
        or "permission_denied" in low
        or "permission denied" in low
    ):
        code_tag = f"HTTP {status_code}" if status_code in (400, 401, 403) else "HTTP 401/403"
        reason = f"API Key 无效、未开通权限或选错了接口通道 [{code_tag}]"
    elif status_code == 400:
        if any(tok in low for tok in ("key", "auth", "credential", "token")):
            reason = "API Key 无效、未开通权限或选错了接口通道 [HTTP 400]"
        else:
            reason = "请求参数被接口拒绝或当前通道不兼容 [HTTP 400]"
    elif status_code == 429 or "resource_exhausted" in low or "quota" in low or "rate limit" in low or "too many requests" in low:
        reason = "API 额度已耗尽或请求过快 [HTTP 429]"
    elif status_code == 404 or "not_found" in low or "not found" in low:
        reason = "模型名称不存在或不可用 [HTTP 404]"
    elif (
        status_code in (500, 502, 503, 504)
        or "service unavailable" in low
        or "overloaded" in low
        or "high demand" in low
        or "internal error" in low
    ):
        code_tag = f"HTTP {status_code}" if status_code in (500, 502, 503, 504) else "HTTP 503"
        reason = f"官方模型服务器暂时繁忙 [{code_tag}]"
    elif "未返回有效候选结果" in raw or "blockreason" in low or "safety" in low:
        reason = "模型未返回有效结果（可能触发内容安全限制）"
    elif "超过安全上限" in raw:
        reason = "接口响应体积超过安全上限，已中止解析"
    elif "非法的模型名称" in raw:
        reason = "模型名称格式不合法，请检查模型配置"
    elif "非受信任的 api 请求地址" in low:
        reason = "请求地址不在受信任白名单内，已安全拦截"
    elif any(
        tok in low
        for tok in (
            "jsondecodeerror",
            "expecting value",
            "unterminated string",
            "extra data",
        )
    ):
        reason = "模型返回的数据格式异常，无法解析为 JSON"
    elif status_code:
        reason = f"接口服务返回异常状态 [HTTP {status_code}]"
    else:
        reason = "网络连接失败或超时，请检查网络设置"

    return f"{prefix}{reason}"


def _resolve_backend(api_key: str, api_mode: str = "auto") -> str:
    k = (api_key or "").strip()
    if api_mode in ("vertex", "gemini"):
        return api_mode
    if k.startswith("AIza"):
        return "gemini"
    if k.startswith("AQ."):
        return "vertex"
    return "vertex"


def _is_vertex_auth_scheme_mismatch(backend: str, status_code: Optional[int], msg: str, key: str) -> bool:
    if backend != "vertex":
        return False
    low = (msg or "").lower()
    if (
        "expected oauth2" in low
        or "api keys are not supported" in low
        or "authentication credentials" in low
        or "unauthenticated" in low
    ):
        return True
    if status_code in (401, 403) and (key or "").strip().startswith("AIza"):
        return True
    return False


def _adapt_schema_for_backend(schema: Optional[dict], backend: str) -> Optional[dict]:
    """Normalize JSON schema type casing dynamically so mixed Vertex/Gemini key pools seamlessly interoperate."""
    if not schema:
        return None
    target_upper = backend == "vertex"

    def _walk(node):
        if isinstance(node, dict):
            out = {}
            for k, v in node.items():
                if k == "type" and isinstance(v, str):
                    out[k] = v.upper() if target_upper else v.lower()
                else:
                    out[k] = _walk(v)
            return out
        if isinstance(node, list):
            return [_walk(x) for x in node]
        return node

    return _walk(copy.deepcopy(schema))


class ApiKeyPool:
    """
    Thread-safe pool supporting an unlimited number of API keys with:
      - Round-robin load distribution ('round_robin') or sticky primary-backup ('failover')
      - Instant automatic failover on quota exhaustion (429), invalid key (400/401/403), 5xx, or timeout
      - Per-key cooldown & permanent invalidation within a run so dead keys are not retried repeatedly
    """

    def __init__(self, raw_keys: Union[str, List[str], "ApiKeyPool"], rotation_mode: str = "round_robin"):
        if isinstance(raw_keys, ApiKeyPool):
            self.keys = list(raw_keys.keys)
            self.rotation_mode = raw_keys.rotation_mode
        else:
            self.keys = parse_api_keys(raw_keys)
            self.rotation_mode = rotation_mode if rotation_mode in ("round_robin", "failover") else "round_robin"

        if not self.keys:
            raise ValueError("请至少填写 1 个有效的 API Key（支持填写无限个 Key 自动轮换）。")

        self._lock = threading.Lock()
        self._cursor = 0
        self._dead_reasons: Dict[int, str] = {}
        self._cooldown_until: Dict[int, float] = {}
        self._last_error_reason: str = ""
        self._last_status: Optional[int] = None

    def __len__(self) -> int:
        return len(self.keys)

    @property
    def all_keys(self) -> List[str]:
        return list(self.keys)

    @property
    def last_error_info(self) -> Tuple[str, Optional[int]]:
        with self._lock:
            return self._last_error_reason, self._last_status

    def get_candidate_order(self) -> List[int]:
        """Return key indices ordered for the next request: healthy keys first, cooling-down keys as fallback."""
        with self._lock:
            n = len(self.keys)
            now = time.monotonic()
            start = self._cursor % n
            if self.rotation_mode == "round_robin":
                self._cursor = (self._cursor + 1) % n

            healthy = []
            cooling = []
            for offset in range(n):
                idx = (start + offset) % n
                if idx in self._dead_reasons:
                    continue
                cd = self._cooldown_until.get(idx, 0.0)
                if cd <= now:
                    healthy.append(idx)
                else:
                    cooling.append((cd, idx))

            cooling.sort(key=lambda x: x[0])
            return healthy + [idx for _, idx in cooling]

    def mark_success(self, idx: int) -> None:
        with self._lock:
            self._cooldown_until.pop(idx, None)
            if self.rotation_mode == "failover":
                self._cursor = idx

    def mark_failure(self, idx: int, status_code: Optional[int], reason: str) -> str:
        """Record failure for key `idx` and return human-readable action tag."""
        with self._lock:
            self._last_error_reason = reason
            self._last_status = status_code
            if status_code in (400, 401, 403):
                self._dead_reasons[idx] = reason
                if self.rotation_mode == "failover" and len(self.keys) > 1:
                    self._cursor = (idx + 1) % len(self.keys)
                return "permanent"
            if status_code == 429:
                self._cooldown_until[idx] = time.monotonic() + 45.0
            else:
                self._cooldown_until[idx] = time.monotonic() + 15.0
            if self.rotation_mode == "failover" and len(self.keys) > 1:
                self._cursor = (idx + 1) % len(self.keys)
            return "cooldown"


class _SafeRedirectHandler(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        validate_https_endpoint(newurl)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def _http_post_json(url: str, headers: dict, payload: dict, timeout: int = 90) -> dict:
    validate_https_endpoint(url)
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers, method="POST")
    ctx = ssl.create_default_context()
    ctx.minimum_version = ssl.TLSVersion.TLSv1_2
    opener = urllib.request.build_opener(
        urllib.request.HTTPSHandler(context=ctx),
        _SafeRedirectHandler(),
    )
    with opener.open(req, timeout=timeout) as resp:
        raw_bytes = resp.read(MAX_API_RESP_BYTES + 1)
        if len(raw_bytes) > MAX_API_RESP_BYTES:
            raise RuntimeError("API 响应体积超过安全上限 (10 MB)，已中止解析。")
        raw = raw_bytes.decode("utf-8")
        return json.loads(raw)


def _http_get_json(url: str, headers: dict, timeout: int = 30) -> dict:
    validate_https_endpoint(url)
    req = urllib.request.Request(url, headers=headers, method="GET")
    ctx = ssl.create_default_context()
    ctx.minimum_version = ssl.TLSVersion.TLSv1_2
    opener = urllib.request.build_opener(
        urllib.request.HTTPSHandler(context=ctx),
        _SafeRedirectHandler(),
    )
    with opener.open(req, timeout=timeout) as resp:
        raw_bytes = resp.read(MAX_API_RESP_BYTES + 1)
        if len(raw_bytes) > MAX_API_RESP_BYTES:
            raise RuntimeError("API 响应体积超过安全上限 (10 MB)，已中止解析。")
        return json.loads(raw_bytes.decode("utf-8"))


def _is_text_generation_gemini_model(model_id: str, supported_methods: Optional[List[str]] = None) -> bool:
    mid = (model_id or "").strip().lower()
    if not mid.startswith("gemini-"):
        return False
    if any(ex in mid for ex in _EXCLUDED_MODEL_SUBSTRINGS):
        return False
    if supported_methods is not None:
        methods_low = {str(m).lower() for m in supported_methods}
        if methods_low and "generatecontent" not in methods_low:
            return False
    try:
        validate_model_id(mid)
    except ValueError:
        return False
    return True


def fetch_available_models(
    raw_keys: Union[str, List[str]],
    api_mode: str = "auto",
) -> Tuple[List[Tuple[str, str]], str]:
    """
    Dynamically fetch the latest available text/JSON generation models using the configured
    Gemini API Key and/or Vertex AI Paid API Key, merged cleanly with `SUPPORTED_MODELS`.
    Returns `(model_pairs_list, human_summary_msg)`.
    """
    keys = parse_api_keys(raw_keys)
    if not keys:
        raise ValueError("请先在密钥池中粘贴至少 1 个 Gemini 或 Vertex 收费版 API Key，再加载模型列表。")

    discovered: Dict[str, str] = {}  # model_id -> display_name
    channels_ok: List[str] = []
    last_err: Optional[Exception] = None

    for k in keys:
        clean_k = k.strip()
        try_gemini = api_mode in ("auto", "gemini")
        try_vertex = api_mode in ("auto", "vertex")

        if try_gemini:
            try:
                res = _http_get_json(
                    "https://generativelanguage.googleapis.com/v1beta/models?pageSize=200",
                    headers={"x-goog-api-key": clean_k},
                    timeout=20,
                )
                for m_obj in res.get("models") or []:
                    raw_name = str(m_obj.get("name") or "").strip()
                    mid = raw_name.split("/")[-1].strip()
                    methods = m_obj.get("supportedGenerationMethods")
                    if _is_text_generation_gemini_model(mid, methods if isinstance(methods, list) else None):
                        disp_title = str(m_obj.get("displayName") or mid).strip()
                        discovered.setdefault(mid, disp_title)
                if "Gemini API" not in channels_ok:
                    channels_ok.append("Gemini API")
            except Exception as ex:
                last_err = ex

        if try_vertex:
            vertex_listed = False
            for ver_seg in ("v1beta1", "v1"):
                try:
                    res_v = _http_get_json(
                        f"https://aiplatform.googleapis.com/{ver_seg}/publishers/google/models?pageSize=200",
                        headers={"x-goog-api-key": clean_k},
                        timeout=20,
                    )
                    items_v = res_v.get("models") or res_v.get("publisherModels") or []
                    for m_obj in items_v:
                        raw_name = str(m_obj.get("name") or "").strip()
                        mid = raw_name.split("/")[-1].strip()
                        if _is_text_generation_gemini_model(mid):
                            disp_title = str(m_obj.get("displayName") or mid).strip()
                            discovered.setdefault(mid, disp_title)
                    vertex_listed = True
                    if "Vertex AI 收费通道" not in channels_ok:
                        channels_ok.append("Vertex AI 收费通道")
                    break
                except Exception as ex_v:
                    last_err = ex_v

            # Vertex Express / Paid API keys sometimes restrict publisher catalog GET while allowing `:generateContent` POST!
            if not vertex_listed:
                try:
                    _call_single_key_once(
                        key=clean_k,
                        safe_model="gemini-2.5-flash",
                        quoted_model="gemini-2.5-flash",
                        system_instruction='Return JSON {"ok": true}.',
                        user_prompt="ping",
                        json_schema=None,
                        api_mode="vertex",
                    )
                    if "Vertex AI 收费通道" not in channels_ok:
                        channels_ok.append("Vertex AI 收费通道")
                    for disp_b, mid_b in SUPPORTED_MODELS:
                        discovered.setdefault(mid_b, disp_b)
                except Exception as ex_probe:
                    # Also probe gemini-3.5-flash-lite on Vertex in case 2.5-flash is disabled
                    try:
                        _call_single_key_once(
                            key=clean_k,
                            safe_model="gemini-3.5-flash-lite",
                            quoted_model="gemini-3.5-flash-lite",
                            system_instruction='Return JSON {"ok": true}.',
                            user_prompt="ping",
                            json_schema=None,
                            api_mode="vertex",
                        )
                        if "Vertex AI 收费通道" not in channels_ok:
                            channels_ok.append("Vertex AI 收费通道")
                        for disp_b, mid_b in SUPPORTED_MODELS:
                            discovered.setdefault(mid_b, disp_b)
                    except Exception:
                        last_err = last_err or ex_probe

        if discovered and len(channels_ok) >= (2 if api_mode == "auto" else 1):
            break

    if not discovered and not channels_ok:
        raise RuntimeError(format_friendly_api_error(last_err or "未能加载模型列表，请检查 API Key 与通道设置。", total_keys=len(keys)))

    # Order models: curated built-in top recommendations first (if supported/merged), then newly discovered live models sorted newest-first!
    merged_out: List[Tuple[str, str]] = []
    seen_ids = set()

    for disp_b, mid_b in SUPPORTED_MODELS:
        if mid_b in discovered or "Vertex AI 收费通道" in channels_ok:
            merged_out.append((disp_b, mid_b))
            seen_ids.add(mid_b)

    extra_ids = sorted(
        (m for m in discovered.keys() if m not in seen_ids),
        key=lambda x: x,
        reverse=True,
    )
    for mid in extra_ids:
        if mid in _BUILTIN_MODEL_LABELS:
            merged_out.append((_BUILTIN_MODEL_LABELS[mid], mid))
        else:
            disp_t = discovered[mid]
            lbl = f"{disp_t} [{mid}]" if mid.lower() not in disp_t.lower() else f"{mid}"
            merged_out.append((lbl, mid))
        seen_ids.add(mid)

    # Always keep built-in fallbacks accessible at the bottom if not already present
    for disp_b, mid_b in SUPPORTED_MODELS:
        if mid_b not in seen_ids:
            merged_out.append((disp_b, mid_b))
            seen_ids.add(mid_b)

    ch_str = " / ".join(channels_ok) if channels_ok else "默认通道"
    return merged_out, f"已从 {ch_str} 成功加载 {len(merged_out)} 个可选模型！"


def _call_single_key_once(
    key: str,
    safe_model: str,
    quoted_model: str,
    system_instruction: str,
    user_prompt: str,
    json_schema: Optional[dict],
    api_mode: str,
) -> tuple[dict, str]:
    """Attempt one logical call with a single key (seamlessly supporting both Paid Vertex AI API keys and Gemini API keys)."""
    clean_k = (key or "").strip()
    primary = _resolve_backend(clean_k, api_mode)
    backends = [primary]
    if api_mode == "auto":
        backends.append("gemini" if primary == "vertex" else "vertex")

    primary_code: Optional[int] = None
    primary_err: Optional[Exception] = None
    primary_mismatch: bool = False

    last_code: Optional[int] = None
    last_err: Optional[Exception] = None

    for b_idx, backend in enumerate(backends):
        adapted_schema = _adapt_schema_for_backend(json_schema, backend)
        try:
            if backend == "vertex":
                headers = {
                    "Content-Type": "application/json; charset=utf-8",
                    "x-goog-api-key": clean_k,
                }
                gen_cfg = {
                    "temperature": 0.1,
                    "responseMimeType": "application/json",
                }
                if adapted_schema:
                    gen_cfg["responseSchema"] = adapted_schema
                payload = {
                    "systemInstruction": {"parts": [{"text": system_instruction}]},
                    "contents": [{"role": "user", "parts": [{"text": user_prompt}]}],
                    "generationConfig": gen_cfg,
                }
                # Try GA paid endpoint (`v1`) first, falling back to `v1beta1` if model is preview-only (`404`)
                res = None
                for ver_seg in ("v1", "v1beta1"):
                    url = f"https://aiplatform.googleapis.com/{ver_seg}/publishers/google/models/{quoted_model}:generateContent"
                    try:
                        res = _http_post_json(url, headers, payload)
                        break
                    except urllib.error.HTTPError as v_he:
                        if v_he.code == 404 and ver_seg == "v1":
                            try:
                                v_he.close()
                            except Exception:
                                pass
                            continue
                        raise
                cands = (res or {}).get("candidates") or []
                if not cands:
                    block_reason = (res or {}).get("promptFeedback", {}).get("blockReason", "EMPTY_CANDIDATES")
                    raise RuntimeError(f"Vertex AI 未返回有效候选结果 ({block_reason})。")
                parts = cands[0].get("content", {}).get("parts", [])
                text = "".join(p.get("text", "") for p in parts if "text" in p).strip()
            else:
                headers_std = {
                    "Content-Type": "application/json; charset=utf-8",
                    "x-goog-api-key": clean_k,
                }
                gen_cfg_g = {
                    "temperature": 0.1,
                    "responseMimeType": "application/json",
                }
                if adapted_schema:
                    gen_cfg_g["responseSchema"] = adapted_schema
                payload_std = {
                    "systemInstruction": {"parts": [{"text": system_instruction}]},
                    "contents": [{"role": "user", "parts": [{"text": user_prompt}]}],
                    "generationConfig": gen_cfg_g,
                }
                try:
                    url_std = f"https://generativelanguage.googleapis.com/v1beta/models/{quoted_model}:generateContent"
                    res = _http_post_json(url_std, headers_std, payload_std)
                    cands_g = res.get("candidates") or []
                    if not cands_g:
                        block_reason = res.get("promptFeedback", {}).get("blockReason", "EMPTY_CANDIDATES")
                        raise RuntimeError(f"Gemini API 未返回有效候选结果 ({block_reason})。")
                    parts_g = cands_g[0].get("content", {}).get("parts", [])
                    text = "".join(p.get("text", "") for p in parts_g if "text" in p).strip()
                except urllib.error.HTTPError as g_he:
                    if g_he.code != 404:
                        raise
                    try:
                        g_he.close()
                    except Exception:
                        pass
                    # Fallback to Interactions API for interaction-only preview endpoints
                    url_int = "https://generativelanguage.googleapis.com/v1beta/interactions"
                    headers_int = {
                        "Content-Type": "application/json; charset=utf-8",
                        "x-goog-api-key": clean_k,
                        "Api-Revision": "2026-05-20",
                    }
                    payload_int = {
                        "model": safe_model,
                        "store": False,
                        "system_instruction": system_instruction,
                        "input": user_prompt,
                        "generation_config": {"temperature": 0.1},
                    }
                    if adapted_schema:
                        payload_int["response_format"] = {
                            "type": "text",
                            "mime_type": "application/json",
                            "schema": adapted_schema,
                        }
                    res = _http_post_json(url_int, headers_int, payload_int)
                    steps = res.get("steps") or []
                    text_chunks = []
                    for st in steps:
                        if st.get("type") == "model_output":
                            for c in st.get("content") or []:
                                if c.get("type") == "text":
                                    text_chunks.append(c.get("text", ""))
                    text = "".join(text_chunks).strip() or res.get("output_text", "").strip()

            return _extract_json_obj(text), backend

        except urllib.error.HTTPError as he:
            body_txt = ""
            try:
                body_txt = he.read(2048).decode("utf-8", errors="ignore")[:320]
            except Exception:
                pass
            finally:
                try:
                    he.close()
                except Exception:
                    pass
            safe_msg = redact_secrets(body_txt or str(he.reason), extra_secret=clean_k)
            cur_err = RuntimeError(f"HTTP {he.code} ({backend}/{safe_model}): {safe_msg}")
            is_mismatch = _is_vertex_auth_scheme_mismatch(backend, he.code, safe_msg, clean_k)

            if b_idx == 0:
                primary_code = he.code
                primary_err = cur_err
                primary_mismatch = is_mismatch
                last_code = he.code
                last_err = cur_err

                # IMPORTANT: Paid Google Cloud Vertex AI API keys often start with `AIza...` and are restricted
                # to `aiplatform.googleapis.com` (returning 400/403 on `generativelanguage.googleapis.com`),
                # so ALWAYS probe the alternate backend in `auto` mode on 400/401/403/404!
                if he.code in (400, 401, 403, 404) and api_mode == "auto" and len(backends) > 1:
                    continue
            else:
                if is_mismatch or (not primary_mismatch and he.code in (400, 401, 403, 404)):
                    last_code = primary_code
                    last_err = primary_err
                else:
                    last_code = he.code
                    last_err = cur_err
            break
        except Exception as ex:
            cur_err = RuntimeError(redact_secrets(str(ex), extra_secret=clean_k))
            if b_idx == 0 or primary_err is None:
                last_code = None
                last_err = cur_err
            else:
                last_code = primary_code
                last_err = primary_err
            break

    err_out = last_err or RuntimeError("调用 API 失败。")
    setattr(err_out, "http_status", last_code)
    raise err_out


def call_gemini_json(
    api_key: Union[str, List[str], ApiKeyPool],
    model: str,
    system_instruction: str,
    user_prompt: str,
    json_schema: Optional[dict] = None,
    api_mode: str = "auto",
    max_retries: int = 4,
    rotation_mode: str = "round_robin",
    log_fn: Optional[Callable[[str], None]] = None,
) -> dict:
    """
    Call Gemini / Vertex AI with multi-key rotation, exponential 503/429 backoff, & model spike fallback.
    Raises a concise human-readable Chinese RuntimeError (preserving `.http_status`) when all keys fail.
    """
    pool = api_key if isinstance(api_key, ApiKeyPool) else ApiKeyPool(api_key, rotation_mode=rotation_mode)
    safe_model = validate_model_id(model)
    quoted_model = urllib.parse.quote(safe_model, safe="")

    total_keys = len(pool)
    last_err: Optional[Exception] = None
    last_status: Optional[int] = None

    for sweep in range(max_retries):
        cand_indices = pool.get_candidate_order()
        if not cand_indices:
            break

        for pos, k_idx in enumerate(cand_indices):
            k_val = pool.keys[k_idx]
            k_lbl = mask_key_label(k_idx, k_val)
            try:
                parsed_json, _used_backend = _call_single_key_once(
                    key=k_val,
                    safe_model=safe_model,
                    quoted_model=quoted_model,
                    system_instruction=system_instruction,
                    user_prompt=user_prompt,
                    json_schema=json_schema,
                    api_mode=api_mode,
                )
                pool.mark_success(k_idx)
                return parsed_json
            except Exception as ex:
                status_code = getattr(ex, "http_status", None)
                last_status = status_code
                clean_reason = redact_secrets(str(ex), extra_secret=pool.all_keys)
                f_kind = pool.mark_failure(k_idx, status_code, clean_reason)
                last_err = RuntimeError(f"{k_lbl} -> {clean_reason}")
                setattr(last_err, "http_status", status_code)

                if total_keys > 1:
                    next_hint = ""
                    if pos + 1 < len(cand_indices):
                        nxt_idx = cand_indices[pos + 1]
                        next_hint = f"，自动切换至 {mask_key_label(nxt_idx, pool.keys[nxt_idx])} 重试"
                    elif sweep + 1 < max_retries:
                        next_hint = "，准备下一轮重试可用密钥"

                    tag_txt = "密钥无效/无权限(本轮已跳过)" if f_kind == "permanent" else "额度不足/超速/高峰拥堵(冷却切换)"
                    if log_fn:
                        log_fn(f"  ⚠️ {k_lbl} {tag_txt} [{status_code or 'ERR'}]{next_hint}...")

        # All available keys failed in this sweep -> exponential backoff (3s -> 6s -> 12s) before next sweep
        remaining_after = pool.get_candidate_order()
        if remaining_after and sweep + 1 < max_retries:
            wait_s = min(15.0, 3.0 * (2 ** sweep))
            if log_fn:
                if total_keys > 1:
                    log_fn(f"  ⏳ 当前轮次可用 Key 均暂时受限 [{last_status or 'ERR'}]，等待 {wait_s:.0f}s 后进行第 {sweep + 2}/{max_retries} 轮重试...")
                else:
                    log_fn(f"  ⏳ 模型接口高峰拥堵/暂时受限 [HTTP {last_status or 'ERR'}]，等待 {wait_s:.0f}s 后自动第 {sweep + 2}/{max_retries} 次重试...")
            time.sleep(wait_s)

    # Automatic fallback model on persistent 502/503/504 high-demand spikes
    if last_status in (502, 503, 504) and safe_model == "gemini-3.5-flash-lite":
        alt_model = "gemini-3.1-flash-lite"
        alt_quoted = urllib.parse.quote(alt_model, safe="")
        cand_indices = pool.get_candidate_order()
        if cand_indices:
            if log_fn:
                log_fn(f"  ↻ 模型 {safe_model} 持续高峰拥堵 (HTTP {last_status})，自动切换备用模型 {alt_model} 继续执行...")
            for k_idx in cand_indices:
                k_val = pool.keys[k_idx]
                try:
                    parsed_json, _ = _call_single_key_once(
                        key=k_val,
                        safe_model=alt_model,
                        quoted_model=alt_quoted,
                        system_instruction=system_instruction,
                        user_prompt=user_prompt,
                        json_schema=json_schema,
                        api_mode=api_mode,
                    )
                    pool.mark_success(k_idx)
                    return parsed_json
                except Exception as alt_ex:
                    alt_code = getattr(alt_ex, "http_status", last_status)
                    last_status = alt_code
                    clean_alt = redact_secrets(str(alt_ex), extra_secret=pool.all_keys)
                    last_err = RuntimeError(f"{mask_key_label(k_idx, k_val)} -> {clean_alt}")
                    setattr(last_err, "http_status", alt_code)

    if last_err is None:
        prev_reason, prev_status = pool.last_error_info
        if prev_reason:
            last_status = prev_status
            last_err = RuntimeError(prev_reason)
            setattr(last_err, "http_status", prev_status)

    friendly_msg = format_friendly_api_error(last_err or "网络连接失败或超时", total_keys=total_keys)
    final_err = RuntimeError(friendly_msg)
    setattr(final_err, "http_status", last_status)
    raise final_err


def _extract_json_obj(text: str):
    t = text.strip()
    if t.startswith("```"):
        t = re.sub(r"^```(?:json)?\s*", "", t)
        t = re.sub(r"\s*```$", "", t)
    return json.loads(t)


def test_all_api_keys(
    raw_keys: Union[str, List[str]],
    model: str,
    api_mode: str = "auto",
    on_each_cb: Optional[Callable[[int, int, dict], None]] = None,
) -> List[dict]:
    """Test every configured key individually and return detailed health report per key."""
    keys = parse_api_keys(raw_keys)
    if not keys:
        raise ValueError("请先填写至少 1 个 API Key。")
    safe_model = validate_model_id(model)
    quoted_model = urllib.parse.quote(safe_model, safe="")
    schema = {
        "type": "object",
        "properties": {"status": {"type": "string"}},
        "required": ["status"],
    }

    results: List[dict] = []
    total = len(keys)
    for idx, k in enumerate(keys):
        masked = mask_key_label(idx, k)
        try:
            res, used_backend = _call_single_key_once(
                key=k,
                safe_model=safe_model,
                quoted_model=quoted_model,
                system_instruction='Return JSON {"status": "ok"}.',
                user_prompt="Ping",
                json_schema=schema,
                api_mode=api_mode,
            )
            ch_label = "Vertex AI 收费通道" if used_backend == "vertex" else "Gemini API 通道"
            item = {
                "index": idx,
                "key": k,
                "masked": masked,
                "ok": True,
                "channel": ch_label,
                "detail": f"可用 ({ch_label} | 响应: {res.get('status', 'ok')})",
            }
        except Exception as ex:
            friendly = format_friendly_api_error(ex, total_keys=1)
            short_detail = friendly[len("API Key 调用失败：") :] if friendly.startswith("API Key 调用失败：") else friendly
            item = {
                "index": idx,
                "key": k,
                "masked": masked,
                "ok": False,
                "channel": "-",
                "detail": short_detail,
                "http_status": getattr(ex, "http_status", None),
            }
        results.append(item)
        if on_each_cb:
            on_each_cb(idx + 1, total, item)
    return results


def test_api_connection(
    api_key: Union[str, List[str]],
    model: str,
    api_mode: str = "auto",
    log_fn: Optional[Callable[[str], None]] = None,
) -> str:
    """Test single key or entire multi-key pool and return human-readable summary."""
    keys = parse_api_keys(api_key)
    if not keys:
        raise ValueError("请先输入至少 1 个 Gemini API Key 或 Vertex API Key。")

    if len(keys) == 1:
        rep = test_all_api_keys(keys, model=model, api_mode=api_mode)[0]
        if not rep["ok"]:
            err = RuntimeError(f"API Key 调用失败：{rep['detail']}")
            setattr(err, "http_status", rep.get("http_status"))
            raise err
        return f"连接成功！{rep['masked']} | 通道：{rep['channel']} | 模型：{model}"

    def _each(done_i: int, tot: int, item: dict):
        if log_fn:
            icon = "✅" if item["ok"] else "❌"
            log_fn(f"  [{done_i}/{tot}] {icon} {item['masked']}：{item['detail']}")

    reports = test_all_api_keys(keys, model=model, api_mode=api_mode, on_each_cb=_each)
    ok_list = [r for r in reports if r["ok"]]
    bad_list = [r for r in reports if not r["ok"]]
    if not ok_list:
        first_bad = bad_list[0]
        err = RuntimeError(f"密钥池中全部 {len(reports)} 个 API Key 均调用失败：{first_bad['detail']}")
        setattr(err, "http_status", first_bad.get("http_status"))
        raise err
    return (
        f"密钥池测试通过！共 {len(reports)} 个 Key：✅ 可用 {len(ok_list)} 个，"
        f"⚠️ 不可用 {len(bad_list)} 个（执行时将自动跳过报错 Key 并轮换可用 Key） | 模型：{model}"
    )


# ---------------- Offline / Fallback Multi-Script Col 3 Approximate Alignment ----------------

_NON_SPACE_SCRIPT_RE = re.compile(
    r"[\u0E00-\u0E7F"  # Thai
    r"\u0E80-\u0EFF"   # Lao
    r"\u0F00-\u0FFF"   # Tibetan
    r"\u1000-\u109F\uAA60-\uAA7F\uA9E0-\uA9FF"  # Myanmar
    r"\u1780-\u17FF\u19E0-\u19FF"  # Khmer
    r"\u3040-\u30FF"   # Hiragana / Katakana
    r"\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF]"  # CJK Unified Ideographs
)
_THAI_LAO_RE = re.compile(r"[\u0E00-\u0EFF]")
_GREEK_RE = re.compile(r"[\u0370-\u03FF\u1F00-\u1FFF]")
_EN_WORD_RE = re.compile(r"[A-Za-z0-9]+(?:[’'][A-Za-z0-9]+)*(?:-[A-Za-z0-9]+)*")
_DIGIT_TOKEN_RE = re.compile(r"\d+(?:[:.]\d+)*")

_CLOSING_TRAILERS = frozenset('”’"\'»›」』)）]】〕}》〉')
_OPENING_LEADERS = frozenset('“‘"\'«‹「『(（[【〔{„‚《〈')
_PREPOSED_VOWELS = frozenset(
    "\u0e40\u0e41\u0e42\u0e43\u0e44"  # Thai เ แ โ ใ ไ
    "\u0ec0\u0ec1\u0ec2\u0ec3\u0ec4"  # Lao ເ ແ ໂ ໃ ໄ
)
_STACKER_VIRAMAS = frozenset(
    "\u1039"  # Myanmar virama / stacker
    "\u17d2"  # Khmer coeng
    "\u094d"  # Devanagari virama
    "\u09cd"  # Bengali virama
    "\u0e3a"  # Thai phinthu
)


def _is_non_space_script_text(text: str) -> bool:
    non_ws = [ch for ch in text if not ch.isspace()]
    if not non_ws:
        return False
    ns_cnt = len(_NON_SPACE_SCRIPT_RE.findall(text))
    return (ns_cnt / len(non_ws)) >= 0.20


def _trim_range(text: str, s: int, e: int) -> Tuple[int, int]:
    n = len(text)
    s = max(0, min(n, s))
    e = max(s, min(n, e))
    while s < e and (text[s].isspace() or text[s] in ("\u200b", "\u200c", "\ufeff")):
        s += 1
    while e > s and (text[e - 1].isspace() or text[e - 1] in ("\u200b", "\u200c", "\ufeff")):
        e -= 1
    return s, e


def _split_sentences_ranges(
    text: str,
    is_english: bool = False,
    include_semi_colon: bool = False,
) -> List[Tuple[int, int]]:
    """
    Split `text` into trimmed non-empty sentence character ranges `[(start, end), ...]`.
    Handles English (.?! and optional ;:), Greek (; question mark, · ano teleia),
    Arabic (؟ ۔ ؛), Devanagari (। ॥), Myanmar (။), Khmer (។ ៕), CJK (。？！), and newlines.
    """
    n = len(text)
    if n == 0:
        return []

    has_greek = (not is_english) and bool(_GREEK_RE.search(text))
    primary_terms = {".", "?", "!", "。", "？", "！", "؟", "۔", "।", "॥", "။", "។", "៕", "።", "།", "՞", "։"}
    if has_greek:
        primary_terms.add(";")
        primary_terms.add("\u037e")  # Greek question mark

    semi_terms = {";", ":", "；", "：", "؛", "·", "·"}
    active_terms = primary_terms | (semi_terms if include_semi_colon else set())

    ranges: List[Tuple[int, int]] = []
    seg_start = 0
    i = 0
    while i < n:
        ch = text[i]
        if ch == "\n":
            ts, te = _trim_range(text, seg_start, i)
            if ts < te:
                ranges.append((ts, te))
            seg_start = i + 1
            i += 1
            continue

        if ch in active_terms:
            # Do not split inside decimal numbers or chapter:verse citations like 3.14 or 1:21
            if ch in (".", ":") and i > 0 and i + 1 < n and text[i - 1].isdigit() and text[i + 1].isdigit():
                i += 1
                continue
            # Do not split ellipsis '...' mid-run; consume consecutive dots/terminators together
            j = i + 1
            while j < n and text[j] in active_terms and text[j] != "\n":
                j += 1
            while j < n and text[j] in _CLOSING_TRAILERS:
                j += 1
            ts, te = _trim_range(text, seg_start, j)
            if ts < te:
                ranges.append((ts, te))
            seg_start = j
            i = j
            continue

        i += 1

    ts, te = _trim_range(text, seg_start, n)
    if ts < te:
        ranges.append((ts, te))
    return ranges


def _split_clauses_ranges(text: str, s: int, e: int, is_english: bool = False) -> List[Tuple[int, int]]:
    """Split a single sentence slice `text[s:e]` into sub-clause ranges at commas, semicolons, colons, or dashes."""
    if s >= e:
        return []
    sub = text[s:e]
    if is_english:
        delims = {",", ";", ":", "—", "–"}
    else:
        delims = {",", "，", "、", "،", "؛", ";", "；", ":", "：", "·", "·", "—", "–", "၊", "፣", "፤"}

    clauses: List[Tuple[int, int]] = []
    c_start = s
    i = s
    while i < e:
        ch = text[i]
        if ch in delims:
            if ch in (".", ":", ",") and i > s and i + 1 < e and text[i - 1].isdigit() and text[i + 1].isdigit():
                i += 1
                continue
            j = i + 1
            if ch in ("—", "–"):
                while j < e and text[j] in ("—", "–"):
                    j += 1
            while j < e and text[j] in _CLOSING_TRAILERS:
                j += 1
            cs, ce = _trim_range(text, c_start, j)
            if cs < ce:
                clauses.append((cs, ce))
            c_start = j
            i = j
            continue
        i += 1

    cs, ce = _trim_range(text, c_start, e)
    if cs < ce:
        clauses.append((cs, ce))

    # For Thai/Lao sentences that use spaces as clause breaks without punctuation commas
    if len(clauses) <= 1 and not is_english and _THAI_LAO_RE.search(sub):
        ws_parts = [
            (s + m.start(), s + m.end())
            for m in re.finditer(r"\S+", sub)
        ]
        if len(ws_parts) >= 2:
            return ws_parts

    return clauses if clauses else [(s, e)]


def _build_grapheme_clusters(text: str, s: int, e: int) -> List[Tuple[int, int]]:
    """
    Segment `text[s:e]` into indivisible orthographic/grapheme clusters `(cs, ce)` (skipping whitespace),
    never splitting Unicode combining marks (`M*`), Brahmic virama/coeng stacks, or Thai/Lao pre-posed vowels.
    """
    clusters: List[Tuple[int, int]] = []
    i = s
    while i < e:
        ch = text[i]
        if ch.isspace() or ch in ("\u200b", "\u200c", "\ufeff"):
            i += 1
            continue

        cs = i
        # Consume any leading Thai/Lao pre-posed vowel(s) onto the following consonant
        while i < e and text[i] in _PREPOSED_VOWELS:
            i += 1
        if i < e and not text[i].isspace():
            i += 1

        while i < e:
            c_next = text[i]
            cat = unicodedata.category(c_next)
            if cat.startswith("M"):
                if c_next in _STACKER_VIRAMAS and i + 1 < e and not text[i + 1].isspace():
                    i += 2
                else:
                    i += 1
            elif c_next == "\u200d" and i + 1 < e:
                i += 2
            else:
                break

        clusters.append((cs, i))
    return clusters


def _snap_safe_codepoint_bounds(text: str, s: int, e: int, min_s: int, max_e: int) -> Tuple[int, int]:
    """Ensure `[s, e]` never starts on a combining mark or cuts off trailing combining marks."""
    s = max(min_s, min(max_e, s))
    e = max(s, min(max_e, e))
    while s < max_e and unicodedata.category(text[s]).startswith("M"):
        s += 1
    if s > min_s and s < max_e and text[s - 1] in _PREPOSED_VOWELS:
        s -= 1
    if e < s:
        e = s
    while e < max_e:
        cat = unicodedata.category(text[e])
        if cat.startswith("M"):
            if text[e] in _STACKER_VIRAMAS and e + 1 < max_e and not text[e + 1].isspace():
                e += 2
            else:
                e += 1
        elif e > s and text[e - 1] in _PREPOSED_VOWELS and not text[e].isspace():
            e += 1
        else:
            break
    return _trim_range(text, s, e)


def _reconstruct_fallback_old_en(spans: List[dict]) -> str:
    """Reconstruct synthetic `old_en_full` from span contexts if legacy task dict omitted `old_en_full`."""
    chunks: List[str] = []
    for idx, sp in enumerate(spans):
        l_ctx = sp.get("left_old_context") or ""
        mid = sp.get("old_en_in_col3") or ""
        r_ctx = sp.get("right_old_context") or ""
        if idx == 0 and l_ctx:
            chunks.append(l_ctx)
        if mid:
            if chunks and not chunks[-1].endswith((" ", "\n")) and not mid.startswith((" ", ",", ".", ";", ":", "?", "!")):
                chunks.append(" ")
            chunks.append(mid)
        if r_ctx:
            chunks.append(r_ctx)
    return "".join(chunks).strip()


def _call_find_fuzzy(
    find_fuzzy_fn: Callable,
    haystack: str,
    needle: str,
    search_from: int = 0,
) -> Optional[Tuple[int, int]]:
    try:
        return find_fuzzy_fn(haystack, needle, search_from=search_from)
    except TypeError:
        try:
            return find_fuzzy_fn(haystack, needle, search_from)
        except TypeError:
            return find_fuzzy_fn(haystack, needle)


def _locate_span_in_old_en(
    old_en_full: str,
    sp: dict,
    en_cursor: int,
    span_idx: int,
    total_spans: int,
    find_fuzzy_fn: Callable,
) -> Tuple[int, int]:
    n = len(old_en_full)
    if n == 0:
        return 0, 0

    cs = sp.get("old_en_char_start")
    ce = sp.get("old_en_char_end")
    if isinstance(cs, int) and isinstance(ce, int) and 0 <= cs <= ce <= n:
        if cs < ce:
            ts, te = _trim_range(old_en_full, cs, ce)
            return (ts, te) if ts < te else (cs, ce)
        return cs, ce

    piece = (sp.get("old_en_in_col3") or "").strip()
    l_ctx = (sp.get("left_old_context") or "").strip()
    r_ctx = (sp.get("right_old_context") or "").strip()

    if piece:
        # Try anchored search with left/right context first
        if l_ctx:
            anchor_l = l_ctx[-20:]
            l_pos = old_en_full.find(anchor_l, en_cursor)
            if l_pos != -1:
                p_pos = old_en_full.find(piece, l_pos + len(anchor_l))
                if p_pos != -1 and p_pos - (l_pos + len(anchor_l)) <= 12:
                    return p_pos, p_pos + len(piece)
        pos = old_en_full.find(piece, en_cursor)
        if pos != -1:
            return pos, pos + len(piece)
        loc = _call_find_fuzzy(find_fuzzy_fn, old_en_full, piece, search_from=en_cursor)
        if loc is not None and loc[0] >= en_cursor:
            return loc[0], loc[1]
        if loc is not None:
            return loc[0], loc[1]
    else:
        # Pure insertion: anchor via left_old_context end or right_old_context start
        if l_ctx:
            anchor_l = l_ctx[-24:]
            l_pos = old_en_full.find(anchor_l, en_cursor)
            if l_pos != -1:
                pt = l_pos + len(anchor_l)
                return pt, pt
        if r_ctx:
            anchor_r = r_ctx[:24]
            r_pos = old_en_full.find(anchor_r, en_cursor)
            if r_pos != -1:
                return r_pos, r_pos

    approx_s = max(en_cursor, int((span_idx / max(1, total_spans)) * n))
    approx_e = max(approx_s, int(((span_idx + 1) / max(1, total_spans)) * n))
    return min(n, approx_s), min(n, approx_e)


def _map_sentence_indices(
    s_first: int,
    s_last: int,
    en_sents: List[Tuple[int, int]],
    for_sents: List[Tuple[int, int]],
    en_total_len: int,
    for_total_len: int,
) -> Tuple[int, int]:
    n_en = len(en_sents)
    n_for = len(for_sents)
    if n_for <= 1:
        return 0, 0
    if n_en == n_for:
        return s_first, s_last
    if abs(n_en - n_for) <= 1 and n_en >= 2:
        f0 = 0 if s_first == 0 else (
            n_for - 1 if s_first >= n_en - 1 else round(s_first * (n_for - 1) / max(1, n_en - 1))
        )
        f1 = n_for - 1 if s_last >= n_en - 1 else (
            f0 if s_last == s_first else round(s_last * (n_for - 1) / max(1, n_en - 1))
        )
        f0 = max(0, min(n_for - 1, f0))
        f1 = max(f0, min(n_for - 1, f1))
        return f0, f1

    # Proportional character mapping snapped to overlapping foreign sentences
    en_s0 = en_sents[s_first][0]
    en_e1 = en_sents[s_last][1]
    r0 = en_s0 / max(1, en_total_len)
    r1 = en_e1 / max(1, en_total_len)
    target_s = r0 * for_total_len
    target_e = max(target_s + 1.0, r1 * for_total_len)

    overlapping = [
        idx for idx, (fs, fe) in enumerate(for_sents)
        if max(fs, target_s) < min(fe, target_e)
    ]
    if overlapping:
        return overlapping[0], overlapping[-1]

    mid = (target_s + target_e) / 2.0
    best_i = min(range(n_for), key=lambda idx: abs(((for_sents[idx][0] + for_sents[idx][1]) / 2.0) - mid))
    return best_i, best_i


def _word_ratio_in_box(en_text: str, box_s: int, box_e: int, c_start: int, c_end: int) -> Tuple[float, float, int, int]:
    """Return `(rel_start, rel_end, span_word_count, box_word_count)` inside `en_text[box_s:box_e]`."""
    box_slice = en_text[box_s:box_e]
    box_words = list(_EN_WORD_RE.finditer(box_slice))
    rel_cs = max(box_s, min(box_e, c_start))
    rel_ce = max(rel_cs, min(box_e, c_end))
    span_words = len(_EN_WORD_RE.findall(en_text[rel_cs:rel_ce]))

    if len(box_words) >= 2:
        hit_indices = [
            idx for idx, m in enumerate(box_words)
            if max(box_s + m.start(), rel_cs) < min(box_s + m.end(), rel_ce)
        ]
        if hit_indices:
            r0 = hit_indices[0] / len(box_words)
            r1 = (hit_indices[-1] + 1) / len(box_words)
            return r0, r1, span_words, len(box_words)

    cjk_units = list(re.finditer(r"[\u3400-\u4dbf\u4e00-\u9fff]|[A-Za-z0-9]+", box_slice))
    if len(cjk_units) >= 2:
        hit_cjk = [
            idx for idx, m in enumerate(cjk_units)
            if max(box_s + m.start(), rel_cs) < min(box_s + m.end(), rel_ce)
        ]
        sp_cjk = len(re.findall(r"[\u3400-\u4dbf\u4e00-\u9fff]|[A-Za-z0-9]+", en_text[rel_cs:rel_ce]))
        est_sp_w = max(1, round(sp_cjk * 0.55)) if sp_cjk > 0 else 0
        est_box_w = max(1, round(len(cjk_units) * 0.55))
        if hit_cjk:
            r0 = hit_cjk[0] / len(cjk_units)
            r1 = (hit_cjk[-1] + 1) / len(cjk_units)
            return r0, r1, est_sp_w, est_box_w

    box_len = max(1, box_e - box_s)
    r0 = max(0.0, min(1.0, (rel_cs - box_s) / box_len))
    r1 = max(r0, min(1.0, (rel_ce - box_s) / box_len))
    return r0, r1, span_words, max(1, len(box_words))


def local_align_col3_spans(
    col3_text: str,
    old_en_full: str,
    spans: List[dict],
    find_fuzzy_fn: Callable,
) -> Dict[str, str]:
    """
    100% offline / zero-API approximate Col 3 span aligner supporting all global scripts:
      - Maps RV spans (and >=75% sentence-coverage spans) to full corresponding foreign sentence(s).
      - Maps UP local spans via clause structure + digit anchors + proportional token/grapheme offsets,
        snapping cleanly to whole words on space-delimited scripts and grapheme clusters (`M*` safe)
        on CJK / Thai / Myanmar / Khmer / Lao scripts.
      - Enforces strict monotonic non-overlapping cursor progression verified by `find_fuzzy_fn`.
    """
    result: Dict[str, str] = {}
    if not col3_text or not col3_text.strip() or not spans:
        return result

    en_full = old_en_full if (old_en_full and old_en_full.strip()) else _reconstruct_fallback_old_en(spans)
    if not en_full:
        en_full = " ".join((s.get("old_en_in_col3") or s.get("new_en") or "edit") for s in spans)

    # 1. Pre-locate all spans `[c_start, c_end]` inside `en_full` monotonically
    en_positions: List[Tuple[int, int]] = []
    en_cur = 0
    for idx, sp in enumerate(spans):
        cs, ce = _locate_span_in_old_en(en_full, sp, en_cur, idx, len(spans), find_fuzzy_fn)
        en_positions.append((cs, ce))
        en_cur = max(en_cur, ce)

    # 2. Split `en_full` and `col3_text` into sentence ranges (supporting both English and 2-col Chinese source)
    src_is_en = not bool(re.search(r"[\u3400-\u4dbf\u4e00-\u9fff]", en_full))
    en_sents = _split_sentences_ranges(en_full, is_english=src_is_en, include_semi_colon=False)
    for_sents = _split_sentences_ranges(col3_text, is_english=False, include_semi_colon=False)

    # If primary sentence counts differ by >1, check whether splitting at `;:` aligns sentence boundaries better
    if abs(len(en_sents) - len(for_sents)) > 1:
        en_semi = _split_sentences_ranges(en_full, is_english=src_is_en, include_semi_colon=True)
        for_semi = _split_sentences_ranges(col3_text, is_english=False, include_semi_colon=True)
        if abs(len(en_semi) - len(for_semi)) < abs(len(en_sents) - len(for_sents)):
            en_sents, for_sents = en_semi, for_semi

    # For multi-sentence English vs single-block Thai/Lao without sentence punctuation, partition space-clauses
    if len(en_sents) >= 2 and len(for_sents) == 1 and _THAI_LAO_RE.search(col3_text):
        tl_clauses = _split_clauses_ranges(col3_text, for_sents[0][0], for_sents[0][1], is_english=False)
        if len(tl_clauses) >= len(en_sents):
            grouped: List[Tuple[int, int]] = []
            n_tc = len(tl_clauses)
            n_es = len(en_sents)
            for k in range(n_es):
                g0 = int(round(k * n_tc / n_es))
                g1 = max(g0 + 1, int(round((k + 1) * n_tc / n_es)))
                grouped.append((tl_clauses[g0][0], tl_clauses[min(n_tc - 1, g1 - 1)][1]))
            for_sents = grouped

    if not en_sents:
        ts, te = _trim_range(en_full, 0, len(en_full))
        en_sents = [(ts, te)] if ts < te else [(0, len(en_full))]
    if not for_sents:
        fs, fe = _trim_range(col3_text, 0, len(col3_text))
        if fs >= fe:
            return result
        for_sents = [(fs, fe)]

    is_non_space = _is_non_space_script_text(col3_text)

    # Pre-map each span to its English `[s_first, s_last]` and Foreign `[f_first, f_last]` sentence indices
    span_sent_maps: List[Tuple[int, int, int, int]] = []
    for cs, ce in en_positions:
        ov = [
            i for i, (es, ee) in enumerate(en_sents)
            if (cs < ce and max(es, cs) < min(ee, ce)) or (cs == ce and es <= cs <= ee)
        ]
        if not ov:
            mid = (cs + ce) / 2.0
            nearest = min(
                range(len(en_sents)),
                key=lambda i: min(abs(en_sents[i][0] - mid), abs(en_sents[i][1] - mid)),
            )
            ov = [nearest]
        sf, sl = ov[0], ov[-1]
        ff, fl = _map_sentence_indices(sf, sl, en_sents, for_sents, len(en_full), len(col3_text))
        span_sent_maps.append((sf, sl, ff, fl))

    col3_end_trimmed = _trim_range(col3_text, 0, len(col3_text))[1]
    cursor = 0

    for idx, sp in enumerate(spans):
        sid = sp["span_id"]
        if cursor >= col3_end_trimmed:
            break

        c_start, c_end = en_positions[idx]
        s_first, s_last, f_first, f_last = span_sent_maps[idx]

        # Advance foreign sentence if cursor already passed it
        while f_last < len(for_sents) - 1 and for_sents[f_last][1] <= cursor:
            f_first = min(len(for_sents) - 1, f_first + 1)
            f_last = max(f_first, f_last + 1)

        en_sent_s, en_sent_e = en_sents[s_first][0], en_sents[s_last][1]
        for_sent_s, for_sent_e = for_sents[f_first][0], for_sents[f_last][1]
        if for_sent_e <= cursor:
            for_sent_s, for_sent_e = cursor, col3_end_trimmed

        # Count how many subsequent spans also land in this exact same `f_last` foreign sentence
        later_same_sent = sum(
            1 for j in range(idx + 1, len(spans))
            if span_sent_maps[j][2] <= f_last
        )

        # Coverage check for Step 4 (`rv` or >=75% of English sentence)
        _, _, span_w_cnt, sent_w_cnt = _word_ratio_in_box(en_full, en_sent_s, en_sent_e, c_start, c_end)
        char_cov = max(0, min(c_end, en_sent_e) - max(c_start, en_sent_s)) / max(1, en_sent_e - en_sent_s)
        word_cov = span_w_cnt / max(1, sent_w_cnt)
        is_full_sentence = (sp.get("type") == "rv") or (max(char_cov, word_cov) >= 0.75 and later_same_sent == 0)

        cand_s, cand_e = for_sent_s, for_sent_e

        if is_full_sentence:
            cand_s = max(cursor, for_sent_s)
            cand_e = for_sent_e
            if later_same_sent > 0 and sp.get("type") == "rv":
                # Reserve a small tail only if another span is sharing this foreign sentence
                clusters_all = _build_grapheme_clusters(col3_text, cand_s, cand_e)
                if len(clusters_all) > later_same_sent + 2:
                    cand_e = clusters_all[len(clusters_all) - 1 - later_same_sent][1]
            cand_s, cand_e = _snap_safe_codepoint_bounds(col3_text, cand_s, cand_e, max(cursor, 0), len(col3_text))
        else:
            # Step 5: Local `up` edit — narrow by sub-clauses when clause structure aligns
            en_box_s, en_box_e = en_sent_s, en_sent_e
            for_box_s, for_box_e = for_sent_s, for_sent_e

            en_clauses = _split_clauses_ranges(en_full, en_sent_s, en_sent_e, is_english=src_is_en)
            for_clauses = _split_clauses_ranges(col3_text, for_sent_s, for_sent_e, is_english=False)

            used_whole_clause = False
            if len(en_clauses) >= 2 and len(for_clauses) >= 2 and abs(len(en_clauses) - len(for_clauses)) <= (0 if len(en_clauses) == 2 else 1):
                cl_ov = [
                    ci for ci, (ecs, ece) in enumerate(en_clauses)
                    if (c_start < c_end and max(ecs, c_start) < min(ece, c_end))
                    or (c_start == c_end and ecs <= c_start <= ece)
                ]
                if cl_ov:
                    cf0, cf1 = _map_sentence_indices(
                        cl_ov[0],
                        cl_ov[-1],
                        en_clauses,
                        for_clauses,
                        en_sent_e - en_sent_s,
                        for_sent_e - for_sent_s,
                    )
                    cand_for_s, cand_for_e = for_clauses[cf0][0], for_clauses[cf1][1]
                    if cand_for_e > cursor:
                        en_box_s, en_box_e = en_clauses[cl_ov[0]][0], en_clauses[cl_ov[-1]][1]
                        for_box_s, for_box_e = max(cursor, cand_for_s), cand_for_e

                        # Check if another span shares this exact foreign clause
                        later_same_clause = any(
                            en_positions[j][0] < en_box_e for j in range(idx + 1, len(spans))
                        )
                        _, _, cl_sp_w, cl_tot_w = _word_ratio_in_box(en_full, en_box_s, en_box_e, c_start, c_end)
                        cl_char_cov = max(0, min(c_end, en_box_e) - max(c_start, en_box_s)) / max(1, en_box_e - en_box_s)
                        if max(cl_char_cov, cl_sp_w / max(1, cl_tot_w)) >= 0.78 and not later_same_clause:
                            cand_s, cand_e = _snap_safe_codepoint_bounds(
                                col3_text, for_box_s, for_box_e, cursor, len(col3_text)
                            )
                            used_whole_clause = True

            if not used_whole_clause:
                avail_s = max(cursor, for_box_s)
                avail_e = for_box_e
                if avail_s >= avail_e:
                    avail_s = cursor
                    avail_e = max(for_sent_e, col3_end_trimmed)

                rel_s, rel_e, sp_words, box_words = _word_ratio_in_box(en_full, en_box_s, en_box_e, c_start, c_end)
                # Adjust relative offsets if `cursor` already consumed the prefix of `for_box`
                if for_box_e > for_box_s and avail_s > for_box_s:
                    consumed_frac = (avail_s - for_box_s) / (for_box_e - for_box_s)
                    if consumed_frac < 0.95:
                        rel_s = max(0.0, (rel_s - consumed_frac) / (1.0 - consumed_frac))
                        rel_e = max(rel_s, min(1.0, (rel_e - consumed_frac) / (1.0 - consumed_frac)))

                # Optional exact digit anchor inside `col3_text[avail_s:avail_e]`
                en_digits = _DIGIT_TOKEN_RE.findall(en_full[c_start:c_end])
                digit_anchor_pos: Optional[int] = None
                if en_digits:
                    d_tok = en_digits[0]
                    d_hits = [m.start() + avail_s for m in re.finditer(re.escape(d_tok), col3_text[avail_s:avail_e])]
                    if len(d_hits) == 1:
                        digit_anchor_pos = d_hits[0]

                is_short_win = sp_words <= 3

                if not is_non_space:
                    # Branch 5A: Space-delimited scripts (Latin, Greek, Cyrillic, Arabic, Hebrew, Hindi, Korean, etc.)
                    raw_toks = [
                        (avail_s + m.start(), avail_s + m.end())
                        for m in re.finditer(r"\S+", col3_text[avail_s:avail_e])
                    ]
                    word_toks = [
                        (ts, te) for (ts, te) in raw_toks
                        if any(ch.isalnum() for ch in col3_text[ts:te])
                    ]
                    if not word_toks:
                        word_toks = raw_toks

                    # Reserve at least 1 word per subsequent span sharing the same sentence
                    usable_w_cnt = len(word_toks)
                    if later_same_sent > 0 and usable_w_cnt > later_same_sent:
                        # Check how many subsequent spans fall inside `[avail_s, avail_e]`
                        same_box_later = sum(
                            1 for j in range(idx + 1, len(spans))
                            if en_positions[j][0] < en_box_e or for_box_e == for_sent_e and span_sent_maps[j][2] <= f_last
                        )
                        if same_box_later > 0 and usable_w_cnt > same_box_later:
                            usable_w_cnt -= same_box_later

                    word_Pool = word_toks[: max(1, usable_w_cnt)]
                    w_n = len(word_Pool)

                    if w_n == 0:
                        cand_s, cand_e = _snap_safe_codepoint_bounds(col3_text, avail_s, avail_e, cursor, len(col3_text))
                    else:
                        if digit_anchor_pos is not None:
                            center_idx = min(
                                range(w_n),
                                key=lambda wi: abs(word_Pool[wi][0] - digit_anchor_pos),
                            )
                        else:
                            center_idx = min(w_n - 1, max(0, int(((rel_s + rel_e) / 2.0) * w_n)))

                        if is_short_win:
                            target_w = min(w_n, 3 if (sp_words == 0 or sp_words >= 3) else max(2 if w_n >= 2 and sp_words == 2 else 1, sp_words))
                            w0 = max(0, min(w_n - target_w, center_idx - (target_w // 2)))
                            w1 = min(w_n - 1, w0 + target_w - 1)
                        else:
                            est_w = max(2, round((sp_words / max(1, box_words)) * w_n))
                            est_w = min(w_n, est_w)
                            w0 = max(0, min(w_n - est_w, int(round(rel_s * w_n))))
                            w1 = max(w0, min(w_n - 1, max(int(round(rel_e * w_n)) - 1, w0 + est_w - 1)))

                        cand_s = word_Pool[w0][0]
                        cand_e = word_Pool[w1][1]

                        # Include detached typographic punctuation at outer boundaries (e.g. French « ... ; / ? »)
                        if w0 == 0 and rel_s <= 0.05 and raw_toks and raw_toks[0][1] <= cand_s:
                            lead_tok = col3_text[raw_toks[0][0] : raw_toks[0][1]]
                            if all(c in _OPENING_LEADERS for c in lead_tok):
                                cand_s = raw_toks[0][0]
                        if w1 == len(word_toks) - 1 and rel_e >= 0.92 and raw_toks and raw_toks[-1][0] >= cand_e:
                            tail_tok = col3_text[raw_toks[-1][0] : raw_toks[-1][1]]
                            if not any(c.isalnum() for c in tail_tok):
                                cand_e = raw_toks[-1][1]

                        cand_s, cand_e = _snap_safe_codepoint_bounds(col3_text, cand_s, cand_e, cursor, len(col3_text))
                else:
                    # Branch 5B: Non-space scripts (CJK, Thai, Myanmar, Khmer, Lao, Tibetan)
                    clusters = _build_grapheme_clusters(col3_text, avail_s, avail_e)
                    usable_c_cnt = len(clusters)
                    if later_same_sent > 0 and usable_c_cnt > later_same_sent * 2:
                        same_box_later = sum(
                            1 for j in range(idx + 1, len(spans))
                            if en_positions[j][0] < en_box_e or (for_box_e == for_sent_e and span_sent_maps[j][2] <= f_last)
                        )
                        if same_box_later > 0 and usable_c_cnt > same_box_later * 2:
                            usable_c_cnt -= same_box_later * 2

                    c_pool = clusters[: max(1, usable_c_cnt)]
                    c_n = len(c_pool)
                    if c_n == 0:
                        cand_s, cand_e = _snap_safe_codepoint_bounds(col3_text, avail_s, avail_e, cursor, len(col3_text))
                    else:
                        if digit_anchor_pos is not None:
                            center_ci = min(
                                range(c_n),
                                key=lambda ci: abs(c_pool[ci][0] - digit_anchor_pos),
                            )
                        else:
                            center_ci = min(c_n - 1, max(0, int(((rel_s + rel_e) / 2.0) * c_n)))

                        if is_short_win:
                            target_c = min(c_n, 6)
                            c0 = max(0, min(c_n - target_c, center_ci - (target_c // 2)))
                            c1 = min(c_n - 1, c0 + target_c - 1)
                        else:
                            est_c = max(5, round((rel_e - rel_s) * c_n))
                            est_c = min(c_n, est_c)
                            c0 = max(0, min(c_n - est_c, int(round(rel_s * c_n))))
                            c1 = max(c0, min(c_n - 1, max(int(round(rel_e * c_n)) - 1, c0 + est_c - 1)))

                        # Ensure minimum ~4-8 char window when available
                        while (c_pool[c1][1] - c_pool[c0][0]) < 4 and (c0 > 0 or c1 + 1 < c_n):
                            if c1 + 1 < c_n:
                                c1 += 1
                            elif c0 > 0:
                                c0 -= 1

                        cand_s = c_pool[c0][0]
                        cand_e = c_pool[c1][1]
                        cand_s, cand_e = _snap_safe_codepoint_bounds(col3_text, cand_s, cand_e, cursor, len(col3_text))

        # Step 6: Verify monotonic non-overlapping match with `find_fuzzy_fn`
        sub = col3_text[cand_s:cand_e].strip()
        if not sub:
            rem_s, rem_e = _snap_safe_codepoint_bounds(col3_text, cursor, min(col3_end_trimmed, cursor + 12), cursor, col3_end_trimmed)
            sub = col3_text[rem_s:rem_e].strip()

        if sub:
            loc = _call_find_fuzzy(find_fuzzy_fn, col3_text, sub, search_from=cursor)
            if loc is not None and loc[0] >= cursor:
                result[sid] = col3_text[loc[0] : loc[1]]
                cursor = loc[1]
            else:
                direct_pos = col3_text.find(sub, cursor)
                if direct_pos != -1:
                    result[sid] = sub
                    cursor = direct_pos + len(sub)

    return result


# ---------------- Offline / Fallback Col 1 (Chinese) Approximate Alignment ----------------

_CN_LEADING_PUNCT = frozenset("，、。；：！？,.;:!?）》】」』”’")
_CN_CLAUSE_DELIMS = frozenset("，、：；,:;")


def _split_cn_sentences_ranges(text: str, include_semicolon: bool = True) -> List[Tuple[int, int]]:
    """Split Chinese `text` into trimmed sentence ranges on `。？！\\n` (and `；;` when `include_semicolon=True`)."""
    n = len(text)
    if n == 0:
        return []
    active_terms = {"。", "？", "！", ".", "?", "!"}
    if include_semicolon:
        active_terms |= {"；", ";"}

    ranges: List[Tuple[int, int]] = []
    seg_start = 0
    i = 0
    while i < n:
        ch = text[i]
        if ch == "\n":
            ts, te = _trim_range(text, seg_start, i)
            if ts < te:
                ranges.append((ts, te))
            seg_start = i + 1
            i += 1
            continue
        if ch in active_terms:
            if ch == "." and i > 0 and i + 1 < n and text[i - 1].isdigit() and text[i + 1].isdigit():
                i += 1
                continue
            j = i + 1
            while j < n and text[j] in active_terms and text[j] != "\n":
                j += 1
            while j < n and text[j] in _CLOSING_TRAILERS:
                j += 1
            ts, te = _trim_range(text, seg_start, j)
            if ts < te:
                ranges.append((ts, te))
            seg_start = j
            i = j
            continue
        i += 1

    ts, te = _trim_range(text, seg_start, n)
    if ts < te:
        ranges.append((ts, te))
    return ranges


def _split_cn_clauses_ranges(
    text: str,
    s: int,
    e: int,
    include_dunhao: bool = True,
) -> List[Tuple[int, int]]:
    """Split Chinese sentence `text[s:e]` into sub-clauses on `，：—–` (plus `、` when `include_dunhao=True`)."""
    if s >= e:
        return []
    delims = {"，", "：", ",", ":", "—", "–"}
    if include_dunhao:
        delims.add("、")

    clauses: List[Tuple[int, int]] = []
    c_start = s
    i = s
    while i < e:
        ch = text[i]
        if ch in delims:
            if ch in (":", ",") and i > s and i + 1 < e and text[i - 1].isdigit() and text[i + 1].isdigit():
                i += 1
                continue
            j = i + 1
            if ch in ("—", "–"):
                while j < e and text[j] in ("—", "–"):
                    j += 1
            while j < e and text[j] in _CLOSING_TRAILERS:
                j += 1
            cs, ce = _trim_range(text, c_start, j)
            if cs < ce:
                clauses.append((cs, ce))
            c_start = j
            i = j
            continue
        i += 1

    cs, ce = _trim_range(text, c_start, e)
    if cs < ce:
        clauses.append((cs, ce))
    return clauses if clauses else [(s, e)]


def _merge_short_en_lead_clauses(
    en_text: str,
    clauses: List[Tuple[int, int]],
    target_count: int,
) -> List[Tuple[int, int]]:
    """
    Merge 1–2 word introductory English discourse clauses (e.g. 'But,', 'However,', 'In fact,')
    into the following clause when English has more comma splits than Chinese (`但就人而言，`).
    """
    merged = list(clauses)
    idx = 0
    while len(merged) > max(1, target_count) and idx + 1 < len(merged):
        cs, ce = merged[idx]
        w_cnt = len(_EN_WORD_RE.findall(en_text[cs:ce]))
        if w_cnt <= 2:
            merged[idx : idx + 2] = [(cs, merged[idx + 1][1])]
        else:
            idx += 1
    return merged


def _build_effective_en_for_col1(
    old_en_full: str,
    spans: List[dict],
    find_fuzzy_fn: Callable,
) -> Tuple[str, List[Tuple[int, int]]]:
    """
    Reconstruct effective English reference text (splicing `new_en` where available, falling back to
    `old_en_in_col3` on pure deletions) together with exact monotonic `(start, end)` bounds per span,
    so Col 1 (latest Chinese) aligns accurately against both replacements and pure insertions.
    """
    base_old = old_en_full if (old_en_full and old_en_full.strip()) else _reconstruct_fallback_old_en(spans)
    if base_old:
        orig_positions: List[Tuple[int, int]] = []
        en_cur = 0
        for idx, sp in enumerate(spans):
            cs, ce = _locate_span_in_old_en(base_old, sp, en_cur, idx, len(spans), find_fuzzy_fn)
            orig_positions.append((cs, ce))
            en_cur = max(en_cur, ce)

        parts: List[str] = []
        eff_positions: List[Tuple[int, int]] = []
        prev_end = 0
        cur_len = 0
        for idx, sp in enumerate(spans):
            ocs, oce = orig_positions[idx]
            ocs = max(prev_end, min(len(base_old), ocs))
            oce = max(ocs, min(len(base_old), oce))

            prefix_seg = base_old[prev_end:ocs]
            if prefix_seg:
                parts.append(prefix_seg)
                cur_len += len(prefix_seg)

            new_piece = (sp.get("new_en") or "").strip()
            old_piece = (sp.get("old_en_in_col3") or "").strip()
            eff_piece = new_piece if new_piece else old_piece

            if ocs == oce and eff_piece:
                if eff_piece[0] in ",;:" and parts:
                    # Pure clause insertion at sentence end (e.g. 'domain.' + ', and ...') replaces trailing '.'
                    stripped_prev = parts[-1].rstrip()
                    if stripped_prev.endswith("."):
                        ws_tail = parts[-1][len(stripped_prev) :]
                        parts[-1] = stripped_prev[:-1] + ws_tail
                        cur_len -= 1
                prev_ch = parts[-1][-1] if parts and parts[-1] else ""
                if prev_ch and not prev_ch.isspace() and eff_piece[0] not in ",.;:?!)]}”’\"'":
                    parts.append(" ")
                    cur_len += 1

            sp_s = cur_len
            if eff_piece:
                parts.append(eff_piece)
                cur_len += len(eff_piece)
            sp_e = cur_len

            if ocs == oce and eff_piece and oce < len(base_old):
                next_ch = base_old[oce]
                if not next_ch.isspace() and next_ch not in ",.;:?!)]}”’\"'":
                    parts.append(" ")
                    cur_len += 1

            eff_positions.append((sp_s, sp_e))
            prev_end = oce

        tail_seg = base_old[prev_end:]
        if tail_seg:
            parts.append(tail_seg)
        eff_en = "".join(parts)
        if eff_en.strip():
            return eff_en, eff_positions

    # Synthetic fallback when only spans with `new_en` / `old_en_in_col3` were supplied
    syn_parts: List[str] = []
    syn_positions: List[Tuple[int, int]] = []
    cur_len = 0
    for idx, sp in enumerate(spans):
        l_ctx = sp.get("left_old_context") or ""
        mid = (sp.get("new_en") or "").strip() or (sp.get("old_en_in_col3") or "").strip() or "edit"
        r_ctx = sp.get("right_old_context") or ""
        if idx == 0 and l_ctx:
            syn_parts.append(l_ctx)
            cur_len += len(l_ctx)
        if syn_parts and not syn_parts[-1].endswith((" ", "\n")) and mid[0] not in ",.;:?!":
            syn_parts.append(" ")
            cur_len += 1
        s0 = cur_len
        syn_parts.append(mid)
        cur_len += len(mid)
        s1 = cur_len
        syn_positions.append((s0, s1))
        if r_ctx:
            if not r_ctx[0].isspace() and r_ctx[0] not in ",.;:?!":
                syn_parts.append(" ")
                cur_len += 1
            syn_parts.append(r_ctx)
            cur_len += len(r_ctx)
    return "".join(syn_parts), syn_positions


def _refine_cn_local_window_bounds(
    col1_text: str,
    cand_s: int,
    cand_e: int,
    min_s: int,
    max_e: int,
    en_span_slice: str,
) -> Tuple[int, int]:
    """Clean Chinese local UP window edges so punctuation/comma boundaries never leave 1-2 stray chars across a comma."""
    s = max(min_s, min(max_e, cand_s))
    e = max(s, min(max_e, cand_e))
    en_trim = (en_span_slice or "").strip()
    en_starts_punct = bool(en_trim and en_trim[0] in ":;,.")
    en_ends_punct = bool(en_trim and en_trim[-1] in ",;:.")

    # Do not begin a normal lexical CJK highlight on closing/mid punctuation
    if not en_starts_punct:
        while s < e and col1_text[s] in _CN_LEADING_PUNCT:
            s += 1

    # Never split in the middle of an ASCII word or number/verse token (e.g. '9:50' or '144,000')
    while s > min_s and s < max_e and col1_text[s].isascii() and (col1_text[s].isalnum() or col1_text[s] in ":.") and col1_text[s - 1].isascii() and col1_text[s - 1].isalnum():
        s -= 1
    while e < max_e and e > s and col1_text[e - 1].isascii() and (col1_text[e - 1].isalnum() or col1_text[e - 1] in ":.") and col1_text[e].isascii() and (col1_text[e].isalnum() or col1_text[e] in ":."):
        e += 1

    # Include immediately preceding opening quote/bracket if window starts right inside it
    if s > min_s and s < e and col1_text[s - 1] in _OPENING_LEADERS:
        s -= 1

    # Complete immediately adjacent closing paren/quote if window opened one (e.g. '（路9:50' -> '（路9:50）')
    _pair_map = {"（": "）", "(": ")", "《": "》", "“": "”", "「": "」", "『": "』", "【": "】", "[": "]"}
    if s < e and col1_text[s] in _pair_map:
        closer = _pair_map[col1_text[s]]
        if closer not in col1_text[s:e] and e < max_e and col1_text[e] == closer:
            e += 1

    # Avoid grabbing 1-2 stray characters before a comma at the left edge (e.g. '此，这样的人' -> '这样的人')
    for pos in range(s, min(e, s + 3)):
        if col1_text[pos] in _CN_CLAUSE_DELIMS and (e - (pos + 1)) >= 4:
            if not (en_starts_punct and pos == s):
                s = pos + 1
                while s < e and col1_text[s].isspace():
                    s += 1
                break

    # Avoid spilling 1-2 stray characters past a comma at the right edge (e.g. '正因如此，这' -> '正因如此，' / '正因如此')
    for pos in range(e - 1, max(s, e - 3) - 1, -1):
        if col1_text[pos] in _CN_CLAUSE_DELIMS and (pos - s) >= 4 and (e - (pos + 1)) in (1, 2):
            e = (pos + 1) if en_ends_punct else pos
            break

    return _snap_safe_codepoint_bounds(col1_text, s, e, min_s, max_e)


def local_align_col1_spans(
    col1_text: str,
    old_en_full: str,
    spans: List[dict],
    find_fuzzy_fn: Callable,
) -> Dict[str, str]:
    """
    100% offline / zero-API Col 1 (Chinese) span aligner:
      - Splits `col1_text` into Chinese sentences (`。？！；\\n`) and sub-clauses (`，、：`).
      - Maps each English span's sentence/clause/character ratio onto `col1_text` using
        `new_en` / `old_en_in_col3` length & position.
      - Selects full corresponding Chinese sentence(s) for `rv` full-sentence spans, and
        corresponding Chinese clause or proportional ~4–12 CJK character window for `up` local spans.
      - Enforces strict monotonic non-overlapping cursor order verified via `find_fuzzy_fn`.
    """
    result: Dict[str, str] = {}
    if not col1_text or not col1_text.strip() or not spans:
        return result

    en_full, en_positions = _build_effective_en_for_col1(old_en_full, spans, find_fuzzy_fn)
    if not en_full:
        return result

    # Choose sentence segmentation (`。？！；\n` vs `。？！\n`) that best aligns Chinese & English sentence counts
    en_semi = _split_sentences_ranges(en_full, is_english=True, include_semi_colon=True)
    cn_semi = _split_cn_sentences_ranges(col1_text, include_semicolon=True)
    en_prim = _split_sentences_ranges(en_full, is_english=True, include_semi_colon=False)
    cn_prim = _split_cn_sentences_ranges(col1_text, include_semicolon=False)

    candidates = [
        (en_semi, cn_semi),
        (en_prim, cn_semi),
        (en_prim, cn_prim),
    ]
    valid_cands = [(es, cs) for (es, cs) in candidates if es and cs]
    if valid_cands:
        en_sents, cn_sents = min(valid_cands, key=lambda pair: abs(len(pair[0]) - len(pair[1])))
    else:
        ts_e, te_e = _trim_range(en_full, 0, len(en_full))
        en_sents = [(ts_e, te_e)] if ts_e < te_e else [(0, len(en_full))]
        ts_c, te_c = _trim_range(col1_text, 0, len(col1_text))
        if ts_c >= te_c:
            return result
        cn_sents = [(ts_c, te_c)]

    # Pre-map each span to English `[s_first, s_last]` and Chinese `[c_first, c_last]` sentence indices
    span_sent_maps: List[Tuple[int, int, int, int]] = []
    for cs, ce in en_positions:
        ov = [
            i for i, (es, ee) in enumerate(en_sents)
            if (cs < ce and max(es, cs) < min(ee, ce)) or (cs == ce and es <= cs <= ee)
        ]
        if not ov:
            mid = (cs + ce) / 2.0
            nearest = min(
                range(len(en_sents)),
                key=lambda i: min(abs(en_sents[i][0] - mid), abs(en_sents[i][1] - mid)),
            )
            ov = [nearest]
        sf, sl = ov[0], ov[-1]
        cf, cl = _map_sentence_indices(sf, sl, en_sents, cn_sents, len(en_full), len(col1_text))
        span_sent_maps.append((sf, sl, cf, cl))

    # Also record which English sentences map to each Chinese sentence (for N-to-1 English->Chinese sentences)
    en_sent_to_cn: List[Tuple[int, int]] = [
        _map_sentence_indices(ei, ei, en_sents, cn_sents, len(en_full), len(col1_text))
        for ei in range(len(en_sents))
    ]

    col1_end_trimmed = _trim_range(col1_text, 0, len(col1_text))[1]
    cursor = 0

    for idx, sp in enumerate(spans):
        sid = sp["span_id"]
        if cursor >= col1_end_trimmed:
            break

        c_start, c_end = en_positions[idx]
        s_first, s_last, c_first, c_last = span_sent_maps[idx]

        while c_last < len(cn_sents) - 1 and cn_sents[c_last][1] <= cursor:
            c_first = min(len(cn_sents) - 1, c_first + 1)
            c_last = max(c_first, c_last + 1)

        en_sent_s, en_sent_e = en_sents[s_first][0], en_sents[s_last][1]
        # Expand English scope for UP ratio/clause mapping if multiple English sentences map into `[c_first, c_last]`
        sharing_en_idxs = [
            ei for ei, (ecf, ecl) in enumerate(en_sent_to_cn)
            if max(ecf, c_first) <= min(ecl, c_last)
        ]
        en_scope_s = min(en_sent_s, en_sents[sharing_en_idxs[0]][0]) if sharing_en_idxs else en_sent_s
        en_scope_e = max(en_sent_e, en_sents[sharing_en_idxs[-1]][1]) if sharing_en_idxs else en_sent_e

        cn_sent_s, cn_sent_e = cn_sents[c_first][0], cn_sents[c_last][1]
        if cn_sent_e <= cursor:
            cn_sent_s, cn_sent_e = cursor, col1_end_trimmed

        later_same_sent = sum(
            1 for j in range(idx + 1, len(spans))
            if span_sent_maps[j][2] <= c_last
        )

        _, _, span_w_cnt, sent_w_cnt = _word_ratio_in_box(en_full, en_scope_s, en_scope_e, c_start, c_end)
        char_cov = max(0, min(c_end, en_scope_e) - max(c_start, en_scope_s)) / max(1, en_scope_e - en_scope_s)
        word_cov = span_w_cnt / max(1, sent_w_cnt)
        is_full_sentence = (sp.get("type") == "rv") or (max(char_cov, word_cov) >= 0.75 and later_same_sent == 0)

        cand_s, cand_e = cn_sent_s, cn_sent_e

        if is_full_sentence:
            cand_s = max(cursor, cn_sent_s)
            cand_e = cn_sent_e
            if later_same_sent > 0 and sp.get("type") == "rv":
                clusters_all = _build_grapheme_clusters(col1_text, cand_s, cand_e)
                if len(clusters_all) > later_same_sent * 2 + 2:
                    cand_e = clusters_all[len(clusters_all) - 1 - later_same_sent * 2][1]
            cand_s, cand_e = _snap_safe_codepoint_bounds(col1_text, cand_s, cand_e, max(cursor, 0), len(col1_text))
        else:
            en_box_s, en_box_e = en_scope_s, en_scope_e
            cn_box_s, cn_box_e = cn_sent_s, cn_sent_e

            # Split `en_scope` into clauses (including `.?!` if multiple English sentences mapped to 1 Chinese sentence)
            en_clauses_raw: List[Tuple[int, int]] = []
            for sub_es, sub_ee in _split_sentences_ranges(en_full[en_scope_s:en_scope_e], is_english=True, include_semi_colon=True):
                en_clauses_raw.extend(
                    _split_clauses_ranges(en_full, en_scope_s + sub_es, en_scope_s + sub_ee, is_english=True)
                )
            if not en_clauses_raw:
                en_clauses_raw = _split_clauses_ranges(en_full, en_scope_s, en_scope_e, is_english=True)
            cn_clauses_major = _split_cn_clauses_ranges(col1_text, cn_sent_s, cn_sent_e, include_dunhao=False)
            cn_clauses_fine = _split_cn_clauses_ranges(col1_text, cn_sent_s, cn_sent_e, include_dunhao=True)

            en_for_major = _merge_short_en_lead_clauses(en_full, en_clauses_raw, len(cn_clauses_major))
            en_for_fine = _merge_short_en_lead_clauses(en_full, en_clauses_raw, len(cn_clauses_fine))

            if abs(len(en_for_fine) - len(cn_clauses_fine)) < abs(len(en_for_major) - len(cn_clauses_major)):
                en_clauses, cn_clauses = en_for_fine, cn_clauses_fine
            else:
                en_clauses, cn_clauses = en_for_major, cn_clauses_major

            used_whole_clause = False
            if (
                len(en_clauses) >= 2
                and len(cn_clauses) >= 2
                and abs(len(en_clauses) - len(cn_clauses)) <= (0 if len(en_clauses) == 2 else 1)
            ):
                cl_ov = [
                    ci for ci, (ecs, ece) in enumerate(en_clauses)
                    if (c_start < c_end and max(ecs, c_start) < min(ece, c_end))
                    or (c_start == c_end and ecs <= c_start <= ece)
                ]
                if cl_ov:
                    cf0, cf1 = _map_sentence_indices(
                        cl_ov[0],
                        cl_ov[-1],
                        en_clauses,
                        cn_clauses,
                        en_scope_e - en_scope_s,
                        cn_sent_e - cn_sent_s,
                    )
                    cand_cn_s, cand_cn_e = cn_clauses[cf0][0], cn_clauses[cf1][1]
                    if cand_cn_e > cursor:
                        en_box_s, en_box_e = en_clauses[cl_ov[0]][0], en_clauses[cl_ov[-1]][1]
                        cn_box_s, cn_box_e = max(cursor, cand_cn_s), cand_cn_e

                        later_same_clause = any(
                            en_positions[j][0] < en_box_e for j in range(idx + 1, len(spans))
                        )
                        _, _, cl_sp_w, cl_tot_w = _word_ratio_in_box(en_full, en_box_s, en_box_e, c_start, c_end)
                        cl_char_cov = max(0, min(c_end, en_box_e) - max(c_start, en_box_s)) / max(1, en_box_e - en_box_s)
                        if max(cl_char_cov, cl_sp_w / max(1, cl_tot_w)) >= 0.75 and not later_same_clause:
                            cand_s, cand_e = _snap_safe_codepoint_bounds(
                                col1_text, cn_box_s, cn_box_e, cursor, len(col1_text)
                            )
                            used_whole_clause = True

            if not used_whole_clause:
                avail_s = max(cursor, cn_box_s)
                avail_e = cn_box_e
                if avail_s >= avail_e:
                    avail_s = cursor
                    avail_e = max(cn_sent_e, col1_end_trimmed)

                rel_s, rel_e, sp_words, _box_words = _word_ratio_in_box(en_full, en_box_s, en_box_e, c_start, c_end)
                if cn_box_e > cn_box_s and avail_s > cn_box_s:
                    consumed_frac = (avail_s - cn_box_s) / (cn_box_e - cn_box_s)
                    if consumed_frac < 0.95:
                        rel_s = max(0.0, (rel_s - consumed_frac) / (1.0 - consumed_frac))
                        rel_e = max(rel_s, min(1.0, (rel_e - consumed_frac) / (1.0 - consumed_frac)))

                en_span_slice = en_full[c_start:c_end]
                en_digits = _DIGIT_TOKEN_RE.findall(en_span_slice)
                digit_anchor_span: Optional[Tuple[int, int]] = None
                if en_digits:
                    d_tok = en_digits[0]
                    d_hits = [
                        (avail_s + m.start(), avail_s + m.end())
                        for m in re.finditer(re.escape(d_tok), col1_text[avail_s:avail_e])
                    ]
                    if len(d_hits) == 1:
                        digit_anchor_span = d_hits[0]

                clusters = _build_grapheme_clusters(col1_text, avail_s, avail_e)
                usable_c_cnt = len(clusters)
                if later_same_sent > 0 and usable_c_cnt > later_same_sent * 2:
                    same_box_later = sum(
                        1 for j in range(idx + 1, len(spans))
                        if en_positions[j][0] < en_box_e or (cn_box_e == cn_sent_e and span_sent_maps[j][2] <= c_last)
                    )
                    if same_box_later > 0 and usable_c_cnt > same_box_later * 2:
                        usable_c_cnt -= same_box_later * 2

                c_pool = clusters[: max(1, usable_c_cnt)]
                c_n = len(c_pool)
                if c_n == 0:
                    cand_s, cand_e = _snap_safe_codepoint_bounds(col1_text, avail_s, avail_e, cursor, len(col1_text))
                else:
                    if digit_anchor_span is not None:
                        d_mid = (digit_anchor_span[0] + digit_anchor_span[1]) / 2.0
                        center_ci = min(
                            range(c_n),
                            key=lambda ci: abs(((c_pool[ci][0] + c_pool[ci][1]) / 2.0) - d_mid),
                        )
                    else:
                        center_ci = min(c_n - 1, max(0, int(((rel_s + rel_e) / 2.0) * c_n)))

                    raw_prop = round((rel_e - rel_s) * c_n)
                    if sp_words <= 3:
                        min_win = 5 if c_n >= 6 else 4
                        target_c = min(c_n, max(min_win, min(8, raw_prop if raw_prop > 0 else 6)))
                        c0 = max(0, min(c_n - target_c, int(round(rel_s * c_n)) if rel_s <= 0.08 else center_ci - (target_c // 2)))
                        c1 = min(c_n - 1, c0 + target_c - 1)
                    else:
                        est_c = min(c_n, max(5, min(12, raw_prop) if (rel_e - rel_s) < 0.65 else raw_prop))
                        c0 = max(0, min(c_n - est_c, int(round(rel_s * c_n))))
                        c1 = max(c0, min(c_n - 1, max(int(round(rel_e * c_n)) - 1, c0 + est_c - 1)))
                        if (c1 - c0 + 1) > 12 and (rel_e - rel_s) < 0.55:
                            c1 = min(c_n - 1, c0 + 11)

                    while (c_pool[c1][1] - c_pool[c0][0]) < 4 and (c0 > 0 or c1 + 1 < c_n):
                        if c1 + 1 < c_n:
                            c1 += 1
                        elif c0 > 0:
                            c0 -= 1

                    cand_s = c_pool[c0][0]
                    cand_e = c_pool[c1][1]
                    if digit_anchor_span is not None:
                        cand_s = min(cand_s, digit_anchor_span[0])
                        cand_e = max(cand_e, min(c_pool[-1][1], digit_anchor_span[1]))
                    cand_s, cand_e = _refine_cn_local_window_bounds(
                        col1_text, cand_s, cand_e, cursor, len(col1_text), en_span_slice
                    )

        sub = col1_text[cand_s:cand_e].strip()
        if not sub:
            rem_s, rem_e = _snap_safe_codepoint_bounds(
                col1_text, cursor, min(col1_end_trimmed, cursor + 8), cursor, col1_end_trimmed
            )
            sub = col1_text[rem_s:rem_e].strip()

        if sub:
            loc = _call_find_fuzzy(find_fuzzy_fn, col1_text, sub, search_from=cursor)
            if loc is not None and loc[0] >= cursor:
                result[sid] = col1_text[loc[0] : loc[1]]
                cursor = loc[1]
            else:
                direct_pos = col1_text.find(sub, cursor)
                if direct_pos != -1:
                    result[sid] = sub
                    cursor = direct_pos + len(sub)

    return result


# ---------------- Step 1 & Step 2 Multi-Threaded Evaluators with Seamless Local Fallback ----------------

STEP1_SYS_PROMPT = """你是专业的多语种文稿修订分级专家。
对于给定的每个修订句（包含 sent_id、heuristic_suggested、change_ratio、edit_hunks、new_words、diff_preview），请严格按以下业务规则判定它属于 "yellow"（UP 局部改点）还是 "cyan"（RV 整句重翻）：

【三大核心量化准则（务必严格遵守）】：
1. **远距局部小改 -> 判为 "yellow"（UP 局部改点）**：
   - 同一段落中各处改点彼此相隔较远（如相隔五六行以上），且每处修改内容较少（仅涉及两三行以内，如替换个别词汇/同义短语、微调连词或句尾单个短语，例如 `The`->`This`、`represent`->`are`、`comes from having experienced it`->`is always gained in experience`），句子整体骨架清晰完整，译者可直接局部修改，无须重翻整句！
2. **长句连续改写 / 句式重构 -> 判为 "cyan"（RV 整句重翻）**：
   - 若修改内容涉及连续较长的句子（超过四五行，如 `new_words >= 42` 且有多处结构改写或增删 >= 12 词），或单句内句式骨架被重构（如改动比例 `change_ratio >= 0.36` 且改动 >= 6 词、或分号 `;` 后半句/前半句整套句型推倒重写），**必须将改点所在的整个完整句子判为 "cyan"**（重翻必须是完整的一整句话，绝不能从逗号或分号开始）！
3. **同段高频近距改点 -> 集中区段判为 "cyan"（RV 整句重翻）**：
   - 若同一段落中需要修改的内容出现频繁、各处改点间隔很短（仅一两行），导致连续多句话句式都在调整（`heuristic_suggested` 已提示 `"cyan"`），请将该集中区段的完整句子统一判为 `"cyan"`，避免黄青交错或满屏散乱碎片。

请严格输出 JSON：{"items": [{"sent_id": "...", "decision": "yellow" 或 "cyan"}]}。"""


def _build_optional_pool(
    api_key: Union[str, List[str], ApiKeyPool],
    rotation_mode: str = "round_robin",
) -> Optional[ApiKeyPool]:
    if isinstance(api_key, ApiKeyPool):
        return api_key if len(api_key) > 0 else None
    keys = parse_api_keys(api_key)
    if not keys:
        return None
    return ApiKeyPool(keys, rotation_mode=rotation_mode)


def evaluate_step1_decisions(
    review_items: List[dict],
    api_key: Union[str, List[str], ApiKeyPool] = "",
    model: str = "gemini-3.5-flash-lite",
    api_mode: str = "auto",
    rotation_mode: str = "round_robin",
    log_fn: Optional[Callable[[str], None]] = None,
    progress_fn: Optional[Callable[[float, str], None]] = None,
    enable_ai: bool = True,
    max_workers: Optional[int] = None,
    stats_out: Optional[dict] = None,
) -> Dict[str, str]:
    stats = {"used_local_fallback": False, "fallback_reason": ""}
    edited = [it for it in review_items if it.get("has_edit")]
    decisions: Dict[str, str] = {}
    if not edited:
        if progress_fn:
            progress_fn(1.0, "无需 AI 评估（0 处改动句）")
        if stats_out is not None:
            stats_out.update(stats)
        return decisions

    pool = _build_optional_pool(api_key, rotation_mode=rotation_mode) if enable_ai else None
    if not enable_ai or pool is None:
        stats["used_local_fallback"] = True
        if enable_ai and pool is None:
            stats["fallback_reason"] = "未填写 API Key"
        if log_fn:
            log_fn("  ℹ️ 初筛分级：使用本地内置规则完成黄(UP)/青(RV)分级...")
        for it in edited:
            sug = it.get("suggested", "yellow")
            decisions[it["sent_id"]] = sug if sug in ("yellow", "cyan") else "yellow"
        if progress_fn:
            progress_fn(1.0, f"本地规则初筛完成（共 {len(edited)} 句）")
        if stats_out is not None:
            stats_out.update(stats)
        return decisions

    schema = {
        "type": "object",
        "properties": {
            "items": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "sent_id": {"type": "string"},
                        "decision": {"type": "string", "enum": ["yellow", "cyan"]},
                    },
                    "required": ["sent_id", "decision"],
                },
            }
        },
        "required": ["items"],
    }

    batch_size = 20
    total_batches = (len(edited) + batch_size - 1) // batch_size
    workers = (
        min(total_batches, max(1, int(max_workers)))
        if max_workers is not None
        else min(total_batches, max(2, min(8, len(pool) * 2)))
    )

    state_lock = threading.Lock()
    io_lock = threading.Lock()
    fallback_event = threading.Event()
    completed_batches = 0

    def _safe_log(msg: str):
        if log_fn:
            with io_lock:
                log_fn(msg)

    def _safe_call_log(msg: str):
        if log_fn and not fallback_event.is_set():
            with io_lock:
                if not fallback_event.is_set():
                    log_fn(msg)

    def _safe_progress(frac: float, msg: str):
        if progress_fn:
            with io_lock:
                progress_fn(frac, msg)

    if total_batches > 1 and workers > 1:
        _safe_log(f"  ⚡ 启动 {workers} 线程并发评估 {len(edited)} 处改动句（共 {total_batches} 批，密钥池 {len(pool)} 个 Key）...")

    def _run_batch(b_idx: int) -> None:
        nonlocal completed_batches
        batch = edited[b_idx * batch_size : (b_idx + 1) * batch_size]
        local_dec: Dict[str, str] = {}

        if not fallback_event.is_set():
            stage_msg = f"AI 初筛评估：第 {b_idx + 1}/{total_batches} 批（本批 {len(batch)} 句 / 共 {len(edited)} 句）"
            _safe_call_log(f"  [{stage_msg}] 正在请求 {model}...")

            prompt_items = [
                {
                    "sent_id": it["sent_id"],
                    "heuristic_suggested": it.get("suggested", "yellow"),
                    "change_ratio": it.get("change_ratio", 0),
                    "edit_hunks": it.get("edit_hunks", 1),
                    "new_words": it.get("new_words", 0),
                    "diff_preview": it.get("diff_preview", ""),
                }
                for it in batch
            ]
            user_prompt = "请评估以下句子的黄(yellow)/青(cyan)分级：\n" + json.dumps(prompt_items, ensure_ascii=False, indent=2)

            try:
                res = call_gemini_json(
                    api_key=pool,
                    model=model,
                    system_instruction=STEP1_SYS_PROMPT,
                    user_prompt=user_prompt,
                    json_schema=schema,
                    api_mode=api_mode,
                    log_fn=_safe_call_log,
                )
                for entry in res.get("items", []):
                    sid = entry.get("sent_id")
                    dec = entry.get("decision")
                    if sid and dec in ("yellow", "cyan"):
                        local_dec[sid] = dec
            except Exception as ex:
                friendly_err = format_friendly_api_error(ex, total_keys=len(pool))
                should_log = False
                with state_lock:
                    if not fallback_event.is_set():
                        fallback_event.set()
                        stats["used_local_fallback"] = True
                        stats["fallback_reason"] = friendly_err
                        should_log = True
                if should_log:
                    _safe_log(f"  ⚠️ {friendly_err} -> 初筛已自动切换内置规则完成分级")

        for it in batch:
            sid = it["sent_id"]
            if sid not in local_dec:
                sug = it.get("suggested", "yellow")
                local_dec[sid] = sug if sug in ("yellow", "cyan") else "yellow"

        with state_lock:
            decisions.update(local_dec)
            completed_batches += 1
            done_cnt = completed_batches

        _safe_progress(done_cnt / total_batches, f"AI 初筛评估：已完成 {done_cnt}/{total_batches} 批")

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = [executor.submit(_run_batch, b_i) for b_i in range(total_batches)]
        for fut in futures:
            fut.result()

    if stats_out is not None:
        stats_out.update(stats)
    return decisions


STEP2_SYS_PROMPT = """你是多语种三列平行语料精准片段对齐专家（支持中文及全球任意外语，包括拉丁语系、斯拉夫语系、希腊语、阿拉伯/波斯等从右向左语系、日韩语、泰/缅/柬/老等东南亚无空格语系、印地/孟加拉等南亚语系）。
给定同一表格行的：
- 第 1 列完整中文段落 `col1_text`（若提供，对应修订后的最新中文）；
- 第 3 列完整外文段落 `col3_text`（对应修订前的旧版外文译文）；
- 以及若干第 2 列英文高亮片段 `spans`（每个片段含 `span_id`、类型 `type`["up"局部改点 / "rv"整句重翻]、对应的旧版英文 `old_en_in_col3`、新版英文 `new_en`、以及前后未高亮的英文上下文 `left_old_context` / `right_old_context`）。

请在**同一次请求**中，为每个 `span` 精准找出：
1. `col1_substring`（当输入提供非空 `col1_text` 时必填）：在 `col1_text` 中与 `new_en` / `old_en_in_col3`（结合前后上下文位置）语义完全对应的**原样连续中文子串**；
2. `col3_substring`：在 `col3_text` 中与 `old_en_in_col3`（结合前后上下文位置）语义完全对应的**原样连续外文子串**。

严格规则：
1. `col1_substring` 必须**一字不差地原样复制**自 `col1_text`，`col3_substring` 必须**一字不差地原样复制**自 `col3_text`（包含原文标点符号，绝不可自行翻译、改写、省略中间词或添加省略号 `...`）。
2. 若 `type == "rv"`（整句重翻），`col1_substring` 与 `col3_substring` 应分别截取对应的完整中文句子与完整外文句子。
3. 若英文片段仅为 ~3 个词的短窗口（如纯标点改点），`col1_substring` 只截取对应位置约 4~8 个汉字的短片段，`col3_substring` 只截取对应位置的约 3 个外文词（若外文为日语、泰语、缅甸语、高棉语等无空格分词语言，则截取约 4~8 个字符的短片段）。
4. 若英文片段是半句/分句（如到逗号、分号或连词处为止），`col1_substring` 与 `col3_substring` 也必须精准停在对应的中文分句与外文分句/连词边界，绝不可把 `left_old_context` 或 `right_old_context` 对应的未改动内容包进来！
5. 同一行内的多个 `spans` 在 `col1_text` 和 `col3_text` 中均按顺序先后出现，互不重叠。
请严格输出 JSON：{"alignments": [{"span_id": "...", "col1_substring": "...", "col3_substring": "..."}]}（若输入未提供 `col1_text`，则每项输出 `span_id` 与 `col3_substring`）。"""


def _resolve_side_with_fallback(
    text: str,
    raw_map: Dict[str, str],
    spans: List[dict],
    local_fallback_fn: Callable[[], Dict[str, str]],
    find_fuzzy_fn: Callable,
) -> Dict[str, str]:
    """
    Validate `raw_map` monotonically against `text` via `find_fuzzy_fn`, recovering any residual
    unmatched spans via bounded head/tail anchoring and `local_fallback_fn()` while preserving
    strict non-overlapping order across all spans in the row.
    """
    resolved: Dict[str, str] = {}
    if not text or not text.strip() or not spans:
        return resolved

    matched_locs: List[Optional[Tuple[int, int]]] = [None] * len(spans)
    probe_cur = 0
    for i, s in enumerate(spans):
        sid = s["span_id"]
        cand = (raw_map.get(sid) or "").strip()
        if not cand:
            continue
        loc = _call_find_fuzzy(find_fuzzy_fn, text, cand, search_from=probe_cur)
        if loc is not None and probe_cur <= loc[0] < loc[1]:
            matched_locs[i] = (loc[0], loc[1])
            probe_cur = loc[1]

    local_backup: Optional[Dict[str, str]] = None
    for i, s in enumerate(spans):
        if matched_locs[i] is not None:
            continue
        sid = s["span_id"]
        low_b = matched_locs[i - 1][1] if (i > 0 and matched_locs[i - 1] is not None) else 0
        high_b = next(
            (matched_locs[j][0] for j in range(i + 1, len(spans)) if matched_locs[j] is not None),
            len(text),
        )
        if low_b >= high_b:
            continue

        cand = (raw_map.get(sid) or "").strip()
        anchored = False
        words = cand.split()
        if len(words) >= 4:
            for n_anchor in (3, 2, 1):
                head = " ".join(words[:n_anchor])
                tail = " ".join(words[-n_anchor:])
                h_loc = _call_find_fuzzy(find_fuzzy_fn, text, head, search_from=low_b)
                if h_loc is not None and low_b <= h_loc[0] < h_loc[1] <= high_b:
                    t_loc = _call_find_fuzzy(find_fuzzy_fn, text, tail, search_from=h_loc[1])
                    if (
                        t_loc is not None
                        and h_loc[1] <= t_loc[0] < t_loc[1] <= high_b
                        and (t_loc[1] - h_loc[0]) <= max(len(cand) * 2 + 120, high_b - low_b)
                    ):
                        matched_locs[i] = (h_loc[0], t_loc[1])
                        anchored = True
                        break
        elif len(cand) >= 8:
            for n_ch in (6, 4, 3, 2):
                if len(cand) < n_ch * 2:
                    continue
                head = cand[:n_ch]
                tail = cand[-n_ch:]
                h_loc = _call_find_fuzzy(find_fuzzy_fn, text, head, search_from=low_b)
                if h_loc is not None and low_b <= h_loc[0] < h_loc[1] <= high_b:
                    t_loc = _call_find_fuzzy(find_fuzzy_fn, text, tail, search_from=h_loc[1])
                    if (
                        t_loc is not None
                        and h_loc[1] <= t_loc[0] < t_loc[1] <= high_b
                        and (t_loc[1] - h_loc[0]) <= max(len(cand) * 2 + 80, high_b - low_b)
                    ):
                        matched_locs[i] = (h_loc[0], t_loc[1])
                        anchored = True
                        break

        if not anchored:
            if local_backup is None:
                local_backup = local_fallback_fn()
            fb_cand = (local_backup.get(sid) or "").strip()
            fb_loc = _call_find_fuzzy(find_fuzzy_fn, text, fb_cand, search_from=low_b) if fb_cand else None
            if fb_loc is not None and low_b <= fb_loc[0] < fb_loc[1] <= high_b:
                matched_locs[i] = (fb_loc[0], fb_loc[1])
            else:
                # Bounded cluster fallback inside `[low_b, high_b]` so monotonic alignment never drops or overlaps
                rem_unres = 1
                for j in range(i + 1, len(spans)):
                    if matched_locs[j] is not None:
                        break
                    rem_unres += 1
                gap_clusters = _build_grapheme_clusters(text, low_b, high_b)
                if gap_clusters:
                    take_n = max(1, len(gap_clusters) // max(1, rem_unres))
                    take_n = min(len(gap_clusters), min(take_n, 10) if s.get("type") != "rv" else take_n)
                    gs, ge = _snap_safe_codepoint_bounds(
                        text, gap_clusters[0][0], gap_clusters[take_n - 1][1], low_b, high_b
                    )
                    sub_gap = text[gs:ge].strip()
                    if sub_gap:
                        g_loc = _call_find_fuzzy(find_fuzzy_fn, text, sub_gap, search_from=low_b)
                        if g_loc is not None and low_b <= g_loc[0] < g_loc[1] <= high_b:
                            matched_locs[i] = (g_loc[0], g_loc[1])

    for s, loc in zip(spans, matched_locs):
        if loc is not None and loc[0] < loc[1]:
            resolved[s["span_id"]] = text[loc[0] : loc[1]]
    return resolved


def _pack_row_alignments(
    spans: List[dict],
    has_col1: bool,
    col1_map: Dict[str, str],
    col3_map: Dict[str, str],
) -> Dict[str, Union[str, Dict[str, str]]]:
    out: Dict[str, Union[str, Dict[str, str]]] = {}
    if has_col1:
        for s in spans:
            sid = s["span_id"]
            c1 = col1_map.get(sid, "")
            c3 = col3_map.get(sid, "")
            if c1 or c3:
                out[sid] = {"col1_substring": c1, "col3_substring": c3}
    else:
        for s in spans:
            sid = s["span_id"]
            c3 = col3_map.get(sid, "")
            if c3:
                out[sid] = c3
    return out


def align_step2_spans(
    tasks_data: dict,
    find_fuzzy_fn: Callable,
    api_key: Union[str, List[str], ApiKeyPool] = "",
    model: str = "gemini-3.5-flash-lite",
    api_mode: str = "auto",
    rotation_mode: str = "round_robin",
    log_fn: Optional[Callable[[str], None]] = None,
    progress_fn: Optional[Callable[[float, str], None]] = None,
    target_lang_hint: str = "",
    enable_ai: bool = True,
    max_workers: Optional[int] = None,
    stats_out: Optional[dict] = None,
) -> Dict[str, Union[str, Dict[str, str]]]:
    stats = {"used_local_fallback": False, "fallback_reason": ""}
    row_tasks = tasks_data.get("row_tasks", [])
    alignments: Dict[str, Union[str, Dict[str, str]]] = {}
    if not row_tasks:
        if progress_fn:
            progress_fn(1.0, "无需对齐外文（0 处高亮）")
        if stats_out is not None:
            stats_out.update(stats)
        return alignments

    pool = _build_optional_pool(api_key, rotation_mode=rotation_mode) if enable_ai else None
    total_rows = len(row_tasks)
    any_col1 = any(
        bool((rt.get("col1_text") or "").strip()) and not bool(rt.get("bilingual_2col"))
        for rt in row_tasks
    )

    if not enable_ai or pool is None:
        stats["used_local_fallback"] = True
        if enable_ai and pool is None:
            stats["fallback_reason"] = "未填写 API Key"
        if log_fn:
            col_desc = "第 1 列中文与第 3 列外文" if any_col1 else "外文列"
            log_fn(f"  ℹ️ 同步标色：使用「本地位置估算模式」对齐共 {total_rows} 行{col_desc}高亮...")
        for r_i, rt in enumerate(row_tasks):
            col1_text = rt.get("col1_text", "")
            col3_text = rt.get("col3_text", "")
            old_en_full = rt.get("old_en_full", "")
            spans = rt.get("spans", [])
            has_col1 = bool(col1_text and col1_text.strip()) and not bool(rt.get("bilingual_2col"))

            col3_map = local_align_col3_spans(
                col3_text=col3_text,
                old_en_full=old_en_full,
                spans=spans,
                find_fuzzy_fn=find_fuzzy_fn,
            )
            col1_map = (
                local_align_col1_spans(
                    col1_text=col1_text,
                    old_en_full=old_en_full,
                    spans=spans,
                    find_fuzzy_fn=find_fuzzy_fn,
                )
                if has_col1
                else {}
            )
            alignments.update(_pack_row_alignments(spans, has_col1, col1_map, col3_map))
            if progress_fn:
                progress_fn((r_i + 1) / total_rows, f"本地位置估算：已完成 {r_i + 1}/{total_rows} 行")
        if stats_out is not None:
            stats_out.update(stats)
        return alignments

    schema_col3_only = {
        "type": "object",
        "properties": {
            "alignments": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "span_id": {"type": "string"},
                        "col3_substring": {"type": "string"},
                    },
                    "required": ["span_id", "col3_substring"],
                },
            }
        },
        "required": ["alignments"],
    }

    schema_dual = {
        "type": "object",
        "properties": {
            "alignments": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "span_id": {"type": "string"},
                        "col1_substring": {"type": "string"},
                        "col3_substring": {"type": "string"},
                    },
                    "required": ["span_id", "col1_substring", "col3_substring"],
                },
            }
        },
        "required": ["alignments"],
    }

    workers = (
        min(total_rows, max(1, int(max_workers)))
        if max_workers is not None
        else min(total_rows, max(2, min(8, len(pool) * 2)))
    )

    state_lock = threading.Lock()
    io_lock = threading.Lock()
    fallback_event = threading.Event()
    completed_rows = 0

    def _safe_log(msg: str):
        if log_fn:
            with io_lock:
                log_fn(msg)

    def _safe_call_log(msg: str):
        if log_fn and not fallback_event.is_set():
            with io_lock:
                if not fallback_event.is_set():
                    log_fn(msg)

    def _safe_progress(frac: float, msg: str):
        if progress_fn:
            with io_lock:
                progress_fn(frac, msg)

    if total_rows > 1 and workers > 1:
        target_lbl = "中/外文同步高亮" if any_col1 else "外文高亮"
        _safe_log(f"  ⚡ 启动 {workers} 线程并发对齐 {total_rows} 行{target_lbl}（密钥池 {len(pool)} 个 Key）...")

    def _align_single_row(r_i: int) -> None:
        nonlocal completed_rows
        rt = row_tasks[r_i]
        col1_text = rt.get("col1_text", "")
        col3_text = rt.get("col3_text", "")
        old_en_full = rt.get("old_en_full", "")
        spans = rt.get("spans", [])
        has_col1 = bool(col1_text and col1_text.strip()) and not bool(rt.get("bilingual_2col"))
        has_col3 = bool(col3_text and col3_text.strip())
        row_align: Dict[str, Union[str, Dict[str, str]]] = {}

        if not spans:
            with state_lock:
                completed_rows += 1
                done_r = completed_rows
            _safe_progress(done_r / total_rows, f"已完成 {done_r}/{total_rows} 行")
            return

        if fallback_event.is_set():
            col3_map = local_align_col3_spans(col3_text, old_en_full, spans, find_fuzzy_fn) if has_col3 else {}
            col1_map = local_align_col1_spans(col1_text, old_en_full, spans, find_fuzzy_fn) if has_col1 else {}
            row_align = _pack_row_alignments(spans, has_col1, col1_map, col3_map)
        else:
            art_short = (rt.get("article_title") or "")[:18]
            lbl_kind = "AI 中/外文对齐" if has_col1 else "AI 外文对齐"
            stage_msg = f"{lbl_kind}：第 {r_i + 1}/{total_rows} 行（《{art_short}》，{len(spans)} 处高亮）"
            _safe_call_log(f"  [{stage_msg}]...")

            row_col1: Dict[str, str] = {}
            row_col3: Dict[str, str] = {}
            pending_spans = list(spans)
            feedback_msg = ""
            switched_to_local = False
            active_schema = schema_dual if has_col1 else schema_col3_only

            for attempt in range(3):
                if fallback_event.is_set():
                    switched_to_local = True
                    break

                payload_obj: dict = {
                    "target_language": target_lang_hint or "auto",
                }
                if has_col1:
                    payload_obj["col1_text"] = col1_text
                payload_obj["col3_text"] = col3_text
                payload_obj["spans"] = [
                    {
                        "span_id": s["span_id"],
                        "type": s.get("type", "up"),
                        "old_en_in_col3": s.get("old_en_in_col3", ""),
                        "new_en": s.get("new_en", ""),
                        "left_old_context": s.get("left_old_context", ""),
                        "right_old_context": s.get("right_old_context", ""),
                    }
                    for s in pending_spans
                ]

                user_prompt = json.dumps(payload_obj, ensure_ascii=False, indent=2)
                if feedback_msg:
                    user_prompt = feedback_msg + "\n\n请重新精确截取以下未匹配成功的片段：\n" + user_prompt

                try:
                    res = call_gemini_json(
                        api_key=pool,
                        model=model,
                        system_instruction=STEP2_SYS_PROMPT,
                        user_prompt=user_prompt,
                        json_schema=active_schema,
                        api_mode=api_mode,
                        log_fn=_safe_call_log,
                    )
                except Exception as ex:
                    friendly_err = format_friendly_api_error(ex, total_keys=len(pool))
                    should_log = False
                    with state_lock:
                        if not fallback_event.is_set():
                            fallback_event.set()
                            stats["used_local_fallback"] = True
                            stats["fallback_reason"] = friendly_err
                            should_log = True
                    if should_log:
                        col_lbl = "第 1 列与第 3 列标色" if any_col1 else "第 3 列标色"
                        _safe_log(f"  ⚠️ {friendly_err} -> 已自动切换「本地位置估算模式」完成{col_lbl}")
                    switched_to_local = True
                    break

                prev_col1 = dict(row_col1)
                prev_col3 = dict(row_col3)
                returned_items = {
                    item["span_id"]: item
                    for item in (res.get("alignments") or [])
                    if isinstance(item, dict) and "span_id" in item
                }
                for s in pending_spans:
                    sid = s["span_id"]
                    it = returned_items.get(sid)
                    if not it:
                        continue
                    if has_col1 and "col1_substring" in it and str(it["col1_substring"]).strip():
                        row_col1[sid] = str(it["col1_substring"]).strip()
                    if "col3_substring" in it and str(it["col3_substring"]).strip():
                        row_col3[sid] = str(it["col3_substring"]).strip()

                err_lines: List[str] = []
                failed_ids = set()

                if has_col1:
                    cur1 = 0
                    for s in spans:
                        sid = s["span_id"]
                        cand1 = row_col1.get(sid, "")
                        loc1 = _call_find_fuzzy(find_fuzzy_fn, col1_text, cand1, search_from=cur1) if cand1 else None
                        if loc1 is None or loc1[0] < cur1:
                            # Keep previous valid col1_substring if retry regressed
                            old_c1 = prev_col1.get(sid, "")
                            old_loc1 = _call_find_fuzzy(find_fuzzy_fn, col1_text, old_c1, search_from=cur1) if old_c1 else None
                            if old_loc1 is not None and old_loc1[0] >= cur1:
                                row_col1[sid] = col1_text[old_loc1[0] : old_loc1[1]]
                                cur1 = old_loc1[1]
                            else:
                                failed_ids.add(sid)
                                err_lines.append(
                                    f"- {sid} 返回的 col1_substring {cand1!r} 在 col1_text 中未按顺序找到！请务必从 col1_text 原文一字不差复制连续中文片段。"
                                )
                        else:
                            row_col1[sid] = col1_text[loc1[0] : loc1[1]]
                            cur1 = loc1[1]

                if has_col3:
                    cur3 = 0
                    for s in spans:
                        sid = s["span_id"]
                        cand3 = row_col3.get(sid, "")
                        loc3 = _call_find_fuzzy(find_fuzzy_fn, col3_text, cand3, search_from=cur3) if cand3 else None
                        if loc3 is None or loc3[0] < cur3:
                            old_c3 = prev_col3.get(sid, "")
                            old_loc3 = _call_find_fuzzy(find_fuzzy_fn, col3_text, old_c3, search_from=cur3) if old_c3 else None
                            if old_loc3 is not None and old_loc3[0] >= cur3:
                                row_col3[sid] = col3_text[old_loc3[0] : old_loc3[1]]
                                cur3 = old_loc3[1]
                            else:
                                failed_ids.add(sid)
                                err_lines.append(
                                    f"- {sid} 返回的 col3_substring {cand3!r} 在 col3_text 中未按顺序找到！请务必从 col3_text 原文一字不差复制完整连续片段（不要用省略号，不要改标点）。"
                                )
                        else:
                            row_col3[sid] = col3_text[loc3[0] : loc3[1]]
                            cur3 = loc3[1]

                if not failed_ids:
                    break

                pending_spans = [s for s in spans if s["span_id"] in failed_ids]
                feedback_msg = "上一次返回的部分子串未能匹配原文：\n" + "\n".join(err_lines)
                if attempt + 1 < 3:
                    _safe_call_log(f"    ↻ 自动修正重试 {len(pending_spans)} 处子串边界（第 {attempt + 2} 次）...")

            if switched_to_local:
                col3_final = local_align_col3_spans(col3_text, old_en_full, spans, find_fuzzy_fn) if has_col3 else {}
                col1_final = local_align_col1_spans(col1_text, old_en_full, spans, find_fuzzy_fn) if has_col1 else {}
                row_align = _pack_row_alignments(spans, has_col1, col1_final, col3_final)
            else:
                col3_final = (
                    _resolve_side_with_fallback(
                        text=col3_text,
                        raw_map=row_col3,
                        spans=spans,
                        local_fallback_fn=lambda: local_align_col3_spans(col3_text, old_en_full, spans, find_fuzzy_fn),
                        find_fuzzy_fn=find_fuzzy_fn,
                    )
                    if has_col3
                    else {}
                )
                col1_final = (
                    _resolve_side_with_fallback(
                        text=col1_text,
                        raw_map=row_col1,
                        spans=spans,
                        local_fallback_fn=lambda: local_align_col1_spans(col1_text, old_en_full, spans, find_fuzzy_fn),
                        find_fuzzy_fn=find_fuzzy_fn,
                    )
                    if has_col1
                    else {}
                )
                row_align = _pack_row_alignments(spans, has_col1, col1_final, col3_final)

        with state_lock:
            alignments.update(row_align)
            completed_rows += 1
            done_r = completed_rows

        mode_lbl = "本地位置估算" if fallback_event.is_set() else ("AI 中/外文对齐" if any_col1 else "AI 外文对齐")
        _safe_progress(done_r / total_rows, f"{mode_lbl}：已完成 {done_r}/{total_rows} 行")

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = [executor.submit(_align_single_row, r_i) for r_i in range(total_rows)]
        for fut in futures:
            fut.result()

    if stats_out is not None:
        stats_out.update(stats)
    return alignments
