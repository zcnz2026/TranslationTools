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
    ("Gemini 3.5 Flash Lite (推荐·均衡快速)", "gemini-3.5-flash-lite"),
    ("Gemini 3.1 Flash Lite (极速·超低成本)", "gemini-3.1-flash-lite"),
]

API_MODES = [
    ("自动识别 (支持混填 Vertex Key 与 Gemini Key)", "auto"),
    ("Vertex AI API Key", "vertex"),
    ("Gemini API Key (AI Studio)", "gemini"),
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


def _call_single_key_once(
    key: str,
    safe_model: str,
    quoted_model: str,
    system_instruction: str,
    user_prompt: str,
    json_schema: Optional[dict],
    api_mode: str,
) -> tuple[dict, str]:
    """Attempt one logical call with a single key (trying alternate backend only when api_mode=='auto')."""
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
                url = f"https://aiplatform.googleapis.com/v1beta1/publishers/google/models/{quoted_model}:generateContent"
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
                res = _http_post_json(url, headers, payload)
                cands = res.get("candidates") or []
                if not cands:
                    block_reason = res.get("promptFeedback", {}).get("blockReason", "EMPTY_CANDIDATES")
                    raise RuntimeError(f"Vertex AI 未返回有效候选结果 ({block_reason})。")
                parts = cands[0].get("content", {}).get("parts", [])
                text = "".join(p.get("text", "") for p in parts if "text" in p).strip()
            else:
                url = "https://generativelanguage.googleapis.com/v1beta/interactions"
                headers = {
                    "Content-Type": "application/json; charset=utf-8",
                    "x-goog-api-key": clean_k,
                    "Api-Revision": "2026-05-20",
                }
                payload = {
                    "model": safe_model,
                    "store": False,
                    "system_instruction": system_instruction,
                    "input": user_prompt,
                    "generation_config": {"temperature": 0.1},
                }
                if adapted_schema:
                    payload["response_format"] = {
                        "type": "text",
                        "mime_type": "application/json",
                        "schema": adapted_schema,
                    }
                res = _http_post_json(url, headers, payload)
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

                # Standard AIza key rejected as outright invalid by Gemini need not probe Vertex (which returns false OAuth2 401)
                low_msg = safe_msg.lower()
                if (
                    api_mode == "auto"
                    and primary == "gemini"
                    and clean_k.startswith("AIza")
                    and ("api_key_invalid" in low_msg or "api key not valid" in low_msg)
                ):
                    break

                if he.code in (400, 401, 403, 404) and api_mode == "auto" and len(backends) > 1:
                    continue
            else:
                # Alternate backend failed: never overwrite primary's error with a misleading Vertex OAuth2/API-key auth mismatch,
                # and keep primary's error when primary had no auth-scheme mismatch and alternate also rejected auth (400/401/403/404).
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
            ch_label = "Vertex AI Express" if used_backend == "vertex" else "Gemini Developer API"
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

_CLOSING_TRAILERS = frozenset('”’"\'»›」』)）]】〕}')
_OPENING_LEADERS = frozenset('“‘"\'«‹「『(（[【〔{„‚')
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
        loc = find_fuzzy_fn(old_en_full, piece, search_from=en_cursor)
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

    # 2. Split `en_full` and `col3_text` into sentence ranges
    en_sents = _split_sentences_ranges(en_full, is_english=True, include_semi_colon=False)
    for_sents = _split_sentences_ranges(col3_text, is_english=False, include_semi_colon=False)

    # If primary sentence counts differ by >1, check whether splitting at `;:` aligns sentence boundaries better
    if abs(len(en_sents) - len(for_sents)) > 1:
        en_semi = _split_sentences_ranges(en_full, is_english=True, include_semi_colon=True)
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

            en_clauses = _split_clauses_ranges(en_full, en_sent_s, en_sent_e, is_english=True)
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
            loc = find_fuzzy_fn(col3_text, sub, search_from=cursor)
            if loc is not None and loc[0] >= cursor:
                result[sid] = col3_text[loc[0] : loc[1]]
                cursor = loc[1]
            else:
                direct_pos = col3_text.find(sub, cursor)
                if direct_pos != -1:
                    result[sid] = sub
                    cursor = direct_pos + len(sub)

    return result


# ---------------- Step 1 & Step 2 Multi-Threaded Evaluators with Seamless Local Fallback ----------------

STEP1_SYS_PROMPT = """你是专业的多语种文稿修订分级专家。
对于给定的每个英文修订句（包含 sent_id、change_ratio、edit_hunks、diff_preview），请判定它属于：
- "yellow"（UP 局部改点）：局部词汇替换、短语微调、增删少量词、标点/引号微调，译者可在外文原句上直接局部修改；
- "cyan"（RV 整句重翻/大改）：整句语序重组、主干句式推倒重写、或改动极其密集导致译者难以在原句局部修补（通常 change_ratio >= 0.40 且多处打散重写，或相邻连续多句大改）。
注意：为了节约翻译成本，凡是译者能在原句顺畅局部修改的，一律优先判为 "yellow"；只有确实整句结构大改、难以局部下手的才判为 "cyan"。
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


STEP2_SYS_PROMPT = """你是多语种平行语料精准片段对齐专家（支持全球任意语种，包括拉丁语系、斯拉夫语系、希腊语、阿拉伯/波斯等从右向左语系、日韩语、泰/缅/柬/老等东南亚无空格语系、印地/孟加拉等南亚语系）。
给定表格第 3 列完整外文段落 `col3_text`，以及若干第 2 列高亮片段 `spans`（每个片段含 `span_id`、对应的旧版英文 `old_en_in_col3`、新版英文 `new_en`、以及前后未高亮的英文上下文 `left_old_context` / `right_old_context`）：
请在 `col3_text` 中找出与每个 `old_en_in_col3` 语义完全对应的**原样连续外文子串**（`col3_substring`）。

严格规则：
1. `col3_substring` 必须**一字不差地原样复制**自 `col3_text`（包含外文原有的标点符号，绝不可自行翻译、改写、省略中间词或添加省略号 `...`）。
2. 若 `old_en_in_col3` 仅为 ~3 个词的短窗口（纯标点改点），在 `col3_text` 中同样只截取对应位置的约 3 个外文词（若外文为日语、泰语、缅甸语、高棉语等无空格分词语言，则截取对应位置约 4~8 个字符的短片段）。
3. 若 `old_en_in_col3` 是半句/分句（如到逗号、分号或连词处为止），`col3_substring` 也必须精准停在对应的外文分句/连词边界，绝不可把 `left_old_context` 或 `right_old_context` 对应的未高亮外文包进来！
4. 同一行内的多个 `spans` 按顺序先后出现，互不重叠。
请严格输出 JSON：{"alignments": [{"span_id": "...", "col3_substring": "..."}]}。"""


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
) -> Dict[str, str]:
    stats = {"used_local_fallback": False, "fallback_reason": ""}
    row_tasks = tasks_data.get("row_tasks", [])
    alignments: Dict[str, str] = {}
    if not row_tasks:
        if progress_fn:
            progress_fn(1.0, "无需对齐外文（0 处高亮）")
        if stats_out is not None:
            stats_out.update(stats)
        return alignments

    pool = _build_optional_pool(api_key, rotation_mode=rotation_mode) if enable_ai else None
    total_rows = len(row_tasks)

    if not enable_ai or pool is None:
        stats["used_local_fallback"] = True
        if enable_ai and pool is None:
            stats["fallback_reason"] = "未填写 API Key"
        if log_fn:
            log_fn(f"  ℹ️ 外文标色：使用「本地位置估算模式」对齐共 {total_rows} 行第 3 列高亮...")
        for r_i, rt in enumerate(row_tasks):
            row_map = local_align_col3_spans(
                col3_text=rt.get("col3_text", ""),
                old_en_full=rt.get("old_en_full", ""),
                spans=rt.get("spans", []),
                find_fuzzy_fn=find_fuzzy_fn,
            )
            alignments.update(row_map)
            if progress_fn:
                progress_fn((r_i + 1) / total_rows, f"本地外文位置估算：已完成 {r_i + 1}/{total_rows} 行")
        if stats_out is not None:
            stats_out.update(stats)
        return alignments

    schema = {
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
        _safe_log(f"  ⚡ 启动 {workers} 线程并发对齐 {total_rows} 行外文高亮（密钥池 {len(pool)} 个 Key）...")

    def _align_single_row(r_i: int) -> None:
        nonlocal completed_rows
        rt = row_tasks[r_i]
        col3_text = rt.get("col3_text", "")
        old_en_full = rt.get("old_en_full", "")
        spans = rt.get("spans", [])
        row_align: Dict[str, str] = {}

        if fallback_event.is_set():
            row_align = local_align_col3_spans(col3_text, old_en_full, spans, find_fuzzy_fn)
        else:
            art_short = (rt.get("article_title") or "")[:18]
            stage_msg = f"AI 外文对齐：第 {r_i + 1}/{total_rows} 行（《{art_short}》，{len(spans)} 处高亮）"
            _safe_call_log(f"  [{stage_msg}]...")

            pending_spans = list(spans)
            feedback_msg = ""
            switched_to_local = False

            for attempt in range(3):
                if fallback_event.is_set():
                    switched_to_local = True
                    break

                payload_obj = {
                    "target_language": target_lang_hint or "auto",
                    "col3_text": col3_text,
                    "spans": [
                        {
                            "span_id": s["span_id"],
                            "old_en_in_col3": s["old_en_in_col3"],
                            "new_en": s["new_en"],
                            "left_old_context": s.get("left_old_context", ""),
                            "right_old_context": s.get("right_old_context", ""),
                        }
                        for s in pending_spans
                    ],
                }
                user_prompt = json.dumps(payload_obj, ensure_ascii=False, indent=2)
                if feedback_msg:
                    user_prompt = feedback_msg + "\n\n请重新精确截取以下未匹配成功的片段：\n" + user_prompt

                try:
                    res = call_gemini_json(
                        api_key=pool,
                        model=model,
                        system_instruction=STEP2_SYS_PROMPT,
                        user_prompt=user_prompt,
                        json_schema=schema,
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
                        _safe_log(f"  ⚠️ {friendly_err} -> 已自动切换「本地位置估算模式」完成第 3 列标色")
                    switched_to_local = True
                    break

                returned = {
                    item["span_id"]: item["col3_substring"]
                    for item in res.get("alignments", [])
                    if "span_id" in item and "col3_substring" in item
                }
                for s in pending_spans:
                    sid = s["span_id"]
                    if sid in returned and str(returned[sid]).strip():
                        row_align[sid] = str(returned[sid]).strip()

                cursor = 0
                failed = []
                for s in spans:
                    sid = s["span_id"]
                    cand = row_align.get(sid, "")
                    loc = find_fuzzy_fn(col3_text, cand, search_from=cursor) if cand else None
                    if loc is None or loc[0] < cursor:
                        failed.append((s, cand))
                    else:
                        cursor = loc[1]

                if not failed:
                    break

                pending_spans = [f[0] for f in failed]
                err_lines = [
                    f"- {s['span_id']} 返回的子串 {cand!r} 在 col3_text 中未找到！请务必从 col3_text 原文一字不差复制完整连续片段（不要用省略号，不要改标点）。"
                    for s, cand in failed
                ]
                feedback_msg = "上一次返回的部分子串未能匹配 col3_text 原文：\n" + "\n".join(err_lines)
                if attempt + 1 < 3:
                    _safe_call_log(f"    ↻ 自动修正重试 {len(failed)} 处子串边界（第 {attempt + 2} 次）...")

            if switched_to_local:
                row_align = local_align_col3_spans(col3_text, old_en_full, spans, find_fuzzy_fn)
            else:
                # Safety anchor fallback + local_align_col3_spans fallback for any residual unmatched spans
                local_backup: Optional[Dict[str, str]] = None
                cursor = 0
                for s in spans:
                    sid = s["span_id"]
                    cand = row_align.get(sid, "")
                    loc = find_fuzzy_fn(col3_text, cand, search_from=cursor) if cand else None
                    if loc is not None and loc[0] >= cursor:
                        cursor = loc[1]
                        continue

                    anchored = False
                    words = cand.split()
                    if len(words) >= 4:
                        for n_anchor in (3, 2, 1):
                            head = " ".join(words[:n_anchor])
                            tail = " ".join(words[-n_anchor:])
                            h_loc = find_fuzzy_fn(col3_text, head, search_from=cursor)
                            if h_loc is not None and h_loc[0] >= cursor:
                                t_loc = find_fuzzy_fn(col3_text, tail, search_from=h_loc[1])
                                if t_loc is not None and t_loc[0] >= h_loc[1] and (t_loc[1] - h_loc[0]) <= max(len(cand) * 2 + 120, len(col3_text)):
                                    row_align[sid] = col3_text[h_loc[0] : t_loc[1]]
                                    cursor = t_loc[1]
                                    anchored = True
                                    break
                    elif len(cand) >= 12:
                        for n_ch in (6, 4, 3):
                            head = cand[:n_ch]
                            tail = cand[-n_ch:]
                            h_loc = find_fuzzy_fn(col3_text, head, search_from=cursor)
                            if h_loc is not None and h_loc[0] >= cursor:
                                t_loc = find_fuzzy_fn(col3_text, tail, search_from=h_loc[1])
                                if t_loc is not None and t_loc[0] >= h_loc[1] and (t_loc[1] - h_loc[0]) <= max(len(cand) * 2 + 80, len(col3_text)):
                                    row_align[sid] = col3_text[h_loc[0] : t_loc[1]]
                                    cursor = t_loc[1]
                                    anchored = True
                                    break

                    if not anchored:
                        if local_backup is None:
                            local_backup = local_align_col3_spans(col3_text, old_en_full, spans, find_fuzzy_fn)
                        fb_cand = local_backup.get(sid, "")
                        fb_loc = find_fuzzy_fn(col3_text, fb_cand, search_from=cursor) if fb_cand else None
                        if fb_loc is not None and fb_loc[0] >= cursor:
                            row_align[sid] = col3_text[fb_loc[0] : fb_loc[1]]
                            cursor = fb_loc[1]

        with state_lock:
            alignments.update(row_align)
            completed_rows += 1
            done_r = completed_rows

        mode_lbl = "本地位置估算" if fallback_event.is_set() else "AI 外文对齐"
        _safe_progress(done_r / total_rows, f"{mode_lbl}：已完成 {done_r}/{total_rows} 行")

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = [executor.submit(_align_single_row, r_i) for r_i in range(total_rows)]
        for fut in futures:
            fut.result()

    if stats_out is not None:
        stats_out.update(stats)
    return alignments
