import copy
import json
import re
import ssl
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from typing import Callable, Dict, List, Optional, Union

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


def _resolve_backend(api_key: str, api_mode: str = "auto") -> str:
    k = (api_key or "").strip()
    if api_mode in ("vertex", "gemini"):
        return api_mode
    if k.startswith("AIza"):
        return "gemini"
    if k.startswith("AQ."):
        return "vertex"
    return "vertex"


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

    def __len__(self) -> int:
        return len(self.keys)

    @property
    def all_keys(self) -> List[str]:
        return list(self.keys)

    def get_candidate_order(self) -> List[int]:
        """Return key indices ordered for the next request: healthy keys first, cooling-down keys as fallback."""
        with self._lock:
            n = len(self.keys)
            now = time.monotonic()
            start = self._cursor % n if self.rotation_mode == "round_robin" else (self._cursor % n)
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
    primary = _resolve_backend(key, api_mode)
    backends = [primary]
    if api_mode == "auto":
        backends.append("gemini" if primary == "vertex" else "vertex")

    last_code: Optional[int] = None
    last_err: Optional[Exception] = None

    for backend in backends:
        adapted_schema = _adapt_schema_for_backend(json_schema, backend)
        try:
            if backend == "vertex":
                url = f"https://aiplatform.googleapis.com/v1beta1/publishers/google/models/{quoted_model}:generateContent"
                headers = {
                    "Content-Type": "application/json; charset=utf-8",
                    "x-goog-api-key": key,
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
                    "x-goog-api-key": key,
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
                body_txt = he.read(2048).decode("utf-8", errors="ignore")[:280]
            except Exception:
                pass
            safe_msg = redact_secrets(body_txt or str(he.reason), extra_secret=key)
            last_code = he.code
            last_err = RuntimeError(f"HTTP {he.code} ({backend}/{safe_model}): {safe_msg}")
            # Try alternate backend in auto mode on 400/401/403/404 (e.g. Gemini key tested against Vertex first)
            if he.code in (400, 401, 403, 404) and api_mode == "auto" and len(backends) > 1:
                continue
            break
        except Exception as ex:
            last_code = None
            last_err = RuntimeError(redact_secrets(str(ex), extra_secret=key))
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
    If a key hits rate limits (429), 503 high demand, or errors out, switches keys or backs off cleanly.
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
                    clean_alt = redact_secrets(str(alt_ex), extra_secret=pool.all_keys)
                    last_err = RuntimeError(f"{mask_key_label(k_idx, k_val)} -> {clean_alt}")

    if total_keys > 1:
        raise RuntimeError(f"密钥池中全部 {total_keys} 个 API Key 均调用失败！最后错误：{last_err}")
    raise last_err or RuntimeError("调用 Gemini/Vertex API 失败。")


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
            item = {
                "index": idx,
                "key": k,
                "masked": masked,
                "ok": False,
                "channel": "-",
                "detail": redact_secrets(str(ex), extra_secret=keys)[:160],
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
            raise RuntimeError(rep["detail"])
        return f"连接成功！{rep['masked']} | 通道：{rep['channel']} | 模型：{model}"

    def _each(done_i: int, tot: int, item: dict):
        if log_fn:
            icon = "✅" if item["ok"] else "❌"
            log_fn(f"  [{done_i}/{tot}] {icon} {item['masked']}：{item['detail']}")

    reports = test_all_api_keys(keys, model=model, api_mode=api_mode, on_each_cb=_each)
    ok_list = [r for r in reports if r["ok"]]
    bad_list = [r for r in reports if not r["ok"]]
    if not ok_list:
        raise RuntimeError(f"密钥池共 {len(reports)} 个 Key 全部不可用！首条错误：{bad_list[0]['detail']}")
    return (
        f"密钥池测试通过！共 {len(reports)} 个 Key：✅ 可用 {len(ok_list)} 个，"
        f"⚠️ 不可用 {len(bad_list)} 个（执行时将自动跳过报错 Key 并轮换可用 Key） | 模型：{model}"
    )


STEP1_SYS_PROMPT = """你是专业的多语种文稿修订分级专家。
对于给定的每个英文修订句（包含 sent_id、change_ratio、edit_hunks、diff_preview），请判定它属于：
- "yellow"（UP 局部改点）：局部词汇替换、短语微调、增删少量词、标点/引号微调，译者可在外文原句上直接局部修改；
- "cyan"（RV 整句重翻/大改）：整句语序重组、主干句式推倒重写、或改动极其密集导致译者难以在原句局部修补（通常 change_ratio >= 0.40 且多处打散重写，或相邻连续多句大改）。
注意：为了节约翻译成本，凡是译者能在原句顺畅局部修改的，一律优先判为 "yellow"；只有确实整句结构大改、难以局部下手的才判为 "cyan"。
请严格输出 JSON：{"items": [{"sent_id": "...", "decision": "yellow" 或 "cyan"}]}。"""


def evaluate_step1_decisions(
    review_items: List[dict],
    api_key: Union[str, List[str], ApiKeyPool],
    model: str,
    api_mode: str = "auto",
    rotation_mode: str = "round_robin",
    log_fn: Optional[Callable[[str], None]] = None,
    progress_fn: Optional[Callable[[float, str], None]] = None,
) -> Dict[str, str]:
    edited = [it for it in review_items if it.get("has_edit")]
    decisions: Dict[str, str] = {}
    if not edited:
        if progress_fn:
            progress_fn(1.0, "无需 AI 评估（0 处改动句）")
        return decisions

    pool = api_key if isinstance(api_key, ApiKeyPool) else ApiKeyPool(api_key, rotation_mode=rotation_mode)
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
    for b_idx in range(total_batches):
        batch = edited[b_idx * batch_size : (b_idx + 1) * batch_size]
        stage_msg = f"AI 初筛评估：第 {b_idx + 1}/{total_batches} 批（本批 {len(batch)} 句 / 共 {len(edited)} 句）"
        if progress_fn:
            progress_fn(b_idx / total_batches, stage_msg)
        if log_fn:
            log_fn(f"  [{stage_msg}] 正在请求 {model}（密钥池共 {len(pool)} 个 Key）...")

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
                log_fn=log_fn,
            )
            for entry in res.get("items", []):
                sid = entry.get("sent_id")
                dec = entry.get("decision")
                if sid and dec in ("yellow", "cyan"):
                    decisions[sid] = dec
        except Exception as ex:
            if log_fn:
                log_fn(f"  [提示] AI 评估第 {b_idx + 1} 批所有密钥均异常（已自动回退内置规则）：{ex}")

        for it in batch:
            if it["sent_id"] not in decisions:
                decisions[it["sent_id"]] = it.get("suggested", "yellow")
        if progress_fn:
            progress_fn((b_idx + 1) / total_batches, f"AI 初筛评估：已完成 {b_idx + 1}/{total_batches} 批")

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
    api_key: Union[str, List[str], ApiKeyPool],
    model: str,
    api_mode: str = "auto",
    rotation_mode: str = "round_robin",
    log_fn: Optional[Callable[[str], None]] = None,
    progress_fn: Optional[Callable[[float, str], None]] = None,
    target_lang_hint: str = "",
) -> Dict[str, str]:
    row_tasks = tasks_data.get("row_tasks", [])
    alignments: Dict[str, str] = {}
    if not row_tasks:
        if progress_fn:
            progress_fn(1.0, "无需对齐外文（0 处高亮）")
        return alignments

    pool = api_key if isinstance(api_key, ApiKeyPool) else ApiKeyPool(api_key, rotation_mode=rotation_mode)
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

    total_rows = len(row_tasks)
    for r_i, rt in enumerate(row_tasks):
        col3_text = rt["col3_text"]
        spans = rt["spans"]
        art_short = rt.get("article_title", "")[:18]
        stage_msg = f"AI 外文对齐：第 {r_i + 1}/{total_rows} 行（《{art_short}》，{len(spans)} 处高亮）"
        if progress_fn:
            progress_fn(r_i / total_rows, stage_msg)
        if log_fn:
            log_fn(f"  [{stage_msg}]...")

        pending_spans = list(spans)
        feedback_msg = ""
        for attempt in range(3):
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

            res = call_gemini_json(
                api_key=pool,
                model=model,
                system_instruction=STEP2_SYS_PROMPT,
                user_prompt=user_prompt,
                json_schema=schema,
                api_mode=api_mode,
                log_fn=log_fn,
            )
            returned = {
                item["span_id"]: item["col3_substring"]
                for item in res.get("alignments", [])
                if "span_id" in item and "col3_substring" in item
            }

            for s in pending_spans:
                sid = s["span_id"]
                if sid in returned and returned[sid].strip():
                    alignments[sid] = returned[sid].strip()

            cursor = 0
            failed = []
            for s in spans:
                sid = s["span_id"]
                cand = alignments.get(sid, "")
                loc = find_fuzzy_fn(col3_text, cand, search_from=cursor) if cand else None
                if loc is None:
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
            if log_fn and attempt + 1 < 3:
                log_fn(f"    ↻ 自动修正重试 {len(failed)} 处子串边界（第 {attempt + 2} 次）...")

        # Safety anchor fallback if any span still failed after 3 attempts
        cursor = 0
        for s in spans:
            sid = s["span_id"]
            cand = alignments.get(sid, "")
            loc = find_fuzzy_fn(col3_text, cand, search_from=cursor) if cand else None
            if loc is not None:
                cursor = loc[1]
                continue
            words = cand.split()
            if len(words) >= 4:
                for n_anchor in (3, 2, 1):
                    head = " ".join(words[:n_anchor])
                    tail = " ".join(words[-n_anchor:])
                    h_loc = find_fuzzy_fn(col3_text, head, search_from=cursor)
                    if h_loc is not None:
                        t_loc = find_fuzzy_fn(col3_text, tail, search_from=h_loc[1])
                        if t_loc is not None and (t_loc[1] - h_loc[0]) <= max(len(cand) * 2 + 120, len(col3_text)):
                            alignments[sid] = col3_text[h_loc[0]:t_loc[1]]
                            cursor = t_loc[1]
                            break
            elif len(cand) >= 12:
                for n_ch in (6, 4, 3):
                    head = cand[:n_ch]
                    tail = cand[-n_ch:]
                    h_loc = find_fuzzy_fn(col3_text, head, search_from=cursor)
                    if h_loc is not None:
                        t_loc = find_fuzzy_fn(col3_text, tail, search_from=h_loc[1])
                        if t_loc is not None and (t_loc[1] - h_loc[0]) <= max(len(cand) * 2 + 80, len(col3_text)):
                            alignments[sid] = col3_text[h_loc[0]:t_loc[1]]
                            cursor = t_loc[1]
                            break

        if progress_fn:
            progress_fn((r_i + 1) / total_rows, f"AI 外文对齐：已完成 {r_i + 1}/{total_rows} 行")

    return alignments
