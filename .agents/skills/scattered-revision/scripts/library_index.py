import hashlib
import json
import os
import re
import tempfile
import threading
import time
import zipfile
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Callable, Dict, List, Optional

from docx_xml_utils import (
    MAX_DOCX_MEMBER_BYTES,
    W,
    _assert_safe_openxml_bytes,
    _validate_docx_zip,
    load_config,
)

INDEX_SCHEMA_VERSION = 3
_FILES_KEY_MARKER = ',"files":{'

_FINAL_OUTPUT_NAME_RE = re.compile(
    r"_\((?:UP\s+\d+\s+words,\s*RV\s+\d+\s+words|up\s*\d+[\s,]+rv\s*\d+)\)_",
    re.IGNORECASE,
)
_WS_RE = re.compile(r"[\s\u3000\u200b\u200c\ufeff]+")
_LOOSE_KEEP_RE = re.compile(r"[\u4e00-\u9fff\u3400-\u4dbfA-Za-z0-9]+")
_CN_SENT_SPLIT_RE = re.compile(r"(?<=[。？！!?；;\n])(?:[”’\"'」』）\)]*)")

_CACHE_LOCKS_GUARD = threading.Lock()
_CACHE_LOCKS: Dict[str, threading.Lock] = {}
_MEM_INDEX_CACHE: Dict[str, tuple] = {}  # norm_lib_dir -> (stat_sig_tuple, files_dict)


def _get_lock_for_key(key: str) -> threading.Lock:
    with _CACHE_LOCKS_GUARD:
        lk = _CACHE_LOCKS.get(key)
        if lk is None:
            lk = threading.Lock()
            _CACHE_LOCKS[key] = lk
        return lk


def normalize_cn(s: str) -> str:
    return _WS_RE.sub("", s or "")


def normalize_cn_loose(s: str) -> str:
    """Keep only CJK ideographs and ASCII alphanumerics for punctuation-agnostic anchor/3-gram matching."""
    return "".join(_LOOSE_KEEP_RE.findall(s or "")).lower()


def split_cn_sentences(text: str) -> List[str]:
    """Split Chinese paragraph into normalized non-empty sentence units."""
    if not text:
        return []
    raw_parts = _CN_SENT_SPLIT_RE.split(text.strip())
    sents = []
    for part in raw_parts:
        ns = normalize_cn(part)
        if ns:
            sents.append(ns)
    return sents


def build_row_anchors(cn_loose: str, max_anchors: int = 12) -> List[str]:
    """Sample evenly spaced 4-char (or 3-char) lookup anchors across normalized Chinese text."""
    n = len(cn_loose)
    if n < 3:
        return [cn_loose] if cn_loose else []
    k = 4 if n >= 6 else 3
    if n <= k:
        return [cn_loose]
    span = n - k
    count = min(max_anchors, span + 1)
    if count <= 1:
        return [cn_loose[:k]]
    anchors = []
    seen = set()
    for i in range(count):
        pos = round(i * span / (count - 1))
        gram = cn_loose[pos : pos + k]
        if gram not in seen:
            seen.add(gram)
            anchors.append(gram)
    return anchors


def _resolve_cache_path(resolved_lib_dir: Path) -> Path:
    norm_key = os.path.normcase(str(resolved_lib_dir))
    sha1_hex = hashlib.sha1(norm_key.encode("utf-8")).hexdigest()
    fname = f"lib_index_{sha1_hex}.json"

    candidates = []
    env_override = os.environ.get("SCATTERED_REV_CACHE_DIR", "").strip()
    if env_override:
        candidates.append(Path(env_override))
    local_app = os.environ.get("LOCALAPPDATA", "").strip()
    if local_app:
        candidates.append(Path(local_app) / "ScatteredRevisionTool")
    candidates.append(Path(tempfile.gettempdir()) / "ScatteredRevisionTool")

    for base_dir in candidates:
        try:
            base_dir.mkdir(parents=True, exist_ok=True)
            return base_dir / fname
        except OSError:
            continue
    return Path(tempfile.gettempdir()) / fname


def is_excluded_library_file(rel_path: Path, exclude_dir_keywords: List[str]) -> bool:
    """
    Determine whether `rel_path` (relative to `lib_dir`) must be excluded from the library index:
      - `~$` Office lock/temp files or non-.docx files
      - Relative parent directories matching `exclude_library_dir_keywords`
      - Filenames starting with `初筛_`, containing `标记结果`, or matching `_(UP X words, RV Y words)_`
    """
    name = rel_path.name
    if not name.lower().endswith(".docx"):
        return True
    if name.startswith("~$") or name.startswith("."):
        return True
    if name.startswith("初筛_") or "标记结果" in name:
        return True
    if _FINAL_OUTPUT_NAME_RE.search(name):
        return True

    kw_lows = [kw.lower() for kw in exclude_dir_keywords if kw]
    for dir_part in rel_path.parent.parts:
        dp_low = dir_part.lower()
        if any(kw in dp_low for kw in kw_lows):
            return True
    return False


def _scan_library_docx_stats(resolved_lib_dir: Path, exclude_dir_keywords: List[str]) -> Dict[str, tuple]:
    """
    Fast recursive os.scandir walk returning {posix_rel_path: (abs_path_str, mtime_ns, size)}.
    On Windows, DirEntry.stat() uses cached FindFirstFileW metadata without extra syscalls.
    """
    kw_lows = [kw.lower() for kw in exclude_dir_keywords if kw]
    found: Dict[str, tuple] = {}
    root_str = str(resolved_lib_dir)

    def _walk(cur_dir: str, rel_prefix: str):
        try:
            with os.scandir(cur_dir) as it:
                entries = list(it)
        except OSError:
            return
        for entry in entries:
            ename = entry.name
            if entry.is_dir(follow_symlinks=False):
                dp_low = ename.lower()
                if ename.startswith(".") or any(kw in dp_low for kw in kw_lows):
                    continue
                next_rel = f"{rel_prefix}/{ename}" if rel_prefix else ename
                _walk(entry.path, next_rel)
            elif entry.is_file(follow_symlinks=False):
                if not ename.lower().endswith(".docx") or ename.startswith("~$") or ename.startswith("."):
                    continue
                if ename.startswith("初筛_") or "标记结果" in ename or _FINAL_OUTPUT_NAME_RE.search(ename):
                    continue
                rel_posix = f"{rel_prefix}/{ename}" if rel_prefix else ename
                try:
                    st = entry.stat(follow_symlinks=False)
                except OSError:
                    continue
                found[rel_posix] = (entry.path, st.st_mtime_ns, st.st_size)

    _walk(root_str, "")
    return found


def _cell_plain_text_fast(tc: ET.Element) -> str:
    t_tag = f"{W}t"
    p_tag = f"{W}p"
    return "\n".join(
        "".join(t.text or "" for t in p.iter(t_tag))
        for p in tc.findall(p_tag)
    )


def _parse_single_library_docx(task: tuple) -> tuple:
    """Worker function for ThreadPoolExecutor: parse one .docx into file index record."""
    rel_posix, abs_path_str, mtime_ns, size = task
    stem = Path(rel_posix).stem.strip()
    p_tag = f"{W}p"
    tbl_tag = f"{W}tbl"
    tr_tag = f"{W}tr"
    tc_tag = f"{W}tc"
    t_tag = f"{W}t"
    body_tag = f"{W}body"

    try:
        with zipfile.ZipFile(abs_path_str, "r") as zf:
            _validate_docx_zip(zf)
            if "word/document.xml" not in zf.namelist():
                return rel_posix, None, "missing word/document.xml"
            info = zf.getinfo("word/document.xml")
            with zf.open(info, "r") as fp:
                xml_bytes = fp.read(MAX_DOCX_MEMBER_BYTES + 1)
        _assert_safe_openxml_bytes(xml_bytes)
        root = ET.fromstring(xml_bytes)
    except Exception as ex:
        return rel_posix, {
            "rel_path": rel_posix,
            "mtime_ns": mtime_ns,
            "size": size,
            "titles": [stem] if stem else [],
            "rows": [],
            "error": str(ex).splitlines()[0][:160],
        }, str(ex)

    body = root.find(body_tag)
    file_titles = []
    if stem:
        file_titles.append(stem)

    tr_to_header: Dict[int, List[str]] = {}
    if body is not None:
        pending_headers: List[str] = []
        for ch in body:
            if ch.tag == p_tag:
                ptxt = "".join(t.text or "" for t in ch.iter(t_tag)).strip()
                if ptxt and len(ptxt) <= 240:
                    pending_headers.append(ptxt)
                    if ptxt not in file_titles and len(file_titles) < 12:
                        file_titles.append(ptxt)
            elif ch.tag == tbl_tag:
                hdr_snapshot = pending_headers[-3:] if pending_headers else []
                for tr in ch.findall(tr_tag):
                    tr_to_header[id(tr)] = hdr_snapshot
                pending_headers = []

    all_trs = root.findall(f".//{tbl_tag}/{tr_tag}")
    rows_out = []
    for r_idx, tr in enumerate(all_trs):
        tcs = tr.findall(tc_tag)
        if len(tcs) < 3:
            continue
        col3_txt = _cell_plain_text_fast(tcs[2])
        if not col3_txt.strip():
            continue
        cn_raw = _cell_plain_text_fast(tcs[0])
        cn_norm = normalize_cn(cn_raw)
        if not cn_norm:
            continue
        en_raw = _cell_plain_text_fast(tcs[1])
        sents = split_cn_sentences(cn_raw)
        open_sents = sents[:3] if sents else [cn_norm[:80]]
        close_sents = sents[-3:] if sents else [cn_norm[-80:]]
        cn_loose = normalize_cn_loose(cn_raw)
        anchors = build_row_anchors(cn_loose)

        row_titles = [stem] if stem else []
        for hp in tr_to_header.get(id(tr), ()):
            if hp not in row_titles:
                row_titles.append(hp)

        rows_out.append({
            "rel_path": rel_posix,
            "title_candidates": row_titles,
            "row_idx": r_idx,
            "cn_norm": cn_norm,
            "opening_sentences": open_sents,
            "closing_sentences": close_sents,
            "anchors": anchors,
            "en_plain": en_raw,
            "col3_preview": col3_txt[:120],
            "has_col3": True,
        })

    return rel_posix, {
        "rel_path": rel_posix,
        "mtime_ns": mtime_ns,
        "size": size,
        "titles": file_titles,
        "rows": rows_out,
    }, None


class LibraryIndex(dict):
    """
    Rich in-memory view over the persistent incremental library index.
    Supports both dict-style access (`idx["total_rows"]`, `idx["rows"]`, `idx["files"]`)
    and fast lazy lookup structures for 5-stage cascading search.
    """

    def __init__(self, lib_dir: Path, cache_path: Path, files_map: Dict[str, dict], rebuilt_files: int, elapsed_sec: float):
        all_rows: List[dict] = []
        rows_by_rel: Dict[str, List[dict]] = {}
        for rel_p in sorted(files_map.keys()):
            f_entry = files_map[rel_p]
            f_rows = f_entry.get("rows", [])
            rows_by_rel[rel_p] = f_rows
            all_rows.extend(f_rows)

        super().__init__(
            version=INDEX_SCHEMA_VERSION,
            lib_dir=str(lib_dir),
            cache_path=str(cache_path),
            total_files=len(files_map),
            total_rows=len(all_rows),
            rebuilt_files=rebuilt_files,
            elapsed_sec=round(elapsed_sec, 4),
            files=files_map,
            rows=all_rows,
        )
        self.lib_dir = lib_dir
        self.cache_path = cache_path
        self.files_map = files_map
        self.rows: List[dict] = all_rows
        self.rows_by_rel: Dict[str, List[dict]] = rows_by_rel

        self._exact_cn_map: Optional[Dict[str, List[dict]]] = None
        self._loose_list: Optional[List[str]] = None
        self._trigram_postings: Optional[Dict[str, List[int]]] = None

    @property
    def exact_cn_map(self) -> Dict[str, List[dict]]:
        if self._exact_cn_map is None:
            mp: Dict[str, List[dict]] = {}
            for r in self.rows:
                mp.setdefault(r["cn_norm"], []).append(r)
            self._exact_cn_map = mp
        return self._exact_cn_map

    @property
    def loose_cn_list(self) -> List[str]:
        if self._loose_list is None:
            self._loose_list = [normalize_cn_loose(r["cn_norm"]) for r in self.rows]
        return self._loose_list

    @property
    def trigram_postings(self) -> Dict[str, List[int]]:
        """Lazily constructed stride-2 3-gram inverted index over all library rows."""
        if self._trigram_postings is None:
            postings: Dict[str, List[int]] = {}
            loose_all = self.loose_cn_list
            for idx, s in enumerate(loose_all):
                n = len(s)
                if n < 3:
                    continue
                seen_local = set()
                positions = list(range(0, n - 2, 2))
                if (n - 3) not in positions:
                    positions.append(n - 3)
                for p in positions:
                    gram = s[p : p + 3]
                    if gram not in seen_local:
                        seen_local.add(gram)
                        postings.setdefault(gram, []).append(idx)
            self._trigram_postings = postings
        return self._trigram_postings

    def resolve_file_path(self, rel_path: str) -> Path:
        return self.lib_dir / Path(rel_path)

    def retrieve_trigram_candidates(self, query_cn: str, top_k: int = 25, allowed_rels: Optional[set] = None) -> List[dict]:
        """Retrieve top candidate rows sharing 3-grams with `query_cn`."""
        q_loose = normalize_cn_loose(query_cn)
        if len(q_loose) < 3:
            return []
        q_grams = {q_loose[i : i + 3] for i in range(len(q_loose) - 2)}
        if not q_grams:
            return []

        if allowed_rels is not None:
            scored = []
            for rel_p in allowed_rels:
                for r in self.rows_by_rel.get(rel_p, ()):
                    rl = normalize_cn_loose(r["cn_norm"])
                    hits = sum(1 for g in q_grams if g in rl)
                    if hits > 0:
                        scored.append((hits, r))
            scored.sort(key=lambda x: x[0], reverse=True)
            return [r for _, r in scored[:top_k]]

        postings = self.trigram_postings
        counts: Dict[int, int] = {}
        for g in q_grams:
            plist = postings.get(g)
            if plist:
                for r_i in plist:
                    counts[r_i] = counts.get(r_i, 0) + 1
        if not counts:
            return []
        ranked = sorted(counts.items(), key=lambda kv: kv[1], reverse=True)[:top_k]
        return [self.rows[r_i] for r_i, _ in ranked]


def _try_fast_header_check(cache_path: Path, disk_stats: Dict[str, tuple]) -> Optional[dict]:
    """
    Read only the compact manifest header at the start of `cache_path` (~60 KB for 650 files)
    to verify 100% cache freshness in ~10ms without parsing multi-megabyte row payloads.
    """
    if not cache_path.exists():
        return None
    max_hdr_bytes = max(65536, len(disk_stats) * 320 + 4096)
    try:
        with open(cache_path, "rb") as fp:
            chunk = fp.read(max_hdr_bytes).decode("utf-8", errors="ignore")
        marker_pos = chunk.find(_FILES_KEY_MARKER)
        if marker_pos == -1:
            return None
        hdr_obj = json.loads(chunk[:marker_pos] + "}")
        if hdr_obj.get("version") != INDEX_SCHEMA_VERSION:
            return None
        manifest = hdr_obj.get("manifest")
        if not isinstance(manifest, dict) or len(manifest) != len(disk_stats):
            return None
        for rel_p, (_, mt, sz) in disk_stats.items():
            m_entry = manifest.get(rel_p)
            if not m_entry or m_entry[0] != mt or m_entry[1] != sz:
                return None
        return {
            "total_files": int(hdr_obj.get("total_files", len(disk_stats))),
            "total_rows": int(hdr_obj.get("total_rows", 0)),
            "rebuilt_files": 0,
            "cache_path": str(cache_path),
        }
    except Exception:
        return None


def _build_or_load_index(
    lib_dir,
    log_fn: Optional[Callable[[str], None]] = None,
    max_workers: int = 8,
) -> LibraryIndex:
    t0 = time.perf_counter()
    resolved_lib = Path(lib_dir).resolve()
    if not resolved_lib.exists() or not resolved_lib.is_dir():
        raise FileNotFoundError(f"找不到语种总库文件夹：{resolved_lib}")

    norm_lib_key = os.path.normcase(str(resolved_lib))
    lock = _get_lock_for_key(norm_lib_key)

    with lock:
        cfg = load_config()
        exclude_kws = cfg.get("exclude_library_dir_keywords", ["标记结果", "初筛", "答案", "output", "temp"])
        cache_path = _resolve_cache_path(resolved_lib)

        disk_stats = _scan_library_docx_stats(resolved_lib, exclude_kws)
        stat_sig = tuple((rel_p, mt, sz) for rel_p, (_, mt, sz) in sorted(disk_stats.items()))

        # Fast path 1: In-process memory cache when disk stats & cache file existence match
        mem_hit = _MEM_INDEX_CACHE.get(norm_lib_key)
        if mem_hit is not None and mem_hit[0] == stat_sig and cache_path.exists():
            elapsed = time.perf_counter() - t0
            return LibraryIndex(resolved_lib, cache_path, mem_hit[1], rebuilt_files=0, elapsed_sec=elapsed)

        # Load persistent JSON cache if present
        cached_files: Dict[str, dict] = {}
        cache_loaded_ok = False
        if cache_path.exists():
            try:
                with open(cache_path, "r", encoding="utf-8") as f:
                    raw_cache = json.load(f)
                if (
                    isinstance(raw_cache, dict)
                    and raw_cache.get("version") == INDEX_SCHEMA_VERSION
                    and isinstance(raw_cache.get("files"), dict)
                ):
                    cached_files = raw_cache["files"]
                    cache_loaded_ok = True
            except Exception:
                cached_files = {}
                cache_loaded_ok = False

        next_files: Dict[str, dict] = {}
        to_parse = []
        for rel_p, (abs_p, mt, sz) in disk_stats.items():
            prev = cached_files.get(rel_p)
            if (
                prev is not None
                and prev.get("mtime_ns") == mt
                and prev.get("size") == sz
                and isinstance(prev.get("rows"), list)
            ):
                next_files[rel_p] = prev
            else:
                to_parse.append((rel_p, abs_p, mt, sz))

        deleted_count = sum(1 for old_k in cached_files if old_k not in disk_stats)
        rebuilt_count = len(to_parse)

        if to_parse:
            workers = max(1, min(int(max_workers or 8), len(to_parse)))
            if log_fn and len(to_parse) >= 5:
                log_fn(f"📚 正在多线程建立/更新总库索引（需解析 {len(to_parse)}/{len(disk_stats)} 个文档，线程数={workers}）...")
            if workers == 1:
                results = [_parse_single_library_docx(t) for t in to_parse]
            else:
                with ThreadPoolExecutor(max_workers=workers) as ex:
                    results = list(ex.map(_parse_single_library_docx, to_parse))

            for rel_p, record, err_msg in results:
                if record is not None:
                    next_files[rel_p] = record
                if err_msg and log_fn:
                    log_fn(f"  ⚠️ 跳过无法读取的总库文件 {rel_p}: {err_msg.splitlines()[0][:100]}")

        need_save = (not cache_loaded_ok) or (rebuilt_count > 0) or (deleted_count > 0)
        if need_save:
            total_rows = sum(len(v.get("rows", ())) for v in next_files.values())
            manifest_hdr = {
                rel_p: [v.get("mtime_ns", 0), v.get("size", 0), len(v.get("rows", ()))]
                for rel_p, v in next_files.items()
            }
            payload = {
                "version": INDEX_SCHEMA_VERSION,
                "lib_dir": str(resolved_lib),
                "updated_at": time.time(),
                "total_files": len(next_files),
                "total_rows": total_rows,
                "manifest": manifest_hdr,
                "files": next_files,
            }
            try:
                cache_path.parent.mkdir(parents=True, exist_ok=True)
                tmp_path = cache_path.with_name(f"{cache_path.name}.tmp_{os.getpid()}_{threading.get_ident()}")
                with open(tmp_path, "w", encoding="utf-8") as f:
                    json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))
                os.replace(tmp_path, cache_path)
            except OSError as ex:
                if log_fn:
                    log_fn(f"  ⚠️ 总库索引缓存写入失败（不影响本次运行）: {ex}")

        _MEM_INDEX_CACHE[norm_lib_key] = (stat_sig, next_files)
        elapsed = time.perf_counter() - t0
        return LibraryIndex(
            resolved_lib,
            cache_path,
            next_files,
            rebuilt_files=rebuilt_count,
            elapsed_sec=elapsed,
        )


def get_library_index(
    lib_dir,
    log_fn: Optional[Callable[[str], None]] = None,
    max_workers: int = 8,
) -> LibraryIndex:
    """Load or incrementally update the persistent library index and return a `LibraryIndex` instance."""
    return _build_or_load_index(lib_dir, log_fn=log_fn, max_workers=max_workers)


def ensure_library_index(
    lib_dir,
    log_fn: Optional[Callable[[str], None]] = None,
    max_workers: int = 8,
) -> dict:
    """
    Ensure the persistent incremental library index for `lib_dir` is up-to-date.
    Returns summary dict:
      {"total_files": int, "total_rows": int, "rebuilt_files": int, "cache_path": str, "elapsed_sec": float}
    """
    t0 = time.perf_counter()
    resolved_lib = Path(lib_dir).resolve()
    if not resolved_lib.exists() or not resolved_lib.is_dir():
        raise FileNotFoundError(f"找不到语种总库文件夹：{resolved_lib}")

    norm_lib_key = os.path.normcase(str(resolved_lib))
    lock = _get_lock_for_key(norm_lib_key)
    with lock:
        cfg = load_config()
        exclude_kws = cfg.get("exclude_library_dir_keywords", ["标记结果", "初筛", "答案", "output", "temp"])
        cache_path = _resolve_cache_path(resolved_lib)
        disk_stats = _scan_library_docx_stats(resolved_lib, exclude_kws)
        stat_sig = tuple((rel_p, mt, sz) for rel_p, (_, mt, sz) in sorted(disk_stats.items()))

        mem_hit = _MEM_INDEX_CACHE.get(norm_lib_key)
        if mem_hit is not None and mem_hit[0] == stat_sig and cache_path.exists():
            tot_rows = sum(len(v.get("rows", ())) for v in mem_hit[1].values())
            return {
                "total_files": len(mem_hit[1]),
                "total_rows": tot_rows,
                "rebuilt_files": 0,
                "cache_path": str(cache_path),
                "elapsed_sec": round(time.perf_counter() - t0, 4),
            }

        fast_res = _try_fast_header_check(cache_path, disk_stats)
        if fast_res is not None:
            fast_res["elapsed_sec"] = round(time.perf_counter() - t0, 4)
            return fast_res

    idx = _build_or_load_index(lib_dir, log_fn=log_fn, max_workers=max_workers)
    return {
        "total_files": int(idx["total_files"]),
        "total_rows": int(idx["total_rows"]),
        "rebuilt_files": int(idx["rebuilt_files"]),
        "cache_path": str(idx["cache_path"]),
        "elapsed_sec": round(time.perf_counter() - t0, 4),
    }
