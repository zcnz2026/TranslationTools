import base64
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import ssl
import subprocess
import sys
import tempfile
from typing import Callable, Optional, Tuple
import urllib.error
import urllib.parse
import urllib.request
import zipfile

try:
    from security_utils import create_secure_update_stagedir
except ImportError:
    create_secure_update_stagedir = None

APP_VERSION = "1.9.0"
GITHUB_REPO = "zcnz2026/TranslationTools"

ZIP_ASSET_NAME = "ScatteredRevisionTool-Windows-x64.zip"
SHA256_ASSET_NAME = "SHA256SUMS.txt"
EXE_ENTRY_NAME = "零散改点整理工具.exe"

MAX_RELEASE_JSON_BYTES = 2 * 1024 * 1024     # 2 MB upper bound for GitHub release metadata
MAX_UPDATE_ZIP_BYTES = 250 * 1024 * 1024     # 250 MB upper bound for portable ZIP download
MAX_EXTRACTED_EXE_BYTES = 300 * 1024 * 1024  # 300 MB upper bound for extracted EXE

_ALLOWED_EXACT_HOSTS = frozenset({
    "github.com",
    "api.github.com",
    "githubusercontent.com",
})
_ALLOWED_HOST_SUFFIXES = (
    ".github.com",
    ".githubusercontent.com",
)

_SEMVER_RE = re.compile(
    r"^\s*[vV]?(\d+)(?:\.(\d+))?(?:\.(\d+))?(?:\.(\d+))?(?:[-+]([0-9A-Za-z.-]+))?\s*$"
)
_SHA256_LINE_RE = re.compile(r"^([0-9a-fA-F]{64})\s+\*?(.+?)\s*$")


def validate_github_https_url(url: str) -> str:
    """Ensure update URLs strictly use HTTPS and belong to GitHub / GitHubUserContent domains."""
    clean = (url or "").strip()
    parsed = urllib.parse.urlsplit(clean)
    host = (parsed.hostname or "").lower().rstrip(".")
    if parsed.scheme != "https" or not host:
        raise ValueError(f"已拦截非安全的更新地址（仅允许 HTTPS GitHub 官方域名）：{url!r}")
    if host not in _ALLOWED_EXACT_HOSTS and not any(host.endswith(sfx) for sfx in _ALLOWED_HOST_SUFFIXES):
        raise ValueError(f"已拦截非受信任的更新域名：{host!r}")
    return clean


def normalize_version_str(ver: str) -> str:
    s = (ver or "").strip()
    if s.startswith(("v", "V")):
        s = s[1:].strip()
    return s or "0.0.0"


def parse_version_key(ver: str) -> Tuple[Tuple[int, int, int, int], int, str]:
    """
    Convert a semantic version tag (e.g. 'v1.7.0', '1.8.0-rc1') into a comparable tuple.
    Stable releases sort higher than pre-release suffixes of the same numeric version.
    """
    m = _SEMVER_RE.match(ver or "")
    if not m:
        nums = [int(x) for x in re.findall(r"\d+", ver or "")[:4]]
        while len(nums) < 4:
            nums.append(0)
        return (nums[0], nums[1], nums[2], nums[3]), 1, ""
    c1 = int(m.group(1) or 0)
    c2 = int(m.group(2) or 0)
    c3 = int(m.group(3) or 0)
    c4 = int(m.group(4) or 0)
    prerelease = (m.group(5) or "").strip()
    is_stable = 0 if prerelease else 1
    return (c1, c2, c3, c4), is_stable, prerelease.lower()


def is_newer_version(latest_version: str, current_version: str) -> bool:
    """Return True strictly when `latest_version` > `current_version`."""
    if not (latest_version or "").strip():
        return False
    return parse_version_key(latest_version) > parse_version_key(current_version)


def parse_sha256sums(sums_text: str, target_filename: str = ZIP_ASSET_NAME) -> str:
    """Extract lowercase 64-char SHA-256 digest for `target_filename` from `SHA256SUMS.txt` content."""
    if not sums_text:
        return ""
    want = Path(target_filename).name.lower()
    clean_text = sums_text.lstrip("\ufeff")
    for raw_line in clean_text.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        m = _SHA256_LINE_RE.match(line)
        if not m:
            continue
        digest, fname = m.group(1).lower(), m.group(2).strip()
        entry_name = Path(fname.replace("\\", "/")).name.lower()
        if entry_name == want:
            return digest
    return ""


def compute_file_sha256(file_path: Path) -> str:
    h = hashlib.sha256()
    with open(file_path, "rb") as fp:
        for chunk in iter(lambda: fp.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().lower()


class _SafeGitHubRedirectHandler(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        resolved_url = urllib.parse.urljoin(req.full_url, newurl)
        validate_github_https_url(resolved_url)
        return super().redirect_request(req, fp, code, msg, headers, resolved_url)


def _build_github_opener(use_system_proxy: bool = True) -> urllib.request.OpenerDirector:
    ctx = ssl.create_default_context()
    ctx.minimum_version = ssl.TLSVersion.TLSv1_2
    proxy_handler = urllib.request.ProxyHandler() if use_system_proxy else urllib.request.ProxyHandler({})
    return urllib.request.build_opener(
        proxy_handler,
        urllib.request.HTTPSHandler(context=ctx),
        _SafeGitHubRedirectHandler(),
    )


def _open_with_proxy_fallback(req: urllib.request.Request, timeout: float):
    validate_github_https_url(req.full_url)
    try:
        opener = _build_github_opener(use_system_proxy=True)
        return opener.open(req, timeout=timeout)
    except urllib.error.HTTPError:
        raise
    except urllib.error.URLError:
        opener_direct = _build_github_opener(use_system_proxy=False)
        return opener_direct.open(req, timeout=timeout)


def _user_agent_header(version_str: str = APP_VERSION) -> str:
    v = normalize_version_str(version_str) or APP_VERSION
    return f"ScatteredRevisionTool/{v}"


def check_for_updates(current_version: str = APP_VERSION, timeout: float = 8.0) -> dict:
    """
    Query GitHub Releases API for the latest published release of `GITHUB_REPO`.
    Returns:
      {
        "has_update": bool,
        "current_version": str,
        "latest_version": str,
        "release_notes": str,
        "html_url": str,
        "zip_url": str,
        "sha256_url": str,
      }
    """
    cur_clean = normalize_version_str(current_version or APP_VERSION)
    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
    validate_github_https_url(api_url)

    req = urllib.request.Request(
        api_url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": _user_agent_header(APP_VERSION),
        },
        method="GET",
    )

    try:
        with _open_with_proxy_fallback(req, timeout=timeout) as resp:
            raw_bytes = resp.read(MAX_RELEASE_JSON_BYTES + 1)
    except urllib.error.HTTPError as he:
        try:
            he.close()
        except Exception:
            pass
        raise RuntimeError(f"检查更新失败：GitHub 接口返回 HTTP {he.code}。") from he
    except Exception as ex:
        raise RuntimeError(f"检查更新失败：无法连接 GitHub 服务器（{str(ex).splitlines()[0]}）。") from ex

    if len(raw_bytes) > MAX_RELEASE_JSON_BYTES:
        raise RuntimeError("检查更新失败：版本元数据体积超出安全上限。")

    data = json.loads(raw_bytes.decode("utf-8"))
    tag_name = str(data.get("tag_name") or "").strip()
    latest_clean = normalize_version_str(tag_name) if tag_name else cur_clean
    release_notes = str(data.get("body") or "").strip()
    html_url = str(data.get("html_url") or f"https://github.com/{GITHUB_REPO}/releases").strip()
    if html_url:
        validate_github_https_url(html_url)

    zip_url = ""
    sha256_url = ""
    for asset in data.get("assets") or []:
        if not isinstance(asset, dict):
            continue
        aname = str(asset.get("name") or "").strip()
        dl_url = str(asset.get("browser_download_url") or "").strip()
        if not aname or not dl_url:
            continue
        try:
            validate_github_https_url(dl_url)
        except ValueError:
            continue
        if aname.lower() == ZIP_ASSET_NAME.lower():
            zip_url = dl_url
        elif aname.lower() == SHA256_ASSET_NAME.lower():
            sha256_url = dl_url

    has_upd = is_newer_version(latest_clean, cur_clean)
    return {
        "has_update": has_upd,
        "current_version": cur_clean,
        "latest_version": latest_clean,
        "release_notes": release_notes,
        "html_url": html_url,
        "zip_url": zip_url,
        "sha256_url": sha256_url,
    }


def _validate_and_extract_exe_from_zip(zip_path: Path, stage_dir: Path) -> Path:
    """
    Validate ZIP entries against Zip Slip & Zip Bomb attacks, and extract ONLY `零散改点整理工具.exe`
    directly into `stage_dir / EXE_ENTRY_NAME`.
    """
    target_info = None
    total_uncompressed = 0

    with zipfile.ZipFile(zip_path, "r") as zf:
        for info in zf.infolist():
            fname = info.filename
            if "\x00" in fname or fname.startswith(("/", "\\")) or ":" in fname:
                raise ValueError(f"更新压缩包含有不安全的路径条目：{fname!r}")
            parts = fname.replace("\\", "/").split("/")
            if ".." in parts:
                raise ValueError(f"检测到 Zip Slip 路径穿越风险，已拦截更新包条目：{fname!r}")
            if info.file_size < 0 or info.file_size > MAX_EXTRACTED_EXE_BYTES:
                raise ValueError(f"更新压缩包内文件体积异常：{fname!r}")
            total_uncompressed += info.file_size
            if total_uncompressed > MAX_EXTRACTED_EXE_BYTES * 2:
                raise ValueError("更新压缩包解压总大小超出安全上限。")

            if not info.is_dir() and parts[-1] == EXE_ENTRY_NAME:
                target_info = info

        if target_info is None:
            raise ValueError(f"更新包中未找到主程序文件：{EXE_ENTRY_NAME}")

        staged_exe = (stage_dir / EXE_ENTRY_NAME).resolve()
        if staged_exe.parent != stage_dir.resolve():
            raise ValueError("提取目标路径校验失败。")

        written = 0
        with zf.open(target_info, "r") as src_fp, open(staged_exe, "wb") as dst_fp:
            while True:
                chunk = src_fp.read(256 * 1024)
                if not chunk:
                    break
                written += len(chunk)
                if written > MAX_EXTRACTED_EXE_BYTES:
                    raise ValueError("主程序解压体积超出安全上限。")
                dst_fp.write(chunk)

    # Verify Windows PE MZ header
    with open(staged_exe, "rb") as f_chk:
        magic = f_chk.read(2)
    if magic != b"MZ":
        raise ValueError("提取的更新程序不是合法的 Windows 可执行文件 (PE MZ)。")

    return staged_exe


def _ps_quote(val: str) -> str:
    """Safely quote a literal string for PowerShell single-quoted literal syntax."""
    return "'" + (val or "").replace("'", "''") + "'"


def _write_and_launch_windows_updater_script(
    stage_dir: Path,
    staged_exe: Path,
    target_exe: Path,
    wait_pid: int,
) -> Path:
    """
    Write a UTF-8-BOM PowerShell script in `stage_dir` that:
      1. Waits for `wait_pid` to exit;
      2. Copies ONLY `staged_exe` over `target_exe` (leaving app_settings.json & languages.json untouched);
      3. Relaunches `target_exe`;
      4. Deletes temporary staging files.
    """
    script_path = stage_dir / "apply_update.ps1"
    target_dir = target_exe.parent

    ps_lines = [
        "$ErrorActionPreference = 'Stop'",
        f"$WaitPid = {int(wait_pid)}",
        f"$StagedExe = {_ps_quote(str(staged_exe))}",
        f"$TargetExe = {_ps_quote(str(target_exe))}",
        f"$TargetDir = {_ps_quote(str(target_dir))}",
        f"$StageDir = {_ps_quote(str(stage_dir))}",
        "",
        "try {",
        "    Wait-Process -Id $WaitPid -Timeout 45 -ErrorAction SilentlyContinue",
        "} catch {}",
        "",
        "$Copied = $false",
        "for ($i = 0; $i -lt 30; $i++) {",
        "    try {",
        "        Copy-Item -LiteralPath $StagedExe -Destination $TargetExe -Force",
        "        $Copied = $true",
        "        break",
        "    } catch {",
        "        Start-Sleep -Milliseconds 400",
        "    }",
        "}",
        "",
        "if ($Copied) {",
        "    Start-Process -FilePath $TargetExe -WorkingDirectory $TargetDir",
        "}",
        "",
        "Start-Sleep -Milliseconds 600",
        "try {",
        "    Remove-Item -LiteralPath $StagedExe -Force -ErrorAction SilentlyContinue",
        "} catch {}",
    ]
    # Pure PowerShell -EncodedCommand cleanup (eliminates shell string interpolation & %VAR% / & hazards)
    clean_ps_code = (
        f"Start-Sleep -Seconds 2; "
        f"Remove-Item -LiteralPath {_ps_quote(str(stage_dir))} -Recurse -Force -ErrorAction SilentlyContinue"
    )
    clean_b64 = base64.b64encode(clean_ps_code.encode("utf-16-le")).decode("ascii")
    ps_lines.append(
        f"Start-Process -FilePath 'powershell.exe' "
        f"-ArgumentList '-NoProfile', '-NonInteractive', '-WindowStyle', 'Hidden', '-EncodedCommand', '{clean_b64}' "
        f"-WindowStyle Hidden"
    )
    script_path.write_text("\r\n".join(ps_lines) + "\r\n", encoding="utf-8-sig")

    creation_flags = 0
    if sys.platform == "win32":
        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000) | getattr(
            subprocess, "DETACHED_PROCESS", 0x00000008
        )

    safe_cwd = str(stage_dir.parent if stage_dir.parent.is_dir() else Path(tempfile.gettempdir()))
    subprocess.Popen(
        [
            "powershell.exe",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-WindowStyle",
            "Hidden",
            "-File",
            str(script_path),
        ],
        cwd=safe_cwd,
        close_fds=True,
        creationflags=creation_flags,
    )
    return script_path


def download_and_apply_update(
    update_info: dict,
    progress_fn: Optional[Callable[[float, str], None]] = None,
    timeout: float = 30.0,
) -> dict:
    """
    Download release ZIP, strictly enforce SHA-256 verification against `SHA256SUMS.txt`
    (both archive and extracted executable, fail-closed), safely extract `零散改点整理工具.exe`,
    and stage detached auto-restart replacement when frozen.
    """
    if not isinstance(update_info, dict):
        raise ValueError("更新信息格式无效。")

    zip_url = str(update_info.get("zip_url") or "").strip()
    sha256_url = str(update_info.get("sha256_url") or "").strip()
    html_url = str(update_info.get("html_url") or f"https://github.com/{GITHUB_REPO}/releases").strip()
    latest_ver = normalize_version_str(str(update_info.get("latest_version") or ""))

    if not zip_url:
        raise ValueError(f"当前新版本 (v{latest_ver}) 尚未包含 Windows 更新包附件 ({ZIP_ASSET_NAME})。")
    if not sha256_url:
        raise ValueError(f"当前新版本 (v{latest_ver}) 缺少官方安全校验文件 ({SHA256_ASSET_NAME})，已拒绝自动更新。")
    validate_github_https_url(zip_url)
    validate_github_https_url(sha256_url)
    if html_url:
        validate_github_https_url(html_url)

    if callable(create_secure_update_stagedir):
        stage_dir = create_secure_update_stagedir()
    else:
        stage_dir = Path(tempfile.mkdtemp(prefix="scattered_rev_update_")).resolve()
    zip_path = stage_dir / ZIP_ASSET_NAME

    try:
        if progress_fn:
            progress_fn(3.0, "正在获取官方 SHA-256 安全校验清单...")
        req_sha = urllib.request.Request(
            sha256_url,
            headers={"User-Agent": _user_agent_header(APP_VERSION)},
            method="GET",
        )
        with _open_with_proxy_fallback(req_sha, timeout=timeout) as resp_sha:
            sums_raw = resp_sha.read(65536).decode("utf-8", errors="replace")
        expected_zip_sha = parse_sha256sums(sums_raw, ZIP_ASSET_NAME)
        expected_exe_sha = parse_sha256sums(sums_raw, EXE_ENTRY_NAME)
        if not expected_zip_sha:
            raise ValueError("SHA256SUMS.txt 中未找到更新压缩包的校验值，已安全中止。")
        if not expected_exe_sha:
            raise ValueError("SHA256SUMS.txt 中未找到主程序的 SHA-256 校验值，已安全中止。")

        if progress_fn:
            progress_fn(6.0, f"正在下载更新包 v{latest_ver} ({ZIP_ASSET_NAME})...")

        req_zip = urllib.request.Request(
            zip_url,
            headers={"User-Agent": _user_agent_header(APP_VERSION)},
            method="GET",
        )
        hasher = hashlib.sha256()
        downloaded = 0
        with _open_with_proxy_fallback(req_zip, timeout=timeout) as resp_zip, open(zip_path, "wb") as fp_out:
            cl_header = resp_zip.headers.get("Content-Length", "").strip()
            total_bytes = int(cl_header) if cl_header.isdigit() and int(cl_header) > 0 else 0
            if total_bytes > MAX_UPDATE_ZIP_BYTES:
                raise ValueError("更新包体积超过安全上限，已中止下载。")

            while True:
                chunk = resp_zip.read(65536)
                if not chunk:
                    break
                downloaded += len(chunk)
                if downloaded > MAX_UPDATE_ZIP_BYTES:
                    raise ValueError("更新包下载体积超过安全上限，已中止。")
                fp_out.write(chunk)
                hasher.update(chunk)
                if progress_fn:
                    mb_done = downloaded / (1024 * 1024)
                    if total_bytes > 0:
                        frac = min(1.0, downloaded / total_bytes)
                        mb_tot = total_bytes / (1024 * 1024)
                        progress_fn(
                            6.0 + frac * 76.0,
                            f"正在下载更新包：{mb_done:.1f} / {mb_tot:.1f} MB ({frac * 100:.0f}%)...",
                        )
                    else:
                        approx_frac = min(0.95, downloaded / (20 * 1024 * 1024))
                        progress_fn(6.0 + approx_frac * 76.0, f"正在下载更新包：已下载 {mb_done:.1f} MB...")

        if downloaded == 0:
            raise ValueError("下载的更新包为空文件。")

        actual_zip_sha = hasher.hexdigest().lower()
        if progress_fn:
            progress_fn(85.0, "正在校验更新包完整性与安全性 (SHA-256)...")
        if actual_zip_sha != expected_zip_sha:
            raise ValueError(
                f"更新包 SHA-256 校验失败（期望 {expected_zip_sha[:12]}...，实际 {actual_zip_sha[:12]}...），已中止更新！"
            )

        if progress_fn:
            progress_fn(91.0, f"正在安全提取新版主程序 ({EXE_ENTRY_NAME})...")
        staged_exe = _validate_and_extract_exe_from_zip(zip_path, stage_dir)

        actual_exe_sha = compute_file_sha256(staged_exe)
        if actual_exe_sha != expected_exe_sha:
            raise ValueError("解压后的主程序 SHA-256 校验不匹配，已安全中止更新！")

        # Remove downloaded ZIP now that EXE is verified & staged
        try:
            zip_path.unlink(missing_ok=True)
        except OSError:
            pass

        is_frozen = bool(getattr(sys, "frozen", False))
        if is_frozen:
            target_exe = Path(sys.executable).resolve()
            if progress_fn:
                progress_fn(97.0, "正在准备自动替换脚本并重启软件...")
            _write_and_launch_windows_updater_script(
                stage_dir=stage_dir,
                staged_exe=staged_exe,
                target_exe=target_exe,
                wait_pid=os.getpid(),
            )
            if progress_fn:
                progress_fn(100.0, "更新就绪，正在自动重启至新版本...")
            return {
                "ready_to_restart": True,
                "staged_exe": str(staged_exe),
                "html_url": html_url,
            }

        if progress_fn:
            progress_fn(100.0, f"当前为源码运行模式，新版 EXE 已下载提取至：{staged_exe}")
        return {
            "ready_to_restart": False,
            "staged_exe": str(staged_exe),
            "html_url": html_url,
        }
    except Exception:
        shutil.rmtree(stage_dir, ignore_errors=True)
        raise
