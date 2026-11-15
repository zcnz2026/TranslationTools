import hashlib
import os
import py_compile
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

TOOL_DIR_NAME = "\u96f6\u6563\u6539\u70b9\u5de5\u5177"  # 零散改点工具
EXE_FINAL_NAME = "\u96f6\u6563\u6539\u70b9\u6574\u7406\u5de5\u5177.exe"  # 零散改点整理工具.exe
DOC_MANUAL_NAME = "\u4f7f\u7528\u8bf4\u660e.md"  # 使用说明.md

SECRET_RX = re.compile(
    r"(AIza[0-9A-Za-z_-]{16,}|AQ\.[0-9A-Za-z._-]{16,}|ya29\.[0-9A-Za-z._-]{16,}|dpapi:v1:[A-Za-z0-9+/]{12,})"
)


def sha256_file(p: Path) -> str:
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().lower()


def main() -> None:
    check_only = "--check-only" in sys.argv
    repo_root = Path(__file__).resolve().parents[2]
    tool_dir = repo_root / TOOL_DIR_NAME
    skill_dir = repo_root / ".agents" / "skills" / "scattered-revision"
    skill_scripts_dir = skill_dir / "scripts"
    skill_resources_dir = skill_dir / "resources"
    req_file = tool_dir / "requirements-build.txt"

    if not tool_dir.is_dir():
        raise FileNotFoundError(f"Tool directory not found: {tool_dir}")
    if not skill_scripts_dir.is_dir():
        raise FileNotFoundError(f"Skill scripts directory not found: {skill_scripts_dir}")

    print("[1/5] Running static syntax & secret-leak checks...")
    py_files = sorted(list(tool_dir.rglob("*.py")) + list(skill_scripts_dir.rglob("*.py")))
    for pf in py_files:
        py_compile.compile(str(pf), doraise=True)
        txt = pf.read_text(encoding="utf-8")
        if SECRET_RX.search(txt):
            raise RuntimeError(f"Hardcoded secret pattern detected in {pf.name}!")
    print(f"      Passed ({len(py_files)} Python files verified clean).")

    if check_only:
        return

    print("[2/5] Installing pinned PyInstaller build dependencies...")
    env = dict(os.environ)
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"], env=env)
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", str(req_file)], env=env)

    print("[3/5] Building standalone Windows x64 executable with PyInstaller...")
    work_root = repo_root / "ci_build_tmp"
    if work_root.exists():
        shutil.rmtree(work_root)
    dist_dir = work_root / "dist"
    work_dir = work_root / "work"
    spec_dir = work_root / "spec"

    add_scripts_arg = f"{skill_scripts_dir}{os.pathsep}skill_scripts"
    add_resources_arg = f"{skill_resources_dir}{os.pathsep}resources"
    pyi_cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name",
        "ScatteredRevisionTool",
        "--distpath",
        str(dist_dir),
        "--workpath",
        str(work_dir),
        "--specpath",
        str(spec_dir),
        "--paths",
        str(tool_dir),
        "--paths",
        str(skill_scripts_dir),
        "--add-data",
        add_scripts_arg,
        "--add-data",
        add_resources_arg,
        "--hidden-import",
        "docx_xml_utils",
        "--hidden-import",
        "step1_extract_and_diff",
        "--hidden-import",
        "step1_render_draft",
        "--hidden-import",
        "step2_parse_reviewed",
        "--hidden-import",
        "step2_render_final",
        "--hidden-import",
        "ai_client",
        "--hidden-import",
        "lang_registry",
        "--hidden-import",
        "security_utils",
        "--hidden-import",
        "pipeline",
        str(tool_dir / "app_gui.py"),
    ]
    subprocess.check_call(pyi_cmd, env=env)

    built_exe = dist_dir / "ScatteredRevisionTool.exe"
    if not built_exe.is_file():
        raise RuntimeError(f"Expected built EXE not found: {built_exe}")

    print("[3.5/5] Verifying frozen EXE self-test (bundled config.json & modules)...")
    subprocess.check_call([str(built_exe), "--self-test"], env=env)
    print("      Frozen EXE self-test passed!")

    print("[4/5] Packaging portable ZIP with UTF-8 filenames...")
    art_dir = repo_root / "artifact_out"
    if art_dir.exists():
        shutil.rmtree(art_dir)
    art_dir.mkdir(parents=True, exist_ok=True)

    zip_path = art_dir / "ScatteredRevisionTool-Windows-x64.zip"
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        zf.write(built_exe, arcname=EXE_FINAL_NAME)
        zf.write(tool_dir / "languages.json", arcname="languages.json")
        zf.write(tool_dir / DOC_MANUAL_NAME, arcname=DOC_MANUAL_NAME)
        zf.write(tool_dir / ".env.example", arcname=".env.example")

    print("[5/5] Generating SHA256SUMS.txt & RELEASE_NOTES.md...")
    exe_sha = sha256_file(built_exe)
    zip_sha = sha256_file(zip_path)
    sums_txt = f"{exe_sha}  {EXE_FINAL_NAME}\n{zip_sha}  ScatteredRevisionTool-Windows-x64.zip\n"
    (art_dir / "SHA256SUMS.txt").write_text(sums_txt, encoding="utf-8")

    notes_md = (
        "## 零散改点自动整理工具 (Windows x64 免安装便携版)\n\n"
        "- 解压 `ScatteredRevisionTool-Windows-x64.zip` 后直接双击 `零散改点整理工具.exe` 即可运行（无需安装 Python 或任何第三方库）。\n"
        "- 支持 **无限多 API Key 自动轮换 & 超额/报错秒切**（兼容混填 Vertex AI Key 与 Gemini Key）。\n"
        "- 支持 **35+ 多语种自由新增扩展**（外置 `languages.json`）。\n"
        "- 本机保存密钥采用 **Windows DPAPI** 系统级加密，磁盘与日志零明文泄露。\n"
        "- 附带 `SHA256SUMS.txt` 文件完整性校验值。\n"
    )
    (art_dir / "RELEASE_NOTES.md").write_text(notes_md, encoding="utf-8")
    print(sums_txt.strip())


if __name__ == "__main__":
    main()
