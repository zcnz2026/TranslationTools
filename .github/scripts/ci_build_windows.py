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
        "library_index",
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
        "word_counter",
        "--hidden-import",
        "win_dnd",
        "--hidden-import",
        "fellowship_processor",
        "--hidden-import",
        "pipeline",
        "--hidden-import",
        "full_article_processor",
        "--hidden-import",
        "dedup_checker",
        "--hidden-import",
        "updater",
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
        "## 多语种改点与分part稿自动整理工具 v1.7.0 (Windows x64 免安装便携版)\n\n"
        "- 解压 `ScatteredRevisionTool-Windows-x64.zip` 后直接双击 `零散改点整理工具.exe` 即可运行（无需安装 Python 或任何第三方库；支持软件内一键检查并自动更新）。\n"
        "- **v1.7.0 重大升级（5 大独立工作台全面成型）**：\n"
        "  1. **📖 新增「全篇改点整理」工作台**：支持粘贴或导入标题清单（`.docx` / `.txt`），兼容单篇底稿与合集大文档自动拆篇，多篇同名标题（如 `第一篇`）弹窗勾选确认，按首行加粗英文标题逐篇独立导出**完整版**与**仅改点版**；\n"
        "  2. **🔍 新增「改点稿查重」工作台 & 零散/全篇联动去重**：基于第 1 列中文四级匹配（整段 -> 段首 -> 段尾 -> 着色片段）对比「未入库改点包」，一键生成《查重结果明细报告》并导出剔除重复段、自动重算字数的发译者定稿；\n"
        "  3. **📝「零散改点整理」深度增强**：修复全局 OOXML `w:id` 冲突与 `people.xml` 残留（彻底根除“发现无法读取的内容”弹窗，同时保留原稿第 2 列与总库第 3 列批注及原版式）；总库实时读取永不陈旧 +「🔄 刷新总库」按钮；2 列自动扩为 3 列横向 A4 + 三语列宽自动拉齐；同篇重复段落底色自动合并；支持 `God's words [in 关键词]` 标准命名、文档开头 `Blue/Yellow highlighted parts...` 说明行自动生成、以及 `0 words` 自动省略；\n"
        "  4. **🔢「英文字数统计」升级**：新增 3 种命名模式（保留原名 / 全篇首行英文标题 / 零散 `God's words [in 关键词]`）、`0 words` 自动省略及可选同步更新开头高亮说明行；\n"
        "  5. **📚「分part稿整理/合并」全面升级**：支持多文件 `Part 1..N` 排序合并与**单文件内多 Part 表格一键合并**；彻底修复合并单元格 `gridSpan` 阶梯错位；严格解耦蓝字、红字（可选同步神话红字）与加粗同步；覆盖圣经与神话摘自出处字号规范；内置《各语系字体》规范（含泰语 `Angsana New` 18/16/14pt）并自动清理句尾多余空格与格底空行；\n"
        "  6. **🔄 软件内置自动更新**：启动后台静默检测新版本并在顶部温和提示，支持一键校验 SHA-256 下载热更新重启（无损保留本机加密 API Key 与自定义语种配置）。\n"
    )
    (art_dir / "RELEASE_NOTES.md").write_text(notes_md, encoding="utf-8")
    print(sums_txt.strip())


if __name__ == "__main__":
    main()
