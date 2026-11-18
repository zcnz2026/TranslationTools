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
        "master_sync",
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
        "## 多语种改点与分part稿自动整理工具 v1.8.0 (Windows x64 免安装便携版)\n\n"
        "- 解压 `ScatteredRevisionTool-Windows-x64.zip` 后直接双击 `零散改点整理工具.exe` 即可运行（无需安装 Python 或任何第三方库；支持软件内一键检查并自动更新）。\n"
        "- **v1.8.0 重大升级（六大板块布局 + 标色精度与全篇/零散工作流全面进化）**：\n"
        "  1. **新增「系统设置」独立配置板块 & 极简全局状态栏**：左侧「全局配置」下集中管理 API 密钥池（内嵌多行编辑器与检测弹窗实时同步）、模型与接口通道切换、**两处黄色高亮自动连片最大间隔词数（默认 ≤3 词，支持 0~20 自由调节）**以及常用语种置顶概览；\n"
        "  2. **语种下拉框「★ 常用置顶 + ⏱ 最近使用」智能排前**：所有板块的目标语种下拉框自动将常用置顶语种与最近使用语种排在最上方，切换语种无需反复滚动查找；\n"
        "  3. **「零散改点整理」无损底稿核对与多列同步升级**：新增「核对最新中英底稿」（支持单文件或文件夹，100% 保留待处理稿浅蓝/粉红底纹、蓝/红字、字号与批注，无损更新文字并输出 `{原名}{日期}更新.docx`）；新增「对比标记高亮颜色」开关（取消勾选仅保留修订痕迹）、默认导出「提取三语原稿」供核对；第 1 列中文自动接受修订并同步黄/青标色；逐字符保留第 2 列圣经引文蓝字/红字不串色；发译者版彻底剔除同表无改点行；单文档平铺输出不套多余子文件夹；\n"
        "  4. **「全篇改点整理」双模式直连与智能分包报告**：新增「现成双文件夹直连对比模式」（免输标题按顺序逐篇对比并以中英底稿命名）；全面兼容 `3-1-39 032 ...（审判）` 等双编号前缀与括号尾注检索并在输出文件名保留前缀；英文无改点篇目自动跳过单篇导出并统一汇总至《全篇对比_英文无改点篇目汇总报告》；多篇对比自动按 `给译者版 / 完整留底版 / 中英最新底稿包 / 外文三语原稿包` 四大子目录归档（单篇自动平铺）；新增「刷新底稿」按钮与每次运行独立临时目录；\n"
        "  5. **标色边界规则精调**：句号改冒号仅从冒号起标黄；半句局部改写严格判为黄色 UP 不误扩青色；零碎全段大改（仅剩 1~2 句未动）自动升级整段青色 RV；同一段两处黄色高亮间隔 ≤N 词自动连片标色。\n"
    )
    (art_dir / "RELEASE_NOTES.md").write_text(notes_md, encoding="utf-8")
    print(sums_txt.strip())


if __name__ == "__main__":
    main()
