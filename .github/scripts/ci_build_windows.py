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
    version_file = tool_dir / "version_info.txt"
    pyi_cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--noupx",
        "--onefile",
        "--windowed",
        "--name",
        "ScatteredRevisionTool",
        "--version-file",
        str(version_file),
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
        "sheets_reader",
        "--hidden-import",
        "library_ingester",
        "--hidden-import",
        "proofread_ingester",
        "--hidden-import",
        "ingest_tab_ui",
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
        "## 多语种改点与分part稿自动整理工具 v1.9.1 (Windows x64 免安装便携版)\n\n"
        "- 解压 `ScatteredRevisionTool-Windows-x64.zip` 后直接双击 `零散改点整理工具.exe` 即可运行（无需安装 Python 或任何第三方库；支持软件内一键检查并自动更新）。\n"
        "- **⚠️ 旧版升级提醒**：由于 `v1.8.0`～`v1.9.0` 客户端内置的旧版热更新启动参数存在兼容缺陷（点击更新后退出未能完成文件替换），**正在使用 `v1.9.0` 及更早版本的用户请手动下载本版覆盖一次**；升级至 `v1.9.1` 后，后续所有新版本均可正常一键自动热替换重启！\n"
        "- **v1.9.1 修复内容（客户端自动热更新彻底根治 & 消除杀软误报）**：\n"
        "  - 彻底移除旧版后台脚本与隐式命令行调用，重构为**纯 Python 原生自复制替换架构**（无任何外部命令行脚本、无隐藏文件属性，避免 Windows Defender / 360 / 火绒等安全软件拦截误杀）；\n"
        "  - 自动清理 PyInstaller 单文件 `_PYI_*` / `_MEI*` 运行时环境变量并重置 `PYINSTALLER_RESET_ENVIRONMENT=1`，等待旧版完全退出释放句柄后安全替换并自动重启新版本，同时补充标准 Windows PE 版本元数据并增加压缩包内直接双击未解压运行的防呆提示。\n"
        "- **v1.9.0 全新核心模块与排版质量全面升级（完整包含于本版）**：\n"
        "  1. **新增「改点入库与通读定稿」完整工作台（3A 改点返稿入库 + 3B 通读定稿入库）**：\n"
        "     - **3A 改点返稿入库**：支持发稿/回稿单文件或文件夹自动配对、全栏位原生拖拽、Step 1A 总库一致性核对（保留总库改动或跳过不一致段）、Step 1B 回稿漏改/错位/多删/青色衔接检查并自动打 Word 批注、Step 2 格式清洗与通读表同步、Step 3 原库零修改安全导出 `改前/`、`改后/`（词句级精准回写保留原有蓝红字与底纹，叠加中文底纹）与 `修订稿/`（支持单份汇总/分批汇总/逐篇输出）；\n"
        "     - **3B 通读定稿入库**：支持粘贴篇目+多个 Google Sheets 链接（严格按 `gid` 仅读指定 Sheet）或本地 `.xlsx/.csv`、导入 Google JSON 授权文件并浏览器登录授权（DPAPI 加密保存）、自定义 5 列列名、改点标点与首字母大写质检、严格精确匹配替换及多 Part 全篇合并核对；\n"
        "  2. **零散与全篇改点体验升级**：新增「仅英外标记黄/青（默认） / 中英外同步标记黄/青 / 仅输出对比修订」三选一高亮模式；第 3 列外文局部改点（UP）升级为柔和浅蓝底纹（`#D9EEF7`）；「提取三语原稿」100% 原样保留总库既有高亮与底纹颜色；统一横向 A4 默认页边距（上 2.0cm、下 2.2cm、左 2.0cm、右 2.0cm）；\n"
        "  3. **分part稿整理/合并深度修复与全语种高度自适应拉齐**：\n"
        "     - 彻底消除印地语、老挝语等输出文档在 Word 打开时的「发现无法读取的内容」提示（清理 `mc:Ignorable` 残留前缀并强制 `CT_PPr` / `CT_RPr` 标准顺序）；\n"
        "     - 修复印地语等语种引号内首尾词漏加粗及 SOV 语序颠倒导致的引号外蓝字错标问题；\n"
        "     - 表格上方蓝色标题自动设为「二级标题层级」（`outlineLvl=1`，导航窗格直接识别显示）；\n"
        "     - 全语种动态模拟折行高度自动拉齐三列列宽，高棉语默认 `Leelawadee UI`、老挝语默认 `DokChampa`，并支持在「语种库」中永久存档各语种自定义标题/正文/摘自字号与行距。\n"
    )
    (art_dir / "RELEASE_NOTES.md").write_text(notes_md, encoding="utf-8")
    print(sums_txt.strip())


if __name__ == "__main__":
    main()
