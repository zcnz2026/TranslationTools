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
    r"(AIza[0-9A-Za-z_-]{16,}|AQ\.[0-9A-Za-z._-]{16,}|ya29\.[0-9A-Za-z._-]{16,}|"
    r"dpapi:v1:[A-Za-z0-9+/=]{12,}|posix:v1:[A-Za-z0-9+/=:]{12,})"
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

    print("[1/5] Running static syntax, eager-annotation & secret-leak checks...", flush=True)
    py_files = sorted(list(tool_dir.rglob("*.py")) + list(skill_scripts_dir.rglob("*.py")))
    for pf in py_files:
        py_compile.compile(str(pf), doraise=True)
        txt = pf.read_text(encoding="utf-8")
        if SECRET_RX.search(txt):
            raise RuntimeError(f"Hardcoded secret pattern detected in {pf.name}!")
    # Verify all module imports & type annotations resolve cleanly (prevents Python 3.13 vs 3.14+ PEP 649 drift)
    import builtins
    import inspect
    import symtable

    builtin_set = set(dir(builtins))
    sys.path.insert(0, str(tool_dir))
    sys.path.insert(0, str(skill_scripts_dir))
    ann_lib = None
    try:
        import annotationlib as _al  # Python 3.14+
        ann_lib = _al
    except ImportError:
        pass
    for pf in py_files:
        mod = __import__(pf.stem)
        st_top = symtable.symtable(pf.read_text(encoding="utf-8"), str(pf), "exec")
        top_defined = {s.get_name() for s in st_top.get_symbols() if s.is_local() or s.is_imported()}
        mod_globals = set(mod.__dict__.keys()) | top_defined | builtin_set

        def _check_symtable(tbl, enclosing_defs):
            local_defs = set(enclosing_defs)
            for sym in tbl.get_symbols():
                if sym.is_local() or sym.is_parameter() or sym.is_imported() or sym.is_declared_global():
                    if not sym.is_declared_global():
                        local_defs.add(sym.get_name())
            if tbl.get_type() == "function":
                for sym in tbl.get_symbols():
                    nm = sym.get_name()
                    if sym.is_global() and sym.is_referenced() and nm not in mod_globals and nm not in local_defs:
                        raise RuntimeError(
                            f"Undefined global symbol {nm!r} in {pf.name}:{tbl.get_name()} (line {tbl.get_lineno()})"
                        )
            next_defs = local_defs if tbl.get_type() == "function" else enclosing_defs
            for child in tbl.get_children():
                _check_symtable(child, next_defs)

        _check_symtable(st_top, set())
        if ann_lib is not None:
            for _, obj in inspect.getmembers(mod):
                if inspect.isfunction(obj) or inspect.isclass(obj):
                    ann_lib.get_annotations(obj, format=ann_lib.Format.VALUE)
                if inspect.isclass(obj):
                    for _, m in inspect.getmembers(obj, inspect.isfunction):
                        ann_lib.get_annotations(m, format=ann_lib.Format.VALUE)
    print(f"      Passed ({len(py_files)} Python files verified clean).", flush=True)

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

    print("[3.5/5] Verifying frozen EXE self-test (bundled config.json & modules)...", flush=True)
    self_test_log = work_root / "self_test_error.log"
    env_st = dict(env)
    env_st["SCATTERED_SELF_TEST_LOG"] = str(self_test_log)
    try:
        subprocess.run([str(built_exe), "--self-test"], env=env_st, check=True, timeout=60)
    except Exception:
        if self_test_log.is_file():
            print("=== Frozen EXE self-test traceback ===", flush=True)
            print(self_test_log.read_text(encoding="utf-8", errors="replace"), flush=True)
        raise
    print("      Frozen EXE self-test passed!", flush=True)

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
        "## 多语种改点与分part稿自动整理工具 v1.9.8 (Windows & macOS 双平台免安装版)\n\n"
        "### 📥 下载与运行指引（按电脑系统选择对应压缩包）\n"
        "- **Windows 电脑 (`x64`)**：下载 `ScatteredRevisionTool-Windows-x64.zip`，解压后直接双击 `零散改点整理工具.exe` 即可运行。\n"
        "- **Mac 电脑 — Apple M1 / M2 / M3 / M4 芯片 (`arm64`)**：下载 `ScatteredRevisionTool-macOS-arm64.zip`，解压后打开 `零散改点整理工具.app`（建议拖入「应用程序 / Applications」文件夹使用）。\n"
        "- **Mac 电脑 — Intel 芯片旧款机型 (`x64`)**：下载 `ScatteredRevisionTool-macOS-x64.zip`，解压后打开 `零散改点整理工具.app`。\n"
        "- **💡 Mac 首次打开提示**：由于未绑定付费 Apple 企业公证证书，首次在 Mac 打开若弹出「无法验证开发者 / 未打开」提示，请**按住 `Control` 右键点击 `零散改点整理工具.app` → 选择「打开」**（或在 `系统设置 → 隐私与安全性` 底部点击「仍要打开」，亦可在终端执行一次 `xattr -cr \"零散改点整理工具.app\"`），之后即可正常双击启动与在线自动更新。\n"
        "- **⚠️ 升级说明**：Windows 端与 Mac 端用户直接在软件左下角点击「检查软件更新 → 立即更新」即可一键自动升级重启至 `v1.9.8`！\n\n"
        "- **v1.9.8 紧急修复内容**：\n"
        "  - **修复总库索引缓存路径解析 `name 'sys' is not defined` 报错**：补齐 `library_index.py` 的 `import sys`（修复 v1.9.7 执行零散改点/全篇改点扫描总库时的异常），并在 CI 打包流水线中新增全函数级 `symtable` 未定义全局变量静态阻断检查。\n"
        "- **v1.9.7 新增内容（正式新增原生 macOS 双架构版本）**：\n"
        "  - **新增原生 macOS 桌面客户端（同时提供 Apple Silicon `arm64` 与 Intel `x64` 独立应用包）**：完整支持全部六大核心工作台（零散改点整理、全篇改点整理、改点稿查重、3A/3B 改点入库与通读定稿、英文字数统计、分part稿整理/合并）及系统设置；\n"
        "  - **针对 macOS Aqua 深度适配字体、按钮渲染、拖拽与安全密钥存储**：自动采用苹方（`PingFang SC`）/ `Menlo` 高清字体比例，修复 macOS 原生按钮白底白字不可见问题，支持 Finder 文件/文件夹拖入，个人配置与加密 API 密钥安全独立存放于 `~/Library/Application Support/ScatteredRevisionTool/`（保持 `.app` 签名完整且升级绝不丢失配置），并支持 Mac 端一键自动更新替换 `.app`。\n"
        "- **v1.9.6 新增与深度升级内容**：\n"
        "  - **支持 Vertex AI 收费版纯 API Key & 一键动态加载最新模型列表**：系统设置支持直接填入 Vertex AI 收费版纯 API Key（无需配置区域或项目），自动兼容 Gemini Developer API 与 Vertex AI `v1` / `v1beta1` 接口；新增「🔄 加载最新模型」按钮动态拉取并缓存最新可用模型列表，下拉框同时支持自由选择或手动输入任意模型；\n"
        "  - **重构黄（`UP` 局部改点）/ 青（`RV` 重翻）三大行数间距量化规则 + 连续单词合并修订**：严格按间隔行数与修改规模分级（远距小改留黄、超四五行长句重构整句标青且绝不从逗号/分号断开、同段高频近距改点集中区段连片标青），并将相邻仅隔空格的连续词替换合并为一处整洁短语修订（消除逐词红绿交替碎片）；\n"
        "  - **补齐 34+1 语系 & 新增韩语（`KO`/`KR`）独立中韩两列改点工作流**：补齐立陶宛语（`LT`）、克利奥语（`HT`）等全部 34 个英转外语系及别名；韩语全程保持 `[中文 | 韩文]` 两列结构，直接对比总库旧中文与新中文生成中文修订及黄/青标色，同步映射韩文底色并按中文字数生成规范命名；\n"
        "  - **3A 改点入库新增中文列无损覆盖更新（保持原有颜色不动）**：`改后/` 第 1 列中文按单字级无损对比同步覆盖最新文本值（自动兼容诗歌/选段局部回写），100% 保留总库第 1 列原有字体颜色、高亮与单元格底纹不动，并在 `修订稿/` 中同步展示中文修订痕迹；\n"
        "  - **`you(s)` / `your(s)` / `you(pl)` / `your(pl)` 保留原有修订与单复数变化标黄区分**：纯新增/移除 `(s)`/`(pl)` 标注原样保留修订痕迹（不接受、不拒绝、不标色、不算词数）；单复数发生实质变化（如 `you(s)` -> `you(pl)`）则保留修订、标记黄色 `UP`、同步映射外文底色并计入字数发给译者修改。\n"
        "- **v1.9.5 新增与深度修复内容**：\n"
        "  - **「零散改点整理」支持批量多文档与整文件夹导入执行**：新增「选文件夹...」及多选文件支持，拖入整个文件夹或多个 `.docx` 时自动识别全部待整理文档，总库索引仅预热一次并按 `给译者版/`、`完整留底版/`、`无改点版/`、`提取三语原稿/` 分目录防重名整齐归档；\n"
        "  - **3A 改点返稿入库全流程核对与青色（`RV`）完整入库修复**：第一步 1A 先彻底拒绝发稿修订，对比总库后自动接受青色（`RV`）正常差异，仅检查并输出剩余真实不一致修订；外文比对自动忽略因突出浅灰句而产生的空行拆段，彻底解决青色英文漏入库、外文分段边界串词及 `步1B` 青色英文无修订痕迹的问题；保留印尼语等省略号前固有空格 ` ....`；\n"
        "  - **3A & 3B 通读定稿表多链接识别、严格校验替换与 5 列完整汇报**：支持多链接（严格按 `#gid=` 读取）及两列带标题直接粘贴；仅当确实有改动（`改动列 != 改前列` 且 `改动列 == 定稿列`）并在回稿/文档中找到 100% 完全一致文本时执行替换；对 `改动列 != 定稿列`、`改动列 == 改前列`、找不到完全一致文本、以及半句↔整句标点突变/篇幅异常/与中文意思不吻合的行，一律在报告横表中把 **5 列（中文、English、改前列、定稿列、改动列）** 完整列出。\n"
        "- **v1.9.4 新增与深度修复内容**：\n"
        "  - **新增顶部「🧹 清理缓存」按钮 & 防串稿自动重置**：一键清理总库内存索引、预览缓存、临时工作目录及残留文件列表；「分part稿整理/合并」（及查重、字数统计）在上一轮执行完成后再次拖入新文稿时自动清空上一轮旧列表，彻底杜绝跨批次合并串稿；\n"
        "  - **第 3 列红色字体默认同步 & 逐处补齐防漏标**：「分part稿整理/合并」默认勾选同步红色字体至第 3 列，升级波兰语等变音符号模糊匹配，并按每一处红色片段逐一核对补齐（彻底解决单段多处红字漏标后半处的问题）；\n"
        "  - **全底层 OOXML 严格校验器（彻底根治「Word 发现无法读取的内容」）**：在全模块统一保存出口自动剥离跨文档迁移产生的失效超链接 `r:id`、清理跨版本扩展属性与孤立 `mc:Ignorable` 声明、补齐空单元格 `<w:p>` 并强制重排全部 `tblPr / trPr / tcPr / pPr / rPr` 标准标签顺序；\n"
        "  - **「分part稿整理/合并」三类批注智能核对与择优保留**：自动识别并核对 ① 直接外文标准译文批注、② 按英文节选截取对应外文批注、③ `Needs to be consistent with...above / Please quote...above` 等与上文译法保持一致的英文指令批注（同步校验对应蓝/红字体颜色）：完全一致自动删除批注，文字或颜色不一致自动原样保留批注并在日志与弹窗中提示具体行号供人工复核。\n"
        "- **v1.9.3 全面体验优化与细节修复**：\n"
        "  - **全自动 `.docx` 无损瘦身**：保存/导出任意文档时自动剥离嵌入字体库（`word/fonts/*.odttf`）、缩略图与冗余 `rsids` 并强制 Deflate 高效压缩，彻底解决原稿 16MB / 2MB 虚胖问题（自动缩减至几十 KB 且 100% 保持原排版、批注与样式）；\n"
        "  - **规范文档命名（`RV` 在前、`UP` 在后）与冗余后缀清理**：字数标签统一调整为 `(RV Y words, UP X words)` 便于公式统一抓取；自动剥离原底稿残留的 `_( words)_旧日期`、`_完整版` 等冗余尾巴，全篇改点统一归档至 `给译者版/` 与 `完整留底版/` 子目录；\n"
        "  - **标色与下划线默认优化**：第 3 列外文局部改点（UP）统一使用柔和浅蓝底纹（`#D9EEF7`），中文列（第 1 列）改为默认勾选「中文列对应改点标记下划线」（绝不叠加黄/青高亮颜色，支持随时取消勾选）；\n"
        "  - **精准校正青色重翻（`RV`）与黄色微调（`UP`）判定边界**：杜绝因长段落内零星小改点数量多而误将整段标青（严格校验真实改动词占比 `< 42%` 绝不整段升青），同时确保句式结构大幅改写打碎的句子准确标为青色重翻；\n"
        "  - **完整保留双行附篇标题 & 引号蓝字/加粗硬边界隔离**：修复 `附篇一 + 副标题` 双行表头在提取/筛选时少一行的问题；强化「分part稿整理/合并」引号边界钳制与引号内首尾词补齐（消除印尼语等引号外叙述文字误标蓝/加粗及引号内首词漏标）；\n"
        "  - **全模块等高三语拉齐 & 新打开软件自动清空全部路径**：统一全部工作台的模拟折行等高列宽求解器；每次新启动软件自动清空全部 6 个板块的文件/目录/总库/标题输入（仅保留系统设置 API 密钥池、模型与语种库存档）。\n"
        "- **v1.9.2 修复与加固内容（3A 改点返稿入库 & 3B 通读定稿入库全流程验证修复）**：\n"
        "  - **修复 3A 验收批注报错**：修复回稿检出漏改/异常打批注时 `'MarkupIdAllocator' object has no attribute 'allocate'` 的报错；\n"
        "  - **3A 诗歌/节选片段精准回写与篇目消歧**：自动追踪表格上方篇目标题辅助定位；当发稿仅引用长段/诗歌中的部分诗节时，仅精准替换对应诗节而不误报冲突、不覆盖其余诗节；自动忽略发稿第 1 列浅蓝定位底色（`DAE9F7`），100% 保持原稿既有底纹与高亮；修正合并修订汇总稿的全局唯一修订 ID 与 `CT_RPr` 顺序，杜绝 Word 打开报错；\n"
        "  - **修复 3B 通读定稿入库兜底匹配与智能列名取值**：修复模糊标题回退时的 `LibraryIndex.files_map` 属性问题，支持直接拖入本地通读 `.xlsx` 时自动根据中文行定位涉及篇目，并统一「改动列 / 定稿列」智能判别逻辑。\n"
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
