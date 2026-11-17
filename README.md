# TranslationTools — 多语种三列改点与分part稿自动整理工具 (v1.7.0)

轻量级 Windows 桌面可视化软件与纯标准库 OpenXML 排版引擎套件，专为多语种三列对照（中文 / 英文 / 外文）Word 文档打造，内置**五大专业工作台**：**📝 零散改点整理**、**📖 全篇改点整理**、**🔍 改点稿查重**、**🔢 英文字数统计**与**📚 分part稿整理/合并**，并支持**软件内一键在线检查与热更新**。

---

## 目录结构

```text
TranslationTools/
├── .agents/skills/scattered-revision/   # 核心 OpenXML 解析、增量总库索引与排版保真脚本
│   ├── SKILL.md                         # 业务规则与工作流规范说明
│   ├── resources/config.json            # 默认颜色、阈值与基础语种映射配置
│   └── scripts/                         # 纯标准库 Word XML 处理引擎（零第三方运行时依赖）
│       ├── docx_xml_utils.py            # OOXML 合规清理、全局唯一 w:id、批注迁移、双列扩三列、同篇重复段合并、三语拉齐
│       ├── library_index.py             # 多线程持久化增量总库索引（实时 mtime 校验 + 支持强制刷新）
│       ├── step1_extract_and_diff.py    # 五级级联检索、总库实时读取、支持浅蓝底色局部对比与全篇强制对比
│       ├── step1_render_draft.py        # 第 2/3 列批注双端保留与总库第 3 列原版式（对齐/字体/字号/粗细）保真合入
│       ├── step2_parse_reviewed.py      # 第 2 列黄(UP)/青(RV)区间与全句坐标提取
│       └── step2_render_final.py        # 第 3 列外文标色、顶部高亮说明行、0词省略命名与三版定稿导出
└── 零散改点工具/                          # 桌面可视化软件 (Tkinter GUI · v1.7.0)
    ├── app_gui.py                       # 左侧五大工作台导航、同名标题消歧弹窗、语系字体管理与自动更新界面
    ├── ai_client.py                     # 多线程 AI 分级/对齐、离线本地位置估算兜底与简洁中文错误提示
    ├── pipeline.py                      # 零散改点一键输出定稿流水线（支持标准命名、三语拉齐与联动未入库查重）
    ├── full_article_processor.py        # 场景二：全篇改点整理引擎（支持单篇与合集大文档拆篇、同名消歧、逐篇双版导出）
    ├── dedup_checker.py                 # 场景三：未入库改点包第 1 列中文四级查重、明细报告生成与自动去重重算引擎
    ├── word_counter.py                  # 第 2 列英文单词统计（3 种命名模式、0 词自动省略、同步更新顶部说明行）
    ├── fellowship_processor.py          # 场景四/五：分part稿整理/合并引擎（多文件/单文件多Part合并、蓝/红/粗细同步、语系字体）
    ├── lang_registry.py                 # 多语种与《各语系字体》标准字号注册表（含泰语 Angsana New 18/16/14pt 等）
    ├── updater.py                       # GitHub Releases 自动检测更新、SHA-256 校验与保留配置无感热替换重启
    ├── win_dnd.py                       # 原生 Win32 文件与文件夹全局拖拽引擎（纯标准库 ctypes）
    ├── security_utils.py                # Windows DPAPI 加密存储、脱敏日志、ZipBomb/XXE/SSRF 防护
    ├── languages.json                   # 外置多语种与标准字体配置文件
    ├── .env.example                     # 环境变量配置示例
    ├── requirements-build.txt           # 打包独立 .exe 所需的构建依赖
    ├── 启动工具.bat                      # 脚本直连启动入口
    └── 使用说明.md                       # 详细操作手册 (v1.7.0)
```

---

## 核心特性 (v1.7.0)

1. **📝「零散改点整理」一键高保真定稿**：
   - **100% 合规 OOXML 与双端批注无损保留**：统一分配全局唯一标记 `w:id`，清理内部临时书签与悬空 `people.xml`，同时完整保留原稿第 2 列批注与总库第 3 列批注及原段落版式（两端/居中对齐、字体、字号、加粗状态），Word 打开绝不弹“无法读取的内容”，支持直接 `Ctrl+S` 保存。
   - **实时总库读取 +「🔄 刷新总库」**：命中总库文件后实时从磁盘读取最新正文（刚在 Word 保存的总库修改立即生效，绝不提取旧缓存）。
   - **双列自动扩三列 + 同篇重复段合并 + 三语拉齐**：2 列表格自动扩展为横向 A4 三列表格并均衡列宽；同一文档内重复出现的中文段落自动合并不同位置的底色改点。
   - **零散标准命名 + 顶部说明行 + 联动查重**：支持 `God's words [in 关键词]` 标准命名、自动生成开头 `Blue/Yellow highlighted parts...` 说明行、`0 words` 自动省略，并可关联「未入库改点包」自动剔除发译者版重复段。

2. **📖「全篇改点整理」独立工作台**：
   - 支持直接粘贴多行标题或拖入/导入 `.docx` / `.txt` 标题清单；同时兼容**单篇底稿文件**与 **5MB+ 多篇合集大文档**自动拆篇检索。
   - 遇到 `第一篇` 等跨卷同名标题时自动弹出**多篇同名标题选择确认弹窗**（展示完整卷名/分部路径与行数供勾选）。
   - 全章逐段对比（无需原稿涂底色），自动提取首行第 2 列英文标题命名，**逐篇独立导出** `完整版` 与 `仅改点版`。

3. **🔍「改点稿查重」独立工作台**：
   - 针对「未入库改点包」执行第 1 列中文四级极速查重（① 整段高相似 → ② 段首 1~3 句 → ③ 段尾 1~3 句 → ④ 带颜色/底色句子片段），自动导出横向 A4 **`查重结果明细报告_YYYYMMDD.docx`**，并可一键导出剔除重复段且重算字数的新定稿。

4. **🔢「英文字数统计」多模式重命名**：
   - 支持**保留原文档名**、**全篇改点（取首行第 2 列英文标题）**、**零散改点（`God's words [in 关键词]`）**三种命名模式；当 UP 或 RV 为 `0` 时自动省略 `0 words` 项；支持一键在文档开头插入/更新 `Blue / Yellow highlighted parts...` 说明行。

5. **📚「分part稿整理/合并」多语系精排**：
   - 同时支持**多份分篇文件 (`Part 1..N`) 排序合并**与**单个文档内多段 `Part` 标题+表格一键合并**，彻底消除 `gridSpan` 合并单元格导致的竖线错位。
   - 严格按第 2 列英文独立同步**蓝字 `#0000FF`**、**红字（可选开启，适用于神话定稿）**与**加粗**（绝不把蓝字引号或段间叙述误加粗，`Question 10:` 等前缀自动同步）。
   - 同时识别圣经出处与神话摘自出处设为五号字（`10.5pt`，泰语 `14pt`）；内置《各语系字体》规范（如泰语 `Angsana New` `18/16/14pt`、印地语 `Mangal`、波兰/法/西/葡/希 `Georgia` 等），自动转为横向 A4 并清理句尾多余空格与单元格底部空行。

6. **🔄 软件内一键自动更新 + 无限 Key 池轮换 + 离线本地标色兜底**。

---

## 快速运行与打包

### 0. 免安装直接下载 & 软件内自动更新

直接前往本仓库 [**Releases 页面**](https://github.com/zcnz2026/TranslationTools/releases) 下载 `ScatteredRevisionTool-Windows-x64.zip`，解压后双击 `零散改点整理工具.exe` 即可运行。后续版本可直接点击软件左下角 **「🔄 检查软件更新」** 一键完成 SHA-256 校验下载与自动替换重启（本机 API Key 与语种配置完整保留）。

### 1. 直接源码运行（仅需 Python 3.10+ 标准库，零第三方运行时依赖）

双击 `零散改点工具/启动工具.bat`，或在终端执行：

```powershell
python .\零散改点工具\app_gui.py
```

### 2. 打包为免安装独立 Windows 单文件 `.exe`

```powershell
pip install -r .\零散改点工具\requirements-build.txt

python -m PyInstaller --noconfirm --clean --onefile --windowed `
  --name "零散改点整理工具" `
  --distpath ".\零散改点工具" `
  --paths ".\零散改点工具" `
  --paths ".\.agents\skills\scattered-revision\scripts" `
  --add-data ".\.agents\skills\scattered-revision\scripts;skill_scripts" `
  --add-data ".\.agents\skills\scattered-revision\resources;resources" `
  --hidden-import docx_xml_utils `
  --hidden-import library_index `
  --hidden-import step1_extract_and_diff `
  --hidden-import step1_render_draft `
  --hidden-import step2_parse_reviewed `
  --hidden-import step2_render_final `
  --hidden-import ai_client `
  --hidden-import lang_registry `
  --hidden-import security_utils `
  --hidden-import word_counter `
  --hidden-import win_dnd `
  --hidden-import fellowship_processor `
  --hidden-import pipeline `
  --hidden-import full_article_processor `
  --hidden-import dedup_checker `
  --hidden-import updater `
  ".\零散改点工具\app_gui.py"
```
