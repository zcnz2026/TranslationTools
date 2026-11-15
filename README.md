# TranslationTools — 多语种三列表格零散改点自动整理工具 (v1.6.0)

轻量级 Windows 桌面工具与核心自动化脚本套件，用于处理多语种三列对照（中文 / 英文 / 外文）Word 文档的**零散改点一键定稿、五级总库检索、外文高亮对齐（AI 精准 / 本地估算双模兜底）、批注无损保留、词数统计与弟兄交通整理**。

---

## 目录结构

```text
TranslationTools/
├── .agents/skills/scattered-revision/   # 核心 OpenXML 解析、增量总库索引与排版保真脚本
│   ├── SKILL.md                         # 业务规则与工作流规范说明
│   ├── resources/config.json            # 默认颜色、阈值与基础语种映射配置
│   └── scripts/                         # 纯标准库 Word XML 处理引擎（零第三方运行时依赖）
│       ├── docx_xml_utils.py            # OpenXML 解析、批注无损迁移与 OOXML 合规校验
│       ├── library_index.py             # 多线程持久化增量总库索引（支持无标题跨篇极速检索）
│       ├── step1_extract_and_diff.py    # 五级级联总库检索、浅蓝底色修订提取与未匹配行优雅跳过
│       ├── step1_render_draft.py        # 排版保真与第 3 列总库段落/批注无损合入
│       ├── step2_parse_reviewed.py      # 第 2 列黄(UP)/青(RV)区间与上下文提取
│       └── step2_render_final.py        # 第 3 列外文标色、词数统计与三版本定稿导出
└── 零散改点工具/                          # 桌面可视化软件 (Tkinter GUI · v1.6.0)
    ├── app_gui.py                       # 一键定稿界面、后台自动建库、多语种与无限 Key 池管理
    ├── ai_client.py                     # 多线程 AI 评估/对齐、本地位置估算兜底与简洁中文错误转换
    ├── pipeline.py                      # 一键输出定稿流水线封装（临时目录自动清理）
    ├── lang_registry.py                 # 多语种自由扩展注册表（预置 35 种语言，支持自定义新增）
    ├── word_counter.py                  # 第2列英文单词统计(UP黄/RV青)与标准文件名新增/更新模块
    ├── fellowship_processor.py          # 场景四：弟兄交通拆框清痕、蓝粗还原、四号/小四/五号排版与合并引擎
    ├── win_dnd.py                       # 原生 Win32 文件与文件夹全局拖拽引擎（纯标准库 ctypes）
    ├── security_utils.py                # Windows DPAPI 加密存储、脱敏日志、ZipBomb/XXE/SSRF 防护
    ├── languages.json                   # 外置多语种配置文件
    ├── .env.example                     # 环境变量配置示例
    ├── requirements-build.txt           # 打包独立 .exe 所需的构建依赖
    ├── 启动工具.bat                      # 脚本直连启动入口
    └── 使用说明.md                       # 详细操作手册 (v1.6.0)
```

---

## 核心特性 (v1.6.0)

1. **⚡ 极简「一键输出定稿」工作流**：
   - 精简操作界面，只需选择待整理文档、语种总库与输出目录，点击 **「⚡ 一键输出定稿」** 即可自动完成总库检索、黄（UP）/ 青（RV）分级、第 3 列外文高亮映射及词数统计，直接导出**发译者版**、**完整留底版**与**无改点内容版**（中间临时文件自动清理，输出目录干净整洁）。
2. **📚 后台自动总库索引 + 全链路多线程并发**：
   - 启动软件、浏览或拖入总库文件夹时，后台守护线程自动建立/增量更新持久化总库索引（按文件修改时间与大小增量解析，二次加载 `<0.2s`）；
   - 总库扫描与 AI 分级/外文对齐均采用多线程并发处理，大幅缩短长篇文档等待时间。
3. **🔍 五级级联智能检索 + 未匹配行优雅跳过**：
   - 支持**无文章标题**、**标题不规范/不匹配**或多篇混排表格，按 **① 标题匹配文件 → ② 全段与开头前几句 → ③ 带底色/高亮/彩色字体的句子片段 → ④ 段落结尾几句 → ⑤ 第 2 列英文词重叠加权校验** 五级级联回退检索；
   - 若个别段落确不在总库中，自动优雅跳过（绝不抛错中断）：**发译者版**与**无改点版**自动剔除跳过行，**完整留底版**保留原中英文并在第 3 列以红字标注提示。
4. **💬 Word 批注无损迁移 & 100% 合规 OOXML（直接 `Ctrl+S` 保存）**：
   - 从总库复制第 3 列外文时完整迁移原批注内容并重分配全局唯一 `w:id`，严格闭合批注起止标记、清理越界 `w14:paraId` 与悬空扩展部件；在 Word 中打开绝不弹出“发现无法读取的内容”，可直接按 `Ctrl+S` 原地保存，无需另存为。
5. **🛡️ 离线 / API Key 失效自动本地位置估算标色 + 简洁中文提示**：
   - 内置「使用 AI 辅助分级与第 3 列精准对齐」开关：未填 API Key、取消勾选 AI，或运行途中全部 Key 失效/超额/断网时，**绝不弹窗阻断或崩溃**，自动切换**本地位置估算模式**按句号/分句比例对齐整词边界完成第 3 列标色（黄→黄、青→浅灰）并准确统计第 2 列词数；
   - 所有接口异常统一转换为单行清晰中文说明，彻底移除原始 JSON 与代码堆栈刷屏。
6. **原稿排版 100% 保真 + 无限 Key 池轮换 + 35+ 语种自由扩展 +「📚 弟兄交通整理」独立工作台**。

---

## 快速运行与打包

### 0. 免安装直接下载（GitHub 自动构建 Windows 版）

直接前往本仓库 [**Releases 页面**](https://github.com/zcnz2026/TranslationTools/releases)（或 [**Actions 构建页**](https://github.com/zcnz2026/TranslationTools/actions)）下载 `ScatteredRevisionTool-Windows-x64.zip`，解压后双击 `零散改点整理工具.exe` 即可运行，无需安装 Python。

### 1. 直接源码运行（仅需 Python 3.10+ 标准库，无需 `pip install` 任何运行时依赖）

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
  ".\零散改点工具\app_gui.py"
```
