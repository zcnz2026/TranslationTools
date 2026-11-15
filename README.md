# TranslationTools — 多语种三列表格零散改点自动整理工具

轻量级 Windows 桌面工具与核心自动化脚本套件，用于处理多语种三列对照（中文 / 英文 / 外文）Word 文档的**零散改点提取、黄(UP)/青(RV)分级初筛、外文高亮同步对齐与词数统计定稿**。

---

## 目录结构

```text
TranslationTools/
├── .agents/skills/scattered-revision/   # 核心 OpenXML 解析、细粒度修订对比与排版保真脚本
│   ├── SKILL.md                         # 业务规则与工作流规范说明
│   ├── resources/config.json            # 默认颜色、阈值与基础语种映射配置
│   └── scripts/                         # 纯标准库 Word XML 处理引擎（零第三方运行时依赖）
│       ├── docx_xml_utils.py
│       ├── step1_extract_and_diff.py
│       ├── step1_render_draft.py
│       ├── step2_parse_reviewed.py
│       └── step2_render_final.py
└── 零散改点工具/                          # 桌面可视化软件 (Tkinter GUI)
    ├── app_gui.py                       # 主界面、多语种管理器、无限多 Key 密钥池管理器
    ├── ai_client.py                     # Gemini / Vertex AI 客户端与多 Key 自动轮换秒切引擎
    ├── pipeline.py                      # 两步走人工复核与一键直出流水线封装
    ├── lang_registry.py                 # 多语种自由扩展注册表（预置 35 种语言，支持自定义新增）
    ├── word_counter.py                  # 第2列英文单词统计(UP黄/RV青)与标准文件名新增/更新模块
    ├── win_dnd.py                       # 原生 Win32 文件与文件夹全局拖拽引擎（纯标准库 ctypes）
    ├── security_utils.py                # Windows DPAPI 加密存储、脱敏日志、ZipBomb/XXE/SSRF 防护
    ├── languages.json                   # 外置多语种配置文件
    ├── .env.example                     # 环境变量配置示例
    ├── requirements-build.txt           # 打包独立 .exe 所需的构建依赖
    ├── 启动工具.bat                      # 脚本直连启动入口
    └── 使用说明.md                       # 详细操作手册
```

---

## 核心特性

1. **原稿字体与排版 100% 保真**：深度操作 Word OpenXML（`w:rPr` / `w:pPr`），完整保留原文档字体（如 `Georgia`）、字号（如 `12pt`）、段落缩进、从右向左（RTL）及复杂文种脚本属性。
2. **无限多 API Key 轮换 & 报错自动秒切**：
   - 支持混填 **Vertex AI API Key**（`AQ.` 开头）与 **Gemini Developer API Key**（`AIza` 开头），数量无上限；
   - 支持 **均衡轮换** 与 **顺序主备** 两种策略，遇到额度耗尽 / 超速（`HTTP 429`）、密钥失效（`401/403`）或网络波动时毫秒级自动切换下一个可用 Key，不中断批处理任务；
   - 支持一键批量测活与剔除失效 Key；本机勾选记住时采用 **Windows DPAPI** 系统级加密存储，磁盘与日志零明文泄露。
3. **支持多语种自由扩展**：内置 35 种常见语言，支持在界面点击「➕ 管理 / 新增语种」随时新增任意语种（实时同步至 `languages.json`）。
4. **双工作流支持**：支持「第一步生成初筛稿人工复核 → 第二步生成定稿」或「⚡ 一键全自动直出全部定稿」，可同时导出**完整留底版**与**发译者版（自动剔除无改点表）**。

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
  --hidden-import step1_extract_and_diff `
  --hidden-import step1_render_draft `
  --hidden-import step2_parse_reviewed `
  --hidden-import step2_render_final `
  --hidden-import ai_client `
  --hidden-import lang_registry `
  --hidden-import security_utils `
  --hidden-import word_counter `
  --hidden-import win_dnd `
  --hidden-import pipeline `
  ".\零散改点工具\app_gui.py"
```
