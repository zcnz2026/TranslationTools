# TranslationTools — 多语种三列改点与分part稿自动整理工具 (v1.8.0)

轻量级 Windows 桌面可视化软件与纯标准库 OpenXML 排版引擎套件，专为多语种三列对照（中文 / 英文 / 外文）Word 文档打造，内置**五大专业工作台 + 独立全局配置中心**：**零散改点整理**、**全篇改点整理**、**改点稿查重**、**英文字数统计**、**分part稿整理/合并**与**系统设置**，并支持**软件内一键在线检查与热更新**。

---

## 目录结构

```text
TranslationTools/
├── .agents/skills/scattered-revision/   # 核心 OpenXML 解析、增量总库索引与排版保真脚本
│   ├── SKILL.md                         # 业务规则与工作流规范说明
│   ├── resources/config.json            # 默认颜色、阈值与基础语种映射配置
│   └── scripts/                         # 纯标准库 Word XML 处理引擎（零第三方运行时依赖）
│       ├── docx_xml_utils.py            # OOXML 合规清理、中文修订接受、逐字符颜色保留、批注迁移、三语拉齐
│       ├── library_index.py             # 多线程持久化增量总库索引（实时 mtime 校验 + 支持强制刷新）
│       ├── step1_extract_and_diff.py    # 五级级联检索、总库实时读取、支持浅蓝底色局部对比与全篇强制对比
│       ├── step1_render_draft.py        # 逐字符内联颜色保真、黄标间隔连片、提取三语原稿与批注保留
│       ├── step2_parse_reviewed.py      # 第 1/2/3 列段落文本与黄(UP)/青(RV)区间坐标提取
│       └── step2_render_final.py        # 第 1/3 列同步标色、无改点行剔除、0词省略命名与三版定稿导出
└── 零散改点工具/                          # 桌面可视化软件 (Tkinter GUI · v1.8.0)
    ├── app_gui.py                       # 左侧六大导航（含系统设置）、常用/最近语种置顶、消歧弹窗与自动更新界面
    ├── ai_client.py                     # 单请求同步中/外双列对齐、严格 RV 判定提示词与离线本地对齐兜底
    ├── master_sync.py                   # 待处理稿核对最新中英底稿引擎（100% 保留原稿底纹/颜色/字体/批注仅无损更新文字）
    ├── pipeline.py                      # 零散改点一键输出定稿流水线（支持底稿核对、三语原稿导出、标色开关与联动查重）
    ├── full_article_processor.py        # 全篇改点整理引擎（按标题提取 + 双文件夹直连对比、前缀保留、无改点报告与智能分包）
    ├── dedup_checker.py                 # 未入库改点包第 1 列中文四级查重、明细报告生成与自动去重重算引擎
    ├── word_counter.py                  # 第 2 列英文单词统计（3 种命名模式、0 词自动省略、同步更新顶部说明行）
    ├── fellowship_processor.py          # 分part稿整理/合并引擎（多文件/单文件多Part合并、蓝/红/粗细同步、语系字体）
    ├── lang_registry.py                 # 多语种常用置顶(★)/最近使用(⏱)排序与《各语系字体》标准字号注册表
    ├── updater.py                       # GitHub Releases 自动检测更新、SHA-256 校验与保留配置无感热替换重启
    ├── win_dnd.py                       # 原生 Win32 文件与文件夹全局拖拽引擎（纯标准库 ctypes）
    ├── security_utils.py                # Windows DPAPI 加密存储、每次运行独立临时工作目录、脱敏日志与安全防护
    ├── languages.json                   # 外置多语种、常用置顶与标准字体配置文件
    ├── .env.example                     # 环境变量配置示例
    ├── requirements-build.txt           # 打包独立 .exe 所需的构建依赖
    ├── 启动工具.bat                      # 脚本直连启动入口
    └── 使用说明.md                       # 详细操作手册 (v1.8.0)
```

---

## 核心特性 (v1.8.0)

1. **「系统设置」独立全局配置中心 & 常用语种置顶**：
   - 左侧「全局配置」分组新增 **「系统设置」** 板块，集中管理 API 密钥池（内嵌多行编辑器与批量检测弹窗实时同步）、模型与接口通道切换、**两处黄色高亮自动连片最大间隔词数（默认 `3`，支持 `0~20` 自由调节）**与常用语种置顶概览；顶部状态栏精炼展示 `模型 | Key 数量 | 黄标连片阈值`。
   - 所有板块的目标语种下拉框支持 **`★ 常用` 置顶 + `⏱ 最近使用` 自动排前**，高频语种一键直达。

2. **「零散改点整理」底稿无损核对与三列同步精排**：
   - **核对最新中英底稿**：支持选择或拖入最新中英底稿 `.docx` 文件或文件夹，在 **100% 保留原始待处理稿的浅蓝/粉红底纹、蓝/红字体颜色、字号与批注** 的前提下无损同步最新中英文字，自动生成 `{原名}{YYYYMMDD}更新.docx` 并无缝衔接提取对比。
   - **中/英/外三列同步标色 + 原稿内联颜色逐字符保真**：自动接受第 1 列中文残留修订并同步标黄（UP）/ 青（RV）；逐字符保留第 2 列圣经引文蓝字（`#0000FF`）与红字不串色；发译者版彻底剔除同表无改点行。
   - **对比标色开关 & 提取三语原稿核对**：支持取消勾选「对比标记高亮颜色」仅输出修订痕迹；默认同步导出对比前的「提取三语原稿」供核对；单文档输出直接平铺在输出目录（不套多余子文件夹）。

3. **「全篇改点整理」双模式对比与智能归档**：
   - **按标题列表提取 + 现成双文件夹直连对比**：既支持粘贴/导入带双编号前缀与括号尾注的标题（如 `3-1-39 032 七雷巨响——预言国度的福音将扩展全宇（审判）`，自动匹配并在输出文件名中保留 `3-1-39 032_` 前缀），也支持取消勾选标题模式直接将「最新中英底稿包」与「外文三语文件包」按顺序逐篇对比并以中英底稿命名。
   - **英文无改点篇目汇总报告 & 智能分包**：若某篇英文对比无改点，自动跳过导出单篇对比稿并统一汇总至 **`全篇对比_英文无改点篇目汇总报告_YYYYMMDD.docx`**；单篇处理直接平铺输出，多篇（`≥2` 篇）自动按 `给译者版/`、`完整留底版/`、`中英最新底稿包/`、`外文三语原稿包/` 四大子目录整洁归档。

4. **标色边界精细规则升级**：
   - 句号改冒号（`flower.` → `flower: Flowers...`）仅从冒号起向后标黄；半句局部改写（如 `as far as man is concerned` → `when it comes to man`）严格判为黄色 UP 不误扩青色；零碎全段大改（仅剩 1~2 句未动）自动升级整段青色 RV；两处黄色高亮间隔词数 `≤N`（默认 3 词）自动连片标色。

5. **「改点稿查重」、「英文字数统计」与「分part稿整理/合并」全套工具链 + 在线热更新**。

---

## 快速运行与打包

### 0. 免安装直接下载 & 软件内自动更新

直接前往本仓库 [**Releases 页面**](https://github.com/zcnz2026/TranslationTools/releases) 下载 `ScatteredRevisionTool-Windows-x64.zip`，解压后双击 `零散改点整理工具.exe` 即可运行。后续版本可直接点击软件左下角 **「检查软件更新」** 一键完成 SHA-256 校验下载与自动替换重启（本机 API Key 与语种配置完整保留）。

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
  --hidden-import master_sync `
  --hidden-import full_article_processor `
  --hidden-import dedup_checker `
  --hidden-import updater `
  ".\零散改点工具\app_gui.py"
```
