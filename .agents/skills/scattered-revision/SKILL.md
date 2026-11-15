---
name: scattered-revision
description: >-
  处理多语种三列表格 Word 文档的“零散改点（场景一）”两步走工作流：从各语种《七卷书》总库自动提取匹配段落填入第3列、按浅蓝底色过滤生成英文原生修订、AI 结合精细边界规则评估黄（UP 局部改点）/青（RV 整句重翻审核）生成初筛稿供人工复核，并在确认后全程保留第2列黄/青两色修订痕迹（不接受青色修订）、将第2列黄色与青色同步精准映射到第3列外文高亮（黄对黄、青对浅灰，纯标点改点同步标约3词，不额外拆段换行）、统计第2列词数并按规范命名输出定稿。当用户要求处理“零散改点”、“英文底稿对比提取”、“改点标色算字数”或继续零散改点第二步时使用此技能。
---

# 零散改点自动化工作流（两步走带人工复核）

本技能通过纯底层 OpenXML 脚本与当前 Agent 语义能力配合，无卡顿完成三语 Word 表格的“零散改点”全流程。默认采用**两步走（方案 A）**，在阶段一完成后暂停交由人工在 Word 中复核第 2 列标色，确认后再执行阶段二生成最终交付稿。

配置文件位于 [config.json](./resources/config.json)（可配置语种代码映射、高亮色值、是否接受青色修订 `accept_col2_rv_revisions`［默认 `false`］、阈值及可选顶部说明模板）。

---

## 阶段一：生成《初筛对比稿》（精细标色 + 保留全部修订供人工复核）

当用户提供《英文底稿.docx》及对应语种总库文件夹时执行阶段一：

### 1. 提取总库匹配行并计算句子级差异
运行 [step1_extract_and_diff.py](./scripts/step1_extract_and_diff.py)：
```powershell
python -X utf8 "<skill_dir>/scripts/step1_extract_and_diff.py" `
  --input "<英文底稿路径>" `
  --library "<对应语种总库目录>" `
  --workdir "<临时工作目录>"
```
脚本会自动：
- 根据表格上方的篇名在总库中定位 `.docx`（自动跳过 `~$` 临时文件及 `标记结果`/`答案`/`初筛` 目录），并按第 1 列中文精确/高置信度匹配对应行；
- 识别第 2 列字符底纹（如 `DEEAF6` 浅蓝底色）：无底色未引用区域自动回退为总库旧版英文，有底色引用区域按词/标点生成精细差异流并切分为同步完整句。

### 2. Agent 语义评估黄（UP）/ 青（RV）句子分类
读取 `<临时工作目录>/step1_review_items.json`，逐句审查每个有改动的句子，结合以下标准判断该句应归为 `"yellow"`（局部改点）还是 `"cyan"`（重翻审核）：
- **`"yellow"`（局部改点 UP）**：仅替换个别实词/短语、单复数、连词或标点微调，译者对照修订痕迹能轻松在原外文句上局部修改。即便连续多句都有局部小改点，只要每句改动较轻（非整段重写），仍保持 `"yellow"`（后续渲染脚本会自动按防碎片规则合并中间极短未标间隙）。
- **`"cyan"`（连续密集大改 RV）**：整句词修改率高（通常 >40%）、句式结构重组，或连续多句修订穿插极其密集（“很花”），译者难以在原句修补、整段重翻审核性价比更高。**硬性铁律**：所有 `"cyan"` 必须以完整句子标点（`.`、`?`、`!`）为起止边界。

将评估结果写入 `<临时工作目录>/step1_decisions.json`，格式为 `{ "<sent_id>": "yellow" | "cyan" | "none", ... }`。

### 3. 渲染输出《初筛对比稿》（内置 5 条精细边界规则）
运行 [step1_render_draft.py](./scripts/step1_render_draft.py)：
```powershell
python -X utf8 "<skill_dir>/scripts/step1_render_draft.py" `
  --manifest "<临时工作目录>/diff_manifest.json" `
  --decisions "<临时工作目录>/step1_decisions.json" `
  --output "<输出目录>/初筛_<原文件名>.docx"
```
渲染脚本与差异分词器会自动执行以下精细规则：
0. **人称单复数角标消歧自动过滤（`you` ⇄ `you(s)`/`you(pl)`）**：
   - 若仅在同词根代词（`you`、`your`、`yours`，同大小写）上**增补或去掉** `(s)` / `(pl)` 角标（如 `you` → `you(pl)` 或 `you(s)` → `you`），**不算改点，直接接受修订**为最新文本，不保留修订痕迹、不标黄；
   - 若发生单复数**互改**（如 `You(s)` ⇄ `you(pl)`、`your(s)` ⇄ `your(pl)`），属于实质单复数改点，正常保留修订并标黄。
1. **纯标点/引号/断句大小写微调 → 3 词窗口**：若某处修订仅涉及标点符号或断句引起的首字母大小写（英文单词本身未变，如 `this {+“+}love{+”+} that`），仅标记改动点前后共 **3 个单词**为黄色高亮，不扩大到整分句。
2. **长分句单处局部改动 → 连词意群切分**：超过 12 词的长分句中若仅一侧有局部改点，在语义完整的并列/从属连词（`and`、`or`、`but`、`yet`、`while`、`nor`，两侧均 $\ge 5$ 词且不拆散 `of A and B` / `both A and B` 短语）处断开，仅标含改点的半句；短分句（$\le 12$ 词）直接标记。
3. **冒号后起标 & 括号内独立标记**：
   - 冒号/分号之前的文字无变化时（含新插入冒号如 `flower{+:+} {-that-}{+Flowers+}`），从冒号**之后**开始标黄，不连带前半句；
   - 括号 `(...)` 内部的局部修改（如 `(including the {-moment-}{+time+} of His crucifixion)`），仅标记括号内部，不波及括号外文字；
   - 若改点簇新文本自身以分句/句末标点收尾（如 `{+. Thus+},{- in which case-}`），紧贴该改点簇停笔，不向后拖入无改动正文。
4. **防碎片化黄色合并（兼顾省钱与防琐碎，两端不扩）**：
   - 极短碎间隙（$\le 4$ 词，如 `"“God,"`、`"I swear to You(s):"`）或句内/夹心短句间隙（$\le 10$ 词，如 `sin does not represent God, and God is sinless.`，两侧黄色均 $\ge 4$ 词且总和 $\ge 3$ 倍间隙词数）被黄色包夹时自动合并；
   - **绝不跨句末标点（`.?!`）将下一句开头无改动的逗号/分号分句（如 `As long as we possess hearts that truly want God,`）连带标黄**，且绝不向外扩张首尾两端。
5. **全程保留黄/青两色修订痕迹**：黄色与青色区域均完整保留原生 `<w:del>`/`<w:ins>` 修订痕迹，不接受修订。

- **暂停点**：向用户汇报匹配篇目、预估 `up`/`rv` 词数及初筛稿路径，提示用户在 Word 中检查第 2 列黄/青标色（如需调整可直接改高亮颜色保存），确认后告知“继续第二步”。

---

## 阶段二：定稿生成（保留修订痕迹 + 第 3 列黄/灰双向精准映射 + 算词命名）

当用户确认初筛稿（或提供人工微调后的初筛稿路径）并要求继续第二步时执行：

### 1. 解析确认后的第 2 列高亮区间
运行 [step2_parse_reviewed.py](./scripts/step2_parse_reviewed.py)（直接从 `.docx` XML 读取，100% 尊重人工在 Word 中做的任何改色或范围微调）：
```powershell
python -X utf8 "<skill_dir>/scripts/step2_parse_reviewed.py" `
  --draft "<确认后的初筛稿路径>" `
  --workdir "<临时工作目录>"
```

### 2. Agent 跨语种精准定位第 3 列高亮子串
读取 `<临时工作目录>/step2_align_tasks.json`。对每一行的 `spans` 列表（按原文顺序排列）：
- 注意第 3 列外文是**旧版英文（`old_en_in_col3`）**的对应译文！请结合 `left_old_context`、`old_en_in_col3`、`right_old_context` 在 `col3_text` 中找到**一模一样的原文连续子串**：
  - 当 `type == "rv"`（对应第 2 列青色整句/整段）：提取 `col3_text` 中对应的**完整外文句子区间**（将渲染为浅灰 `lightGray` 高亮，**不额外拆段换行**）。
  - 当 `type == "up"`（对应第 2 列黄色区间）：提取 `col3_text` 中与该英文区间精准对应的外文片段（将渲染为黄色 `yellow` 高亮）：
    - 若第 2 列是**纯标点 3 词窗口**（如 `"this love that"`），第 3 列也同步提取对应的**约 3 个外文词**（如法语 `"cet amour que"`）；
    - 若第 2 列是**冒号后起标**或**括号内起标**，第 3 列也同步从对应分句/括号处起标；
    - 若第 2 列经过防碎片合并跨越了中间短句，第 3 列也同步覆盖对应的连续外文区间。
- 将映射表写入 `<临时工作目录>/step2_alignments.json`，格式为：
  ```json
  {
    "R0_S0": "alors il te sera impossible de faire le travail que tu devrais faire pour Lui.",
    "R2_S1": "cet amour que"
  }
  ```

### 3. 生成最终定稿并校验
运行 [step2_render_final.py](./scripts/step2_render_final.py)：
```powershell
python -X utf8 "<skill_dir>/scripts/step2_render_final.py" `
  --tasks "<临时工作目录>/step2_align_tasks.json" `
  --alignments "<临时工作目录>/step2_alignments.json" `
  --outdir "<输出目录>" `
  --lang "<语种代码如FR/PT/GR>" `
  --doc-name "<稿件类型或篇名>" `
  --date "<MMDD可选>" `
  [--remove-unchanged-tables]
```
> 若批次中存在外文总库已是最新英文、第 2 列完全无改点的表格，可加 `--remove-unchanged-tables` 输出剔除无改点表格的“发译者版”，不加该参数则输出保留全部表格的“完整留底版”。
脚本会自动执行并校验：
1. **最终定稿同样保留青色（RV）与黄色（UP）区域的 `<w:del>`/`<w:ins>` 修订痕迹**（由 `config.json` 中 `"accept_col2_rv_revisions": false` 控制，仅归一化人工可能涂成 `blue`/`darkCyan` 的色值，不接受青色修订）；
2. 将第 3 列外文中对应的 UP 片段标 `yellow`、RV 整句标 `lightGray`（不插入多余换行，严格保持 OpenXML `w:rPr` 标签顺序，自动兼容法语不间断空格 `\xa0` 与弯引号匹配）；若有任何 1 处子串未命中会报错退出（返回码 `2`），修正子串重跑即可确保 100% 命中；
3. 默认保留文档顶部空行（不自动插入说明文字，与标准测试集一致）；
4. **仅统计第 2 列英文节点**的接受态有效词数（正常 `<w:t>` + `<w:ins>` 内 `<w:t>`，自动排除 `<w:delText>`，第 3 列标色绝不影响计费），按 `[语种]_[篇名]_(upX rvY)_[MMDD].docx` 缺省规则命名输出。
