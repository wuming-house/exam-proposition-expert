# 命题专家 · Exam Proposition Expert

一个**跨平台、可被任意 AI 智能体复用**的命题技能（Skill）。它以一位「从事教育行业 50 年、刚退休的命题教授」的身份，按学校真实的试卷命题流程，以**「命题人 / 审题人」双角色独立把关**，友好地采集考试信息，并生成**可直接打印的 Word 试卷（.docx）**。

核心保证：**试卷里的 LaTeX 公式会渲染为 Unicode 数学符号（兼容 Word、WPS、LibreOffice，可显示分数、根号、上/下标等），不是图片、绝不会乱码**，配图正确嵌入；且**未经审题人独立审查通过，绝不输出最终试卷**。

本技能遵循通用的 `SKILL.md` 规范编写，**不绑定任何单一产品**：不仅能在 WorkBuddy 中运行，也兼容 Claude Code、Cursor、Codex、通义灵码等任何支持 `SKILL.md` 的 AI 智能体框架——换句话说，**任何会读文件、能执行命令的 AI 助手都能把它装来用**。

---

## 特性

- 🎓 **专业命题流程**：依标命题 → 双向细目表 → 出题 → 试做 → 审题 → 组卷 → 制卷 → 考后分析，完整闭环。
- 👥 **命题人 / 审题人双角色把关**：命题与审题由**两个独立视角**完成（AI 切换角色模拟），审题人专挑毛病并产出《审题报告》；**审题不通过，不得定稿生成 docx**（硬门）。
- 🔍 **审题报告 + 机械校验**：附 `review_paper.py` 做客观检查（分值合计=满分、题号连续、选择题选项完整、公式可解析、图片存在），结果贴入《审题报告》。
- 💬 **友好采集**：分两轮（结构化多选 + 自然语言追问）引导你填写考试信息，不会一次性甩长表单。
- 📐 **公式不丢不乱**：LaTeX 公式经解析后渲染为 Unicode 数学符号（分数、根号、上/下标、希腊字母等），在 Word / WPS / LibreOffice 中均可清晰显示，不是图片也不会乱码。
- 🖼️ **配图正确嵌入**：`![说明](figs/xxx.png)` 引用的图片直接嵌入文档，不溢出、不丢失。
- 🖨️ **开箱即打印**：统一中文字体（宋体/黑体 + Times New Roman）、A4 版面、页边距、页码，可选左侧「密封线」。
- 📦 **零外部依赖**：完全自包含，不需要 pandoc 等外部二进制。

---

## 目录结构

```
exam-proposition-expert/
├── SKILL.md              # 技能定义（角色、双角色机制、流程、采集方式、书写约定、审题报告模板）
├── scripts/
│   ├── build_paper.py    # 试卷生成引擎：Markdown → 可打印 docx
│   └── review_paper.py   # 审题人的机械校验工具（分值/题号/选项/公式/图片）
├── demo/                 # 演示样例
│   ├── paper.md          # 一份七年级(上)数学期中卷的 Markdown 源文件（含公式与配图引用）
│   ├── make_figs.py      # 用 matplotlib 生成演示配图的脚本
│   ├── figs/             # 演示配图
│   ├── 试卷样例.docx      # 由本技能生成的成品样例（可直接打开查看效果）
│   └── 审题报告_样例.md   # 在 paper.md 上跑 review_paper.py 产出的机械校验样例
└── README.md
```

---

## 安装

### 方式一：手动安装

把整个 `exam-proposition-expert/` 文件夹复制到所用 AI 助手的「用户级 skills 目录」。不同框架的路径略有差异：

```bash
# WorkBuddy（Windows PowerShell）
Copy-Item -Recurse exam-proposition-expert "$env:USERPROFILE\.workbuddy\skills\"

# WorkBuddy / Claude Code / 其它（macOS / Linux）
cp -r exam-proposition-expert ~/.workbuddy/skills/
# Claude Code 用户请将目标目录换成 ~/.claude/skills/

# Cursor（项目级）
cp -r exam-proposition-expert .cursor/skills/
```

重启/刷新对话后，对助手说「帮我出一份 XX 学科的试卷」即可触发。

### 方式二：让 AI 智能体自动安装

如果你正在使用AI 助手（WorkBuddy、Claude、ChatGPT、Codex、通义灵码、Copilot、Gemini 等），**只需把下面这段提示词原样发给它**，它会自动帮你克隆仓库、放到正确的 skills 目录、并装好依赖——你不用手动敲任何命令：

```
帮我安装这个Skill：https://github.com/wuming-house/exam-proposition-expert
```

> 这段提示词是**框架无关**的——任何能读写文件、执行命令的 AI 助手读完都能自行完成安装，无需你了解各框架的目录约定。

### 运行依赖

生成引擎 `build_paper.py` 需要 Python 3 与以下库（无论哪个框架，装好后都一样）：

```bash
pip install python-docx latex2mathml matplotlib
```

---

## 依赖

生成引擎 `build_paper.py` 依赖：

- `python-docx`（排版与 docx 写出，**必须**）
- `latex2mathml`（LaTeX → Unicode 数学符号，**必须**）
- `matplotlib`（生成配图，按需）

安装：

```bash
pip install python-docx latex2mathml matplotlib
```

> WorkBuddy 本机已内置隔离 Python 环境（预装上述库），技能可开箱即用。

---

## 使用方法（给 Skill 调用者/开发者参考）

技能内置的标准双角色流程：

1. **命题人出题** → 写成 `paper_draft.md`。
2. **审题人独立审查** → 切换角色，按 `SKILL.md` 第五节清单逐条核对，并跑机械校验：
   ```bash
   python scripts/review_paper.py paper_draft.md -o 审题报告.md
   ```
   结论为「需修订后复审」时，回到第 1 步修订并重审，直到「通过」。
3. **审题通过后** → 定稿 `paper.md`，生成可打印 docx：
   ```bash
   python scripts/build_paper.py paper.md -o 试卷.docx --seamless
   ```

`build_paper.py` 参数：

- `-o, --output`：输出 docx 路径（必填）
- `--seamless`：添加左侧竖排「密封线」
- `--no-page-number`：取消页码

`review_paper.py` 参数：

- `-o, --output`：将机械校验报告写入文件（缺省打印到标准输出）

`paper.md` 书写约定（要点）：

- 第一行 `# 标题` 作为试卷大标题（自动居中加粗）。
- 考生信息栏用 Markdown 表格（下划线留空供手写）。
- 题型用 `## 二级标题` 并写明分值（如"每小题 3 分，共 30 分"）。
- 公式：行内 `$x^2+1$`，独立成行居中 `$$...$$`。
- 图片：`![图n：说明](figs/xxx.png)`，路径相对 `paper.md` 所在目录。

完整约定与双角色机制见 `SKILL.md`。

---

## 协议

本仓库默认采用 **MIT License**，欢迎 Fork、改进并分享给更多人使用。
（如需其它协议，可替换根目录的 `LICENSE` 文件。）

---

## 致谢

流程设计参考中小学及升学考试的真实命题规范（双向细目表、依标命题、审题分离、保密管理等）。

---

**Created by 吾鳴**
- 微信公众号：**蕪鳴**
- 微信：**wu_ming_2025**
- Built for 老师|家长 · Made with ❤️
