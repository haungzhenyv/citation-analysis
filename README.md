# citation-analysis

一个 [Claude Code](https://claude.com/claude-code) **技能（Skill）**：**论文被引分析**。

给定一篇论文（DOI 或标题），自动完成：

1. 在 **Web of Science** 查被引总数与引用清单
2. 按 **中科院分区表（最新版）· 大类学科 1 区** 筛选
3. 逐条**溯源引用原句**、核查作者单位/职称/期刊编委
4. 撰写**中文介绍**
5. 排版成 **PDF**（黑体蓝标题 + 宋体正文）

## 安装

把本仓库放进 Claude Code 的技能目录：

```bash
git clone https://github.com/<你的用户名>/citation-analysis.git ~/.claude/skills/citation-analysis
```

- 个人级（全局可用）：`~/.claude/skills/citation-analysis/`
- 项目级：`<项目>/.claude/skills/citation-analysis/`

## 使用

```
/citation-analysis <DOI 或标题>
```

或直接用自然语言描述需求（如「帮我做某篇论文的被引分析」），Claude 会自动调用本技能。

## 目录结构

```
citation-analysis/
├── SKILL.md                  # 技能说明（7 步流程）
└── scripts/
    └── make_intro_pdf.py     # 把「引用示例」中文介绍 md 排成规范 PDF
```

## 依赖

- **Python**：`python-docx`、`pywin32`
- **Windows + 已安装 Microsoft Word**（`make_intro_pdf.py` 用 Word COM 导出 PDF）

```bash
pip install python-docx pywin32
```

## 脚本用法

```bash
python scripts/make_intro_pdf.py 输入.md 输出.pdf --title "文档总标题"
```

- 输入 md：每个条目以 `## 引用示例 N` 开头，其下为正文段落。
- 排版规范：总标题黑体 16pt 加粗蓝色 `#1F4E79` 居中；条目标题「引用示例 N」黑体 14pt 加粗蓝色居中；正文宋体 12pt、两端对齐、首行缩进 2 字、行距 1.12。

## 说明

- 分区筛选**始终使用中科院分区表最新版**（截至 2026-10 为「2025年3月升级版」）；动手前请确认是否有更新版本。
- 编委信息**查不到就留空，禁止编造**。
- 本技能不包含任何账号、密钥等个人信息；使用 WoS 请用你自己的机构账号登录。

## 许可

[MIT](LICENSE)
