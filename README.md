# interview-minutes（访谈纪要工作台）

一个 Claude Code skill：把语音转写稿、现场手记、要点草稿等原始素材，整理成公司版式的《XX访谈纪要》 docx。

## 它做什么

- **固定结构**：元信息块 → 项目简介 → 中文序号分节正文（节内用加粗引导短语）→ 带提问方归属的 Q&A
- **固定版式**：宋体 12pt、标题 16pt 居中加粗、行距 1.15，由脚本统一渲染，不用手工调
- **忠实记录**：素材里没有的数字/人名/职务/机构一律不编造，缺的标「（待补）」；不写「谁谁认为」；信息只出现一遍；拒答与敏感事项如实记录
- **轻度核查**：对市场规模、融资额、估值、专利/论文数量、公开头衔做联网复核，存疑处只括注、不改写原话

## 目录结构

```
interview-minutes/
├── SKILL.md                  # skill 主入口：产物、约束、8 步工作流、自检清单
├── README.md
├── references/
│   ├── outline-spec.md       # 骨架规范：必备五节、备选节清单、子节用法
│   ├── writing-spec.md       # 行文规范：引导短语、人名/机构/数字、清洗、核查
│   └── qa-spec.md            # Q&A 规范：收录标准、编号归属、拒答处理
└── assets/
    ├── render_minutes.py     # 把 minutes_data.json 渲染成 docx
    └── sample_data.json      # 可直接运行的完整样例（已脱敏）
```

## 安装

把整个目录放到 Claude Code 的 skills 目录下即可：

```
~/.claude/skills/interview-minutes/
```

## 依赖

- [python-docx](https://python-docx.readthedocs.io/)（渲染脚本用）
- Python 3

## 用法

触发后按 skill 的 8 步工作流走：素材盘点 → 元信息确认 → 骨架确认 → 逐节撰写 → Q&A 整理 → 轻度核查 → 渲染 docx → 自检。

渲染命令：

```bash
PYTHONIOENCODING=utf-8 python assets/render_minutes.py \
  --data minutes_data.json \
  --out "XX访谈纪要20260920.docx"
```

> Windows 下必须带 `PYTHONIOENCODING=utf-8`，否则中文报错。

## 说明

- 仓库内的样例数据、示例标的公司与人名均已**脱敏**，仅作结构演示用。
- 骨架由多份公司纪要样本的交集分析得出：公司概况与定位 / 创始团队 / 股权结构与融资历程 / 技术路线与产品 / Q&A 五节必备，其余十余节按素材备选。
