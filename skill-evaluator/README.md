# C4 Skill Submission Evaluator（C4A 技能提交自动评审器）

> 把「WeChat 群文件夹」变成一份结构化的 C4 提交评审报告。
> C4A 挑战交付技能包：由 WangChunYan 开发。

## 一句话

输入一个本地文件夹路径，输出每位作者的完整性检查、四条件质量评分、综合排名与改进建议。

## 目录结构

```
skill-evaluator/
├── SKILL.md                  # 技能定义（frontmatter + 工作流 + 边界情况）
├── c4_skill_evaluator.py     # 核心实现（扫描→完整性→质量评审→报告）
├── requirements.txt          # 可选依赖（pypdf）
└── references/
    ├── c4_rubric.json        # 机器可读评分表（运行时加载）
    └── c4_rubric.yaml        # 人类可读评分表（同语义，供编辑）
```

## 快速开始

```bash
# 1. 进入技能目录
cd skill-evaluator

# 2. 运行（零依赖，Python 3.8+ 即可）
python c4_skill_evaluator.py <你的C4提交文件夹路径>

# 3. 查看输出
#    review_output/评审报告.md   — Markdown 报告
#    review_output/评审详表.csv  — Excel 兼容详表
```

## 四层级覆盖

| 层级 | 能力 | 实现位置 |
|------|------|----------|
| Level 1 | 文件采集与作者识别 | `scan_folder()` / `extract_author()` |
| Level 2 | 完整性检查（5 文件） | `check_completeness()` |
| Level 3 | 质量评审（四条件） | `evaluate_quality()` |
| Level 4 | 完整报告（排名/建议/版本追踪/HTML） | `render_markdown/csv/html()` |

## 设计要点

- **零依赖**：核心脚本仅用 Python 标准库，助教无需 pip install 即可复用。
- **规则 + 可选 LLM 混合**：规则层确定、可离线；`--llm` 启用 LLM 定性深审。
- **可配置评分表**：权重、信号、检查项都放在 `references/c4_rubric.json`，可改可换。
- **版本追踪**：自动识别 `_v2`/`_v3`，展示迭代轨迹。

详见 `../WangChunYan_C4A_方案设计.md` 与 `../WangChunYan_C4A_拿来说明.md`。
