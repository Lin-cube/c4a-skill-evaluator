# WangChunYan_C4A_教学说明

> 怎么安装、怎么用、输入什么得到什么 —— C4A 技能提交自动评审器的上手教程。

---

## 1. 这是什么

一个「评审技能」的技能：给它一个装满 C4 提交的本地文件夹，它自动产出每位作者的
完整性检查、四条件质量评分、综合排名与改进建议。

- **输入**：一个本地文件夹路径（WeChat 同步文件夹 / 手动下载文件夹 / 任何含群文件的目录）
- **输出**：`评审报告.md`（人类读）+ `评审详表.csv`（Excel 打开）+ 可选 `评审报告.html`（可视化）

---

## 2. 安装（30 秒）

### 前提

- Python 3.8+（本技能核心**零第三方依赖**，不用 pip install 就能跑）

### 两种方式

**方式 A：直接用源码（推荐，零安装）**

把 `skill-evaluator/` 目录拷到任何地方，直接运行里面的脚本即可。

**方式 B：安装技能包**

解压 `WangChunYan_C4A_skill-evaluator.skill`（本质是 zip），得到 `skill-evaluator/` 目录。

### 可选依赖（需要时才装）

```bash
pip install pypdf      # 想让 PDF 也被解析正文（不装则 PDF 仅按文件名匹配）
```

---

## 3. 使用（3 分钟上手）

### 第一步：准备文件夹

把大家的 C4 提交文件放到一个文件夹（可含子目录），文件名遵循 `姓名拼音_C4_内容描述.扩展名`。

### 第二步：跑命令

```bash
cd skill-evaluator

# 最基本用法：扫描并生成 Markdown + CSV 报告到 review_output/
python c4_skill_evaluator.py "D:\你的C4提交文件夹"

# 指定输出目录 + 三种格式
python c4_skill_evaluator.py "D:\你的C4提交文件夹" --output report --format md,csv,html

# 用自定义评分表（改权重/信号）
python c4_skill_evaluator.py "D:\你的C4提交文件夹" --rubric references\c4_rubric.json

# 启用可选 LLM 深审（需先设置 OPENAI_API_KEY）
python c4_skill_evaluator.py "D:\你的C4提交文件夹" --llm
```

### 第三步：看结果

| 文件 | 内容 | 谁读 |
|------|------|------|
| `review_output/评审报告.md` | 总览 + 作者详情 + 排名 + 改进建议 | 老师/助教 |
| `review_output/评审详表.csv` | 一行一人的打分明细 | Excel 分析 |
| `review_output/评审报告.html` | 排名表网页版 | 投屏展示 |

---

## 4. 常见坑

| 坑 | 现象 | 解法 |
|----|------|------|
| 提示"未找到 C4 相关文件" | 文件名没带 `_C4_` 标记 | 按命名规范重命名，或确认路径正确 |
| 控制台打印乱码 | Windows GBK 控制台 | 已在脚本里强制 UTF-8，仍乱码可 `chcp 65001` |
| PDF 只按文件名匹配 | 没装 pypdf | `pip install pypdf` 后重跑 |
| 想改评分权重 | 默认 0.4/0.6 | 编辑 `references/c4_rubric.json` 的 `scoring` 节点 |
| 中文文件名 | — | 已 UTF-8 优先、GBK 回退，无需处理 |

---

## 5. 用真实群文件怎么跑（对应 CHALLENGE 的开放题）

文件怎么从 WeChat 群到本地，这一步留给你自己，本技能只要求「一个本地文件夹路径」。
三个可行路径：

1. **手动下载**：群文件逐个保存到一个文件夹（最简单）
2. **WeChat 桌面端同步**：找到本地缓存目录，直接指过去
3. **导出工具**：用现成群文件导出工具（拿来主义）

无论哪条路，最后都落到同一条命令：

```bash
python c4_skill_evaluator.py "<你的文件夹路径>"
```

---

## 6. 输入输出速查

| | 说明 |
|--|------|
| **输入** | 文件夹路径（含 `_C4_` 命名文件），支持 .md/.py/.skill/.docx/.pdf/.mp4/.png 等混合格式 |
| **输出** | 评审报告：完整性矩阵（✅/❌）+ 四条件评级（✅/⚠️/❌）+ 综合分 + 排名 + 个性化改进建议 |
| **副作用** | 无（只读扫描，输出到独立 `--output` 目录） |

---

## 7. 验证本技能可跑（自带测试数据）

仓库自带 3 份测试提交，一条命令即可复现评审报告：

```bash
python c4_skill_evaluator.py test_data --output review_output --format md,csv,html
```

预期结果：ZhangSan 5/5、LiSi 3/5、WangWu 2/5，综合分降序排名。
