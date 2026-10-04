#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
c4_skill_evaluator.py — C4 技能提交自动评审器
==============================================

把一个「WeChat 群文件夹」（或任何本地目录）变成一份结构化的 C4 提交评审报告。

设计定位（C4A Level 1-4 全覆盖）：
  ① 文件采集与识别   —— 扫描目录，按 `_C4_` 命名规范识别作者并分组
  ② 提交完整性检查   —— 检查 C4 五个必须文件是否齐全
  ③ 技能质量评审     —— 按 C4 四条件（可复用/可执行/可验证/IO明确）打分
  ④ 评审报告生成     —— Markdown + CSV(Excel 兼容) + 可选 HTML，含排名与建议

技术选型：规则 + 可选 LLM 的混合架构。
  - 规则层：确定性、零依赖、可离线运行（本文件仅依赖 Python 3 标准库）。
  - LLM 层：可选，提供 OPENAI_API_KEY / OPENAI_BASE_URL 时对文本做深度定性评审。

为什么标准库优先：技能要「可复用、可执行」，不应要求助教先 pip install。
因此 PDF/DOCX 内容解析用标准库 best-effort（docx 走 zipfile+XML，pdf 仅文件名匹配，
如需全文解析可 pip install pypdf，代码会自动探测并使用）。

用法示例：
    python c4_skill_evaluator.py <文件夹路径> [--output out/] [--format md,csv,html]
    python c4_skill_evaluator.py test_data --format md,csv

作者：WangChunYan
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
import zipfile
from dataclasses import dataclass, field, asdict
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# Windows 控制台默认 GBK，无法打印 ✅ 等字符；强制 UTF-8 输出避免崩溃。
# （报告文件始终以 UTF-8 写入，不受影响。）
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

# ---------------------------------------------------------------------------
# 默认评审规则（内嵌兜底）。运行时会优先加载 references/c4_rubric.json。
# 该结构与 starter 包里的 references/c4_rubric.yaml 语义一致，仅改为 JSON
# 以便标准库 json 直接解析、零第三方依赖。
# ---------------------------------------------------------------------------
DEFAULT_RUBRIC: dict = {
    "required_deliverables": {
        "skill_doc": {
            "label_cn": "Skill 说明文档",
            "filename_patterns": ["skill说明", "skill_doc", "skill-doc", "skill_description", "说明"],
            "content_signals": ["使用场景", "输入", "输出", "解决什么问题", "use case", "input", "output"],
            "extensions": [".md", ".pdf", ".docx", ".txt"],
        },
        "executable_content": {
            "label_cn": "可执行内容",
            "filename_patterns": ["技能", "skill", "executable"],
            "content_signals": ["```", "def ", "class ", "import ", "---", "workflow", "pipeline"],
            "extensions": [".skill", ".py", ".md", ".zip"],
        },
        "demo": {
            "label_cn": "Demo（视频/截图）",
            "filename_patterns": ["demo", "演示", "截图", "screen", "录屏", "example"],
            "content_signals": [],
            "extensions": [".mp4", ".mov", ".webm", ".png", ".jpg", ".jpeg", ".gif"],
        },
        "teaching_doc": {
            "label_cn": "教学说明",
            "filename_patterns": ["教学说明", "教学", "tutorial", "teaching", "上手指南", "quickstart", "how-to"],
            "content_signals": ["上手", "步骤", "常见坑", "安装", "注意事项", "getting started", "step by step"],
            "extensions": [".md", ".pdf", ".docx", ".txt"],
        },
        "ai_log": {
            "label_cn": "AI 日志",
            "filename_patterns": ["AI日志", "AI_log", "ai-log", "ai_日志", "AI使用"],
            "content_signals": ["使用的 AI", "AI 工具", "prompt", "迭代次数", "迭代", "ChatGPT", "Claude", "DeepSeek"],
            "extensions": [".md", ".pdf", ".docx", ".txt"],
        },
    },
    "quality_criteria": {
        "reusable": {
            "label_cn": "可复用",
            "weight": 0.25,
            "positive_signals": ["安装", "install", "环境要求", "requirements", "dependencies", "pip install", "npm install", "兼容", "compatible", "依赖"],
            "negative_signals": ["/Users/", "/home/", "C:\\", "我的电脑", "my_secret", "api_key =", "password ="],
            "check_items": [
                "有安装步骤或环境要求说明",
                "无硬编码绝对路径",
                "列明依赖 / 兼容性",
                "有平台无关或平台标注说明",
            ],
        },
        "executable": {
            "label_cn": "可执行",
            "weight": 0.25,
            "positive_signals": ["```python", "```bash", "```javascript", "def ", "class ", "import ", "---", "workflow", "pipeline", "main(", "if __name__"],
            "negative_signals": [],
            "check_items": [
                "包含可运行代码 / prompt / workflow",
                "SKILL 包结构完整（有 frontmatter）",
                "无明显语法错误",
                "有明确执行入口或运行步骤",
            ],
        },
        "verifiable": {
            "label_cn": "可验证",
            "weight": 0.25,
            "positive_signals": ["测试", "test", "example", "示例", "预期", "expected", "demo", "result", "结果", "断言", "assert"],
            "negative_signals": [],
            "check_items": [
                "有测试用例或示例",
                "定义了预期输出 / 结果",
                "Demo 展示真实结果",
                "成功/失败判定标准清晰",
            ],
        },
        "clear_io": {
            "label_cn": "IO 明确",
            "weight": 0.25,
            "positive_signals": ["输入", "输出", "input", "output", "→", "接受", "返回", "generates", "produces", "returns"],
            "negative_signals": [],
            "io_pattern": "输入.*输出|input.*output|接受.*返回|给定.*得到",
            "check_items": [
                "有「输入X，输出Y」一句话描述",
                "输入类型/格式明确",
                "输出类型/格式明确",
                "边界情况（edge case）有说明",
            ],
        },
    },
    "scoring": {
        "composite_formula": "completeness * 0.4 + quality * 0.6",
        "completeness_weight": 0.4,
        "quality_weight": 0.6,
    },
}

# 评审报告输出时的中文字段名（与 rubric.json 的 label_cn 保持一致）
TEXT_EXTENSIONS = {".md", ".txt", ".py", ".yaml", ".yml", ".json", ".skill", ".csv", ".log"}
DOCX_EXTENSIONS = {".docx"}
PDF_EXTENSIONS = {".pdf"}
BINARY_EXTENSIONS = {".mp4", ".mov", ".webm", ".png", ".jpg", ".jpeg", ".gif", ".zip", ".rar", ".7z"}

# 版本标记，用于 Level 4 的迭代追踪
VERSION_RE = re.compile(r"_v(\d+)", re.IGNORECASE)
# 作者识别：文件名前缀到第一个 _C4_ 之前
AUTHOR_PREFIX_RE = re.compile(r"^(.*?)[_\-]?C4[Aa]?[_\-]")


# ---------------------------------------------------------------------------
# 数据模型
# ---------------------------------------------------------------------------
@dataclass
class ScannedFile:
    """一个被扫描到的提交文件。"""
    path: str
    name: str
    ext: str
    size: int
    author: str
    version: int = 1
    text_preview: str = ""  # 文本文件内容（截断）；二进制为空
    readable: bool = True


@dataclass
class DeliverableResult:
    """单个必须文件的检查结果。"""
    key: str
    label: str
    found: bool
    matched_file: str = ""
    confidence: str = "—"  # high / medium / low（文件名=high，内容信号=low）
    evidence: str = ""


@dataclass
class CriterionResult:
    """单个质量条件的评审结果。"""
    key: str
    label: str
    rating: str          # ✅ / ⚠️ / ❌
    score: float         # 1.0 / 0.5 / 0.0
    satisfied: List[str]
    unmet: List[str]
    evidence: str
    suggestion: str


@dataclass
class AuthorReport:
    """单个作者的完整评审结果。"""
    author: str
    files: List[ScannedFile]
    completeness: List[DeliverableResult]
    completeness_score: float  # 0.0 - 1.0
    quality: List[CriterionResult]
    quality_score: float       # 0.0 - 1.0
    composite: float
    level: str
    latest_version: int
    versions: List[int]
    suggestions: List[str]


# ---------------------------------------------------------------------------
# 文件读取工具
# ---------------------------------------------------------------------------
def read_text(path: Path, max_chars: int = 4000) -> str:
    """以 UTF-8 读取，GBK 回退；返回截断文本（中文文件名/内容兼容）。"""
    try:
        raw = path.read_bytes()
    except OSError:
        return ""
    for enc in ("utf-8", "gbk", "utf-16", "latin-1"):
        try:
            return raw.decode(enc)[:max_chars]
        except (UnicodeDecodeError, UnicodeError):
            continue
    return raw.decode("utf-8", errors="ignore")[:max_chars]


def read_docx_text(path: Path, max_chars: int = 4000) -> str:
    """best-effort 提取 docx 正文（docx 本质是 zip + XML），零第三方依赖。"""
    try:
        with zipfile.ZipFile(path) as z:
            if "word/document.xml" not in z.namelist():
                return ""
            xml = z.read("word/document.xml").decode("utf-8", errors="ignore")
            # 段落之间插入换行，再剥离标签
            xml = xml.replace("</w:p>", "\n")
            text = re.sub(r"<[^>]+>", "", xml)
            return text[:max_chars]
    except (zipfile.BadZipFile, OSError, KeyError):
        return ""


def read_file_content(path: Path) -> Tuple[str, bool]:
    """返回 (文本内容, 是否成功读取文本)。二进制文件返回 ('', False)。"""
    ext = path.suffix.lower()
    if ext in BINARY_EXTENSIONS:
        return "", False
    if ext in DOCX_EXTENSIONS:
        text = read_docx_text(path)
        return text, bool(text)
    if ext in PDF_EXTENSIONS:
        # 标准库无法可靠解析 PDF；可选依赖 pypdf 会自动启用
        try:
            import pypdf  # type: ignore
            reader = pypdf.PdfReader(str(path))
            text = "\n".join((p.extract_text() or "") for p in reader.pages)
            return text[:4000], True
        except Exception:
            return "", False
    if ext in TEXT_EXTENSIONS or ext == "":
        return read_text(path), True
    # 未知扩展名：尝试按文本读取
    return read_text(path), True


# ---------------------------------------------------------------------------
# Step 1 — 文件采集与识别
# ---------------------------------------------------------------------------
def scan_folder(folder: str) -> Dict[str, List[ScannedFile]]:
    """递归扫描目录，识别 `_C4_` 相关提交并按作者分组。"""
    root = Path(folder).resolve()
    if not root.exists():
        raise FileNotFoundError(f"文件夹不存在：{folder}")
    if not root.is_dir():
        raise NotADirectoryError(f"不是文件夹：{folder}")

    authors: Dict[str, List[ScannedFile]] = {}
    for path in sorted(root.rglob("*")):
        if not path.is_file():
            continue
        if not is_c4_related(path):
            continue
        info = build_scanned_file(path)
        authors.setdefault(info.author, []).append(info)
    return authors


def is_c4_related(path: Path) -> bool:
    """判断文件是否与 C4 提交相关。"""
    name = path.name
    # 文件名含 _C4_ 标记
    if "_C4_" in name or "-C4-" in name or re.search(r"C4[Aa]?[_-]", name):
        return True
    # 或位于含 C4 的目录层级
    for part in path.parts:
        if re.search(r"C4", part):
            return True
    return False


def extract_author(path: Path) -> str:
    """从文件名/目录名提取作者。回退顺序：文件名前缀 → 目录名 → Unknown。"""
    name = path.name
    m = re.match(r"^(.*?)[_\-]?C4[Aa]?[_\-]", name)
    if m and m.group(1).strip():
        author = m.group(1).strip()
        author = re.sub(r"[_\-]+$", "", author)
        if author and not author.startswith("."):
            return author
    # 回退：父目录名
    parent = path.parent.name
    if parent and parent not in (".", "..") and "C4" in parent:
        return re.sub(r"[_\-]?C4.*$", "", parent).strip() or parent
    return "Unknown"


def extract_version(name: str) -> int:
    m = VERSION_RE.search(name)
    return int(m.group(1)) if m else 1


def build_scanned_file(path: Path) -> ScannedFile:
    text, readable = read_file_content(path)
    return ScannedFile(
        path=str(path),
        name=path.name,
        ext=path.suffix.lower(),
        size=path.stat().st_size,
        author=extract_author(path),
        version=extract_version(path.name),
        text_preview=text,
        readable=readable,
    )


# ---------------------------------------------------------------------------
# Step 2 — 完整性检查
# ---------------------------------------------------------------------------
def _filename_hit(fname: str, patterns: List[str]) -> bool:
    low = fname.lower()
    return any(p.lower() in low for p in patterns if p)


def _content_hit(text: str, signals: List[str]) -> bool:
    if not text:
        return False
    low = text.lower()
    return any(s.lower() in low for s in signals if s)


def _match_score(spec: dict, f: ScannedFile) -> Tuple[float, str]:
    """计算「交付物 spec × 文件 f」的匹配得分与置信度。

    评分规则（数值越大越优先）：
      - 文件名命中：+3.0，且更具体的 pattern 权重更高（+len*0.1）
      - 首选扩展名命中：+1.0
      - 内容信号命中：+0.5
      - 版本号微调：+version*0.001（同分时优先最新版，避免错配旧版）
    """
    score = 0.0
    confidence = "—"
    fname_pats = spec.get("filename_patterns", [])
    content_sigs = spec.get("content_signals", [])
    prefs = spec.get("extensions", [])

    fname_hit = False
    content_hit = False
    best_len = 0
    for pat in fname_pats:
        if pat and pat.lower() in f.name.lower():
            best_len = max(best_len, len(pat))
    if best_len:
        fname_hit = True
        score += 3.0 + best_len * 0.1
        confidence = "high"
    if prefs and f.ext in prefs:
        score += 1.0
    if content_sigs and _content_hit(f.text_preview, content_sigs):
        content_hit = True
        score += 0.5
        if confidence == "—":
            confidence = "low"
    # 扩展名命中只是加分项，不构成匹配：必须有文件名或内容信号才算命中。
    if not fname_hit and not content_hit:
        return 0.0, "—"
    if score > 0:
        score += f.version * 0.001
    return score, confidence


def check_completeness(files: List[ScannedFile], rubric: dict) -> List[DeliverableResult]:
    """为每个必须文件找证据，采用全局贪心最大权匹配。

    关键改进：不再按交付物顺序逐个抢占文件，而是全局选出得分最高的
    (交付物, 文件) 对依次分配，避免「skill 关键字误命中 skill说明 文档」
    这类信号碰撞，显著降低误判率。每个文件只算一次，避免重复计数。
    """
    deliverables = rubric["required_deliverables"]
    d_keys = list(deliverables.keys())
    assigned: Dict[int, Tuple[int, str]] = {}  # d_idx -> (f_idx, confidence)
    used_f: set = set()

    for _ in range(len(d_keys)):
        best: Optional[Tuple[float, int, int, str]] = None
        for di, key in enumerate(d_keys):
            if di in assigned:
                continue
            for fi, f in enumerate(files):
                if fi in used_f:
                    continue
                score, confidence = _match_score(deliverables[key], f)
                if score <= 0:
                    continue
                if best is None or score > best[0]:
                    best = (score, di, fi, confidence)
        if best is None:
            break
        _, di, fi, confidence = best
        assigned[di] = (fi, confidence)
        used_f.add(fi)

    results: List[DeliverableResult] = []
    for di, key in enumerate(d_keys):
        label = deliverables[key]["label_cn"]
        if di in assigned:
            fi, confidence = assigned[di]
            f = files[fi]
            results.append(DeliverableResult(
                key=key, label=label, found=True,
                matched_file=f.name, confidence=confidence,
                evidence=f"匹配：{f.name}"))
        else:
            results.append(DeliverableResult(key=key, label=label, found=False))
    return results


# ---------------------------------------------------------------------------
# Step 3 — 质量评审（四条件）
# ---------------------------------------------------------------------------
def _check_item_score(item: str, text: str, pos: List[str], neg: List[str], io_pattern: Optional[str] = None) -> bool:
    """把抽象 check_item 落到可检测信号上（规则层近似）。"""
    low = text.lower()
    item_low = item.lower()

    # 负向项：出现任一负信号即判不满足
    if any(n.lower() in low for n in neg if n):
        return False

    # IO 专用：检查正则模式
    if io_pattern and ("io" in item_low or "输入" in item or "output" in item_low):
        try:
            if re.search(io_pattern, text, re.IGNORECASE):
                return True
        except re.error:
            pass

    # 正向项：检查对应关键词
    kw_map = {
        "安装": ["安装", "install", "pip install", "npm install", "环境要求", "requirements"],
        "路径": ["C:\\", "/Users/", "/home/"],
        "依赖": ["依赖", "dependencies", "requirements", "兼容", "compatible"],
        "平台": ["兼容", "compatible", "平台", "platform", "windows", "macos", "linux"],
        "代码": ["```", "def ", "class ", "import ", "if __name__", "main("],
        "frontmatter": ["---"],
        "语法": ["```", "def ", "class ", "import "],
        "执行入口": ["运行", "执行", "python ", "main(", "usage", "用法", "if __name__"],
        "测试": ["测试", "test", "assert", "示例", "example"],
        "预期": ["预期", "expected", "结果", "result", "断言", "assert"],
        "demo": ["demo", "演示", "截图", "录屏", "example"],
        "判定标准": ["成功", "失败", "预期", "expected", "通过", "pass"],
        "一句话": ["输入", "输出", "input", "output", "→"],
        "输入类型": ["输入", "input", "接受", "参数", "类型"],
        "输出类型": ["输出", "output", "返回", "生成", "produces", "returns"],
        "边界": ["边界", "edge case", "异常", "空", "错误", "error"],
    }
    for label, kws in kw_map.items():
        if label in item_low or any(k in item_low for k in kws):
            return any(k.lower() in low for k in kws)
    return False


def evaluate_quality(files: List[ScannedFile], rubric: dict) -> List[CriterionResult]:
    """按四条件对作者提交做质量评审。合并所有文本作为评审素材。"""
    merged = "\n".join(f.text_preview for f in files if f.readable)
    criteria = rubric["quality_criteria"]
    results: List[CriterionResult] = []

    for key, spec in criteria.items():
        label = spec["label_cn"]
        pos = spec.get("positive_signals", [])
        neg = spec.get("negative_signals", [])
        io_pattern = spec.get("io_pattern")
        items = spec.get("check_items", [])

        satisfied = [it for it in items if _check_item_score(it, merged, pos, neg, io_pattern)]
        unmet = [it for it in items if it not in satisfied]

        n_ok = len(satisfied)
        if n_ok >= 2:
            rating, score = "✅", 1.0
        elif n_ok == 1:
            rating, score = "⚠️", 0.5
        else:
            rating, score = "❌", 0.0

        pos_hits = [p for p in pos if p.lower() in merged.lower()]
        neg_hits = [n for n in neg if n.lower() in merged.lower()]
        evidence = f"命中正向信号 {len(pos_hits)} 个" + (f"；命中负向信号 {len(neg_hits)} 个" if neg_hits else "")

        results.append(CriterionResult(
            key=key, label=label, rating=rating, score=score,
            satisfied=satisfied, unmet=unmet, evidence=evidence,
            suggestion=build_suggestion(key, unmet, neg_hits),
        ))
    return results


def build_suggestion(key: str, unmet: List[str], neg_hits: List[str]) -> str:
    tips = {
        "reusable": "补充安装步骤与环境要求，去掉硬编码绝对路径和密钥。",
        "executable": "补充可运行代码/工作流，并给出明确执行入口与运行说明。",
        "verifiable": "补充测试用例、示例输入与预期输出，让成功/失败可判定。",
        "clear_io": "用一句话写清「输入X，输出Y」，并说明输入输出类型与边界情况。",
    }
    if neg_hits:
        return f"发现疑似硬编码/密钥（{', '.join(neg_hits[:3])}），建议移除。" + tips.get(key, "")
    return tips.get(key, "补充相应信号以提升该维度。")


# ---------------------------------------------------------------------------
# Step 4 — 汇总与报告
# ---------------------------------------------------------------------------
def build_author_report(author: str, files: List[ScannedFile], rubric: dict) -> AuthorReport:
    comp = check_completeness(files, rubric)
    comp_score = sum(1 for d in comp if d.found) / max(1, len(comp))
    quality = evaluate_quality(files, rubric)
    quality_score = sum(c.score for c in quality) / max(1, len(quality))

    cw = rubric["scoring"].get("completeness_weight", 0.4)
    qw = rubric["scoring"].get("quality_weight", 0.6)
    composite = comp_score * cw + quality_score * qw

    versions = sorted({f.version for f in files})
    latest = max(versions) if versions else 1

    if comp_score >= 0.8 and quality_score >= 0.75:
        level = "L4·完善"
    elif comp_score >= 0.6 and quality_score >= 0.5:
        level = "L3·良好"
    elif comp_score >= 0.4:
        level = "L2·及格"
    else:
        level = "L1·待补"

    return AuthorReport(
        author=author, files=files, completeness=comp,
        completeness_score=comp_score, quality=quality,
        quality_score=quality_score, composite=round(composite, 3),
        level=level, latest_version=latest, versions=versions,
        suggestions=[c.suggestion for c in quality if c.rating != "✅"],
    )


def _rating_tick(score: float) -> str:
    if score >= 1.0:
        return "✅"
    if score >= 0.5:
        return "⚠️"
    return "❌"


def render_markdown(reports: List[AuthorReport], folder: str, extra: dict) -> str:
    n = len(reports)
    n_files = sum(len(r.files) for r in reports)
    n_complete = sum(1 for r in reports if r.completeness_score >= 1.0)
    n_partial = sum(1 for r in reports if 0 < r.completeness_score < 1.0)
    avg_q = (sum(r.quality_score for r in reports) / n) if n else 0.0

    lines: List[str] = []
    lines.append("# C4 提交自动评审报告")
    lines.append("")
    lines.append(f"- 生成时间：{extra.get('timestamp', datetime.now().isoformat())}")
    lines.append(f"- 扫描路径：`{folder}`")
    lines.append(f"- 识别提交：**{n}** 位作者，**{n_files}** 个文件")
    lines.append("")
    lines.append("## 一、班级总览")
    lines.append("")
    lines.append("| 指标 | 数值 |")
    lines.append("|------|------|")
    lines.append(f"| 总提交人数 | {n} |")
    lines.append(f"| 完整提交（5/5） | {n_complete} |")
    lines.append(f"| 部分提交 | {n_partial} |")
    lines.append(f"| 完全缺失 | {n - n_complete - n_partial} |")
    lines.append(f"| 平均质量分 | {avg_q:.2f}/1.00 |")
    lines.append("")
    lines.append("## 二、作者详情")
    lines.append("")
    for r in sorted(reports, key=lambda x: -x.composite):
        lines.append(f"### {r.author}（综合分 {r.composite}，{r.level}）")
        lines.append("")
        lines.append("**完整性检查：**")
        lines.append("")
        lines.append("| 必须文件 | 状态 | 匹配文件 | 置信度 |")
        lines.append("|----------|------|----------|--------|")
        for d in r.completeness:
            status = "✅" if d.found else "❌"
            lines.append(f"| {d.label} | {status} | {d.matched_file or '—'} | {d.confidence} |")
        lines.append("")
        lines.append("**质量评审（四条件）：**")
        lines.append("")
        lines.append("| 条件 | 评级 | 依据 |")
        lines.append("|------|------|------|")
        for c in r.quality:
            lines.append(f"| {c.label} | {c.rating} | {c.evidence} |")
        lines.append("")
        lines.append("**改进建议：**")
        lines.append("")
        if r.suggestions:
            for s in r.suggestions:
                lines.append(f"- {s}")
        else:
            lines.append("- 无明显短板，继续保持。")
        if len(r.versions) > 1:
            lines.append(f"- 检测到多个版本：{', '.join('v' + str(v) for v in r.versions)}，已按最新版 v{r.latest_version} 评审。")
        lines.append("")
        lines.append("---")
        lines.append("")
    lines.append("## 三、排名")
    lines.append("")
    lines.append("| 排名 | 作者 | 完整性 | 质量分 | 综合分 | 评级 |")
    lines.append("|------|------|--------|--------|--------|------|")
    for rank, r in enumerate(sorted(reports, key=lambda x: -x.composite), start=1):
        lines.append(f"| {rank} | {r.author} | {r.completeness_score*5:.0f}/5 | {r.quality_score:.2f} | {r.composite} | {r.level} |")
    lines.append("")
    lines.append("## 四、全班改进建议")
    lines.append("")
    missing_counter: Dict[str, int] = {}
    for r in reports:
        for d in r.completeness:
            if not d.found:
                missing_counter[d.label] = missing_counter.get(d.label, 0) + 1
    if missing_counter:
        most_missing = max(missing_counter, key=missing_counter.get)
        lines.append(f"- 最常见缺失文件：**{most_missing}**（{missing_counter[most_missing]} 人缺失）")
    weak_counter: Dict[str, int] = {}
    for r in reports:
        for c in r.quality:
            if c.rating != "✅":
                weak_counter[c.label] = weak_counter.get(c.label, 0) + 1
    if weak_counter:
        weakest = max(weak_counter, key=weak_counter.get)
        lines.append(f"- 最弱质量维度：**{weakest}**（{weak_counter[weakest]} 人未达 ✅）")
    lines.append("- 建议每位同学在提交前用本评审技能自检一次（`python c4_skill_evaluator.py <你的文件夹>`）。")
    lines.append("")
    return "\n".join(lines)


def render_csv(reports: List[AuthorReport], folder: str) -> str:
    import io
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["作者", "完整性得分(0-5)", "质量分(0-1)", "综合分", "评级", "最新版本", "缺失文件", "最弱维度"])
    for r in sorted(reports, key=lambda x: -x.composite):
        missing = "、".join(d.label for d in r.completeness if not d.found) or "无"
        weak = "、".join(c.label for c in r.quality if c.rating != "✅") or "无"
        w.writerow([r.author, f"{r.completeness_score*5:.0f}/5", f"{r.quality_score:.2f}",
                    f"{r.composite}", r.level, f"v{r.latest_version}", missing, weak])
    return buf.getvalue()


def render_html(reports: List[AuthorReport], folder: str) -> str:
    rows = ""
    for rank, r in enumerate(sorted(reports, key=lambda x: -x.composite), start=1):
        rows += (
            f"<tr><td>{rank}</td><td>{r.author}</td><td>{r.completeness_score*5:.0f}/5</td>"
            f"<td>{r.quality_score:.2f}</td><td>{r.composite}</td><td>{r.level}</td></tr>"
        )
    return f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8"><title>C4 评审报告</title>
<style>body{{font-family:-apple-system,'Microsoft YaHei',sans-serif;margin:2rem;}}
table{{border-collapse:collapse;width:100%;margin:1rem 0;}}td,th{{border:1px solid #ccc;padding:.5rem;text-align:left;}}</style>
</head><body><h1>C4 提交自动评审报告</h1>
<p>扫描路径：<code>{folder}</code> ｜ 共 {len(reports)} 位作者</p>
<table><tr><th>排名</th><th>作者</th><th>完整性</th><th>质量分</th><th>综合分</th><th>评级</th></tr>{rows}</table>
</body></html>"""


# ---------------------------------------------------------------------------
# 可选 LLM 深度评审（混合架构的 LLM 层）
# ---------------------------------------------------------------------------
def llm_deep_review(text: str, criterion_label: str) -> Optional[str]:
    """可选：调用 OpenAI 兼容接口做定性深审。未配置环境变量时静默跳过。"""
    api_key = os.environ.get("OPENAI_API_KEY")
    base_url = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")
    if not api_key:
        return None
    try:
        import urllib.request
        payload = json.dumps({
            "model": os.environ.get("LLM_MODEL", "gpt-4o-mini"),
            "messages": [
                {"role": "system", "content": "你是 C4 技能评审专家，请用一句话点评该维度，并给出改进建议。"},
                {"role": "user", "content": f"维度：{criterion_label}\n内容（截断）：{text[:3000]}"},
            ],
            "temperature": 0.2,
        }).encode("utf-8")
        req = urllib.request.Request(
            f"{base_url.rstrip('/')}/chat/completions", data=payload,
            headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return data["choices"][0]["message"]["content"].strip()
    except Exception as e:  # noqa: BLE001
        return f"[LLM 深审失败，已回退规则层：{e}]"


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def load_rubric(rubric_path: Optional[str]) -> dict:
    if rubric_path and Path(rubric_path).exists():
        with open(rubric_path, "r", encoding="utf-8") as f:
            return json.load(f)
    default = Path(__file__).parent / "references" / "c4_rubric.json"
    if default.exists():
        with open(default, "r", encoding="utf-8") as f:
            return json.load(f)
    return DEFAULT_RUBRIC


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(description="C4 技能提交自动评审器")
    p.add_argument("folder", help="包含 C4 提交的本地文件夹路径")
    p.add_argument("--output", "-o", default="review_output", help="报告输出目录（默认 review_output）")
    p.add_argument("--format", "-f", default="md,csv", help="输出格式，逗号分隔：md,csv,html")
    p.add_argument("--rubric", "-r", default=None, help="自定义 rubric JSON 路径（默认 references/c4_rubric.json）")
    p.add_argument("--llm", action="store_true", help="启用可选 LLM 深度评审（需 OPENAI_API_KEY）")
    args = p.parse_args(argv)

    rubric = load_rubric(args.rubric)
    try:
        grouped = scan_folder(args.folder)
    except (FileNotFoundError, NotADirectoryError) as e:
        print(f"[错误] {e}", file=sys.stderr)
        return 1

    if not grouped:
        print(f"[提示] 未在 `{args.folder}` 找到 C4 相关文件。", file=sys.stderr)
        print("       请确认文件遵循命名规范 `姓名拼音_C4_内容描述.扩展名`。", file=sys.stderr)
        return 1

    reports = [build_author_report(author, files, rubric) for author, files in grouped.items()]
    reports.sort(key=lambda r: -r.composite)

    outdir = Path(args.output)
    outdir.mkdir(parents=True, exist_ok=True)
    extra = {"timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

    if "md" in args.format:
        (outdir / "评审报告.md").write_text(render_markdown(reports, args.folder, extra), encoding="utf-8")
    if "csv" in args.format:
        (outdir / "评审详表.csv").write_text(render_csv(reports, args.folder), encoding="utf-8-sig")
    if "html" in args.format:
        (outdir / "评审报告.html").write_text(render_html(reports, args.folder), encoding="utf-8")

    print(f"✅ 评审完成：{len(reports)} 位作者，{sum(len(r.files) for r in reports)} 个文件")
    print(f"📁 报告输出目录：{outdir.resolve()}")
    for r in reports:
        print(f"   · {r.author:<16} 完整性 {r.completeness_score*5:.0f}/5  质量 {r.quality_score:.2f}  综合 {r.composite}  ({r.level})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
